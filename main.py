"""
main.py — Serviço "gerar_parecer" (chamado por uma ação HTTP dentro do
flow do Power Automate, no lugar da ação nativa "Popular um modelo do
Word" — que não funciona porque os Content Controls do template estão
dentro de uma tabela).
-----------------------------------------------------------------
Recebe do FLOW (não mais direto do Copilot Studio) o JSON com os 138
campos que o agente já montou, no MESMO formato de `dados_parecer_json`
descrito nas instruções do agente (ver instrucoes_agente_v3_flow.md).

O que este serviço faz, na ordem:
  1. Recebe o JSON de 138 campos direto no corpo da requisição (POST).
  2. Preenche o ModeloParecer_jinja.docx via docxtpl (headless, sem
     precisar do Word — e sem se importar se os campos estão dentro de
     tabela ou não, ao contrário da ação nativa do Word Online).
  3. Sobe o .docx pro SharePoint via Microsoft Graph API.
  4. Baixa o base_talentos.xlsx atual do SharePoint, faz upsert (mesma
     regra de dedupe do registrar_talentos.py original) e reenvia.
  5. Devolve JSON pro flow: link do documento + status da base. O flow
     só repassa isso pro "Respond to the agent".

Rodar localmente pra testar:
    uvicorn main:app --host 0.0.0.0 --port 8000

Variáveis de ambiente necessárias — ver .env.example
"""

import io
import os
import re
import time
import unicodedata
import datetime
from pathlib import Path

import requests
from fastapi import FastAPI, Header, HTTPException
from docxtpl import DocxTemplate
from openpyxl import Workbook, load_workbook
import msal

# ============================================================
# Config (variáveis de ambiente — ver .env.example)
# ============================================================
TENANT_ID = os.environ.get("TENANT_ID", "")
CLIENT_ID = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
SITE_ID = os.environ.get("SITE_ID", "")  # id do site do SharePoint (não a URL)

# Caminhos DENTRO da biblioteca de documentos do site (Graph usa path relativo à raiz do drive)
SHAREPOINT_DOCS_FOLDER = os.environ.get("SHAREPOINT_DOCS_FOLDER", "Apoio/Documentos/Parecer/Saida")
SHAREPOINT_EXCEL_PATH = os.environ.get("SHAREPOINT_EXCEL_PATH", "Apoio/Documentos/Parecer/Base/base_talentos.xlsx")

# Chave simples que o Copilot Studio envia no header pra autenticar a chamada
API_KEY = os.environ.get("PARECER_API_KEY", "")

# MODO DEMO: liga automaticamente se as credenciais do Graph não estiverem
# configuradas (ex.: enquanto o admin não concede o Sites.Selected). Nesse
# modo o serviço gera o Word e atualiza o Excel LOCALMENTE, em vez de subir
# pro SharePoint — serve pra validar a conexão com o Copilot Studio e a
# geração do documento sem depender do acesso ao SharePoint ainda.
DEMO_MODE = os.environ.get("PARECER_DEMO_MODE", "").lower() == "true" or not all(
    [TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID]
)
LOCAL_OUTPUT_DIR = Path(__file__).parent / "output_demo"

TEMPLATE_PATH = Path(__file__).parent / "ModeloParecer_jinja.docx"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

BASE_TALENTOS_FIELDS = [
    "Nome_Completo", "Telefone", "Cidade", "UF", "LinkedIn_URL",
    "Cliente", "Cargo_Aplicado", "Senioridade_Aplicada", "Data_Entrevista",
    "Anos_Exp_No_Cargo", "Anos_Exp_Geral", "Recomendacao",
    "Escolaridade", "Area_Formacao", "Recrutador_Responsavel",
    "Palavras_Chave_5", "Observacoes",
]

# Campos do parecer (mesmo FIELD_ORDER do gerar_parecer.py original)
FIELD_ORDER = [
    "Cargo", "Cliente", "Recrutador", "Data_Envio", "Recomendacao",
    "Nome_Candidato", "Tempo_Experiencia",
]
for i in range(1, 16):
    FIELD_ORDER += [f"REQO{i:02d}", f"REQO{i:02d}_Analise", f"REQO{i:02d}_Comprovante"]
for i in range(1, 11):
    FIELD_ORDER += [f"REQD{i:02d}", f"REQD{i:02d}_Analise", f"REQD{i:02d}_Comprovante"]
for i in range(1, 11):
    FIELD_ORDER += [f"CERO{i:02d}", f"CERO{i:02d}_Analise", f"CERO{i:02d}_Comprovante"]
FIELD_ORDER += [
    "Comunicacao", "Comunicacao_Analise", "Comunicacao_Comprovante",
    "Organizacao", "Organizacao_Analise", "Organizacao_Comprovante",
    "Analise_Critica", "Analise_Critica_Analise", "Analise_Critica_Comprovante",
    "Criatividade", "Criatividade_Analise", "Criatividade_Comprovante",
    "Inovacao", "Inovacao_Analise", "Inovacao_Comprovante",
    "Parecer_Final",
]

# Rótulos estáticos das competências — o agente não precisa mais enviar
# isso, o servidor já preenche sozinho (evita 5 campos redundantes).
STATIC_LABELS = {
    "Comunicacao": "Comunicação",
    "Organizacao": "Organização",
    "Analise_Critica": "Análise Crítica",
    "Criatividade": "Criatividade",
    "Inovacao": "Inovação",
}

ALL_TEMPLATE_FIELDS = FIELD_ORDER  # os 128 campos que o docxtpl espera


# ============================================================
# Normalização do payload — o agente já manda JSON estruturado (via
# Parse JSON no flow), então aqui só garantimos que as 138 chaves
# existem (preenche com "" o que faltar) e limpamos tipos estranhos.
# ============================================================
def _norm(s: str) -> str:
    s = (s or "").strip().replace("-", "_").replace(" ", "_")
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.lower()


def sanitize_filename(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", (s or "").strip())


def parse_date_flexible(s: str):
    if not s:
        return None
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def normalize_payload(raw: dict) -> dict:
    """Garante que todas as 138 chaves esperadas existem, mesmo que o
    agente tenha mandado alguma vazia/faltando. Valores None viram ''."""
    all_keys = FIELD_ORDER + [k for k in BASE_TALENTOS_FIELDS if k not in FIELD_ORDER]
    return {k: ("" if raw.get(k) is None else str(raw.get(k))) for k in all_keys}


# ============================================================
# Microsoft Graph — autenticação e I/O no SharePoint
# ============================================================
def get_graph_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Falha ao autenticar no Graph: {result.get('error_description')}")
    return result["access_token"]


def graph_upload_file(token: str, relative_path: str, content: bytes) -> dict:
    url = f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{relative_path}:/content"
    resp = requests.put(
        url,
        headers={"Authorization": f"Bearer {token}"},
        data=content,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()


def graph_download_file(token: str, relative_path: str) -> bytes | None:
    url = f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{relative_path}:/content"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.content


# ============================================================
# Excel — mesma lógica de upsert do registrar_talentos.py original
# ============================================================
def ensure_excel_headers(ws):
    if ws.max_row == 1 and all(ws.cell(1, c + 1).value is None for c in range(len(BASE_TALENTOS_FIELDS))):
        for i, h in enumerate(BASE_TALENTOS_FIELDS, start=1):
            ws.cell(1, i, h)


def _columns_map():
    return {h: i for i, h in enumerate(BASE_TALENTOS_FIELDS, start=1)}


def _row_key_tuple(row_dict: dict):
    return (
        _norm(row_dict.get("Nome_Completo", "")),
        _norm(row_dict.get("Cliente", "")),
        _norm(row_dict.get("Cargo_Aplicado", "")),
        _norm(row_dict.get("Recrutador_Responsavel", "")),
    )


def _find_matching_row(ws, row_dict: dict):
    key = _row_key_tuple(row_dict)
    if any(not part for part in key):
        return None
    cols = _columns_map()
    max_row = ws.max_row or 1
    for r in range(2, max_row + 1):
        k2 = (
            _norm(ws.cell(r, cols["Nome_Completo"]).value or ""),
            _norm(ws.cell(r, cols["Cliente"]).value or ""),
            _norm(ws.cell(r, cols["Cargo_Aplicado"]).value or ""),
            _norm(ws.cell(r, cols["Recrutador_Responsavel"]).value or ""),
        )
        if k2 == key:
            return r
    return None


def _is_newer(new_date_str: str, old_date_str: str) -> bool:
    nd = parse_date_flexible(new_date_str)
    od = parse_date_flexible(old_date_str)
    if nd and od:
        return nd >= od
    return True


def upsert_talentos(existing_bytes: bytes | None, row: dict) -> tuple[bytes, str]:
    if existing_bytes:
        wb = load_workbook(io.BytesIO(existing_bytes))
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active

    ensure_excel_headers(ws)
    cols = _columns_map()
    r_match = _find_matching_row(ws, row)

    if r_match is None:
        ws.append([row.get(k, "") for k in BASE_TALENTOS_FIELDS])
        status = "appended"
    else:
        old_date = ws.cell(r_match, cols["Data_Entrevista"]).value or ""
        new_date = row.get("Data_Entrevista", "")
        if _is_newer(new_date, old_date):
            for i, field in enumerate(BASE_TALENTOS_FIELDS, start=1):
                ws.cell(r_match, i, row.get(field, ""))
            status = "updated"
        else:
            status = "skipped"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), status


# ============================================================
# Geração do Word via docxtpl
# ============================================================
def render_docx(payload: dict) -> bytes:
    ctx = {f: payload.get(f, "") for f in ALL_TEMPLATE_FIELDS}
    ctx.update(STATIC_LABELS)  # garante os rótulos das competências
    tpl = DocxTemplate(str(TEMPLATE_PATH))
    tpl.render(ctx)
    buf = io.BytesIO()
    tpl.save(buf)
    return buf.getvalue()


# ============================================================
# API
# ============================================================
app = FastAPI(title="Gerar Parecer", version="1.0")


@app.post("/gerar-parecer")
def gerar_parecer(body: dict, x_api_key: str = Header(default="")):
    """Recebe o JSON de 138 campos direto (o mesmo que o Parse JSON do
    flow monta a partir de dados_parecer_json). Chamado pela ação HTTP
    do flow — não é mais chamado direto pelo Copilot Studio."""
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key inválida.")

    if not body:
        raise HTTPException(status_code=400, detail="Corpo da requisição vazio.")

    payload = normalize_payload(body)

    # nome do arquivo
    hoje = datetime.date.today()
    data_envio = parse_date_flexible(payload.get("Data_Envio")) or hoje
    cliente = sanitize_filename(payload.get("Cliente"))
    cargo = sanitize_filename(payload.get("Cargo"))
    cand = sanitize_filename(payload.get("Nome_Candidato"))
    dt_str = data_envio.strftime("%Y-%m-%d")
    filename = f"Parecer - {cliente} - {cargo} - {cand} - {dt_str}.docx"

    # 1) Gera o docx (sempre local, independe do SharePoint)
    try:
        docx_bytes = render_docx(payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Falha ao gerar o documento Word: {e}")

    tem_dados_base = any(
        (payload.get(k) or "").strip() for k in ("Nome_Completo", "LinkedIn_URL", "Telefone")
    )

    if DEMO_MODE:
        # ---- Caminho DEMO: salva local em vez de ir pro SharePoint ----
        LOCAL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        doc_path = LOCAL_OUTPUT_DIR / filename
        doc_path.write_bytes(docx_bytes)

        base_status = "sem_dados_minimos"
        if tem_dados_base:
            excel_path = LOCAL_OUTPUT_DIR / "base_talentos.xlsx"
            existing = excel_path.read_bytes() if excel_path.exists() else None
            new_bytes, base_status = upsert_talentos(existing, payload)
            excel_path.write_bytes(new_bytes)

        return {
            "status": "sucesso",
            "modo": "DEMO — arquivo salvo localmente no servidor, SharePoint ainda não configurado",
            "documento_url": f"(demo) {doc_path}",
            "documento_nome": filename,
            "base_talentos_status": base_status,
        }

    # ---- Caminho real: sobe pro SharePoint via Microsoft Graph ----
    try:
        token = get_graph_token()
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Falha ao autenticar no Microsoft Graph: {e}")

    try:
        doc_result = graph_upload_file(token, f"{SHAREPOINT_DOCS_FOLDER}/{filename}", docx_bytes)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Falha ao subir o documento no SharePoint: {e}")

    base_status = "erro"
    if tem_dados_base:
        try:
            existing = graph_download_file(token, SHAREPOINT_EXCEL_PATH)
            new_bytes, base_status = upsert_talentos(existing, payload)
            graph_upload_file(token, SHAREPOINT_EXCEL_PATH, new_bytes)
        except Exception as e:
            base_status = f"erro: {e}"
    else:
        base_status = "sem_dados_minimos"

    return {
        "status": "sucesso",
        "modo": "SharePoint",
        "documento_url": doc_result.get("webUrl"),
        "documento_nome": filename,
        "base_talentos_status": base_status,
    }


@app.get("/health")
def health():
    return {"status": "ok", "modo": "DEMO (sem SharePoint)" if DEMO_MODE else "SharePoint"}
