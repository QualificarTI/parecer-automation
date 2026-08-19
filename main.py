# Índices das tabelas do template que têm linhas "opcionais" (slots de
# REQO/REQD/CERO que nem sempre são todos usados pela vaga). As duas
# primeiras linhas de cada uma são cabeçalho (título do grupo + nomes das
# colunas) — os dados começam na linha de índice 2. Depende da estrutura
# atual do ModeloParecer_jinja.docx (tabela 2 = Requisitos Obrigatórios,
# 3 = Desejáveis, 4 = Certificações) — se o template mudar de estrutura,
# esses índices precisam ser revisados.
REQUIREMENT_TABLE_INDEXES = [2, 3, 4]
HEADER_ROWS_PER_TABLE = 2

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _row_text(tr) -> str:
    """Extrai todo o texto de uma linha de tabela (<w:tr>), incluindo texto
    dentro de células envolvidas em controles de conteúdo (<w:sdt>). O
    python-docx padrão (row.cells) não enxerga células assim — só olha
    filhos diretos de <w:tr> — por isso precisamos ler o XML na mão."""
    return "".join(t.text or "" for t in tr.iter(_W + "t"))


def remove_empty_requirement_rows(doc):
    """Remove linhas de tabela cujos campos ficaram todos vazios (slots de
    REQO/REQD/CERO não usados pela vaga) — evita linhas em branco feias no
    documento final."""
    for t_idx in REQUIREMENT_TABLE_INDEXES:
        if t_idx >= len(doc.tables):
            continue
        table = doc.tables[t_idx]
        all_trs = table._tbl.findall(_W + "tr")
        data_trs = all_trs[HEADER_ROWS_PER_TABLE:]
        for tr in data_trs:
            if _row_text(tr).strip() == "":
                tr.getparent().remove(tr)


def remove_empty_requirement_blocks(doc):
    """Remove a tabela inteira (título + cabeçalho de colunas, ambos em
    azul) quando nenhuma linha de dado sobrou nela — ex.: vaga sem nenhum
    requisito desejável ou sem nenhuma certificação obrigatória. Precisa
    rodar DEPOIS de remove_empty_requirement_rows, senão nunca vai achar
    tabela "vazia" (as linhas em branco ainda estariam lá)."""
    target_tables = [doc.tables[i] for i in REQUIREMENT_TABLE_INDEXES if i < len(doc.tables)]
    for table in target_tables:
        remaining_trs = table._tbl.findall(_W + "tr")
        if len(remaining_trs) <= HEADER_ROWS_PER_TABLE:
            table._tbl.getparent().remove(table._tbl)


def render_docx(payload: dict) -> bytes:
    ctx = {f: payload.get(f, "") for f in ALL_TEMPLATE_FIELDS}
    ctx.update(STATIC_LABELS)  # garante os rótulos das competências
    tpl = DocxTemplate(str(TEMPLATE_PATH))
    tpl.render(ctx)
    buf = io.BytesIO()
    tpl.save(buf)
    data = buf.getvalue()

    # Autoverificação: tenta reabrir o .docx recém-gerado antes de devolvê-lo.
    # Um .zip pode estar tecnicamente íntegro (unzip -t não acusa nada) e
    # mesmo assim ter XML interno inválido, que o Word recusa abrir. Se
    # isso acontecer aqui, preferimos falhar com um erro 500 claro do que
    # devolver pro flow um arquivo que só vai quebrar na hora de abrir no
    # Word.
    try:
        doc = Document(io.BytesIO(data))
        _ = len(doc.paragraphs)  # força o parse completo do document.xml
    except Exception as e:
        raise RuntimeError(
            f"O .docx gerado não passou na verificação de integridade "
            f"(provavelmente algum campo do payload contém caracteres "
            f"inválidos para XML): {e}"
        )

    # Remove linhas vazias (slots de requisito/certificação não usados) e,
    # se a tabela inteira ficou sem nenhuma linha de dado, remove o bloco
    # inteiro (título + cabeçalho). Depois resalva e reverifica de novo
    # pra garantir que a limpeza não quebrou nada.
    remove_empty_requirement_rows(doc)
    remove_empty_requirement_blocks(doc)   # <- linha nova
    buf2 = io.BytesIO()
    doc.save(buf2)
    data = buf2.getvalue()
    try:
        Document(io.BytesIO(data))
    except Exception as e:
        raise RuntimeError(f"Falha ao limpar linhas vazias do documento: {e}")

    return data
