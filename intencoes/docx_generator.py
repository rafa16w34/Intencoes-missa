import os
import io
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from django.conf import settings

COR_PRETA = RGBColor(0, 0, 0)

LINHA1_PAROQUIA = "MITRA DIOCESANA DE DIVINÓPOLIS"
LINHA2_PAROQUIA = "PARÓQUIA DIVINO ESPÍRITO SANTO E SÃO FRANCISCO DE PAULA"
CNPJ = "20.169.827/0005-64"
ENDERECO = "PÇA. DOM CRISTIANO, S/Nº - CENTRO - DIVINÓPOLIS"
EMAIL = "catedral.divinopolis@diocesedivinopolis.org.br"
FONE = "(37) 3221-9242"

# Mapeamento do código do modelo IntencaoMissa para o título oficial do bloco no Word
ORDEM_BLOCOS = [
    ("honra", "HONRA E LOUVOR"),
    ("niver", "ANIVERSÁRIO NATALÍCIO"),
    ("casamento", "ANIVERSÁRIO DE CASAMENTO"),
    ("acao_gracas", "AÇÃO DE GRAÇAS"),
    ("intencao", "INTENÇÃO"),
    ("saude", "RECUPERAÇÃO DE SAÚDE"),
    ("setimo", "7º DIA"),
    ("falecido", "IRMÃOS FALECIDOS"),
]


def quebrar_linha(texto, limite=100):
    """Quebra textos muito longos com hífen para não desformatar a folha."""
    palavras = texto.split()
    linhas = []
    linha_atual = ""

    for palavra in palavras:
        if len(linha_atual) + len(palavra) + 1 <= limite:
            if linha_atual:
                linha_atual += " "
            linha_atual += palavra
        else:
            linhas.append(linha_atual + "-")
            linha_atual = "-" + palavra

    if linha_atual:
        linhas.append(linha_atual)

    return linhas


def _remover_margens_celula(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for lado in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _definir_bordas_celula(cell, **bordas):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for borda, dados in bordas.items():
        tag = f"w:{borda}"
        elemento = tcBorders.find(qn(tag))
        if elemento is None:
            elemento = OxmlElement(tag)
            tcBorders.append(elemento)
        elemento.set(qn("w:val"), dados.get("val", "single"))
        elemento.set(qn("w:sz"), str(dados.get("sz", 5)))
        elemento.set(qn("w:space"), str(dados.get("space", 0)))
        elemento.set(qn("w:color"), dados.get("color", "000000"))


def _remover_bordas_tabela(tabela):
    tbl = tabela._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for nome in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{nome}")
        el.set(qn("w:val"), "none")
        tblBorders.append(el)
    tblPr.append(tblBorders)


def _formatar_run(run, fonte="Arial", tamanho=12, negrito=False, italico=False):
    run.font.name = fonte
    run.font.size = Pt(tamanho)
    run.font.bold = negrito
    run.font.italic = italico
    run.font.color.rgb = COR_PRETA
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), fonte)


def _sem_espacamento(paragrafo):
    pf = paragrafo.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def _montar_cabecalho(section):
    header = section.header
    header.is_linked_to_previous = False
    header.paragraphs[0].text = ""

    largura_util = section.page_width - section.left_margin - section.right_margin
    tabela = header.add_table(rows=1, cols=3, width=largura_util)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.autofit = False
    _remover_bordas_tabela(tabela)

    larguras = [Cm(2.98), Cm(12.94), Cm(2.98)]
    for coluna, largura in zip(tabela.columns, larguras):
        coluna.width = largura

    celula_esq, celula_meio, celula_dir = tabela.rows[0].cells
    for celula in (celula_esq, celula_meio, celula_dir):
        celula.width = larguras[0]
        _definir_bordas_celula(celula, bottom={"sz": 5, "color": "000000"})

    _remover_margens_celula(celula_esq)
    _remover_margens_celula(celula_dir)

    logo_esq = os.path.join(settings.BASE_DIR, "static", "img", "logo_diocese.jpg")
    logo_dir = os.path.join(settings.BASE_DIR, "static", "img", "logo_cated.jpg")
    if not os.path.exists(logo_esq):
        logo_esq = os.path.join(settings.BASE_DIR, "src", "assets", "logo_diocese.jpg")
    if not os.path.exists(logo_dir):
        logo_dir = os.path.join(settings.BASE_DIR, "src", "assets", "logo_cated.jpg")

    p = celula_esq.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(logo_esq):
        p.add_run().add_picture(logo_esq, width=Cm(3))

    p = celula_dir.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(logo_dir):
        p.add_run().add_picture(logo_dir, width=Cm(3))

    linhas = [
        [(LINHA1_PAROQUIA, True)],
        [(LINHA2_PAROQUIA, True)],
        [("CNPJ: ", True), (CNPJ, False)],
        [(ENDERECO + " - ", False), ("CEP:", True), (" 35500-004", False)],
        [("E-mail: ", True), (EMAIL, False)],
        [("Fone: ", True), (FONE, False)],
    ]
    celula_meio.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    primeiro = True
    for partes in linhas:
        p = celula_meio.paragraphs[0] if primeiro else celula_meio.add_paragraph()
        primeiro = False
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _sem_espacamento(p)
        for texto, negrito in partes:
            _formatar_run(p.add_run(texto), fonte="Times New Roman", tamanho=9, negrito=negrito)


def _adicionar_campo(paragrafo, nome_campo):
    run = paragrafo.add_run()
    _formatar_run(run, fonte="Times New Roman", tamanho=8)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {nome_campo} "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._element.append(fld_char1)
    run._element.append(instr)
    run._element.append(fld_char2)


def _montar_rodape(section, data_str):
    footer = section.footer
    footer.is_linked_to_previous = False
    footer.paragraphs[0].text = ""

    largura_util = section.page_width - section.left_margin - section.right_margin
    tabela = footer.add_table(rows=1, cols=2, width=largura_util)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.autofit = False
    _remover_bordas_tabela(tabela)

    larguras = [Cm(14.0), Cm(4.9)]
    for coluna, largura in zip(tabela.columns, larguras):
        coluna.width = largura

    celula_esq, celula_dir = tabela.rows[0].cells
    for celula in (celula_esq, celula_dir):
        _definir_bordas_celula(celula, top={"sz": 5, "color": "000000"})

    celula_esq.paragraphs[0].text = ""

    p = celula_dir.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _sem_espacamento(p)
    _formatar_run(p.add_run(f"Data: {data_str}"), fonte="Times New Roman", tamanho=8)

    p2 = celula_dir.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _sem_espacamento(p2)
    run = p2.add_run("Página ")
    _formatar_run(run, fonte="Times New Roman", tamanho=8)
    _adicionar_campo(p2, "PAGE")
    run = p2.add_run(" de ")
    _formatar_run(run, fonte="Times New Roman", tamanho=8)
    _adicionar_campo(p2, "NUMPAGES")


def _nova_linha(tabela):
    tabela.add_row()
    return tabela.rows[-1].cells[0]


def _escrever_titulo_documento(tabela, texto):
    celula = tabela.rows[0].cells[0] if len(tabela.rows) == 1 and not tabela.rows[0].cells[0].paragraphs[0].runs else _nova_linha(tabela)
    p = celula.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _sem_espacamento(p)
    _formatar_run(p.add_run(texto), fonte="Times New Roman", tamanho=18, negrito=True)


def _escrever_horario(tabela, texto):
    celula = _nova_linha(tabela)
    p = celula.paragraphs[0]
    _sem_espacamento(p)
    _formatar_run(p.add_run(texto), fonte="Times New Roman", tamanho=16, negrito=True)


def _escrever_espacador(tabela):
    _nova_linha(tabela)


def _escrever_bloco(tabela, titulo, itens):
    celula_titulo = _nova_linha(tabela)
    _definir_bordas_celula(celula_titulo, bottom={"sz": 10, "color": "000000"})
    p = celula_titulo.paragraphs[0]
    _sem_espacamento(p)
    _formatar_run(p.add_run(f"{titulo}:"), fonte="Times New Roman", tamanho=16, negrito=True, italico=True)

    celula_itens = _nova_linha(tabela)
    _definir_bordas_celula(celula_itens, bottom={"sz": 5, "color": "000000"})
    primeiro = True
    for item in itens:
        linhas_quebradas = quebrar_linha(item)
        for l in linhas_quebradas:
            p = celula_itens.paragraphs[0] if primeiro else celula_itens.add_paragraph()
            primeiro = False
            _sem_espacamento(p)
            _formatar_run(p.add_run(l), fonte="Arial", tamanho=12)

    if primeiro:
        _sem_espacamento(celula_itens.paragraphs[0])

    _escrever_espacador(tabela)


def gerar_docx_missa(missa):
    """
    Gera o documento .docx formatado para a Missa informada
    e retorna um buffer io.BytesIO pronto para download.
    """
    documento = Document()

    section = documento.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.2)
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.5)

    data_formatada = missa.data.strftime("%d/%m/%Y")
    hora_formatada = missa.horario.strftime("%H:%M")

    _montar_cabecalho(section)
    _montar_rodape(section, f"{data_formatada} - {hora_formatada}")

    largura_util = section.page_width - section.left_margin - section.right_margin
    tabela = documento.add_table(rows=1, cols=1)
    tabela.autofit = False
    tabela.columns[0].width = largura_util
    tabela.rows[0].cells[0].width = largura_util
    _remover_bordas_tabela(tabela)

    _escrever_titulo_documento(tabela, f"Intenções de {data_formatada}")
    _escrever_horario(tabela, f"Missa das {hora_formatada}h")
    _escrever_espacador(tabela)

    # Apenas intenções com pagamento confirmado (status='pago') são impressas no Word
    intencoes_qs = missa.intencoes.filter(
        pagamento__isnull=False,
        pagamento__status='pago'
    ).select_related('pagamento').order_by('id')

    for tipo_chave, titulo_bloco in ORDEM_BLOCOS:
        itens_cat = [it for it in intencoes_qs if it.tipo == tipo_chave]
        sec_itens = [it for it in itens_cat if it.pagamento and it.pagamento.forma == 'secretaria']
        ses_itens = [it for it in itens_cat if not (it.pagamento and it.pagamento.forma == 'secretaria')]

        itens_formatados = []
        for it in sec_itens:
            itens_formatados.append(it.nome_intencao.rstrip(" *"))

        for idx, it in enumerate(ses_itens):
            n = it.nome_intencao.rstrip(" *")
            if idx == 0:
                # Primeira intenção da sessão nesta categoria recebe '*' para diferenciação no docx
                itens_formatados.append(f"{n} *")
            else:
                itens_formatados.append(n)

        _escrever_bloco(tabela, titulo_bloco, itens_formatados)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)
    return buffer
