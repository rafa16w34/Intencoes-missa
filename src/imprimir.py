"""
Geração das "Intenções" em .docx, no padrão visual usado pela paróquia
(cabeçalho com brasão + dados da paróquia, título centralizado, blocos com
título em negrito/itálico e linhas horizontais para preenchimento manual).

Substitui o antigo módulo baseado em openpyxl. As variáveis importadas de
config.py continuam as mesmas (data_formatada, categorias, valor_total) —
só a forma de "imprimir" mudou, de planilha para Word.
"""

import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from config import data_formatada, categorias,hora_formatada, registro  # valor_total mantido por compatibilidade (não usado na impressão)

def remover_margens_celula(cell):
    tcPr = cell._tc.get_or_add_tcPr()

    tcMar = OxmlElement("w:tcMar")

    for lado in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)

    tcPr.append(tcMar)
# ------------------------------------------------------------------------------------------------------------------
# DADOS FIXOS DO CABEÇALHO — ajuste aqui se algum dado da paróquia mudar
# ------------------------------------------------------------------------------------------------------------------

PASTA_ASSETS = os.path.join(os.path.dirname(__file__), "assets")
LOGO_ESQUERDA = os.path.join(PASTA_ASSETS, "logo_diocese.jpg")
LOGO_DIREITA = os.path.join(PASTA_ASSETS, "logo_cated.jpg")

LINHA1_PAROQUIA = "MITRA DIOCESANA DE DIVINÓPOLIS"
LINHA2_PAROQUIA = "PARÓQUIA DIVINO ESPÍRITO SANTO E SÃO FRANCISCO DE PAULA"
CNPJ = "20.169.827/0005-64"
ENDERECO = "PÇA. DOM CRISTIANO, S/Nº - CENTRO - DIVINÓPOLIS"
EMAIL = "catedral.divinopolis@diocesedivinopolis.org.br"
FONE = "(37) 3221-9242"

# Horário(s) de missa exibido(s) logo abaixo do título. Ajuste conforme a celebração.
HORARIO_MISSA = "Missa das " + hora_formatada

# Ordem em que os blocos aparecem no documento (substitui o "layout" de colunas do Excel,
# já que agora tudo flui verticalmente em uma coluna só, como no modelo novo)
ORDEM_BLOCOS = [
    "HONRA E LOUVOR",
    "ANIVERSÁRIO NATALÍCIO",
    "ANIVERSÁRIO DE CASAMENTO",
    "AÇÃO DE GRAÇAS",
    "INTENÇÃO",
    "RECUPERAÇÃO DE SAÚDE",
    "7º DIA",
    "IRMÃOS FALECIDOS",
]

COR_PRETA = RGBColor(0, 0, 0)

# ------------------------------------------------------------------------------------------------------------------
# HELPERS DE BAIXO NÍVEL (bordas de célula — o python-docx não expõe isso diretamente)
# ------------------------------------------------------------------------------------------------------------------

def _definir_bordas_celula(cell, **bordas):

    """
    Define bordas individuais de uma célula de tabela.
    Uso: _definir_bordas_celula(celula, top={"sz": 10}, bottom={"sz": 5})
    Cada valor é um dict opcional com sz (espessura em oitavos de ponto), color e val.
    """

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

    """Remove qualquer borda padrão da tabela — as linhas visuais são feitas célula a célula."""

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

    """Remove espaçamento extra antes/depois do parágrafo (visual mais compacto, como no modelo)."""

    pf = paragrafo.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


# ------------------------------------------------------------------------------------------------------------------
# CABEÇALHO E RODAPÉ
# ------------------------------------------------------------------------------------------------------------------

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

    remover_margens_celula(celula_esq)
    remover_margens_celula(celula_dir)

    # Logo esquerda
    p = celula_esq.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(LOGO_ESQUERDA):
        p.add_run().add_picture(LOGO_ESQUERDA, width=Cm(3))

    # Logo direita
    p = celula_dir.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(LOGO_DIREITA):
        p.add_run().add_picture(LOGO_DIREITA, width=Cm(3))

    # Texto central (partes em negrito misturadas com texto normal, como no modelo)
    linhas = [
        [(LINHA1_PAROQUIA, True)],
        [(LINHA2_PAROQUIA, True)],
        [("CNPJ: ", True), (CNPJ, False)],
        [(ENDERECO + " - ", False), ("CEP:", True), (" 35500-004".replace("35500-004", "35500-004"), False)],
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


def _montar_rodape(section, data_impressao):
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

    celula_esq.paragraphs[0].text = ""  # coluna vazia (só para manter a régua alinhada)

    p = celula_dir.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _sem_espacamento(p)
    _formatar_run(p.add_run(f"Data: {data_impressao}"), fonte="Times New Roman", tamanho=8)

    p2 = celula_dir.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _sem_espacamento(p2)
    run = p2.add_run("Página ")
    _formatar_run(run, fonte="Times New Roman", tamanho=8)
    _adicionar_campo(p2, "PAGE")
    run = p2.add_run(" de ")
    _formatar_run(run, fonte="Times New Roman", tamanho=8)
    _adicionar_campo(p2, "NUMPAGES")


def _adicionar_campo(paragrafo, nome_campo):
    """Insere um campo dinâmico do Word (ex.: número de página automático)."""
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


# ------------------------------------------------------------------------------------------------------------------
# CORPO DO DOCUMENTO (título + blocos com linhas para preenchimento manual)
# ------------------------------------------------------------------------------------------------------------------

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
    # Linha do título do bloco — negrito + itálico, com régua grossa embaixo
    celula_titulo = _nova_linha(tabela)
    _definir_bordas_celula(celula_titulo, bottom={"sz": 10, "color": "000000"})
    p = celula_titulo.paragraphs[0]
    _sem_espacamento(p)
    _formatar_run(p.add_run(f"{titulo}:"), fonte="Times New Roman", tamanho=16, negrito=True, italico=True)

    # Linha única contendo todos os itens (um parágrafo por item), com régua fina embaixo
    celula_itens = _nova_linha(tabela)
    _definir_bordas_celula(celula_itens, bottom={"sz": 5, "color": "000000"})
    primeiro = True
    for item in itens:
        p = celula_itens.paragraphs[0] if primeiro else celula_itens.add_paragraph()
        primeiro = False
        _sem_espacamento(p)
        _formatar_run(p.add_run(item), fonte="Arial", tamanho=12)
    if primeiro:
        # bloco sem nenhum item: ainda assim deixa a célula com um parágrafo vazio
        _sem_espacamento(celula_itens.paragraphs[0])


    # Espaço extra entre blocos
    _escrever_espacador(tabela)


# ------------------------------------------------------------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ------------------------------------------------------------------------------------------------------------------

def imprimir(data_impressao=None):
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

    _montar_cabecalho(section)
    _montar_rodape(section, data_impressao or data_formatada)

    largura_util = section.page_width - section.left_margin - section.right_margin
    tabela = documento.add_table(rows=1, cols=1)
    tabela.autofit = False
    tabela.columns[0].width = largura_util
    tabela.rows[0].cells[0].width = largura_util
    _remover_bordas_tabela(tabela)

    _escrever_titulo_documento(tabela, f"Intenções de {data_formatada}")
    _escrever_horario(tabela, HORARIO_MISSA)
    _escrever_espacador(tabela)

    for titulo in ORDEM_BLOCOS:
        itens = categorias.get(titulo, [])
        _escrever_bloco(tabela, titulo, itens)

    os.makedirs("output", exist_ok=True)
    caminho_saida = os.path.join("output", "Intenções.docx")
    documento.save(caminho_saida)

    print(f"\nArquivo Word '{caminho_saida}' criado com sucesso!\n")
    return caminho_saida


def imprimir_registro():

    with open ('output/Registro.txt','w') as arq:

        arq.write(f'Registro {data_formatada}:\n')

        for i in range(len(registro)):

            arq.write(f'{registro[i]}')

    print(f"\nArquivo Registro.txt criado com sucesso!\n")
 