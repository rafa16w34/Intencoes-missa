from docx import Document

from intencao import quebrar_linha
from config import categorias

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#FUNÇÃO DE ESCANEAR O DOCUMENTO WORD

def extrair_docx():

    def normalizar(txt):
        return " ".join(txt.strip().upper().split())

    arquivo = "input/intencoes.docx"
    doc = Document(arquivo)

    TITULOS = [
        "HONRA E LOUVOR",
        "INTENÇÃO",
        "ANIVERSÁRIO NATALÍCIO",
        "ANIVERSÁRIO DE CASAMENTO",
        "RECUPERAÇÃO DE SAÚDE",
        "AÇÃO DE GRAÇAS",
        "IRMÃOS FALECIDOS",
    ]

    TITULOS_NORM = [normalizar(t) for t in TITULOS]

    dados = {}
    titulo_atual = None

    # LER TABELAS EM VEZ DE PARÁGRAFOS
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:

                linha = cell.text.strip()
                if not linha:
                    continue

                ln = normalizar(linha)

                if ln in TITULOS_NORM:  # detecta título
                    titulo_atual = ln
                    dados[titulo_atual] = []
                    continue

                if titulo_atual:
                    
                    if (linha != "+") and not (linha.startswith("INTENÇÕES")):

                        linhas_quebradas = quebrar_linha(linha)

                        for l in linhas_quebradas:
                            dados[titulo_atual].append(l)



    return dados



dados_importados = extrair_docx()



for titulo, itens in dados_importados.items():

    for item in itens:

        # se NÃO começar com 7º, mantém na categoria original
        if not item.startswith("7º"):
            categorias[titulo].append(item)

        # se começar com 7º, manda para a categoria 7º DIA
        else:
            categorias["7º DIA"].append(item)


