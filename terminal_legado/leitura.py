import re
import unicodedata

from docx import Document

from intencao import quebrar_linha
from config import categorias

# ---------------------------------------------------------------------------------------------------------------
# FUNÇÃO DE ESCANEAR O DOCUMENTO WORD
# ---------------------------------------------------------------------------------------------------------------
#
# No arquivo real da paróquia os títulos das seções NÃO usam exatamente o mesmo texto das chaves de
# `categorias` (ex.: o documento traz "PELOS ANIVERSARIANTES", enquanto a categoria interna é
# "ANIVERSÁRIO NATALÍCIO"). Além disso, algumas seções trazem o título e os nomes juntos na MESMA
# célula da tabela (ex.: "PELA RECUPERAÇÃO DA SAÚDE DE:" seguido do nome logo abaixo, sem quebra de
# célula). Por isso a extração agora funciona linha a linha (parágrafo a parágrafo) dentro de cada
# célula, em vez de célula a célula inteira.

# título como aparece de fato no .docx (sem acento, sem ":", em maiúsculas) -> chave usada em categorias
MAPA_TITULOS = {
    "HONRA E LOUVOR": "HONRA E LOUVOR",
    "PELOS ANIVERSARIANTES": "ANIVERSÁRIO NATALÍCIO",
    "ANIVERSARIO NATALICIO": "ANIVERSÁRIO NATALÍCIO",
    "ANIVERSARIO MATRIMONIAL DE": "ANIVERSÁRIO DE CASAMENTO",
    "ANIVERSARIO DE CASAMENTO": "ANIVERSÁRIO DE CASAMENTO",
    "POR INTENCAO": "INTENÇÃO",
    "INTENCAO": "INTENÇÃO",
    "PELA RECUPERACAO DA SAUDE DE": "RECUPERAÇÃO DE SAÚDE",
    "RECUPERACAO DE SAUDE": "RECUPERAÇÃO DE SAÚDE",
    "ACAO DE GRACAS": "AÇÃO DE GRAÇAS",
    "PELOS NOSSOS IRMAOS FALECIDOS": "IRMÃOS FALECIDOS",
    "IRMAOS FALECIDOS": "IRMÃOS FALECIDOS",
    "7 DIA": "7º DIA",
}

# prefixos de linha que devem ser sempre ignorados (título do documento, marcador solto "+", etc.)
PREFIXOS_IGNORADOS = ("INTENCOES", "MISSA DAS")

# fragmentos com poucos caracteres são quase sempre resíduo de digitação em linhas em branco
# (ex.: "TOT", "AA" encontrados soltos no documento original) — não são nomes válidos
TAMANHO_MINIMO_ITEM = 4


def _normalizar(txt):
    """Maiúsculas, sem espaços duplicados, sem ':' no final."""
    txt = txt.strip().upper().rstrip(":").strip()
    return " ".join(txt.split())


def _sem_acento(txt):
    nfkd = unicodedata.normalize("NFKD", txt)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _mapear_titulo(linha_normalizada):
    """Retorna a chave de `categorias` correspondente, ou None se a linha não for um título."""
    chave = _sem_acento(linha_normalizada).upper()
    # cobre "7º DIA", "7O DIA", "7º Dia", etc. (o "º" vira "o" minúsculo ao remover o acento)
    chave = re.sub(r"^(\d+)\s*[ºO]?\s*DIA$", r"\1 DIA", chave)
    return MAPA_TITULOS.get(chave)


def extrair_docx(caminho="input/intencoes.docx"):
    doc = Document(caminho)

    dados = {}
    titulo_atual = None

    # lê as tabelas (o modelo da paróquia usa uma tabela de uma coluna só, linha por linha)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if not cell.text.strip():
                    continue

                # uma célula pode conter várias linhas (parágrafos): às vezes só título,
                # às vezes só itens, e às vezes título + itens juntos na mesma célula
                for linha in cell.text.split("\n"):
                    linha = linha.strip()
                    if not linha:
                        continue

                    linha_norm = _normalizar(linha)

                    categoria_mapeada = _mapear_titulo(linha_norm)
                    if categoria_mapeada:
                        titulo_atual = categoria_mapeada
                        dados.setdefault(titulo_atual, [])
                        continue

                    if not titulo_atual:
                        continue  # ainda não achamos nenhum título válido — ignora

                    if linha == "+" or linha_norm.startswith(PREFIXOS_IGNORADOS):
                        continue

                    if len(linha) < TAMANHO_MINIMO_ITEM:
                        continue  # resíduo de digitação (ex.: "TOT", "AA")

                    for l in quebrar_linha(linha):
                        if l:
                            dados[titulo_atual].append(l)

    return dados


dados_importados = extrair_docx()

for titulo, itens in dados_importados.items():
    for item in itens:

        # se NÃO começar com 7º, mantém na categoria original
        if not item.upper().lstrip().startswith("7º") and not item.upper().lstrip().startswith("7O "):
            categorias.setdefault(titulo, []).append(item)

        # se começar com 7º, manda para a categoria 7º DIA
        else:
            categorias.setdefault("7º DIA", []).append(item)