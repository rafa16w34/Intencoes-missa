import re
import unicodedata
from datetime import date, time
from docx import Document
from .models import IntencaoMissa, Missa, Pagamento, RegistroCaixa, VALOR_UNITARIO_INTENCAO

MAPA_TITULOS = {
    "HONRA E LOUVOR": "honra",
    "PELOS ANIVERSARIANTES": "niver",
    "ANIVERSARIO NATALICIO": "niver",
    "ANIVERSARIO MATRIMONIAL DE": "casamento",
    "ANIVERSARIO DE CASAMENTO": "casamento",
    "POR INTENCAO": "intencao",
    "INTENCAO": "intencao",
    "PELA RECUPERACAO DA SAUDE DE": "saude",
    "RECUPERACAO DE SAUDE": "saude",
    "ACAO DE GRACAS": "acao_gracas",
    "PELOS NOSSOS IRMAOS FALECIDOS": "falecido",
    "IRMAOS FALECIDOS": "falecido",
    "7 DIA": "setimo",
}

PREFIXOS_IGNORADOS = ("INTENCOES", "MISSA DAS")
TAMANHO_MINIMO_ITEM = 4

MESES = {
    'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'março': 3, 'abril': 4,
    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
    'outubro': 10, 'novembro': 11, 'dezembro': 12
}


def _normalizar(txt):
    txt = txt.strip().upper().rstrip(":").strip()
    return " ".join(txt.split())


def _sem_acento(txt):
    nfkd = unicodedata.normalize("NFKD", txt)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _mapear_titulo(linha_normalizada):
    chave = _sem_acento(linha_normalizada).upper()
    chave = re.sub(r"^(\d+)\s*[ºO]?\s*DIA$", r"\1 DIA", chave)
    return MAPA_TITULOS.get(chave)


def extrair_data_e_horario_docx(doc):
    """
    Tenta identificar a data e o horário da missa a partir dos primeiros parágrafos/células do documento.
    Suporta formatos como:
      - 'Intenções de sábado, 07 de novembro de 2026'
      - 'Intenções de 04/08/2026 - 19:00'
      - 'Missa das 19h' ou 'Missa das 19h30' ou 'Missa das 19:00'
    """
    amostras = []
    for table in doc.tables[:1]:
        for row in table.rows[:5]:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    amostras.append(txt)
    for p in doc.paragraphs[:5]:
        txt = p.text.strip()
        if txt:
            amostras.append(txt)

    texto_cabecalho = "\n".join(amostras)

    data_detectada = None
    horario_detectado = None

    # 1. Procura data por extenso (ex: 07 de novembro de 2026)
    match_extenso = re.search(r'(\d{1,2})\s+de\s+([a-zA-Zç]+)\s+de\s+(\d{4})', texto_cabecalho, re.IGNORECASE)
    # 2. Procura data numérica (ex: 07/11/2026)
    match_num = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', texto_cabecalho)

    if match_extenso:
        dia = int(match_extenso.group(1))
        mes_nome = match_extenso.group(2).lower()
        ano = int(match_extenso.group(3))
        mes = MESES.get(mes_nome)
        if mes:
            try:
                data_detectada = date(ano, mes, dia)
            except ValueError:
                pass
    elif match_num:
        dia = int(match_num.group(1))
        mes = int(match_num.group(2))
        ano = int(match_num.group(3))
        try:
            data_detectada = date(ano, mes, dia)
        except ValueError:
            pass

    # Procura horário (ex: Missa das 19h, 19:00, 19h30, 07h)
    match_hora = re.search(r'(?:missa\s+das\s+)?(\d{1,2})(?:h|:)(\d{2})?', texto_cabecalho, re.IGNORECASE)
    if match_hora:
        h = int(match_hora.group(1))
        m = int(match_hora.group(2)) if match_hora.group(2) else 0
        if 0 <= h <= 23 and 0 <= m <= 59:
            horario_detectado = time(h, m)

    return data_detectada, horario_detectado


def importar_docx_para_missa(arquivo, missa_padrao=None):
    """
    Lê um arquivo .docx e:
    1. Detecta data e horário para buscar ou criar a Missa correspondente (caso não informada).
    2. Importa todas as intenções mapeadas.
    3. Cria o registro de Pagamento com forma 'secretaria' para que as intenções
       já fiquem cadastradas como PAGAS por padrão.
    Retorna uma tupla (missa_final, qtd_itens, missa_foi_criada).
    """
    doc = Document(arquivo)
    data_detectada, horario_detectado = extrair_data_e_horario_docx(doc)

    missa_final = missa_padrao
    missa_criada = False

    # Se detectou data e horário, busca ou cria a missa correspondente
    if data_detectada and horario_detectado:
        missa_final, missa_criada = Missa.objects.get_or_create(
            data=data_detectada,
            horario=horario_detectado
        )
    elif not missa_final:
        # Fallback se não detectou e nenhuma missa foi passada
        data_hoje = date.today()
        horario_padrao = time(19, 0)
        missa_final, missa_criada = Missa.objects.get_or_create(
            data=data_hoje,
            horario=horario_padrao
        )

    titulo_atual = None
    eh_sessao_categoria = False
    itens_secretaria = []
    itens_sessao = []

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if not cell.text.strip():
                    continue

                for linha in cell.text.split("\n"):
                    linha = linha.strip()
                    if not linha:
                        continue

                    linha_norm = _normalizar(linha)
                    categoria_mapeada = _mapear_titulo(linha_norm)
                    if categoria_mapeada:
                        titulo_atual = categoria_mapeada
                        eh_sessao_categoria = False  # Reset ao mudar de categoria
                        continue

                    if not titulo_atual:
                        continue

                    if linha == "+" or linha_norm.startswith(PREFIXOS_IGNORADOS):
                        continue

                    # Se a linha contiver '*', identifica o início das intenções da sessão nesta categoria
                    if "*" in linha:
                        eh_sessao_categoria = True
                        linha = linha.replace("*", "").strip()

                    if len(linha) < TAMANHO_MINIMO_ITEM:
                        continue

                    tipo_final = titulo_atual
                    if linha.upper().lstrip().startswith("7º") or linha.upper().lstrip().startswith("7O "):
                        tipo_final = "setimo"

                    novo_item = IntencaoMissa(
                        missa=missa_final,
                        nome_intencao=linha.strip().title(),
                        tipo=tipo_final,
                    )

                    if eh_sessao_categoria:
                        itens_sessao.append(novo_item)
                    else:
                        itens_secretaria.append(novo_item)

    if itens_secretaria:
        valor_secretaria = len(itens_secretaria) * VALOR_UNITARIO_INTENCAO
        pagamento_secretaria = Pagamento.objects.create(
            missa=missa_final,
            forma='secretaria',
            status='pago',
            valor_total=valor_secretaria,
            valor_recebido=valor_secretaria,
            doacao=0,
            nome_pix='Importação Secretaria Paroquial'
        )
        for item in itens_secretaria:
            item.pagamento = pagamento_secretaria
        IntencaoMissa.objects.bulk_create(itens_secretaria)

        RegistroCaixa.objects.create(
            missa=missa_final,
            texto=f"Lote prévio da SECRETARIA importado via DOCX ({len(itens_secretaria)} intenções - R$ {valor_secretaria:.2f})"
        )

    if itens_sessao:
        valor_sessao = len(itens_sessao) * VALOR_UNITARIO_INTENCAO
        pagamento_sessao = Pagamento.objects.create(
            missa=missa_final,
            forma='dinheiro',
            status='pago',
            valor_total=valor_sessao,
            valor_recebido=valor_sessao,
            doacao=0,
            nome_pix='Importação Sessão (com *)'
        )
        for item in itens_sessao:
            item.pagamento = pagamento_sessao
        IntencaoMissa.objects.bulk_create(itens_sessao)

        RegistroCaixa.objects.create(
            missa=missa_final,
            texto=f"Lote da SESSÃO identificado com asterisco (*) importado via DOCX ({len(itens_sessao)} intenções - R$ {valor_sessao:.2f})"
        )

    total_importados = len(itens_secretaria) + len(itens_sessao)
    return missa_final, total_importados, missa_criada
