import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from django.utils import timezone
from .models import VALOR_UNITARIO_INTENCAO

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


def gerar_excel_missa(missa):
    """
    Gera uma planilha Excel (.xlsx) completa com duas abas:
    1. 'Intenções da Missa': Organizada em 4 colunas clássicas para visualização e impressão.
    2. 'Fechamento de Caixa': Resumo operacional de balcão, arrecadação, detalhamento Pix e auditoria.
    """
    wb = Workbook()

    # -------------------------------------------------------------------------
    # ABA 1: INTENÇÕES DA MISSA
    # -------------------------------------------------------------------------
    ws_intencoes = wb.active
    ws_intencoes.title = "Intenções da Missa"
    ws_intencoes.views.sheetView[0].showGridLines = True

    # Estilos
    fonte_titulo_principal = Font(name="Arial", size=13, bold=True, color="1F497D")
    fonte_subtitulo = Font(name="Arial", size=10, bold=True, color="333333")
    fonte_bloco_titulo = Font(name="Arial", size=11, bold=True, color="000000")
    fonte_item = Font(name="Arial", size=10, bold=False, color="000000")

    fill_cabecalho = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_bloco = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")

    borda_fina = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    borda_inferior_grossa = Border(
        bottom=Side(style="medium", color="1F497D"),
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
    )

    # Cabeçalho da planilha
    ws_intencoes.merge_cells("A1:D1")
    c1 = ws_intencoes["A1"]
    c1.value = "CATEDRAL DO DIVINO ESPÍRITO SANTO — INTENÇÕES DE MISSA"
    c1.font = fonte_titulo_principal
    c1.alignment = Alignment(horizontal="center", vertical="center")
    c1.fill = fill_cabecalho

    ws_intencoes.merge_cells("A2:D2")
    c2 = ws_intencoes["A2"]
    data_formatada = missa.data.strftime("%d/%m/%Y")
    hora_formatada = missa.horario.strftime("%H:%M")
    c2.value = f"Missa de {data_formatada} às {hora_formatada}h"
    c2.font = fonte_subtitulo
    c2.alignment = Alignment(horizontal="center", vertical="center")

    # Obter intenções pagas organizadas
    intencoes_qs = missa.intencoes.filter(
        pagamento__isnull=False,
        pagamento__status='pago'
    ).select_related('pagamento').order_by('id')

    # Agrupadas por tipo
    categorias_itens = {}
    for tipo_chave, titulo_bloco in ORDEM_BLOCOS:
        itens_cat = [it for it in intencoes_qs if it.tipo == tipo_chave]
        sec_itens = [it for it in itens_cat if it.pagamento and it.pagamento.forma == 'secretaria']
        ses_itens = [it for it in itens_cat if not (it.pagamento and it.pagamento.forma == 'secretaria')]

        nomes = [it.nome_intencao.rstrip(" *") for it in sec_itens]
        for idx, it in enumerate(ses_itens):
            n = it.nome_intencao.rstrip(" *")
            nomes.append(f"{n} *" if idx == 0 else n)
        categorias_itens[tipo_chave] = nomes

    layout = [
        (1, 4, "honra", "HONRA E LOUVOR"),
        (1, None, "niver", "ANIVERSÁRIO NATALÍCIO"),
        (1, None, "casamento", "ANIVERSÁRIO DE CASAMENTO"),
        (1, None, "acao_gracas", "AÇÃO DE GRAÇAS"),
        (2, 4, "intencao", "INTENÇÃO"),
        (2, None, "saude", "RECUPERAÇÃO DE SAÚDE"),
        (3, 4, "setimo", "7º DIA"),
    ]

    def escrever_bloco_excel(ws, col, lin_inicio, titulo, lista):
        c = ws.cell(row=lin_inicio, column=col, value=titulo)
        c.font = fonte_bloco_titulo
        c.fill = fill_bloco
        c.border = borda_inferior_grossa
        c.alignment = Alignment(horizontal="left", vertical="center")

        lin = lin_inicio + 1
        for item in lista:
            cell = ws.cell(row=lin, column=col, value=item)
            cell.font = fonte_item
            cell.border = borda_fina
            cell.alignment = Alignment(horizontal="left", vertical="center")
            lin += 1

        if not lista:
            cell = ws.cell(row=lin, column=col, value="—")
            cell.font = Font(name="Arial", size=9, italic=True, color="888888")
            cell.border = borda_fina
            cell.alignment = Alignment(horizontal="center", vertical="center")
            lin += 1

        return lin

    proxima_linha = {1: 4, 2: 4, 3: 4, 4: 4}
    for col, lin_inicio, chave, tit in layout:
        start_row = lin_inicio if lin_inicio is not None else proxima_linha[col] + 1
        lista = categorias_itens.get(chave, [])
        fim = escrever_bloco_excel(ws_intencoes, col, start_row, tit, lista)
        proxima_linha[col] = fim

    # IRMÃOS FALECIDOS (dividido em colunas 3 e 4)
    falecidos = categorias_itens.get("falecido", [])
    metade = (len(falecidos) + 1) // 2
    fale_col3 = falecidos[:metade]
    fale_col4 = falecidos[metade:]

    lin_base_falecidos = max(proxima_linha[3] + 1, 9)
    fim_col3 = escrever_bloco_excel(ws_intencoes, 3, lin_base_falecidos, "IRMÃOS FALECIDOS", fale_col3)
    fim_col4 = escrever_bloco_excel(ws_intencoes, 4, lin_base_falecidos, "IRMÃOS FALECIDOS (CONT.)", fale_col4)

    # Ajustar largura das colunas
    for col_idx in range(1, 5):
        ws_intencoes.column_dimensions[get_column_letter(col_idx)].width = 38

    # -------------------------------------------------------------------------
    # ABA 2: FECHAMENTO DE CAIXA
    # -------------------------------------------------------------------------
    ws_caixa = wb.create_sheet(title="Fechamento de Caixa")
    ws_caixa.views.sheetView[0].showGridLines = True

    ws_caixa.merge_cells("A1:D1")
    c_cx1 = ws_caixa["A1"]
    c_cx1.value = "FECHAMENTO DE CAIXA — CATEDRAL DO DIVINO ESPÍRITO SANTO"
    c_cx1.font = fonte_titulo_principal
    c_cx1.fill = fill_cabecalho
    c_cx1.alignment = Alignment(horizontal="center", vertical="center")

    ws_caixa.merge_cells("A2:D2")
    c_cx2 = ws_caixa["A2"]
    c_cx2.value = f"Missa: {data_formatada} às {hora_formatada}h | Emitido em: {timezone.localtime().strftime('%d/%m/%Y %H:%M')}"
    c_cx2.font = fonte_subtitulo
    c_cx2.alignment = Alignment(horizontal="center", vertical="center")

    # Indicadores
    intencoes_dinheiro = missa.intencoes.filter(pagamento__isnull=False, pagamento__status='pago', pagamento__forma='dinheiro').count()
    intencoes_pix = missa.intencoes.filter(pagamento__isnull=False, pagamento__status='pago', pagamento__forma='pix').count()
    doacoes_feitas = missa.pagamentos.filter(status='pago', doacao__gt=0).count()
    intencoes_sec = missa.intencoes_secretaria_count

    resumo_linhas = [
        ("RESUMO OPERACIONAL DE INTENÇÕES", ""),
        ("Total de Intenções Pagas (Sessão):", missa.intencoes_sessao_count),
        (" - Intenções marcadas com dinheiro:", intencoes_dinheiro),
        (" - Intenções marcadas com Pix:", intencoes_pix),
        (" - Doações do troco registradas:", doacoes_feitas),
        (" - Intenções importadas da secretaria:", intencoes_sec),
        ("", ""),
        ("RESUMO FINANCEIRO (ARRECADAÇÃO)", ""),
        ("Valor Total Arrecadado (Confirmado na Sessão):", f"R$ {missa.total_sessao:.2f}"),
        (" - Dinheiro em Caixa (Sessão):", f"R$ {missa.total_dinheiro:.2f}"),
        (" - Pix Confirmado (Sessão):", f"R$ {missa.total_pix:.2f}"),
        (" - Doações Recebidas:", f"R$ {missa.total_doacao:.2f}"),
        (" - Secretaria / Importado (Prévio):", f"R$ {missa.total_secretaria:.2f}"),
    ]

    curr_row = 4
    for rotulo, valor in resumo_linhas:
        if rotulo.startswith("RESUMO"):
            ws_caixa.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=2)
            cell = ws_caixa.cell(row=curr_row, column=1, value=rotulo)
            cell.font = Font(name="Arial", size=11, bold=True, color="1F497D")
            cell.fill = fill_bloco
            cell.border = borda_inferior_grossa
        elif rotulo == "":
            pass
        else:
            c_r = ws_caixa.cell(row=curr_row, column=1, value=rotulo)
            c_r.font = Font(name="Arial", size=10, bold=rotulo.startswith("Total") or rotulo.startswith("Valor"))
            c_r.border = borda_fina

            c_v = ws_caixa.cell(row=curr_row, column=2, value=valor)
            c_v.font = Font(name="Arial", size=10, bold=rotulo.startswith("Total") or rotulo.startswith("Valor"))
            c_v.border = borda_fina
            c_v.alignment = Alignment(horizontal="right")
        curr_row += 1

    # Detalhamento Pix
    curr_row += 1
    ws_caixa.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=4)
    c_p = ws_caixa.cell(row=curr_row, column=1, value="DETALHAMENTO DE PAGAMENTOS PIX CONFIRMADOS")
    c_p.font = Font(name="Arial", size=11, bold=True, color="1F497D")
    c_p.fill = fill_bloco
    curr_row += 1

    headers_pix = ["ID Pagamento", "Titular do Pix", "Horário", "Valor Confirmado"]
    for col_idx, h in enumerate(headers_pix, start=1):
        cell = ws_caixa.cell(row=curr_row, column=col_idx, value=h)
        cell.font = Font(name="Arial", size=10, bold=True)
        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        cell.border = borda_fina
    curr_row += 1

    pagamentos_pix = missa.pagamentos.filter(status='pago', forma='pix').order_by('id')
    if pagamentos_pix.exists():
        for p in pagamentos_pix:
            hora = p.pago_em.strftime('%d/%m/%Y %H:%M') if p.pago_em else p.criado_em.strftime('%d/%m/%Y %H:%M')
            ws_caixa.cell(row=curr_row, column=1, value=f"#{p.id}").border = borda_fina
            ws_caixa.cell(row=curr_row, column=2, value=p.nome_pix or 'Não informado').border = borda_fina
            ws_caixa.cell(row=curr_row, column=3, value=hora).border = borda_fina
            ws_caixa.cell(row=curr_row, column=4, value=f"R$ {p.valor_total:.2f}").border = borda_fina
            curr_row += 1
    else:
        ws_caixa.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=4)
        c_vazio = ws_caixa.cell(row=curr_row, column=1, value="Nenhum pagamento via Pix registrado.")
        c_vazio.font = Font(name="Arial", size=10, italic=True)
        c_vazio.border = borda_fina
        curr_row += 1

    # Trilha de auditoria
    curr_row += 1
    ws_caixa.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=4)
    c_a = ws_caixa.cell(row=curr_row, column=1, value="TRILHA DE AUDITORIA E MOVIMENTAÇÃO (HISTÓRICO)")
    c_a.font = Font(name="Arial", size=11, bold=True, color="1F497D")
    c_a.fill = fill_bloco
    curr_row += 1

    ws_caixa.cell(row=curr_row, column=1, value="Data/Hora").font = Font(name="Arial", size=10, bold=True)
    ws_caixa.cell(row=curr_row, column=1).border = borda_fina
    ws_caixa.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)
    ws_caixa.cell(row=curr_row, column=2, value="Descrição da Operação").font = Font(name="Arial", size=10, bold=True)
    ws_caixa.cell(row=curr_row, column=2).border = borda_fina
    curr_row += 1

    registros = missa.registros_auditoria.all().order_by('criado_em')
    if registros.exists():
        for reg in registros:
            c_dh = ws_caixa.cell(row=curr_row, column=1, value=reg.criado_em.strftime('%d/%m/%Y %H:%M:%S'))
            c_dh.border = borda_fina
            c_dh.font = Font(name="Arial", size=9)
            ws_caixa.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)
            c_tx = ws_caixa.cell(row=curr_row, column=2, value=reg.texto)
            c_tx.border = borda_fina
            c_tx.font = Font(name="Arial", size=9)
            curr_row += 1
    else:
        ws_caixa.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=4)
        c_v = ws_caixa.cell(row=curr_row, column=1, value="Nenhum registro de auditoria gerado ainda.")
        c_v.font = Font(name="Arial", size=10, italic=True)
        c_v.border = borda_fina
        curr_row += 1

    ws_caixa.column_dimensions["A"].width = 25
    ws_caixa.column_dimensions["B"].width = 35
    ws_caixa.column_dimensions["C"].width = 25
    ws_caixa.column_dimensions["D"].width = 20

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
