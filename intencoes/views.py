from django.shortcuts import render, redirect, get_object_or_404
from django.http import FileResponse, HttpResponse
from django.contrib import messages
from django.utils import timezone
from .models import IntencaoMissa, Missa, Pagamento, RegistroCaixa, VALOR_UNITARIO_INTENCAO
from .forms import IntencaoMissaForm, MissaForm, PagamentoForm, UploadDocxForm
from .docx_generator import gerar_docx_missa
from .docx_importer import importar_docx_para_missa
from .excel_generator import gerar_excel_missa
from .pdf_generator import gerar_pdf_missa

CATEGORIAS_LITURGICAS = [
    ('honra', 'Honra e Louvor', 'bi-stars', 'bg-secondary'),
    ('niver', 'Aniversários Natalícios', 'bi-cake2', 'bg-warning text-dark'),
    ('casamento', 'Aniversário de Casamento', 'bi-heart', 'bg-primary'),
    ('acao_gracas', 'Ação de Graças', 'bi-hand-thumbs-up', 'bg-success'),
    ('intencao', 'Intenções Particulares', 'bi-bookmark-heart', 'bg-secondary'),
    ('saude', 'Recuperação de Saúde', 'bi-heart-pulse', 'bg-info text-dark'),
    ('setimo', '7° Dia', 'bi-flower1', 'bg-danger'),
    ('falecido', 'Irmãos Falecidos', 'bi-cloud-sun', 'bg-dark'),
]


def _chave_lote_sessao(missa_id):
    """Retorna a chave usada na sessão para identificar o lote de pagamento pendente desta pessoa."""
    return f'lote_pendente_missa_{missa_id}'


def _obter_ou_criar_lote(request, missa):
    """
    Retorna o objeto Pagamento (status='pendente') vinculado à rodada da pessoa atual.
    Garante que intenções de pessoas diferentes não sejam misturadas.
    """
    chave = _chave_lote_sessao(missa.id)
    lote_id = request.session.get(chave)

    lote = None
    if lote_id:
        lote = Pagamento.objects.filter(id=lote_id, missa=missa, status='pendente').first()

    if not lote:
        lote = Pagamento.objects.create(
            missa=missa,
            forma='pendente',
            status='pendente',
            valor_total=0,
            valor_recebido=0,
            doacao=0
        )
        request.session[chave] = lote.id

    return lote


def painel_missas(request):
    """Tela inicial do sistema com os cards sequenciais e timers coloridos."""
    lista_missas = Missa.objects.prefetch_related('intencoes', 'pagamentos').all()
    return render(request, 'intencoes/painel_missas.html', {
        'lista_missas': lista_missas,
    })


def detalhes_missa(request, missa_id):
    """
    Tela de gerenciamento exclusiva da missa selecionada.
    Mostra as intenções organizadas em dropdowns por categoria litúrgica,
    alerta sobre pagamentos pendentes e exibe os totais da sessão abaixo.
    """
    missa = get_object_or_404(Missa.objects.prefetch_related('intencoes__pagamento', 'pagamentos'), pk=missa_id)

    todas_intencoes = list(missa.intencoes.all())
    intencoes_por_categoria = []
    for chave, nome, icone, badge_classe in CATEGORIAS_LITURGICAS:
        itens = [it for it in todas_intencoes if it.tipo == chave]
        intencoes_por_categoria.append({
            'chave': chave,
            'nome': nome,
            'icone': icone,
            'badge_classe': badge_classe,
            'itens': itens,
            'total': len(itens),
        })

    lotes_pendentes = missa.pagamentos.filter(status='pendente').prefetch_related('intencoes_pagas')

    return render(request, 'intencoes/detalhes_missa.html', {
        'missa': missa,
        'intencoes_por_categoria': intencoes_por_categoria,
        'total_intencoes': len(todas_intencoes),
        'lotes_pendentes': lotes_pendentes,
    })


def criar_missa(request):
    """Cadastra uma nova celebração de missa."""
    if request.method == 'POST':
        form = MissaForm(request.POST)
        if form.is_valid():
            missa = form.save()
            messages.success(request, f"Missa de {missa.data.strftime('%d/%m/%Y')} às {missa.horario.strftime('%H:%M')} criada com sucesso!")
            return redirect('detalhes_missa', missa_id=missa.id)
    else:
        form = MissaForm()

    return render(request, 'intencoes/criar_missa.html', {'form': form})


def cadastrar_intencao(request, missa_id):
    """Loop de cadastro para uma mesma pessoa (lote fechado) antes de confirmar o pagamento."""
    missa = get_object_or_404(Missa, pk=missa_id)

    if missa.esta_bloqueada:
        messages.error(request, "Esta missa já encerrou há mais de 3 horas e está bloqueada para novas marcações.")
        return redirect('detalhes_missa', missa_id=missa.id)

    lote_pagamento = _obter_ou_criar_lote(request, missa)

    if request.method == 'POST':
        form = IntencaoMissaForm(request.POST)
        if form.is_valid():
            intencao = form.save(commit=False)
            intencao.missa = missa
            intencao.pagamento = lote_pagamento
            intencao.save()

            # Atualiza o total do lote daquela pessoa
            qtd_lote = lote_pagamento.intencoes_pagas.count()
            lote_pagamento.valor_total = qtd_lote * VALOR_UNITARIO_INTENCAO
            lote_pagamento.save()

            messages.success(request, f"Intenção '{intencao.nome_intencao}' adicionada à rodada.")
            return redirect('cadastrar_intencao', missa_id=missa.id)
    else:
        form = IntencaoMissaForm()

    intencoes_lote = lote_pagamento.intencoes_pagas.all()
    valor_total = lote_pagamento.valor_total

    return render(request, 'intencoes/cadastrar.html', {
        'form': form,
        'missa': missa,
        'lote': lote_pagamento,
        'intencoes_pendentes': intencoes_lote,
        'valor_total': valor_total,
        'valor_unitario': VALOR_UNITARIO_INTENCAO,
    })


def remover_intencao_pendente(request, missa_id, intencao_id):
    """Remove uma intenção da rodada atual."""
    missa = get_object_or_404(Missa, pk=missa_id)
    lote = _obter_ou_criar_lote(request, missa)

    intencao = IntencaoMissa.objects.filter(id=intencao_id, pagamento=lote).first()
    if intencao:
        intencao.delete()
        lote.valor_total = lote.intencoes_pagas.count() * VALOR_UNITARIO_INTENCAO
        lote.save()
        messages.info(request, "Intenção removida da rodada.")

    return redirect('cadastrar_intencao', missa_id=missa_id)


def salvar_lote_pendente(request, missa_id):
    """
    Suspende a rodada da pessoa atual e a salva explicitamente com PAGAMENTO PENDENTE.
    Registra na auditoria e avisa ao atendente que o valor e intenções NÃO vão para o DOCX nem para o caixa.
    """
    missa = get_object_or_404(Missa, pk=missa_id)
    chave = _chave_lote_sessao(missa_id)
    lote_id = request.session.get(chave)

    if lote_id:
        lote = Pagamento.objects.filter(id=lote_id, missa=missa, status='pendente').first()
        if lote and lote.intencoes_pagas.exists():
            nomes = ", ".join(item.nome_intencao for item in lote.intencoes_pagas.all())
            qtd = lote.intencoes_pagas.count()
            valor = lote.valor_total

            # Registra na auditoria de caixa
            RegistroCaixa.objects.create(
                missa=missa,
                texto=f"Lote #{lote.id} com {qtd} intenção(ões) marcado com PAGAMENTO PENDENTE (R$ {valor:.2f}): {nomes}"
            )

            del request.session[chave]
            messages.warning(
                request,
                f"Lote #{lote.id} ({qtd} intenções) salvo como PAGAMENTO PENDENTE. "
                f"Atenção: Enquanto o pagamento não for confirmado, elas NÃO serão impressas no Word nem somadas ao total do caixa."
            )
        elif lote:
            # Lote vazio sem nenhuma intenção: descarta
            lote.delete()
            del request.session[chave]

    return redirect('detalhes_missa', missa_id=missa_id)


def cancelar_rodada(request, missa_id):
    """Descarta a rodada da pessoa atual e exclui suas intenções não pagas."""
    missa = get_object_or_404(Missa, pk=missa_id)
    chave = _chave_lote_sessao(missa_id)
    lote_id = request.session.get(chave)

    if lote_id:
        lote = Pagamento.objects.filter(id=lote_id, missa=missa, status='pendente').first()
        if lote:
            lote.intencoes_pagas.all().delete()
            lote.delete()
        del request.session[chave]
        messages.info(request, "Rodada cancelada e intenções descartadas.")

    return redirect('detalhes_missa', missa_id=missa_id)


def pagamento(request, missa_id, pagamento_id=None):
    """
    Registra e quita o pagamento (dinheiro ou pix, troco e doação) para TODO o lote de uma pessoa.
    Permite quitar tanto a rodada ativa da sessão quanto um lote pendente aberto a partir do menu da missa.
    """
    missa = get_object_or_404(Missa, pk=missa_id)

    if missa.esta_bloqueada:
        messages.error(request, "Esta missa já encerrou há mais de 3 horas e está bloqueada para pagamentos.")
        return redirect('detalhes_missa', missa_id=missa.id)

    # Identifica o lote a pagar (pelo ID da URL ou da sessão)
    chave = _chave_lote_sessao(missa_id)
    if not pagamento_id:
        pagamento_id = request.session.get(chave)

    lote_obj = get_object_or_404(Pagamento, id=pagamento_id, missa=missa)

    intencoes_lote = lote_obj.intencoes_pagas.all()
    if not intencoes_lote.exists():
        messages.info(request, "Este lote não possui intenções cadastradas.")
        return redirect('detalhes_missa', missa_id=missa.id)

    valor_total = intencoes_lote.count() * VALOR_UNITARIO_INTENCAO
    lote_obj.valor_total = valor_total

    if request.method == 'POST':
        form = PagamentoForm(request.POST)
        if form.is_valid():
            lote_obj.forma = form.cleaned_data['forma']
            lote_obj.valor_recebido = form.cleaned_data['valor_recebido']
            lote_obj.doacao = form.cleaned_data['doacao']
            lote_obj.nome_pix = form.cleaned_data['nome_pix']
            lote_obj.status = 'pago'
            lote_obj.pago_em = timezone.now()
            lote_obj.save()

            # Registra a confirmação na auditoria do caixa
            titular_pix = f" - Titular Pix: {lote_obj.nome_pix}" if lote_obj.forma == 'pix' else ""
            doacao_str = f" (Doação: R$ {lote_obj.doacao:.2f})" if lote_obj.doacao > 0 else ""
            RegistroCaixa.objects.create(
                missa=missa,
                texto=f"Pagamento CONFIRMADO (Lote #{lote_obj.id} - R$ {lote_obj.valor_total:.2f} via {lote_obj.get_forma_display()}{titular_pix}{doacao_str})"
            )

            # Se era o lote ativo da sessão, limpa a sessão
            if request.session.get(chave) == lote_obj.id:
                del request.session[chave]

            messages.success(
                request,
                f"Pagamento de R$ {valor_total:.2f} do Lote #{lote_obj.id} confirmado com sucesso! "
                f"As intenções foram liberadas e agora serão impressas na folha oficial."
            )
            return redirect('detalhes_missa', missa_id=missa.id)
    else:
        form = PagamentoForm(initial={'valor_recebido': valor_total, 'doacao': 0.00})

    return render(request, 'intencoes/pagamento.html', {
        'form': form,
        'missa': missa,
        'lote': lote_obj,
        'intencoes_pendentes': intencoes_lote,
        'valor_total': valor_total,
    })


def excluir_lote_pendente(request, missa_id, pagamento_id):
    """Exclui definitivamente um lote pendente e todas as suas intenções vinculadas."""
    missa = get_object_or_404(Missa, pk=missa_id)
    lote = get_object_or_404(Pagamento, id=pagamento_id, missa=missa, status='pendente')

    qtd = lote.intencoes_pagas.count()
    lote.intencoes_pagas.all().delete()
    lote.delete()

    RegistroCaixa.objects.create(
        missa=missa,
        texto=f"Lote #{pagamento_id} com {qtd} intenção(ões) pendente(s) foi CANCELADO e descartado."
    )

    chave = _chave_lote_sessao(missa_id)
    if request.session.get(chave) == pagamento_id:
        del request.session[chave]

    messages.info(request, f"Lote #{pagamento_id} pendente foi descartado.")
    return redirect('detalhes_missa', missa_id=missa.id)


def editar_intencao(request, intencao_id):
    """Permite editar uma intenção já cadastrada na missa."""
    intencao = get_object_or_404(IntencaoMissa, pk=intencao_id)
    missa = intencao.missa

    if missa.esta_bloqueada:
        messages.error(request, "Esta celebração já encerrou há mais de 3 horas e não aceita edições.")
        return redirect('detalhes_missa', missa_id=missa.id)

    if request.method == 'POST':
        form = IntencaoMissaForm(request.POST, instance=intencao)
        if form.is_valid():
            form.save()
            messages.success(request, "Intenção atualizada com sucesso!")
            return redirect('detalhes_missa', missa_id=missa.id)
    else:
        form = IntencaoMissaForm(instance=intencao)

    return render(request, 'intencoes/editar_intencao.html', {
        'form': form,
        'intencao': intencao,
        'missa': missa,
    })


def excluir_intencao(request, intencao_id):
    """Exclui uma intenção da missa."""
    intencao = get_object_or_404(IntencaoMissa, pk=intencao_id)
    missa = intencao.missa

    if missa.esta_bloqueada:
        messages.error(request, "Esta celebração já encerrou há mais de 3 horas e não permite exclusões.")
        return redirect('detalhes_missa', missa_id=missa.id)

    if request.method == 'POST':
        nome = intencao.nome_intencao
        intencao.delete()
        messages.success(request, f"Intenção '{nome}' excluída com sucesso.")
        return redirect('detalhes_missa', missa_id=missa.id)

    return render(request, 'intencoes/excluir_intencao.html', {
        'intencao': intencao,
        'missa': missa,
    })


def baixar_docx(request, missa_id):
    """Gera e retorna o documento Word no padrão oficial da Mitra Diocesana (somente intenções pagas)."""
    missa = get_object_or_404(Missa, pk=missa_id)
    buffer = gerar_docx_missa(missa)

    data_str = missa.data.strftime("%d_%m_%Y")
    hora_str = missa.horario.strftime("%Hh%M")
    nome_arquivo = f"Intencoes_{data_str}_{hora_str}.docx"

    return FileResponse(
        buffer,
        as_attachment=True,
        filename=nome_arquivo,
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


def baixar_pdf(request, missa_id):
    """Gera e retorna o documento em PDF no mesmo padrão oficial do Word."""
    missa = get_object_or_404(Missa, pk=missa_id)
    try:
        buffer = gerar_pdf_missa(missa)
    except Exception as e:
        messages.error(request, f"Não foi possível converter o documento em PDF: {e}")
        return redirect('detalhes_missa', missa_id=missa.id)

    data_str = missa.data.strftime("%d_%m_%Y")
    hora_str = missa.horario.strftime("%Hh%M")
    nome_arquivo = f"Intencoes_{data_str}_{hora_str}.pdf"

    return FileResponse(
        buffer,
        as_attachment=True,
        filename=nome_arquivo,
        content_type='application/pdf'
    )


def baixar_excel(request, missa_id):
    """Gera e faz download da planilha Excel (.xlsx) da missa e fechamento de caixa."""
    missa = get_object_or_404(Missa, pk=missa_id)
    buffer = gerar_excel_missa(missa)
    data_str = missa.data.strftime("%d_%m_%Y")
    hora_str = missa.horario.strftime("%H_%M")
    nome_arquivo = f"Intencoes_{data_str}_{hora_str}.xlsx"
    return FileResponse(
        buffer,
        as_attachment=True,
        filename=nome_arquivo,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def gerar_texto_registro_caixa(missa):
    """Gera o texto padronizado do Registro de Caixa conforme modelo oficial."""
    intencoes_dinheiro = missa.intencoes.filter(
        pagamento__isnull=False, pagamento__status='pago', pagamento__forma='dinheiro'
    ).count()
    intencoes_pix = missa.intencoes.filter(
        pagamento__isnull=False, pagamento__status='pago', pagamento__forma='pix'
    ).count()
    doacoes_feitas = missa.pagamentos.filter(status='pago', doacao__gt=0).count()
    intencoes_secretaria = missa.intencoes_secretaria_count
    total_sessao_intencoes = missa.intencoes_sessao_count

    pagamentos_confirmados = missa.pagamentos.filter(status='pago').order_by('-pago_em', '-criado_em')
    pagamentos_pix = pagamentos_confirmados.filter(forma='pix')
    registros_auditoria = missa.registros_auditoria.all().order_by('criado_em')

    linhas = [
        "=== REGISTRO DE CAIXA - CATEDRAL DIVINO ESPÍRITO SANTO ===",
        f"Missa: {missa.data.strftime('%d/%m/%Y')} às {missa.horario.strftime('%H:%M')}",
        "===========================================================================",
        f"Total de Intenções Pagas: {total_sessao_intencoes}",
        f"- Numero de intenções marcadas com dinheiro: {intencoes_dinheiro}",
        f"- Numero de inteções marcadas com pix: {intencoes_pix}",
        f"- Numero de doações feitas: {doacoes_feitas}",
        f"- Numero de intenções marcadas na secretária: {intencoes_secretaria}",
        "===========================================================================",
        f"Valor Total Arrecadado (Confirmado): R$ {missa.total_sessao:.2f}",
        f" - Dinheiro em Caixa (Sessão): R$ {missa.total_dinheiro:.2f}",
        f" - Pix Confirmado (Sessão): R$ {missa.total_pix:.2f}",
        f" - Doações Recebidas: R$ {missa.total_doacao:.2f}",
        f" - Secretaria / Importado: R$ {missa.total_secretaria:.2f}",
        "===========================================================================",
        "",
        "--- DETALHAMENTO DE PAGAMENTOS PIX CONFIRMADOS ---"
    ]

    if pagamentos_pix.exists():
        for p in pagamentos_pix:
            hora = p.pago_em.strftime('%d/%m/%Y %H:%M') if p.pago_em else p.criado_em.strftime('%d/%m/%Y %H:%M')
            linhas.append(f"• Titular: {p.nome_pix or 'Não informado'} | Valor: R$ {p.valor_total:.2f} (em {hora})")
    else:
        linhas.append("Nenhum pagamento via Pix registrado.")

    linhas.append("")
    linhas.append("--- TRILHA DE AUDITORIA E MOVIMENTAÇÃO (HISTÓRICO) ---")
    if registros_auditoria.exists():
        for reg in registros_auditoria:
            linhas.append(f"[{reg.criado_em.strftime('%d/%m/%Y %H:%M:%S')}] {reg.texto}")
    else:
        linhas.append("Nenhum registro de auditoria gerado ainda.")

    linhas.append("")
    linhas.append(f"Relatório gerado pelo Sistema de Intenções em {timezone.localtime().strftime('%d/%m/%Y %H:%M:%S')}.")

    return "\n".join(linhas)


def caixa_missa(request, missa_id):
    """Exibe o fechamento de caixa financeiro da missa com histórico de auditoria completo."""
    missa = get_object_or_404(Missa, pk=missa_id)
    pagamentos_confirmados = missa.pagamentos.filter(status='pago').order_by('-pago_em', '-criado_em')
    pagamentos_pendentes = missa.pagamentos.filter(status='pendente').order_by('-criado_em')
    pagamentos_pix = pagamentos_confirmados.filter(forma='pix')
    registros_auditoria = missa.registros_auditoria.all().order_by('criado_em')

    formato = request.GET.get('exportar')
    if formato == 'xlsx':
        return redirect('baixar_excel', missa_id=missa.id)
    if formato == 'pdf':
        return redirect('baixar_pdf', missa_id=missa.id)
    if formato in ('docx', 'word'):
        return redirect('baixar_docx', missa_id=missa.id)

    if formato == 'txt':
        conteudo = gerar_texto_registro_caixa(missa)
        response = HttpResponse(conteudo, content_type='text/plain; charset=utf-8')
        response['Content-Disposition'] = f'attachment; filename="Registro_Caixa_{missa.data.strftime("%d_%m_%Y")}.txt"'
        return response

    intencoes_dinheiro = missa.intencoes.filter(
        pagamento__isnull=False, pagamento__status='pago', pagamento__forma='dinheiro'
    ).count()
    intencoes_pix = missa.intencoes.filter(
        pagamento__isnull=False, pagamento__status='pago', pagamento__forma='pix'
    ).count()
    doacoes_feitas = missa.pagamentos.filter(status='pago', doacao__gt=0).count()

    texto_registro = gerar_texto_registro_caixa(missa)

    return render(request, 'intencoes/caixa.html', {
        'missa': missa,
        'pagamentos': pagamentos_confirmados,
        'pagamentos_pendentes': pagamentos_pendentes,
        'pagamentos_pix': pagamentos_pix,
        'registros_auditoria': registros_auditoria,
        'intencoes_dinheiro': intencoes_dinheiro,
        'intencoes_pix': intencoes_pix,
        'doacoes_feitas': doacoes_feitas,
        'texto_registro': texto_registro,
    })


def importar_docx_view(request, missa_id=None):
    """Permite fazer upload de um arquivo .docx e importar intenções legadas da secretaria."""
    missa_padrao = get_object_or_404(Missa, pk=missa_id) if missa_id else None

    if missa_padrao and missa_padrao.esta_bloqueada:
        messages.error(request, "Esta celebração já encerrou há mais de 3 horas e está bloqueada para importações.")
        return redirect('detalhes_missa', missa_id=missa_padrao.id)

    if request.method == 'POST':
        form = UploadDocxForm(request.POST, request.FILES)
        if form.is_valid():
            arquivo = request.FILES['arquivo']
            try:
                missa_final, qtd, criada = importar_docx_para_missa(arquivo, missa_padrao)
                msg_missa = "criada automaticamente" if criada else "atualizada"

                RegistroCaixa.objects.create(
                    missa=missa_final,
                    texto=f"Lote prévio da SECRETARIA importado via DOCX ({qtd} intenções - R$ {qtd*VALOR_UNITARIO_INTENCAO:.2f})"
                )

                messages.success(
                    request,
                    f"Sucesso! Missa de {missa_final.data.strftime('%d/%m/%Y')} às {missa_final.horario.strftime('%H:%M')} "
                    f"{msg_missa} com {qtd} intenção(ões) importada(s) como pagas pela Secretaria!"
                )
                return redirect('detalhes_missa', missa_id=missa_final.id)
            except Exception as e:
                messages.error(request, f"Erro ao processar o arquivo Word: {e}")
    else:
        form = UploadDocxForm()

    return render(request, 'intencoes/importar_docx.html', {
        'form': form,
        'missa': missa_padrao,
    })
