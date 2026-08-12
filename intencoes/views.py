from django.shortcuts import render, redirect, get_object_or_404
from .models import IntencaoMissa, Missa, Pagamento, VALOR_UNITARIO_INTENCAO
from .forms import IntencaoMissaForm, MissaForm, PagamentoForm


def _chave_sessao(missa_id):
    """Chave usada na sessão para lembrar quais intenções foram cadastradas
    nesta rodada, para esta missa, até a pessoa ir para o pagamento."""
    return f'intencoes_pendentes_missa_{missa_id}'


# Página para criar uma nova missa (data + horário)
def criar_missa(request):
    if request.method == 'POST':
        form = MissaForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('listar_intencoes')
    else:
        form = MissaForm()

    return render(request, 'intencoes/criar_missa.html', {'form': form})


# Loop de cadastro: vai salvando intenção por intenção na MESMA missa,
# até a pessoa clicar em "Finalizar e ir para pagamento"
def cadastrar_intencao(request, missa_id):
    missa = get_object_or_404(Missa, pk=missa_id)
    chave = _chave_sessao(missa_id)
    ids_pendentes = request.session.get(chave, [])

    if request.method == 'POST':
        form = IntencaoMissaForm(request.POST)
        if form.is_valid():
            intencao = form.save(commit=False)
            intencao.missa = missa
            intencao.save()

            ids_pendentes.append(intencao.id)
            request.session[chave] = ids_pendentes

            # volta para a MESMA página (efeito "loop"), pronta para a próxima intenção
            return redirect('cadastrar_intencao', missa_id=missa.id)
    else:
        form = IntencaoMissaForm()

    intencoes_pendentes = IntencaoMissa.objects.filter(id__in=ids_pendentes)
    valor_total = len(ids_pendentes) * VALOR_UNITARIO_INTENCAO

    return render(request, 'intencoes/cadastrar.html', {
        'form': form,
        'missa': missa,
        'intencoes_pendentes': intencoes_pendentes,
        'valor_total': valor_total,
        'valor_unitario': VALOR_UNITARIO_INTENCAO,
    })


# Tela final: mostra o valor total (nº de intenções da rodada x valor unitário)
# e registra a forma de pagamento
def pagamento(request, missa_id):

    missa = get_object_or_404(Missa, pk=missa_id)
    chave = _chave_sessao(missa_id)
    ids_pendentes = request.session.get(chave, [])

    if not ids_pendentes:
        # ninguém cadastrou nada nesta rodada ainda — não faz sentido pagar
        return redirect('cadastrar_intencao', missa_id=missa.id)

    intencoes_pendentes = IntencaoMissa.objects.filter(id__in=ids_pendentes)
    valor_total = len(ids_pendentes) * VALOR_UNITARIO_INTENCAO

    if request.method == 'POST':
        
        form = PagamentoForm(request.POST)
        if form.is_valid():
            pagamento_obj = form.save(commit=False)
            pagamento_obj.missa = missa
            pagamento_obj.valor_total = valor_total
            pagamento_obj.save()

            intencoes_pendentes.update(pagamento=pagamento_obj)

            del request.session[chave]  # limpa a rodada
            return redirect('listar_intencoes')
    else:
        form = PagamentoForm(initial={'valor_recebido': valor_total})

    return render(request, 'intencoes/pagamento.html', {
        'form': form,
        'missa': missa,
        'intencoes_pendentes': intencoes_pendentes,
        'valor_total': valor_total,
    })


def listar_intencoes(request):
    lista_missas = Missa.objects.prefetch_related('intencoes').all()
    return render(request, 'intencoes/listar.html', {'lista_missas': lista_missas})
