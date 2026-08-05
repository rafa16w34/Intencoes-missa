from django.shortcuts import render, redirect
from .models import IntencaoMissa
from .forms import IntencaoMissaForm

# Página para cadastrar uma intenção
def cadastrar_intencao(request):
    if request.method == 'POST':
        form = IntencaoMissaForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('listar_intencoes')
    else:
        form = IntencaoMissaForm()
    
    return render(request, 'intencoes/cadastrar.html', {'form': form})

# Página para listar as intenções registradas
def listar_intencoes(request):
    intencoes = IntencaoMissa.objects.all()
    return render(request, 'intencoes/listar.html', {'intencoes': intencoes})