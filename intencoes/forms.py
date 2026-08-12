from django import forms
from .models import IntencaoMissa, Missa, Pagamento


class MissaForm(forms.ModelForm):
    class Meta:
        model = Missa
        fields = ['data', 'horario']
        widgets = {
            'data': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'horario': forms.TimeInput(attrs={'class': 'form-control', 'type': 'time'}),
        }


class IntencaoMissaForm(forms.ModelForm):
    class Meta:
        model = IntencaoMissa
        fields = ['nome_intencao', 'tipo']
        widgets = {
            'nome_intencao': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Nome da pessoa/intenção', 'autofocus': True}),
            'tipo': forms.Select(attrs={'class': 'form-select'}),
        }


class PagamentoForm(forms.ModelForm):
    class Meta:
        model = Pagamento
        fields = ['forma', 'valor_recebido', 'doacao', 'nome_pix']
        widgets = {
            'forma': forms.Select(attrs={'class': 'form-select'}),
            'valor_recebido': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'}),
            'doacao': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'}),
            'nome_pix': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Preencher se a forma for Pix'}),
        }
