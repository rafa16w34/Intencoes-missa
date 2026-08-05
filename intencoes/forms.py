from django import forms
from .models import IntencaoMissa

class IntencaoMissaForm(forms.ModelForm):
    class Meta:
        model = IntencaoMissa
        fields = ['nome_intencao', 'tipo', 'data_missa', 'horario_missa', 'observacao']
        widgets = {
            'nome_intencao': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Nome da pessoa/intenção'}),
            'tipo': forms.Select(attrs={'class': 'form-select'}),
            'data_missa': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'horario_missa': forms.TimeInput(attrs={'class': 'form-control', 'type': 'time'}),
            'observacao': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
        }