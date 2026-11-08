from django import forms
from django.core.exceptions import ValidationError
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
            'nome_intencao': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Nome da pessoa ou intenção',
                'autofocus': True,
                'autocomplete': 'off'
            }),
            'tipo': forms.Select(attrs={'class': 'form-select'}),
        }


class PagamentoForm(forms.ModelForm):
    forma = forms.ChoiceField(
        label="Forma de Pagamento",
        choices=[
            ('', '-- Selecione a forma de pagamento --'),
            ('dinheiro', 'Dinheiro'),
            ('pix', 'Pix'),
        ],
        required=True,
        error_messages={
            'required': 'É obrigatório selecionar uma forma de pagamento (Dinheiro ou Pix).'
        },
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_forma'})
    )

    class Meta:
        model = Pagamento
        fields = ['forma', 'valor_recebido', 'doacao', 'nome_pix']
        widgets = {
            'valor_recebido': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.50', 'id': 'id_valor_recebido'}),
            'doacao': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.50', 'id': 'id_doacao', 'value': '0.00'}),
            'nome_pix': forms.TextInput(attrs={'class': 'form-control', 'id': 'id_nome_pix', 'placeholder': 'Nome completo de quem fez o Pix'}),
        }

    def clean_nome_pix(self):
        nome = (self.cleaned_data.get('nome_pix') or '').strip()
        if nome:
            return nome.title()
        return nome

    def clean(self):
        cleaned_data = super().clean()
        forma = cleaned_data.get('forma')
        nome_pix = (cleaned_data.get('nome_pix') or '').strip()

        if forma == 'pix' and not nome_pix:
            self.add_error('nome_pix', 'É obrigatório informar o nome do responsável pelo Pix para conferência.')

        return cleaned_data


class UploadDocxForm(forms.Form):
    arquivo = forms.FileField(
        label="Selecione o arquivo Word (.docx)",
        widget=forms.FileInput(attrs={'class': 'form-control', 'accept': '.docx'})
    )
