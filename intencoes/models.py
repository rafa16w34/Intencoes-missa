from django.db import models

VALOR_UNITARIO_INTENCAO = 3  # mesmo valor fixo por intenção usado no CLI original


class Missa(models.Model):
    data = models.DateField(verbose_name="Data da Missa")
    horario = models.TimeField(verbose_name="Horário da Missa")

    class Meta:
        verbose_name = "Missa"
        verbose_name_plural = "Missas"
        ordering = ['data', 'horario']
        unique_together = ('data', 'horario')

    def __str__(self):
        return f"{self.data.strftime('%d/%m/%Y')} - {self.horario.strftime('%H:%M')}"


class Pagamento(models.Model):
    FORMA_CHOICES = [
        ('dinheiro', 'Dinheiro'),
        ('pix', 'Pix'),
    ]

    missa = models.ForeignKey(Missa, on_delete=models.CASCADE, related_name='pagamentos')
    forma = models.CharField(max_length=10, choices=FORMA_CHOICES)
    valor_total = models.DecimalField(max_digits=8, decimal_places=2, verbose_name="Valor Total")
    valor_recebido = models.DecimalField(max_digits=8, decimal_places=2, verbose_name="Valor Recebido")
    doacao = models.DecimalField(max_digits=8, decimal_places=2, default=0, verbose_name="Doação (do troco)")
    nome_pix = models.CharField(max_length=100, blank=True, verbose_name="Nome de quem fez o Pix")
    criado_em = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Pagamento"
        verbose_name_plural = "Pagamentos"

    @property
    def troco(self):
        return self.valor_recebido - self.valor_total - self.doacao

    def __str__(self):
        return f"Pagamento de R$ {self.valor_total} - {self.get_forma_display()}"


class IntencaoMissa(models.Model):
    TIPO_INTENCAO_CHOICES = [
        ('honra', 'Honra e Louvor'),
        ('niver', 'Aniversários'),
        ('casamento', 'Aniversário de Casamento'),
        ('acao_gracas', 'Ação de Graças'),
        ('intencao', 'Intenções'),
        ('saude', 'Recuperação de Saúde'),
        ('setimo', '7° Dia'),
        ('falecido', 'Irmãos Falecidos'),
    ]

    missa = models.ForeignKey(
        Missa, on_delete=models.CASCADE, related_name='intencoes', verbose_name="Missa"
    )
    nome_intencao = models.CharField(max_length=200, verbose_name="Nome para a Intenção")
    tipo = models.CharField(
        max_length=20, choices=TIPO_INTENCAO_CHOICES, default='falecido', verbose_name="Tipo de Intenção"
    )
    pagamento = models.ForeignKey(
        Pagamento, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='intencoes_pagas', verbose_name="Pagamento"
    )
    criado_em = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Intenção de Missa"
        verbose_name_plural = "Intenções de Missa"
        ordering = ['criado_em']

    def __str__(self):
        return f"{self.nome_intencao} ({self.get_tipo_display()})"
