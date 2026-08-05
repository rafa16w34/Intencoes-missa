from django.db import models

class IntencaoMissa(models.Model):
    TIPO_INTENCAO_CHOICES = [
        ('honra','Honra e Louvor'),
        ('niver','Aniversários'),
        ('casamento','Aniversário de Casamento'),
        ('acao_gracas', 'Ação de Graças'),
        ('intencao','Intenções'),
        ('saude', 'Recuperação de Saúde'),
        ('setimo','7° Dia'),
        ('falecido', 'Irmãos Falecidos'),
        
    ]

    
    nome_intencao = models.CharField(max_length=200, verbose_name="Nome para a Intenção")
    tipo = models.CharField(max_length=20, choices=TIPO_INTENCAO_CHOICES, default='falecido', verbose_name="Tipo de Intenção")
    data_missa = models.DateField(verbose_name="Data da Missa")
    horario_missa = models.TimeField(verbose_name="Horário da Missa")
    observacao = models.TextField(blank=True, null=True, verbose_name="Observação")
    criado_em = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Intenção de Missa"
        verbose_name_plural = "Intenções de Missa"
        ordering = ['data_missa', 'horario_missa']

    def __str__(self):
        return f"{self.nome_intencao} - {self.data_missa.strftime('%d/%m/%Y')} ({self.get_tipo_display()})"