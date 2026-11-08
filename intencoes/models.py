from datetime import datetime, timedelta
from django.db import models
from django.utils import timezone

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

    @property
    def data_hora(self):
        dt = datetime.combine(self.data, self.horario)
        tz = timezone.get_current_timezone()
        return timezone.make_aware(dt, tz) if timezone.is_naive(dt) else dt

    @property
    def esta_bloqueada(self):
        """Bloqueia novas intenções, edições e exclusões após 3 horas do início da celebração."""
        agora = timezone.localtime()
        return agora >= self.data_hora + timedelta(hours=3)

    @property
    def status_timer(self):
        """
        Retorna cor, badges e texto explicativo do tempo até o horário da missa:
        - Verde: Falta mais de 1 dia (> 24h)
        - Amarelo: Faltam algumas horas (< 24h e > 10 min)
        - Vermelho: Faltam 10 minutos ou menos (horário de imprimir)
        - Cinza: Horário já começou ou encerrou há mais de 3 horas
        """
        agora = timezone.localtime()
        dt = self.data_hora
        delta = dt - agora

        if agora >= dt + timedelta(hours=3):
            return {
                'cor': 'cinza',
                'badge_class': 'bg-secondary',
                'border_class': 'border-secondary',
                'card_border': 'border-secondary',
                'texto': 'Encerrada (+3h)',
                'bloqueada': True,
                'minutos_restantes': 0,
            }
        elif agora >= dt:
            diff_passou = agora - dt
            minutos = int(diff_passou.total_seconds() // 60)
            return {
                'cor': 'cinza',
                'badge_class': 'bg-dark',
                'border_class': 'border-dark',
                'card_border': 'border-dark',
                'texto': f'Iniciada há {minutos} min',
                'bloqueada': False,
                'minutos_restantes': 0,
            }
        else:
            segundos = int(delta.total_seconds())
            minutos = segundos // 60
            horas = minutos // 60
            dias = delta.days

            if minutos <= 10:
                return {
                    'cor': 'vermelho',
                    'badge_class': 'bg-danger text-white',
                    'border_class': 'border-danger border-2',
                    'card_border': 'border-danger border-2 shadow-danger',
                    'texto': f'Faltam {max(1, minutos)} min (Imprimir!)',
                    'bloqueada': False,
                    'minutos_restantes': minutos,
                }
            elif dias < 1:
                minutos_rest = minutos % 60
                return {
                    'cor': 'amarelo',
                    'badge_class': 'bg-warning text-dark',
                    'border_class': 'border-warning border-2',
                    'card_border': 'border-warning',
                    'texto': f'Faltam {horas}h {minutos_rest}m',
                    'bloqueada': False,
                    'minutos_restantes': minutos,
                }
            else:
                return {
                    'cor': 'verde',
                    'badge_class': 'bg-success text-white',
                    'border_class': 'border-success',
                    'card_border': 'border-success',
                    'texto': f'Faltam {dias} dia(s)',
                    'bloqueada': False,
                    'minutos_restantes': minutos,
                }

    @property
    def total_arrecadado(self):
        """Total consolidado de todos os pagamentos confirmados (sessão + secretaria)."""
        return sum(p.valor_total for p in self.pagamentos.filter(status='pago'))

    @property
    def total_sessao(self):
        """Total arrecadado EXCLUSIVAMENTE nesta sessão de balcão (Dinheiro + Pix confirmados)."""
        return sum(p.valor_total for p in self.pagamentos.filter(forma__in=['dinheiro', 'pix'], status='pago'))

    @property
    def intencoes_sessao_count(self):
        """Quantidade de intenções pagas na sessão de balcão."""
        return self.intencoes.filter(pagamento__forma__in=['dinheiro', 'pix'], pagamento__status='pago').count()

    @property
    def intencoes_secretaria_count(self):
        """Quantidade de intenções importadas da Secretaria e confirmadas."""
        return self.intencoes.filter(pagamento__forma='secretaria', pagamento__status='pago').count()

    @property
    def intencoes_pendentes_count(self):
        """Quantidade de intenções aguardando confirmação de pagamento."""
        return self.intencoes.filter(models.Q(pagamento__status='pendente') | models.Q(pagamento__isnull=True)).count()

    @property
    def lotes_pendentes(self):
        """Lotes de pagamentos pendentes aguardando quitação nesta missa."""
        return self.pagamentos.filter(status='pendente').prefetch_related('intencoes_pagas')

    @property
    def total_dinheiro(self):
        return sum(p.valor_total for p in self.pagamentos.filter(forma='dinheiro', status='pago'))

    @property
    def total_pix(self):
        return sum(p.valor_total for p in self.pagamentos.filter(forma='pix', status='pago'))

    @property
    def total_secretaria(self):
        return sum(p.valor_total for p in self.pagamentos.filter(forma='secretaria', status='pago'))

    @property
    def total_doacao(self):
        return sum(p.doacao for p in self.pagamentos.filter(status='pago'))


class Pagamento(models.Model):
    STATUS_CHOICES = [
        ('pago', 'Pago'),
        ('pendente', 'Pendente'),
    ]

    FORMA_CHOICES = [
        ('dinheiro', 'Dinheiro'),
        ('pix', 'Pix'),
        ('secretaria', 'Secretaria (Docx)'),
        ('pendente', 'Aguardando Pagamento'),
    ]

    missa = models.ForeignKey(Missa, on_delete=models.CASCADE, related_name='pagamentos')
    forma = models.CharField(max_length=20, choices=FORMA_CHOICES, default='pendente')
    status = models.CharField(max_length=15, choices=STATUS_CHOICES, default='pago')
    valor_total = models.DecimalField(max_digits=8, decimal_places=2, verbose_name="Valor Total")
    valor_recebido = models.DecimalField(max_digits=8, decimal_places=2, default=0, verbose_name="Valor Recebido")
    doacao = models.DecimalField(max_digits=8, decimal_places=2, default=0, verbose_name="Doação (do troco)")
    nome_pix = models.CharField(max_length=100, blank=True, verbose_name="Nome de quem fez o Pix")
    criado_em = models.DateTimeField(auto_now_add=True)
    pago_em = models.DateTimeField(null=True, blank=True, verbose_name="Data/Hora da Quitação")

    class Meta:
        verbose_name = "Pagamento"
        verbose_name_plural = "Pagamentos"

    @property
    def troco(self):
        if self.status != 'pago':
            return 0
        return self.valor_recebido - self.valor_total - self.doacao

    def save(self, *args, **kwargs):
        if self.nome_pix:
            self.nome_pix = self.nome_pix.strip().title()
        super().save(*args, **kwargs)

    def __str__(self):
        status_txt = "PAGO" if self.status == 'pago' else "PENDENTE"
        return f"Pagamento #{self.id} (R$ {self.valor_total}) - [{status_txt}]"


class RegistroCaixa(models.Model):
    missa = models.ForeignKey(Missa, on_delete=models.CASCADE, related_name='registros_auditoria')
    texto = models.TextField(verbose_name="Texto do Registro")
    criado_em = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Registro de Caixa"
        verbose_name_plural = "Registros de Caixa"
        ordering = ['criado_em']

    def __str__(self):
        return f"[{self.criado_em.strftime('%d/%m/%Y %H:%M')}] {self.texto}"


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

    def save(self, *args, **kwargs):
        if self.nome_intencao:
            self.nome_intencao = self.nome_intencao.strip().title()
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.nome_intencao} ({self.get_tipo_display()})"
