from django.contrib import admin
from .models import IntencaoMissa, Missa, Pagamento


class IntencaoMissaInline(admin.TabularInline):
    model = IntencaoMissa
    extra = 1


@admin.register(Missa)
class MissaAdmin(admin.ModelAdmin):
    list_display = ('data', 'horario')
    list_filter = ('data',)
    inlines = [IntencaoMissaInline]


@admin.register(IntencaoMissa)
class IntencaoMissaAdmin(admin.ModelAdmin):
    list_display = ('nome_intencao', 'tipo', 'missa', 'pagamento')
    list_filter = ('missa', 'tipo')


@admin.register(Pagamento)
class PagamentoAdmin(admin.ModelAdmin):
    list_display = ('missa', 'forma', 'valor_total', 'valor_recebido', 'troco', 'criado_em')
    list_filter = ('missa', 'forma')
