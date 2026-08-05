from django.contrib import admin
from .models import IntencaoMissa

@admin.register(IntencaoMissa)
class IntencaoMissaAdmin(admin.ModelAdmin):
    list_display = ('nome_intencao', 'tipo', 'data_missa', 'horario_missa')
    list_filter = ('data_missa', 'tipo')
    #search_fields = ('nome_intencao')