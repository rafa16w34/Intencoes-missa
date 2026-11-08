from django.urls import path
from . import views

urlpatterns = [
    # Painel inicial com os cards ("quadrados em sequência")
    path('', views.painel_missas, name='painel_missas'),
    path('painel/', views.painel_missas, name='listar_intencoes'),

    # Tela exclusiva de uma celebração selecionada
    path('missas/<int:missa_id>/', views.detalhes_missa, name='detalhes_missa'),

    # Operações da missa
    path('missas/nova/', views.criar_missa, name='criar_missa'),
    path('missas/<int:missa_id>/cadastrar/', views.cadastrar_intencao, name='cadastrar_intencao'),
    path('missas/<int:missa_id>/remover-pendente/<int:intencao_id>/', views.remover_intencao_pendente, name='remover_intencao_pendente'),
    path('missas/<int:missa_id>/salvar-pendente/', views.salvar_lote_pendente, name='salvar_lote_pendente'),
    path('missas/<int:missa_id>/cancelar-rodada/', views.cancelar_rodada, name='cancelar_rodada'),
    path('missas/<int:missa_id>/pagamento/', views.pagamento, name='pagamento'),
    path('missas/<int:missa_id>/pagamento/<int:pagamento_id>/', views.pagamento, name='pagamento_lote'),
    path('missas/<int:missa_id>/excluir-pendente/<int:pagamento_id>/', views.excluir_lote_pendente, name='excluir_lote_pendente'),
    path('missas/<int:missa_id>/imprimir-docx/', views.baixar_docx, name='baixar_docx'),
    path('missas/<int:missa_id>/imprimir-pdf/', views.baixar_pdf, name='baixar_pdf'),
    path('missas/<int:missa_id>/imprimir-excel/', views.baixar_excel, name='baixar_excel'),
    path('missas/<int:missa_id>/caixa/', views.caixa_missa, name='caixa_missa'),
    path('missas/<int:missa_id>/importar-docx/', views.importar_docx_view, name='importar_docx'),
    path('importar-docx/', views.importar_docx_view, name='importar_docx_geral'),
    path('intencoes/<int:intencao_id>/editar/', views.editar_intencao, name='editar_intencao'),
    path('intencoes/<int:intencao_id>/excluir/', views.excluir_intencao, name='excluir_intencao'),
]
