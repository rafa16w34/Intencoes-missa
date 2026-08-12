from django.urls import path
from . import views

urlpatterns = [
    path('', views.listar_intencoes, name='listar_intencoes'),
    path('missas/nova/', views.criar_missa, name='criar_missa'),
    path('missas/<int:missa_id>/cadastrar/', views.cadastrar_intencao, name='cadastrar_intencao'),
    path('missas/<int:missa_id>/pagamento/', views.pagamento, name='pagamento'),
]
