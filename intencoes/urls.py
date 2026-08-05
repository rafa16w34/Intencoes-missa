from django.urls import path
from . import views

urlpatterns = [
    path('', views.listar_intencoes, name='listar_intencoes'),
    path('cadastrar/', views.cadastrar_intencao, name='cadastrar_intencao'),
]