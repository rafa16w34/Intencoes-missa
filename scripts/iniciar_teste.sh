#!/bin/bash
# Iniciar o sistema no Banco de Teste
cd "$(dirname "$0")/.."
if [ -d "venv" ]; then
    source venv/bin/activate
fi
echo "=========================================="
echo "   INICIANDO SISTEMA NO MODO DE TESTES"
echo "   Banco de dados: db_teste.sqlite3"
echo "   (Não afeta os dados oficiais da Catedral)"
echo "=========================================="
export DJANGO_DB=db_teste.sqlite3
python manage.py runserver
