#!/bin/bash
# Iniciar o sistema no Banco Oficial
cd "$(dirname "$0")/.."
if [ -d "venv" ]; then
    source venv/bin/activate
fi
echo "=========================================="
echo "   INICIANDO SISTEMA NO BANCO OFICIAL"
echo "   Banco de dados: db.sqlite3"
echo "=========================================="
unset DJANGO_DB
python manage.py runserver
