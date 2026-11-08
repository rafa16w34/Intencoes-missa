#!/bin/bash
# Zerar apenas o banco de testes
cd "$(dirname "$0")/.."
if [ -d "venv" ]; then
    source venv/bin/activate
fi
echo "=========================================="
echo "    ZERANDO O BANCO DE DADOS DE TESTES"
echo "=========================================="
rm -f db_teste.sqlite3
echo "Recriando tabelas no db_teste.sqlite3..."
DJANGO_DB=db_teste.sqlite3 python manage.py migrate
echo "Concluído! O banco de testes está zerado e pronto para novos testes."
