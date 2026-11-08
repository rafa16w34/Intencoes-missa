@echo off
chcp 65001 > nul
cd /d "%~dp0.."
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)
echo ==========================================
echo    INICIANDO SISTEMA NO MODO DE TESTES
echo    Banco de dados: db_teste.sqlite3
echo    (Nao afeta os dados oficiais da Catedral)
echo ==========================================
set DJANGO_DB=db_teste.sqlite3
python manage.py runserver
pause
