@echo off
chcp 65001 > nul
cd /d "%~dp0.."
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)
echo ==========================================
echo     ZERANDO O BANCO DE DADOS DE TESTES
echo ==========================================
if exist db_teste.sqlite3 del db_teste.sqlite3
echo Recriando tabelas no db_teste.sqlite3...
set DJANGO_DB=db_teste.sqlite3
python manage.py migrate
set DJANGO_DB=
echo Concluido! O banco de testes esta zerado e pronto para novos testes.
pause
