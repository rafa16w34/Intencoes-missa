@echo off
chcp 65001 > nul
cd /d "%~dp0.."
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)
echo ==========================================
echo    INICIANDO SISTEMA NO BANCO OFICIAL
echo    Banco de dados: db.sqlite3
echo ==========================================
set DJANGO_DB=
python manage.py runserver
pause
