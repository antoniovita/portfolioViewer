@echo off
echo ===============================
echo Criando ambiente virtual e instalando dependencias...
echo ===============================

REM cria a venv se nao existir
if not exist venv (
    python -m venv venv
)

REM ativa a venv
call venv\Scripts\activate

REM instala as bibliotecas
pip install --upgrade pip
pip install pandas yfinance openpyxl

echo ===============================
echo Instalacao concluida!
echo ===============================
pause
