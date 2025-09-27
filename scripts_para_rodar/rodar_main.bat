@echo off
echo ===============================
echo Rodando o Monitor de Portfolio...
echo ===============================

REM ativa a venv
call venv\Scripts\activate

REM executa o main.py
python main.py

echo ===============================
echo Execucao concluida!
echo ===============================
pause
