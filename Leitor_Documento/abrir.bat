@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" echo Ambiente .venv nao encontrado. Rode: python -m venv .venv e depois pip install -r requirements.txt & pause & exit /b
set LEITOR_ESCONDER=1
".venv\Scripts\python.exe" janela.py
if errorlevel 1 echo. & echo Ocorreu um erro. Copie a mensagem acima e envie para pedir ajuda. & pause
