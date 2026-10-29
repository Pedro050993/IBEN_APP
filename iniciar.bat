@echo off
REM Iben App: instala as dependencias na primeira vez e abre o app no navegador.
cd /d "%~dp0"
where python >nul 2>nul || (echo Python nao encontrado. Instale o Python 3.11 ou mais novo em python.org e marque "Add python.exe to PATH". & pause & exit /b 1)
if not exist ".venv\Scripts\python.exe" (
  echo Preparando o ambiente na primeira execucao, leva alguns minutos...
  python -m venv .venv || (pause & exit /b 1)
  ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt || (pause & exit /b 1)
)
".venv\Scripts\python.exe" -m streamlit run iben_app\app.py
