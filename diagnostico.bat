@echo off
cd /d "%~dp0"
echo === Diagnostico del Transcriptor ===
echo Instalando componentes que faltan para YouTube...
.venv\Scripts\python.exe -m pip install -q -r requirements.txt
.venv\Scripts\python.exe transcriptor.py --diagnostico
pause
