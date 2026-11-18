@echo off
cd /d "%~dp0"
echo === Actualizando Transcriptor ===
.venv\Scripts\python.exe -m pip install --upgrade -r requirements.txt
.venv\Scripts\python.exe transcriptor.py --menu
echo.
echo Listo. Clic derecho en cualquier video: Enviar a - Transcriptor
pause
