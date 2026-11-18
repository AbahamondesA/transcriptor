@echo off
cd /d "%~dp0"
echo === Instalando Transcriptor ===
where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")
%PY% --version || (echo No se encontro Python. Instala Python 3.11 o 3.12 desde python.org marcando "Add python.exe to PATH". & pause & exit /b 1)
if exist .venv rmdir /s /q .venv
%PY% -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe transcriptor.py --menu
echo.
echo Listo. Abre la aplicacion con iniciar.bat
pause
