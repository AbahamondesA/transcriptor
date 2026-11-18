@echo off
rem Mantiene la ventana abierta pase lo que pase
if /i not "%~1"=="ejecutar" (
  cmd /k ""%~f0" ejecutar"
  exit /b
)
cd /d "%~dp0"
set "LOG=%~dp0build.log"
echo Inicio %date% %time% > "%LOG%"
echo ==========================================================
echo   Creando Transcriptor.exe y su instalador
echo   Tarda entre 5 y 20 minutos. No cierres esta ventana.
echo ==========================================================

echo [1/4] Verificando el entorno de Python de este computador...
echo PASO 1 verificar >> "%LOG%"
if not exist .venv\Scripts\python.exe call :crear_venv
call :verificar
if errorlevel 1 (
  echo     El entorno no corresponde a este computador. Reinstalando, puede tardar...
  echo PASO 1b reinstalar entorno >> "%LOG%"
  rmdir /s /q .venv >> "%LOG%" 2>&1
  call :crear_venv
  call :verificar
  if errorlevel 1 goto fallo
)
.venv\Scripts\python.exe -m pip install -q --upgrade pyinstaller >> "%LOG%" 2>&1
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [2/4] Empaquetando la aplicacion...
echo PASO 2 pyinstaller >> "%LOG%"
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --windowed --name Transcriptor --icon transcriptor.ico ^
  --collect-all faster_whisper --collect-all ctranslate2 --collect-all av --collect-all onnxruntime ^
  --collect-all tkinterdnd2 --collect-all imageio_ffmpeg --collect-all yt_dlp --collect-all yt_dlp_ejs ^
  --collect-all soundcard --collect-all sentencepiece --collect-all tokenizers --collect-all numpy ^
  --collect-data docx --collect-submodules huggingface_hub ^
  --hidden-import tkinter.colorchooser --hidden-import tkinter.font ^
  transcriptor.py >> "%LOG%" 2>&1
if errorlevel 1 goto fallo
if not exist dist\Transcriptor\Transcriptor.exe goto fallo
if exist .venv\Scripts\deno.exe copy /y .venv\Scripts\deno.exe dist\Transcriptor\ >> "%LOG%" 2>&1

echo [3/4] Buscando Inno Setup...
echo PASO 3 inno >> "%LOG%"
call :buscar_iscc
if not defined ISCC (
  echo     Instalando Inno Setup con winget...
  winget install --id JRSoftware.InnoSetup -e --silent --accept-package-agreements --accept-source-agreements >> "%LOG%" 2>&1
  call :buscar_iscc
)
if not defined ISCC goto portable

echo [4/4] Creando el instalador...
"%ISCC%" transcriptor.iss >> "%LOG%" 2>&1
if errorlevel 1 goto portable
echo.
echo LISTO: instalador\Transcriptor-Setup.exe
echo Fin OK >> "%LOG%"
explorer instalador
pause
exit /b 0

:portable
echo     No se pudo crear el instalador; se genera una version portable.
powershell -NoProfile -Command "Compress-Archive -Force -Path 'dist\Transcriptor' -DestinationPath 'Transcriptor-portable.zip'" >> "%LOG%" 2>&1
echo.
echo LISTO: dist\Transcriptor\Transcriptor.exe  y  Transcriptor-portable.zip
echo Fin portable >> "%LOG%"
explorer dist\Transcriptor
pause
exit /b 0

:fallo
echo.
echo ERROR al crear el .exe. El detalle quedo en build.log
echo Fin ERROR >> "%LOG%"
pause
exit /b 1

:verificar
echo   verificando con: >> "%LOG%"
.venv\Scripts\python.exe --version >> "%LOG%" 2>&1
.venv\Scripts\python.exe -c "import numpy, av, ctranslate2, faster_whisper, sentencepiece, onnxruntime, soundcard, tkinterdnd2, imageio_ffmpeg, yt_dlp, docx; print('entorno OK')" >> "%LOG%" 2>&1
exit /b %errorlevel%

:crear_venv
echo   creando entorno >> "%LOG%"
where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")
%PY% -m venv .venv >> "%LOG%" 2>&1
.venv\Scripts\python.exe -m pip install -q --upgrade pip >> "%LOG%" 2>&1
.venv\Scripts\python.exe -m pip install -q -r requirements.txt >> "%LOG%" 2>&1
exit /b 0

:buscar_iscc
set "ISCC="
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
exit /b 0
