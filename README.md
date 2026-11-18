# Transcriptor

Aplicación de escritorio para Windows que transcribe y traduce audio y video **en el propio computador** (el audio no se envía a ningún servicio).

## Qué hace

- **Archivos**: transcribe audios y videos de cualquier carpeta (varios a la vez, arrastrar y soltar, o clic derecho → *Enviar a → Transcriptor*).
- **Enlaces**: descarga y transcribe videos de YouTube, Vimeo, Facebook y otros sitios.
- **En vivo**: transcribe reuniones (Meet, Zoom, Teams, Webex) o cualquier audio del equipo y el micrófono, con subtítulos flotantes.
- **Visor sincronizado**: el video con la transcripción al lado, resaltando la frase actual.
- **Video con subtítulos**: genera un MP4 con el texto estampado; editor con vista previa (fuente, tamaño, colores, borde, caja, posición).
- **Traductor inglés ↔ español**: pestaña de traducción libre, traducción de toda la transcripción o línea por línea (clic derecho / Ctrl+T), y traducción en vivo.
- **Exportar**: TXT, SRT, VTT y Word.

## Instalar

Descarga `Transcriptor-Setup.exe` desde la sección **Releases** (o desde los *Artifacts* de la última ejecución en **Actions**) y ejecútalo. No requiere permisos de administrador ni tener Python instalado.

La primera vez que transcribes o traduces se descargan los modelos (voz ≈ 500 MB, cada sentido de traducción ≈ 150 MB).

## Desarrollo

```bat
instalar.bat          :: crea el entorno .venv e instala dependencias
iniciar.bat           :: abre la aplicación
diagnostico.bat       :: prueba cada componente
crear_instalador.bat  :: arma el .exe y el instalador en el propio PC
```

Cada `push` a `main` arma el instalador automáticamente en GitHub Actions (Windows, Python 3.12) y verifica el `.exe` con una autoprueba. Al crear una etiqueta `v1.0`, `v1.1`, etc., se publica una versión descargable en *Releases*.

Tecnologías: [faster-whisper](https://github.com/SYSTRAN/faster-whisper), OPUS-MT (Universidad de Helsinki) vía CTranslate2, yt-dlp, FFmpeg, PyInstaller e Inno Setup.
