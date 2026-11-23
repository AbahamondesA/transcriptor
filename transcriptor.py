#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Transcriptor — transcribe archivos de audio/video, enlaces (YouTube y otros)
y reuniones en vivo (Meet, Zoom, Teams, Webex o cualquier audio del equipo),
mostrando el texto a medida que se produce. Todo se procesa localmente con
Whisper (faster-whisper); el audio no sale del computador.
"""
import os
import re
import sys
import json
import html
import time
import queue
import datetime
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import numpy as np

APP = "Transcriptor"
CONGELADO = getattr(sys, "frozen", False)   # True cuando corre como Transcriptor.exe
SR = 16000          # frecuencia que usa Whisper
CAP_SR = 48000      # frecuencia de captura (se reduce a 16 kHz)
OUT_DIR = os.path.join(os.path.expanduser("~"), "Documents", "Transcriptor")
os.makedirs(OUT_DIR, exist_ok=True)
LOG = os.path.join(OUT_DIR, "registro.log")

import logging
import traceback
logging.basicConfig(filename=LOG, level=logging.INFO, encoding="utf-8",
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(APP)
sys.excepthook = lambda t, v, tb: log.error("".join(traceback.format_exception(t, v, tb)))

# Programas instalados junto a Python (p. ej. deno, que YouTube exige a yt-dlp)
_scripts = os.path.dirname(sys.executable)
os.environ["PATH"] = _scripts + os.pathsep + os.environ.get("PATH", "")

MODELOS = ["tiny", "base", "small", "medium", "large-v3-turbo", "large-v3"]
IDIOMAS = [("Español", "es"), ("Detectar automáticamente", None), ("Inglés", "en"),
           ("Portugués", "pt"), ("Francés", "fr"), ("Italiano", "it"), ("Alemán", "de")]

# Frases que Whisper suele "inventar" sobre silencio o música
EXT_MEDIA = (".mp4", ".mkv", ".mov", ".avi", ".webm", ".m4v", ".wmv", ".flv", ".mpg", ".mpeg",
             ".3gp", ".ts", ".mts", ".mp3", ".wav", ".m4a", ".ogg", ".opus", ".flac", ".aac", ".wma", ".amr")

ALUCINACIONES = ("amara.org", "subtítulos realizados por", "subtitulado por",
                 "subtítulos por la comunidad", "gracias por ver", "suscríbete",
                 "thanks for watching", "thank you for watching", "¡suscríbete")


# ───────────────────────── utilidades ─────────────────────────
def ts(sec, sep=","):
    sec = max(0.0, float(sec))
    ms = int(round(sec * 1000))
    h, ms = divmod(ms, 3_600_000)
    m, ms = divmod(ms, 60_000)
    s, ms = divmod(ms, 1000)
    return f"{h:02d}:{m:02d}:{s:02d}{sep}{ms:03d}"


def ts_corto(sec):
    return ts(sec)[:8]


def es_alucinacion(texto):
    t = texto.lower().strip()
    if not t or len(t) > 90:
        return False
    return any(a in t for a in ALUCINACIONES)


def etiqueta(seg):
    return seg.get("wall") or ts_corto(seg["start"])


def _com_init():
    """soundcard en Windows necesita COM inicializado en cada hilo."""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.ole32.CoInitializeEx(None, 0)
        except Exception:
            pass


# ───────────────────────── exportación ─────────────────────────
def a_txt(segs, titulo=""):
    out = [titulo, "=" * len(titulo), ""] if titulo else []
    for s in segs:
        quien = f"{s['src']}: " if s.get("src") else ""
        out.append(f"[{etiqueta(s)}] {quien}{s['text']}")
    return "\n".join(out) + "\n"


def a_srt(segs):
    out = []
    for i, s in enumerate(segs, 1):
        quien = f"{s['src']}: " if s.get("src") else ""
        out.append(f"{i}\n{ts(s['start'])} --> {ts(s['end'])}\n{quien}{s['text']}\n")
    return "\n".join(out)


def a_vtt(segs):
    out = ["WEBVTT", ""]
    for s in segs:
        quien = f"<v {s['src']}>" if s.get("src") else ""
        out.append(f"{ts(s['start'], '.')} --> {ts(s['end'], '.')}\n{quien}{s['text']}\n")
    return "\n".join(out)


def a_docx(segs, ruta, titulo):
    from docx import Document
    from docx.shared import Pt, RGBColor
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    doc.add_heading(titulo or "Transcripción", level=1)
    p = doc.add_paragraph()
    r = p.add_run(f"Generado el {datetime.datetime.now():%d-%m-%Y %H:%M} · {len(segs)} segmentos")
    r.italic = True
    r.font.size = Pt(9)
    for s in segs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(f"[{etiqueta(s)}] ")
        r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        r.font.size = Pt(9)
        if s.get("src"):
            r = p.add_run(f"{s['src']}: ")
            r.bold = True
        p.add_run(s["text"])
    doc.save(ruta)


VISOR_HTML = r"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<title>__TITULO__</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box}
body{margin:0;font-family:"Segoe UI",system-ui,sans-serif;background:#f4f5f7;color:#1d1f23;height:100vh;display:flex;flex-direction:column}
header{padding:10px 18px;background:#1d2a44;color:#fff;display:flex;gap:14px;align-items:center;flex-wrap:wrap}
header h1{font-size:16px;margin:0;font-weight:600;flex:1;min-width:200px}
header input{padding:6px 10px;border-radius:6px;border:0;width:220px}
header label{font-size:13px;display:flex;gap:6px;align-items:center}
header select{border-radius:6px;padding:4px}
main{flex:1;display:flex;min-height:0}
#izq{flex:1.3;display:flex;flex-direction:column;padding:14px;min-width:0}
video{width:100%;max-height:70vh;background:#000;border-radius:8px}
#sub{margin-top:12px;font-size:22px;line-height:1.35;min-height:3em;background:#fff;border-radius:8px;padding:12px 16px;box-shadow:0 1px 3px rgba(0,0,0,.08)}
#der{flex:1;overflow-y:auto;padding:10px 14px 40px;border-left:1px solid #dfe2e6;background:#fff}
.l{padding:7px 10px;border-radius:6px;cursor:pointer;display:flex;gap:10px;line-height:1.4}
.l:hover{background:#eef2f8}
.l .t{color:#7a8190;font-size:12px;font-variant-numeric:tabular-nums;padding-top:2px;flex:none}
.l.act{background:#fff3c4}
.l.oculta{display:none}
mark{background:#ffd54a}
@media (max-width:800px){main{flex-direction:column}#der{border-left:0}}
</style></head><body>
<header><h1>__TITULO__</h1>
<input id="q" placeholder="Buscar en la transcripción…">
<label><input type="checkbox" id="seguir" checked> Seguir el video</label>
<label>Velocidad <select id="vel"><option>0.75</option><option selected>1</option><option>1.25</option><option>1.5</option><option>2</option></select></label>
</header>
<main><section id="izq"><video id="v" controls preload="metadata" src="__MEDIA__"></video><div id="sub"></div></section>
<section id="der"></section></main>
<script>
const S = __SEGMENTOS__;
const v=document.getElementById('v'), der=document.getElementById('der'), sub=document.getElementById('sub');
const fmt=t=>{t=Math.max(0,t|0);const h=(t/3600)|0,m=((t%3600)/60)|0,s=t%60;return (h?h+':':'')+String(m).padStart(h?2:1,'0')+':'+String(s).padStart(2,'0')};
const esc=x=>x.replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
S.forEach((s,i)=>{const d=document.createElement('div');d.className='l';d.dataset.i=i;
 d.innerHTML='<span class="t">'+fmt(s.start)+'</span><span class="x">'+(s.src?'<b>'+esc(s.src)+':</b> ':'')+esc(s.text)+'</span>';
 d.onclick=()=>{v.currentTime=s.start+0.01;v.play()};der.appendChild(d)});
const filas=[...der.children];let act=-1;
function buscar(t){let lo=0,hi=S.length-1,r=-1;while(lo<=hi){const m=(lo+hi)>>1;if(S[m].start<=t){r=m;lo=m+1}else hi=m-1}return r}
v.addEventListener('timeupdate',()=>{const i=buscar(v.currentTime);if(i===act)return;
 if(act>=0)filas[act].classList.remove('act');act=i;
 if(i>=0){filas[i].classList.add('act');sub.textContent=S[i].text;
  if(document.getElementById('seguir').checked)filas[i].scrollIntoView({block:'center',behavior:'smooth'})}else sub.textContent=''});
document.getElementById('vel').onchange=e=>v.playbackRate=+e.target.value;
document.getElementById('q').oninput=e=>{const q=e.target.value.trim().toLowerCase();
 filas.forEach((f,i)=>{const s=S[i];const x=f.querySelector('.x');const base=(s.src?'<b>'+esc(s.src)+':</b> ':'');
  if(!q){f.classList.remove('oculta');x.innerHTML=base+esc(s.text);return}
  const ok=s.text.toLowerCase().includes(q);f.classList.toggle('oculta',!ok);
  if(ok){const re=new RegExp('('+q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi');x.innerHTML=base+esc(s.text).replace(re,'<mark>$1</mark>')}})};
</script></body></html>"""


def a_visor(segs, media, ruta_html, titulo):
    rel = os.path.relpath(media, os.path.dirname(ruta_html)).replace("\\", "/")
    from urllib.parse import quote
    datos = [{"start": round(s["start"], 2), "end": round(s["end"], 2),
              "text": s["text"], "src": s.get("src", "")} for s in segs]
    pagina = (VISOR_HTML.replace("__TITULO__", html.escape(titulo))
              .replace("__MEDIA__", quote(rel))
              .replace("__SEGMENTOS__", json.dumps(datos, ensure_ascii=False).replace("</", "<\\/")))
    with open(ruta_html, "w", encoding="utf-8") as f:
        f.write(pagina)


def segs_segun(segs, vista):
    """Original, Traducción o Ambos (texto y traducción en dos líneas)."""
    out = []
    for s in segs:
        tr = s.get("trad")
        if vista == "Traducción" and tr:
            out.append(dict(s, text=tr))
        elif vista == "Ambos" and tr:
            out.append(dict(s, text=s["text"] + "\n" + tr))
        else:
            out.append(s)
    return out


# ───────────────────────── lectura de audio ─────────────────────────
def decodificar(ruta):
    """Extrae el audio de cualquier archivo como mono 16 kHz float32.
    Se hace aquí (y no dentro de faster-whisper) para funcionar con
    cualquier versión de PyAV."""
    import av
    partes = []
    with av.open(ruta, mode="r") as cont:
        if not cont.streams.audio:
            raise ValueError("El archivo no tiene pista de audio.")
        pista = cont.streams.audio[0]
        res = av.audio.resampler.AudioResampler(format="s16", layout="mono", rate=SR)

        def agregar(frames):
            for f in frames or []:
                partes.append(f.to_ndarray().reshape(-1))

        for paquete in cont.demux(pista):
            try:
                for frame in paquete.decode():
                    agregar(res.resample(frame))
            except av.error.InvalidDataError:
                continue        # salta trozos dañados
        agregar(res.resample(None))
    if not partes:
        raise ValueError("No se pudo leer audio del archivo.")
    return np.concatenate(partes).astype(np.float32) / 32768.0


# ───────────────────────── motor Whisper ─────────────────────────
class Motor:
    def __init__(self):
        self.model = None
        self.nombre = None
        self.lock = threading.Lock()

    def cargar(self, nombre, log):
        with self.lock:
            if self.model is not None and self.nombre == nombre:
                return
            from faster_whisper import WhisperModel
            log(f"Cargando modelo «{nombre}»… (la primera vez se descarga; puede tardar)")
            modelo = None
            try:
                import ctranslate2
                if ctranslate2.get_cuda_device_count() > 0:
                    modelo = WhisperModel(nombre, device="cuda", compute_type="float16")
                    list(modelo.transcribe(np.zeros(SR, np.float32), language="es")[0])
                    log("Usando tarjeta gráfica (GPU).")
            except Exception:
                modelo = None
            if modelo is None:
                modelo = WhisperModel(nombre, device="cpu", compute_type="int8",
                                      cpu_threads=max(2, (os.cpu_count() or 4) - 1))
                log("Usando procesador (CPU).")
            self.model, self.nombre = modelo, nombre

    def archivo(self, ruta, idioma):
        audio = decodificar(ruta)
        return self.model.transcribe(audio, language=idioma, beam_size=5, vad_filter=True,
                                     condition_on_previous_text=False)

    def fragmento(self, audio, idioma, contexto=None, beam=1):
        segs, _ = self.model.transcribe(
            audio, language=idioma, beam_size=beam, vad_filter=False,
            condition_on_previous_text=False, without_timestamps=True,
            initial_prompt=contexto or None, no_speech_threshold=0.6)
        partes = []
        for s in segs:
            if s.no_speech_prob > 0.7 and s.avg_logprob < -0.8:
                continue
            if s.text.strip():
                partes.append(s.text.strip())
        texto = " ".join(partes).strip()
        return "" if es_alucinacion(texto) else texto


# ───────────────────────── traductor (local) ─────────────────────────
DIRECCIONES = {"Inglés → Español": "en-es", "Español → Inglés": "es-en", "Detectar automáticamente": "auto"}
MODELOS_TRAD = {  # OPUS-MT (Universidad de Helsinki) convertidos a CTranslate2
    "en-es": ["michaelfeil/ct2fast-opus-mt-en-es", "gaudi/opus-mt-en-es-ctranslate2"],
    "es-en": ["michaelfeil/ct2fast-opus-mt-es-en", "gaudi/opus-mt-es-en-ctranslate2"],
}
_PAL_ES = set("el la los las de del que y en un una por con para es no se su al lo como más pero sus le ya o este "
              "sí porque esta entre cuando muy sin sobre también me hasta hay donde quien desde todo nos durante "
              "todos uno les ni contra otros ese eso ante ellos esto mí antes algunos qué unos yo otro otras otra "
              "él tanto esa estos mucho quienes nada muchos cual poco ella estar estas algunas algo nosotros".split())
_PAL_EN = set("the be to of and a in that have i it for not on with he as you do at this but his by from they we "
              "say her she or an will my one all would there their what so up out if about who get which go me "
              "when make can like time no just him know take people into year your good some could them see "
              "other than then now look only come its over think also back after use two how our work is are was".split())


def detectar_idioma(texto):
    pal = re.findall(r"[a-záéíóúñü']+", texto.lower())
    es = sum(p in _PAL_ES for p in pal) + 2 * len(re.findall(r"[ñáéíóú¿¡]", texto.lower()))
    en = sum(p in _PAL_EN for p in pal)
    return "es" if es >= en else "en"


def _frases(texto):
    """Divide en frases para traducir textos largos sin cortar ideas."""
    partes = []
    for parrafo in texto.split("\n"):
        if not parrafo.strip():
            partes.append(("\n", None))
            continue
        for f in re.split(r"(?<=[.!?…])\s+", parrafo.strip()):
            if f:
                partes.append((f, True))
        partes.append(("\n", None))
    return partes


class Traductor:
    def __init__(self):
        self.modelos = {}
        self.lock = threading.Lock()

    def _cargar(self, par, log=lambda m: None):
        with self.lock:
            if par in self.modelos:
                return self.modelos[par]
            import ctranslate2
            import sentencepiece as spm
            from huggingface_hub import snapshot_download
            ultimo = None
            for repo in MODELOS_TRAD[par]:
                try:
                    log(f"Preparando traductor {par} (la primera vez se descarga ~150 MB)…")
                    ruta = snapshot_download(repo, allow_patterns=["*.bin", "*.json", "*.txt", "*.spm"])
                    tr = ctranslate2.Translator(ruta, device="cpu", compute_type="int8",
                                                inter_threads=1, intra_threads=max(2, (os.cpu_count() or 4) // 2))
                    src = spm.SentencePieceProcessor(model_file=os.path.join(ruta, "source.spm"))
                    tgt = spm.SentencePieceProcessor(model_file=os.path.join(ruta, "target.spm"))
                    self.modelos[par] = (tr, src, tgt)
                    log("Traductor listo.")
                    return self.modelos[par]
                except Exception as e:
                    log_err = f"{repo}: {e}"
                    logging.getLogger(APP).warning("traductor " + log_err)
                    ultimo = e
            raise RuntimeError(f"No se pudo preparar el traductor ({ultimo}).")

    def _lote(self, frases, par, log):
        tr, src, tgt = self._cargar(par, log)
        res = []
        for i in range(0, len(frases), 24):
            trozo = frases[i:i + 24]
            tokens = [src.encode(f, out_type=str)[:500] + ["</s>"] for f in trozo]
            out = tr.translate_batch(tokens, beam_size=4, max_decoding_length=512)
            res += [tgt.decode(o.hypotheses[0]).replace("▁", " ").strip() for o in out]
        return res

    def traducir_lista(self, textos, direccion, log=lambda m: None):
        """Traduce cada texto de la lista; con 'auto' decide el sentido texto por texto."""
        pares = [(detectar_idioma(t) + "-" + ("en" if detectar_idioma(t) == "es" else "es"))
                 if direccion == "auto" else direccion for t in textos]
        salida = [""] * len(textos)
        for par in set(pares):
            idx = [i for i, p in enumerate(pares) if p == par and textos[i].strip()]
            if idx:
                for i, t in zip(idx, self._lote([textos[i] for i in idx], par, log)):
                    salida[i] = t
        return salida

    def traducir_texto(self, texto, direccion, log=lambda m: None):
        partes = _frases(texto)
        frases = [p for p, es_frase in partes if es_frase]
        if direccion == "auto":
            direccion = "es-en" if detectar_idioma(texto) == "es" else "en-es"
        trad = iter(self._lote(frases, direccion, log)) if frases else iter([])
        out, linea = [], []
        for p, es_frase in partes:
            if es_frase:
                linea.append(next(trad))
            else:
                out.append(" ".join(linea))
                linea = []
        return "\n".join(out).strip("\n"), direccion


# ───────────────────────── captura en vivo ─────────────────────────
class Canal:
    def __init__(self, nombre, mic):
        self.nombre, self.mic = nombre, mic
        self.q = queue.Queue()
        self.buf = np.zeros(0, np.float32)
        self.t0 = None              # hora de inicio del búfer
        self.ult_llegada = 0.0
        self.ult_voz = 0.0
        self.hay_voz = False
        self.ult_parcial = 0.0
        self.ruido = 0.003


class Captura:
    MAX_SEG = 12.0      # corta frases largas
    SILENCIO = 0.7      # pausa que cierra una frase

    def __init__(self, app, canales, vista_previa, idioma):
        self.app, self.canales = app, canales
        self.vista_previa, self.idioma = vista_previa, idioma
        self.stop = threading.Event()
        self.t_inicio = time.time()
        self.contexto = ""

    def iniciar(self):
        for c in self.canales:
            threading.Thread(target=self._grabar, args=(c,), daemon=True).start()
        threading.Thread(target=self._procesar, daemon=True).start()

    def detener(self):
        self.stop.set()

    def _grabar(self, c):
        _com_init()
        try:
            with c.mic.recorder(samplerate=CAP_SR, blocksize=2400) as r:
                while not self.stop.is_set():
                    d = r.record(numframes=12000)           # 0,25 s
                    if d.ndim > 1:
                        d = d.mean(axis=1)
                    n = len(d) // 3 * 3
                    if n:
                        c.q.put((time.time(), d[:n].reshape(-1, 3).mean(axis=1).astype(np.float32)))
        except Exception as e:
            self.app.emit(("error", f"No se pudo capturar «{c.nombre}»: {e}"))

    def _cerrar(self, c, final=True):
        audio = c.buf
        if len(audio) < SR * 0.4:
            return ""
        beam = 3 if final else 1
        texto = self.app.motor.fragmento(audio, self.idioma, self.contexto[-200:] if final else None, beam)
        return texto

    def _procesar(self):
        try:
            self._procesar_bucle()
        except Exception as e:
            log.exception("procesar")
            self.app.emit(("error", f"Falló la transcripción en vivo:\n{e}"))

    def _procesar_bucle(self):
        while not self.stop.is_set():
            ahora = time.time()
            for c in self.canales:
                while True:
                    try:
                        t, d = c.q.get_nowait()
                    except queue.Empty:
                        break
                    if c.t0 is None or len(c.buf) == 0:
                        c.t0 = t - len(d) / SR
                    c.buf = np.concatenate([c.buf, d])
                    c.ult_llegada = t
                    rms = float(np.sqrt(np.mean(d * d))) if len(d) else 0.0
                    if rms < c.ruido:
                        c.ruido = rms * 0.9 + c.ruido * 0.1
                    else:
                        c.ruido = c.ruido * 0.998 + rms * 0.002
                    if rms > max(0.004, c.ruido * 3):
                        c.ult_voz, c.hay_voz = t, True

                dur = len(c.buf) / SR
                if dur == 0:
                    continue
                silencio = ahora - c.ult_voz
                if (c.hay_voz and silencio >= self.SILENCIO and dur >= 1.0) or dur >= self.MAX_SEG:
                    texto = self._cerrar(c, final=True) if c.hay_voz else ""
                    inicio = c.t0 - self.t_inicio
                    if texto:
                        self.contexto = (self.contexto + " " + texto)[-400:]
                        seg = {"start": inicio, "end": inicio + dur, "text": texto, "src": c.nombre,
                               "wall": datetime.datetime.fromtimestamp(c.t0).strftime("%H:%M:%S")}
                        if self.app.trad_vivo.get():
                            try:
                                seg["trad"] = self.app.traductor.traducir_lista(
                                    [texto], DIRECCIONES[self.app.dir_vivo.get()], self.app._log)[0]
                            except Exception as e:
                                log.exception("traducción en vivo")
                                self.app._log(f"Traducción en vivo no disponible: {e}")
                        self.app.emit(("seg", seg))
                    self.app.emit(("parcial", c.nombre, ""))
                    c.buf = np.zeros(0, np.float32)
                    c.t0, c.hay_voz = None, False
                elif not c.hay_voz and dur > 3:
                    c.buf = c.buf[-int(SR * 0.5):]
                    c.t0 = ahora - len(c.buf) / SR
                elif (self.vista_previa and c.hay_voz and dur >= 1.2
                      and ahora - c.ult_parcial >= 1.5):
                    c.ult_parcial = ahora
                    texto = self._cerrar(c, final=False)
                    if texto:
                        self.app.emit(("parcial", c.nombre, texto))
            time.sleep(0.1)


# ───────────────────────── subtítulos flotantes ─────────────────────────
class Flotante(tk.Toplevel):
    def __init__(self, master, al_cerrar):
        super().__init__(master)
        self.al_cerrar = al_cerrar
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        try:
            self.attributes("-alpha", 0.88)
        except tk.TclError:
            pass
        self.configure(bg="#111")
        ancho = min(1100, self.winfo_screenwidth() - 80)
        x = (self.winfo_screenwidth() - ancho) // 2
        y = self.winfo_screenheight() - 210
        self.geometry(f"{ancho}x130+{x}+{y}")
        self.lbl = tk.Label(self, text="Esperando voz…", fg="white", bg="#111",
                            font=("Segoe UI", 17), wraplength=ancho - 30, justify="left", anchor="sw")
        self.lbl.pack(fill="both", expand=True, padx=14, pady=(8, 2))
        self.par = tk.Label(self, text="", fg="#b8c4d6", bg="#111", font=("Segoe UI", 14, "italic"),
                            wraplength=ancho - 30, justify="left", anchor="w")
        self.par.pack(fill="x", padx=14, pady=(0, 8))
        tk.Label(self, text="arrastra para mover · doble clic: tamaño · clic derecho: cerrar",
                 fg="#666", bg="#111", font=("Segoe UI", 8)).place(relx=1, rely=0, anchor="ne", x=-6, y=2)
        for w in (self, self.lbl, self.par):
            w.bind("<ButtonPress-1>", self._ini)
            w.bind("<B1-Motion>", self._mover)
            w.bind("<Button-3>", lambda e: self.cerrar())
            w.bind("<Double-Button-1>", self._tam)
        self.grande = False

    def _ini(self, e):
        self._dx, self._dy = e.x_root - self.winfo_x(), e.y_root - self.winfo_y()

    def _mover(self, e):
        self.geometry(f"+{e.x_root - self._dx}+{e.y_root - self._dy}")

    def _tam(self, e):
        self.grande = not self.grande
        alto = 260 if self.grande else 130
        self.geometry(f"{self.winfo_width()}x{alto}")

    def mostrar(self, lineas, parcial):
        self.lbl.config(text="\n".join(lineas[-(4 if self.grande else 2):]) or "Esperando voz…")
        self.par.config(text=parcial)

    def cerrar(self):
        self.al_cerrar()
        self.destroy()


# ───────────────────────── aplicación ─────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        root.title(f"{APP} — archivos, enlaces y reuniones en vivo")
        root.geometry("1000x720")
        root.minsize(820, 560)
        self.motor = Motor()
        self.traductor = Traductor()
        self.q = queue.Queue()
        self.segs = []
        self.parciales = {}
        self.media = None
        self.titulo = "Transcripción"
        self.ocupado = False
        self.cola = []
        self._ultima_carpeta = os.path.expanduser("~")
        self.stop_evt = threading.Event()
        self.captura = None
        self.flotante = None
        self.altavoces, self.micros = [], []
        os.makedirs(OUT_DIR, exist_ok=True)
        self._estilo()
        self._ui()
        root.protocol("WM_DELETE_WINDOW", self._salir)
        root.after(100, self._bucle)
        # arrastrar y soltar archivos o carpetas sobre la ventana
        try:
            from tkinterdnd2 import DND_FILES
            root.drop_target_register(DND_FILES)
            root.dnd_bind("<<Drop>>", self._soltar)
        except Exception:
            log.info("Arrastrar y soltar no disponible (falta tkinterdnd2)")
        # archivos recibidos desde "Enviar a" o "Abrir con"
        args = [a for a in sys.argv[1:] if os.path.isfile(a) or os.path.isdir(a)]
        if args:
            rutas = []
            for a in args:
                rutas += self._medios_de_carpeta(a) if os.path.isdir(a) else [a]
            root.after(500, lambda: self._poner_archivos(rutas, iniciar=True))

    # ---------- interfaz ----------
    def _estilo(self):
        s = ttk.Style()
        try:
            s.theme_use("vista" if sys.platform == "win32" else "clam")
        except tk.TclError:
            pass
        s.configure("Acc.TButton", font=("Segoe UI", 10, "bold"))
        s.configure("Tit.TLabel", font=("Segoe UI", 10, "bold"))

    def _ui(self):
        r = self.root
        top = ttk.Frame(r, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Modelo:").pack(side="left")
        self.modelo = tk.StringVar(value="small")
        ttk.Combobox(top, textvariable=self.modelo, values=MODELOS, width=14, state="readonly").pack(side="left", padx=(4, 14))
        ttk.Label(top, text="Idioma:").pack(side="left")
        self.idioma = tk.StringVar(value=IDIOMAS[0][0])
        ttk.Combobox(top, textvariable=self.idioma, values=[i[0] for i in IDIOMAS], width=22, state="readonly").pack(side="left", padx=4)
        self.estado = ttk.Label(top, text="Listo.", foreground="#3a5a8c")
        self.estado.pack(side="right")

        nb = ttk.Notebook(r)
        nb.pack(fill="x", padx=10)
        self.nb = nb

        # 1) Archivos en cola
        f1 = ttk.Frame(nb, padding=10)
        nb.add(f1, text="  1 · Archivos (cola automática)  ")
        bar1 = ttk.Frame(f1)
        bar1.grid(row=0, column=0, columnspan=2, sticky="ew")
        for texto, cmd in (("Agregar archivos…", self._examinar), ("Agregar carpeta…", self._agregar_carpeta),
                           ("Agregar enlaces…", self._agregar_enlaces), ("Quitar", self._quitar_cola),
                           ("Vaciar", self._vaciar_cola), ("▲", lambda: self._mover_cola(-1)),
                           ("▼", lambda: self._mover_cola(1))):
            ttk.Button(bar1, text=texto, command=cmd, width=3 if len(texto) == 1 else None).pack(side="left", padx=(0, 4))
        self.btn_cola = ttk.Button(bar1, text="▶ Procesar cola", style="Acc.TButton", command=self._procesar_cola)
        self.btn_cola.pack(side="right")
        self.arbol = ttk.Treeview(f1, columns=("n", "nombre", "estado"), show="headings", height=5,
                                  selectmode="extended")
        for col, tit, ancho, est in (("n", "#", 36, False), ("nombre", "Archivo o enlace", 520, True),
                                     ("estado", "Estado", 230, False)):
            self.arbol.heading(col, text=tit, anchor="w")
            self.arbol.column(col, width=ancho, stretch=est, anchor="w")
        self.arbol.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        sb = ttk.Scrollbar(f1, orient="vertical", command=self.arbol.yview)
        sb.grid(row=1, column=1, sticky="ns", pady=(6, 0))
        self.arbol.configure(yscrollcommand=sb.set)
        self.arbol.bind("<Delete>", lambda e: self._quitar_cola())
        opc = ttk.Frame(f1)
        opc.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(opc, text="Al terminar cada uno:").grid(row=0, column=0, sticky="w")
        self.op_word = tk.BooleanVar(value=False)
        self.op_trad = tk.BooleanVar(value=False)
        self.op_dir = tk.StringVar(value="Detectar automáticamente")
        self.op_video = tk.BooleanVar(value=False)
        self.op_subs = tk.StringVar(value="Original")
        ttk.Label(opc, text="guarda TXT y SRT siempre", foreground="#666").grid(row=0, column=1, sticky="w", padx=8)
        ttk.Checkbutton(opc, text="Guardar también en Word", variable=self.op_word).grid(row=0, column=2, sticky="w", padx=8)
        ttk.Checkbutton(opc, text="Traducir:", variable=self.op_trad).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Combobox(opc, textvariable=self.op_dir, values=list(DIRECCIONES), width=24,
                     state="readonly").grid(row=1, column=1, sticky="w", padx=8, pady=(4, 0))
        fv = ttk.Frame(opc)
        fv.grid(row=1, column=2, columnspan=2, sticky="w", padx=8, pady=(4, 0))
        ttk.Checkbutton(fv, text="Crear video con subtítulos en:", variable=self.op_video).pack(side="left")
        ttk.Combobox(fv, textvariable=self.op_subs, values=["Original", "Traducción", "Ambos"], width=11,
                     state="readonly").pack(side="left", padx=4)
        ttk.Label(fv, text="(usa tu formato de subtítulos guardado)", foreground="#666").pack(side="left")
        fd = ttk.Frame(opc)
        fd.grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))
        ttk.Label(fd, text="Guardar resultados en:").pack(side="left")
        self.op_destino = tk.StringVar(value=OUT_DIR)
        ttk.Label(fd, textvariable=self.op_destino, foreground="#1d4f91").pack(side="left", padx=6)
        ttk.Button(fd, text="Cambiar…", command=self._cambiar_destino).pack(side="left")
        ttk.Button(fd, text="Abrir carpeta", command=lambda: self._abrir_carpeta(self.op_destino.get())).pack(side="left", padx=4)
        ttk.Label(f1, text="Puedes agregar más archivos mientras la cola avanza; si uno falla, sigue con el siguiente. "
                           "También puedes arrastrar archivos o carpetas a la ventana.",
                  foreground="#666").grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        f1.columnconfigure(0, weight=1)

        # 2) Enlace
        f2 = ttk.Frame(nb, padding=10)
        nb.add(f2, text="  2 · Enlace (YouTube y otros)  ")
        self.url = tk.StringVar()
        ttk.Entry(f2, textvariable=self.url).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(f2, text="Descargar y transcribir", style="Acc.TButton", command=self._transcribir_url).grid(row=0, column=1)
        self.con_video = tk.BooleanVar(value=True)
        ttk.Checkbutton(f2, text="Bajar también el video (para verlo con la transcripción sincronizada)",
                        variable=self.con_video).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(f2, text="YouTube, Vimeo, Facebook, X, Instagram, enlaces directos a .mp4/.mp3 y cientos de sitios más.",
                  foreground="#666").grid(row=2, column=0, columnspan=2, sticky="w")
        f2.columnconfigure(0, weight=1)

        # 3) En vivo
        f3 = ttk.Frame(nb, padding=10)
        nb.add(f3, text="  3 · En vivo (reuniones y videos)  ")
        self.usar_sistema = tk.BooleanVar(value=True)
        self.usar_mic = tk.BooleanVar(value=False)
        ttk.Checkbutton(f3, text="Audio del equipo (lo que escuchas: Meet, Zoom, Teams, Webex, YouTube…)",
                        variable=self.usar_sistema).grid(row=0, column=0, sticky="w")
        self.cb_alt = ttk.Combobox(f3, width=42, state="readonly")
        self.cb_alt.grid(row=0, column=1, padx=6, sticky="ew")
        ttk.Checkbutton(f3, text="Mi micrófono (usa audífonos para no duplicar)",
                        variable=self.usar_mic).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.cb_mic = ttk.Combobox(f3, width=42, state="readonly")
        self.cb_mic.grid(row=1, column=1, padx=6, sticky="ew", pady=(4, 0))
        ttk.Button(f3, text="↻ Dispositivos", command=self._cargar_dispositivos).grid(row=0, column=2, rowspan=2)
        bar = ttk.Frame(f3)
        bar.grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.btn_vivo = ttk.Button(bar, text="● Iniciar transcripción en vivo", style="Acc.TButton", command=self._vivo)
        self.btn_vivo.pack(side="left")
        ttk.Button(bar, text="Subtítulos flotantes", command=self._toggle_flotante).pack(side="left", padx=8)
        self.vista_previa = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="Vista previa instantánea (texto provisional mientras hablan)",
                        variable=self.vista_previa).pack(side="left", padx=8)
        tv = ttk.Frame(f3)
        tv.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.trad_vivo = tk.BooleanVar(value=False)
        self.dir_vivo = tk.StringVar(value="Detectar automáticamente")
        ttk.Checkbutton(tv, text="Traducir en vivo:", variable=self.trad_vivo).pack(side="left")
        ttk.Combobox(tv, textvariable=self.dir_vivo, values=list(DIRECCIONES), width=24,
                     state="readonly").pack(side="left", padx=6)
        f3.columnconfigure(1, weight=1)

        # 4) Traductor de texto
        f4 = ttk.Frame(nb, padding=10)
        nb.add(f4, text="  4 · Traductor  ")
        barra = ttk.Frame(f4)
        barra.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        self.dir_texto = tk.StringVar(value="Detectar automáticamente")
        ttk.Combobox(barra, textvariable=self.dir_texto, values=list(DIRECCIONES), width=24,
                     state="readonly").pack(side="left")
        ttk.Button(barra, text="Traducir", style="Acc.TButton", command=self._traducir_texto).pack(side="left", padx=6)
        ttk.Button(barra, text="⇄ Intercambiar", command=self._intercambiar).pack(side="left")
        ttk.Button(barra, text="Usar la transcripción", command=self._usar_transcripcion).pack(side="left", padx=6)
        ttk.Button(barra, text="Copiar traducción", command=self._copiar_traduccion).pack(side="left")
        ttk.Label(barra, text="Ctrl+Enter traduce", foreground="#777").pack(side="right")
        self.tr_in = tk.Text(f4, height=6, wrap="word", font=("Segoe UI", 10), undo=True)
        self.tr_out = tk.Text(f4, height=6, wrap="word", font=("Segoe UI", 10), background="#f6f8fb")
        self.tr_in.grid(row=1, column=0, sticky="nsew", padx=(0, 4))
        self.tr_out.grid(row=1, column=1, sticky="nsew", padx=(4, 0))
        self.tr_in.bind("<Control-Return>", lambda e: (self._traducir_texto(), "break")[1])
        f4.columnconfigure(0, weight=1)
        f4.columnconfigure(1, weight=1)
        nb.bind("<<NotebookTabChanged>>", lambda e: self._cargar_dispositivos(silencioso=True)
                if nb.index("current") == 2 and not self.altavoces else None)

        # progreso
        pf = ttk.Frame(r, padding=(10, 6, 10, 0))
        pf.pack(fill="x")
        self.prog = ttk.Progressbar(pf, mode="determinate", maximum=100)
        self.prog.pack(side="left", fill="x", expand=True)
        self.btn_det = ttk.Button(pf, text="Detener", command=self._detener, state="disabled")
        self.btn_det.pack(side="left", padx=(8, 0))

        # transcripción
        tf = ttk.Frame(r, padding=(10, 6))
        tf.pack(fill="both", expand=True)
        cab = ttk.Frame(tf)
        cab.pack(fill="x")
        ttk.Label(cab, text="Transcripción", style="Tit.TLabel").pack(side="left")
        ttk.Label(cab, text="  clic derecho en una o varias líneas para traducirlas (Ctrl+T)",
                  foreground="#777").pack(side="left")
        self.seguir = tk.BooleanVar(value=True)
        ttk.Checkbutton(cab, text="Desplazamiento automático", variable=self.seguir).pack(side="right")
        self.vista = tk.StringVar(value="Original")
        cbv = ttk.Combobox(cab, textvariable=self.vista, values=["Original", "Traducción", "Ambos"],
                           width=11, state="readonly")
        cbv.pack(side="right", padx=(4, 14))
        cbv.bind("<<ComboboxSelected>>", lambda e: self._repintar())
        ttk.Label(cab, text="Mostrar y exportar:").pack(side="right")
        ttk.Button(cab, text="Traducir transcripción…", command=self._dialogo_traducir).pack(side="right", padx=14)
        self.txt = ScrolledText(tf, wrap="word", font=("Segoe UI", 11), padx=10, pady=8,
                                relief="flat", borderwidth=1, background="#fbfbfc")
        self.txt.pack(fill="both", expand=True, pady=(4, 0))
        self.txt.tag_config("ts", foreground="#8a8f99", font=("Consolas", 9))
        self.txt.tag_config("yo", foreground="#1f6feb", font=("Segoe UI", 11, "bold"))
        self.txt.tag_config("otros", foreground="#b3471d", font=("Segoe UI", 11, "bold"))
        self.txt.tag_config("trad", foreground="#1b7a4b", font=("Segoe UI", 11, "italic"))
        self.txt.tag_raise("sel")
        self.txt.bind("<Button-3>", self._menu_texto)
        self.txt.bind("<Control-t>", self._atajo_traducir)
        self.txt.bind("<Control-T>", self._atajo_traducir)
        self.txt.bind("<Double-Button-1>", lambda e: (self._marcar(self._segs_en(self.txt.index(f"@{e.x},{e.y}"))), "break")[1])
        self.txt.bind("<Button-1>", lambda e: self.txt.focus_set())
        self.txt.configure(state="disabled")
        self.lbl_parcial = tk.Label(tf, text="", fg="#6b7280", bg=r.cget("bg"), anchor="w", justify="left",
                                    font=("Segoe UI", 10, "italic"), wraplength=940)
        self.lbl_parcial.pack(fill="x", pady=(4, 0))

        # acciones
        af = ttk.Frame(r, padding=(10, 0, 10, 10))
        af.pack(fill="x")
        for texto, cmd in (("Copiar todo", self._copiar), ("Guardar TXT", lambda: self._guardar("txt")),
                           ("SRT (subtítulos)", lambda: self._guardar("srt")), ("VTT", lambda: self._guardar("vtt")),
                           ("Word", lambda: self._guardar("docx")), ("Limpiar", self._limpiar)):
            ttk.Button(af, text=texto, command=cmd).pack(side="left", padx=(0, 6))
        self.btn_visor = ttk.Button(af, text="▶ Ver video con transcripción sincronizada",
                                    style="Acc.TButton", command=self._visor, state="disabled")
        self.btn_visor.pack(side="right")
        self.btn_exportar = ttk.Button(af, text="Descargar video con subtítulos",
                                       command=self._dialogo_video, state="disabled")
        self.btn_exportar.pack(side="right", padx=6)

    # ---------- comunicación entre hilos ----------
    def emit(self, msg):
        self.q.put(msg)

    def _log(self, texto):
        self.emit(("log", texto))

    def _bucle(self):
        try:
            while True:
                m = self.q.get_nowait()
                tipo = m[0]
                if tipo == "log":
                    self.estado.config(text=m[1])
                elif tipo == "seg":
                    self.segs.append(m[1])
                    self._pintar(m[1], len(self.segs) - 1)
                elif tipo == "parcial":
                    if m[2]:
                        self.parciales[m[1]] = m[2]
                    else:
                        self.parciales.pop(m[1], None)
                    self._pintar_parcial()
                elif tipo == "progreso":
                    self.prog["value"] = m[1]
                elif tipo == "cola_item":
                    if self.arbol.exists(m[1]):
                        v = list(self.arbol.item(m[1], "values"))
                        v[2] = m[2]
                        self.arbol.item(m[1], values=v)
                        if m[2].startswith("En curso"):
                            self.arbol.see(m[1])
                elif tipo == "cola_nuevo":
                    self._limpiar(forzar=True)
                    self.titulo = os.path.splitext(m[1])[0]
                    self.root.title(f"{APP} — {m[1]}" + (f" (quedan {m[2]} en cola)" if m[2] > 0 else ""))
                elif tipo == "cola_traducido":
                    if self.vista.get() == "Original":
                        self.vista.set("Ambos")
                    self._repintar()
                elif tipo == "cola_media":
                    self.media, self.titulo = m[1], m[2]
                    self.btn_visor.config(state="normal")
                    try:
                        self.btn_exportar.config(state="normal" if info_media(self.media)[0] else "disabled")
                    except Exception:
                        pass
                elif tipo == "cola_fin":
                    self._cola_fin(*m[1:])
                elif tipo == "lineas_trad":
                    self._lineas_traducidas(*m[1:])
                elif tipo == "traducido":
                    self._traducido(m[1])
                elif tipo == "trad_texto":
                    self.tr_out.delete("1.0", "end")
                    self.tr_out.insert("1.0", m[1])
                    if m[2]:
                        self.estado.config(text="Traducido " + ("español → inglés" if m[2] == "es-en" else "inglés → español"))
                elif tipo == "video_listo":
                    self._video_listo(m[1])
                elif tipo == "fin":
                    self._fin(m[1])
                elif tipo == "error":
                    log.error(m[1])
                    self.estado.config(text="Error.")
                    messagebox.showerror(APP, m[1])
                    if not self.captura:
                        self._fin(None)
        except queue.Empty:
            pass
        self.root.after(100, self._bucle)

    def _pintar(self, s, i=None, mover=True):
        t = self.txt
        if i is None:
            i = len(self.segs) - 1
        tag = f"s{i}"
        t.configure(state="normal")
        t.insert("end", f"[{etiqueta(s)}] ", ("ts", tag))
        if s.get("src"):
            t.insert("end", f"{s['src']}: ", ("yo" if s["src"] == "Yo" else "otros", tag))
        vista = self.vista.get()
        if vista == "Traducción" and s.get("trad"):
            t.insert("end", s["trad"] + "\n", ("trad", tag))
        else:
            t.insert("end", s["text"] + "\n", (tag,))
            if s.get("trad") and (vista == "Ambos" or s.get("ver_trad")):
                t.insert("end", "      " + s["trad"] + "\n", ("trad", tag))
        t.configure(state="disabled")
        if mover and self.seguir.get():
            t.see("end")
        self._actualizar_flotante()

    def _repintar(self, conservar=False):
        pos = self.txt.yview()[0]
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")
        for i, s in enumerate(self.segs):
            self._pintar(s, i, mover=not conservar)
        if conservar:
            self.txt.yview_moveto(pos)

    def segs_vista(self):
        """Segmentos tal como se ven (original, traducción o ambos) para exportar."""
        v = self.vista.get()
        out = []
        for s in self.segs:
            tr = s.get("trad")
            if v == "Traducción" and tr:
                out.append(dict(s, text=tr))
            elif tr and (v == "Ambos" or s.get("ver_trad")):
                out.append(dict(s, text=s["text"] + "\n" + tr))
            else:
                out.append(s)
        return out

    # ---------- traducir líneas directamente en el texto ----------
    def _segs_en(self, indice=None):
        """Segmentos bajo la selección; si no hay selección, el de la línea indicada."""
        t = self.txt
        idx = set()
        try:
            a, b = t.index("sel.first"), t.index("sel.last")
        except tk.TclError:
            a = b = None
        if a and (indice is None or (t.compare(a, "<=", indice) and t.compare(indice, "<=", b))):
            pos = a
            while t.compare(pos, "<", b):
                for tg in t.tag_names(pos):
                    if re.fullmatch(r"s\d+", tg):
                        idx.add(int(tg[1:]))
                pos = t.index(f"{pos} +1 line linestart")
        elif indice is not None:
            for tg in t.tag_names(t.index(f"{indice} linestart")):
                if re.fullmatch(r"s\d+", tg):
                    idx.add(int(tg[1:]))
        return sorted(i for i in idx if i < len(self.segs))

    def _marcar(self, idxs):
        t = self.txt
        t.tag_remove("sel", "1.0", "end")
        for i in idxs:
            r = t.tag_ranges(f"s{i}")
            if r:
                t.tag_add("sel", r[0], r[-1])

    def _menu_texto(self, e):
        indice = self.txt.index(f"@{e.x},{e.y}")
        idxs = self._segs_en(indice)
        if not idxs:
            return
        self._marcar(idxs)
        n = len(idxs)
        q = "esta línea" if n == 1 else f"las {n} líneas seleccionadas"
        m = tk.Menu(self.root, tearoff=0)
        for titulo, modo in ((f"Traducir {q} (agregar debajo)", "debajo"),
                             (f"Traducir {q} (reemplazar el texto)", "reemplazar")):
            sub = tk.Menu(m, tearoff=0)
            for et, d in DIRECCIONES.items():
                sub.add_command(label=et, command=lambda d=d, modo=modo: self._traducir_lineas(idxs, d, modo))
            m.add_cascade(label=titulo, menu=sub)
        m.add_separator()
        if any(self.segs[i].get("orig") for i in idxs):
            m.add_command(label="Restaurar texto original", command=lambda: self._restaurar(idxs))
        if any(self.segs[i].get("trad") for i in idxs):
            m.add_command(label="Quitar traducción", command=lambda: self._quitar_trad(idxs))
        if n == 1:
            m.add_command(label="Editar línea…", command=lambda: self._editar_linea(idxs[0]))
        m.add_command(label="Copiar", command=lambda: self._copiar_lineas(idxs))
        m.add_separator()
        m.add_command(label="Seleccionar todo", command=lambda: self._marcar(range(len(self.segs))))
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    def _atajo_traducir(self, e=None):
        idxs = self._segs_en(self.txt.index("insert"))
        if idxs:
            self._traducir_lineas(idxs, "auto", "debajo")
        return "break"

    def _traducir_lineas(self, idxs, direccion, modo):
        if self.ocupado:
            return
        textos = [self.segs[i].get("orig") or self.segs[i]["text"] for i in idxs]
        self._log(f"Traduciendo {len(idxs)} línea(s)…")

        def job():
            try:
                trads = self.traductor.traducir_lista(textos, direccion, self._log)
                self.emit(("lineas_trad", idxs, trads, modo, None))
            except Exception as e:
                log.exception("traducir líneas")
                self.emit(("lineas_trad", idxs, [], modo, str(e)))
        threading.Thread(target=job, daemon=True).start()

    def _lineas_traducidas(self, idxs, trads, modo, error):
        if error:
            messagebox.showerror(APP, f"No se pudo traducir:\n{error}")
            self.estado.config(text="Error al traducir.")
            return
        for i, tr in zip(idxs, trads):
            if i >= len(self.segs) or not tr:
                continue
            s = self.segs[i]
            if modo == "reemplazar":
                s.setdefault("orig", s["text"])
                s["text"] = tr
                s.pop("trad", None)
                s.pop("ver_trad", None)
            else:
                s["trad"] = tr
                s["ver_trad"] = True
        self._repintar(conservar=True)
        self._marcar(idxs)
        self.estado.config(text=f"{len(idxs)} línea(s) traducida(s). Clic derecho para restaurar o quitar.")
        self._autoguardar()

    def _restaurar(self, idxs):
        for i in idxs:
            s = self.segs[i]
            if s.get("orig"):
                s["text"] = s.pop("orig")
        self._repintar(conservar=True)

    def _quitar_trad(self, idxs):
        for i in idxs:
            self.segs[i].pop("trad", None)
            self.segs[i].pop("ver_trad", None)
        self._repintar(conservar=True)

    def _editar_linea(self, i):
        s = self.segs[i]
        d = tk.Toplevel(self.root)
        d.title("Editar línea")
        d.transient(self.root)
        f = ttk.Frame(d, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"[{etiqueta(s)}] Texto:").pack(anchor="w")
        ta = tk.Text(f, height=4, width=80, wrap="word", font=("Segoe UI", 10))
        ta.pack(fill="both", expand=True)
        ta.insert("1.0", s["text"])
        tb = None
        if s.get("trad"):
            ttk.Label(f, text="Traducción:").pack(anchor="w", pady=(8, 0))
            tb = tk.Text(f, height=4, width=80, wrap="word", font=("Segoe UI", 10))
            tb.pack(fill="both", expand=True)
            tb.insert("1.0", s["trad"])

        def guardar():
            s["text"] = ta.get("1.0", "end").strip() or s["text"]
            if tb is not None:
                s["trad"] = tb.get("1.0", "end").strip()
            d.destroy()
            self._repintar(conservar=True)
            self._autoguardar()
        ttk.Button(f, text="Guardar", style="Acc.TButton", command=guardar).pack(anchor="e", pady=(8, 0))
        ta.focus_set()
        d.grab_set()

    def _copiar_lineas(self, idxs):
        sv = self.segs_vista()
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(sv[i]["text"] for i in idxs))
        self.estado.config(text="Copiado al portapapeles.")

    # ---------- traductor ----------
    def _dialogo_traducir(self):
        if self.ocupado or not self.segs:
            if not self.segs:
                messagebox.showinfo(APP, "Primero transcribe algo.")
            return
        d = tk.Toplevel(self.root)
        d.title("Traducir transcripción")
        d.transient(self.root)
        d.resizable(False, False)
        f = ttk.Frame(d, padding=14)
        f.pack()
        ttk.Label(f, text="Sentido de la traducción:").grid(row=0, column=0, sticky="w")
        dv = tk.StringVar(value="Detectar automáticamente")
        ttk.Combobox(f, textvariable=dv, values=list(DIRECCIONES), width=26, state="readonly").grid(row=0, column=1, padx=6)
        ttk.Label(f, text="La traducción se agrega bajo cada frase; elige qué mostrar y exportar\n"
                          "(original, traducción o ambos) en el selector sobre la transcripción.",
                  foreground="#555").grid(row=1, column=0, columnspan=2, sticky="w", pady=8)

        def ir():
            direccion = DIRECCIONES[dv.get()]
            d.destroy()
            self._ocupar(True)
            self.stop_evt.clear()
            segs = list(self.segs)

            def job():
                try:
                    n = len(segs)
                    for i in range(0, n, 24):
                        if self.stop_evt.is_set():
                            break
                        lote = segs[i:i + 24]
                        trads = self.traductor.traducir_lista([s["text"] for s in lote], direccion, self._log)
                        for s, t in zip(lote, trads):
                            s["trad"] = t
                        self.emit(("progreso", min(100, (i + len(lote)) / n * 100)))
                        self._log(f"Traduciendo… {min(n, i + len(lote))} de {n} frases")
                    self.emit(("traducido", None))
                except Exception as e:
                    log.exception("traducir transcripción")
                    self.emit(("traducido", str(e)))
            threading.Thread(target=job, daemon=True).start()
        ttk.Button(f, text="Traducir", style="Acc.TButton", command=ir).grid(row=2, column=1, sticky="e")
        d.grab_set()

    def _traducido(self, error):
        self._ocupar(False)
        if error:
            messagebox.showerror(APP, f"No se pudo traducir:\n{error}")
            self.estado.config(text="Error al traducir.")
            return
        if self.vista.get() == "Original":
            self.vista.set("Ambos")
        self._repintar()
        self.prog["value"] = 100
        self.estado.config(text="Traducción lista. Elige «Original», «Traducción» o «Ambos» para ver y exportar.")
        self._autoguardar()

    def _traducir_texto(self):
        texto = self.tr_in.get("1.0", "end").strip()
        if not texto:
            return
        direccion = DIRECCIONES[self.dir_texto.get()]
        self.tr_out.delete("1.0", "end")
        self.tr_out.insert("1.0", "Traduciendo…")

        def job():
            try:
                trad, sentido = self.traductor.traducir_texto(texto, direccion, self._log)
                self.emit(("trad_texto", trad, sentido))
            except Exception as e:
                log.exception("traducir texto")
                self.emit(("trad_texto", f"[No se pudo traducir: {e}]", None))
        threading.Thread(target=job, daemon=True).start()

    def _intercambiar(self):
        a, b = self.tr_in.get("1.0", "end-1c"), self.tr_out.get("1.0", "end-1c")
        self.tr_in.delete("1.0", "end")
        self.tr_in.insert("1.0", b)
        self.tr_out.delete("1.0", "end")
        self.tr_out.insert("1.0", a)
        inv = {"Inglés → Español": "Español → Inglés", "Español → Inglés": "Inglés → Español"}
        self.dir_texto.set(inv.get(self.dir_texto.get(), self.dir_texto.get()))

    def _usar_transcripcion(self):
        self.tr_in.delete("1.0", "end")
        self.tr_in.insert("1.0", "\n".join(s["text"] for s in self.segs))

    def _copiar_traduccion(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.tr_out.get("1.0", "end-1c"))
        self.estado.config(text="Traducción copiada al portapapeles.")

    def _pintar_parcial(self):
        txt = "   ".join(f"{k}: {v}…" for k, v in self.parciales.items())
        self.lbl_parcial.config(text=("✎ " + txt) if txt else "")
        self._actualizar_flotante()

    def _actualizar_flotante(self):
        if self.flotante:
            lineas = [(f"{s['src']}: " if self.captura and len(self.captura.canales) > 1 else "") + s["text"]
                      for s in self.segs_vista()[-4:]]
            parcial = "   ".join(f"{v}…" for v in self.parciales.values())
            self.flotante.mostrar(lineas, parcial)

    # ---------- utilidades UI ----------
    def _idioma(self):
        return dict(IDIOMAS).get(self.idioma.get(), "es")

    def _ocupar(self, si):
        self.ocupado = si
        self.btn_det.config(state="normal" if si else "disabled")
        if si:
            self.prog["value"] = 0

    def _limpiar(self, forzar=False):
        if self.ocupado and not forzar:
            return
        self.segs.clear()
        self.parciales.clear()
        self.media = None
        self.btn_visor.config(state="disabled")
        self.btn_exportar.config(state="disabled")
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")
        self.lbl_parcial.config(text="")
        self.prog["value"] = 0

    def _examinar(self):
        ps = filedialog.askopenfilenames(
            title="Elige uno o varios archivos de audio o video (de cualquier carpeta o unidad)",
            initialdir=self._ultima_carpeta, filetypes=[
                ("Audio y video", " ".join("*" + e for e in EXT_MEDIA)), ("Todos", "*.*")])
        if ps:
            self._poner_archivos(list(ps))

    def _agregar_carpeta(self):
        c = filedialog.askdirectory(title="Elige una carpeta: se agregan todos sus audios y videos",
                                    initialdir=self._ultima_carpeta)
        if c:
            self._poner_archivos(self._medios_de_carpeta(c))

    @staticmethod
    def _medios_de_carpeta(c):
        rutas = []
        for raiz, _, fs in os.walk(c):
            rutas += [os.path.join(raiz, f) for f in sorted(fs) if os.path.splitext(f)[1].lower() in EXT_MEDIA]
        return rutas

    def _agregar_enlaces(self):
        d = tk.Toplevel(self.root)
        d.title("Agregar enlaces a la cola")
        d.transient(self.root)
        f = ttk.Frame(d, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Pega uno o varios enlaces (YouTube u otros), uno por línea:").pack(anchor="w")
        t = tk.Text(f, width=80, height=10, font=("Segoe UI", 10))
        t.pack(fill="both", expand=True, pady=6)
        try:
            clip = self.root.clipboard_get()
            if re.match(r"https?://", clip.strip()):
                t.insert("1.0", clip.strip() + "\n")
        except tk.TclError:
            pass

        def ok():
            urls = re.findall(r"https?://\S+", t.get("1.0", "end"))
            d.destroy()
            self._poner_en_cola([{"url": u} for u in urls])
        ttk.Button(f, text="Agregar a la cola", style="Acc.TButton", command=ok).pack(anchor="e")
        t.focus_set()
        d.grab_set()

    def _poner_archivos(self, rutas, iniciar=False):
        rutas = [r for r in rutas if os.path.isfile(r)]
        if not rutas:
            return
        self._ultima_carpeta = os.path.dirname(rutas[0])
        self._poner_en_cola([{"ruta": r} for r in rutas], iniciar)

    def _poner_en_cola(self, items, iniciar=False):
        existentes = {it.get("ruta") or it.get("url") for it in self.cola if it["estado"] == "Pendiente"}
        n = 0
        for it in items:
            clave = it.get("ruta") or it.get("url")
            if clave in existentes:
                continue
            existentes.add(clave)
            it["estado"] = "Pendiente"
            it["nombre"] = os.path.basename(it["ruta"]) if it.get("ruta") else it["url"]
            it["iid"] = self.arbol.insert("", "end", values=("", it["nombre"], "Pendiente"))
            self.cola.append(it)
            n += 1
        self._numerar()
        self.nb.select(0)
        if n:
            self.estado.config(text=f"{n} elemento(s) agregados a la cola.")
        if iniciar and not self.ocupado:
            self._procesar_cola()

    def _numerar(self):
        for i, iid in enumerate(self.arbol.get_children(), 1):
            v = list(self.arbol.item(iid, "values"))
            v[0] = i
            self.arbol.item(iid, values=v)

    def _item_de(self, iid):
        return next((it for it in self.cola if it["iid"] == iid), None)

    def _quitar_cola(self):
        for iid in self.arbol.selection():
            it = self._item_de(iid)
            if it and it["estado"].startswith("En curso"):
                continue
            if it:
                self.cola.remove(it)
            self.arbol.delete(iid)
        self._numerar()

    def _vaciar_cola(self):
        for it in list(self.cola):
            if not it["estado"].startswith("En curso"):
                self.cola.remove(it)
                self.arbol.delete(it["iid"])
        self._numerar()

    def _mover_cola(self, paso):
        sel = self.arbol.selection()
        if not sel:
            return
        hijos = list(self.arbol.get_children())
        for iid in (sel if paso < 0 else reversed(sel)):
            i = hijos.index(iid)
            j = max(0, min(len(hijos) - 1, i + paso))
            self.arbol.move(iid, "", j)
            hijos = list(self.arbol.get_children())
        orden = {iid: k for k, iid in enumerate(hijos)}
        self.cola.sort(key=lambda it: orden.get(it["iid"], 1e9))
        self._numerar()

    def _cambiar_destino(self):
        c = filedialog.askdirectory(title="Carpeta donde guardar los resultados", initialdir=self.op_destino.get())
        if c:
            self.op_destino.set(os.path.normpath(c))

    def _abrir_carpeta(self, c):
        try:
            os.makedirs(c, exist_ok=True)
            os.startfile(c)
        except AttributeError:
            webbrowser.open("file://" + c)
        except Exception as e:
            messagebox.showerror(APP, str(e))

    def _soltar(self, evento):
        rutas = []
        for r in self.root.tk.splitlist(evento.data):
            rutas += self._medios_de_carpeta(r) if os.path.isdir(r) else [r]
        self._poner_archivos(rutas, iniciar=True)

    # ---------- 1) cola automática ----------
    def _procesar_cola(self):
        if self.ocupado:
            return
        for it in self.cola:          # lo detenido antes vuelve a quedar pendiente
            if it["estado"] == "Detenido" or it["estado"].startswith("En curso"):
                it["estado"] = "Pendiente"
                v = list(self.arbol.item(it["iid"], "values"))
                v[2] = "Pendiente"
                self.arbol.item(it["iid"], values=v)
        if not any(it["estado"] == "Pendiente" for it in self.cola):
            messagebox.showinfo(APP, "La cola está vacía. Agrega archivos, una carpeta o enlaces.")
            return
        op = {"word": self.op_word.get(), "trad": self.op_trad.get(), "dir": DIRECCIONES[self.op_dir.get()],
              "video": self.op_video.get(), "subs": self.op_subs.get(), "destino": self.op_destino.get(),
              "modelo": self.modelo.get(), "idioma": self._idioma(), "con_video": self.con_video.get()}
        self._ocupar(True)
        self.stop_evt.clear()
        self.btn_cola.config(state="disabled")
        threading.Thread(target=self._job_cola, args=(op,), daemon=True).start()

    def _job_cola(self, op):
        hechos, fallos = 0, 0
        os.makedirs(op["destino"], exist_ok=True)
        try:
            self.motor.cargar(op["modelo"], self._log)
        except Exception as e:
            log.exception("modelo")
            self.emit(("cola_fin", 0, 0, f"No se pudo cargar el modelo: {e}"))
            return
        while not self.stop_evt.is_set():
            it = next((x for x in self.cola if x["estado"] == "Pendiente"), None)
            if it is None:
                break
            pend = sum(1 for x in self.cola if x["estado"] == "Pendiente") - 1
            self._estado_item(it, "En curso…")
            self.emit(("cola_nuevo", it["nombre"], pend))
            try:
                if it.get("url"):
                    self._estado_item(it, "En curso: descargando…")
                    ruta, titulo = self._descargar(it["url"], op["con_video"])
                else:
                    ruta, titulo = it["ruta"], os.path.splitext(it["nombre"])[0]
                self._estado_item(it, "En curso: transcribiendo…")
                segs = self._transcribir_lista(ruta, op["idioma"], lambda pct: self._estado_item(
                    it, f"En curso: transcribiendo {pct:.0f}%"))
                if self.stop_evt.is_set():
                    self._estado_item(it, "Detenido")
                    break
                if not segs:
                    raise RuntimeError("No se detectó voz en el archivo.")
                if op["trad"]:
                    self._estado_item(it, "En curso: traduciendo…")
                    for i in range(0, len(segs), 24):
                        lote = segs[i:i + 24]
                        for sg, tr in zip(lote, self.traductor.traducir_lista([x["text"] for x in lote], op["dir"], self._log)):
                            sg["trad"] = tr
                    self.emit(("cola_traducido",))
                self._estado_item(it, "En curso: guardando…")
                base = os.path.join(op["destino"], re.sub(r'[\\/:*?"<>|]+', "_", titulo)[:90])
                vista_txt = "Ambos" if op["trad"] else "Original"
                with open(base + ".txt", "w", encoding="utf-8") as f:
                    f.write(a_txt(segs_segun(segs, vista_txt), titulo))
                with open(base + ".srt", "w", encoding="utf-8") as f:
                    f.write(a_srt(segs))
                if op["trad"]:
                    with open(base + " - traducción.srt", "w", encoding="utf-8") as f:
                        f.write(a_srt(segs_segun(segs, "Traducción")))
                if op["word"]:
                    a_docx(segs_segun(segs, vista_txt), base + ".docx", titulo)
                extra = ""
                if op["video"]:
                    if info_media(ruta)[0]:
                        self._estado_item(it, "En curso: creando video…")
                        ok = video_subtitulado(ruta, segs_segun(segs, op["subs"]), base + " - subtitulado.mp4",
                                               cargar_estilo(), lambda pct, t, d: self._estado_item(
                                                   it, f"En curso: creando video {pct:.0f}%"), self.stop_evt)
                        if not ok:
                            self._estado_item(it, "Detenido")
                            break
                        extra = " + video"
                    else:
                        extra = " (sin video: es solo audio)"
                self.emit(("cola_media", ruta, titulo))
                self._estado_item(it, "Listo ✓" + extra)
                hechos += 1
            except Exception as e:
                log.exception("cola: " + it["nombre"])
                self._estado_item(it, "Error: " + str(e).splitlines()[0][:120] if str(e) else "Error")
                fallos += 1
        self.emit(("cola_fin", hechos, fallos, op["destino"]))

    def _estado_item(self, it, estado):
        it["estado"] = estado
        self.emit(("cola_item", it["iid"], estado))

    def _transcribir_lista(self, ruta, idioma, al_avanzar):
        segs_w, info = self.motor.archivo(ruta, idioma)
        dur = max(info.duration or 1, 1)
        out, ult = [], -5
        for s in segs_w:
            if self.stop_evt.is_set():
                break
            texto = s.text.strip()
            if texto and not es_alucinacion(texto):
                sg = {"start": s.start, "end": s.end, "text": texto, "src": "", "wall": None}
                out.append(sg)
                self.emit(("seg", sg))
            pct = min(100, s.end / dur * 100)
            self.emit(("progreso", pct))
            if pct - ult >= 1:
                ult = pct
                al_avanzar(pct)
        return out

    def _descargar(self, url, con_video):
        import yt_dlp
        destino = os.path.join(OUT_DIR, "descargas")
        os.makedirs(destino, exist_ok=True)

        def hook(d):
            if self.stop_evt.is_set():
                raise yt_dlp.utils.DownloadCancelled("Cancelado")
            if d.get("status") == "downloading":
                tot = d.get("total_bytes") or d.get("total_bytes_estimate") or 0
                if tot:
                    self.emit(("progreso", d.get("downloaded_bytes", 0) / tot * 100))
                self._log(f"Descargando… {d.get('_percent_str', '').strip()}")

        fmt = ("18/best[ext=mp4][vcodec!=none][acodec!=none]/best[acodec!=none][vcodec!=none]/best"
               if con_video else "bestaudio[ext=m4a]/bestaudio/best")
        opts = {"format": fmt, "noplaylist": True, "quiet": True, "no_warnings": True,
                "outtmpl": os.path.join(destino, "%(title).80s [%(id)s].%(ext)s"),
                "windowsfilenames": True, "progress_hooks": [hook]}
        self._log("Obteniendo información del enlace…")
        with yt_dlp.YoutubeDL(opts) as ydl:
            info = ydl.extract_info(url, download=True)
            ruta = ydl.prepare_filename(info)
        if not os.path.exists(ruta):
            base = os.path.splitext(ruta)[0]
            cand = [os.path.join(destino, f) for f in os.listdir(destino)
                    if os.path.join(destino, f).startswith(base)]
            ruta = cand[0] if cand else ruta
        return ruta, info.get("title") or os.path.splitext(os.path.basename(ruta))[0]

    def _cola_fin(self, hechos, fallos, destino_o_error):
        self._ocupar(False)
        self.btn_cola.config(state="normal")
        self.root.title(f"{APP} — archivos, enlaces y reuniones en vivo")
        if hechos == 0 and fallos == 0 and not os.path.isdir(str(destino_o_error)):
            messagebox.showerror(APP, destino_o_error)
            return
        det = " (detenida)" if self.stop_evt.is_set() else ""
        self.estado.config(text=f"Cola terminada{det}: {hechos} listo(s), {fallos} con error.")
        self.prog["value"] = 100
        try:
            self.root.bell()
        except tk.TclError:
            pass
        msg = f"Cola terminada{det}.\n\n{hechos} archivo(s) listo(s)" + (f", {fallos} con error" if fallos else "") + \
              f".\nResultados en:\n{destino_o_error}\n\n¿Abrir la carpeta?"
        if messagebox.askyesno(APP, msg):
            self._abrir_carpeta(destino_o_error)

    def _job_archivo(self, ruta, titulo=None):
        try:
            self.motor.cargar(self.modelo.get(), self._log)
            self._log("Transcribiendo… el texto aparece a medida que avanza.")
            segs, info = self.motor.archivo(ruta, self._idioma())
            dur = max(info.duration or 1, 1)
            for s in segs:
                if self.stop_evt.is_set():
                    break
                texto = s.text.strip()
                if texto and not es_alucinacion(texto):
                    self.emit(("seg", {"start": s.start, "end": s.end, "text": texto, "src": "", "wall": None}))
                self.emit(("progreso", min(100, s.end / dur * 100)))
                self._log(f"Transcribiendo… {ts_corto(s.end)} de {ts_corto(dur)}")
            self.emit(("fin", {"media": ruta, "titulo": titulo or os.path.splitext(os.path.basename(ruta))[0]}))
        except Exception as e:
            log.exception("archivo")
            self.emit(("error", f"No se pudo transcribir:\n{e}\n\nDetalle en {LOG}"))

    # ---------- 2) enlace ----------
    def _transcribir_url(self):
        url = self.url.get().strip()
        if self.ocupado:
            return
        if not re.match(r"https?://", url):
            messagebox.showwarning(APP, "Pega un enlace que empiece con http:// o https://")
            return
        self._limpiar()
        self._ocupar(True)
        self.stop_evt.clear()
        threading.Thread(target=self._job_url, args=(url, self.con_video.get()), daemon=True).start()

    def _job_url(self, url, con_video):
        try:
            ruta, titulo = self._descargar(url, con_video)
        except Exception as e:
            log.exception("enlace")
            self.emit(("error", f"No se pudo descargar el enlace:\n{e}\n\n"
                                "Si es YouTube, ejecuta actualizar.bat (YouTube cambia seguido)."))
            return
        self.emit(("progreso", 0))
        self._job_archivo(ruta, titulo)

    # ---------- 3) en vivo ----------
    def _cargar_dispositivos(self, silencioso=False):
        try:
            _com_init()
            import soundcard as sc
            self.altavoces = sc.all_speakers()
            self.micros = sc.all_microphones()
            def_alt, def_mic = sc.default_speaker(), sc.default_microphone()
        except Exception as e:
            if not silencioso:
                messagebox.showerror(APP, f"No se pudieron leer los dispositivos de audio:\n{e}")
            return
        self.cb_alt["values"] = [a.name for a in self.altavoces]
        self.cb_mic["values"] = [m.name for m in self.micros]
        if self.altavoces:
            nombres = [a.name for a in self.altavoces]
            self.cb_alt.current(nombres.index(def_alt.name) if def_alt and def_alt.name in nombres else 0)
        if self.micros:
            nombres = [m.name for m in self.micros]
            self.cb_mic.current(nombres.index(def_mic.name) if def_mic and def_mic.name in nombres else 0)

    def _vivo(self):
        if self.captura:
            self._detener()
            return
        if self.ocupado:
            return
        if not (self.usar_sistema.get() or self.usar_mic.get()):
            messagebox.showwarning(APP, "Activa al menos una fuente: audio del equipo o micrófono.")
            return
        if not self.altavoces and not self.micros:
            self._cargar_dispositivos()
        try:
            import soundcard as sc
            canales = []
            if self.usar_sistema.get():
                alt = self.altavoces[self.cb_alt.current()] if self.altavoces else sc.default_speaker()
                lb = sc.get_microphone(id=str(alt.name), include_loopback=True)
                canales.append(Canal("Reunión", lb))
            if self.usar_mic.get():
                mic = self.micros[self.cb_mic.current()] if self.micros else sc.default_microphone()
                canales.append(Canal("Yo", mic))
        except Exception as e:
            log.exception("abrir audio")
            messagebox.showerror(APP, f"No se pudo abrir el audio:\n{e}")
            return
        self._limpiar()
        self._ocupar(True)
        self.prog.config(mode="indeterminate")
        self.prog.start(15)
        self.btn_vivo.config(text="■ Detener transcripción en vivo")

        def arrancar():
            try:
                self.motor.cargar(self.modelo.get(), self._log)
            except Exception as e:
                log.exception("modelo")
                self.emit(("error", f"No se pudo cargar el modelo:\n{e}"))
                return
            self.captura = Captura(self, canales, self.vista_previa.get(), self._idioma())
            self.captura.iniciar()
            self._log("● Escuchando… habla o reproduce el audio de la reunión/video.")
        threading.Thread(target=arrancar, daemon=True).start()

    def _toggle_flotante(self):
        if self.flotante:
            self.flotante.cerrar()
            return
        self.flotante = Flotante(self.root, lambda: setattr(self, "flotante", None))
        self._actualizar_flotante()

    # ---------- detener / fin ----------
    def _detener(self):
        self.stop_evt.set()
        if self.captura:
            self.captura.detener()
            self.captura = None
            self.prog.stop()
            self.prog.config(mode="determinate")
            self.btn_vivo.config(text="● Iniciar transcripción en vivo")
            self.parciales.clear()
            self._pintar_parcial()
            self._ocupar(False)
            self.titulo = f"Reunión {datetime.datetime.now():%d-%m-%Y %H%M}"
            self.estado.config(text=f"Sesión detenida · {len(self.segs)} segmentos. Puedes guardarla.")
            self._autoguardar()

    def _fin(self, datos):
        self._ocupar(False)
        if datos:
            self.media, self.titulo = datos["media"], datos["titulo"]
            self.prog["value"] = 100
            self.btn_visor.config(state="normal")
            try:
                self.btn_exportar.config(state="normal" if info_media(self.media)[0] else "disabled")
            except Exception:
                self.btn_exportar.config(state="disabled")
            self.estado.config(text=f"Listo · {len(self.segs)} segmentos." +
                               (" (detenido)" if self.stop_evt.is_set() else ""))
            self._autoguardar(srt=True)
        else:
            self.estado.config(text="Listo.")

    def _autoguardar(self, srt=False):
        if not self.segs:
            return
        base = re.sub(r'[\\/:*?"<>|]+', "_", self.titulo)[:90]
        try:
            with open(os.path.join(OUT_DIR, base + ".txt"), "w", encoding="utf-8") as f:
                f.write(a_txt(self.segs_vista(), self.titulo))
            if srt:
                with open(os.path.join(OUT_DIR, base + ".srt"), "w", encoding="utf-8") as f:
                    f.write(a_srt(self.segs_vista()))
        except Exception:
            log.exception("autoguardado")

    # ---------- exportar ----------
    def _copiar(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(a_txt(self.segs_vista()))
        self.estado.config(text="Copiado al portapapeles.")

    def _guardar(self, ext):
        if not self.segs:
            messagebox.showinfo(APP, "Aún no hay transcripción.")
            return
        base = re.sub(r'[\\/:*?"<>|]+', "_", self.titulo)[:90]
        p = filedialog.asksaveasfilename(initialdir=OUT_DIR, initialfile=f"{base}.{ext}",
                                         defaultextension="." + ext, filetypes=[(ext.upper(), "*." + ext)])
        if not p:
            return
        try:
            if ext == "docx":
                a_docx(self.segs_vista(), p, self.titulo)
            else:
                sv = self.segs_vista()
                contenido = {"txt": lambda: a_txt(sv, self.titulo),
                             "srt": lambda: a_srt(sv), "vtt": lambda: a_vtt(sv)}[ext]()
                with open(p, "w", encoding="utf-8") as f:
                    f.write(contenido)
            self.estado.config(text=f"Guardado: {p}")
        except Exception as e:
            messagebox.showerror(APP, f"No se pudo guardar:\n{e}")

    def _dialogo_video(self):
        if self.ocupado or not (self.media and self.segs):
            return
        from tkinter import colorchooser, font as tkfont
        import tempfile
        cfg = cargar_estilo()
        _, dur = info_media(self.media)
        media = self.media
        d = tk.Toplevel(self.root)
        d.title("Video con subtítulos — formato del texto")
        d.transient(self.root)
        f = ttk.Frame(d, padding=12)
        f.pack(fill="both", expand=True)
        izq = ttk.Frame(f)
        izq.grid(row=0, column=0, sticky="n")
        der = ttk.Frame(f)
        der.grid(row=0, column=1, sticky="n", padx=(14, 0))

        v = {k: (tk.BooleanVar(value=cfg[k]) if isinstance(cfg[k], bool) else
                 tk.IntVar(value=cfg[k]) if isinstance(cfg[k], int) else tk.StringVar(value=cfg[k]))
             for k in SUB_DEFECTO}
        fuentes = sorted({x for x in tkfont.families() if not x.startswith("@")}, key=str.lower)
        fila = [0]

        def campo(et, w):
            ttk.Label(izq, text=et).grid(row=fila[0], column=0, sticky="w", pady=3)
            w.grid(row=fila[0], column=1, sticky="w", padx=6, pady=3)
            fila[0] += 1
            return w

        campo("Tipo de letra:", ttk.Combobox(izq, textvariable=v["fuente"], values=fuentes, width=26, state="readonly"))
        campo("Tamaño:", ttk.Spinbox(izq, from_=8, to=80, textvariable=v["tamano"], width=6))
        est = ttk.Frame(izq)
        ttk.Checkbutton(est, text="Negrita", variable=v["negrita"]).pack(side="left")
        ttk.Checkbutton(est, text="Cursiva", variable=v["cursiva"]).pack(side="left", padx=8)
        ttk.Checkbutton(est, text="MAYÚSCULAS", variable=v["mayusculas"]).pack(side="left")
        campo("Estilo:", est)

        def boton_color(clave):
            b = tk.Button(izq, width=8, relief="groove", bg=v[clave].get(), activebackground=v[clave].get())

            def elegir():
                c = colorchooser.askcolor(color=v[clave].get(), parent=d, title="Elige un color")[1]
                if c:
                    v[clave].set(c.upper())
                    b.config(bg=c, activebackground=c)
            b.config(command=elegir)
            return b

        campo("Color del texto:", boton_color("color"))
        campo("Forma:", ttk.Combobox(izq, textvariable=v["forma"], values=FORMAS_SUB, width=20, state="readonly"))
        campo("Color de borde / caja:", boton_color("color_borde"))
        campo("Grosor del borde:", ttk.Spinbox(izq, from_=0, to=12, textvariable=v["grosor"], width=6))
        campo("Opacidad caja / sombra:", tk.Scale(izq, from_=0, to=100, orient="horizontal", variable=v["opacidad"],
                                                  length=170, showvalue=True))
        campo("Posición:", ttk.Combobox(izq, textvariable=v["posicion"], values=list(POSICIONES_SUB), width=10, state="readonly"))
        campo("Separación del borde:", ttk.Spinbox(izq, from_=0, to=150, textvariable=v["margen"], width=6))

        # vista previa real
        ttk.Label(der, text="Vista previa (así quedará en el video)").pack(anchor="w")
        lbl = tk.Label(der, text="Generando vista previa…", width=70, height=18, bg="#222", fg="#bbb")
        lbl.pack(pady=(4, 6))
        t_prev = tk.DoubleVar(value=self.segs[0]["start"] + 0.5 if self.segs else min(5, dur / 3))
        mf = ttk.Frame(der)
        mf.pack(fill="x")
        ttk.Label(mf, text="Momento del video:").pack(side="left")
        tk.Scale(mf, from_=0, to=max(1, dur - 0.5), orient="horizontal", variable=t_prev, resolution=0.5,
                 showvalue=False, length=330).pack(side="left", padx=6)
        lbl_t = ttk.Label(mf, text="")
        lbl_t.pack(side="left")
        png = os.path.join(tempfile.mkdtemp(prefix="transcriptor_vp_"), "prev.png")
        estado = {"job": None, "gen": 0, "img": None}

        def leer():
            c = {}
            for k, var in v.items():
                try:
                    c[k] = var.get()
                except (tk.TclError, ValueError):
                    c[k] = SUB_DEFECTO[k]
            return c

        def texto_en(t):
            sv = self.segs_vista()
            for s in sv:
                if s["start"] <= t <= s["end"] + 0.3:
                    return s["text"]
            cerca = min(sv, key=lambda s: abs(s["start"] - t))
            return cerca["text"]

        def generar():
            estado["gen"] += 1
            gen, c, t = estado["gen"], leer(), t_prev.get()
            lbl_t.config(text=ts_corto(t))

            def trabajo():
                try:
                    vista_previa_subs(media, texto_en(t), c, t, png)
                    d.after(0, lambda: mostrar(gen))
                except Exception as e:
                    log.exception("vista previa")
                    msg = str(e)
                    d.after(0, lambda: lbl.config(image="", text="No se pudo generar la vista previa:\n" + msg[-200:]))
            threading.Thread(target=trabajo, daemon=True).start()

        def mostrar(gen):
            if gen != estado["gen"] or not d.winfo_exists():
                return
            try:
                estado["img"] = tk.PhotoImage(file=png)
                lbl.config(image=estado["img"], text="", width=estado["img"].width(), height=estado["img"].height())
            except tk.TclError as e:
                lbl.config(text=f"Vista previa no disponible: {e}")

        def programar(*_):
            if estado["job"]:
                d.after_cancel(estado["job"])
            estado["job"] = d.after(450, generar)

        for var in list(v.values()) + [t_prev]:
            var.trace_add("write", programar)

        def restablecer():
            for k, val in SUB_DEFECTO.items():
                v[k].set(val)
            d.destroy()
            guardar_estilo(dict(SUB_DEFECTO))
            self._dialogo_video()

        def crear():
            c = leer()
            guardar_estilo(c)
            d.destroy()
            base = os.path.splitext(os.path.basename(media))[0]
            salida = filedialog.asksaveasfilename(
                title="Guardar video con subtítulos", initialdir=os.path.dirname(media),
                initialfile=f"{base} - subtitulado.mp4", defaultextension=".mp4", filetypes=[("Video MP4", "*.mp4")])
            if not salida:
                return
            if os.path.abspath(salida) == os.path.abspath(media):
                messagebox.showwarning(APP, "Elige otro nombre: no se puede sobrescribir el video original.")
                return
            self._ocupar(True)
            self.stop_evt.clear()

            def prog(pct, t, dur_):
                self.emit(("progreso", pct))
                self._log(f"Creando video con subtítulos… {pct:.0f}% ({ts_corto(t)} de {ts_corto(dur_)})")

            def job():
                try:
                    self._log("Creando video con subtítulos…")
                    ok_ = video_subtitulado(media, self.segs_vista(), salida, c, prog, self.stop_evt)
                    self.emit(("video_listo", salida if ok_ else None))
                except Exception as e:
                    log.exception("video subtitulado")
                    self.emit(("video_listo", None))
                    self.emit(("error", str(e)))
            threading.Thread(target=job, daemon=True).start()

        b = ttk.Frame(f)
        b.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(b, text="Restablecer", command=restablecer).pack(side="left")
        ttk.Button(b, text="Cancelar", command=lambda: (guardar_estilo(leer()), d.destroy())).pack(side="right")
        ttk.Button(b, text="Crear video", style="Acc.TButton", command=crear).pack(side="right", padx=6)
        generar()
        d.grab_set()

    def _video_listo(self, salida):
        self._ocupar(False)
        if not salida:
            self.estado.config(text="Creación de video cancelada.")
            return
        self.prog["value"] = 100
        self.estado.config(text=f"Video listo: {salida}")
        if messagebox.askyesno(APP, f"Video con subtítulos creado:\n{salida}\n\n¿Quieres reproducirlo ahora?"):
            try:
                os.startfile(salida)
            except AttributeError:
                webbrowser.open("file://" + salida)

    def _visor(self):
        if not (self.media and self.segs):
            return
        carpeta = os.path.dirname(self.media)
        ruta = os.path.join(carpeta, os.path.splitext(os.path.basename(self.media))[0] + " - transcripcion.html")
        try:
            a_visor(self.segs_vista(), self.media, ruta, self.titulo)
            with open(os.path.splitext(self.media)[0] + ".srt", "w", encoding="utf-8") as f:
                f.write(a_srt(self.segs_vista()))   # VLC y otros reproductores lo cargan solos
        except PermissionError:
            ruta = os.path.join(OUT_DIR, os.path.basename(ruta))
            a_visor(self.segs_vista(), self.media, ruta, self.titulo)
        webbrowser.open("file:///" + ruta.replace("\\", "/").lstrip("/"))

    def _salir(self):
        self.stop_evt.set()
        if self.captura:
            self.captura.detener()
        self.root.destroy()


# ───────────────────────── video con subtítulos ─────────────────────────
SUB_DEFECTO = {
    "fuente": "Arial", "tamano": 20, "negrita": False, "cursiva": False, "mayusculas": False,
    "color": "#FFFFFF", "forma": "Caja de fondo", "color_borde": "#000000",
    "grosor": 3, "opacidad": 60, "posicion": "Abajo", "margen": 22,
}
FORMAS_SUB = ["Caja de fondo", "Contorno", "Sombra", "Contorno y sombra", "Sin borde"]
POSICIONES_SUB = {"Abajo": 2, "Centro": 10, "Arriba": 6}   # numeración SSA de libass
CFG_SUB = os.path.join(OUT_DIR, "estilo_subtitulos.json")


def cargar_estilo():
    cfg = dict(SUB_DEFECTO)
    try:
        with open(CFG_SUB, encoding="utf-8") as f:
            cfg.update({k: v for k, v in json.load(f).items() if k in SUB_DEFECTO})
    except Exception:
        pass
    return cfg


def guardar_estilo(cfg):
    try:
        with open(CFG_SUB, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=1)
    except Exception:
        log.exception("guardar estilo")


def _ass_color(hexcolor, alpha=0):
    h = hexcolor.lstrip("#")
    r, g, b = h[0:2], h[2:4], h[4:6]
    return f"&H{alpha:02X}{b}{g}{r}".upper()


def estilo_ass(cfg):
    fuente = re.sub(r"[',:\\\\]", "", cfg["fuente"]) or "Arial"
    g = max(0, int(cfg["grosor"]))
    alfa_caja = int(round(255 * (1 - max(0, min(100, int(cfg["opacidad"]))) / 100)))
    partes = [f"FontName={fuente}", f"FontSize={int(cfg['tamano'])}",
              f"PrimaryColour={_ass_color(cfg['color'])}",
              f"Bold={-1 if cfg['negrita'] else 0}", f"Italic={-1 if cfg['cursiva'] else 0}",
              f"Alignment={POSICIONES_SUB.get(cfg['posicion'], 2)}",
              f"MarginV={int(cfg['margen'])}", "MarginL=30", "MarginR=30"]
    borde = cfg["color_borde"]
    forma = cfg["forma"]
    if forma == "Caja de fondo":
        partes += ["BorderStyle=3", f"Outline={max(1, g)}", "Shadow=0",
                   f"OutlineColour={_ass_color(borde, alfa_caja)}", f"BackColour={_ass_color(borde, alfa_caja)}"]
    elif forma == "Contorno":
        partes += ["BorderStyle=1", f"Outline={g}", "Shadow=0", f"OutlineColour={_ass_color(borde)}"]
    elif forma == "Sombra":
        partes += ["BorderStyle=1", "Outline=0", f"Shadow={max(1, g)}",
                   f"BackColour={_ass_color(borde, alfa_caja)}"]
    elif forma == "Contorno y sombra":
        partes += ["BorderStyle=1", f"Outline={g}", f"Shadow={max(1, g // 2 + 1)}",
                   f"OutlineColour={_ass_color(borde)}", f"BackColour={_ass_color(borde, alfa_caja)}"]
    else:
        partes += ["BorderStyle=1", "Outline=0", "Shadow=0"]
    return ",".join(partes)


def _filtro_subs(cfg):
    vf = "subtitles=subs.srt:charenc=UTF-8:force_style='%s'" % estilo_ass(cfg)
    if sys.platform == "win32":
        vf += ":fontsdir='C\\:/Windows/Fonts'"
    return vf


def _srt_estilo(segs, cfg):
    if cfg.get("mayusculas"):
        segs = [dict(s, text=s["text"].upper(), src=(s.get("src") or "").upper()) for s in segs]
    return a_srt(segs)


def ffmpeg_exe(instalar=True):
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        pass
    import shutil
    exe = shutil.which("ffmpeg")
    if exe or not instalar or CONGELADO:
        return exe
    # instala el componente automáticamente y vuelve a intentar
    import subprocess
    py = os.path.join(os.path.dirname(sys.executable), "python.exe")
    if not os.path.exists(py):
        py = sys.executable
    try:
        log.info("Instalando imageio-ffmpeg…")
        subprocess.run([py, "-m", "pip", "install", "-q", "imageio-ffmpeg"], check=True,
                       timeout=600, creationflags=_SIN_VENTANA)
        import importlib
        importlib.invalidate_caches()
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        log.exception("instalar imageio-ffmpeg")
        return None


def info_media(ruta):
    import av
    with av.open(ruta) as c:
        dur = (c.duration or 0) / 1_000_000
        return bool(c.streams.video), dur


_SIN_VENTANA = 0x08000000 if sys.platform == "win32" else 0


def vista_previa_subs(media, texto, cfg, t, salida_png, ancho=560):
    """Dibuja un cuadro del video con el subtítulo, usando el mismo motor del video final."""
    import subprocess
    import tempfile
    exe = ffmpeg_exe()
    if not exe:
        raise RuntimeError("Falta el componente de video. Ejecuta actualizar.bat.")
    tmp = tempfile.mkdtemp(prefix="transcriptor_prev_")
    with open(os.path.join(tmp, "subs.srt"), "w", encoding="utf-8") as f:
        f.write(_srt_estilo([{"start": 0, "end": 3600, "text": texto, "src": ""}], cfg))
    cmd = [exe, "-y", "-hide_banner", "-loglevel", "error", "-ss", f"{max(0, t):.2f}",
           "-i", os.path.abspath(media), "-frames:v", "1",
           "-vf", _filtro_subs(cfg) + f",scale={ancho}:-2", os.path.abspath(salida_png)]
    r = subprocess.run(cmd, cwd=tmp, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=60, creationflags=_SIN_VENTANA)
    if r.returncode != 0 or not os.path.exists(salida_png):
        raise RuntimeError(r.stderr.strip()[-400:] or "No se pudo generar la vista previa.")
    return salida_png


def video_subtitulado(media, segs, salida, cfg, progreso, stop):
    """Crea un MP4 con la transcripción estampada en la imagen."""
    import subprocess
    import tempfile
    exe = ffmpeg_exe()
    if not exe:
        raise RuntimeError("Falta el componente de video. Ejecuta actualizar.bat.")
    _, dur = info_media(media)
    tmp = tempfile.mkdtemp(prefix="transcriptor_")
    with open(os.path.join(tmp, "subs.srt"), "w", encoding="utf-8") as f:
        f.write(_srt_estilo(segs, cfg))
    cmd = [exe, "-y", "-hide_banner", "-i", os.path.abspath(media), "-vf", _filtro_subs(cfg),
           "-map", "0:v:0", "-map", "0:a?", "-c:v", "libx264", "-preset", "veryfast", "-crf", "20",
           "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "160k", "-movflags", "+faststart",
           os.path.abspath(salida)]
    log.info("ffmpeg: " + " ".join(cmd))
    proc = subprocess.Popen(cmd, cwd=tmp, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL,
                            stdin=subprocess.DEVNULL, text=True, encoding="utf-8", errors="replace",
                            creationflags=_SIN_VENTANA)
    ultimas = []
    for linea in proc.stderr:
        ultimas = (ultimas + [linea.strip()])[-15:]
        m = re.search(r"time=(\d+):(\d+):(\d+\.?\d*)", linea)
        if m and dur:
            t = int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3))
            progreso(min(100.0, t / dur * 100), t, dur)
        if stop.is_set():
            proc.terminate()
            break
    rc = proc.wait()
    if stop.is_set():
        try:
            os.remove(salida)
        except OSError:
            pass
        return False
    if rc != 0:
        log.error("ffmpeg falló:\n" + "\n".join(ultimas))
        raise RuntimeError("No se pudo crear el video:\n" + "\n".join(ultimas[-4:]))
    return True


def crear_menu_enviar_a(forzar=False):
    """Crea el acceso «Enviar a → Transcriptor» del clic derecho de Windows."""
    if sys.platform != "win32" or (CONGELADO and not forzar):
        return False   # en la versión instalada el instalador crea el acceso
    sendto = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "SendTo")
    lnk = os.path.join(sendto, "Transcriptor.lnk")
    if os.path.exists(lnk) and not forzar:
        return True
    exe = sys.executable
    pyw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    if os.path.exists(pyw):
        exe = pyw
    script = "" if CONGELADO else os.path.abspath(__file__)
    q = lambda x: "'" + x.replace("'", "''") + "'"
    ps = ("$s=(New-Object -ComObject WScript.Shell).CreateShortcut(" + q(lnk) + ");"
          "$s.TargetPath=" + q(exe) + ";"
          "$s.Arguments=" + q('"' + script + '"' if script else "") + ";"
          "$s.WorkingDirectory=" + q(os.path.dirname(script or exe)) + ";"
          "$s.Description='Transcribir con Transcriptor';$s.Save()")
    import subprocess
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                       check=True, timeout=30, creationflags=0x08000000)
        log.info("Acceso Enviar a creado: " + lnk)
        return True
    except Exception:
        log.exception("crear acceso Enviar a")
        return False


def autoprueba(ruta_informe):
    """Comprueba que el paquete (.exe) trae todo lo necesario. Devuelve 0 si está bien."""
    import importlib
    import wave
    lineas, ok = [], True
    for mod in ("numpy", "av", "ctranslate2", "faster_whisper", "onnxruntime", "sentencepiece",
                "huggingface_hub", "tkinterdnd2", "imageio_ffmpeg", "yt_dlp", "docx", "soundcard"):
        try:
            importlib.import_module(mod)
            lineas.append(f"OK {mod}")
        except Exception as e:
            lineas.append(f"ERROR {mod}: {e!r}")
            ok = False
    try:
        lineas.append("OK ffmpeg " + str(ffmpeg_exe(instalar=False)))
    except Exception as e:
        lineas.append(f"ERROR ffmpeg: {e!r}")
        ok = False
    try:
        prueba = os.path.join(OUT_DIR, "_autoprueba.wav")
        with wave.open(prueba, "wb") as w:
            w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR)
            w.writeframes((np.sin(np.arange(SR) * 0.1) * 5000).astype(np.int16).tobytes())
        lineas.append(f"OK lectura de audio ({len(decodificar(prueba)) / SR:.1f} s)")
        os.remove(prueba)
    except Exception as e:
        lineas.append(f"ERROR lectura de audio: {e!r}")
        ok = False
    try:
        from faster_whisper.vad import get_vad_model  # modelo VAD incluido en el paquete
        get_vad_model()
        lineas.append("OK modelo VAD")
    except Exception as e:
        lineas.append(f"ERROR modelo VAD: {e!r}")
        ok = False
    with open(ruta_informe, "w", encoding="utf-8") as f:
        f.write("\n".join(lineas) + "\n")
    return 0 if ok else 1


def diagnostico():
    def p(msg):
        print(msg, flush=True)
        log.info("[diag] " + msg)
    p(f"Python {sys.version.split()[0]} · {sys.platform}")
    for mod in ("numpy", "av", "ctranslate2", "faster_whisper", "soundcard", "yt_dlp", "docx"):
        try:
            m = __import__(mod)
            p(f"  OK  {mod} {getattr(m, '__version__', '')}")
        except Exception as e:
            p(f"  ERROR {mod}: {e}")
    import wave
    prueba = os.path.join(OUT_DIR, "_prueba.wav")
    with wave.open(prueba, "wb") as w:
        w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR)
        w.writeframes((np.sin(np.arange(SR * 2) * 2 * np.pi * 220 / SR) * 8000).astype(np.int16).tobytes())
    try:
        a = decodificar(prueba)
        p(f"  OK  lectura de audio ({len(a) / SR:.1f} s)")
    except Exception:
        p("  ERROR lectura de audio:\n" + traceback.format_exc())
    p("Cargando modelo «small» (si es la primera vez, se descarga ~500 MB)…")
    try:
        m = Motor()
        m.cargar("small", p)
        t0 = time.time()
        segs, info = m.archivo(prueba, "es")
        list(segs)
        p(f"  OK  modelo y transcripción ({time.time() - t0:.1f} s)")
    except Exception:
        p("  ERROR modelo:\n" + traceback.format_exc())
    try:
        _com_init()
        import soundcard as sc
        alt, mic = sc.default_speaker(), sc.default_microphone()
        p(f"  Altavoz predeterminado: {alt.name}")
        p(f"  Micrófono predeterminado: {mic.name if mic else '(ninguno)'}")
        for nombre, dev in (("audio del equipo", sc.get_microphone(id=str(alt.name), include_loopback=True)),
                            ("micrófono", mic)):
            if dev is None:
                continue
            res = {}

            def grabar():
                _com_init()
                with dev.recorder(samplerate=CAP_SR, blocksize=2400) as r:
                    res["d"] = r.record(numframes=CAP_SR)
            th = threading.Thread(target=grabar, daemon=True)
            th.start(); th.join(5)
            if "d" in res:
                p(f"  OK  captura {nombre} (nivel {float(np.abs(res['d']).max()):.4f})")
            else:
                p(f"  AVISO captura {nombre}: no llegó audio en 5 s (normal si no sonaba nada)")
    except Exception:
        p("  ERROR captura de audio:\n" + traceback.format_exc())
    try:
        import yt_dlp
        with yt_dlp.YoutubeDL({"quiet": True, "no_warnings": True, "noplaylist": True}) as y:
            info = y.extract_info("https://www.youtube.com/watch?v=jNQXAC9IVRw", download=False)
        fm = [f.get("format_id") for f in info.get("formats", [])]
        p(f"  OK  YouTube ({len(fm)} formatos; 18 disponible: {'18' in fm})")
    except Exception:
        p("  ERROR YouTube:\n" + traceback.format_exc())
    try:
        os.remove(prueba)
    except OSError:
        pass
    p(f"Diagnóstico terminado. Registro: {LOG}")


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    if "--autoprueba" in sys.argv:
        sys.exit(autoprueba(sys.argv[sys.argv.index("--autoprueba") + 1]
                            if len(sys.argv) > sys.argv.index("--autoprueba") + 1 else "autoprueba.txt"))
    if "--menu" in sys.argv:
        print("Acceso «Enviar a → Transcriptor»:", "creado" if crear_menu_enviar_a(True) else "no se pudo crear")
        return
    if "--diagnostico" in sys.argv:
        diagnostico()
        return
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()

    def _err(exc, val, tb):
        log.error("".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror(APP, f"{val}\n\nDetalle en {LOG}")
    root.report_callback_exception = _err
    log.info("Inicio de la aplicación")
    threading.Thread(target=crear_menu_enviar_a, daemon=True).start()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
