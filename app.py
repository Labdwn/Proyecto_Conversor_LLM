"""
Interfaz web local para el conversor .txt/.md -> .docx.

No requiere pip install de nada: usa unicamente la libreria estandar de
Python (http.server, webbrowser, json, base64, tempfile, subprocess).
El usuario solo necesita tener instalados Python y Pandoc.

Flujo:
  1. Este script levanta un servidor HTTP en 127.0.0.1 en un puerto libre.
  2. Abre el navegador por defecto en esa direccion.
  3. El usuario arrastra o selecciona su archivo en la pagina.
  4. El navegador lo sube via POST /convert; este script corre la misma
     logica de limpieza y verificacion que clean_and_convert.py y
     devuelve el .docx (en base64) + el reporte de verificacion.
  5. El navegador arma la descarga del .docx en el momento (sin volver
     a tocar el disco del lado del servidor mas que en un archivo temporal).
"""

import base64
import http.server
import json
import os
import re
import subprocess
import tempfile
import webbrowser

import clean_and_convert as cc  # misma logica de limpieza, sin tocarla


def convert_text_to_docx(raw_text):
    """Corre la misma secuencia que main() en clean_and_convert.py, pero
    trabajando en memoria/temporal y devolviendo (reporte, bytes_docx)
    en vez de escribir archivos permanentes."""

    cleaned = cc.clean_math_blocks(raw_text)

    with tempfile.TemporaryDirectory() as tmp:
        md_path = os.path.join(tmp, "documento.md")
        docx_path = os.path.join(tmp, "documento.docx")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(cleaned)

        result = subprocess.run(
            ["pandoc", "-f", "markdown+tex_math_single_backslash-raw_tex",
             md_path, "-o", docx_path],
            capture_output=True, text=True, encoding="utf-8"
        )
        if result.returncode != 0:
            raise RuntimeError(f"Pandoc fallo al convertir:\n{result.stderr}")

        # Misma logica de reporte (resumen + detalle) que usa clean_and_convert.py
        # desde la linea de comandos, para que ambas vias muestren lo mismo.
        report, _ = cc.build_report(
            raw_text, cleaned, docx_path, result,
            path_label="(archivo subido)", docx_label="documento.docx",
        )

        with open(docx_path, "rb") as f:
            docx_bytes = f.read()

    return report, docx_bytes


def parse_multipart(body, boundary):
    """Parser minimo de multipart/form-data (solo lo que necesitamos:
    un unico campo de archivo). Sin dependencias externas."""
    delimiter = b"--" + boundary
    parts = body.split(delimiter)
    for part in parts:
        if not part or part in (b"--\r\n", b"--"):
            continue
        if part.startswith(b"\r\n"):
            part = part[2:]
        if part.endswith(b"\r\n"):
            part = part[:-2]
        if not part:
            continue
        header_blob, sep, data = part.partition(b"\r\n\r\n")
        if not sep:
            continue
        headers = header_blob.decode("utf-8", errors="replace")
        if "filename=" not in headers:
            continue
        m = re.search(r'filename="([^"]*)"', headers)
        filename = m.group(1) if m else None
        return {"filename": filename, "data": data}
    return None


HTML_PAGE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Conversor .txt / .md a .docx</title>
<style>
  body { font-family: system-ui, sans-serif; max-width: 640px; margin: 40px auto; padding: 0 16px; color: #222; }
  h1 { font-size: 1.25rem; }
  h2 { font-size: 0.95rem; margin: 28px 0 10px; color: #444; }
  label { display: block; font-size: 0.8rem; font-weight: 600; color: #444; margin-bottom: 4px; }
  .field { margin-bottom: 14px; }
  .note { font-size: 0.8rem; color: #666; margin: -4px 0 16px; }
  input[type="text"], select {
    width: 100%; box-sizing: border-box; padding: 8px 10px; font-size: 0.95rem;
    border: 1px solid #ccc; border-radius: 6px; background: white;
  }
  input[type="text"]:invalid, input.invalid { border-color: #b00020; }
  #drop { border: 2px dashed #999; border-radius: 10px; padding: 48px 16px; text-align: center; cursor: pointer; transition: border-color .15s, background .15s; margin-top: 8px; }
  #drop.hover { border-color: #3366cc; background: #f0f5ff; }
  #drop p { margin: 0; color: #555; }
  #filename { margin-top: 12px; font-weight: 600; }
  #filenamePreview { margin-top: 10px; font-size: 0.85rem; color: #555; }
  #filenamePreview code { background: #f0f0f0; padding: 2px 6px; border-radius: 4px; }
  #hint { margin-top: 10px; font-size: 0.85rem; color: #b00020; min-height: 1.2em; }
  button { margin-top: 16px; padding: 10px 20px; font-size: 1rem; border: none; border-radius: 6px; background: #3366cc; color: white; cursor: pointer; }
  button:disabled { background: #aaa; cursor: default; }
  #status { margin-top: 16px; }
  #report { white-space: pre-wrap; background: #f7f7f7; border: 1px solid #ddd; border-radius: 6px; padding: 12px; margin-top: 16px; font-size: 0.85rem; max-height: 340px; overflow-y: auto; }
  .alerta { color: #b00020; }
  a.download { display: inline-block; margin-top: 12px; padding: 10px 20px; background: #1a8a3d; color: white; text-decoration: none; border-radius: 6px; }
  #timestamp { margin-top: 14px; font-size: 0.9rem; }
  #timestamp button { margin-top: 0; margin-left: 8px; padding: 4px 12px; font-size: 0.85rem; }
</style>
</head>
<body>
  <h1>Conversor .txt / .md a .docx</h1>

  <h2>Datos del archivo</h2>
  <p class="note">Ningun campo es obligatorio. Si dejas todo vacio, se conserva el nombre del archivo subido. Si completas alguno, hay que completar el resto (la version es la unica excepcion).</p>
  <div class="field">
    <label for="problemId">Nombre del problema</label>
    <input type="text" id="problemId" placeholder="ej. FIS-FI2-001">
  </div>
  <div class="field">
    <label for="iaSelect">IA</label>
    <select id="iaSelect">
      <option value="">-- (opcional) --</option>
      <option value="ChatGPT">ChatGPT</option>
      <option value="Claude">Claude</option>
      <option value="Copilot">Copilot</option>
      <option value="DeepSeek">DeepSeek</option>
      <option value="Gemini">Gemini</option>
      <option value="MathGPT">MathGPT</option>
      <option value="Meta">Meta</option>
      <option value="Grok">Grok</option>
    </select>
  </div>
  <div class="field">
    <label for="versionInput">Version del modelo (opcional siempre)</label>
    <input type="text" id="versionInput" placeholder="ej. 5.6">
  </div>
  <div class="field">
    <label for="asistenteInput">Asistente (A + numero)</label>
    <input type="text" id="asistenteInput" placeholder="ej. A1">
  </div>
  <div class="field">
    <label for="numSelect">Origen de la respuesta</label>
    <select id="numSelect">
      <option value="">-- (opcional) --</option>
      <option value="1">1 - Respuesta a un PDF</option>
      <option value="2">2 - Respuesta a una imagen</option>
    </select>
  </div>
  <div id="filenamePreview"></div>

  <h2>Archivo a convertir</h2>
  <div id="drop">
    <p>Arrastra tu archivo aqui, o hace clic para elegirlo</p>
    <div id="filename"></div>
  </div>
  <input type="file" id="fileInput" accept=".txt,.md" style="display:none">
  <div id="hint"></div>
  <button id="convertBtn" disabled>Convertir</button>
  <div id="status"></div>
  <div id="result"></div>

<script>
const drop = document.getElementById('drop');
const fileInput = document.getElementById('fileInput');
const filenameEl = document.getElementById('filename');
const convertBtn = document.getElementById('convertBtn');
const statusEl = document.getElementById('status');
const resultEl = document.getElementById('result');
const hintEl = document.getElementById('hint');
const previewEl = document.getElementById('filenamePreview');

const problemIdEl = document.getElementById('problemId');
const iaSelectEl = document.getElementById('iaSelect');
const versionEl = document.getElementById('versionInput');
const asistenteEl = document.getElementById('asistenteInput');
const numSelectEl = document.getElementById('numSelect');

let selectedFile = null;

function sanitize(s) {
  return s.trim().replace(/[\\\\\\/:*?"<>|]/g, '-');
}

function pad2(n) { return n.toString().padStart(2, '0'); }

function isFilled(v) { return v.trim() !== ''; }

// version no cuenta para decidir si el usuario "empezo a llenar" el formulario
function anyMetadataFilled() {
  return isFilled(problemIdEl.value) || isFilled(iaSelectEl.value) ||
         isFilled(versionEl.value) || isFilled(asistenteEl.value) ||
         isFilled(numSelectEl.value);
}

function getMissing() {
  const missing = [];
  if (!selectedFile) missing.push('archivo');
  if (anyMetadataFilled()) {
    if (!isFilled(problemIdEl.value)) missing.push('nombre del problema');
    if (!isFilled(iaSelectEl.value)) missing.push('IA');
    if (!/^A\\d+$/i.test(asistenteEl.value.trim())) missing.push('asistente (formato A+numero, ej. A7)');
    if (!isFilled(numSelectEl.value)) missing.push('origen (PDF/imagen)');
  }
  return missing;
}

function buildFilename() {
  if (!anyMetadataFilled()) {
    // nada lleno: se conserva el nombre del archivo subido
    const base = selectedFile ? selectedFile.name.replace(/\\.[^.]+$/, '') : 'documento';
    return sanitize(base) + '.docx';
  }
  const problemId = sanitize(problemIdEl.value);
  const ia = iaSelectEl.value.trim();
  const version = sanitize(versionEl.value);
  const asistente = sanitize(asistenteEl.value.toUpperCase());
  const num = numSelectEl.value;
  const modelo = version ? `${ia}-${version}` : ia;
  return `${problemId}_${modelo}_${asistente}_${num}.docx`;
}

function updateState() {
  const missing = getMissing();
  if (missing.length) {
    convertBtn.disabled = true;
    hintEl.textContent = missing.length === 1 && missing[0] === 'archivo'
      ? 'Selecciona un archivo para continuar.'
      : 'Falta completar: ' + missing.join(', ') + '.';
    previewEl.textContent = '';
  } else {
    convertBtn.disabled = false;
    hintEl.textContent = '';
    const suffix = anyMetadataFilled() ? '' : ' (nombre original del archivo)';
    previewEl.innerHTML = 'Se descargara como: <code>' + buildFilename() + '</code>' + suffix;
  }
}

[problemIdEl, iaSelectEl, versionEl, asistenteEl, numSelectEl].forEach(el => {
  el.addEventListener('input', updateState);
  el.addEventListener('change', updateState);
});

drop.addEventListener('click', () => fileInput.click());
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('hover'); });
drop.addEventListener('dragleave', () => drop.classList.remove('hover'));
drop.addEventListener('drop', e => {
  e.preventDefault();
  drop.classList.remove('hover');
  if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) setFile(fileInput.files[0]);
});

function setFile(f) {
  selectedFile = f;
  filenameEl.textContent = f.name;
  resultEl.innerHTML = '';
  statusEl.textContent = '';
  updateState();
}

updateState();

convertBtn.addEventListener('click', async () => {
  if (!selectedFile || getMissing().length) return;
  convertBtn.disabled = true;
  statusEl.textContent = 'Convirtiendo...';
  resultEl.innerHTML = '';

  const formData = new FormData();
  formData.append('file', selectedFile);

  try {
    const resp = await fetch('/convert', { method: 'POST', body: formData });
    const data = await resp.json();
    if (!data.ok) {
      statusEl.innerHTML = '<span class="alerta">' + data.message + '</span>';
      convertBtn.disabled = false;
      return;
    }
    statusEl.textContent = 'Listo.';
    const bytes = atob(data.docx_b64);
    const arr = new Uint8Array(bytes.length);
    for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
    const blob = new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const url = URL.createObjectURL(blob);
    const outName = buildFilename();

    const link = document.createElement('a');
    link.href = url;
    link.download = outName;
    link.className = 'download';
    link.textContent = 'Descargar ' + outName;
    resultEl.appendChild(link);

    const now = new Date();
    const fecha = `${pad2(now.getDate())}/${pad2(now.getMonth() + 1)}/${now.getFullYear()}`;
    const hora = `${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
    const tsDiv = document.createElement('div');
    tsDiv.id = 'timestamp';
    tsDiv.innerHTML = '<span id="tsText">Generado: ' + fecha + '  ' + hora + '</span>' +
      '<button id="copyTsBtn" type="button">Copiar</button>';
    resultEl.appendChild(tsDiv);
    document.getElementById('copyTsBtn').addEventListener('click', () => {
      navigator.clipboard.writeText(fecha + '  ' + hora).then(() => {
        const btn = document.getElementById('copyTsBtn');
        btn.textContent = 'Copiado';
        setTimeout(() => { btn.textContent = 'Copiar'; }, 1500);
      });
    });

    const reportDiv = document.createElement('div');
    reportDiv.id = 'report';
    reportDiv.textContent = data.report;
    resultEl.appendChild(reportDiv);

    convertBtn.disabled = false;
  } catch (err) {
    statusEl.innerHTML = '<span class="alerta">Error de conexion: ' + err + '</span>';
    convertBtn.disabled = false;
  }
});
</script>
</body>
</html>
"""


class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # silenciar el log de accesos en la consola

    def do_GET(self):
        if self.path == "/":
            page = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(page)))
            self.end_headers()
            self.wfile.write(page)
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path != "/convert":
            self.send_error(404)
            return

        content_type = self.headers.get("Content-Type", "")
        if "boundary=" not in content_type:
            self.send_json_error("Solicitud invalida (falta boundary).")
            return
        boundary = content_type.split("boundary=")[1].strip()
        if boundary.startswith('"') and boundary.endswith('"'):
            boundary = boundary[1:-1]

        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        parsed = parse_multipart(body, boundary.encode())
        if not parsed or not parsed.get("data"):
            self.send_json_error("No se recibio ningun archivo.")
            return

        try:
            raw_text = parsed["data"].decode("utf-8")
        except UnicodeDecodeError:
            self.send_json_error("El archivo no parece ser texto UTF-8 valido.")
            return

        try:
            report, docx_bytes = convert_text_to_docx(raw_text)
        except FileNotFoundError:
            self.send_json_error("No se encontro Pandoc. Confirma que este instalado y en el PATH.")
            return
        except Exception as e:
            self.send_json_error(f"Error inesperado durante la conversion: {e}")
            return

        payload = json.dumps({
            "ok": True,
            "report": report,
            "docx_b64": base64.b64encode(docx_bytes).decode("ascii"),
        }).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def send_json_error(self, message):
        payload = json.dumps({"ok": False, "message": message}).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def main():
    httpd = http.server.HTTPServer(("127.0.0.1", 0), Handler)
    port = httpd.server_port
    url = f"http://127.0.0.1:{port}/"
    print(f"Servidor local iniciado en {url}")
    print("No cierres esta ventana mientras uses el conversor.")
    webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()