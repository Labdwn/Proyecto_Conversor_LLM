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
import io
import json
import os
import re
import subprocess
import tempfile
import webbrowser
import zipfile

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


def parse_multipart_all(body, boundary):
    """Como parse_multipart, pero devuelve TODAS las partes que traen archivo
    (para /convert-batch), en el mismo orden en que llegaron."""
    delimiter = b"--" + boundary
    parts = body.split(delimiter)
    files = []
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
        files.append({"filename": filename, "data": data})
    return files


def sanitize_zip_name(name):
    """Limpia un nombre para que sea seguro como entrada de zip / archivo en
    disco: sin separadores de ruta ni caracteres invalidos en Windows."""
    name = os.path.basename(name or "documento")
    name = re.sub(r'[\\/:*?"<>|]', '-', name).strip()
    return name or "documento"


def convert_batch(files):
    """Convierte una lista de {'filename', 'data'} (nombre final deseado +
    contenido crudo del .md/.txt) y arma un .zip en memoria con todos los
    .docx generados mas un reporte.txt combinado.

    Devuelve (zip_bytes, resultados) donde resultados es una lista de
    {"filename": str, "ok": bool, "message": str} en el mismo orden de
    entrada, para que el cliente pueda mostrar que archivo fallo (si alguno)."""
    resultados = []
    reporte_partes = []
    buffer = io.BytesIO()
    usados = {}

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in files:
            nombre_original = item.get("filename") or "documento"
            base = sanitize_zip_name(os.path.splitext(nombre_original)[0])

            try:
                raw_text = item["data"].decode("utf-8")
            except UnicodeDecodeError:
                resultados.append({
                    "filename": nombre_original, "ok": False,
                    "message": "El archivo no parece ser texto UTF-8 valido.",
                })
                reporte_partes.append(
                    f"### {nombre_original}\n[ERROR] El archivo no parece ser texto UTF-8 valido.\n")
                continue

            try:
                report, docx_bytes = convert_text_to_docx(raw_text)
            except FileNotFoundError:
                resultados.append({
                    "filename": nombre_original, "ok": False,
                    "message": "No se encontro Pandoc. Confirma que este instalado y en el PATH.",
                })
                reporte_partes.append(
                    f"### {nombre_original}\n[ERROR] No se encontro Pandoc.\n")
                continue
            except Exception as e:
                resultados.append({
                    "filename": nombre_original, "ok": False,
                    "message": f"Error inesperado durante la conversion: {e}",
                })
                reporte_partes.append(
                    f"### {nombre_original}\n[ERROR] {e}\n")
                continue

            # Evitar colisiones de nombre dentro del zip (ej. dos archivos
            # que terminaron con el mismo problema+IA+asistente).
            nombre_zip = base
            if nombre_zip in usados:
                usados[nombre_zip] += 1
                nombre_zip = f"{base}_{usados[nombre_zip]}"
            else:
                usados[nombre_zip] = 0
            nombre_zip += ".docx"

            zf.writestr(nombre_zip, docx_bytes)
            resultados.append({"filename": nombre_zip, "ok": True, "message": "Convertido."})
            reporte_partes.append(f"### {nombre_zip}\n{report}\n")

        zf.writestr("reporte.txt", "\n".join(reporte_partes))

    return buffer.getvalue(), resultados


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

  /* --- Modo lote --- */
  .batch-field { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; flex-wrap: wrap; }
  .batch-field label.fieldLabel { flex: 0 0 150px; margin-bottom: 0; }
  .batch-field .toggleWrap { display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: #555; font-weight: 400; }
  .batch-field .toggleWrap input { width: auto; }
  .batch-field .sharedInput { flex: 1; min-width: 160px; }
  .batch-field .sharedInput input[type="text"], .batch-field .sharedInput select { width: 100%; }
  table#batchTable { width: 100%; border-collapse: collapse; margin-top: 14px; font-size: 0.85rem; }
  table#batchTable th, table#batchTable td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: middle; }
  table#batchTable th { background: #f0f0f0; font-size: 0.75rem; }
  table#batchTable input[type="text"], table#batchTable select { padding: 4px 6px; font-size: 0.82rem; }
  table#batchTable td.filecell { font-size: 0.8rem; color: #444; max-width: 160px; overflow-wrap: anywhere; }
  table#batchTable td.previewcell { font-size: 0.75rem; color: #666; max-width: 180px; overflow-wrap: anywhere; }
  #batchResults { margin-top: 14px; font-size: 0.85rem; }
  #batchResults .row-ok { color: #1a8a3d; }
  #batchResults .row-err { color: #b00020; }
  #modeNote { font-size: 0.8rem; color: #666; margin-top: 4px; }
</style>
</head>
<body>
  <h1>Conversor .txt / .md a .docx</h1>

  <h2>Archivo(s) a convertir</h2>
  <p class="note">Podes soltar un solo archivo, o varios a la vez para convertirlos en lote.</p>
  <div id="drop">
    <p>Arrastra tu(s) archivo(s) aqui, o hace clic para elegirlos</p>
    <div id="filename"></div>
  </div>
  <input type="file" id="fileInput" accept=".txt,.md" style="display:none" multiple>
  <div id="hint"></div>

  <div id="singleModeSection">
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
  </div>

  <div id="batchSection" style="display:none">
    <h2>Datos del lote (<span id="batchCount">0</span> archivos)</h2>
    <p class="note">Para cada dato, elegi si aplica igual a todos los archivos o si varia por archivo. Los que varian se llenan en la tabla de abajo.</p>

    <div class="batch-field">
      <label class="fieldLabel">Nombre del problema</label>
      <span class="toggleWrap"><input type="checkbox" id="fixProblemId" checked> Mismo para todos</span>
      <span class="sharedInput" id="sharedProblemIdWrap"><input type="text" id="sharedProblemId" placeholder="ej. FIS-FI2-001"></span>
    </div>
    <div class="batch-field">
      <label class="fieldLabel">IA</label>
      <span class="toggleWrap"><input type="checkbox" id="fixIa"> Misma para todos</span>
      <span class="sharedInput" id="sharedIaWrap" style="display:none">
        <select id="sharedIa">
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
      </span>
    </div>
    <div class="batch-field">
      <label class="fieldLabel">Version (opcional)</label>
      <span class="toggleWrap"><input type="checkbox" id="fixVersion" checked> Misma para todos</span>
      <span class="sharedInput" id="sharedVersionWrap"><input type="text" id="sharedVersion" placeholder="ej. 5.6"></span>
    </div>
    <div class="batch-field">
      <label class="fieldLabel">Asistente (A+num)</label>
      <span class="toggleWrap"><input type="checkbox" id="fixAsistente"> Mismo para todos</span>
      <span class="sharedInput" id="sharedAsistenteWrap" style="display:none"><input type="text" id="sharedAsistente" placeholder="ej. A1"></span>
    </div>
    <div class="batch-field">
      <label class="fieldLabel">Origen</label>
      <span class="toggleWrap"><input type="checkbox" id="fixNum" checked> Mismo para todos</span>
      <span class="sharedInput" id="sharedNumWrap">
        <select id="sharedNum">
          <option value="">-- (opcional) --</option>
          <option value="1">1 - Respuesta a un PDF</option>
          <option value="2">2 - Respuesta a una imagen</option>
        </select>
      </span>
    </div>

    <table id="batchTable">
      <thead><tr id="batchTableHead"></tr></thead>
      <tbody id="batchTableBody"></tbody>
    </table>
  </div>

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

const singleModeSectionEl = document.getElementById('singleModeSection');
const batchSectionEl = document.getElementById('batchSection');
const batchCountEl = document.getElementById('batchCount');

const fixProblemIdEl = document.getElementById('fixProblemId');
const sharedProblemIdEl = document.getElementById('sharedProblemId');
const sharedProblemIdWrapEl = document.getElementById('sharedProblemIdWrap');
const fixIaEl = document.getElementById('fixIa');
const sharedIaEl = document.getElementById('sharedIa');
const sharedIaWrapEl = document.getElementById('sharedIaWrap');
const fixVersionEl = document.getElementById('fixVersion');
const sharedVersionEl = document.getElementById('sharedVersion');
const sharedVersionWrapEl = document.getElementById('sharedVersionWrap');
const fixAsistenteEl = document.getElementById('fixAsistente');
const sharedAsistenteEl = document.getElementById('sharedAsistente');
const sharedAsistenteWrapEl = document.getElementById('sharedAsistenteWrap');
const fixNumEl = document.getElementById('fixNum');
const sharedNumEl = document.getElementById('sharedNum');
const sharedNumWrapEl = document.getElementById('sharedNumWrap');

const fieldDefs = {
  problemId: { fixedEl: fixProblemIdEl, sharedEl: sharedProblemIdEl, wrapEl: sharedProblemIdWrapEl, label: 'Problema', inputType: 'text' },
  ia:        { fixedEl: fixIaEl,        sharedEl: sharedIaEl,        wrapEl: sharedIaWrapEl,        label: 'IA',        inputType: 'select' },
  version:   { fixedEl: fixVersionEl,   sharedEl: sharedVersionEl,   wrapEl: sharedVersionWrapEl,   label: 'Version',   inputType: 'text' },
  asistente: { fixedEl: fixAsistenteEl, sharedEl: sharedAsistenteEl, wrapEl: sharedAsistenteWrapEl, label: 'Asistente', inputType: 'text' },
  num:       { fixedEl: fixNumEl,       sharedEl: sharedNumEl,       wrapEl: sharedNumWrapEl,       label: 'Origen',    inputType: 'select' },
};
const FIELD_ORDER = ['problemId', 'ia', 'version', 'asistente', 'num'];

let selectedFile = null;
let selectedFiles = [];
let batchMode = false;
let rowData = [];

function sanitize(s) {
  return s.trim().replace(/[\\\\\\/:*?"<>|]/g, '-');
}

function pad2(n) { return n.toString().padStart(2, '0'); }

function renderTimestamp(container) {
  const now = new Date();
  const fecha = `${pad2(now.getDate())}/${pad2(now.getMonth() + 1)}/${now.getFullYear()}`;
  const hora = `${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
  const tsDiv = document.createElement('div');
  tsDiv.id = 'timestamp';
  tsDiv.innerHTML = '<span id="tsText">Generado: ' + fecha + '  ' + hora + '</span>' +
    '<button id="copyTsBtn" type="button">Copiar</button>';
  container.appendChild(tsDiv);
  tsDiv.querySelector('#copyTsBtn').addEventListener('click', () => {
    navigator.clipboard.writeText(fecha + '  ' + hora).then(() => {
      const btn = tsDiv.querySelector('#copyTsBtn');
      btn.textContent = 'Copiado';
      setTimeout(() => { btn.textContent = 'Copiar'; }, 1500);
    });
  });
}

function isFilled(v) { return (v || '').trim() !== ''; }

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

// --- Modo lote ---

function getFieldValue(i, key) {
  const def = fieldDefs[key];
  if (def.fixedEl.checked) return def.sharedEl.value;
  const row = rowData[i] || {};
  return row[key] || '';
}

function buildBatchFilename(i) {
  const file = selectedFiles[i];
  const problemId = getFieldValue(i, 'problemId');
  const ia = getFieldValue(i, 'ia');
  const version = getFieldValue(i, 'version');
  const asistente = getFieldValue(i, 'asistente');
  const num = getFieldValue(i, 'num');
  const anyFilled = isFilled(problemId) || isFilled(ia) || isFilled(version) || isFilled(asistente) || isFilled(num);
  if (!anyFilled) {
    const base = file.name.replace(/\\.[^.]+$/, '');
    return sanitize(base) + '.docx';
  }
  const pid = sanitize(problemId);
  const iaClean = ia.trim();
  const ver = sanitize(version);
  const asis = sanitize(asistente.toUpperCase());
  const modelo = ver ? `${iaClean}-${ver}` : iaClean;
  return `${pid}_${modelo}_${asis}_${num}.docx`;
}

function computeAllFilenames() {
  const used = {};
  const names = [];
  for (let i = 0; i < selectedFiles.length; i++) {
    let base = buildBatchFilename(i).replace(/\\.docx$/, '');
    if (used[base] !== undefined) {
      used[base]++;
      base = base + '_' + used[base];
    } else {
      used[base] = 0;
    }
    names.push(base + '.docx');
  }
  return names;
}

function getMissingBatch() {
  const missing = [];
  if (!selectedFiles.length) { missing.push('archivos'); return missing; }
  let incompleteCount = 0;
  for (let i = 0; i < selectedFiles.length; i++) {
    const problemId = getFieldValue(i, 'problemId');
    const ia = getFieldValue(i, 'ia');
    const version = getFieldValue(i, 'version');
    const asistente = getFieldValue(i, 'asistente');
    const num = getFieldValue(i, 'num');
    const anyFilled = isFilled(problemId) || isFilled(ia) || isFilled(version) || isFilled(asistente) || isFilled(num);
    if (anyFilled) {
      const ok = isFilled(problemId) && isFilled(ia) && /^A\\d+$/i.test(asistente.trim()) && isFilled(num);
      if (!ok) incompleteCount++;
    }
  }
  if (incompleteCount > 0) missing.push(incompleteCount + ' archivo(s) con datos incompletos');
  return missing;
}

function updateAllPreviews() {
  const names = computeAllFilenames();
  names.forEach((name, i) => {
    const cell = document.getElementById('preview-' + i);
    if (cell) cell.textContent = name;
  });
}

function updateBatchState() {
  const missing = getMissingBatch();
  if (missing.length) {
    convertBtn.disabled = true;
    hintEl.textContent = missing[0] === 'archivos'
      ? 'Selecciona varios archivos para continuar.'
      : 'Falta completar: ' + missing.join(', ') + '.';
  } else {
    convertBtn.disabled = false;
    hintEl.textContent = '';
  }
  updateAllPreviews();
}

function renderBatchTable() {
  const theadRow = document.getElementById('batchTableHead');
  const tbody = document.getElementById('batchTableBody');
  theadRow.innerHTML = '';
  tbody.innerHTML = '';

  const varyingFields = FIELD_ORDER.filter(k => !fieldDefs[k].fixedEl.checked);

  const thFile = document.createElement('th');
  thFile.textContent = 'Archivo';
  theadRow.appendChild(thFile);
  varyingFields.forEach(k => {
    const th = document.createElement('th');
    th.textContent = fieldDefs[k].label;
    theadRow.appendChild(th);
  });
  const thPreview = document.createElement('th');
  thPreview.textContent = 'Nombre final';
  theadRow.appendChild(thPreview);

  selectedFiles.forEach((file, i) => {
    const tr = document.createElement('tr');
    const tdFile = document.createElement('td');
    tdFile.className = 'filecell';
    tdFile.textContent = file.name;
    tr.appendChild(tdFile);

    varyingFields.forEach(k => {
      const def = fieldDefs[k];
      const td = document.createElement('td');
      let input;
      if (def.inputType === 'select') {
        input = document.createElement('select');
        input.innerHTML = def.sharedEl.innerHTML;
      } else {
        input = document.createElement('input');
        input.type = 'text';
        if (k === 'asistente') input.placeholder = 'A1';
      }
      input.value = (rowData[i] && rowData[i][k]) || '';
      input.addEventListener('input', () => { rowData[i][k] = input.value; updateBatchState(); });
      input.addEventListener('change', () => { rowData[i][k] = input.value; updateBatchState(); });
      td.appendChild(input);
      tr.appendChild(td);
    });

    const tdPreview = document.createElement('td');
    tdPreview.className = 'previewcell';
    tdPreview.id = 'preview-' + i;
    tr.appendChild(tdPreview);

    tbody.appendChild(tr);
  });

  updateAllPreviews();
}

[['fixProblemId', 'sharedProblemIdWrap'], ['fixIa', 'sharedIaWrap'], ['fixVersion', 'sharedVersionWrap'],
 ['fixAsistente', 'sharedAsistenteWrap'], ['fixNum', 'sharedNumWrap']].forEach(([fixId, wrapId]) => {
  const fixEl = document.getElementById(fixId);
  const wrapEl = document.getElementById(wrapId);
  fixEl.addEventListener('change', () => {
    wrapEl.style.display = fixEl.checked ? '' : 'none';
    renderBatchTable();
    updateBatchState();
  });
});

[sharedProblemIdEl, sharedIaEl, sharedVersionEl, sharedAsistenteEl, sharedNumEl].forEach(el => {
  el.addEventListener('input', updateBatchState);
  el.addEventListener('change', updateBatchState);
});

// --- Seleccion de archivos (single o lote) ---

drop.addEventListener('click', () => fileInput.click());
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('hover'); });
drop.addEventListener('dragleave', () => drop.classList.remove('hover'));
drop.addEventListener('drop', e => {
  e.preventDefault();
  drop.classList.remove('hover');
  if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) handleFiles(fileInput.files);
});

function handleFiles(fileList) {
  const files = Array.from(fileList);
  if (!files.length) return;
  resultEl.innerHTML = '';
  statusEl.textContent = '';

  if (files.length === 1) {
    batchMode = false;
    selectedFile = files[0];
    selectedFiles = [files[0]];
    filenameEl.textContent = files[0].name;
    singleModeSectionEl.style.display = '';
    batchSectionEl.style.display = 'none';
    updateState();
  } else {
    batchMode = true;
    selectedFile = null;
    selectedFiles = files;
    rowData = files.map(() => ({ problemId: '', ia: '', version: '', asistente: '', num: '' }));
    filenameEl.textContent = files.length + ' archivos seleccionados';
    singleModeSectionEl.style.display = 'none';
    batchSectionEl.style.display = '';
    batchCountEl.textContent = files.length;
    renderBatchTable();
    updateBatchState();
  }
}

updateState();

convertBtn.addEventListener('click', async () => {
  if (batchMode) {
    if (!selectedFiles.length || getMissingBatch().length) return;
    await convertBatchFiles();
  } else {
    if (!selectedFile || getMissing().length) return;
    await convertSingleFile();
  }
});

async function convertSingleFile() {
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

    renderTimestamp(resultEl);

    const reportDiv = document.createElement('div');
    reportDiv.id = 'report';
    reportDiv.textContent = data.report;
    resultEl.appendChild(reportDiv);

    convertBtn.disabled = false;
  } catch (err) {
    statusEl.innerHTML = '<span class="alerta">Error de conexion: ' + err + '</span>';
    convertBtn.disabled = false;
  }
}

async function convertBatchFiles() {
  convertBtn.disabled = true;
  statusEl.textContent = 'Convirtiendo lote...';
  resultEl.innerHTML = '';

  const names = computeAllFilenames();
  const formData = new FormData();
  for (let i = 0; i < selectedFiles.length; i++) {
    formData.append('files', selectedFiles[i], names[i]);
  }

  try {
    const resp = await fetch('/convert-batch', { method: 'POST', body: formData });
    const data = await resp.json();
    if (!data.ok) {
      statusEl.innerHTML = '<span class="alerta">' + (data.message || 'Error') + '</span>';
      convertBtn.disabled = false;
      return;
    }
    statusEl.textContent = 'Listo.';
    const bytes = atob(data.zip_b64);
    const arr = new Uint8Array(bytes.length);
    for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
    const blob = new Blob([arr], { type: 'application/zip' });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = 'lote.zip';
    link.className = 'download';
    link.textContent = 'Descargar lote.zip (' + data.results.length + ' archivo(s), incluye reporte.txt)';
    resultEl.appendChild(link);

    renderTimestamp(resultEl);

    const resDiv = document.createElement('div');
    resDiv.id = 'batchResults';
    data.results.forEach(r => {
      const row = document.createElement('div');
      row.className = r.ok ? 'row-ok' : 'row-err';
      row.textContent = (r.ok ? 'OK - ' : 'ERROR - ') + r.filename + (r.ok ? '' : ': ' + r.message);
      resDiv.appendChild(row);
    });
    resultEl.appendChild(resDiv);

    convertBtn.disabled = false;
  } catch (err) {
    statusEl.innerHTML = '<span class="alerta">Error de conexion: ' + err + '</span>';
    convertBtn.disabled = false;
  }
}
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
        if self.path == "/convert":
            self._handle_convert()
        elif self.path == "/convert-batch":
            self._handle_convert_batch()
        else:
            self.send_error(404)

    def _read_multipart_body(self):
        """Devuelve (boundary_bytes, body) o None si la solicitud es invalida
        (ya se encarga de mandar el error JSON en ese caso)."""
        content_type = self.headers.get("Content-Type", "")
        if "boundary=" not in content_type:
            self.send_json_error("Solicitud invalida (falta boundary).")
            return None
        boundary = content_type.split("boundary=")[1].strip()
        if boundary.startswith('"') and boundary.endswith('"'):
            boundary = boundary[1:-1]
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        return boundary.encode(), body

    def _handle_convert(self):
        parsed_boundary = self._read_multipart_body()
        if parsed_boundary is None:
            return
        boundary, body = parsed_boundary

        parsed = parse_multipart(body, boundary)
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

    def _handle_convert_batch(self):
        parsed_boundary = self._read_multipart_body()
        if parsed_boundary is None:
            return
        boundary, body = parsed_boundary

        files = parse_multipart_all(body, boundary)
        if not files:
            self.send_json_error("No se recibio ningun archivo.")
            return

        zip_bytes, resultados = convert_batch(files)

        payload = json.dumps({
            "ok": True,
            "results": resultados,
            "zip_b64": base64.b64encode(zip_bytes).decode("ascii"),
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