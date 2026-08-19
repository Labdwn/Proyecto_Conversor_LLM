import sys, re, subprocess, os

# Comandos de fuente viejos de TeX que Pandoc NO entiende dentro de $...$ o $$...$$.
# Se mapean a su equivalente moderno que si soporta.
FONT_CMDS = {
    'rm': 'mathrm',
    'bf': 'mathbf',
    'it': 'mathit',
    'sf': 'mathsf',
    'tt': 'mathtt',
    'cal': 'mathcal',
    'frak': 'mathfrak',
    'Bbb': 'mathbb',
}

def find_matching_brace(text, start):
    """start apunta justo despues de un '{'. Devuelve el indice del '}' que cierra,
    respetando llaves anidadas. Si no encuentra cierre, devuelve len(text)."""
    depth = 1
    i = start
    while i < len(text):
        if text[i] == '{':
            depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return len(text)  # sin cierre: se toma hasta el final (no se pierde texto)

def fix_font_commands(s):
    """Reemplaza \\rm{...}, \\rm ..., \\bf{...}, etc. por \\mathrm{...}, \\mathbf{...}, etc.
    Maneja llaves anidadas correctamente y comandos sin llaves (aplican hasta el
    proximo '}' que los cierre implicitamente, o hasta el final de la formula)."""
    cmd_pattern = re.compile(r'\\(' + '|'.join(FONT_CMDS.keys()) + r')\b\s*')
    out = []
    i = 0
    while True:
        m = cmd_pattern.search(s, i)
        if not m:
            out.append(s[i:])
            break
        out.append(s[i:m.start()])
        cmd = m.group(1)
        new_cmd = FONT_CMDS[cmd]
        pos = m.end()
        if pos < len(s) and s[pos] == '{':
            # Caso {\rm ...} ya con llave propia justo despues -> \mathrm{...}
            close = find_matching_brace(s, pos + 1)
            inner = s[pos + 1:close]
            inner = fix_font_commands(inner)  # por si hay comandos anidados dentro
            out.append(f'\\{new_cmd}{{{inner}}}')
            i = close + 1
        else:
            # Caso sin llave propia: \rm aplica hasta el '}' que lo contiene (si lo hay)
            # o hasta el final de la formula.
            depth = 0
            j = pos
            while j < len(s):
                if s[j] == '{':
                    depth += 1
                elif s[j] == '}':
                    if depth == 0:
                        break
                    depth -= 1
                j += 1
            inner = s[pos:j]
            inner = fix_font_commands(inner)
            out.append(f'\\{new_cmd}{{{inner}}}')
            i = j
    return ''.join(out)

def collapse_whitespace(inner):
    inner = re.sub(r'\s*\n\s*', ' ', inner)
    inner = re.sub(r'\s+', ' ', inner).strip()
    return inner

def clean_math_blocks(text):
    """Procesa $$...$$, \\[...\\] y $...$ (matematica), dejando todo lo demas intacto.
    Evita procesar contenido dentro de bloques de codigo ```...``` para no romper
    ejemplos de codigo que las IAs a veces incluyen junto con la matematica."""

    # 1) Proteger bloques de codigo ```...``` guardandolos aparte temporalmente.
    code_blocks = []
    def stash_code(m):
        code_blocks.append(m.group(0))
        return f'\x00CODEBLOCK{len(code_blocks) - 1}\x00'
    text = re.sub(r'```.*?```', stash_code, text, flags=re.DOTALL)

    # 2) $$...$$ (bloques)
    def collapse_dd(m):
        inner = collapse_whitespace(m.group(1))
        inner = fix_font_commands(inner)
        return f'$${inner}$$'
    text = re.sub(r'\$\$(.*?)\$\$', collapse_dd, text, flags=re.DOTALL)

    # 3) \[ ... \] (bloques, estilo alternativo)
    def collapse_br(m):
        inner = collapse_whitespace(m.group(1))
        inner = fix_font_commands(inner)
        return f'\\[{inner}\\]'
    text = re.sub(r'\\\[(.*?)\\\]', collapse_br, text, flags=re.DOTALL)

    # 4) $...$ inline (una sola linea, sin tocar $$ ya procesados)
    def fix_inline(m):
        inner = fix_font_commands(m.group(1))
        return f'${inner}$'
    text = re.sub(r'(?<!\$)\$(?!\$)([^\n\$]+?)(?<!\$)\$(?!\$)', fix_inline, text)

    # 5) Restaurar bloques de codigo
    def restore_code(m):
        idx = int(m.group(1))
        return code_blocks[idx]
    text = re.sub(r'\x00CODEBLOCK(\d+)\x00', restore_code, text)

    return text

def count_math_blocks(text):
    dd = len(re.findall(r'\$\$.*?\$\$', text, flags=re.DOTALL))
    br = len(re.findall(r'\\\[.*?\\\]', text, flags=re.DOTALL))
    inline = len(re.findall(r'(?<!\$)\$(?!\$)[^\n\$]+?(?<!\$)\$(?!\$)', text))
    return dd, br, inline

def build_report(raw, cleaned, docx_path, result, path_label, docx_label):
    """Corre todos los chequeos de verificacion y arma el reporte final en
    texto (resumen ejecutivo arriba + detalle abajo). La usan tanto main()
    (linea de comandos) como app.py (interfaz web), para no duplicar logica.

    raw/cleaned : texto original y texto ya limpiado.
    docx_path   : ruta real en disco del .docx generado (se necesita para
                  poder leerlo de vuelta con Pandoc y verificar el resultado).
    result      : subprocess.CompletedProcess de la conversion md -> docx.
    path_label / docx_label : nombres para mostrar en el encabezado del
                  reporte (pueden ser rutas reales o nombres logicos, como
                  en app.py donde se trabaja con archivos temporales).

    Devuelve (texto_del_reporte, todo_ok: bool).
    """
    dd0, br0, in0 = count_math_blocks(raw)
    dd1, br1, in1 = count_math_blocks(cleaned)

    warnings = [l for l in result.stderr.splitlines() if l.startswith("[WARNING] Could not convert TeX math")]

    # Vamos a construir el reporte en dos pasadas: primero corremos todos los
    # chequeos y guardamos su resultado (ok/alerta + detalle); al final armamos
    # un resumen ejecutivo arriba y el detalle completo abajo.
    checks = []  # cada item: {"titulo", "ok": bool, "resumen": str, "detalle": [lineas]}

    # --- Chequeo 1: integridad de bloques de matematica ---
    integridad_ok = (dd0 == dd1 and br0 == br1 and in0 == in1)
    detalle1 = [
        f"Bloques $$...$$      -> originales: {dd0}   despues de limpiar: {dd1}",
        f"Bloques \\[...\\]      -> originales: {br0}   despues de limpiar: {br1}",
        f"Formulas $...$ inline -> originales: {in0}   despues de limpiar: {in1}",
    ]
    if not integridad_ok:
        detalle1.append("")
        detalle1.append("La cantidad de bloques de matematica cambio durante la limpieza.")
        detalle1.append("Esto NO deberia pasar. Revisa el archivo _limpio.md contra el original a mano.")
    checks.append({
        "titulo": "Integridad de bloques matematicos",
        "ok": integridad_ok,
        "resumen": "El numero de formulas es identico antes y despues de limpiar."
                   if integridad_ok else "Cambio la cantidad de formulas durante la limpieza.",
        "detalle": detalle1,
    })

    # --- Chequeo 2: formulas que Pandoc no logro convertir ---
    conversion_ok = (len(warnings) == 0)
    detalle2 = [f"Formulas con warning de conversion: {len(warnings)}"]
    if warnings:
        detalle2.append("Quedaron como texto/codigo LaTeX crudo en el .docx final. Revisalas manualmente:")
        for w in warnings:
            detalle2.append(f"  - {w}")
    else:
        detalle2.append("Ninguna formula genero warning de conversion.")
    detalle2.append("")
    detalle2.append("Nota: 'sin warnings' significa que Pandoc logro parsear la sintaxis TeX; no")
    detalle2.append("garantiza que el resultado visual sea identico al original. Para trabajos")
    detalle2.append("formales, se recomienda revisar el .docx contra el .md formula por formula.")
    checks.append({
        "titulo": "Conversion de formulas (Pandoc)",
        "ok": conversion_ok,
        "resumen": "Pandoc convirtio todas las formulas sin problemas."
                   if conversion_ok else f"{len(warnings)} formula(s) no se pudieron convertir.",
        "detalle": detalle2,
    })

    # --- Verificacion ida y vuelta: docx -> md, comparado contra el .md limpio ---

    def strip_math(t):
        # $$...$$ y \[...\] son bloque (pueden cruzar lineas); $...$ y \(...\)
        # son inline (una sola linea). Si falta alguno de estos 4, la
        # comparacion queda asimetrica: el lado "cleaned" conserva el LaTeX
        # crudo pero el lado "roundtrip" ya lo reconvirtio a otra notacion,
        # y aparecen diffs falsos linea por linea.
        t = re.sub(r'\$\$.*?\$\$', ' ', t, flags=re.DOTALL)
        t = re.sub(r'\\\[.*?\\\]', ' ', t, flags=re.DOTALL)
        t = re.sub(r'(?<!\$)\$(?!\$)[^\n\$]+?(?<!\$)\$(?!\$)', ' ', t)
        t = re.sub(r'\\\(.*?\\\)', ' ', t)
        return t

    def extract_numbers(t):
        # Pandoc a veces separa el signo menos del numero con un espacio al
        # reformatear ("-3.65" -> "- 3.65"). Lo normalizamos antes de extraer.
        t = re.sub(r'-\s+(?=\d)', '-', t)
        # Solo cuenta como "numero" un entero o un decimal con cifras despues
        # del punto (ej. 3, -5.08, 1.34e4). Un "3." suelto (el numero de un
        # encabezado tipo "# 3. Potencial...") NO cuenta: no es un dato de
        # una formula, es numeracion de seccion, y Pandoc a veces le quita
        # el punto al reescribirlo, generando una falsa alerta.
        nums = re.findall(r'-?\d+\.\d+|-?\d+(?!\.\D|\.$)', t)
        return sorted(nums)

    def unescape_backslashes(t):
        # Pandoc escapa "\" como "\\" al reescribir texto normal a Markdown
        # (convencion de escape, no perdida de datos). Lo revertimos.
        return t.replace('\\\\', '\\')

    # encoding="utf-8" es clave aqui: en Windows, subprocess con text=True sin
    # especificar encoding usa la codificacion por defecto del sistema (a
    # menudo cp1252 en Windows en espanol), no UTF-8. Pandoc SI escribe UTF-8
    # en su salida, asi que sin esto, cualquier tilde o enie se lee mal
    # (aparece como "Ã³" en vez de "ó") y genera diffs falsos en el reporte.
    # No afecta al .docx en si (Pandoc lo escribe directo a disco), solo a
    # esta verificacion.
    roundtrip = subprocess.run(
        ["pandoc", docx_path, "-t", "markdown", "--wrap=none"],
        capture_output=True, text=True, encoding="utf-8"
    )

    if roundtrip.returncode != 0:
        checks.append({
            "titulo": "Texto normal (fuera de formulas)",
            "ok": False,
            "resumen": "No se pudo verificar automaticamente (Pandoc fallo al leer el .docx).",
            "detalle": ["No se pudo hacer la verificacion automatica (Pandoc fallo al leer el .docx)."],
        })
        checks.append({
            "titulo": "Numeros dentro de las formulas",
            "ok": False,
            "resumen": "No se pudo verificar (depende de la lectura del .docx, que fallo).",
            "detalle": ["No se pudo hacer la verificacion automatica (Pandoc fallo al leer el .docx)."],
        })
    else:
        from collections import Counter

        def extract_words(t):
            # Comparamos "bolsas de palabras" en vez de linea por linea.
            # Esto es a proposito: Pandoc reescribe el formato al convertir
            # docx -> markdown (las tablas markdown "|...|" se vuelven tablas
            # de texto con guiones y espacios, las lineas "---" se alargan,
            # los parrafos se re-parten distinto) sin que eso signifique que
            # se perdio contenido. Comparar solo las palabras reales ignora
            # ese formato y detecta unicamente perdida o alteracion de texto.
            t = strip_math(t)
            t = unescape_backslashes(t)
            t = re.sub(r'-{3,}', ' ', t)     # lineas horizontales / bordes de tabla
            t = re.sub(r'\|', ' ', t)        # separadores de tabla markdown
            t = re.sub(r'^```.*?```$', ' ', t, flags=re.DOTALL | re.MULTILINE)
            return re.findall(r"[A-Za-zÀ-ÿ0-9']+", t.lower())

        # --- Chequeo 3: texto normal (fuera de formulas), por palabras ---
        orig_words = Counter(extract_words(cleaned))
        back_words = Counter(extract_words(roundtrip.stdout))
        faltan = orig_words - back_words   # palabras que estaban y ya no estan
        sobran = back_words - orig_words   # palabras que aparecieron de mas
        prose_ok = (not faltan and not sobran)

        def formatear_conteo(counter, limite=25):
            items = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))
            piezas = [f"{w} (x{n})" if n > 1 else w for w, n in items[:limite]]
            texto = ", ".join(piezas)
            if len(items) > limite:
                texto += f", ... y {len(items) - limite} palabra(s) mas"
            return texto

        if prose_ok:
            detalle3 = ["El texto normal (fuera de formulas) coincide palabra por palabra",
                        "entre el .md y el .docx (se ignoran diferencias de formato: tablas,",
                        "lineas horizontales, saltos de linea)."]
        else:
            detalle3 = []
            if faltan:
                total_faltan = sum(faltan.values())
                detalle3.append(f"FALTAN en el .docx {total_faltan} palabra(s) que si estaban en el original:")
                detalle3.append("  " + formatear_conteo(faltan))
                detalle3.append("")
            if sobran:
                total_sobran = sum(sobran.values())
                detalle3.append(f"SOBRAN en el .docx {total_sobran} palabra(s) que NO estaban en el original:")
                detalle3.append("  " + formatear_conteo(sobran))
                detalle3.append("")
            detalle3.append("Revisa el .docx buscando estas palabras/zonas para confirmar si es")
            detalle3.append("perdida real de contenido o solo una reformulacion de Pandoc.")
        checks.append({
            "titulo": "Texto normal (fuera de formulas)",
            "ok": prose_ok,
            "resumen": "El texto fuera de las formulas coincide palabra por palabra."
                       if prose_ok else
                       f"Faltan {sum(faltan.values())} palabra(s) y sobran {sum(sobran.values())} en el .docx.",
            "detalle": detalle3,
        })

        # --- Chequeo 4: numeros dentro de las formulas ---
        orig_nums = extract_numbers(cleaned)
        back_nums = extract_numbers(roundtrip.stdout)
        nums_ok = (orig_nums == back_nums)
        if nums_ok:
            detalle4 = [
                f"Los {len(orig_nums)} numeros encontrados en las formulas coinciden exactamente",
                "(mismos valores, mismo conteo) entre el .md original y el .docx generado.",
            ]
        else:
            faltantes, extras = [], list(back_nums)
            for n in orig_nums:
                if n in extras:
                    extras.remove(n)
                else:
                    faltantes.append(n)
            detalle4 = [
                "Los numeros no coinciden entre el .md y el .docx.",
                f"Numeros del original que NO aparecen en el docx: {faltantes if faltantes else 'ninguno'}",
                f"Numeros en el docx que NO estaban en el original: {extras if extras else 'ninguno'}",
                "",
                "Esto NO necesariamente es un error del script (Pandoc a veces reordena o",
                "reformatea exponentes), pero se recomienda revisar manualmente estas formulas.",
            ]
        checks.append({
            "titulo": "Numeros dentro de las formulas",
            "ok": nums_ok,
            "resumen": "Los numeros de las formulas coinciden entre el .md y el .docx."
                       if nums_ok else "Hay numeros que no coinciden entre el .md y el .docx.",
            "detalle": detalle4,
        })

    # --- Ensamblado del reporte final: resumen ejecutivo arriba, detalle abajo ---
    ANCHO = 70
    todo_ok = all(c["ok"] for c in checks) and result.returncode == 0
    lines = []
    lines.append("=" * ANCHO)
    lines.append("REPORTE DE CONVERSION MD -> DOCX")
    lines.append("=" * ANCHO)
    lines.append(f"Archivo original : {path_label}")
    lines.append(f"Documento final  : {docx_label}")
    lines.append("")

    if result.returncode != 0:
        lines.append("RESULTADO GENERAL: [ERROR] Pandoc no pudo generar el .docx.")
    elif todo_ok:
        lines.append("RESULTADO GENERAL: [OK] Todo salio bien, sin observaciones.")
    else:
        pendientes = sum(1 for c in checks if not c["ok"])
        lines.append(f"RESULTADO GENERAL: [REVISAR] El documento se genero, pero hay "
                      f"{pendientes} punto(s) con observaciones (ver detalle abajo).")
    lines.append("")

    lines.append("-" * ANCHO)
    lines.append("RESUMEN")
    lines.append("-" * ANCHO)
    for c in checks:
        marca = "[OK]     " if c["ok"] else "[REVISAR]"
        lines.append(f"{marca} {c['titulo']}")
        lines.append(f"          {c['resumen']}")
    lines.append("")

    lines.append("-" * ANCHO)
    lines.append("DETALLE")
    lines.append("-" * ANCHO)
    for n, c in enumerate(checks, start=1):
        marca = "OK" if c["ok"] else "REVISAR"
        lines.append(f"\n{n}. {c['titulo']}  [{marca}]")
        lines.append("-" * ANCHO)
        lines.extend(c["detalle"])

    lines.append("")
    lines.append("=" * ANCHO)
    lines.append("Fin del reporte.")
    lines.append("=" * ANCHO)

    return "\n".join(lines), todo_ok


def main():
    if len(sys.argv) < 2:
        print("Uso: arrastra un archivo .md sobre este script, o corre: python clean_and_convert.py archivo.md")
        input("Presiona Enter para salir...")
        return
    path = sys.argv[1]
    with open(path, encoding="utf-8") as f:
        raw = f.read()

    cleaned = clean_math_blocks(raw)

    base, _ = os.path.splitext(path)
    clean_path = base + "_limpio.md"
    docx_path = base + ".docx"
    report_path = base + "_reporte.txt"

    with open(clean_path, "w", encoding="utf-8") as f:
        f.write(cleaned)

    # -raw_tex: evita que Pandoc interprete comandos \algo desconocidos como LaTeX
    # crudo y los descarte en silencio si no los reconoce. Con esto, si algo no se
    # puede convertir, se mantiene VISIBLE como texto en vez de desaparecer.
    result = subprocess.run(
        ["pandoc", "-f", "markdown+tex_math_single_backslash-raw_tex", clean_path, "-o", docx_path],
        capture_output=True, text=True, encoding="utf-8"
    )

    report, _ = build_report(raw, cleaned, docx_path, result, path_label=path, docx_label=docx_path)

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)

    print(report)
    if result.returncode != 0:
        print("\n[ERROR] Pandoc fallo. Revisa el mensaje completo abajo:")
        print(result.stderr)
    else:
        print(f"\nListo: {docx_path}")
        print(f"Reporte guardado en: {report_path}")

if __name__ == "__main__":
    main()