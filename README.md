# Conversor .txt / .md a .docx

Herramienta para limpiar y convertir archivos `.txt`/`.md` (con fórmulas en
LaTeX) a documentos de Word (`.docx`), con verificación automática de que no
se perdió contenido en la conversión. Incluye una interfaz web local.

## Descarga

Descargar la última versión desde la sección
**[Releases](../../releases)** de este repositorio (buscar el archivo
`.zip` adjunto a la versión más reciente).

## Requisitos

El programa requiere de dos programas instalados en la computadora:

1. **Python 3** (para correr `app.py`)
2. **Pandoc** (hace la conversión real de Markdown a `.docx`)

Ambos se instalan una sola vez. Después de eso, el programa se abre con
un doble clic.

### Instalar Python

1. Entrar a **https://www.python.org/downloads/** y descargar la versión
   más reciente para Windows.
2. Al ejecutar el instalador, marcar la casilla **"Add python.exe to PATH"**
   (aparece abajo del todo en la primera pantalla del instalador). Esto es
   importante: si no se marca, `iniciar.bat` no va a encontrar Python.
3. Hacer clic en "Install Now" y esperar a que termine.
4. Para confirmar que quedó bien instalado, abrir la terminal (`cmd`) y
   escribir:
   ```
   python --version
   ```
   Debería mostrar algo como `Python 3.12.x`.

### Instalar Pandoc

1. Entrar a **https://pandoc.org/installing.html** y descargar el instalador
   `.msi` para Windows.
2. Ejecutar el instalador con las opciones por defecto.
3. Para confirmar, abrir la terminal y escribir:
   ```
   pandoc --version
   ```
   Debería mostrar la versión instalada.

> Nota: hace falta reiniciar la terminal (o la sesión de Windows) después
> de instalar Python y Pandoc para que el `PATH` se actualice.

## Uso

1. Descomprimir el `.zip` descargado en cualquier carpeta.
2. Hacer doble clic en **`iniciar.bat`**.
3. Se va a abrir una ventana negra (no se debe cerrar mientras se use el
   programa) y el navegador con la interfaz web.
4. Arrastrar su archivo `.txt` o `.md` al recuadro de la página (o varios
   a la vez para convertir en lote).
5. Completar los datos del archivo (problema, IA, versión, asistente,
   origen) si corresponde.
6. Hacer clic en **Convertir** y descargar el `.docx` (o el `.zip` con
   todos, en modo lote). Debajo va a aparecer la fecha y hora de
   generación, y un reporte de verificación.

## Actualizar a una nueva versión

Cuando haya una versión nueva, simplemente necesita descargar el `.zip` más
reciente desde **[Releases](../../releases)** y reemplazar los archivos
de su carpeta actual. No hace falta reinstalar Python ni Pandoc: esos
quedan instalados en la computadora, no en la carpeta del programa.

## Archivos del proyecto

| Archivo                | Qué hace                                              |
|-------------------------|--------------------------------------------------------|
| `iniciar.bat`            | Doble clic para abrir el programa. Únicamente debe interactuar con este archivo.                    |
| `app.py`                  | Levanta la interfaz web local.                        |
| `clean_and_convert.py`  | Lógica de limpieza de fórmulas y conversión con Pandoc.|

## Historial de versiones

Ver la sección [Releases](../../releases) para el changelog de cada
versión publicada.
