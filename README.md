# md2docx (v2.1)

Herramienta en Python con **interfaz gráfica nativa** y **modo línea de comandos (CLI)** que convierte archivos Markdown (`.md`) a documentos de Word (`.docx`) utilizando [pandoc](https://pandoc.org). Permite aplicar una plantilla `.docx` para controlar estilos, márgenes, encabezados y pies de página sin dependencias externas de Python.

---

## Novedades de la v2.1

- **Interfaz gráfica (GUI) por defecto:** Al abrir el programa (por terminal o doble clic) se muestra una ventana intuitiva para seleccionar archivos y opciones sin escribir comandos.
- **Ruta de guardado automática:** El archivo `.docx` se genera automáticamente en la **misma carpeta** y con el **mismo nombre** que el `.md` de origen, mostrando una vista previa de la ruta en tiempo real.
- **Gestión inteligente de conflictos y bloqueos:** Permite elegir entre preguntar antes de sobrescribir, sobrescribir directamente o crear versiones incrementales (`_v2.docx`, `_v3.docx`...). También detecta si el archivo está abierto en Word para evitar errores de permisos.
- **Accesos rápidos post-conversión:** Botones para abrir directamente el `.docx` generado en Word o abrir la carpeta contenedora.
- **Blindaje y compatibilidad:** Compatible desde **Python 3.8** en adelante (incluyendo Python 3.15). Incluye captura de errores fatales mediante cuadros de diálogo nativos de Windows para evitar cierres silenciosos al ejecutar con doble clic.

---

## Qué hace y qué no hace

### Qué hace
- Convierte `.md` y `.markdown` a `.docx` mediante `pandoc`.
- Aplica los estilos, márgenes, encabezados y pies de página de una plantilla `.docx` (`--reference-doc`).
- Crea la plantilla base de pandoc con un clic desde la interfaz (o mediante `--init-template`).
- Resuelve rutas relativas de imágenes tomando como base la carpeta del archivo `.md` (`--resource-path`).
- Inserta tabla de contenidos (`--toc`) y numeración de capítulos/secciones (`--number-sections`) de forma opcional.

### Qué NO hace
- No crea portadas corporativas dinámicas, logos ni metadatos complejos: pandoc hereda únicamente **estilos, márgenes, encabezados y pies** de la plantilla.
- No postprocesa el `.docx` tras generarlo (por ejemplo, ajuste automático de anchos de columna en tablas).
- No admite plantillas `.dotx`: la plantilla debe guardarse siempre como `.docx`.
- No realiza la conversión inversa (de `.docx` a `.md`).

---

## Requisitos

| Requisito | Detalle |
|---|---|
| **Python** | 3.8 o superior (`tkinter` incluido en las instalaciones estándar de Windows/macOS) |
| **pandoc** | Instalado y accesible en el `PATH` del sistema (`pandoc --version` debe responder) |
| **Dependencias Python** | **Ninguna externa.** Solo usa la biblioteca estándar (`tkinter`, `argparse`, `subprocess`, `shutil`, `pathlib`, `os`, `sys`, `traceback`, `ctypes`) |
| **Microsoft Word** | Opcional, solo necesario para editar la plantilla de estilos y revisar los `.docx` |

---

## Estructura de carpetas

```text
md2docx/
├── md2docx.py                  # Código principal (GUI + CLI)
├── README.md                   # Esta documentación
└── plantillas/
    └── plantilla_base.docx     # Se genera desde la GUI o con --init-template
```

> Los archivos `.md` pueden estar ubicados en cualquier carpeta o unidad (local o sincronizada como Google Drive). El `.docx` resultante se guardará siempre junto al `.md` original (salvo que uses `-o` en modo CLI).

---

## Puesta en marcha

1. **Instala pandoc** desde [pandoc.org/installing.html](https://pandoc.org/installing.html) y verifica en una terminal que `pandoc --version` funciona.
2. **Inicia la aplicación** haciendo doble clic sobre `md2docx.py` (o ejecutando `python md2docx.py` en la terminal).
3. **Crea tu plantilla base (una sola vez):**
   - Pulsa el botón **"Crear / regenerar plantilla_base.docx"** en la interfaz (o ejecuta `python md2docx.py --init-template`).
   - Abre `plantillas/plantilla_base.docx` en Word.
   - Edita los estilos usando **Modificar estilo** (nunca formato directo sobre el texto):
     - `Title`, `Heading 1`, `Heading 2`, `Heading 3`
     - `Body Text`, `First Paragraph`, `Compact`
     - `Table`, `Source Code`
     - Márgenes, encabezados y pies de página.
   - Guarda el archivo como `.docx` (nunca `.dotx`).

---

## Guía de uso

### 1. Modo interfaz gráfica (recomendado)

Ejecuta `python md2docx.py` o haz doble clic sobre `md2docx.py`.

1. **Archivo de origen y destino:** Pulsa **Examinar...** y selecciona tu archivo `.md`. En el recuadro *"Se guardará en:"* verás la ruta exacta del `.docx` que se creará en esa misma carpeta, junto con un aviso si ya existe un archivo previo.
2. **Plantilla de estilos:** Deja marcada la casilla si quieres aplicar `plantilla_base.docx` (o pulsa **Elegir...** para seleccionar otra plantilla `.docx` distinta). Desmárcala si prefieres el estilo neutro por defecto de pandoc.
3. **Opciones de documento:**
   - Marca si deseas incluir **tabla de contenidos / índice (TOC)** o **numerar capítulos y secciones**.
   - Selecciona el comportamiento si el `.docx` ya existe: *Preguntar antes*, *Crear versión (`_v2`, `_v3`...)* o *Sobrescribir directamente*.
4. **Convertir:** Pulsa **Convertir a Word (.docx)**. Al terminar, puedes usar los botones **Abrir Word** o **Abrir carpeta** para revisar el resultado de inmediato.

> **Nota para usuarios de Windows:** Si quieres abrir la interfaz gráfica sin que aparezca ninguna consola negra detrás al hacer doble clic, puedes crear un archivo `Iniciar_md2docx.bat` en la misma carpeta con el siguiente contenido:
> ```bat
> @echo off
> start "" pythonw "%~dp0md2docx.py"
> ```

---

### 2. Modo línea de comandos (CLI)

Si pasas un archivo `.md` como argumento o usas `--cli`, el programa funciona en consola sin abrir ventanas gráficas:

```bash
# Conversión rápida usando plantilla_base.docx si existe
python md2docx.py informe.md

# Especificar plantilla personalizada, índice y numeración
python md2docx.py informe.md -t otra_plantilla.docx --toc --number-sections

# Guardar en una ruta distinta a la carpeta del .md
python md2docx.py informe.md -o C:\Salidas\informe_final.docx

# Convertir sin aplicar ninguna plantilla
python md2docx.py informe.md --no-template

# Generar la plantilla base por consola
python md2docx.py --init-template

# Forzar el modo interactivo clásico por preguntas en terminal
python md2docx.py --cli
```

#### Referencia de argumentos CLI

| Opción | Descripción |
|---|---|
| *(sin argumentos)* | Abre la interfaz gráfica (`tkinter`) |
| `input` | Ruta del archivo `.md` de entrada |
| `-t`, `--template` | Ruta de la plantilla `.docx` a usar como `--reference-doc` |
| `--no-template` | Ignora la plantilla por defecto y usa los estilos base de pandoc |
| `-o`, `--output` | Ruta personalizada de salida (por defecto: junto al `.md`) |
| `--toc` | Inserta una tabla de contenidos (índice) al inicio del documento |
| `--number-sections` | Numera automáticamente los encabezados y secciones |
| `--lang` | Código de idioma del documento (por defecto `es-ES`) |
| `--init-template` | Crea `plantillas/plantilla_base.docx` y termina |
| `--cli` | Fuerza el modo interactivo en consola en lugar de abrir la ventana |

---

## Flujo de trabajo recomendado

1. Redacta el documento en Markdown (`.md`), incluyendo opcionalmente un bloque de metadatos YAML al inicio:
   ```yaml
   ---
   title: Memoria técnica del proyecto
   author: Nombre Apellidos
   date: Noviembre 2026
   ---
   ```
2. Convierte el archivo con **md2docx**.
3. Si activaste la opción de **Índice (TOC)**, al abrir el `.docx` en Word haz **clic derecho sobre el índice > Actualizar campo** (o pulsa `F9`) para que Word calcule los números de página.
4. **Mantén el `.md` como fuente única de verdad:** si quieres cambiar tipografías, colores o márgenes, modifica los estilos en `plantilla_base.docx` y vuelve a convertir, en lugar de editar el `.docx` final a mano.

---

## Arquitectura interna (`md2docx.py`)

El script está diseñado en un único archivo autocontenido con separación clara entre el motor de conversión y las interfaces de usuario:

| Componente / Función | Responsabilidad |
|---|---|
| `check_pandoc()` | Verifica mediante `shutil.which` que `pandoc` esté disponible en el `PATH` |
| `create_default_template()` | Extrae `reference.docx` desde pandoc y lo guarda en `plantillas/plantilla_base.docx` |
| `resolve_output_path()` | Calcula la ruta destino junto al `.md`, gestiona colisiones (`_v2`, `_v3`...) y detecta bloqueos de archivo en Windows |
| `run_conversion()` | Valida entradas, construye el comando de `pandoc`, ejecuta el subproceso y captura avisos/errores |
| `launch_gui()` | Construye y gestiona la interfaz gráfica en `tkinter` / `ttk` |
| `show_fatal_error()` | Muestra cuadros de error nativos (`MessageBoxW` en Windows / `tkinter`) si ocurre un fallo crítico al arrancar sin consola |
| `main()` | Gestiona los argumentos con `argparse` y decide si lanzar la GUI, el modo CLI o `--init-template` |

---

## Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| `pandoc no está en el PATH` | Pandoc no está instalado o la sesión no se ha reiniciado tras instalarlo | Instala pandoc y cierra/abre la aplicación o terminal |
| `pandoc necesita la plantilla en formato .docx` | Guardaste la plantilla en Word como `.dotx` | Ábrela en Word y guárdala como **Documento de Word (`.docx`)** |
| Los cambios en la plantilla no se reflejan | Se aplicó formato directo al texto o se editó el estilo `Normal` | En la plantilla, edita los estilos específicos de pandoc (`Body Text`, `First Paragraph`, `Heading 1`...) |
| Se genera `archivo_v2.docx` automáticamente | El `.docx` original estaba abierto en Word o tienes marcada la opción de crear versión | Cierra el documento en Word antes de sobrescribirlo |
| Al hacer doble clic en `.py` se abre IDLE o el Bloc de notas | Windows tiene asociado `.py` a un editor en lugar del intérprete de Python | Haz clic derecho > *Abrir con* > *Python*, o usa un acceso directo `.bat` con `pythonw` |
| El índice (TOC) aparece vacío al abrir Word | Pandoc inserta el campo dinámico de Word pero no lo calcula | Haz clic derecho sobre el índice en Word > *Actualizar campo* (o pulsa `F9`) |