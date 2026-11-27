# md2docx

Script en Python que convierte archivos Markdown (`.md`) a Word (`.docx`) usando [pandoc](https://pandoc.org), con plantilla opcional para controlar el aspecto del documento final.

## Estado

**v1. No se ha ejecutado contra una instalación real de pandoc en el momento de escribir esto.** Antes de usarlo en un entregable, haz una primera prueba con un `.md` corto y otro con una tabla ancha. Si algo falla, el error de pandoc se muestra tal cual en la terminal.

## Qué hace

- Pide el archivo `.md` (por argumento o de forma interactiva).
- Aplica una plantilla `.docx` si se desea (estilos, márgenes, encabezado y pie).
- Genera el `.docx` junto al `.md` de origen, o en la ruta que indiques.
- Genera la plantilla base de pandoc con un solo comando, para que la edites en Word.

## Qué NO hace (v1)

- No crea portadas corporativas, logos ni datos de proyecto: pandoc solo hereda **estilos, márgenes, encabezados y pies** de la plantilla.
- No postprocesa el `.docx` (anchos de tabla, campos, numeración avanzada). Está previsto para una v2 con `python-docx`.
- No acepta plantillas `.dotx`: hay que guardarlas como `.docx`.
- No convierte de `.docx` a `.md`.

## Requisitos

| Requisito | Detalle |
|---|---|
| Python | 3.10 o superior (usa la sintaxis `Path \| None`) |
| pandoc | Instalado y accesible en el PATH (`pandoc --version` debe funcionar) |
| Dependencias Python | Ninguna: solo biblioteca estándar (`argparse`, `subprocess`, `shutil`, `pathlib`, `sys`) |
| Word | Solo para editar la plantilla y abrir el resultado |

## Estructura de carpetas

```
md2docx/
├── md2docx.py
├── README.md
└── plantillas/
    └── plantilla_base.docx     (se crea con --init-template)
```

Los `.md` pueden estar en cualquier sitio. La salida se guarda junto al origen salvo que indiques `-o`.

## Puesta en marcha (una sola vez)

1. Instala pandoc y comprueba: `pandoc --version`.
2. Crea la plantilla base:
   ```
   python md2docx.py --init-template
   ```
3. Abre `plantillas/plantilla_base.docx` en Word y modifica los estilos con **Modificar estilo** (no con formato directo):
   - `Title`, `Heading 1`, `Heading 2`, `Heading 3`
   - `Body Text`, `First Paragraph`, `Compact`
   - `Table`, `Source Code`
   - Márgenes, encabezado y pie de página
4. Guarda como `.docx` (nunca `.dotx`).

## Uso

### Modo interactivo

```
python md2docx.py
```

Pregunta, por este orden: archivo `.md`, plantilla (Enter = la de por defecto si existe, si no, ninguna), índice y numeración de secciones. Limpia comillas y espacios, así que puedes arrastrar archivos a la terminal.

### Modo línea de comandos

```
python md2docx.py informe.md
python md2docx.py informe.md -t otra_plantilla.docx
python md2docx.py informe.md --toc --number-sections
python md2docx.py informe.md -o C:\Salidas\informe_final.docx
python md2docx.py informe.md --no-template
python md2docx.py --init-template
```

### Opciones

| Opción | Efecto |
|---|---|
| `input` | Archivo `.md` de entrada (opcional: si falta, modo interactivo) |
| `-t`, `--template` | Plantilla `.docx` a usar como `--reference-doc` |
| `--no-template` | No usar plantilla, ni siquiera la de por defecto |
| `-o`, `--output` | Ruta del `.docx` de salida |
| `--toc` | Inserta índice de contenidos |
| `--number-sections` | Numera los capítulos y apartados |
| `--lang` | Idioma del documento (por defecto `es-ES`) |
| `--init-template` | Crea `plantillas/plantilla_base.docx` y termina |

### Qué plantilla se usa

1. La indicada con `-t`, si existe.
2. Si no, `plantillas/plantilla_base.docx` junto al script, si existe.
3. Si no, ninguna (estilos por defecto de pandoc).
4. `--no-template` salta los pasos 1 y 2.

## Modus operandi recomendado

1. Escribe el contenido en `.md` con metadatos YAML al inicio:
   ```yaml
   ---
   title: Memoria técnica de climatización
   author: Nombre Apellidos
   date: 26 de noviembre de 2026
   ---
   ```
2. Ejecuta el script.
3. Abre el `.docx`. Si usaste `--toc`, actualiza el índice (clic derecho > *Actualizar campo*, o F9).
4. **Si necesitas cambiar el formato, edita la plantilla, no el resultado.** El `.md` es la fuente de verdad: si retocas el `.docx` a mano y vuelves a convertir, pierdes esos cambios.

## Cómo está hecho

Un único archivo, `md2docx.py`, sin dependencias externas de Python. Llama a pandoc directamente con `subprocess`, sin `pypandoc`, para tener una dependencia menos y ver los errores reales de pandoc.

### Funciones

| Función | Responsabilidad |
|---|---|
| `clean_path(raw)` | Quita comillas y espacios de una ruta, y expande `~` |
| `check_pandoc()` | Comprueba con `shutil.which` que pandoc está en el PATH; si no, termina con error |
| `init_template()` | Ejecuta `pandoc --print-default-data-file reference.docx` y guarda `plantillas/plantilla_base.docx`. Pregunta antes de sobrescribir. Si la versión de pandoc devuelve el contenido por stdout en vez de escribirlo con `-o`, lo guarda desde stdout |
| `ask_interactive()` | Pide archivo, plantilla, índice y numeración cuando no hay argumentos |
| `resolve_output(md, output)` | Decide la ruta de salida; si existe, pregunta entre sobrescribir, crear versión nueva (`_v2`, `_v3`...) o cancelar; detecta archivos bloqueados |
| `convert(args)` | Valida entradas, monta el comando de pandoc, lo ejecuta y muestra avisos y errores |
| `main()` | Define los argumentos con `argparse` y decide entre `--init-template`, modo interactivo o conversión |

### Comando de pandoc que se ejecuta

```
pandoc entrada.md -f markdown -t docx -o salida.docx \
  --resource-path <carpeta del .md> \
  -M lang=es-ES \
  [--reference-doc plantilla.docx] [--toc] [--number-sections]
```

- `--resource-path` hace que las imágenes con rutas relativas se resuelvan desde la carpeta del `.md`.
- `-M lang=es-ES` fija el idioma para ortografía y comillas.

### Validaciones previas

- pandoc disponible en el PATH.
- El `.md` existe (si la extensión no es `.md` o `.markdown`, solo avisa).
- La plantilla existe y es `.docx`; se rechaza `.dotx`.
- Si el `.docx` de salida existe, se pregunta antes de tocarlo.

### Archivo abierto en Word (Windows)

Word bloquea el `.docx` que tiene abierto. Si el destino existe, el script intenta abrirlo en modo añadir; si salta `PermissionError`, guarda con sufijo de versión y lo avisa, en vez de fallar.

### Salida de errores

Los avisos de pandoc (`stderr`) se muestran siempre, aunque la conversión termine bien. Si pandoc devuelve un código distinto de 0, el script termina con ese código en el mensaje.

## Limitaciones conocidas

- Pandoc usa sus propios nombres de estilo (`Body Text`, `First Paragraph`, `Compact`...), no `Normal`. Modificar solo `Normal` en la plantilla no basta.
- Las pipe tables calculan los anchos de columna según la longitud de las líneas separadoras (`---`). Las tablas anchas de cálculo pueden salir mal repartidas; la solución v1 es ajustar los guiones del `.md` a mano.
- El índice (`--toc`) es un campo de Word y no se rellena hasta que lo actualizas al abrir el documento.
- Si una imagen no se encuentra, pandoc avisa pero genera el documento igualmente.

## Solución de problemas

| Síntoma | Causa probable | Qué hacer |
|---|---|---|
| `pandoc no está en el PATH` | No instalado o terminal abierta antes de instalar | Instalar y abrir una terminal nueva |
| `la plantilla debe ser .docx` | Plantilla guardada como `.dotx` | Guardar como `.docx` en Word |
| El formato no cambia | Formato directo en vez de estilos en la plantilla | Modificar los estilos de pandoc, no el texto |
| Sale `..._v2.docx` | El original estaba abierto en Word | Cerrarlo y reconvertir, o usar el `_v2` |
| `--init-template` falla | Diferencia entre versiones de pandoc | Probar a mano `pandoc -o plantilla_base.docx --print-default-data-file reference.docx` y revisar el error |
| Caracteres raros en la terminal | Codificación de la consola de Windows | Ejecutar `chcp 65001` antes del script |

## Hoja de ruta (v2)

- Postprocesado con `python-docx`: portada corporativa, anchos de tabla, pies con numeración.
- Selector gráfico de archivos con `tkinter` como alternativa al modo interactivo de terminal.
- Procesado por lotes de una carpeta de `.md`.
