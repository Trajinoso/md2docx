#!/usr/bin/env python3
"""
md2docx.py (v2.1) - Convierte Markdown a Word (.docx) usando pandoc.
Compatible con Python 3.8+ y ejecución por doble clic en Windows.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Callable, Optional

SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = SCRIPT_DIR / "plantillas"
DEFAULT_TEMPLATE = TEMPLATE_DIR / "plantilla_base.docx"


class ConversionError(Exception):
    """Error controlado durante las validaciones o la ejecución de pandoc."""


def show_fatal_error(title: str, message: str) -> None:
    """Muestra un error visible incluso si se ejecuta con doble clic sin consola."""
    print(f"\n[{title}]\n{message}", file=sys.stderr)
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)
            return
        except Exception:
            pass
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:
        input("\nPulsa Enter para salir...")


def clean_path(raw: str) -> Path:
    """Limpia comillas y espacios (rutas arrastradas a la terminal)."""
    return Path(raw.strip().strip('"').strip("'")).expanduser()


def check_pandoc() -> None:
    """Comprueba que pandoc esté disponible en el PATH."""
    if shutil.which("pandoc") is None:
        raise ConversionError(
            "pandoc no está en el PATH.\n"
            "Instálalo desde https://pandoc.org/installing.html y reinicia la aplicación."
        )


def create_default_template(overwrite: bool = False) -> Path:
    """Genera la plantilla base de pandoc y devuelve su ruta."""
    check_pandoc()
    TEMPLATE_DIR.mkdir(exist_ok=True)
    if DEFAULT_TEMPLATE.exists() and not overwrite:
        raise FileExistsError(f"La plantilla {DEFAULT_TEMPLATE.name} ya existe.")

    result = subprocess.run(
        ["pandoc", "-o", str(DEFAULT_TEMPLATE), "--print-default-data-file", "reference.docx"],
        capture_output=True,
    )
    if result.returncode != 0:
        raise ConversionError(
            f"Error generando plantilla:\n{result.stderr.decode(errors='replace')}"
        )
    if not DEFAULT_TEMPLATE.exists() or DEFAULT_TEMPLATE.stat().st_size == 0:
        DEFAULT_TEMPLATE.write_bytes(result.stdout)

    return DEFAULT_TEMPLATE


def next_available_version(target: Path) -> Path:
    """Devuelve la siguiente ruta libre con sufijo _v2, _v3..."""
    n = 2
    while target.with_name(f"{target.stem}_v{n}.docx").exists():
        n += 1
    return target.with_name(f"{target.stem}_v{n}.docx")


def is_file_locked(path: Path) -> bool:
    """Comprueba si un archivo existente está bloqueado (ej. abierto en Word en Windows)."""
    if not path.exists():
        return False
    try:
        with open(path, "ab"):
            return False
    except PermissionError:
        return True


def resolve_output_path(
    md: Path,
    output: Optional[Path] = None,
    conflict_mode: str = "ask",
    ask_callback: Optional[Callable[[Path], str]] = None,
) -> tuple[Path, list[str]]:
    """
    Determina la ruta del .docx de salida (por defecto junto al .md).
    conflict_mode: 'ask' | 'overwrite' | 'version'
    """
    notes: list[str] = []
    out = output if output else md.with_suffix(".docx")

    if out.exists():
        if conflict_mode == "version":
            out = next_available_version(out)
            notes.append(f"El archivo ya existía; se guardará como nueva versión: {out.name}")
        elif conflict_mode == "ask":
            if ask_callback is not None:
                choice = ask_callback(out)
            else:
                resp = input(
                    f"{out.name} ya existe. [s]obrescribir / [v]ersión nueva / [c]ancelar: "
                ).strip().lower()
                choice = {"s": "overwrite", "v": "version"}.get(resp, "cancel")

            if choice == "version":
                out = next_available_version(out)
            elif choice != "overwrite":
                raise ConversionError("Operación cancelada por el usuario.")

    if is_file_locked(out):
        alt = next_available_version(out)
        notes.append(f"AVISO: {out.name} está abierto o bloqueado. Se guardará como {alt.name}")
        out = alt

    return out, notes


def run_conversion(
    md: Path,
    template: Optional[Path] = None,
    no_template: bool = False,
    output: Optional[Path] = None,
    toc: bool = False,
    number_sections: bool = False,
    lang: str = "es-ES",
    conflict_mode: str = "ask",
    ask_callback: Optional[Callable[[Path], str]] = None,
) -> tuple[Path, str]:
    """Ejecuta la conversión completa y devuelve (ruta_salida, registro_mensajes)."""
    check_pandoc()
    logs: list[str] = []

    if not md or not md.is_file():
        raise ConversionError(f"No existe el archivo de entrada:\n{md}")

    if md.suffix.lower() not in {".md", ".markdown"}:
        logs.append(f"AVISO: {md.name} no tiene extensión .md; se intenta convertir igualmente.")

    resolved_template: Optional[Path] = None
    if not no_template:
        if template:
            resolved_template = template
        elif DEFAULT_TEMPLATE.exists():
            resolved_template = DEFAULT_TEMPLATE

    if resolved_template:
        if not resolved_template.is_file():
            raise ConversionError(f"No existe la plantilla seleccionada:\n{resolved_template}")
        if resolved_template.suffix.lower() != ".docx":
            raise ConversionError(
                "pandoc necesita la plantilla en formato .docx (no .dotx). Guárdala como .docx."
            )

    out, resolution_notes = resolve_output_path(
        md=md,
        output=output,
        conflict_mode=conflict_mode,
        ask_callback=ask_callback,
    )
    logs.extend(resolution_notes)

    cmd = [
        "pandoc",
        str(md),
        "-f",
        "markdown",
        "-t",
        "docx",
        "-o",
        str(out),
        "--resource-path",
        str(md.parent),
        "-M",
        f"lang={lang}",
    ]
    if resolved_template:
        cmd += ["--reference-doc", str(resolved_template)]
    if toc:
        cmd.append("--toc")
    if number_sections:
        cmd.append("--number-sections")

    logs.append("Comando: " + " ".join(f'"{c}"' if " " in c else c for c in cmd))

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    if result.stderr.strip():
        logs.append("--- Avisos de pandoc ---")
        logs.append(result.stderr.strip())

    if result.returncode != 0:
        raise ConversionError(
            f"pandoc terminó con código de error {result.returncode}.\n\n{result.stderr.strip()}"
        )

    logs.append(f"Documento generado correctamente en:\n{out}")
    if toc:
        logs.append(
            "Nota: al abrir el .docx en Word, actualiza el índice (clic derecho > Actualizar campo, o F9)."
        )

    return out, "\n".join(logs)


# ==============================================================================
# INTERFAZ GRÁFICA (TKINTER)
# ==============================================================================

def open_path_in_os(target: Path) -> None:
    """Abre un archivo o carpeta con la aplicación predeterminada del sistema."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(target))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(target)], check=False)
        else:
            subprocess.run(["xdg-open", str(target)], check=False)
    except Exception as exc:
        raise ConversionError(f"No se pudo abrir {target}:\n{exc}") from exc


def launch_gui() -> None:
    """Inicia la interfaz gráfica de usuario."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError:
        show_fatal_error(
            "Falta tkinter",
            "Tu instalación de Python no incluye el módulo gráfico 'tkinter'.\n"
            "Puedes usar el script desde la terminal con:\npython md2docx.py --cli",
        )
        return

    root = tk.Tk()
    root.title("md2docx — Conversor de Markdown a Word")
    root.geometry("680x550")
    root.minsize(620, 510)

    root.lift()
    root.attributes("-topmost", True)
    root.after(200, lambda: root.attributes("-topmost", False))

    style = ttk.Style(root)
    available_themes = style.theme_names()
    for preferred in ("vista", "xpnative", "clam", "aqua"):
        if preferred in available_themes:
            style.theme_use(preferred)
            break

    bg_color = "#f6f7f9"
    root.configure(bg=bg_color)
    style.configure("TFrame", background=bg_color)
    style.configure("TLabelframe", background=bg_color)
    style.configure("TLabelframe.Label", background=bg_color, font=("Segoe UI", 10, "bold"))
    style.configure("TLabel", background=bg_color, font=("Segoe UI", 9))
    style.configure("Header.TLabel", background=bg_color, font=("Segoe UI", 14, "bold"), foreground="#1f2937")
    style.configure("Subheader.TLabel", background=bg_color, font=("Segoe UI", 9), foreground="#4b5563")
    style.configure("Path.TLabel", background="#eef2f6", foreground="#1e3a8a", font=("Consolas", 9), padding=6)
    style.configure("Warn.TLabel", background=bg_color, foreground="#b45309", font=("Segoe UI", 8, "italic"))
    style.configure("TCheckbutton", background=bg_color, font=("Segoe UI", 9))
    style.configure("TRadiobutton", background=bg_color, font=("Segoe UI", 9))
    style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 7))

    md_var = tk.StringVar()
    tpl_var = tk.StringVar(value=str(DEFAULT_TEMPLATE) if DEFAULT_TEMPLATE.exists() else "")
    use_tpl_var = tk.BooleanVar(value=True)
    toc_var = tk.BooleanVar(value=False)
    num_var = tk.BooleanVar(value=False)
    conflict_var = tk.StringVar(value="ask")
    dest_preview_var = tk.StringVar(value="Selecciona un archivo .md para ver la ruta de salida...")
    dest_status_var = tk.StringVar(value="")

    last_generated_file: dict[str, Optional[Path]] = {"path": None}

    def update_destination_preview(*_args) -> None:
        raw_md = md_var.get().strip()
        if not raw_md:
            dest_preview_var.set("Selecciona un archivo .md para ver la ruta de salida...")
            dest_status_var.set("")
            return

        md_path = clean_path(raw_md)
        out_path = md_path.with_suffix(".docx")
        dest_preview_var.set(str(out_path))

        if out_path.exists():
            dest_status_var.set("Aviso: ya existe un .docx con ese nombre en la carpeta.")
        else:
            dest_status_var.set("Se creará un archivo nuevo en la misma carpeta del .md.")

    def browse_md() -> None:
        initial_dir = str(clean_path(md_var.get()).parent) if md_var.get().strip() else str(Path.home())
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Markdown (.md)",
            initialdir=initial_dir,
            filetypes=[("Archivos Markdown", "*.md *.markdown"), ("Todos los archivos", "*.*")],
        )
        if file_path:
            md_var.set(file_path)

    def browse_template() -> None:
        initial_dir = str(TEMPLATE_DIR) if TEMPLATE_DIR.exists() else str(Path.home())
        file_path = filedialog.askopenfilename(
            title="Seleccionar plantilla Word (.docx)",
            initialdir=initial_dir,
            filetypes=[("Documentos de Word", "*.docx")],
        )
        if file_path:
            tpl_var.set(file_path)
            use_tpl_var.set(True)

    def toggle_template_state() -> None:
        state = "normal" if use_tpl_var.get() else "disabled"
        tpl_entry.configure(state=state)
        tpl_browse_btn.configure(state=state)

    def handle_init_template() -> None:
        try:
            if DEFAULT_TEMPLATE.exists():
                confirm = messagebox.askyesno(
                    "Sobrescribir plantilla",
                    f"Ya existe '{DEFAULT_TEMPLATE.name}'.\n¿Deseas regenerarla desde cero?",
                )
                if not confirm:
                    return
            created = create_default_template(overwrite=True)
            tpl_var.set(str(created))
            use_tpl_var.set(True)
            toggle_template_state()
            append_log(f"Plantilla base creada en:\n{created}")
            if messagebox.askyesno(
                "Plantilla creada",
                f"Plantilla creada en:\n{created}\n\n¿Quieres abrirla ahora en Word para editar sus estilos?",
            ):
                open_path_in_os(created)
        except Exception as exc:
            messagebox.showerror("Error al crear plantilla", str(exc))

    def append_log(text: str, clear: bool = False) -> None:
        log_text.configure(state="normal")
        if clear:
            log_text.delete("1.0", tk.END)
        log_text.insert(tk.END, text + "\n")
        log_text.see(tk.END)
        log_text.configure(state="disabled")

    def gui_ask_conflict(existing_path: Path) -> str:
        answer = messagebox.askyesnocancel(
            "El archivo ya existe",
            f"El archivo '{existing_path.name}' ya existe en esa carpeta.\n\n"
            "• Sí: Sobrescribir el archivo existente\n"
            "• No: Guardar como nueva versión (_v2, _v3...)\n"
            "• Cancelar: Abortar conversión",
        )
        if answer is True:
            return "overwrite"
        if answer is False:
            return "version"
        return "cancel"

    def handle_convert() -> None:
        raw_md = md_var.get().strip()
        if not raw_md:
            messagebox.showwarning("Falta archivo", "Por favor, selecciona un archivo Markdown (.md).")
            return

        md_path = clean_path(raw_md)
        use_tpl = use_tpl_var.get()
        tpl_path = clean_path(tpl_var.get()) if (use_tpl and tpl_var.get().strip()) else None

        btn_open_doc.configure(state="disabled")
        btn_open_folder.configure(state="disabled")
        last_generated_file["path"] = None

        try:
            out_path, log_output = run_conversion(
                md=md_path,
                template=tpl_path,
                no_template=not use_tpl,
                output=None,
                toc=toc_var.get(),
                number_sections=num_var.get(),
                lang="es-ES",
                conflict_mode=conflict_var.get(),
                ask_callback=gui_ask_conflict,
            )
            append_log(log_output, clear=True)
            last_generated_file["path"] = out_path
            btn_open_doc.configure(state="normal")
            btn_open_folder.configure(state="normal")
            update_destination_preview()
            messagebox.showinfo("Conversión completada", f"Archivo guardado en:\n{out_path}")

        except ConversionError as exc:
            append_log(f"ERROR: {exc}", clear=True)
            if "cancelada" not in str(exc).lower():
                messagebox.showerror("Error en la conversión", str(exc))
        except Exception as exc:
            append_log(f"ERROR INESPERADO: {exc}", clear=True)
            messagebox.showerror("Error inesperado", str(exc))

    def open_generated_doc() -> None:
        if last_generated_file["path"] and last_generated_file["path"].exists():
            try:
                open_path_in_os(last_generated_file["path"])
            except ConversionError as exc:
                messagebox.showerror("Error", str(exc))

    def open_generated_folder() -> None:
        if last_generated_file["path"]:
            try:
                open_path_in_os(last_generated_file["path"].parent)
            except ConversionError as exc:
                messagebox.showerror("Error", str(exc))

    main_frame = ttk.Frame(root, padding=18)
    main_frame.pack(fill="both", expand=True)

    ttk.Label(main_frame, text="Conversor Markdown a Word (.docx)", style="Header.TLabel").pack(anchor="w")
    ttk.Label(
        main_frame,
        text="Selecciona el archivo .md; el documento Word se guardará automáticamente en esa misma ruta.",
        style="Subheader.TLabel",
    ).pack(anchor="w", pady=(2, 12))

    io_frame = ttk.LabelFrame(main_frame, text=" 1. Archivo de origen y destino ", padding=10)
    io_frame.pack(fill="x", pady=(0, 10))
    io_frame.columnconfigure(1, weight=1)

    ttk.Label(io_frame, text="Archivo .md:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
    md_entry = ttk.Entry(io_frame, textvariable=md_var)
    md_entry.grid(row=0, column=1, sticky="ew", pady=4)
    ttk.Button(io_frame, text="Examinar...", command=browse_md).grid(row=0, column=2, padx=(8, 0), pady=4)

    ttk.Label(io_frame, text="Se guardará en:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(6, 2))
    ttk.Label(io_frame, textvariable=dest_preview_var, style="Path.TLabel", anchor="w").grid(
        row=1, column=1, columnspan=2, sticky="ew", pady=(6, 2)
    )
    ttk.Label(io_frame, textvariable=dest_status_var, style="Warn.TLabel").grid(
        row=2, column=1, columnspan=2, sticky="w", pady=(2, 0)
    )

    md_var.trace_add("write", update_destination_preview)

    tpl_frame = ttk.LabelFrame(main_frame, text=" 2. Plantilla de estilos (.docx) ", padding=10)
    tpl_frame.pack(fill="x", pady=(0, 10))
    tpl_frame.columnconfigure(1, weight=1)

    ttk.Checkbutton(
        tpl_frame,
        text="Aplicar plantilla de Word (estilos, márgenes, encabezado y pie)",
        variable=use_tpl_var,
        command=toggle_template_state,
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

    ttk.Label(tpl_frame, text="Plantilla:").grid(row=1, column=0, sticky="w", padx=(0, 8))
    tpl_entry = ttk.Entry(tpl_frame, textvariable=tpl_var)
    tpl_entry.grid(row=1, column=1, sticky="ew")
    tpl_browse_btn = ttk.Button(tpl_frame, text="Elegir...", command=browse_template)
    tpl_browse_btn.grid(row=1, column=2, padx=(8, 0))

    ttk.Button(
        tpl_frame,
        text="Crear / regenerar plantilla_base.docx",
        command=handle_init_template,
    ).grid(row=2, column=1, columnspan=2, sticky="e", pady=(8, 0))

    opts_frame = ttk.LabelFrame(main_frame, text=" 3. Opciones de documento ", padding=10)
    opts_frame.pack(fill="x", pady=(0, 12))
    opts_frame.columnconfigure(0, weight=1)
    opts_frame.columnconfigure(1, weight=1)

    left_opts = ttk.Frame(opts_frame)
    left_opts.grid(row=0, column=0, sticky="nw")
    ttk.Checkbutton(left_opts, text="Incluir tabla de contenidos / índice (TOC)", variable=toc_var).pack(anchor="w", pady=2)
    ttk.Checkbutton(left_opts, text="Numerar capítulos y secciones", variable=num_var).pack(anchor="w", pady=2)

    right_opts = ttk.Frame(opts_frame)
    right_opts.grid(row=0, column=1, sticky="nw")
    ttk.Label(right_opts, text="Si el .docx ya existe:").pack(anchor="w", pady=(0, 2))
    ttk.Radiobutton(right_opts, text="Preguntar antes", variable=conflict_var, value="ask").pack(anchor="w")
    ttk.Radiobutton(right_opts, text="Crear versión (_v2, _v3...)", variable=conflict_var, value="version").pack(anchor="w")
    ttk.Radiobutton(right_opts, text="Sobrescribir directamente", variable=conflict_var, value="overwrite").pack(anchor="w")

    actions_frame = ttk.Frame(main_frame)
    actions_frame.pack(fill="x", pady=(0, 10))

    ttk.Button(
        actions_frame,
        text="Convertir a Word (.docx)",
        style="Accent.TButton",
        command=handle_convert,
    ).pack(side="left")

    btn_open_folder = ttk.Button(
        actions_frame,
        text="Abrir carpeta",
        state="disabled",
        command=open_generated_folder,
    )
    btn_open_folder.pack(side="right", padx=(6, 0))

    btn_open_doc = ttk.Button(
        actions_frame,
        text="Abrir Word",
        state="disabled",
        command=open_generated_doc,
    )
    btn_open_doc.pack(side="right")

    log_frame = ttk.LabelFrame(main_frame, text=" Estado y avisos de pandoc ", padding=8)
    log_frame.pack(fill="both", expand=True)

    log_text = tk.Text(log_frame, height=6, wrap="word", font=("Consolas", 9), bg="#ffffff", fg="#1f2937", relief="flat")
    log_text.pack(side="left", fill="both", expand=True)
    scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    scrollbar.pack(side="right", fill="y")
    log_text.configure(yscrollcommand=scrollbar.set, state="disabled")

    append_log("Listo. Selecciona un archivo .md para comenzar.")
    toggle_template_state()

    root.mainloop()


# ==============================================================================
# MODO TERMINAL (CLI)
# ==============================================================================

def init_template_cli() -> None:
    """Genera la plantilla base desde terminal."""
    try:
        if DEFAULT_TEMPLATE.exists():
            resp = input(f"{DEFAULT_TEMPLATE.name} ya existe. ¿Sobrescribir? [s/N]: ").strip().lower()
            if resp != "s":
                print("Cancelado.")
                return
        created = create_default_template(overwrite=True)
        print(f"Plantilla creada: {created}")
    except ConversionError as exc:
        sys.exit(f"ERROR: {exc}")


def ask_interactive_cli() -> argparse.Namespace:
    """Modo interactivo clásico por consola (--cli)."""
    md = clean_path(input("Archivo .md: "))
    tpl_hint = f" [Enter = {DEFAULT_TEMPLATE.name}]" if DEFAULT_TEMPLATE.exists() else " [Enter = sin plantilla]"
    tpl_raw = input(f"Plantilla .docx{tpl_hint}: ").strip()
    toc = input("¿Índice (TOC)? [s/N]: ").strip().lower() == "s"
    num = input("¿Numerar secciones? [s/N]: ").strip().lower() == "s"
    return argparse.Namespace(
        input=md,
        template=clean_path(tpl_raw) if tpl_raw else None,
        no_template=False,
        output=None,
        toc=toc,
        number_sections=num,
        lang="es-ES",
        init_template=False,
        cli=True,
    )


def convert_cli(args: argparse.Namespace) -> None:
    """Ejecuta la conversión en modo línea de comandos."""
    try:
        out, log_output = run_conversion(
            md=args.input,
            template=args.template,
            no_template=args.no_template,
            output=args.output,
            toc=args.toc,
            number_sections=args.number_sections,
            lang=args.lang,
            conflict_mode="ask",
        )
        print(log_output)
    except ConversionError as exc:
        sys.exit(f"ERROR: {exc}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convierte .md a .docx con pandoc.")
    parser.add_argument("input", nargs="?", type=clean_path, help="Archivo .md de entrada")
    parser.add_argument("-t", "--template", type=clean_path, help="Plantilla .docx (reference-doc)")
    parser.add_argument("--no-template", action="store_true", help="No usar plantilla, ni la por defecto")
    parser.add_argument("-o", "--output", type=clean_path, help="Ruta del .docx de salida")
    parser.add_argument("--toc", action="store_true", help="Incluir índice")
    parser.add_argument("--number-sections", action="store_true", help="Numerar secciones")
    parser.add_argument("--lang", default="es-ES", help="Idioma del documento (por defecto es-ES)")
    parser.add_argument("--init-template", action="store_true", help="Crear plantillas/plantilla_base.docx")
    parser.add_argument("--cli", action="store_true", help="Forzar modo interactivo en consola en lugar de ventana")
    args = parser.parse_args()

    if args.init_template:
        init_template_cli()
        return

    if args.input is None:
        if args.cli:
            args = ask_interactive_cli()
            convert_cli(args)
        else:
            launch_gui()
        return

    convert_cli(args)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        show_fatal_error("Error inesperado en md2docx", traceback.format_exc())
