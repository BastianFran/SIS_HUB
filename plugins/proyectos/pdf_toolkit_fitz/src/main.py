from __future__ import annotations

import base64
import json
import re
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz

IMAGE_FILE_TYPES = [("Imagenes", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp")]
PDF_FONT_OPTIONS = ("helv", "cour", "tiro", "symb", "zapfdingbats")
EXPORT_IMAGE_FORMATS = ("png", "jpg")


def _ensure_pdf_path(path: str) -> Path:
    p = Path(path.strip())
    if not p.exists():
        raise FileNotFoundError(f"No existe el archivo: {p}")
    if p.suffix.lower() != ".pdf":
        raise ValueError(f"El archivo no es PDF: {p.name}")
    return p


def _safe_stem(text: str, fallback: str) -> str:
    value = re.sub(r"[^\w\-]+", "_", (text or "").strip(), flags=re.UNICODE).strip("_")
    return value or fallback


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for i in range(1, 1000):
        candidate = path.with_name(f"{path.stem}_{i}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"No se pudo generar un nombre unico para: {path}")


def _parse_page_num(token: str, total_pages: int) -> int:
    if not token.isdigit():
        raise ValueError(f"Pagina invalida: '{token}'")
    num = int(token)
    if num < 1 or num > total_pages:
        raise ValueError(
            f"Pagina fuera de rango ({num}). El documento tiene {total_pages} pagina(s)."
        )
    return num - 1


def _parse_page_spec(spec: str, total_pages: int) -> list[int]:
    text = (spec or "").strip()
    if not text:
        return list(range(total_pages))

    pages: set[int] = set()
    for token in text.split(","):
        part = token.strip()
        if not part:
            continue
        if "-" in part:
            left, right = part.split("-", 1)
            start = _parse_page_num(left.strip(), total_pages)
            end = _parse_page_num(right.strip(), total_pages)
            if start > end:
                raise ValueError(f"Rango invalido: '{part}'. Usa formato menor-mayor.")
            for pno in range(start, end + 1):
                pages.add(pno)
        else:
            pages.add(_parse_page_num(part, total_pages))

    if not pages:
        raise ValueError("No se detectaron paginas validas.")
    return sorted(pages)


def _parse_split_groups(spec: str, total_pages: int) -> list[tuple[str, list[int]]]:
    text = (spec or "").strip()
    if not text:
        return [(f"p{i + 1}", [i]) for i in range(total_pages)]

    groups: list[tuple[str, list[int]]] = []
    for token in text.split(","):
        part = token.strip()
        if not part:
            continue
        if "-" in part:
            left, right = part.split("-", 1)
            start = _parse_page_num(left.strip(), total_pages)
            end = _parse_page_num(right.strip(), total_pages)
            if start > end:
                raise ValueError(f"Rango invalido: '{part}'. Usa formato menor-mayor.")
            start_1 = start + 1
            end_1 = end + 1
            groups.append((f"p{start_1}-{end_1}", list(range(start, end + 1))))
        else:
            pno = _parse_page_num(part, total_pages)
            groups.append((f"p{pno + 1}", [pno]))

    if not groups:
        raise ValueError("No se detectaron rangos validos.")
    return groups


def _parse_ordered_page_spec(spec: str, total_pages: int) -> list[int]:
    text = (spec or "").strip()
    if not text:
        return list(range(total_pages))

    pages: list[int] = []
    for token in text.split(","):
        part = token.strip()
        if not part:
            continue
        if "-" in part:
            left, right = part.split("-", 1)
            start = _parse_page_num(left.strip(), total_pages)
            end = _parse_page_num(right.strip(), total_pages)
            step = 1 if end >= start else -1
            for pno in range(start, end + step, step):
                pages.append(pno)
        else:
            pages.append(_parse_page_num(part, total_pages))

    if not pages:
        raise ValueError("No se detectaron paginas validas para reordenar.")
    return pages


def _save_doc(doc: fitz.Document, out_path: Path, optimize: bool) -> None:
    if optimize:
        doc.save(out_path, garbage=4, deflate=True)
    else:
        doc.save(out_path)


def _assert_not_encrypted(doc: fitz.Document, path: Path) -> None:
    if getattr(doc, "needs_pass", False):
        raise ValueError(
            f"El PDF '{path.name}' esta protegido con contrasena. "
            "Desbloquealo antes de procesarlo en el Hub."
        )


def _parse_hex_color(text: str) -> tuple[float, float, float]:
    value = (text or "").strip().lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        raise ValueError("Color invalido. Usa formato #RRGGBB.")
    try:
        return tuple(int(value[i : i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError as exc:
        raise ValueError("Color invalido. Usa formato #RRGGBB.") from exc


class PDFToolkitUI:
    def __init__(self, parent, context):
        self.context = context
        self.log = context.logger()

        self.root = ttk.Frame(parent, style="TFrame")
        self.root.pack(fill="both", expand=True, padx=12, pady=12)

        self.running = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Listo.")

        self.output_dir_var = tk.StringVar(value=str(context.data_dir / "pdf_toolkit"))
        self.output_name_var = tk.StringVar(value="resultado_pdf")
        self.output_file_var = tk.StringVar(value="")

        self.merge_files: list[Path] = []
        self.create_images: list[Path] = []
        self.template_actions: list[dict] = []

        self.template_pdf_doc: fitz.Document | None = None
        self.template_pdf_path: Path | None = None
        self.template_page_index = 0
        self.preview_photo: tk.PhotoImage | None = None
        self.preview_origin = (12.0, 12.0)
        self.preview_image_size = (0.0, 0.0)
        self.preview_drag_start: tuple[float, float] | None = None
        self.preview_drag_rect_id: int | None = None
        self.preview_action_indexes: list[int] = []

        self.action_buttons: list[ttk.Button] = []

        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(
            self.root,
            text="Herramienta PDF con PyMuPDF (fitz)",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            self.root,
            text=(
                "Funciones comunes de manejo PDF en un solo plugin: combinar, dividir, "
                "editar, crear y extraer."
            ),
        ).pack(anchor="w", pady=(0, 8))

        common = ttk.LabelFrame(self.root, text="Salida del proceso")
        common.pack(fill="x", pady=(0, 8))

        row1 = ttk.Frame(common)
        row1.pack(fill="x", padx=8, pady=8)
        ttk.Label(row1, text="Carpeta destino:").pack(side="left", padx=(0, 6))
        ttk.Entry(row1, textvariable=self.output_dir_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(row1, text="Examinar...", command=self._choose_output_dir).pack(
            side="left", padx=6
        )

        row2 = ttk.Frame(common)
        row2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(row2, text="Nombre base:").pack(side="left", padx=(0, 6))
        ttk.Entry(row2, textvariable=self.output_name_var, width=30).pack(side="left")
        ttk.Button(row2, text="Limpiar log", command=self._clear_log).pack(side="left", padx=8)

        row3 = ttk.Frame(common)
        row3.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(row3, text="Guardar como (1 archivo):").pack(side="left", padx=(0, 6))
        ttk.Entry(row3, textvariable=self.output_file_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row3, text="Elegir archivo...", command=self._choose_output_file).pack(
            side="left", padx=6
        )
        ttk.Button(row3, text="Limpiar", command=lambda: self.output_file_var.set("")).pack(
            side="left"
        )

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True)

        self._build_merge_tab()
        self._build_split_tab()
        self._build_edit_tab()
        self._build_security_tab()
        self._build_organize_tab()
        self._build_create_tab()
        self._build_extract_tab()
        self._build_export_images_tab()
        self._build_template_tab()
        self._build_info_tab()

        footer = ttk.Frame(self.root)
        footer.pack(fill="x", pady=(8, 0))
        self.progress = ttk.Progressbar(footer, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Label(footer, textvariable=self.status_var).pack(side="left")

        self.log_box = tk.Text(self.root, height=8, wrap="word")
        self.log_box.pack(fill="both", expand=False, pady=(8, 0))
        self._append_log("Plugin PDF inicializado.")

    def _build_merge_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Combinar")

        top = ttk.Frame(tab)
        top.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(
            top, text="Selecciona 2 o mas PDFs para combinarlos en un archivo."
        ).pack(anchor="w")

        buttons = ttk.Frame(tab)
        buttons.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Button(buttons, text="Agregar PDFs...", command=self._add_merge_files).pack(
            side="left"
        )
        ttk.Button(buttons, text="Quitar seleccionado", command=self._remove_merge_file).pack(
            side="left", padx=6
        )
        ttk.Button(buttons, text="Limpiar lista", command=self._clear_merge_files).pack(
            side="left", padx=6
        )

        self.merge_list = tk.Listbox(tab, height=8, exportselection=False)
        self.merge_list.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        opts = ttk.Frame(tab)
        opts.pack(fill="x", padx=10, pady=(0, 10))
        self.merge_opt_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts,
            text="Optimizar al guardar (reduce tamano cuando es posible)",
            variable=self.merge_opt_var,
        ).pack(side="left")
        btn = ttk.Button(opts, text="Combinar PDFs", style="Primary.TButton", command=self._run_merge)
        btn.pack(side="right")
        self.action_buttons.append(btn)

    def _build_split_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Dividir")

        row1 = ttk.Frame(tab)
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(row1, text="PDF origen:").pack(side="left", padx=(0, 6))
        self.split_file_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.split_file_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row1, text="Buscar...", command=lambda: self._pick_pdf(self.split_file_var)).pack(
            side="left", padx=6
        )

        row2 = ttk.Frame(tab)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        self.split_mode_var = tk.StringVar(value="each")
        ttk.Radiobutton(
            row2, text="Una salida por pagina", value="each", variable=self.split_mode_var
        ).pack(side="left")
        ttk.Radiobutton(
            row2, text="Usar rangos", value="ranges", variable=self.split_mode_var
        ).pack(side="left", padx=12)

        row3 = ttk.Frame(tab)
        row3.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(row3, text="Rangos (si aplica):").pack(side="left", padx=(0, 6))
        self.split_ranges_var = tk.StringVar(value="1-3,4-6")
        ttk.Entry(row3, textvariable=self.split_ranges_var).pack(side="left", fill="x", expand=True)

        self.split_opt_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            tab,
            text="Optimizar archivos de salida",
            variable=self.split_opt_var,
        ).pack(anchor="w", padx=10, pady=(0, 10))

        btn = ttk.Button(tab, text="Dividir PDF", style="Primary.TButton", command=self._run_split)
        btn.pack(anchor="e", padx=10, pady=(0, 10))
        self.action_buttons.append(btn)

    def _build_edit_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Editar")

        row1 = ttk.Frame(tab)
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(row1, text="PDF origen:").pack(side="left", padx=(0, 6))
        self.edit_file_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.edit_file_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row1, text="Buscar...", command=lambda: self._pick_pdf(self.edit_file_var)).pack(
            side="left", padx=6
        )

        row2 = ttk.Frame(tab)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row2, text="Paginas (opcional):").pack(side="left", padx=(0, 6))
        self.edit_pages_var = tk.StringVar(value="")
        ttk.Entry(row2, textvariable=self.edit_pages_var).pack(side="left", fill="x", expand=True)
        ttk.Label(row2, text="Ej: 1-3,7").pack(side="left", padx=6)

        row3 = ttk.Frame(tab)
        row3.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row3, text="Rotar:").pack(side="left", padx=(0, 6))
        self.edit_rotate_var = tk.StringVar(value="0")
        combo = ttk.Combobox(
            row3,
            textvariable=self.edit_rotate_var,
            values=["0", "90", "180", "270"],
            state="readonly",
            width=8,
        )
        combo.pack(side="left")
        ttk.Label(row3, text="grados").pack(side="left", padx=6)

        row4 = ttk.Frame(tab)
        row4.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row4, text="Marca de agua (texto):").pack(side="left", padx=(0, 6))
        self.edit_watermark_var = tk.StringVar(value="")
        ttk.Entry(row4, textvariable=self.edit_watermark_var).pack(
            side="left", fill="x", expand=True
        )

        meta = ttk.LabelFrame(tab, text="Metadatos (opcional)")
        meta.pack(fill="x", padx=10, pady=(0, 6))
        self.meta_title_var = tk.StringVar(value="")
        self.meta_author_var = tk.StringVar(value="")
        self.meta_subject_var = tk.StringVar(value="")
        self.meta_keywords_var = tk.StringVar(value="")

        self._meta_row(meta, "Titulo:", self.meta_title_var)
        self._meta_row(meta, "Autor:", self.meta_author_var)
        self._meta_row(meta, "Asunto:", self.meta_subject_var)
        self._meta_row(meta, "Keywords:", self.meta_keywords_var)

        self.edit_opt_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(tab, text="Optimizar al guardar", variable=self.edit_opt_var).pack(
            anchor="w", padx=10, pady=(0, 10)
        )

        btn = ttk.Button(tab, text="Aplicar cambios", style="Primary.TButton", command=self._run_edit)
        btn.pack(anchor="e", padx=10, pady=(0, 10))
        self.action_buttons.append(btn)

    def _build_security_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Seguridad")

        self.security_file_var = tk.StringVar(value="")
        self.security_mode_var = tk.StringVar(value="encrypt")
        self.security_current_pw_var = tk.StringVar(value="")
        self.security_user_pw_var = tk.StringVar(value="")
        self.security_owner_pw_var = tk.StringVar(value="")
        self.security_algo_var = tk.StringVar(value="AES-256")
        self.security_opt_var = tk.BooleanVar(value=True)

        row1 = ttk.Frame(tab)
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(row1, text="PDF origen:").pack(side="left", padx=(0, 6))
        ttk.Entry(row1, textvariable=self.security_file_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            row1, text="Buscar...", command=lambda: self._pick_pdf(self.security_file_var)
        ).pack(side="left", padx=6)

        row2 = ttk.Frame(tab)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Radiobutton(
            row2, text="Proteger PDF", value="encrypt", variable=self.security_mode_var
        ).pack(side="left")
        ttk.Radiobutton(
            row2, text="Quitar contrasena", value="decrypt", variable=self.security_mode_var
        ).pack(side="left", padx=10)

        row3 = ttk.Frame(tab)
        row3.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row3, text="Contrasena actual (si el PDF ya esta protegido):").pack(
            side="left", padx=(0, 6)
        )
        ttk.Entry(row3, textvariable=self.security_current_pw_var, show="*").pack(
            side="left", fill="x", expand=True
        )

        row4 = ttk.Frame(tab)
        row4.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row4, text="Contrasena usuario:").pack(side="left", padx=(0, 6))
        ttk.Entry(row4, textvariable=self.security_user_pw_var, show="*").pack(
            side="left", fill="x", expand=True
        )
        ttk.Label(row4, text="Contrasena owner:").pack(side="left", padx=(10, 6))
        ttk.Entry(row4, textvariable=self.security_owner_pw_var, show="*").pack(
            side="left", fill="x", expand=True
        )

        row5 = ttk.Frame(tab)
        row5.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Label(row5, text="Algoritmo:").pack(side="left", padx=(0, 6))
        ttk.Combobox(
            row5,
            textvariable=self.security_algo_var,
            values=["AES-256", "AES-128", "RC4-128"],
            state="readonly",
            width=12,
        ).pack(side="left")
        ttk.Checkbutton(row5, text="Optimizar al guardar", variable=self.security_opt_var).pack(
            side="left", padx=12
        )
        ttk.Label(
            row5,
            text=(
                "Nota: en modo 'Quitar contrasena' solo se necesita contrasena actual si aplica."
            ),
        ).pack(side="left", padx=8)

        btn = ttk.Button(
            tab,
            text="Aplicar seguridad",
            style="Primary.TButton",
            command=self._run_security,
        )
        btn.pack(anchor="e", padx=10, pady=(0, 10))
        self.action_buttons.append(btn)

    def _build_organize_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Organizar")

        row1 = ttk.Frame(tab)
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(row1, text="PDF origen:").pack(side="left", padx=(0, 6))
        self.organize_file_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.organize_file_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            row1, text="Buscar...", command=lambda: self._pick_pdf(self.organize_file_var)
        ).pack(side="left", padx=6)

        row2 = ttk.Frame(tab)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row2, text="Nuevo orden de paginas:").pack(side="left", padx=(0, 6))
        self.organize_order_var = tk.StringVar(value="")
        ttk.Entry(row2, textvariable=self.organize_order_var).pack(
            side="left", fill="x", expand=True
        )

        row3 = ttk.Frame(tab)
        row3.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(
            row3,
            text=(
                "Ejemplos: 1-5 (igual), 3,1,2,4-5 (reordenar), 1,1,2 (duplicar), "
                "8-1 (invertir). Vacio = sin cambios."
            ),
            wraplength=900,
            justify="left",
        ).pack(anchor="w")

        self.organize_opt_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(tab, text="Optimizar al guardar", variable=self.organize_opt_var).pack(
            anchor="w", padx=10, pady=(0, 10)
        )

        btn = ttk.Button(
            tab, text="Aplicar organizacion", style="Primary.TButton", command=self._run_organize
        )
        btn.pack(anchor="e", padx=10, pady=(0, 10))
        self.action_buttons.append(btn)

    def _build_create_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Crear PDF")

        ttk.Label(
            tab,
            text="Crear un PDF a partir de imagenes (PNG, JPG, BMP, TIFF, WEBP).",
        ).pack(anchor="w", padx=10, pady=(10, 6))

        btns = ttk.Frame(tab)
        btns.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Button(btns, text="Agregar imagenes...", command=self._add_images).pack(side="left")
        ttk.Button(btns, text="Quitar seleccionada", command=self._remove_image).pack(
            side="left", padx=6
        )
        ttk.Button(btns, text="Limpiar lista", command=self._clear_images).pack(side="left", padx=6)

        self.images_list = tk.Listbox(tab, height=8, exportselection=False)
        self.images_list.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        self.create_opt_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            tab, text="Optimizar al guardar", variable=self.create_opt_var
        ).pack(anchor="w", padx=10, pady=(0, 10))

        btn = ttk.Button(tab, text="Crear PDF", style="Primary.TButton", command=self._run_create)
        btn.pack(anchor="e", padx=10, pady=(0, 10))
        self.action_buttons.append(btn)

    def _build_extract_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Extraer")

        row1 = ttk.Frame(tab)
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(row1, text="PDF origen:").pack(side="left", padx=(0, 6))
        self.extract_file_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.extract_file_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            row1, text="Buscar...", command=lambda: self._pick_pdf(self.extract_file_var)
        ).pack(side="left", padx=6)

        row2 = ttk.Frame(tab)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row2, text="Paginas (opcional):").pack(side="left", padx=(0, 6))
        self.extract_pages_var = tk.StringVar(value="")
        ttk.Entry(row2, textvariable=self.extract_pages_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Label(row2, text="Ej: 1-2,5").pack(side="left", padx=6)

        self.extract_text_var = tk.BooleanVar(value=True)
        self.extract_images_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(tab, text="Extraer texto a .txt", variable=self.extract_text_var).pack(
            anchor="w", padx=10, pady=(0, 2)
        )
        ttk.Checkbutton(tab, text="Extraer imagenes embebidas", variable=self.extract_images_var).pack(
            anchor="w", padx=10, pady=(0, 10)
        )

        btn = ttk.Button(tab, text="Extraer contenido", style="Primary.TButton", command=self._run_extract)
        btn.pack(anchor="e", padx=10, pady=(0, 10))
        self.action_buttons.append(btn)

    def _build_export_images_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="PDF a Imagenes")

        row1 = ttk.Frame(tab)
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(row1, text="PDF origen:").pack(side="left", padx=(0, 6))
        self.export_img_file_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.export_img_file_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            row1, text="Buscar...", command=lambda: self._pick_pdf(self.export_img_file_var)
        ).pack(side="left", padx=6)

        row2 = ttk.Frame(tab)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row2, text="Paginas (opcional):").pack(side="left", padx=(0, 6))
        self.export_img_pages_var = tk.StringVar(value="")
        ttk.Entry(row2, textvariable=self.export_img_pages_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Label(row2, text="Ej: 1-3,8").pack(side="left", padx=6)

        row3 = ttk.Frame(tab)
        row3.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row3, text="Formato:").pack(side="left", padx=(0, 6))
        self.export_img_format_var = tk.StringVar(value="png")
        ttk.Combobox(
            row3,
            textvariable=self.export_img_format_var,
            values=list(EXPORT_IMAGE_FORMATS),
            state="readonly",
            width=8,
        ).pack(side="left")
        ttk.Label(row3, text="DPI:").pack(side="left", padx=(10, 6))
        self.export_img_dpi_var = tk.StringVar(value="150")
        ttk.Entry(row3, textvariable=self.export_img_dpi_var, width=8).pack(side="left")
        self.export_img_gray_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row3, text="Escala de grises", variable=self.export_img_gray_var
        ).pack(side="left", padx=12)

        btn = ttk.Button(
            tab,
            text="Exportar paginas a imagenes",
            style="Primary.TButton",
            command=self._run_export_images,
        )
        btn.pack(anchor="e", padx=10, pady=(0, 10))
        self.action_buttons.append(btn)

    def _build_template_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Plantilla/Hotspots")

        self.template_file_var = tk.StringVar(value="")
        self.template_zoom_var = tk.StringVar(value="125")
        self.template_mode_var = tk.StringVar(value="none")
        self.template_page_var = tk.StringVar(value="Pagina 0/0")
        self.hotspot_name_var = tk.StringVar(value="campo_1")
        self.hotspot_default_var = tk.StringVar(value="")
        self.hotspot_font_var = tk.StringVar(value="helv")
        self.hotspot_size_var = tk.StringVar(value="12")
        self.hotspot_color_var = tk.StringVar(value="#111111")
        self.template_opt_var = tk.BooleanVar(value=True)
        self.template_flatten_var = tk.BooleanVar(value=False)
        self.template_redact_var = tk.BooleanVar(value=False)
        self.template_fields_var = tk.BooleanVar(value=True)
        self.template_dpi_var = tk.StringVar(value="144")

        row1 = ttk.Frame(tab)
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(row1, text="PDF origen:").pack(side="left", padx=(0, 6))
        ttk.Entry(row1, textvariable=self.template_file_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            row1, text="Buscar...", command=lambda: self._pick_pdf(self.template_file_var)
        ).pack(side="left", padx=6)
        ttk.Button(row1, text="Cargar preview", command=self._load_template_pdf).pack(
            side="left"
        )

        row2 = ttk.Frame(tab)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Button(row2, text="< Pagina", command=self._template_prev_page).pack(side="left")
        ttk.Button(row2, text="Pagina >", command=self._template_next_page).pack(
            side="left", padx=6
        )
        ttk.Label(row2, textvariable=self.template_page_var).pack(side="left", padx=(6, 16))
        ttk.Label(row2, text="Zoom %:").pack(side="left")
        ttk.Combobox(
            row2,
            textvariable=self.template_zoom_var,
            values=["50", "75", "100", "125", "150", "200", "300"],
            state="readonly",
            width=8,
        ).pack(side="left", padx=6)
        ttk.Button(row2, text="Aplicar zoom", command=self._render_template_preview).pack(side="left")

        row3 = ttk.Frame(tab)
        row3.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(row3, text="Modo dibujo:").pack(side="left", padx=(0, 6))
        ttk.Radiobutton(
            row3, text="Ninguno", value="none", variable=self.template_mode_var
        ).pack(side="left")
        ttk.Radiobutton(
            row3, text="Cuadro blanco", value="white", variable=self.template_mode_var
        ).pack(side="left", padx=8)
        ttk.Radiobutton(
            row3, text="Hotspot texto", value="hotspot", variable=self.template_mode_var
        ).pack(side="left")

        row4 = ttk.Frame(tab)
        row4.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(row4, text="Campo:").pack(side="left")
        ttk.Entry(row4, textvariable=self.hotspot_name_var, width=16).pack(side="left", padx=6)
        ttk.Label(row4, text="Texto por defecto:").pack(side="left")
        ttk.Entry(row4, textvariable=self.hotspot_default_var, width=22).pack(
            side="left", padx=6
        )
        ttk.Label(row4, text="Fuente:").pack(side="left")
        ttk.Combobox(
            row4,
            textvariable=self.hotspot_font_var,
            values=list(PDF_FONT_OPTIONS),
            state="readonly",
            width=12,
        ).pack(side="left", padx=6)
        ttk.Label(row4, text="Tam:").pack(side="left")
        ttk.Entry(row4, textvariable=self.hotspot_size_var, width=6).pack(side="left", padx=4)
        ttk.Label(row4, text="Color:").pack(side="left")
        ttk.Entry(row4, textvariable=self.hotspot_color_var, width=10).pack(side="left", padx=4)

        row5 = ttk.Frame(tab)
        row5.pack(fill="x", padx=10, pady=(0, 8))
        btn_validate = ttk.Button(
            row5, text="Validar plantilla", command=self._validate_template_ui
        )
        btn_validate.pack(side="left")
        btn_apply_main = ttk.Button(
            row5,
            text="Aplicar plantilla y guardar PDF",
            style="Primary.TButton",
            command=self._run_template_apply,
        )
        btn_apply_main.pack(side="left", padx=8)
        btn_flatten_main = ttk.Button(
            row5, text="Solo digitalizar y guardar", command=self._run_template_flatten
        )
        btn_flatten_main.pack(side="left")
        ttk.Label(
            row5,
            text="Flujo recomendado: cargar plantilla -> validar -> aplicar y guardar.",
        ).pack(side="left", padx=10)

        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        right = ttk.LabelFrame(body, text="Acciones y plantilla")
        right.pack(side="left", fill="y", padx=(10, 0))

        canvas_wrap = ttk.Frame(left)
        canvas_wrap.pack(fill="both", expand=True)
        canvas_wrap.columnconfigure(0, weight=1)
        canvas_wrap.rowconfigure(0, weight=1)

        self.template_canvas = tk.Canvas(
            canvas_wrap, bg="#101820", highlightthickness=0, cursor="crosshair"
        )
        self.template_canvas.grid(row=0, column=0, sticky="nsew")
        sx = ttk.Scrollbar(canvas_wrap, orient="horizontal", command=self.template_canvas.xview)
        sy = ttk.Scrollbar(canvas_wrap, orient="vertical", command=self.template_canvas.yview)
        sx.grid(row=1, column=0, sticky="ew")
        sy.grid(row=0, column=1, sticky="ns")
        self.template_canvas.configure(xscrollcommand=sx.set, yscrollcommand=sy.set)
        self.template_canvas.bind("<ButtonPress-1>", self._template_canvas_press)
        self.template_canvas.bind("<B1-Motion>", self._template_canvas_drag)
        self.template_canvas.bind("<ButtonRelease-1>", self._template_canvas_release)

        ttk.Label(
            right,
            text=(
                "Uso: selecciona modo, arrastra con mouse en la previsualizacion y luego "
                "aplica plantilla."
            ),
            wraplength=300,
            justify="left",
        ).pack(anchor="w", padx=8, pady=(8, 6))

        self.template_actions_list = tk.Listbox(right, height=11, exportselection=False)
        self.template_actions_list.pack(fill="x", padx=8, pady=(0, 6))

        act_btns = ttk.Frame(right)
        act_btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(act_btns, text="Quitar seleccion", command=self._remove_template_action).pack(
            side="left"
        )
        ttk.Button(act_btns, text="Deshacer ultima", command=self._undo_template_action).pack(
            side="left", padx=6
        )
        ttk.Button(act_btns, text="Limpiar pagina", command=self._clear_template_page).pack(
            side="left"
        )

        tpl_btns = ttk.Frame(right)
        tpl_btns.pack(fill="x", padx=8, pady=(0, 6))
        ttk.Button(tpl_btns, text="Guardar plantilla...", command=self._save_template_file).pack(
            side="left"
        )
        ttk.Button(tpl_btns, text="Cargar plantilla...", command=self._load_template_file).pack(
            side="left", padx=6
        )

        ttk.Label(
            right,
            text="Valores de hotspots (formato campo=valor, una linea por campo):",
            wraplength=300,
            justify="left",
        ).pack(anchor="w", padx=8, pady=(0, 4))
        self.hotspot_values_box = tk.Text(right, height=9, wrap="word")
        self.hotspot_values_box.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        ttk.Button(
            right, text="Autocompletar campos", command=self._template_seed_values
        ).pack(anchor="w", padx=8, pady=(0, 8))

        ttk.Checkbutton(
            right,
            text="Digitalizar antes de aplicar (flatten)",
            variable=self.template_flatten_var,
        ).pack(anchor="w", padx=8, pady=(0, 2))
        ttk.Checkbutton(
            right,
            text="Borrar texto real en cuadros blancos (redaccion)",
            variable=self.template_redact_var,
        ).pack(anchor="w", padx=8, pady=(0, 2))
        ttk.Checkbutton(
            right,
            text="Guardar hotspots como campos editables (AcroForm)",
            variable=self.template_fields_var,
        ).pack(anchor="w", padx=8, pady=(0, 2))
        ttk.Checkbutton(
            right, text="Optimizar al guardar", variable=self.template_opt_var
        ).pack(anchor="w", padx=8, pady=(0, 2))
        dpi_row = ttk.Frame(right)
        dpi_row.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(dpi_row, text="DPI digitalizacion:").pack(side="left")
        ttk.Entry(dpi_row, textvariable=self.template_dpi_var, width=8).pack(side="left", padx=6)

        self.action_buttons.append(btn_validate)
        self.action_buttons.append(btn_apply_main)
        self.action_buttons.append(btn_flatten_main)

    def _build_info_tab(self) -> None:
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Referencia")

        text = tk.Text(tab, wrap="word")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        text.insert(
            "1.0",
            (
                "Tipos/subconjuntos PDF mas comunes:\n"
                "- PDF (general, ISO 32000).\n"
                "- PDF/A (archivo y preservacion a largo plazo).\n"
                "- PDF/X (preprensa e impresion).\n"
                "- PDF/UA (accesibilidad universal).\n"
                "- PDF/E (ingenieria), PDF/VT (impresion variable).\n\n"
                "Herramientas mas usadas en flujos PDF:\n"
                "- Combinar y dividir documentos.\n"
                "- Rotar/reordenar/eliminar paginas.\n"
                "- Marcas de agua y metadatos.\n"
                "- Extraer texto e imagenes.\n"
                "- Crear PDF desde imagenes/escaneos.\n\n"
                "Funciones agregadas en este plugin:\n"
                "- Guardar como (ruta exacta para salidas de 1 archivo).\n"
                "- Organizar paginas por orden personalizado.\n"
                "- Exportar paginas PDF a PNG/JPG.\n"
                "- Proteger o quitar contrasena del PDF.\n"
                "- Plantillas con hotspots y campos editables AcroForm.\n\n"
                "Complicaciones frecuentes:\n"
                "- PDFs con contrasena: primero desbloquear.\n"
                "- Escaneos sin texto: requiere OCR externo.\n"
                "- PDFs firmados: cualquier edicion invalida firmas.\n"
                "- Subestandares (PDF/A, PDF/X, PDF/UA) requieren validacion dedicada.\n\n"
                "Este plugin se enfoca en esas operaciones comunes con PyMuPDF.\n"
                "No valida cumplimiento formal PDF/A, PDF/X o PDF/UA."
            ),
        )
        text.configure(state="disabled")

    def _meta_row(self, parent, label_text: str, variable: tk.StringVar) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label_text).pack(side="left", padx=(0, 6))
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta destino")
        if path:
            self.output_dir_var.set(path)

    def _choose_output_file(self) -> None:
        initial = f"{self._base_name()}.pdf"
        path = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            defaultextension=".pdf",
            initialfile=initial,
            filetypes=[("PDF", "*.pdf")],
        )
        if path:
            self.output_file_var.set(path)

    def _pick_pdf(self, target_var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(title="Selecciona PDF", filetypes=[("PDF", "*.pdf")])
        if path:
            target_var.set(path)

    def _append_log(self, text: str) -> None:
        def _do() -> None:
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")

        self.root.after(0, _do)

    def _clear_log(self) -> None:
        self.log_box.delete("1.0", "end")
        self._append_log("Log reiniciado.")

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _set_busy(self, busy: bool) -> None:
        self.running.set(busy)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()
        for button in self.action_buttons:
            if busy:
                button.state(["disabled"])
            else:
                button.state(["!disabled"])

    def _start_task(self, task_name: str, fn) -> None:
        if self.running.get():
            return
        self._set_busy(True)
        self._set_status(f"{task_name} en proceso...")
        self._append_log(f"[INICIO] {task_name}")
        self.log.info("pdf_toolkit_fitz: start %s", task_name)

        def _worker() -> None:
            try:
                message = fn()
                self.root.after(0, lambda: self._on_task_success(task_name, message))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_task_error(task_name, e))
            finally:
                self.root.after(0, lambda: self._set_busy(False))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_task_success(self, task_name: str, message: str) -> None:
        self._set_status(f"{task_name} completado.")
        if message:
            self._append_log(f"[OK] {message}")
        self.log.info("pdf_toolkit_fitz: success %s", task_name)
        detail = f"{task_name} finalizado."
        if message:
            detail += f"\n\n{message}"
        messagebox.showinfo("Completado", detail)

    def _on_task_error(self, task_name: str, exc: Exception) -> None:
        self._set_status(f"{task_name} con error.")
        self._append_log(f"[ERROR] {exc}")
        self.log.exception("pdf_toolkit_fitz: error en %s", task_name)
        messagebox.showerror("Error", str(exc))

    def _output_dir(self) -> Path:
        out_dir = Path((self.output_dir_var.get() or "").strip())
        if not str(out_dir):
            raise ValueError("Debes indicar carpeta destino.")
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir

    def _base_name(self) -> str:
        return _safe_stem(self.output_name_var.get(), "resultado_pdf")

    def _single_output_pdf_path(self, suffix: str) -> Path:
        raw = (self.output_file_var.get() or "").strip()
        if raw:
            out = Path(raw)
            if out.suffix.lower() != ".pdf":
                out = out.with_suffix(".pdf")
            if out.parent:
                out.parent.mkdir(parents=True, exist_ok=True)
            return _unique_path(out)
        return _unique_path(self._output_dir() / f"{self._base_name()}{suffix}.pdf")

    def _parse_dpi(self, value: str, *, min_dpi: int = 72, max_dpi: int = 600) -> int:
        raw = (value or "").strip()
        if not raw.isdigit():
            raise ValueError("DPI invalido. Usa un numero entero.")
        dpi = int(raw)
        if dpi < min_dpi or dpi > max_dpi:
            raise ValueError(f"DPI fuera de rango. Usa un valor entre {min_dpi} y {max_dpi}.")
        return dpi

    def _refresh_listbox(self, box: tk.Listbox, items: list[Path]) -> None:
        box.delete(0, "end")
        for item in items:
            box.insert("end", str(item))

    def _add_merge_files(self) -> None:
        paths = filedialog.askopenfilenames(title="Selecciona PDFs", filetypes=[("PDF", "*.pdf")])
        for path in paths:
            p = Path(path)
            if p not in self.merge_files:
                self.merge_files.append(p)
        self._refresh_listbox(self.merge_list, self.merge_files)

    def _remove_merge_file(self) -> None:
        sel = self.merge_list.curselection()
        if not sel:
            return
        idx = sel[0]
        self.merge_files.pop(idx)
        self._refresh_listbox(self.merge_list, self.merge_files)

    def _clear_merge_files(self) -> None:
        self.merge_files.clear()
        self._refresh_listbox(self.merge_list, self.merge_files)

    def _add_images(self) -> None:
        paths = filedialog.askopenfilenames(title="Selecciona imagenes", filetypes=IMAGE_FILE_TYPES)
        for path in paths:
            p = Path(path)
            if p not in self.create_images:
                self.create_images.append(p)
        self._refresh_listbox(self.images_list, self.create_images)

    def _remove_image(self) -> None:
        sel = self.images_list.curselection()
        if not sel:
            return
        idx = sel[0]
        self.create_images.pop(idx)
        self._refresh_listbox(self.images_list, self.create_images)

    def _clear_images(self) -> None:
        self.create_images.clear()
        self._refresh_listbox(self.images_list, self.create_images)

    def _run_merge(self) -> None:
        self._start_task("Combinar PDFs", self._task_merge)

    def _run_split(self) -> None:
        self._start_task("Dividir PDF", self._task_split)

    def _run_edit(self) -> None:
        self._start_task("Editar PDF", self._task_edit)

    def _run_security(self) -> None:
        self._start_task("Seguridad PDF", self._task_security)

    def _run_organize(self) -> None:
        self._start_task("Organizar PDF", self._task_organize)

    def _run_create(self) -> None:
        self._start_task("Crear PDF", self._task_create)

    def _run_extract(self) -> None:
        self._start_task("Extraer contenido", self._task_extract)

    def _run_export_images(self) -> None:
        self._start_task("Exportar PDF a imagenes", self._task_export_images)

    def _run_template_apply(self) -> None:
        payload = self._template_build_payload(apply_template=True)
        try:
            report = self._validate_template_payload(payload, require_actions=True)
            self._append_log(f"[OK] Pre-validacion plantilla: {report}")
        except Exception as exc:
            messagebox.showerror("No se puede aplicar plantilla", str(exc))
            return
        self._start_task(
            "Aplicar plantilla PDF", lambda p=payload: self._task_template_apply(p)
        )

    def _run_template_flatten(self) -> None:
        payload = self._template_build_payload(apply_template=False)
        try:
            self._validate_template_payload(payload, require_actions=False)
        except Exception as exc:
            messagebox.showerror("No se puede digitalizar", str(exc))
            return
        self._start_task("Digitalizar PDF", lambda p=payload: self._task_template_flatten(p))

    def _validate_template_ui(self) -> None:
        try:
            payload = self._template_build_payload(apply_template=True)
            report = self._validate_template_payload(payload, require_actions=True)
            self._append_log(f"[OK] Validacion plantilla: {report}")
            self._set_status("Plantilla valida. Lista para aplicar.")
            messagebox.showinfo("Plantilla valida", report)
        except Exception as exc:
            self._append_log(f"[ERROR] Validacion plantilla: {exc}")
            self._set_status("Plantilla con errores.")
            messagebox.showerror("Plantilla invalida", str(exc))

    def _template_build_payload(self, apply_template: bool) -> dict:
        return {
            "input_pdf": self.template_file_var.get(),
            "actions": [dict(item) for item in self.template_actions],
            "values_text": self.hotspot_values_box.get("1.0", "end"),
            "flatten": bool(self.template_flatten_var.get()),
            "redact": bool(self.template_redact_var.get()),
            "editable_fields": bool(self.template_fields_var.get()),
            "optimize": bool(self.template_opt_var.get()),
            "dpi": self.template_dpi_var.get(),
            "apply_template": apply_template,
        }

    def _template_parse_zoom(self) -> float:
        raw = (self.template_zoom_var.get() or "").strip()
        try:
            pct = float(raw)
        except ValueError as exc:
            raise ValueError("Zoom invalido. Usa porcentaje numerico (ej: 125).") from exc
        if pct < 25 or pct > 400:
            raise ValueError("Zoom fuera de rango. Usa un valor entre 25 y 400.")
        return pct / 100.0

    def _template_parse_dpi(self, value: str) -> int:
        return self._parse_dpi(value, min_dpi=72, max_dpi=600)

    def _template_parse_hotspot_size(self) -> float:
        raw = (self.hotspot_size_var.get() or "").strip()
        try:
            size = float(raw)
        except ValueError as exc:
            raise ValueError("Tamano de fuente invalido para hotspot.") from exc
        if size < 4 or size > 120:
            raise ValueError("Tamano de fuente fuera de rango (4-120).")
        return size

    def _load_template_pdf(self) -> None:
        try:
            in_path = _ensure_pdf_path(self.template_file_var.get())
            self._close_template_doc()
            doc = fitz.open(in_path)
            _assert_not_encrypted(doc, in_path)
            if doc.page_count == 0:
                doc.close()
                raise ValueError("El PDF no tiene paginas.")
            self.template_pdf_doc = doc
            self.template_pdf_path = in_path
            self.template_page_index = 0
            self._render_template_preview()
            self._append_log(f"Preview cargado: {in_path}")
        except Exception as exc:
            messagebox.showerror("Error cargando preview", str(exc))

    def _close_template_doc(self) -> None:
        if self.template_pdf_doc is not None:
            try:
                self.template_pdf_doc.close()
            except Exception:
                pass
            self.template_pdf_doc = None
            self.template_pdf_path = None
            self.preview_photo = None

    def _template_prev_page(self) -> None:
        if self.template_pdf_doc is None:
            return
        if self.template_page_index > 0:
            self.template_page_index -= 1
            self._render_template_preview()

    def _template_next_page(self) -> None:
        if self.template_pdf_doc is None:
            return
        if self.template_page_index < self.template_pdf_doc.page_count - 1:
            self.template_page_index += 1
            self._render_template_preview()

    def _render_template_preview(self) -> None:
        canvas = self.template_canvas
        canvas.delete("all")
        self.preview_photo = None
        self.preview_image_size = (0.0, 0.0)
        self.preview_action_indexes = []

        if self.template_pdf_doc is None:
            self.template_page_var.set("Pagina 0/0")
            self._template_refresh_actions()
            canvas.configure(scrollregion=(0, 0, 10, 10))
            return

        total = self.template_pdf_doc.page_count
        self.template_page_index = max(0, min(self.template_page_index, total - 1))
        self.template_page_var.set(f"Pagina {self.template_page_index + 1}/{total}")

        try:
            zoom = self._template_parse_zoom()
        except Exception as exc:
            messagebox.showerror("Zoom invalido", str(exc))
            return
        page = self.template_pdf_doc[self.template_page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        png_data = pix.tobytes("png")
        encoded = base64.b64encode(png_data).decode("ascii")
        self.preview_photo = tk.PhotoImage(data=encoded)

        ox, oy = self.preview_origin
        width = float(self.preview_photo.width())
        height = float(self.preview_photo.height())
        self.preview_image_size = (width, height)
        canvas.create_image(ox, oy, image=self.preview_photo, anchor="nw")
        canvas.create_rectangle(ox, oy, ox + width, oy + height, outline="#6b7280", width=1)
        self._template_draw_overlays()
        canvas.configure(
            scrollregion=(0, 0, max(24, int(ox + width + 24)), max(24, int(oy + height + 24)))
        )
        self._template_refresh_actions()

    def _template_draw_overlays(self) -> None:
        if self.template_pdf_doc is None:
            return
        zoom = self._template_parse_zoom()
        ox, oy = self.preview_origin
        for action in self.template_actions:
            page = int(action.get("page", -1))
            if page != self.template_page_index:
                continue
            rect = fitz.Rect(action["rect"])
            x0 = ox + rect.x0 * zoom
            y0 = oy + rect.y0 * zoom
            x1 = ox + rect.x1 * zoom
            y1 = oy + rect.y1 * zoom
            if action.get("type") == "white":
                outline = "#38bdf8"
                label = "WHITE"
            else:
                outline = "#34d399"
                label = f"HOTSPOT:{action.get('name', '')}"
            self.template_canvas.create_rectangle(
                x0, y0, x1, y1, outline=outline, width=2, dash=(5, 3)
            )
            self.template_canvas.create_text(
                x0 + 4,
                y0 + 4,
                anchor="nw",
                text=label,
                fill=outline,
                font=("Segoe UI", 8, "bold"),
            )

    def _template_canvas_press(self, event) -> None:
        if self.template_pdf_doc is None:
            return
        if self.template_mode_var.get() not in {"white", "hotspot"}:
            return
        cx = float(self.template_canvas.canvasx(event.x))
        cy = float(self.template_canvas.canvasy(event.y))
        ox, oy = self.preview_origin
        width, height = self.preview_image_size
        if width <= 0 or height <= 0:
            return
        max_x = ox + width
        max_y = oy + height
        if cx < ox or cx > max_x or cy < oy or cy > max_y:
            return
        self.preview_drag_start = (max(ox, min(cx, max_x)), max(oy, min(cy, max_y)))
        self.preview_drag_rect_id = self.template_canvas.create_rectangle(
            cx, cy, cx, cy, outline="#f8fafc", width=2, dash=(4, 2)
        )

    def _template_canvas_drag(self, event) -> None:
        if self.preview_drag_start is None or self.preview_drag_rect_id is None:
            return
        ox, oy = self.preview_origin
        width, height = self.preview_image_size
        max_x = ox + width
        max_y = oy + height
        cx = float(self.template_canvas.canvasx(event.x))
        cy = float(self.template_canvas.canvasy(event.y))
        x = max(ox, min(cx, max_x))
        y = max(oy, min(cy, max_y))
        x0, y0 = self.preview_drag_start
        self.template_canvas.coords(self.preview_drag_rect_id, x0, y0, x, y)

    def _template_canvas_release(self, event) -> None:
        if self.preview_drag_start is None or self.preview_drag_rect_id is None:
            return
        drag_id = self.preview_drag_rect_id
        self.preview_drag_rect_id = None
        self.template_canvas.delete(drag_id)

        x0, y0 = self.preview_drag_start
        self.preview_drag_start = None
        ox, oy = self.preview_origin
        width, height = self.preview_image_size
        max_x = ox + width
        max_y = oy + height
        x = max(ox, min(float(self.template_canvas.canvasx(event.x)), max_x))
        y = max(oy, min(float(self.template_canvas.canvasy(event.y)), max_y))
        left, right = sorted([x0, x])
        top, bottom = sorted([y0, y])
        if right - left < 6 or bottom - top < 6:
            return

        try:
            pdf_rect = self._template_canvas_to_pdf_rect(left, top, right, bottom)
            self._template_add_action(pdf_rect)
            self._render_template_preview()
        except Exception as exc:
            messagebox.showerror("Error creando accion", str(exc))

    def _template_canvas_to_pdf_rect(
        self, x0: float, y0: float, x1: float, y1: float
    ) -> fitz.Rect:
        zoom = self._template_parse_zoom()
        ox, oy = self.preview_origin
        return fitz.Rect((x0 - ox) / zoom, (y0 - oy) / zoom, (x1 - ox) / zoom, (y1 - oy) / zoom)

    def _template_add_action(self, rect: fitz.Rect) -> None:
        mode = self.template_mode_var.get()
        normalized = [round(rect.x0, 2), round(rect.y0, 2), round(rect.x1, 2), round(rect.y1, 2)]
        if mode == "white":
            action = {"type": "white", "page": self.template_page_index, "rect": normalized}
            self.template_actions.append(action)
            self._append_log(
                f"Accion WHITE agregada en pagina {self.template_page_index + 1}: {normalized}"
            )
            return

        if mode == "hotspot":
            name = _safe_stem(
                self.hotspot_name_var.get(),
                f"campo_{len([a for a in self.template_actions if a.get('type') == 'hotspot']) + 1}",
            )
            size = self._template_parse_hotspot_size()
            color = self.hotspot_color_var.get().strip() or "#111111"
            _parse_hex_color(color)
            font = (self.hotspot_font_var.get() or "").strip()
            if font not in PDF_FONT_OPTIONS:
                raise ValueError(f"Fuente no soportada. Usa una de: {', '.join(PDF_FONT_OPTIONS)}")
            action = {
                "type": "hotspot",
                "page": self.template_page_index,
                "rect": normalized,
                "name": name,
                "font": font,
                "size": round(size, 2),
                "color": color,
                "default": self.hotspot_default_var.get(),
            }
            self.template_actions.append(action)
            self.hotspot_name_var.set(f"{name}_next")
            self._append_log(
                f"Hotspot '{name}' agregado en pagina {self.template_page_index + 1}: {normalized}"
            )

    def _template_refresh_actions(self) -> None:
        self.template_actions_list.delete(0, "end")
        self.preview_action_indexes = []
        for idx, action in enumerate(self.template_actions):
            page = int(action.get("page", -1))
            if page != self.template_page_index:
                continue
            self.preview_action_indexes.append(idx)
            rect = action.get("rect", [0, 0, 0, 0])
            if action.get("type") == "white":
                label = f"WHITE [{rect[0]:.0f},{rect[1]:.0f},{rect[2]:.0f},{rect[3]:.0f}]"
            else:
                label = (
                    f"HOTSPOT {action.get('name', '')} "
                    f"[{rect[0]:.0f},{rect[1]:.0f},{rect[2]:.0f},{rect[3]:.0f}]"
                )
            self.template_actions_list.insert("end", label)

    def _remove_template_action(self) -> None:
        selected = self.template_actions_list.curselection()
        if not selected:
            return
        list_idx = selected[0]
        if list_idx < 0 or list_idx >= len(self.preview_action_indexes):
            return
        action_idx = self.preview_action_indexes[list_idx]
        removed = self.template_actions.pop(action_idx)
        self._append_log(f"Accion eliminada: {removed.get('type')} pagina {removed.get('page', 0) + 1}")
        self._render_template_preview()

    def _undo_template_action(self) -> None:
        for idx in range(len(self.template_actions) - 1, -1, -1):
            if int(self.template_actions[idx].get("page", -1)) == self.template_page_index:
                removed = self.template_actions.pop(idx)
                self._append_log(
                    f"Accion removida (deshacer): {removed.get('type')} pagina {self.template_page_index + 1}"
                )
                self._render_template_preview()
                return

    def _clear_template_page(self) -> None:
        keep: list[dict] = []
        removed = 0
        for action in self.template_actions:
            if int(action.get("page", -1)) == self.template_page_index:
                removed += 1
            else:
                keep.append(action)
        if removed == 0:
            return
        self.template_actions = keep
        self._append_log(f"Acciones eliminadas en pagina {self.template_page_index + 1}: {removed}")
        self._render_template_preview()

    def _save_template_file(self) -> None:
        initial_name = "pdf_template.json"
        if self.template_pdf_path is not None:
            initial_name = f"{_safe_stem(self.template_pdf_path.stem, 'pdf')}_template.json"
        out_path = filedialog.asksaveasfilename(
            title="Guardar plantilla JSON",
            defaultextension=".json",
            initialfile=initial_name,
            filetypes=[("JSON", "*.json")],
        )
        if not out_path:
            return
        data = {
            "version": 1,
            "source_pdf": str(self.template_pdf_path) if self.template_pdf_path else "",
            "actions": self.template_actions,
        }
        Path(out_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self._append_log(f"Plantilla guardada: {out_path}")

    def _load_template_file(self) -> None:
        in_path = filedialog.askopenfilename(
            title="Cargar plantilla JSON", filetypes=[("JSON", "*.json")]
        )
        if not in_path:
            return
        try:
            payload = json.loads(Path(in_path).read_text(encoding="utf-8"))
            raw_actions = payload.get("actions")
            if not isinstance(raw_actions, list):
                raise ValueError("Plantilla invalida: falta lista de acciones.")
            parsed: list[dict] = []
            for item in raw_actions:
                action = self._normalize_template_action(item)
                if action is not None:
                    parsed.append(action)
            if not parsed:
                raise ValueError("La plantilla no contiene acciones validas.")
            self.template_actions = parsed

            source_pdf = str(payload.get("source_pdf", "")).strip()
            if source_pdf and Path(source_pdf).exists():
                self.template_file_var.set(source_pdf)
                self._load_template_pdf()
            else:
                self._render_template_preview()
            self._template_seed_values()
            self._set_status("Plantilla cargada. Valida y aplica para guardar el PDF final.")
            self._append_log(f"Plantilla cargada: {in_path} ({len(parsed)} acciones)")
        except Exception as exc:
            messagebox.showerror("Error cargando plantilla", str(exc))

    def _normalize_template_action(self, item: object) -> dict | None:
        if not isinstance(item, dict):
            return None
        action_type = str(item.get("type", "")).strip().lower()
        if action_type not in {"white", "hotspot"}:
            return None

        page = item.get("page")
        if not isinstance(page, int) or page < 0:
            return None

        rect = item.get("rect")
        if not isinstance(rect, list) or len(rect) != 4:
            return None
        try:
            x0, y0, x1, y1 = [float(v) for v in rect]
        except (TypeError, ValueError):
            return None
        if x1 <= x0 or y1 <= y0:
            return None

        if action_type == "white":
            return {"type": "white", "page": page, "rect": [x0, y0, x1, y1]}

        name = _safe_stem(str(item.get("name", "")), "campo")
        font = str(item.get("font", "helv")).strip()
        if font not in PDF_FONT_OPTIONS:
            font = "helv"
        try:
            size = float(item.get("size", 12))
        except (TypeError, ValueError):
            size = 12.0
        size = max(4.0, min(size, 120.0))
        color = str(item.get("color", "#111111")).strip() or "#111111"
        _parse_hex_color(color)
        default = str(item.get("default", ""))
        return {
            "type": "hotspot",
            "page": page,
            "rect": [x0, y0, x1, y1],
            "name": name,
            "font": font,
            "size": size,
            "color": color,
            "default": default,
        }

    def _template_seed_values(self) -> None:
        current = self._parse_hotspot_values(self.hotspot_values_box.get("1.0", "end"))
        names = sorted(
            {
                _safe_stem(str(action.get("name", "")), "")
                for action in self.template_actions
                if action.get("type") == "hotspot"
            }
        )
        lines: list[str] = []
        for name in names:
            if not name:
                continue
            lines.append(f"{name}={current.get(name, '')}")
        self.hotspot_values_box.delete("1.0", "end")
        if lines:
            self.hotspot_values_box.insert("1.0", "\n".join(lines) + "\n")

    def _parse_hotspot_values(self, text: str) -> dict[str, str]:
        result: dict[str, str] = {}
        for raw_line in (text or "").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            field = _safe_stem(key, "")
            if field:
                result[field] = value.strip()
        return result

    def _validate_template_payload(self, payload: dict, require_actions: bool) -> str:
        in_path = _ensure_pdf_path(payload.get("input_pdf", ""))
        actions = payload.get("actions") or []
        if not isinstance(actions, list):
            raise ValueError("Formato invalido: acciones de plantilla no es una lista.")

        normalized: list[dict] = []
        for item in actions:
            action = self._normalize_template_action(item)
            if action is not None:
                normalized.append(action)
        if require_actions and not normalized:
            raise ValueError("No hay acciones validas para aplicar.")

        values = self._parse_hotspot_values(str(payload.get("values_text", "")))
        dpi = self._template_parse_dpi(str(payload.get("dpi", "144")))
        flatten = bool(payload.get("flatten", False))
        redact = bool(payload.get("redact", False))
        editable_fields = bool(payload.get("editable_fields", True))

        with fitz.open(in_path) as doc:
            _assert_not_encrypted(doc, in_path)
            if doc.page_count == 0:
                raise ValueError("El PDF no tiene paginas.")
            total_pages = doc.page_count

            white_count = 0
            hotspot_count = 0
            skipped_page = 0
            skipped_rect = 0
            missing_without_default: list[str] = []

            for action in normalized:
                page_idx = int(action["page"])
                if page_idx < 0 or page_idx >= doc.page_count:
                    skipped_page += 1
                    continue
                page = doc[page_idx]
                rect = fitz.Rect(action["rect"]) & page.rect
                if rect.is_empty:
                    skipped_rect += 1
                    continue
                if action["type"] == "white":
                    white_count += 1
                else:
                    hotspot_count += 1
                    name = _safe_stem(str(action.get("name", "")), "")
                    if name:
                        default_val = str(action.get("default", "")).strip()
                        user_val = values.get(name, "").strip()
                        if not default_val and not user_val:
                            missing_without_default.append(name)

        applied_candidate = white_count + hotspot_count
        if require_actions and applied_candidate == 0:
            raise ValueError(
                "Las acciones de la plantilla no aplican al PDF actual "
                "(paginas fuera de rango o rectangulos fuera de pagina)."
            )

        missing_fields = sorted(set(missing_without_default))
        missing_label = ", ".join(missing_fields[:8]) if missing_fields else "ninguno"
        if len(missing_fields) > 8:
            missing_label += ", ..."

        mode = "editable (AcroForm)" if editable_fields else "texto fijo"
        return (
            f"PDF: {total_pages} pagina(s) | "
            f"acciones validas: {applied_candidate} (white={white_count}, hotspot={hotspot_count}) | "
            f"omitidas: pagina={skipped_page}, rect={skipped_rect} | "
            f"faltan valores (sin default): {missing_label} | "
            f"modo hotspots: {mode} | "
            f"redaccion={redact} | flatten={flatten} (dpi={dpi})"
        )

    def _clone_document(self, src: fitz.Document) -> fitz.Document:
        out = fitz.open()
        out.insert_pdf(src)
        return out

    def _flatten_document(self, src: fitz.Document, dpi: int) -> fitz.Document:
        out = fitz.open()
        for pno in range(src.page_count):
            page = src[pno]
            pix = page.get_pixmap(dpi=dpi, alpha=False, annots=True)
            new_page = out.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, stream=pix.tobytes("png"))
        return out

    def _widget_font_name(self, name: str) -> str:
        normalized = (name or "").strip().lower()
        mapping = {
            "helv": "Helv",
            "cour": "Cour",
            "tiro": "TiRo",
            "symb": "Symb",
            "zapfdingbats": "ZaDb",
            "zadb": "ZaDb",
        }
        return mapping.get(normalized, "Helv")

    def _task_merge(self) -> str:
        if len(self.merge_files) < 2:
            raise ValueError("Debes seleccionar al menos 2 PDFs.")

        out = self._single_output_pdf_path("_merge")
        out_doc = fitz.open()
        try:
            for path in self.merge_files:
                with fitz.open(path) as src:
                    _assert_not_encrypted(src, path)
                    out_doc.insert_pdf(src)
            _save_doc(out_doc, out, optimize=self.merge_opt_var.get())
        finally:
            out_doc.close()
        return f"PDF combinado: {out}"

    def _task_split(self) -> str:
        in_path = _ensure_pdf_path(self.split_file_var.get())
        out_dir = self._output_dir() / f"{self._base_name()}_split"
        out_dir.mkdir(parents=True, exist_ok=True)

        with fitz.open(in_path) as src:
            _assert_not_encrypted(src, in_path)
            total = src.page_count
            if total == 0:
                raise ValueError("El PDF no tiene paginas.")

            if self.split_mode_var.get() == "each":
                groups = [(f"p{i + 1}", [i]) for i in range(total)]
            else:
                groups = _parse_split_groups(self.split_ranges_var.get(), total)

            generated = 0
            for label, pages in groups:
                part = fitz.open()
                try:
                    for pno in pages:
                        part.insert_pdf(src, from_page=pno, to_page=pno)
                    out_file = _unique_path(
                        out_dir / f"{_safe_stem(in_path.stem, 'pdf')}_{label}.pdf"
                    )
                    _save_doc(part, out_file, optimize=self.split_opt_var.get())
                    generated += 1
                finally:
                    part.close()

        return f"Division completada. Archivos generados: {generated} en {out_dir}"

    def _task_edit(self) -> str:
        in_path = _ensure_pdf_path(self.edit_file_var.get())
        out = self._single_output_pdf_path("_editado")

        with fitz.open(in_path) as doc:
            _assert_not_encrypted(doc, in_path)
            total = doc.page_count
            pages = _parse_page_spec(self.edit_pages_var.get(), total)

            changed = False
            rotate = int((self.edit_rotate_var.get() or "0").strip())
            if rotate not in {0, 90, 180, 270}:
                raise ValueError("Rotacion invalida. Usa 0, 90, 180 o 270.")
            if rotate:
                for pno in pages:
                    page = doc[pno]
                    page.set_rotation((page.rotation + rotate) % 360)
                changed = True

            watermark = self.edit_watermark_var.get().strip()
            if watermark:
                for pno in pages:
                    page = doc[pno]
                    rect = page.rect
                    box = fitz.Rect(
                        rect.x0 + 30,
                        rect.y0 + rect.height * 0.45,
                        rect.x1 - 30,
                        rect.y0 + rect.height * 0.55,
                    )
                    page.insert_textbox(
                        box,
                        watermark,
                        fontsize=24,
                        color=(0.65, 0.65, 0.65),
                        align=fitz.TEXT_ALIGN_CENTER,
                        overlay=True,
                    )
                changed = True

            metadata_changes = {
                "title": self.meta_title_var.get().strip(),
                "author": self.meta_author_var.get().strip(),
                "subject": self.meta_subject_var.get().strip(),
                "keywords": self.meta_keywords_var.get().strip(),
            }
            if any(metadata_changes.values()):
                metadata = dict(doc.metadata or {})
                for key, val in metadata_changes.items():
                    if val:
                        metadata[key] = val
                doc.set_metadata(metadata)
                changed = True

            optimize = self.edit_opt_var.get()
            if not changed and not optimize:
                raise ValueError("No hay cambios para aplicar.")

            _save_doc(doc, out, optimize=optimize)

        return f"PDF editado: {out}"

    def _task_security(self) -> str:
        in_path = _ensure_pdf_path(self.security_file_var.get())
        mode = (self.security_mode_var.get() or "").strip().lower()
        optimize = bool(self.security_opt_var.get())
        algo_name = (self.security_algo_var.get() or "AES-256").strip().upper()
        algo_map = {
            "AES-256": fitz.PDF_ENCRYPT_AES_256,
            "AES-128": fitz.PDF_ENCRYPT_AES_128,
            "RC4-128": fitz.PDF_ENCRYPT_RC4_128,
        }
        if algo_name not in algo_map:
            raise ValueError("Algoritmo invalido. Usa AES-256, AES-128 o RC4-128.")

        with fitz.open(in_path) as doc:
            if getattr(doc, "needs_pass", False):
                current_pw = self.security_current_pw_var.get()
                if not current_pw:
                    raise ValueError(
                        "El PDF esta protegido. Debes indicar contrasena actual para continuar."
                    )
                if doc.authenticate(current_pw) <= 0:
                    raise ValueError("Contrasena actual invalida.")

            if mode == "decrypt":
                out = self._single_output_pdf_path("_sin_password")
                doc.save(
                    out,
                    encryption=fitz.PDF_ENCRYPT_NONE,
                    garbage=4 if optimize else 0,
                    deflate=bool(optimize),
                )
                return f"Proteccion removida: {out}"

            if mode != "encrypt":
                raise ValueError("Modo de seguridad invalido.")

            user_pw = self.security_user_pw_var.get()
            owner_pw = self.security_owner_pw_var.get()
            if not user_pw and not owner_pw:
                raise ValueError("Debes indicar al menos una contrasena (usuario u owner).")
            if not owner_pw:
                owner_pw = user_pw

            permissions = (
                fitz.PDF_PERM_ACCESSIBILITY
                | fitz.PDF_PERM_PRINT
                | fitz.PDF_PERM_COPY
                | fitz.PDF_PERM_ANNOTATE
                | fitz.PDF_PERM_FORM
                | fitz.PDF_PERM_ASSEMBLE
            )

            out = self._single_output_pdf_path("_protegido")
            doc.save(
                out,
                encryption=algo_map[algo_name],
                user_pw=user_pw,
                owner_pw=owner_pw,
                permissions=permissions,
                garbage=4 if optimize else 0,
                deflate=bool(optimize),
            )
            return f"PDF protegido ({algo_name}): {out}"

    def _task_organize(self) -> str:
        in_path = _ensure_pdf_path(self.organize_file_var.get())
        out = self._single_output_pdf_path("_organizado")
        spec = self.organize_order_var.get()

        with fitz.open(in_path) as doc:
            _assert_not_encrypted(doc, in_path)
            if doc.page_count == 0:
                raise ValueError("El PDF no tiene paginas.")
            order = _parse_ordered_page_spec(spec, doc.page_count)
            doc.select(order)
            _save_doc(doc, out, optimize=self.organize_opt_var.get())

        return f"PDF organizado: {out} | paginas de salida: {len(order)}"

    def _task_create(self) -> str:
        if not self.create_images:
            raise ValueError("Debes seleccionar al menos una imagen.")

        out = self._single_output_pdf_path("_imagenes")
        doc = fitz.open()
        try:
            for img_path in self.create_images:
                with fitz.open(img_path) as img_doc:
                    pdf_bytes = img_doc.convert_to_pdf()
                with fitz.open("pdf", pdf_bytes) as img_pdf:
                    doc.insert_pdf(img_pdf)
            _save_doc(doc, out, optimize=self.create_opt_var.get())
        finally:
            doc.close()

        return f"PDF creado desde imagenes: {out}"

    def _task_extract(self) -> str:
        in_path = _ensure_pdf_path(self.extract_file_var.get())
        extract_text = self.extract_text_var.get()
        extract_images = self.extract_images_var.get()
        if not extract_text and not extract_images:
            raise ValueError("Selecciona al menos una opcion de extraccion.")

        out_dir = self._output_dir() / f"{self._base_name()}_extraido"
        out_dir.mkdir(parents=True, exist_ok=True)

        generated: list[Path] = []
        image_count = 0
        with fitz.open(in_path) as doc:
            _assert_not_encrypted(doc, in_path)
            pages = _parse_page_spec(self.extract_pages_var.get(), doc.page_count)

            if extract_text:
                txt_path = _unique_path(out_dir / f"{_safe_stem(in_path.stem, 'pdf')}_texto.txt")
                chunks: list[str] = []
                for pno in pages:
                    chunks.append(f"\n--- Pagina {pno + 1} ---\n")
                    chunks.append(doc[pno].get_text("text"))
                txt_path.write_text("".join(chunks), encoding="utf-8")
                generated.append(txt_path)

            if extract_images:
                image_count = self._extract_images(doc, pages, out_dir, in_path.stem)

        parts: list[str] = []
        if generated:
            parts.append("Archivos: " + ", ".join(str(p) for p in generated))
        if extract_images:
            parts.append(f"Imagenes extraidas: {image_count}")
        return " | ".join(parts) if parts else f"Extraccion completada en {out_dir}"

    def _task_export_images(self) -> str:
        in_path = _ensure_pdf_path(self.export_img_file_var.get())
        fmt = (self.export_img_format_var.get() or "").strip().lower()
        if fmt not in EXPORT_IMAGE_FORMATS:
            raise ValueError(f"Formato invalido. Usa: {', '.join(EXPORT_IMAGE_FORMATS)}.")
        dpi = self._parse_dpi(self.export_img_dpi_var.get(), min_dpi=72, max_dpi=1200)
        grayscale = bool(self.export_img_gray_var.get())

        out_dir = self._output_dir() / f"{self._base_name()}_imagenes"
        out_dir.mkdir(parents=True, exist_ok=True)
        base = _safe_stem(in_path.stem, "pdf")
        exported = 0

        with fitz.open(in_path) as doc:
            _assert_not_encrypted(doc, in_path)
            pages = _parse_page_spec(self.export_img_pages_var.get(), doc.page_count)
            colorspace = fitz.csGRAY if grayscale else fitz.csRGB
            for pno in pages:
                page = doc[pno]
                pix = page.get_pixmap(dpi=dpi, colorspace=colorspace, alpha=False)
                out_file = _unique_path(out_dir / f"{base}_p{pno + 1}.{fmt}")
                if fmt == "jpg":
                    out_file.write_bytes(pix.tobytes("jpg"))
                else:
                    out_file.write_bytes(pix.tobytes("png"))
                exported += 1

        return f"Imagenes exportadas: {exported} en {out_dir}"

    def _task_template_apply(self, payload: dict) -> str:
        in_path = _ensure_pdf_path(payload.get("input_pdf", ""))
        actions = payload.get("actions") or []
        if not isinstance(actions, list) or not actions:
            raise ValueError("No hay acciones de plantilla para aplicar.")
        normalized_actions: list[dict] = []
        for item in actions:
            action = self._normalize_template_action(item)
            if action is not None:
                normalized_actions.append(action)
        if not normalized_actions:
            raise ValueError("No se encontraron acciones validas en la plantilla.")

        values = self._parse_hotspot_values(str(payload.get("values_text", "")))
        flatten = bool(payload.get("flatten", False))
        redact = bool(payload.get("redact", False))
        editable_fields = bool(payload.get("editable_fields", True))
        optimize = bool(payload.get("optimize", True))
        dpi = self._template_parse_dpi(str(payload.get("dpi", "144")))

        out = self._single_output_pdf_path("_plantilla")

        with fitz.open(in_path) as src:
            _assert_not_encrypted(src, in_path)
            doc = self._flatten_document(src, dpi) if flatten else self._clone_document(src)

        try:
            white_count = 0
            hotspot_count = 0
            redact_pages: set[int] = set()

            for action in normalized_actions:
                page_idx = int(action["page"])
                if page_idx < 0 or page_idx >= doc.page_count:
                    continue
                page = doc[page_idx]
                if action["type"] == "white":
                    rect = fitz.Rect(action["rect"]) & page.rect
                    if rect.is_empty:
                        continue
                    if redact:
                        page.add_redact_annot(rect, fill=(1, 1, 1))
                        redact_pages.add(page_idx)
                    else:
                        page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), width=0, overlay=True)
                    white_count += 1

            if redact_pages:
                for pno in sorted(redact_pages):
                    doc[pno].apply_redactions()

            field_name_counter: dict[str, int] = {}
            for action in normalized_actions:
                if action["type"] != "hotspot":
                    continue
                page_idx = int(action["page"])
                if page_idx < 0 or page_idx >= doc.page_count:
                    continue
                page = doc[page_idx]
                rect = fitz.Rect(action["rect"]) & page.rect
                if rect.is_empty:
                    continue
                name = _safe_stem(str(action.get("name", "")), "")
                value = values.get(name, str(action.get("default", "")))
                if editable_fields:
                    base_name = name or "campo"
                    idx = field_name_counter.get(base_name, 0) + 1
                    field_name_counter[base_name] = idx
                    field_name = base_name if idx == 1 else f"{base_name}_{idx}"
                    widget = fitz.Widget()
                    widget.field_name = field_name
                    widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
                    widget.rect = rect
                    widget.field_value = value
                    widget.text_font = self._widget_font_name(str(action.get("font", "helv")))
                    widget.text_fontsize = float(action.get("size", 12))
                    widget.text_color = _parse_hex_color(str(action.get("color", "#111111")))
                    widget.fill_color = (1, 1, 1)
                    widget.border_color = (0.78, 0.78, 0.78)
                    widget.border_width = 1
                    page.add_widget(widget)
                else:
                    page.insert_textbox(
                        rect,
                        value,
                        fontsize=float(action.get("size", 12)),
                        fontname=str(action.get("font", "helv")),
                        color=_parse_hex_color(str(action.get("color", "#111111"))),
                        align=fitz.TEXT_ALIGN_LEFT,
                        overlay=True,
                    )
                hotspot_count += 1

            if white_count == 0 and hotspot_count == 0:
                raise ValueError(
                    "Ninguna accion fue aplicable al PDF destino. "
                    "Revisa que la plantilla corresponda al documento cargado."
                )

            if editable_fields:
                try:
                    doc.need_appearances(True)
                except Exception:
                    pass
            _save_doc(doc, out, optimize=optimize)
        finally:
            doc.close()

        mode_txt = "con digitalizacion previa" if flatten else "sin digitalizar"
        field_mode_txt = "campos editables" if editable_fields else "texto fijo"
        return (
            f"Plantilla aplicada ({mode_txt}): {out} | "
            f"cuadros: {white_count}, hotspots: {hotspot_count} ({field_mode_txt})"
        )

    def _task_template_flatten(self, payload: dict) -> str:
        in_path = _ensure_pdf_path(payload.get("input_pdf", ""))
        optimize = bool(payload.get("optimize", True))
        dpi = self._template_parse_dpi(str(payload.get("dpi", "144")))
        out = self._single_output_pdf_path("_digitalizado")

        with fitz.open(in_path) as src:
            _assert_not_encrypted(src, in_path)
            doc = self._flatten_document(src, dpi)
        try:
            _save_doc(doc, out, optimize=optimize)
        finally:
            doc.close()
        return f"PDF digitalizado: {out}"

    def _extract_images(
        self, doc: fitz.Document, pages: list[int], out_dir: Path, prefix: str
    ) -> int:
        seen_xrefs: set[int] = set()
        count = 0
        base = _safe_stem(prefix, "pdf")
        for pno in pages:
            page = doc[pno]
            for img in page.get_images(full=True):
                xref = img[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)
                data = doc.extract_image(xref)
                ext = data.get("ext", "bin")
                payload = data.get("image")
                if not payload:
                    continue
                out_file = _unique_path(out_dir / f"{base}_p{pno + 1}_{xref}.{ext}")
                out_file.write_bytes(payload)
                count += 1
        return count


def crear_interfaz(parent, context):
    ui = PDFToolkitUI(parent, context)

    def on_destroy(_evt=None):
        if _evt is not None and _evt.widget is not ui.root:
            return
        ui._close_template_doc()

    ui.root.bind("<Destroy>", on_destroy)
    return ui.root
