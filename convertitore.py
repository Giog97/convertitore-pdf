# -*- coding: utf-8 -*-
"""
Convertitore PDF
----------------
Applicazione desktop con interfaccia grafica (due schede) per:
  * Convertire immagini e documenti Word in PDF (con anteprima del PDF risultante,
    unione, formato pagina)
  * Dividere un PDF in piu' file (con anteprima delle pagine)

Immagini -> Pillow | Word -> Microsoft Word (automazione COM)
Dividi/Unisci PDF -> pypdf | Anteprima PDF -> pypdfium2
Interfaccia disponibile in Italiano e Inglese.
"""

import os
import sys
import queue
import shutil
import tempfile
import threading
import traceback
from pathlib import Path
from contextlib import nullcontext


def _fix_tcl_tk():
    """Corregge il caricamento di Tcl/Tk negli ambienti virtuali Windows."""
    if sys.platform != "win32":
        return
    if os.environ.get("TCL_LIBRARY") and os.environ.get("TK_LIBRARY"):
        return
    base = getattr(sys, "base_prefix", sys.prefix)
    tcl_root = os.path.join(base, "tcl")
    if not os.path.isdir(tcl_root):
        return
    for name in os.listdir(tcl_root):
        full = os.path.join(tcl_root, name)
        if not os.path.isdir(full):
            continue
        if name.startswith("tcl8") and os.path.isfile(os.path.join(full, "init.tcl")):
            os.environ.setdefault("TCL_LIBRARY", full)
        elif name.startswith("tk8"):
            os.environ.setdefault("TK_LIBRARY", full)


_fix_tcl_tk()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
    _BaseTk = TkinterDnD.Tk
except Exception:
    DND_AVAILABLE = False
    _BaseTk = tk.Tk

try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except Exception:
    PYPDF_AVAILABLE = False

try:
    import pypdfium2 as pdfium
    PDFIUM_AVAILABLE = True
except Exception:
    PDFIUM_AVAILABLE = False


# ---------------------------------------------------------------------------
# Palette e costanti
# ---------------------------------------------------------------------------
BG      = "#EEF1F6"
CARD    = "#FFFFFF"
ACCENT  = "#4C6EF5"
ACCENT2 = "#3B5BDB"
TEXT    = "#2B2F36"
MUTED   = "#5B6470"
BORDER  = "#E1E5EC"
FIELD   = "#F6F7FA"

IMAGE_EXTS = {
    ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff",
    ".webp", ".ico", ".ppm", ".pgm", ".pnm", ".jfif",
}
WORD_EXTS = {".doc", ".docx", ".docm", ".rtf", ".odt"}
SUPPORTED_EXTS = IMAGE_EXTS | WORD_EXTS

PAGE_DPI = 150
PAGE_SIZES = {
    "fit":    None,
    "a4":     (1240, 1754),
    "letter": (1275, 1650),
}


# ---------------------------------------------------------------------------
# Traduzioni
# ---------------------------------------------------------------------------
TR = {
    "it": {
        "subtitle": "Immagini e Word → PDF  ·  Dividi PDF",
        "tab_convert": "Converti in PDF",
        "tab_split": "Dividi PDF",
        # --- Converti ---
        "c_files_title": "1 · File da convertire",
        "c_hint_dnd": "Trascina qui i file o le cartelle, oppure usa \"Aggiungi file\"",
        "c_hint_nodnd": "Usa \"Aggiungi file\" per selezionare i file",
        "c_add": "＋ Aggiungi file",
        "c_remove": "Rimuovi",
        "c_clear": "Svuota",
        "c_preview_title": "Anteprima del PDF risultante",
        "c_preview_select": "Seleziona un file\nper vedere l'anteprima",
        "c_preview_generating": "Generazione anteprima…\n(apertura di Word)",
        "c_opts_title": "2 · Opzioni",
        "c_save_to": "Salva in:",
        "c_out_default": "Stessa cartella dei file di origine",
        "c_choose": "Scegli…",
        "c_reset": "Ripristina",
        "c_merge": "Unisci tutti i file in un unico PDF (nell'ordine dell'elenco)",
        "c_page_format": "Formato pagina immagini:",
        "c_fit": "Adatta all'immagine",
        "c_a4": "A4",
        "c_letter": "Lettera",
        "c_convert_btn": "Converti in PDF",
        "open_folder": "📂 Apri cartella",
        "c_ready": "Pronto. Aggiungi i file e premi \"Converti in PDF\".",
        "c_info_img": "{name}\n{w}×{h} px · {size}",
        "c_info_word": "{name}\nPDF da Word · {size}",
        "m_added": "Aggiunti {n} file.",
        "m_ignored": "Ignorati {n} file con formato non supportato.",
        "m_no_files_title": "Nessun file",
        "m_no_files_msg": "Aggiungi almeno un file da convertire.",
        "m_pillow_missing": "Pillow non installato: impossibile convertire.",
        "m_start_word": "Avvio di Microsoft Word…",
        "m_ok_named": "OK  {name} → {out}",
        "m_ok_simple": "OK  {name}",
        "m_error_file": "ERRORE {name}: {e}",
        "m_merge_created": "→ PDF unico creato: {name}",
        "m_generic_error": "ERRORE: {e}",
        "m_word_hint": "Verifica che Microsoft Word sia installato per i documenti Word.",
        "m_completed_errors": "Completato con {n} errore/i.",
        "m_completed_ok": "Conversione completata con successo.",
        "t_completed_title": "Completato",
        "t_completed_warn": "Conversione terminata con {n} errore/i.\nControlla il log.",
        "t_done_title": "Fatto",
        "t_all_converted": "Tutti i file sono stati convertiti in PDF.",
        "err_title": "Errore",
        # --- Dividi ---
        "s_file_title": "1 · Scegli il PDF da dividere",
        "s_pdf_hint_dnd": "Trascina qui un PDF oppure premi \"Scegli PDF\"",
        "s_pdf_hint_nodnd": "Premi \"Scegli PDF\"",
        "s_choose_pdf": "Scegli PDF…",
        "s_pages_count": "📄 {name} · {n} pagine",
        "s_how_title": "2 · Come dividere",
        "s_mode_each": "Una pagina per ogni file PDF",
        "s_every_1": "Ogni",
        "s_every_2": "pagine per file",
        "s_extract": "Estrai pagine:",
        "s_extract_hint": "es. 1-3, 5, 8",
        "s_out_default": "Stessa cartella del PDF",
        "s_preview_title": "Anteprima pagine",
        "s_preview_select": "Seleziona un PDF\nper vedere le pagine",
        "s_preview_nopdfium": "Anteprima non disponibile\n(pypdfium2 mancante)",
        "s_page_x_of_y": "Pagina {i} di {n}",
        "s_pages_only": "{n} pagine",
        "s_split_btn": "Dividi PDF",
        "s_ready": "Pronto. Scegli un PDF, imposta la modalità e premi \"Dividi PDF\".",
        "s_pypdf_missing": "ATTENZIONE: la libreria pypdf non è disponibile: la divisione non funziona.",
        "s_loaded": "Caricato: {name} ({n} pagine)",
        "s_created": "Creato: {name}",
        "s_no_valid_pages": "Nessuna pagina valida negli intervalli indicati.",
        "s_done_msg": "Divisione completata: creati {n} file PDF.",
        "s_done_box": "Creati {n} file PDF.",
        "s_err_split": "Divisione non riuscita. Controlla il log.",
        "s_pdf_invalid_title": "PDF non valido",
        "s_pdf_invalid_msg": "Impossibile leggere il PDF.\n{e}",
        "s_no_pdf_title": "Nessun PDF",
        "s_no_pdf_msg": "Scegli prima un file PDF da dividere.",
        "s_pypdf_err": "pypdf non è installato: impossibile leggere il PDF.",
    },
    "en": {
        "subtitle": "Images and Word → PDF  ·  Split PDF",
        "tab_convert": "Convert to PDF",
        "tab_split": "Split PDF",
        "c_files_title": "1 · Files to convert",
        "c_hint_dnd": "Drag files or folders here, or use \"Add files\"",
        "c_hint_nodnd": "Use \"Add files\" to select files",
        "c_add": "＋ Add files",
        "c_remove": "Remove",
        "c_clear": "Clear",
        "c_preview_title": "Preview of the resulting PDF",
        "c_preview_select": "Select a file\nto see the preview",
        "c_preview_generating": "Generating preview…\n(opening Word)",
        "c_opts_title": "2 · Options",
        "c_save_to": "Save to:",
        "c_out_default": "Same folder as source files",
        "c_choose": "Choose…",
        "c_reset": "Reset",
        "c_merge": "Merge all files into a single PDF (in list order)",
        "c_page_format": "Image page size:",
        "c_fit": "Fit to image",
        "c_a4": "A4",
        "c_letter": "Letter",
        "c_convert_btn": "Convert to PDF",
        "open_folder": "📂 Open folder",
        "c_ready": "Ready. Add files and press \"Convert to PDF\".",
        "c_info_img": "{name}\n{w}×{h} px · {size}",
        "c_info_word": "{name}\nPDF from Word · {size}",
        "m_added": "Added {n} file(s).",
        "m_ignored": "Ignored {n} file(s) with unsupported format.",
        "m_no_files_title": "No files",
        "m_no_files_msg": "Add at least one file to convert.",
        "m_pillow_missing": "Pillow not installed: cannot convert.",
        "m_start_word": "Starting Microsoft Word…",
        "m_ok_named": "OK  {name} → {out}",
        "m_ok_simple": "OK  {name}",
        "m_error_file": "ERROR {name}: {e}",
        "m_merge_created": "→ Single PDF created: {name}",
        "m_generic_error": "ERROR: {e}",
        "m_word_hint": "Make sure Microsoft Word is installed for Word documents.",
        "m_completed_errors": "Completed with {n} error(s).",
        "m_completed_ok": "Conversion completed successfully.",
        "t_completed_title": "Completed",
        "t_completed_warn": "Conversion finished with {n} error(s).\nCheck the log.",
        "t_done_title": "Done",
        "t_all_converted": "All files were converted to PDF.",
        "err_title": "Error",
        "s_file_title": "1 · Choose the PDF to split",
        "s_pdf_hint_dnd": "Drag a PDF here or press \"Choose PDF\"",
        "s_pdf_hint_nodnd": "Press \"Choose PDF\"",
        "s_choose_pdf": "Choose PDF…",
        "s_pages_count": "📄 {name} · {n} pages",
        "s_how_title": "2 · How to split",
        "s_mode_each": "One page per PDF file",
        "s_every_1": "Every",
        "s_every_2": "pages per file",
        "s_extract": "Extract pages:",
        "s_extract_hint": "e.g. 1-3, 5, 8",
        "s_out_default": "Same folder as the PDF",
        "s_preview_title": "Page preview",
        "s_preview_select": "Select a PDF\nto see its pages",
        "s_preview_nopdfium": "Preview not available\n(pypdfium2 missing)",
        "s_page_x_of_y": "Page {i} of {n}",
        "s_pages_only": "{n} pages",
        "s_split_btn": "Split PDF",
        "s_ready": "Ready. Choose a PDF, set the mode and press \"Split PDF\".",
        "s_pypdf_missing": "WARNING: the pypdf library is not available: splitting won't work.",
        "s_loaded": "Loaded: {name} ({n} pages)",
        "s_created": "Created: {name}",
        "s_no_valid_pages": "No valid pages in the given ranges.",
        "s_done_msg": "Split complete: created {n} PDF files.",
        "s_done_box": "Created {n} PDF files.",
        "s_err_split": "Split failed. Check the log.",
        "s_pdf_invalid_title": "Invalid PDF",
        "s_pdf_invalid_msg": "Could not read the PDF.\n{e}",
        "s_no_pdf_title": "No PDF",
        "s_no_pdf_msg": "Choose a PDF file to split first.",
        "s_pypdf_err": "pypdf is not installed: cannot read the PDF.",
    },
}


def kind_of(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in IMAGE_EXTS:
        return "image"
    if ext in WORD_EXTS:
        return "word"
    if ext == ".pdf":
        return "pdf"
    return "unknown"


def human_size(path) -> str:
    try:
        n = float(Path(path).stat().st_size)
    except Exception:
        return ""
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def enable_dpi_awareness():
    """Rende l'app 'DPI aware' su Windows: testo nitido su schermi con scala 125/150%."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Conversione IMMAGINI -> PDF (e generazione pagina per l'anteprima)
# ---------------------------------------------------------------------------
def _prepare_image(path: Path) -> "Image.Image":
    img = Image.open(path)
    if getattr(img, "is_animated", False):
        img.seek(0)
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        background = Image.new("RGB", img.size, (255, 255, 255))
        rgba = img.convert("RGBA")
        background.paste(rgba, mask=rgba.split()[-1])
        return background
    return img.convert("RGB")


def _fit_on_page(img: "Image.Image", page_px, margin=0.04) -> "Image.Image":
    pw, ph = page_px
    if img.width > img.height and pw < ph:
        pw, ph = ph, pw
    page = Image.new("RGB", (pw, ph), "white")
    avail_w = int(pw * (1 - 2 * margin))
    avail_h = int(ph * (1 - 2 * margin))
    scale = min(avail_w / img.width, avail_h / img.height)
    nw, nh = max(1, int(img.width * scale)), max(1, int(img.height * scale))
    resized = img.resize((nw, nh), Image.LANCZOS)
    page.paste(resized, ((pw - nw) // 2, (ph - nh) // 2))
    return page


def image_page(path: Path, page_key="fit") -> "Image.Image":
    """Costruisce l'immagine della pagina PDF risultante (usata anche per l'anteprima)."""
    img = _prepare_image(path)
    size = PAGE_SIZES.get(page_key)
    return _fit_on_page(img, size) if size else img


def image_to_pdf(input_path: Path, output_path: Path, page_key="fit") -> None:
    page = image_page(input_path, page_key)
    res = PAGE_DPI if PAGE_SIZES.get(page_key) else 100.0
    page.save(output_path, "PDF", resolution=res)


# ---------------------------------------------------------------------------
# Conversione WORD -> PDF (automazione Microsoft Word)
# ---------------------------------------------------------------------------
WD_EXPORT_FORMAT_PDF = 17


class WordConverter:
    def __init__(self):
        self.word = None
        self._com = False

    def __enter__(self):
        import comtypes
        import comtypes.client
        comtypes.CoInitialize()
        self._com = True
        self.word = comtypes.client.CreateObject("Word.Application")
        self.word.Visible = False
        try:
            self.word.DisplayAlerts = 0
        except Exception:
            pass
        return self

    def convert(self, input_path: Path, output_path: Path) -> None:
        doc = self.word.Documents.Open(
            str(input_path), ConfirmConversions=False,
            ReadOnly=True, AddToRecentFiles=False,
        )
        try:
            doc.ExportAsFixedFormat(str(output_path), WD_EXPORT_FORMAT_PDF)
        finally:
            doc.Close(False)

    def __exit__(self, *exc):
        try:
            if self.word is not None:
                self.word.Quit()
        except Exception:
            pass
        finally:
            if self._com:
                try:
                    import comtypes
                    comtypes.CoUninitialize()
                except Exception:
                    pass
        return False


# ---------------------------------------------------------------------------
# Operazioni su PDF
# ---------------------------------------------------------------------------
def unique_output_path(folder: Path, stem: str) -> Path:
    candidate = folder / f"{stem}.pdf"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{stem} ({counter}).pdf"
        counter += 1
    return candidate


def merge_pdfs(paths, output_path: Path) -> None:
    writer = PdfWriter()
    for p in paths:
        writer.append(str(p))
    with open(output_path, "wb") as f:
        writer.write(f)
    writer.close()


def pdf_page_count(path) -> int:
    return len(PdfReader(str(path)).pages)


def split_each_page(path: Path, out_dir: Path):
    reader = PdfReader(str(path))
    outs = []
    for i, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        target = unique_output_path(out_dir, f"{path.stem}_pagina_{i}")
        with open(target, "wb") as f:
            writer.write(f)
        outs.append(target)
    return outs


def split_every_n(path: Path, n: int, out_dir: Path):
    reader = PdfReader(str(path))
    pages = reader.pages
    total = len(pages)
    outs = []
    part = 1
    for start in range(0, total, n):
        writer = PdfWriter()
        for page in pages[start:start + n]:
            writer.add_page(page)
        target = unique_output_path(out_dir, f"{path.stem}_parte_{part}")
        with open(target, "wb") as f:
            writer.write(f)
        outs.append(target)
        part += 1
    return outs


def extract_pages(path: Path, indices, out_dir: Path):
    reader = PdfReader(str(path))
    pages = reader.pages
    writer = PdfWriter()
    for i in indices:
        if 0 <= i < len(pages):
            writer.add_page(pages[i])
    target = unique_output_path(out_dir, f"{path.stem}_estratto")
    with open(target, "wb") as f:
        writer.write(f)
    return [target]


def parse_page_ranges(text: str, total: int):
    result = []
    for chunk in text.replace(" ", "").split(","):
        if not chunk:
            continue
        try:
            if "-" in chunk:
                a, b = chunk.split("-", 1)
                a, b = int(a), int(b)
                step = 1 if a <= b else -1
                for p in range(a, b + step, step):
                    result.append(p - 1)
            else:
                result.append(int(chunk) - 1)
        except ValueError:
            continue
    seen, out = set(), []
    for i in result:
        if 0 <= i < total and i not in seen:
            seen.add(i)
            out.append(i)
    return out


def render_pdf_page(path, index: int, scale=1.5) -> "Image.Image":
    doc = pdfium.PdfDocument(str(path))
    try:
        page = doc[index]
        bitmap = page.render(scale=scale)
        return bitmap.to_pil().convert("RGB")
    finally:
        try:
            doc.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Widget di supporto
# ---------------------------------------------------------------------------
def primary_button(parent, text, command):
    b = tk.Button(parent, text=text, command=command, bg=ACCENT, fg="white",
                  activebackground=ACCENT2, activeforeground="white",
                  relief="flat", bd=0, font=("Segoe UI", 11, "bold"),
                  padx=20, pady=9, cursor="hand2")
    b.bind("<Enter>", lambda e: b.config(bg=ACCENT2))
    b.bind("<Leave>", lambda e: b.config(bg=ACCENT))
    return b


def soft_button(parent, text, command, small=True):
    b = tk.Button(parent, text=text, command=command, bg="#EDF0F6", fg=TEXT,
                  activebackground="#DCE2EE", activeforeground=TEXT,
                  relief="flat", bd=0, font=("Segoe UI", 9 if small else 10),
                  padx=10, pady=5, cursor="hand2")
    b.bind("<Enter>", lambda e: b.config(bg="#DCE2EE"))
    b.bind("<Leave>", lambda e: b.config(bg="#EDF0F6"))
    return b


def make_card(parent):
    return tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                    highlightthickness=1, bd=0)


def card_title(card, text):
    lbl = tk.Label(card, text=text, bg=CARD, fg=TEXT, font=("Segoe UI", 11, "bold"))
    lbl.pack(anchor="w", padx=12, pady=(10, 4))
    return lbl


# ---------------------------------------------------------------------------
# Frame base con gestione thread + coda messaggi
# ---------------------------------------------------------------------------
class TaskFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, style="Tab.TFrame", padding=14)
        self.app = app
        self.t = app.t
        self.msg_queue = queue.Queue()
        self.worker = None
        self._last_folder = None
        self.after(120, self._poll)

    def start_worker(self, target, *args):
        if self.worker and self.worker.is_alive():
            return False
        self.worker = threading.Thread(target=target, args=args, daemon=True)
        self.worker.start()
        return True

    def _poll(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    self.progress.config(value=payload)
                elif kind == "progress_max":
                    self.progress.config(maximum=max(1, payload), value=0)
                elif kind == "done":
                    self._on_done(payload)
                else:
                    self._extra_message(kind, payload)
        except queue.Empty:
            pass
        try:
            self.after(120, self._poll)
        except Exception:
            pass

    def _extra_message(self, kind, payload):
        pass

    def _on_done(self, payload):
        pass

    def _log(self, text):
        self.log.config(state="normal")
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.config(state="disabled")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.config(state="disabled")

    def _open_folder(self):
        if self._last_folder and os.path.isdir(self._last_folder):
            os.startfile(self._last_folder)


# ---------------------------------------------------------------------------
# TAB 1 - Conversione in PDF
# ---------------------------------------------------------------------------
class ConvertTab(TaskFrame):
    PREVIEW_W, PREVIEW_H = 320, 420

    def __init__(self, master, app):
        super().__init__(master, app)
        self.files = []
        self.output_dir = None
        self._preview_photo = None
        self._current_pil = None
        self._resize_job = None
        self._preview_cache = {}      # {(path, mtime): PIL Image | None}
        self._preview_seq = 0
        self._preview_lock = threading.Lock()
        self._drag_index = None       # indice della voce trascinata per il riordino
        self.scale = getattr(self.winfo_toplevel(), "scale", 1.0)
        self.pw = int(self.PREVIEW_W * self.scale)
        self.ph = int(self.PREVIEW_H * self.scale)
        self._build()
        self._enable_dnd()

    def _build(self):
        # ---- riga superiore: due colonne di uguale larghezza ----
        top = ttk.Frame(self, style="Tab.TFrame")
        top.pack(fill="both", expand=True)
        top.columnconfigure(0, weight=1, uniform="cols")
        top.columnconfigure(1, weight=1, uniform="cols")
        top.rowconfigure(0, weight=1)

        # --- Scheda file (colonna sinistra) ---
        files_card = make_card(top)
        files_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        card_title(files_card, self.t("c_files_title"))
        hint = self.t("c_hint_dnd") if DND_AVAILABLE else self.t("c_hint_nodnd")
        tk.Label(files_card, text=hint, bg=CARD, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=12)

        list_wrap = tk.Frame(files_card, bg=CARD)
        list_wrap.pack(fill="both", expand=True, padx=12, pady=8)
        sb = ttk.Scrollbar(list_wrap, orient="vertical")
        self.listbox = tk.Listbox(
            list_wrap, selectmode=tk.EXTENDED, yscrollcommand=sb.set,
            activestyle="none", bd=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
            bg="white", fg=TEXT, font=("Segoe UI", 10),
            selectbackground=ACCENT, selectforeground="white",
            exportselection=False,
        )
        sb.config(command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._update_preview)
        # Riordino trascinando il nome del file dentro l'elenco
        self.listbox.bind("<Button-1>", self._drag_start, add="+")
        self.listbox.bind("<B1-Motion>", self._drag_motion, add="+")
        self.listbox.bind("<ButtonRelease-1>", self._drag_end, add="+")

        btns = tk.Frame(files_card, bg=CARD)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        soft_button(btns, self.t("c_add"), self.add_files).pack(side="left")
        soft_button(btns, "▲", self._move_up).pack(side="left", padx=(6, 2))
        soft_button(btns, "▼", self._move_down).pack(side="left", padx=(0, 6))
        soft_button(btns, self.t("c_remove"), self._remove).pack(side="left")
        soft_button(btns, self.t("c_clear"), self._clear).pack(side="left", padx=6)

        # --- Scheda anteprima (colonna destra) ---
        prev_card = make_card(top)
        prev_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        card_title(prev_card, self.t("c_preview_title"))
        self.preview_box = tk.Frame(prev_card, bg=FIELD, highlightthickness=1,
                                    highlightbackground=BORDER)
        self.preview_box.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self.preview_img = tk.Label(self.preview_box, bg=FIELD, fg=MUTED,
                                    text=self.t("c_preview_select"),
                                    font=("Segoe UI", 11), justify="center")
        self.preview_img.pack(expand=True)
        self.preview_box.bind("<Configure>", self._on_box_resize)
        self.preview_info = tk.Label(prev_card, text="", bg=CARD, fg=TEXT,
                                     font=("Segoe UI", 9), justify="center",
                                     wraplength=self.pw)
        self.preview_info.pack(padx=12, pady=(0, 12))

        # ---- scheda opzioni ----
        opt_card = make_card(self)
        opt_card.pack(fill="x", pady=(12, 0))
        card_title(opt_card, self.t("c_opts_title"))

        row1 = tk.Frame(opt_card, bg=CARD)
        row1.pack(fill="x", padx=12, pady=(0, 8))
        tk.Label(row1, text=self.t("c_save_to"), bg=CARD, fg=TEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        self.out_var = tk.StringVar(value=self.t("c_out_default"))
        self.out_entry = ttk.Entry(row1, textvariable=self.out_var, state="readonly")
        self.out_entry.pack(side="left", fill="x", expand=True, padx=8)
        soft_button(row1, self.t("c_choose"), self._choose_output).pack(side="left")
        soft_button(row1, self.t("c_reset"), self._reset_output).pack(side="left", padx=(6, 0))

        row2 = tk.Frame(opt_card, bg=CARD)
        row2.pack(fill="x", padx=12, pady=(0, 8))
        self.merge_var = tk.BooleanVar(value=False)
        tk.Checkbutton(row2, text=self.t("c_merge"), variable=self.merge_var,
                       bg=CARD, fg=TEXT, activebackground=CARD, font=("Segoe UI", 9),
                       selectcolor="white", highlightthickness=0, bd=0).pack(anchor="w")

        row3 = tk.Frame(opt_card, bg=CARD)
        row3.pack(fill="x", padx=12, pady=(0, 12))
        tk.Label(row3, text=self.t("c_page_format"), bg=CARD, fg=TEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        self.page_var = tk.StringVar(value="fit")
        for val, key in (("fit", "c_fit"), ("a4", "c_a4"), ("letter", "c_letter")):
            tk.Radiobutton(row3, text=self.t(key), value=val, variable=self.page_var,
                           bg=CARD, fg=TEXT, activebackground=CARD, selectcolor="white",
                           font=("Segoe UI", 9), highlightthickness=0, bd=0).pack(side="left", padx=(10, 0))

        # aggiorna l'anteprima quando cambiano le opzioni
        self.page_var.trace_add("write", lambda *a: self._update_preview())
        self.merge_var.trace_add("write", lambda *a: self._update_preview())

        # ---- azioni + progress ----
        action = tk.Frame(self, bg=BG)
        action.pack(fill="x", pady=(12, 6))
        self.convert_btn = primary_button(action, self.t("c_convert_btn"), self.start)
        self.convert_btn.pack(side="left")
        self.open_btn = soft_button(action, self.t("open_folder"), self._open_folder, small=False)
        self.open_btn.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(action, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(12, 0))

        # ---- log ----
        log_card = make_card(self)
        log_card.pack(fill="both", expand=False, pady=(6, 0))
        self.log = tk.Text(log_card, height=5, wrap="word", state="disabled",
                           bd=0, bg="white", fg=TEXT, font=("Consolas", 9),
                           highlightthickness=0)
        self.log.pack(fill="both", expand=True, padx=10, pady=10)
        self._log(self.t("c_ready"))

    # -- Drag & drop --------------------------------------------------------
    def _enable_dnd(self):
        if not DND_AVAILABLE:
            return
        try:
            for w in (self.listbox, self):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        try:
            raw = self.tk.splitlist(event.data)
        except Exception:
            raw = [event.data]
        self._add_paths([Path(str(p)) for p in raw])

    # -- Elenco file --------------------------------------------------------
    def add_files(self):
        filetypes = [
            ("Immagini / Word",
             "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp *.ico *.jfif "
             "*.doc *.docx *.docm *.rtf *.odt"),
            ("Immagini", "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp *.ico *.jfif"),
            ("Word", "*.doc *.docx *.docm *.rtf *.odt"),
            ("Tutti i file / All files", "*.*"),
        ]
        paths = filedialog.askopenfilenames(title=self.t("c_add"), filetypes=filetypes)
        self._add_paths([Path(p) for p in paths])

    def _icon_for(self, path):
        return "🖼" if kind_of(path) == "image" else "📄"

    def _add_one(self, path: Path) -> bool:
        if kind_of(path) not in ("image", "word") or path in self.files:
            return False
        self.files.append(path)
        self.listbox.insert(tk.END, f"  {self._icon_for(path)}  {path.name}")
        return True

    def _add_paths(self, paths):
        added, skipped = 0, 0
        for path in paths:
            if path.is_dir():
                for child in sorted(path.iterdir()):
                    if child.is_file() and self._add_one(child):
                        added += 1
                continue
            if not path.is_file():
                continue
            if kind_of(path) not in ("image", "word"):
                skipped += 1
                continue
            if self._add_one(path):
                added += 1
        if added:
            self._log(self.t("m_added", n=added))
            if not self.listbox.curselection() and self.listbox.size() > 0:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(0)
                self.listbox.activate(0)
                self._update_preview()
        if skipped:
            self._log(self.t("m_ignored", n=skipped))

    def _selected_index(self):
        sel = self.listbox.curselection()
        return sel[0] if sel else None

    def _refresh_list(self, select=None):
        self.listbox.delete(0, tk.END)
        for p in self.files:
            self.listbox.insert(tk.END, f"  {self._icon_for(p)}  {p.name}")
        if select is not None and 0 <= select < len(self.files):
            self.listbox.selection_set(select)
            self.listbox.activate(select)
        self._update_preview()

    def _move_up(self):
        i = self._selected_index()
        if i is None or i == 0:
            return
        self.files[i - 1], self.files[i] = self.files[i], self.files[i - 1]
        self._refresh_list(select=i - 1)

    def _move_down(self):
        i = self._selected_index()
        if i is None or i >= len(self.files) - 1:
            return
        self.files[i + 1], self.files[i] = self.files[i], self.files[i + 1]
        self._refresh_list(select=i + 1)

    # -- Riordino trascinando il nome --------------------------------------
    def _drag_start(self, event):
        self._drag_index = self.listbox.nearest(event.y)

    def _drag_motion(self, event):
        if self._drag_index is None:
            return
        new = self.listbox.nearest(event.y)
        if new < 0 or new >= len(self.files) or new == self._drag_index:
            return
        # sposta la voce (file + riga della lista) nella nuova posizione
        item = self.files.pop(self._drag_index)
        self.files.insert(new, item)
        text = self.listbox.get(self._drag_index)
        self.listbox.delete(self._drag_index)
        self.listbox.insert(new, text)
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(new)
        self.listbox.activate(new)
        self._drag_index = new
        self.listbox.config(cursor="fleur")

    def _drag_end(self, event):
        moved = self._drag_index is not None
        self._drag_index = None
        self.listbox.config(cursor="")
        if moved:
            self._update_preview()

    def _remove(self):
        for idx in sorted(self.listbox.curselection(), reverse=True):
            self.listbox.delete(idx)
            del self.files[idx]
        self._update_preview()

    def _clear(self):
        self.listbox.delete(0, tk.END)
        self.files.clear()
        self._update_preview()

    def _choose_output(self):
        folder = filedialog.askdirectory(title=self.t("c_save_to"))
        if folder:
            self.output_dir = Path(folder)
            self.out_var.set(folder)

    def _reset_output(self):
        self.output_dir = None
        self.out_var.set(self.t("c_out_default"))

    # -- Anteprima del PDF risultante --------------------------------------
    def _placeholder(self, text, big=False):
        self._current_pil = None
        self._preview_photo = None
        self.preview_img.config(image="", text=text,
                                font=("Segoe UI", 40 if big else 11))
        self.preview_info.config(text="")

    def _show_pil(self, pil_img, info_text):
        self._current_pil = pil_img
        self.preview_info.config(text=info_text)
        self._render_current()

    def _render_current(self):
        if self._current_pil is None or ImageTk is None:
            return
        self.preview_box.update_idletasks()
        bw = self.preview_box.winfo_width() - 16
        bh = self.preview_box.winfo_height() - 16
        if bw < 60 or bh < 60:
            bw, bh = self.pw - 16, self.ph - 16
        thumb = self._current_pil.copy()
        thumb.thumbnail((bw, bh), Image.LANCZOS)
        self._preview_photo = ImageTk.PhotoImage(thumb)
        self.preview_img.config(image=self._preview_photo, text="")
        self.preview_info.config(wraplength=max(self.pw, self.preview_box.winfo_width()))

    def _on_box_resize(self, event=None):
        if self._resize_job:
            try:
                self.after_cancel(self._resize_job)
            except Exception:
                pass
        self._resize_job = self.after(90, self._render_current)

    def _cache_key(self, path):
        try:
            return (str(path), os.path.getmtime(path))
        except OSError:
            return (str(path), 0)

    def _update_preview(self, event=None):
        i = self._selected_index()
        if i is None or i >= len(self.files):
            self._placeholder(self.t("c_preview_select"))
            return
        path = self.files[i]
        kind = kind_of(path)
        page_key = self.page_var.get()

        if kind == "image" and Image and ImageTk:
            try:
                base = _prepare_image(path)
                w, h = base.size
                size = PAGE_SIZES.get(page_key)
                page_img = _fit_on_page(base, size) if size else base
                self._show_pil(page_img, self.t("c_info_img", name=path.name, w=w, h=h,
                                                 size=human_size(path)))
            except Exception:
                self._placeholder(self.t("c_preview_select"))
            return

        if kind == "word":
            # anteprima reale del PDF generato da Word (in background, con cache)
            if not (PDFIUM_AVAILABLE and Image and ImageTk):
                self._word_info_placeholder(path)
                return
            key = self._cache_key(path)
            if key in self._preview_cache:
                img = self._preview_cache[key]
                if img is not None:
                    self._show_pil(img, self.t("c_info_word", name=path.name, size=human_size(path)))
                else:
                    self._word_info_placeholder(path)
                return
            if self.worker and self.worker.is_alive():
                # conversione in corso: evito una seconda istanza di Word
                self._word_info_placeholder(path)
                return
            self._placeholder(self.t("c_preview_generating"))
            self._preview_seq += 1
            threading.Thread(target=self._make_word_preview,
                             args=(path, key, self._preview_seq), daemon=True).start()
            return

        self._placeholder(self.t("c_preview_select"))

    def _word_info_placeholder(self, path):
        self._placeholder("📄", big=True)
        self.preview_info.config(text=self.t("c_info_word", name=path.name, size=human_size(path)))

    def _make_word_preview(self, path, key, seq):
        img = None
        try:
            with self._preview_lock:
                tmp = tempfile.mkdtemp(prefix="prev_")
                out = Path(tmp) / "p.pdf"
                with WordConverter() as wc:
                    wc.convert(path, out)
                img = render_pdf_page(out, 0, scale=1.5)
                shutil.rmtree(tmp, ignore_errors=True)
        except Exception:
            img = None
        self.msg_queue.put(("wpreview", (key, seq, img)))

    def _extra_message(self, kind, payload):
        if kind != "wpreview":
            return
        key, seq, img = payload
        self._preview_cache[key] = img
        if seq != self._preview_seq:
            return
        i = self._selected_index()
        if i is None or i >= len(self.files):
            return
        path = self.files[i]
        if self._cache_key(path) != key:
            return
        if img is not None:
            self._show_pil(img, self.t("c_info_word", name=path.name, size=human_size(path)))
        else:
            self._word_info_placeholder(path)

    # -- Stato (per il cambio lingua) --------------------------------------
    def get_state(self):
        return {
            "files": list(self.files),
            "output_dir": self.output_dir,
            "merge": self.merge_var.get(),
            "page": self.page_var.get(),
            "sel": self._selected_index(),
        }

    def set_state(self, s):
        self.files = list(s.get("files", []))
        self.output_dir = s.get("output_dir")
        if self.output_dir:
            self.out_var.set(str(self.output_dir))
        self.merge_var.set(s.get("merge", False))
        self.page_var.set(s.get("page", "fit"))
        self._refresh_list(select=s.get("sel"))

    # -- Avvio conversione --------------------------------------------------
    def start(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.files:
            messagebox.showinfo(self.t("m_no_files_title"), self.t("m_no_files_msg"))
            return
        if Image is None:
            messagebox.showerror(self.t("err_title"), self.t("m_pillow_missing"))
            return
        self.convert_btn.config(state="disabled")
        self._clear_log()
        self.start_worker(self._work, list(self.files), self.output_dir,
                          self.merge_var.get(), self.page_var.get())

    def _work(self, files, out_dir, merge, page_key):
        q = self.msg_queue
        q.put(("progress_max", len(files) + (1 if merge else 0)))
        errors, done, last_folder = 0, 0, None
        words_present = any(kind_of(p) == "word" for p in files)

        def convert_one(path, target, wc):
            k = kind_of(path)
            if k == "image":
                image_to_pdf(path, target, page_key)
            elif k == "word":
                if wc is None:
                    raise RuntimeError("Word non disponibile")
                wc.convert(path, target)
            else:
                raise RuntimeError("formato non supportato")

        try:
            wc_ctx = WordConverter() if words_present else nullcontext()
            if merge:
                tmp = tempfile.mkdtemp(prefix="conv_")
                temp_pdfs = []
                try:
                    if words_present:
                        q.put(("log", self.t("m_start_word")))
                    with wc_ctx as wc:
                        for idx, p in enumerate(files):
                            tp = Path(tmp) / f"{idx:04d}.pdf"
                            try:
                                convert_one(p, tp, wc)
                                temp_pdfs.append(tp)
                                q.put(("log", self.t("m_ok_simple", name=p.name)))
                            except Exception as e:
                                errors += 1
                                q.put(("log", self.t("m_error_file", name=p.name, e=e)))
                            done += 1
                            q.put(("progress", done))
                    if temp_pdfs:
                        folder = out_dir or files[0].parent
                        folder.mkdir(parents=True, exist_ok=True)
                        target = unique_output_path(folder, "documento_unito")
                        merge_pdfs(temp_pdfs, target)
                        last_folder = folder
                        q.put(("log", self.t("m_merge_created", name=target.name)))
                    done += 1
                    q.put(("progress", done))
                finally:
                    shutil.rmtree(tmp, ignore_errors=True)
            else:
                if words_present:
                    q.put(("log", self.t("m_start_word")))
                with wc_ctx as wc:
                    for p in files:
                        folder = out_dir or p.parent
                        folder.mkdir(parents=True, exist_ok=True)
                        target = unique_output_path(folder, p.stem)
                        try:
                            convert_one(p, target, wc)
                            last_folder = folder
                            q.put(("log", self.t("m_ok_named", name=p.name, out=target.name)))
                        except Exception as e:
                            errors += 1
                            q.put(("log", self.t("m_error_file", name=p.name, e=e)))
                        done += 1
                        q.put(("progress", done))
        except Exception as e:
            errors += 1
            q.put(("log", self.t("m_generic_error", e=e)))
            q.put(("log", self.t("m_word_hint")))
        finally:
            q.put(("done", {"errors": errors, "folder": str(last_folder) if last_folder else None}))

    def _on_done(self, payload):
        self.convert_btn.config(state="normal")
        self._last_folder = payload.get("folder")
        errors = payload.get("errors", 0)
        if errors:
            self._log(self.t("m_completed_errors", n=errors))
            messagebox.showwarning(self.t("t_completed_title"), self.t("t_completed_warn", n=errors))
        else:
            self._log(self.t("m_completed_ok"))
            messagebox.showinfo(self.t("t_done_title"), self.t("t_all_converted"))


# ---------------------------------------------------------------------------
# TAB 2 - Dividi PDF
# ---------------------------------------------------------------------------
class SplitTab(TaskFrame):
    PREVIEW_W, PREVIEW_H = 320, 420

    def __init__(self, master, app):
        super().__init__(master, app)
        self.pdf_path = None
        self.output_dir = None
        self.page_count = 0
        self.page_index = 0
        self._preview_photo = None
        self.scale = getattr(self.winfo_toplevel(), "scale", 1.0)
        self.pw = int(self.PREVIEW_W * self.scale)
        self.ph = int(self.PREVIEW_H * self.scale)
        self._build()
        self._enable_dnd()

    def _build(self):
        file_card = make_card(self)
        file_card.pack(fill="x")
        card_title(file_card, self.t("s_file_title"))
        row = tk.Frame(file_card, bg=CARD)
        row.pack(fill="x", padx=12, pady=(0, 12))
        self.pdf_var = tk.StringVar(
            value=self.t("s_pdf_hint_dnd") if DND_AVAILABLE else self.t("s_pdf_hint_nodnd"))
        self.pdf_entry = ttk.Entry(row, textvariable=self.pdf_var, state="readonly")
        self.pdf_entry.pack(side="left", fill="x", expand=True)
        soft_button(row, self.t("s_choose_pdf"), self._choose_pdf).pack(side="left", padx=(8, 0))
        self.count_label = tk.Label(file_card, text="", bg=CARD, fg=ACCENT2,
                                    font=("Segoe UI", 9, "bold"))
        self.count_label.pack(anchor="w", padx=12, pady=(0, 10))

        mid = ttk.Frame(self, style="Tab.TFrame")
        mid.pack(fill="both", expand=True, pady=(12, 0))
        mid.columnconfigure(0, weight=1, uniform="cols")
        mid.columnconfigure(1, weight=1, uniform="cols")
        mid.rowconfigure(0, weight=1)

        opt_card = make_card(mid)
        opt_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        card_title(opt_card, self.t("s_how_title"))

        self.mode_var = tk.StringVar(value="each")
        tk.Radiobutton(opt_card, text=self.t("s_mode_each"), value="each", variable=self.mode_var,
                       bg=CARD, fg=TEXT, activebackground=CARD, selectcolor="white",
                       font=("Segoe UI", 9), highlightthickness=0, bd=0).pack(anchor="w", padx=12, pady=2)

        r_every = tk.Frame(opt_card, bg=CARD)
        r_every.pack(fill="x", padx=12, pady=2)
        tk.Radiobutton(r_every, text=self.t("s_every_1"), value="every", variable=self.mode_var,
                       bg=CARD, fg=TEXT, activebackground=CARD, selectcolor="white",
                       font=("Segoe UI", 9), highlightthickness=0, bd=0).pack(side="left")
        self.every_var = tk.StringVar(value="2")
        ttk.Spinbox(r_every, from_=1, to=999, width=4, textvariable=self.every_var).pack(side="left", padx=4)
        tk.Label(r_every, text=self.t("s_every_2"), bg=CARD, fg=TEXT,
                 font=("Segoe UI", 9)).pack(side="left")

        r_range = tk.Frame(opt_card, bg=CARD)
        r_range.pack(fill="x", padx=12, pady=2)
        tk.Radiobutton(r_range, text=self.t("s_extract"), value="ranges", variable=self.mode_var,
                       bg=CARD, fg=TEXT, activebackground=CARD, selectcolor="white",
                       font=("Segoe UI", 9), highlightthickness=0, bd=0).pack(side="left")
        self.ranges_var = tk.StringVar(value="")
        ttk.Entry(r_range, textvariable=self.ranges_var, width=16).pack(side="left", padx=4)
        tk.Label(r_range, text=self.t("s_extract_hint"), bg=CARD, fg=MUTED,
                 font=("Segoe UI", 8)).pack(side="left")

        row_out = tk.Frame(opt_card, bg=CARD)
        row_out.pack(fill="x", padx=12, pady=(12, 12))
        tk.Label(row_out, text=self.t("c_save_to"), bg=CARD, fg=TEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        self.out_var = tk.StringVar(value=self.t("s_out_default"))
        ttk.Entry(row_out, textvariable=self.out_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=8)
        soft_button(row_out, self.t("c_choose"), self._choose_output).pack(side="left")

        prev_card = make_card(mid)
        prev_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        card_title(prev_card, self.t("s_preview_title"))
        self.preview_box = tk.Frame(prev_card, bg=FIELD, highlightthickness=1,
                                    highlightbackground=BORDER)
        self.preview_box.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        msg = self.t("s_preview_select") if PDFIUM_AVAILABLE else self.t("s_preview_nopdfium")
        self.preview_img = tk.Label(self.preview_box, bg=FIELD, fg=MUTED, text=msg,
                                    font=("Segoe UI", 11), justify="center")
        self.preview_img.pack(expand=True)
        self.preview_box.bind("<Configure>", self._on_box_resize)
        nav = tk.Frame(prev_card, bg=CARD)
        nav.pack(pady=(0, 12))
        self.prev_btn = soft_button(nav, "◀", self._prev_page)
        self.prev_btn.pack(side="left")
        self.page_lbl = tk.Label(nav, text="—", bg=CARD, fg=TEXT, font=("Segoe UI", 9), width=14)
        self.page_lbl.pack(side="left", padx=6)
        self.next_btn = soft_button(nav, "▶", self._next_page)
        self.next_btn.pack(side="left")

        action = tk.Frame(self, bg=BG)
        action.pack(fill="x", pady=(12, 6))
        self.split_btn = primary_button(action, self.t("s_split_btn"), self.start)
        self.split_btn.pack(side="left")
        self.open_btn = soft_button(action, self.t("open_folder"), self._open_folder, small=False)
        self.open_btn.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(action, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(12, 0))

        log_card = make_card(self)
        log_card.pack(fill="both", expand=False, pady=(6, 0))
        self.log = tk.Text(log_card, height=4, wrap="word", state="disabled",
                           bd=0, bg="white", fg=TEXT, font=("Consolas", 9),
                           highlightthickness=0)
        self.log.pack(fill="both", expand=True, padx=10, pady=10)
        self._resize_job = None
        self._current_pil = None
        self._log(self.t("s_ready") if PYPDF_AVAILABLE else self.t("s_pypdf_missing"))

    # -- Drag & drop --------------------------------------------------------
    def _enable_dnd(self):
        if not DND_AVAILABLE:
            return
        try:
            for w in (self.pdf_entry, self):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        try:
            raw = self.tk.splitlist(event.data)
        except Exception:
            raw = [event.data]
        for p in raw:
            path = Path(str(p))
            if path.suffix.lower() == ".pdf" and path.is_file():
                self._load_pdf(path)
                return

    # -- Selezione PDF ------------------------------------------------------
    def _choose_pdf(self):
        p = filedialog.askopenfilename(title=self.t("s_choose_pdf"),
                                       filetypes=[("PDF", "*.pdf"), ("Tutti i file / All files", "*.*")])
        if p:
            self._load_pdf(Path(p))

    def _load_pdf(self, path: Path):
        if not PYPDF_AVAILABLE:
            messagebox.showerror(self.t("err_title"), self.t("s_pypdf_err"))
            return
        try:
            self.page_count = pdf_page_count(path)
        except Exception as e:
            messagebox.showerror(self.t("s_pdf_invalid_title"), self.t("s_pdf_invalid_msg", e=e))
            return
        self.pdf_path = path
        self.pdf_var.set(str(path))
        self.count_label.config(text=self.t("s_pages_count", name=path.name, n=self.page_count))
        self.page_index = 0
        self._render_preview()
        self._log(self.t("s_loaded", name=path.name, n=self.page_count))

    def _choose_output(self):
        folder = filedialog.askdirectory(title=self.t("c_save_to"))
        if folder:
            self.output_dir = Path(folder)
            self.out_var.set(folder)

    # -- Anteprima pagine ---------------------------------------------------
    def _render_preview(self):
        if not (self.pdf_path and PDFIUM_AVAILABLE and Image and ImageTk):
            self.page_lbl.config(text=self.t("s_pages_only", n=self.page_count) if self.page_count else "—")
            return
        try:
            self._current_pil = render_pdf_page(self.pdf_path, self.page_index, scale=1.5)
            self._render_current()
        except Exception:
            self._current_pil = None
            self.preview_img.config(image="", text=self.t("s_preview_select"))
        self.page_lbl.config(text=self.t("s_page_x_of_y", i=self.page_index + 1, n=self.page_count))

    def _render_current(self):
        if self._current_pil is None or ImageTk is None:
            return
        self.preview_box.update_idletasks()
        bw = self.preview_box.winfo_width() - 16
        bh = self.preview_box.winfo_height() - 16
        if bw < 60 or bh < 60:
            bw, bh = self.pw - 16, self.ph - 16
        thumb = self._current_pil.copy()
        thumb.thumbnail((bw, bh), Image.LANCZOS)
        self._preview_photo = ImageTk.PhotoImage(thumb)
        self.preview_img.config(image=self._preview_photo, text="")

    def _on_box_resize(self, event=None):
        if self._resize_job:
            try:
                self.after_cancel(self._resize_job)
            except Exception:
                pass
        self._resize_job = self.after(90, self._render_current)

    def _prev_page(self):
        if self.pdf_path and self.page_index > 0:
            self.page_index -= 1
            self._render_preview()

    def _next_page(self):
        if self.pdf_path and self.page_index < self.page_count - 1:
            self.page_index += 1
            self._render_preview()

    # -- Stato (per il cambio lingua) --------------------------------------
    def get_state(self):
        return {
            "pdf_path": self.pdf_path, "output_dir": self.output_dir,
            "mode": self.mode_var.get(), "every": self.every_var.get(),
            "ranges": self.ranges_var.get(), "page_index": self.page_index,
        }

    def set_state(self, s):
        self.mode_var.set(s.get("mode", "each"))
        self.every_var.set(s.get("every", "2"))
        self.ranges_var.set(s.get("ranges", ""))
        if s.get("output_dir"):
            self.output_dir = s["output_dir"]
            self.out_var.set(str(self.output_dir))
        if s.get("pdf_path"):
            self._load_pdf(s["pdf_path"])
            self.page_index = min(s.get("page_index", 0), max(0, self.page_count - 1))
            self._render_preview()

    # -- Avvio divisione ----------------------------------------------------
    def start(self):
        if self.worker and self.worker.is_alive():
            return
        if not PYPDF_AVAILABLE:
            messagebox.showerror(self.t("err_title"), self.t("s_pypdf_err"))
            return
        if not self.pdf_path:
            messagebox.showinfo(self.t("s_no_pdf_title"), self.t("s_no_pdf_msg"))
            return
        mode = self.mode_var.get()
        try:
            n = int(self.every_var.get())
        except ValueError:
            n = 1
        self.split_btn.config(state="disabled")
        self.progress.config(mode="indeterminate")
        self.progress.start(12)
        self._clear_log()
        self.start_worker(self._work, self.pdf_path, self.output_dir, mode, n,
                          self.ranges_var.get(), self.page_count)

    def _work(self, pdf_path, out_dir, mode, n, ranges_text, total):
        q = self.msg_queue
        try:
            folder = out_dir or pdf_path.parent
            folder.mkdir(parents=True, exist_ok=True)
            if mode == "each":
                outs = split_each_page(pdf_path, folder)
            elif mode == "every":
                outs = split_every_n(pdf_path, max(1, n), folder)
            else:
                idx = parse_page_ranges(ranges_text, total)
                if not idx:
                    q.put(("log", self.t("s_no_valid_pages")))
                    q.put(("done", {"errors": 1, "folder": None}))
                    return
                outs = extract_pages(pdf_path, idx, folder)
            for o in outs:
                q.put(("log", self.t("s_created", name=o.name)))
            q.put(("done", {"errors": 0, "folder": str(folder), "count": len(outs)}))
        except Exception as e:
            q.put(("log", self.t("m_generic_error", e=e)))
            q.put(("done", {"errors": 1, "folder": None}))

    def _on_done(self, payload):
        self.progress.stop()
        self.progress.config(mode="determinate", value=0)
        self.split_btn.config(state="normal")
        self._last_folder = payload.get("folder")
        if payload.get("errors"):
            messagebox.showwarning(self.t("err_title"), self.t("s_err_split"))
        else:
            count = payload.get("count", 0)
            self._log(self.t("s_done_msg", n=count))
            messagebox.showinfo(self.t("t_done_title"), self.t("s_done_box", n=count))


# ---------------------------------------------------------------------------
# Finestra principale
# ---------------------------------------------------------------------------
class App(_BaseTk):
    def __init__(self):
        super().__init__()
        self.lang = "it"
        try:
            self.scale = max(1.0, self.winfo_fpixels("1i") / 96.0)
        except Exception:
            self.scale = 1.0
        try:
            self.tk.call("tk", "scaling", self.winfo_fpixels("1i") / 72.0)
        except Exception:
            pass
        self.title("Convertitore PDF")
        self.geometry(f"{int(960 * self.scale)}x{int(820 * self.scale)}")
        self.minsize(int(880 * self.scale), int(740 * self.scale))
        self.configure(bg=BG)
        self._set_icon()
        self._setup_style()
        self._build_header()
        self._build_notebook()

    def t(self, key, **kw):
        s = TR.get(self.lang, TR["it"]).get(key, key)
        try:
            return s.format(**kw) if kw else s
        except Exception:
            return s

    def _set_icon(self):
        try:
            self.iconbitmap(resource_path("assets/icon.ico"))
        except Exception:
            pass

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=("Segoe UI", 10))
        style.configure("Tab.TFrame", background=BG)
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(26, 12), font=("Segoe UI", 10, "bold"),
                        background="#D5DAE4", foreground=MUTED, borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", CARD)],
                  foreground=[("selected", ACCENT2)],
                  expand=[("selected", [0, 0, 0, 0])])
        style.configure("TProgressbar", background=ACCENT, troughcolor=FIELD,
                        borderwidth=0, thickness=10)
        style.configure("TEntry", fieldbackground="white", bordercolor=BORDER)
        style.configure("TSpinbox", fieldbackground="white", bordercolor=BORDER)
        style.configure("TScrollbar", background=BG, troughcolor=BG, borderwidth=0)

    def _build_header(self):
        # Nessuna altezza fissa: l'intestazione si adatta al contenuto (sottotitolo sempre visibile)
        header = tk.Frame(self, bg=ACCENT2)
        header.pack(fill="x")

        left = tk.Frame(header, bg=ACCENT2)
        left.pack(side="left", padx=22, pady=(12, 14))
        tk.Label(left, text="Convertitore PDF", bg=ACCENT2, fg="white",
                 font=("Segoe UI", 18, "bold")).pack(anchor="w")
        self.subtitle = tk.Label(left, text=self.t("subtitle"), bg=ACCENT2,
                                 fg="#DCE3FF", font=("Segoe UI", 10))
        self.subtitle.pack(anchor="w")

        right = tk.Frame(header, bg=ACCENT2)
        right.pack(side="right", padx=18, pady=12)
        tk.Label(right, text="🌐", bg=ACCENT2, fg="white",
                 font=("Segoe UI", 12)).pack(side="left", padx=(0, 6))
        self.btn_it = self._lang_button(right, "it", "IT")
        self.btn_it.pack(side="left", padx=(0, 4))
        self.btn_en = self._lang_button(right, "en", "EN")
        self.btn_en.pack(side="left")
        self._update_lang_buttons()

    def _lang_button(self, parent, code, label):
        return tk.Button(parent, text=label, command=lambda: self._switch_language(code),
                         relief="flat", bd=0, font=("Segoe UI", 9, "bold"),
                         padx=12, pady=4, cursor="hand2")

    def _update_lang_buttons(self):
        for code, b in (("it", self.btn_it), ("en", self.btn_en)):
            if code == self.lang:
                b.config(bg="white", fg=ACCENT2, activebackground="white", activeforeground=ACCENT2)
            else:
                b.config(bg=ACCENT2, fg="white", activebackground=ACCENT, activeforeground="white")

    def _build_notebook(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        self.convert_tab = ConvertTab(self.nb, self)
        self.split_tab = SplitTab(self.nb, self)
        self.nb.add(self.convert_tab, text=self.t("tab_convert").center(20))
        self.nb.add(self.split_tab, text=self.t("tab_split").center(20))

    def _switch_language(self, lang):
        if lang == self.lang:
            return
        conv_state = self.convert_tab.get_state()
        split_state = self.split_tab.get_state()
        try:
            active = self.nb.index(self.nb.select())
        except Exception:
            active = 0
        self.lang = lang
        self._update_lang_buttons()
        self.subtitle.config(text=self.t("subtitle"))
        self.nb.destroy()
        self._build_notebook()
        self.convert_tab.set_state(conv_state)
        self.split_tab.set_state(split_state)
        try:
            self.nb.select(active)
        except Exception:
            pass


def main():
    enable_dpi_awareness()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
