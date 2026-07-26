#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import csv
import hashlib
import io
import json
import logging
import os
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from pathlib import PurePosixPath
import tkinter as tk
from tkinter import filedialog, messagebox

from data_manager import (
    DATA_DIR,
    T,
    days_to_next_wartung,
    kat_ui,
    nacti_poruchy,
    nacti_stroje,
    slozka_stroje,
)

LOGGER = logging.getLogger("sgm.export")
BACKUP_FORMAT_VERSION = 2
BACKUP_MANIFEST = "sgm_backup_manifest.json"


def vyber_fotky_dialog(parent, image_paths: list):
    """
    Dialog pro výběr fotek k exportu do PDF s miniaturami a tooltip náhledem.
    Vrací list vybraných cest k obrázkům nebo None při zrušení.
    """
    if not image_paths:
        return []

    try:
        from PIL import Image as PILImage, ImageTk
    except ImportError:
        messagebox.showwarning(
            T("Miniatury", "Miniaturansicht"),
            T(
                "Knihovna Pillow není nainstalována.\nNainstaluj: py -m pip install Pillow\n\nDialog se otevře bez miniatur.",
                "Pillow-Bibliothek nicht installiert.\nInstallieren: py -m pip install Pillow\n\nDialog wird ohne Miniaturansichten geöffnet.",
            ),
            parent=parent,
        )
        return vyber_fotky_dialog_bez_miniatur(parent, image_paths)

    win = tk.Toplevel(parent)
    win.title(T("Výběr fotek pro PDF export", "Fotos für PDF-Export auswählen"))
    win.geometry("650x500")
    win.transient(parent)
    win.grab_set()
    win.resizable(True, True)

    header = tk.Frame(win, padx=10, pady=10)
    header.pack(fill="x")
    tk.Label(
        header,
        text=T("Vyberte fotky pro export do PDF:", "Wählen Sie Fotos für PDF-Export:"),
        font=("Segoe UI", 11, "bold"),
    ).pack(anchor="w")
    tk.Label(
        header,
        text=T(
            f"Celkem nalezeno: {len(image_paths)} fotek • Najeďte myší pro náhled",
            f"Insgesamt gefunden: {len(image_paths)} Fotos • Mit Maus überfahren für Vorschau",
        ),
        fg="gray",
    ).pack(anchor="w")

    btn_frame = tk.Frame(win, padx=10, pady=5)
    btn_frame.pack(fill="x")

    list_frame = tk.Frame(win)
    list_frame.pack(fill="both", expand=True, padx=10, pady=5)

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")

    canvas = tk.Canvas(list_frame, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=canvas.yview)

    inner_frame = tk.Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    check_vars = []
    thumbnails = []
    tooltip_window = None
    tooltip_label = None

    preview = {"win": None, "lbl": None, "img": None, "path": None}

    def _hide_preview(_=None):
        w = preview.get("win")
        if w is not None:
            try:
                w.destroy()
            except Exception:
                pass
        preview.update({"win": None, "lbl": None, "img": None, "path": None})

    def show_tooltip(event, image_path):
        nonlocal tooltip_window, tooltip_label
        hide_tooltip()
        try:
            img = PILImage.open(image_path)
            img.thumbnail((700, 520), PILImage.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            tooltip_window = tk.Toplevel(win)
            tooltip_window.wm_overrideredirect(True)
            tooltip_window.attributes("-topmost", True)
            x = event.x_root + 15
            y = event.y_root - 50
            screen_width = win.winfo_screenwidth()
            screen_height = win.winfo_screenheight()
            tooltip_window.geometry(f"+{x}+{y}")
            frame = tk.Frame(tooltip_window, relief="solid", borderwidth=2, bg="white")
            frame.pack()
            tooltip_label = tk.Label(frame, image=photo, bg="white")
            tooltip_label.image = photo
            tooltip_label.pack(padx=2, pady=2)
            info = tk.Label(frame, text=image_path.name, bg="white", font=("Segoe UI", 8), wraplength=280)
            info.pack(padx=5, pady=(0, 5))
            tooltip_window.update_idletasks()
            tw = tooltip_window.winfo_width()
            th = tooltip_window.winfo_height()
            if x + tw > screen_width - 10:
                x = event.x_root - tw - 15
            if y + th > screen_height - 10:
                y = event.y_root - th - 10
                if y < 10:
                    y = 10
            tooltip_window.geometry(f"+{x}+{y}")
        except Exception:
            hide_tooltip()

    def hide_tooltip(event=None):
        nonlocal tooltip_window
        if tooltip_window:
            try:
                tooltip_window.destroy()
            except Exception:
                pass
            tooltip_window = None

    def move_tooltip(event):
        nonlocal tooltip_window
        if not tooltip_window:
            return
        try:
            x = event.x_root + 15
            y = event.y_root - 50
            tooltip_window.geometry(f"+{x}+{y}")
        except Exception:
            pass

    for i, path in enumerate(image_paths):
        var = tk.BooleanVar(value=True)
        check_vars.append(var)
        frame_item = tk.Frame(inner_frame, bg="white", relief="flat")
        frame_item.pack(fill="x", padx=2, pady=1)

        try:
            img = PILImage.open(path)
            img.thumbnail((30, 30), PILImage.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            thumbnails.append(photo)
            img_label = tk.Label(frame_item, image=photo, bg="white")
            img_label.pack(side="left", padx=(5, 5))
        except Exception:
            spacer = tk.Label(frame_item, text="✕", width=2, bg="white")
            spacer.pack(side="left", padx=(5, 5))

        cb = tk.Checkbutton(
            frame_item,
            variable=var,
            text=f"{i+1}. {path.name}",
            anchor="w",
            font=("Segoe UI", 9),
            bg="white",
        )
        cb.pack(side="left", fill="x", expand=True)

        def _row_enter(e, f=frame_item, p=path):
            f.config(bg="#f0f0f0")
            show_tooltip(e, p)

        def _row_motion(e):
            move_tooltip(e)

        def _row_leave(e, f=frame_item):
            f.config(bg="white")
            hide_tooltip()

        frame_item.bind("<Enter>", _row_enter)
        frame_item.bind("<Motion>", _row_motion)
        frame_item.bind("<Leave>", _row_leave)
        for child in frame_item.winfo_children():
            child.bind("<Enter>", _row_enter)
            child.bind("<Motion>", _row_motion)
            child.bind("<Leave>", _row_leave)

    def select_all():
        for var in check_vars:
            var.set(True)

    def deselect_all():
        for var in check_vars:
            var.set(False)

    def invert_selection():
        for var in check_vars:
            var.set(not var.get())

    tk.Button(btn_frame, text=T("✓ Vybrat vše", "✓ Alles wählen"), command=select_all, width=12).pack(side="left", padx=2)
    tk.Button(btn_frame, text=T("✗ Zrušit vše", "✗ Alles abwählen"), command=deselect_all, width=12).pack(side="left", padx=2)
    tk.Button(btn_frame, text=T("⇄ Invertovat", "⇄ Invertieren"), command=invert_selection, width=12).pack(side="left", padx=2)

    def update_scroll(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(canvas_window, width=canvas.winfo_width())

    inner_frame.bind("<Configure>", update_scroll)
    canvas.bind("<Configure>", update_scroll)

    def on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", on_mousewheel)
    result = {"paths": None}

    def confirm():
        hide_tooltip()
        result["paths"] = [path for i, path in enumerate(image_paths) if check_vars[i].get()]
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    def cancel():
        hide_tooltip()
        result["paths"] = None
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    bottom_frame = tk.Frame(win, padx=10, pady=10)
    bottom_frame.pack(fill="x")
    info_label = tk.Label(bottom_frame, text="", fg="blue")
    info_label.pack(side="left")

    def update_count(*args):
        count = sum(1 for var in check_vars if var.get())
        info_label.config(
            text=T(f"Vybráno: {count} z {len(image_paths)}", f"Ausgewählt: {count} von {len(image_paths)}")
        )

    for var in check_vars:
        var.trace_add("write", update_count)
    update_count()

    tk.Button(
        bottom_frame,
        text=T("✓ Export s vybranými", "✓ Export mit Auswahl"),
        command=confirm,
        bg="#4CAF50",
        fg="white",
        width=18,
        font=("Segoe UI", 10, "bold"),
    ).pack(side="right", padx=5)
    tk.Button(bottom_frame, text=T("✗ Zrušit", "✗ Abbrechen"), command=cancel, width=12).pack(side="right")

    win.bind("<Return>", lambda e: confirm())
    win.bind("<Escape>", lambda e: cancel())
    win.protocol("WM_DELETE_WINDOW", lambda: (hide_tooltip(), cancel()))

    parent.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - 650) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - 500) // 2
    win.geometry(f"+{max(0, x)}+{max(0, y)}")

    win.wait_window()
    return result["paths"]


def vyber_fotky_dialog_bez_miniatur(parent, image_paths: list):
    if not image_paths:
        return []

    win = tk.Toplevel(parent)
    win.title(T("Výběr fotek pro PDF export", "Fotos für PDF-Export auswählen"))
    win.geometry("600x500")
    win.transient(parent)
    win.grab_set()
    win.resizable(True, True)

    header = tk.Frame(win, padx=10, pady=10)
    header.pack(fill="x")
    tk.Label(
        header,
        text=T("Vyberte fotky pro export do PDF:", "Wählen Sie Fotos für PDF-Export:"),
        font=("Segoe UI", 11, "bold"),
    ).pack(anchor="w")
    tk.Label(header, text=T(f"Celkem nalezeno: {len(image_paths)} fotek", f"Insgesamt gefunden: {len(image_paths)} Fotos"), fg="gray").pack(anchor="w")

    btn_frame = tk.Frame(win, padx=10, pady=5)
    btn_frame.pack(fill="x")
    list_frame = tk.Frame(win)
    list_frame.pack(fill="both", expand=True, padx=10, pady=5)
    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")
    canvas = tk.Canvas(list_frame, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=canvas.yview)
    inner_frame = tk.Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    check_vars = []
    for i, path in enumerate(image_paths):
        var = tk.BooleanVar(value=True)
        check_vars.append(var)
        frame_item = tk.Frame(inner_frame)
        frame_item.pack(fill="x", padx=5, pady=2)
        cb = tk.Checkbutton(frame_item, variable=var, text=f"{i+1}. {path.name}", anchor="w", font=("Segoe UI", 9))
        cb.pack(side="left", fill="x", expand=True)

    def select_all():
        for var in check_vars:
            var.set(True)

    def deselect_all():
        for var in check_vars:
            var.set(False)

    def invert_selection():
        for var in check_vars:
            var.set(not var.get())

    tk.Button(btn_frame, text=T("✓ Vybrat vše", "✓ Alles wählen"), command=select_all, width=12).pack(side="left", padx=2)
    tk.Button(btn_frame, text=T("✗ Zrušit vše", "✗ Alles abwählen"), command=deselect_all, width=12).pack(side="left", padx=2)
    tk.Button(btn_frame, text=T("⇄ Invertovat", "⇄ Invertieren"), command=invert_selection, width=12).pack(side="left", padx=2)

    def update_scroll(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(canvas_window, width=canvas.winfo_width())

    inner_frame.bind("<Configure>", update_scroll)
    canvas.bind("<Configure>", update_scroll)
    result = {"paths": None}

    def confirm():
        result["paths"] = [path for i, path in enumerate(image_paths) if check_vars[i].get()]
        win.destroy()

    def cancel():
        result["paths"] = None
        win.destroy()

    bottom_frame = tk.Frame(win, padx=10, pady=10)
    bottom_frame.pack(fill="x")
    info_label = tk.Label(bottom_frame, text="", fg="blue")
    info_label.pack(side="left")

    def update_count(*args):
        count = sum(1 for var in check_vars if var.get())
        info_label.config(
            text=T(f"Vybráno: {count} z {len(image_paths)}", f"Ausgewählt: {count} von {len(image_paths)}")
        )

    for var in check_vars:
        var.trace_add("write", update_count)
    update_count()

    tk.Button(
        bottom_frame,
        text=T("✓ Export s vybranými", "✓ Export mit Auswahl"),
        command=confirm,
        bg="#4CAF50",
        fg="white",
        width=18,
        font=("Segoe UI", 10, "bold"),
    ).pack(side="right", padx=5)
    tk.Button(bottom_frame, text=T("✗ Zrušit", "✗ Abbrechen"), command=cancel, width=12).pack(side="right")

    win.bind("<Return>", lambda e: confirm())
    win.bind("<Escape>", lambda e: cancel())

    parent.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - 600) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - 500) // 2
    win.geometry(f"+{max(0, x)}+{max(0, y)}")

    win.wait_window()
    return result["paths"]


def export_poruchy_pdf(parent, cislo: str, stroje: dict, alarm_filter: str | None = None, reseni_filter: str | None = None, selected_ids: list[str] | None = None):
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
        from reportlab.lib import colors
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.units import mm
        from reportlab.lib.utils import ImageReader
    except Exception:
        messagebox.showerror(
            T("Export PDF", "PDF-Export"),
            T(
                "Chybí knihovna reportlab.\nNainstaluj: py -m pip install reportlab",
                "Reportlab-Bibliothek fehlt.\nInstallieren: py -m pip install reportlab",
            ),
            parent=parent,
        )
        return

    def _img_fit(path: str, max_w, max_h):
        ir = ImageReader(path)
        iw, ih = ir.getSize()
        if not iw or not ih:
            return Image(path, width=max_w, height=max_h)
        scale = min(max_w / iw, max_h / ih)
        return Image(path, width=iw * scale, height=ih * scale)

    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "C:/Windows/Fonts/DejaVuSans.ttf"))
        font_name = "DejaVu"
    except Exception:
        pdfmetrics.registerFont(TTFont("Arial", "C:/Windows/Fonts/arial.ttf"))
        font_name = "Arial"

    cislo = str(cislo)
    vse = nacti_poruchy()
    poruchy = [p for p in vse if str(p.get("cislo")) == cislo]
    if selected_ids is not None:
        selected_set = {str(i) for i in selected_ids}
        poruchy = [p for p in poruchy if str(p.get("id", "")) in selected_set]
    if alarm_filter:
        poruchy = [p for p in poruchy if (p.get("alarm") or "").strip() == alarm_filter]
    if reseni_filter:
        poruchy = [p for p in poruchy if (p.get("reseni") or "").strip() == reseni_filter]

    if not poruchy:
        if alarm_filter or reseni_filter:
            parts = []
            if alarm_filter:
                parts.append(f"{T('alarm', 'Alarm')}: {alarm_filter}")
            if reseni_filter:
                parts.append(f"{T('reseni', 'Lösung')}: {reseni_filter}")
            filtr_txt = " | ".join(parts)
            messagebox.showinfo(
                T("Export PDF", "PDF-Export"),
                T(
                    f"Stroj {cislo} nema pro zvoleny filtr zadne poruchy k exportu.\nFiltr: {filtr_txt}",
                    f"Maschine {cislo} hat fur den gewahlten Filter keine Störungen zum Export.\nFilter: {filtr_txt}",
                ),
                parent=parent,
            )
            return
        messagebox.showinfo(
            T("Export PDF", "PDF-Export"),
            T(f"Stroj {cislo} nemá žádné poruchy k exportu.", f"Maschine {cislo} hat keine Störungen zum Export."),
            parent=parent,
        )
        return

    stroj = stroje.get(cislo, {})
    vyrobce = stroj.get("vyrobce", "")
    typ_stroje = stroj.get("typ", "")

    fname = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".pdf",
        filetypes=[(T("PDF soubor", "PDF-Datei"), "*.pdf")],
        initialfile=f"poruchy_stroj_{cislo}.pdf",
    )
    if not fname:
        return

    styles = getSampleStyleSheet()
    base_style = ParagraphStyle("Base", parent=styles["Normal"], fontName=font_name, fontSize=9, leading=11)
    title_style = ParagraphStyle("Title", parent=base_style, fontSize=16, leading=18, spaceAfter=4 * mm)
    sub_style = ParagraphStyle("Sub", parent=base_style, fontSize=11, leading=13, spaceAfter=6 * mm)
    header_style = ParagraphStyle("Header", parent=base_style, fontSize=9, leading=11)

    doc = SimpleDocTemplate(
        fname,
        pagesize=landscape(A4),
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )

    story = []
    logo_path = os.path.join(os.path.dirname(__file__), "sgm_logo_metal.png")
    if os.path.exists(logo_path):
        try:
            img = Image(logo_path, width=40 * mm, height=15 * mm)
            header_table = Table(
                [[Paragraph(f"<b>SGM-Wartung — {T('Poruchy stroje', 'Störungen Maschine')} {cislo}</b>", title_style), img]],
                colWidths=[160 * mm, 40 * mm],
            )
            header_table.setStyle(
                TableStyle(
                    [
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                        ("TOPPADDING", (0, 0), (-1, -1), 0),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                    ]
                )
            )
            story.append(header_table)
        except Exception:
            story.append(Paragraph(f"<b>SGM-Wartung — {T('Poruchy stroje', 'Störungen Maschine')} {cislo}</b>", title_style))
    else:
        story.append(Paragraph(f"<b>SGM-Wartung — {T('Poruchy stroje', 'Störungen Maschine')} {cislo}</b>", title_style))

    story.append(Paragraph(f"{T('Výrobce', 'Hersteller')}: {vyrobce}&nbsp;&nbsp;&nbsp;&nbsp;{T('Typ', 'Typ')}: {typ_stroje}", sub_style))
    story.append(Spacer(0, 4 * mm))

    header = ["ID", T("Datum", "Datum"), T("Kat", "Kat"), T("Alarm", "Alarm"), T("Popis / Řešení", "Beschr. / Lösung"), T("Operátor", "Operator")]
    data = [[Paragraph(h, header_style) for h in header]]

    for p in poruchy:
        lines = []
        popis = p.get("popis", "") or ""
        reseni = p.get("reseni", "") or ""
        if popis.strip():
            lines.append(f"{T('Popis', 'Beschreibung')}: {popis.strip()}")
        if reseni.strip():
            lines.append(f"{T('Řešení', 'Lösung')}: {reseni.strip()}")
        combo_text = "<br/>".join(lines) if lines else ""
        data.append(
            [
                Paragraph(str(p.get("id", "")), base_style),
                Paragraph(p.get("cas", "") or "", base_style),
                Paragraph(kat_ui(p.get("kategorie", "")), base_style),
                Paragraph(p.get("alarm", "") or "", base_style),
                Paragraph(combo_text, base_style),
                Paragraph(p.get("operator_uzavrel", "") or "", base_style),
            ]
        )

    table = Table(
        data,
        colWidths=[15 * mm, 35 * mm, 25 * mm, 35 * mm, 110 * mm, 30 * mm],
        repeatRows=1,
    )
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (0, -1), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(table)

    foto_root = None
    base = Path(slozka_stroje(cislo))
    if base.is_dir():
        foto_root = base
    if foto_root is None:
        base = Path(DATA_DIR) / "soubory" / str(cislo)
        if (base / "Fotodokumentace").is_dir():
            foto_root = base / "Fotodokumentace"
        elif base.is_dir():
            foto_root = base

    image_paths = []
    if foto_root and foto_root.is_dir():
        for ext in ("*.jpg", "*.jpeg", "*.png", "*.bmp"):
            image_paths.extend(foto_root.rglob(ext))
        image_paths = sorted(image_paths)

    if image_paths:
        selected_paths = vyber_fotky_dialog(parent, image_paths)
        if selected_paths is None:
            selected_paths = []
        if selected_paths:
            story.append(PageBreak())
            story.append(Paragraph(f"<b>{T('Fotodokumentace', 'Fotodokumentation')}</b>", title_style))
            story.append(Spacer(0, 4 * mm))
            max_w = doc.width
            max_h = doc.height - (20 * mm)
            for path in selected_paths:
                fname_img = os.path.basename(str(path))
                story.append(Paragraph(fname_img, base_style))
                story.append(Spacer(0, 3 * mm))
                try:
                    story.append(_img_fit(str(path), max_w, max_h))
                except Exception:
                    story.append(Paragraph(T("Chyba načtení obrázku:", "Fehler beim Laden des Bildes:") + f" {fname_img}", base_style))
                story.append(PageBreak())

    try:
        doc.build(story)
    except PermissionError:
        messagebox.showerror(
            T("Export PDF", "PDF-Export"),
            T(
                "Soubor se nepodařilo zapsat.\nJe možné, že je otevřený v prohlížeči.",
                "Datei konnte nicht geschrieben werden.\nMöglicherweise ist sie im Browser geöffnet.",
            ),
            parent=parent,
        )
        return

    messagebox.showinfo(
        T("Export PDF", "PDF-Export"),
        T(f"PDF bylo uloženo jako:\n{fname}", f"PDF wurde gespeichert als:\n{fname}"),
        parent=parent,
    )


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _sha256_zip_member(archive: zipfile.ZipFile, member: str) -> str:
    digest = hashlib.sha256()
    with archive.open(member) as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _backup_files(data_dir: Path):
    for path in sorted(data_dir.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(data_dir)
        if relative.parts and relative.parts[0].lower() == "logs":
            continue
        if path.name.endswith(".tmp"):
            continue
        yield path, relative


def create_backup_archive(destination: Path, data_dir: Path = DATA_DIR) -> Path:
    """Vytvoří úplnou ověřitelnou ZIP zálohu uživatelských dat."""
    destination = Path(destination)
    data_dir = Path(data_dir)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination_resolved = destination.resolve()
    files = [
        (path, relative)
        for path, relative in _backup_files(data_dir)
        if path.resolve() != destination_resolved
    ]
    manifest = {
        "format": BACKUP_FORMAT_VERSION,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "files": {
            relative.as_posix(): {
                "size": path.stat().st_size,
                "sha256": _sha256_file(path),
            }
            for path, relative in files
        },
    }

    temp_path = destination.with_name(f".{destination.name}.tmp")
    try:
        with zipfile.ZipFile(
            temp_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
        ) as archive:
            for path, relative in files:
                archive.write(path, arcname=f"data/{relative.as_posix()}")
            archive.writestr(
                BACKUP_MANIFEST,
                json.dumps(manifest, ensure_ascii=False, indent=2),
            )
        inspect_backup_archive(temp_path, validate_data=False)
        os.replace(temp_path, destination)
        return destination
    except Exception:
        LOGGER.exception("Vytvoření zálohy selhalo: %s", destination)
        raise
    finally:
        temp_path.unlink(missing_ok=True)


def _safe_zip_name(name: str) -> PurePosixPath:
    if "\\" in name:
        raise ValueError(f"Neplatná cesta v ZIP: {name}")
    path = PurePosixPath(name)
    if path.is_absolute() or ".." in path.parts:
        raise ValueError(f"Neplatná cesta v ZIP: {name}")
    if path.parts and ":" in path.parts[0]:
        raise ValueError(f"Neplatná cesta v ZIP: {name}")
    return path


def _validate_csv_bytes(content: bytes, required: set[str], name: str) -> None:
    try:
        text = content.decode("utf-8-sig")
        header = next(csv.reader(io.StringIO(text)), [])
    except (UnicodeDecodeError, csv.Error) as exc:
        raise ValueError(f"Soubor {name} není platné UTF-8 CSV.") from exc
    missing = required.difference(header)
    if missing:
        raise ValueError(
            f"Soubor {name} nemá povinné sloupce: {', '.join(sorted(missing))}"
        )


def inspect_backup_archive(archive_path: Path, *, validate_data: bool = True) -> dict:
    """Ověří bezpečnost, CSV strukturu a kontrolní součty zálohy."""
    archive_path = Path(archive_path)
    with zipfile.ZipFile(archive_path) as archive:
        infos = [info for info in archive.infolist() if not info.is_dir()]
        if len(infos) > 100_000:
            raise ValueError("Záloha obsahuje nepřiměřeně mnoho souborů.")
        total_size = sum(info.file_size for info in infos)
        if total_size > 20 * 1024**3:
            raise ValueError("Rozbalená záloha je větší než 20 GB.")
        for info in infos:
            _safe_zip_name(info.filename)
        broken = archive.testzip()
        if broken:
            raise ValueError(f"ZIP je poškozený: {broken}")

        names = {info.filename for info in infos}
        if len(names) != len(infos):
            raise ValueError("Záloha obsahuje duplicitní názvy souborů.")
        is_current = BACKUP_MANIFEST in names
        if is_current:
            try:
                manifest = json.loads(archive.read(BACKUP_MANIFEST))
            except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                raise ValueError("Manifest zálohy je poškozený.") from exc
            if manifest.get("format") != BACKUP_FORMAT_VERSION:
                raise ValueError("Nepodporovaná verze zálohy.")
            manifest_files = manifest.get("files")
            if not isinstance(manifest_files, dict):
                raise ValueError("Manifest neobsahuje seznam souborů.")
            for relative, metadata in manifest_files.items():
                if not isinstance(relative, str) or not isinstance(metadata, dict):
                    raise ValueError("Manifest obsahuje neplatný záznam souboru.")
                _safe_zip_name(relative)
                member = f"data/{relative}"
                if member not in names:
                    raise ValueError(f"V záloze chybí soubor: {relative}")
                if archive.getinfo(member).file_size != metadata.get("size"):
                    raise ValueError(f"Nesouhlasí velikost souboru: {relative}")
                checksum = _sha256_zip_member(archive, member)
                if checksum != metadata.get("sha256"):
                    raise ValueError(f"Nesouhlasí kontrolní součet: {relative}")
            entries = {
                relative: f"data/{relative}" for relative in manifest_files.keys()
            }
            archive_type = "complete"
        else:
            legacy = ("stroje.csv", "poruchy.csv", "sablony_alarmu.csv")
            entries = {name: name for name in legacy if name in names}
            archive_type = "legacy"

        if validate_data:
            for required_name in ("stroje.csv", "poruchy.csv"):
                if required_name not in entries:
                    raise ValueError(
                        f"V záloze chybí povinný soubor {required_name}."
                    )
            _validate_csv_bytes(
                archive.read(entries["stroje.csv"]), {"cislo"}, "stroje.csv"
            )
            _validate_csv_bytes(
                archive.read(entries["poruchy.csv"]),
                {"id", "cislo", "stav"},
                "poruchy.csv",
            )
        return {
            "type": archive_type,
            "entries": entries,
            "file_count": len(entries),
            "total_size": sum(
                archive.getinfo(member).file_size for member in entries.values()
            ),
        }


def _copy_file_atomic(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = destination.parent / f".{destination.name}.restore.tmp"
    try:
        shutil.copy2(source, temp_path)
        os.replace(temp_path, destination)
    finally:
        temp_path.unlink(missing_ok=True)


def restore_backup_archive(
    archive_path: Path,
    data_dir: Path = DATA_DIR,
    safety_backup_dir: Path | None = None,
) -> Path:
    """Obnoví soubory ze ZIP a při chybě vrátí již změněné soubory zpět."""
    archive_path = Path(archive_path)
    data_dir = Path(data_dir)
    info = inspect_backup_archive(archive_path)
    safety_backup_dir = (
        Path(safety_backup_dir)
        if safety_backup_dir is not None
        else data_dir.parent / "backups"
    )
    safety_backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    safety_path = safety_backup_dir / f"pred_obnovou_{timestamp}.zip"
    create_backup_archive(safety_path, data_dir)

    stage_root = data_dir.parent / ".sgm_restore_work"
    if stage_root.exists():
        shutil.rmtree(stage_root)
    stage_root.mkdir(parents=True)
    rollback_root = stage_root / "rollback"
    incoming_root = stage_root / "incoming"
    changed = []
    try:
        with zipfile.ZipFile(archive_path) as archive:
            for relative_text, member in info["entries"].items():
                relative = Path(*PurePosixPath(relative_text).parts)
                incoming = incoming_root / relative
                incoming.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, open(incoming, "wb") as target:
                    shutil.copyfileobj(source, target)

        _validate_csv_bytes(
            (incoming_root / "stroje.csv").read_bytes(), {"cislo"}, "stroje.csv"
        )
        _validate_csv_bytes(
            (incoming_root / "poruchy.csv").read_bytes(),
            {"id", "cislo", "stav"},
            "poruchy.csv",
        )

        for relative_text in info["entries"]:
            relative = Path(*PurePosixPath(relative_text).parts)
            destination = data_dir / relative
            rollback = rollback_root / relative
            existed = destination.exists()
            if existed:
                rollback.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(destination, rollback)
            changed.append((destination, rollback, existed))
            _copy_file_atomic(incoming_root / relative, destination)
        LOGGER.info(
            "Data obnovena z %s; bezpečnostní záloha: %s",
            archive_path,
            safety_path,
        )
        return safety_path
    except Exception:
        LOGGER.exception("Obnova zálohy selhala, probíhá návrat změn")
        for destination, rollback, existed in reversed(changed):
            try:
                if existed:
                    _copy_file_atomic(rollback, destination)
                else:
                    destination.unlink(missing_ok=True)
            except Exception:
                LOGGER.critical(
                    "Návrat souboru po chybě obnovy selhal: %s",
                    destination,
                    exc_info=True,
                )
        raise
    finally:
        shutil.rmtree(stage_root, ignore_errors=True)


def backup_zip(parent=None):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".zip",
        initialfile=f"sgm_backup_{ts}.zip",
    )
    if not fname:
        return False
    try:
        create_backup_archive(Path(fname))
    except Exception as exc:
        messagebox.showerror(
            T("Záloha", "Sicherung"),
            T(
                f"Zálohu se nepodařilo vytvořit:\n{exc}",
                f"Sicherung konnte nicht erstellt werden:\n{exc}",
            ),
            parent=parent,
        )
        return False
    messagebox.showinfo(
        T("Záloha", "Sicherung"),
        f"{T('Uloženo', 'Gespeichert')} {fname}",
        parent=parent,
    )
    return True


def restore_zip(parent=None):
    fname = filedialog.askopenfilename(parent=parent, filetypes=[("ZIP", "*.zip")])
    if not fname:
        return False
    try:
        info = inspect_backup_archive(Path(fname))
    except Exception as exc:
        LOGGER.warning("Odmítnuta neplatná záloha %s", fname, exc_info=True)
        messagebox.showerror(
            T("Obnova", "Wiederherstellung"),
            T(
                f"Záloha není platná a nebude obnovena:\n{exc}",
                f"Die Sicherung ist ungültig und wird nicht wiederhergestellt:\n{exc}",
            ),
            parent=parent,
        )
        return False

    kind = (
        T("úplná", "vollständig")
        if info["type"] == "complete"
        else T("starší – pouze CSV", "älter – nur CSV")
    )
    if not messagebox.askyesno(
        T("Potvrdit obnovu", "Wiederherstellung bestätigen"),
        T(
            f"Typ zálohy: {kind}\nPočet souborů: {info['file_count']}\n\n"
            "Před obnovou se automaticky uloží úplná kopie současných dat. "
            "Pokračovat?",
            f"Sicherungstyp: {kind}\nDateien: {info['file_count']}\n\n"
            "Vor der Wiederherstellung wird automatisch eine vollständige Kopie "
            "der aktuellen Daten erstellt. Fortfahren?",
        ),
        parent=parent,
    ):
        return False

    try:
        safety_path = restore_backup_archive(Path(fname))
    except Exception as exc:
        messagebox.showerror(
            T("Obnova", "Wiederherstellung"),
            T(
                f"Obnova selhala; provedené změny byly vráceny:\n{exc}",
                f"Wiederherstellung fehlgeschlagen; Änderungen wurden "
                f"zurückgesetzt:\n{exc}",
            ),
            parent=parent,
        )
        return False
    messagebox.showinfo(
        T("Obnova", "Wiederherstellung"),
        T(
            f"Data obnovena.\nNávratová záloha:\n{safety_path}",
            f"Daten wiederhergestellt.\nRücksicherung:\n{safety_path}",
        ),
        parent=parent,
    )
    return True


def export_wartung_csv(parent):
    stroje = nacti_stroje()

    mode_var = getattr(parent, "wartung_mode", None)
    mode = mode_var.get() if mode_var is not None else T("≤ 30 dní", "≤ 30 Tage")

    rows = []
    for cislo, s in stroje.items():
        dny = days_to_next_wartung(s)
        if dny is None:
            continue

        if mode == T("prošlé", "überfällig"):
            if dny > 0:
                continue
        elif mode == T("≤ 30 dní", "≤ 30 Tage"):
            if dny > 30:
                continue
        elif mode == T("vše s údržbou", "Alle mit Wartung"):
            pass
        else:
            if dny > 30:
                continue

        if dny <= 0:
            stav = T("prošlá", "überfällig")
            status_ikona = "🔴 PROŠLÉ"
        elif dny == 1:
            stav = T("za 1 den", "in 1 Tag")
            status_ikona = "🟡 BRZY"
        else:
            stav = f"{T('za', 'in')} {dny} {T('dní', 'Tagen')}"
            status_ikona = "🟡 BRZY" if dny <= 30 else "🟢 OK"

        rows.append(
            {
                "cislo": cislo,
                "vyrobce": s.get("vyrobce", ""),
                "typ": s.get("typ", ""),
                "rok": s.get("rok", ""),
                "spm": s.get("spm", ""),
                "seriove": s.get("seriove", ""),
                "wartung_last": s.get("wartung_last", ""),
                "dny_do_wartung": dny,
                "wartung_stav": stav,
                "status_ikona": status_ikona,
            }
        )

    if not rows:
        messagebox.showinfo(
            T("Údržba", "Wartung"),
            T(
                "Není žádný stroj odpovídající zvolenému filtru.",
                "Keine Maschine entspricht dem gewählten Filter.",
            ),
            parent=parent,
        )
        return False

    rows.sort(key=lambda r: r["dny_do_wartung"])

    fname = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv")],
        initialfile="SGM_Wartung_seznam.csv",
        title=T("Uložit seznam údržby", "Wartungsliste speichern"),
    )
    if not fname:
        return False

    fieldnames = [
        "cislo",
        "vyrobce",
        "typ",
        "rok",
        "spm",
        "seriove",
        "wartung_last",
        "dny_do_wartung",
        "wartung_stav",
        "status_ikona",
    ]

    with open(fname, "w", newline="", encoding="utf-8") as f:
        import csv

        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

    messagebox.showinfo(
        T("Údržba", "Wartung"),
        f"{T('Seznam strojů pro údržbu byl uložen do', 'Liste der Maschinen für Wartung gespeichert in')}:\n{fname}",
        parent=parent,
    )
    return True
