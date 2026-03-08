#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from data_manager import (
    DATA_DIR,
    SOUBOR_PORUCHY,
    SOUBOR_SABLONY,
    SOUBOR_STROJE,
    days_to_next_wartung,
    kat_ui,
    nacti_stroje,
    nacti_poruchy,
    slozka_stroje,
)


LANG = os.environ.get("SGM_LANG", "de").strip().lower()


def T(cz: str, de: str | None = None) -> str:
    if LANG.upper() == "DE":
        return de if de is not None else cz
    return cz


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


def export_poruchy_pdf(parent, cislo: str, stroje: dict):
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

    if not poruchy:
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


def backup_zip(parent=None):
    ts = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".zip",
        initialfile=f"sgm_backup_{ts}.zip",
    )
    if not fname:
        return False
    with zipfile.ZipFile(fname, "w") as z:
        for p in [SOUBOR_STROJE, SOUBOR_PORUCHY, SOUBOR_SABLONY]:
            if p.exists():
                z.write(p, arcname=p.name)
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
    with zipfile.ZipFile(fname) as z:
        for name in ["stroje.csv", "poruchy.csv", "sablony_alarmu.csv"]:
            if name in z.namelist():
                with z.open(name) as src, open(DATA_DIR / name, "wb") as dst:
                    dst.write(src.read())
    messagebox.showinfo(
        T("Obnova", "Wiederherstellung"),
        T("Data obnovena.", "Daten wiederhergestellt."),
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
        elif mode == T("vše s Wartung", "Alle mit Wartung"):
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
            T("Wartung", "Wartung"),
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
        title=T("Uložit seznam Wartung", "Wartungsliste speichern"),
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
        T("Wartung", "Wartung"),
        f"{T('Seznam strojů pro Wartung byl uložen do', 'Liste der Maschinen für Wartung gespeichert in')}:\n{fname}",
        parent=parent,
    )
    return True
