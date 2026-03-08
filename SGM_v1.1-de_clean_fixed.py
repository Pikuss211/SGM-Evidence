#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from datetime import datetime
from tkinter import simpledialog
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import csv
import zipfile
import re
from pathlib import Path
from datetime import date, datetime, timedelta
from pathlib import Path
import data_manager as dm
import export_manager as em

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

SOUBOR_STROJE = DATA_DIR / "stroje.csv"
SOUBOR_PORUCHY = DATA_DIR / "poruchy.csv"
SOUBOR_SABLONY = DATA_DIR / "sablony_alarmu.csv"

# --- Přepínač emoji v UI (0 = vypnout emoji, 1 = zapnout) ---
USE_EMOJI = os.environ.get("SGM_USE_EMOJI", "1") != "0"


def UI(emoji_text: str, plain_text: str) -> str:
    return emoji_text if USE_EMOJI else plain_text
# --------------------------------------------------------------
# ===== SORT MODE (interní klíče vs. UI text) =====


# --- Jazyk UI (CZ/DE) ---
# Interní klíče (cislo, stav, kategorie, ...) se NIKDY nepřekládají.
LANG = os.environ.get("SGM_LANG", "de").strip().lower()


def T(cz: str, de: str | None = None) -> str:
    """Překlad pro UI. Když DE chybí, použije se CZ."""
    if LANG.upper() == "DE":
        return de if de is not None else cz
    return cz


STAV_LABELS = {
    "bezi":     T("běží", "läuft"),
    "porucha":  T("porucha", "Störung"),
}


def UIT(emoji: str, cz: str) -> str:
    """Text pro tlačítka: pokud jsou emoji zapnuté, přidá emoji + přeložený text."""
    if USE_EMOJI:
        return f"{emoji} {T(cz)}".strip()
    return T(cz)


SORT_LABELS = {
    "cislo":        T("Číslo", "Nr."),
    "otevrene_desc": T("Otevřené ↓", "Offen ↓"),
    "poruchy_30d":  T("Poruchy 30d", "Stör. 30T"),
    "poruchy_all":  T("Poruchy celkem", "Stör. ges."),
    "last_open":    T("Posl. otevřená", "Letzte offen"),
}

SORT_KEYS = {v: k for k, v in SORT_LABELS.items()}

TILE_FIELD_LABELS = {
    "cislo_only": T("Číslo", "Nr."),
    "vyrobce":    T("Výrobce", "Hersteller"),
    "rok":        T("Rok", "Jahr"),
    "spm":        "SPM",
    "seriove":    "S/N",
}

TILE_FIELD_BY_LABEL = {v: k for k, v in TILE_FIELD_LABELS.items()}


# --- Správná cesta k datům (funguje i pro .exe z PyInstalleru) ---
if getattr(sys, "frozen", False):  # běží jako EXE
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

SOUBOR_STROJE = DATA_DIR / "stroje.csv"

# --- Složky pro soubory strojů (fotky, dokumenty) ---


def otevrit_slozku(cesta: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(cesta)  # type: ignore
        elif sys.platform == "darwin":
            os.system(f'open "{cesta}"')
        else:
            os.system(f'xdg-open "{cesta}"')
    except Exception as e:
        messagebox.showerror(
            T("Chyba", "Fehler"), f"{T('Nepodařilo se otevřít složku', 'Ordner konnte nicht geöffnet werden')}:\n{e}")


SOUBOR_PORUCHY = DATA_DIR / "poruchy.csv"
SOUBOR_SABLONY = DATA_DIR / "sablony_alarmu.csv"
# --------------------------------------------------------------

# Pomocné funkce pro práci se soubory


def ask_kategorie_combobox(parent) -> str | None:
    """
    Otevře modální dialog s Comboboxem a vrátí
    'elektricka' / 'mechanicka' / 'jina' nebo None při zrušení.
    """
    # mapování label -> interní hodnota
    LABEL2VAL = {
        T("Elektrická", "Elektrisch"): "elektricka",
        T("Mechanická", "Mechanisch"): "mechanicka",
        T("Jiná", "Sonstige"): "jina",
    }

    win = tk.Toplevel(parent)
    win.title(T("Nová porucha – kategorie", "Neue Störung – Kategorie"))
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = tk.Frame(win, padx=12, pady=10)
    frm.pack(fill="both", expand=True)

    tk.Label(frm, text=T("Kategorie poruchy:", "Störungskategorie:")).grid(
        row=0, column=0, sticky="w")

    var = tk.StringVar()
    cb = ttk.Combobox(frm, textvariable=var, state="readonly",
                      values=[T("Elektrická", "Elektrisch"), T("Mechanická", "Mechanisch"), T("Jiná", "Sonstige")], width=22)
    cb.grid(row=1, column=0, sticky="ew", pady=(4, 10))
    cb.focus_set()

    result = {"val": None}

    def do_ok(_e=None):
        label = var.get().strip()
        if not label:
            messagebox.showwarning(
                T("Upozornění", "Hinweis"), T("Vyber kategorii.", "Wählen Sie eine Kategorie."), parent=win)
            return
        result["val"] = LABEL2VAL[label]
        win.destroy()

    def do_cancel(_e=None):
        result["val"] = None
        win.destroy()

    btns = tk.Frame(frm)
    btns.grid(row=2, column=0, sticky="e")
    tk.Button(btns, text="OK", width=10, command=do_ok).pack(
        side="left", padx=(0, 6))
    tk.Button(btns, text=T("Zrušit", "Abbrechen"), width=10,
              command=do_cancel).pack(side="left")

    win.bind("<Return>", do_ok)
    win.bind("<Escape>", do_cancel)

    # zarovnání k rodiči
    win.update_idletasks()
    px = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_width()) // 2
    py = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_height()) // 2
    win.geometry(f"+{max(0, px)}+{max(0, py)}")

    parent.wait_window(win)
    return result["val"]


def bulk_uzavrit_dialog(parent, poruchy: list) -> list[str] | None:
    opened = [p for p in poruchy if p.get("stav") == "otevrena"]
    if not opened:
        messagebox.showinfo(T("Hromadné uzavření", "Sammelschließung"),
                            T("Žádné otevřené poruchy.", "Keine offenen Störungen."), parent=parent)
        return None
    win = tk.Toplevel(parent)
    win.title(T("Hromadné uzavření poruch", "Störungen sammelweise schließen"))
    win.transient(parent)
    win.grab_set()
    frm = tk.Frame(win, padx=10, pady=10)
    frm.pack(fill="both", expand=True)
    tk.Label(
        frm, text=T("Vyber záznamy k uzavření (Ctrl/Shift pro více):", "Wählen Sie Einträge zum Schließen (Ctrl/Shift für mehrere):")).pack(anchor="w")
    lb = tk.Listbox(frm, selectmode="extended", width=110, height=18)
    items = []
    for p in opened:
        line = (
            f"ID {p.get('id')} | "
            f"{T('stroj', 'Maschine')} {p.get('cislo')} | "
            f"{p.get('cas')} | "
            f"{p.get('alarm')} | "
            f"{kat_ui(p.get('kategorie'))}"
            )
        items.append((p.get("id"), line))
        lb.insert("end", line)
    lb.pack(fill="both", expand=True, pady=6)
    res = {"ids": None}

    def do_ok():
        sel = lb.curselection()
        res["ids"] = [items[i][0] for i in sel] if sel else []
        win.destroy()

    def do_cancel(): res["ids"] = None; win.destroy()
    btns = tk.Frame(frm)
    btns.pack(anchor="e")
    tk.Button(btns, text=T("Uzavřít", "Schließen"), command=do_ok,
              width=12).pack(side="left", padx=(0, 6))
    tk.Button(btns, text=T("Zrušit", "Abbrechen"), command=do_cancel,
              width=12).pack(side="left")
    parent.wait_window(win)
    return res["ids"]


STAV_UI = {
    "bezi":    T("běží", "läuft"),
    "porucha": T("porucha", "Störung"),
}

STAV_UI_REV = {v: k for k, v in STAV_UI.items()}


# --- jednoduchý tooltip ---


def create_tooltip(widget, text: str):
    """
    Jednoduchý tooltip pro Tkinter widgety.

    Oprava proti náhodným TclErrorům při rychlém přejetí myší:
    - tooltip se vytváří se zpožděním
    - všechny winfo() výpočty jsou chráněné (widget/toplevel může mezitím zaniknout)
    """
    tip = {"w": None, "after": None}

    def _do_show():
        tip["after"] = None
        # widget už nemusí existovat (např. přepočet mřížky)
        try:
            if not widget.winfo_exists():
                return
        except Exception:
            return

        if tip["w"] is not None:
            return

        try:
            tw = tk.Toplevel(widget)
            tip["w"] = tw
            tw.wm_overrideredirect(True)
            tw.attributes("-topmost", True)

            lbl = tk.Label(
                tw,
                text=text,
                justify="left",
                background="#ffffe0",
                relief="solid",
                borderwidth=1,
                font=("Segoe UI", 9),
                padx=6,
                pady=4,
            )
            lbl.pack()

            # umístění poblíž kurzoru, ale uvnitř obrazovky
            try:
                x = widget.winfo_pointerx() + 12
                y = widget.winfo_pointery() + 12

                tw.update_idletasks()
                w = tw.winfo_width()
                h = tw.winfo_height()
                sw = tw.winfo_screenwidth()
                sh = tw.winfo_screenheight()

                if x + w > sw:
                    x = max(0, sw - w - 8)
                if y + h > sh:
                    y = max(0, sh - h - 8)

                tw.wm_geometry(f"+{x}+{y}")
            except tk.TclError:
                # někdo mezitím zavřel okno / widget zanikl
                _hide()
        except tk.TclError:
            # ochrana pro případy, kdy se Toplevel nestihne vytvořit
            tip["w"] = None
            return

    def _show(_event=None):
        # zruš případné předchozí naplánování
        try:
            if tip["after"] is not None:
                widget.after_cancel(tip["after"])
        except Exception:
            pass
        tip["after"] = None

        # naplánuj zpožděné zobrazení (anti-flicker)
        try:
            tip["after"] = widget.after(250, _do_show)
        except Exception:
            tip["after"] = None

    def _hide(_event=None):
        # zruš pending show
        try:
            if tip["after"] is not None:
                widget.after_cancel(tip["after"])
        except Exception:
            pass
        tip["after"] = None

        tw = tip.get("w")
        tip["w"] = None
        if tw is not None:
            try:
                if tw.winfo_exists():
                    tw.destroy()
            except Exception:
                pass

    widget.bind("<Enter>", _show)
    widget.bind("<Leave>", _hide)


def center_over(child: "tk.Toplevel", parent: "tk.Misc"):
    """Vycentruje child okno nad parent oknem (spolehlivé na Windows)."""
    child.update_idletasks()

    # rozměry child
    cw = child.winfo_width()
    ch = child.winfo_height()

    # pozice/rozměry parent
    px = parent.winfo_rootx()
    py = parent.winfo_rooty()
    pw = parent.winfo_width()
    ph = parent.winfo_height()

    x = px + max(0, (pw - cw) // 2)
    y = py + max(0, (ph - ch) // 2)

    child.geometry(f"+{x}+{y}")
    child.lift()
    try:
        child.focus_force()
    except Exception:
        pass


# ===== GUI Aplikace =====
BASE_DIR = dm.BASE_DIR
COLORS = dm.COLORS
DATA_DIR = dm.DATA_DIR
SOUBOR_PORUCHY = dm.SOUBOR_PORUCHY
SOUBOR_SABLONY = dm.SOUBOR_SABLONY
SOUBOR_STROJE = dm.SOUBOR_STROJE
_safe_int = dm._safe_int
barva_dlazdice = dm.barva_dlazdice
color_by_cat = dm.color_by_cat
days_to_next_wartung = dm.days_to_next_wartung
kat_ui = dm.kat_ui
last_open_dt = dm.last_open_dt
last_open_issue = dm.last_open_issue
nacti_poruchy = dm.nacti_poruchy
nacti_sablony = dm.nacti_sablony
nacti_stroje = dm.nacti_stroje
next_free_machine_number = dm.next_free_machine_number
normalize_dt = dm.normalize_dt
normalize_kategorie = dm.normalize_kategorie
normalize_stav = dm.normalize_stav
nove_id = dm.nove_id
porucha_stav_ui = dm.porucha_stav_ui
slozka_stroje = dm.slozka_stroje
stav_ui = dm.stav_ui
uloz_poruchy = dm.uloz_poruchy
uloz_stroje = dm.uloz_stroje
export_poruchy_pdf = em.export_poruchy_pdf
vyber_fotky_dialog = em.vyber_fotky_dialog
vyber_fotky_dialog_bez_miniatur = em.vyber_fotky_dialog_bez_miniatur

class StrojeGrid(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(T("SGM – přehled strojů", "SGM – Maschinenübersicht"))
        self.geometry("1200x740")
        self.configure(bg="#e9f2f7")
        self.sort_mode = tk.StringVar(value="cislo")
        self.tile_field = tk.StringVar(
            value="cislo_only")              # interní
        self.tile_field_ui = tk.StringVar(
            # UI tex)
            value=TILE_FIELD_LABELS["cislo_only"])

        # === DATA ===
        self.stroje = nacti_stroje()     # dict cislo -> dict
        # jméno operátora (start: Windows login)
        self.operator = os.getenv("USERNAME", "")
        self.poruchy = nacti_poruchy()    # list
        self.sablony = nacti_sablony()    # dict alarm -> reseni
        # --- STAV UI (musí být připraveno před tvorbou widgetů) ---
        self.status = tk.StringVar(value=T("Zadej číslo a Enter (nebo dvojklik). N = přidat stroj.",
                                           "Nummer eingeben und Enter (oder Doppelklick). N = Maschine hinzufügen."))
        self.filter_only_problem = tk.BooleanVar(
            value=False)  # jen stroje s otevřenou poruchou
        self.last_selected = None
        self.vstup_str = ""  # číselný vstup z klávesnice
        self._resize_after_id = None
        self._pending_columns = None
        self._last_draw_columns = None
        self._scrollregion_after_id = None
        self._wheel_after_id = None
        self._wheel_accum = 0
        self._wheel_remainder = 0.0

        # filtr kategorie (vše/elektrická/mechanická/jiná)
        self.filtr_kat = tk.StringVar(value="vse")

        # režim exportu Wartung
        self.wartung_mode = tk.StringVar(value=T("≤ 30 dní", "≤ 30 Tage"))

        # === HORNÍ LIŠTA ===
        top = tk.Frame(self, bg="#e9f2f7")
        top.pack(fill="x")

        # Levá část – hlavní akce
        tk.Button(top, text=UIT("💾", T("Zálohovat", "Sichern")),
                  command=self.backup_zip).pack(side="left", padx=(10, 0), pady=6)
        tk.Button(top, text=UIT("⤓", T("Obnovit", "Wiederh.")),
                  command=self.restore_zip).pack(side="left", padx=(6, 0), pady=6)
        tk.Button(top, text=UIT("🔎", T("Hledat poruchy", "Stör. suchen")),
                  command=self.global_search_gui).pack(side="left", padx=(10, 0), pady=6)
        tk.Button(top, text=T("Hromadně uzavřít", "Sammel-Schl."),
                  command=lambda: self.hromadne_uzavrit(self)).pack(side="left", padx=(10, 0), pady=6)
        tk.Button(top, text=UIT("⏯️", T("Přepnout stav", "Status")),
                  command=self.prepnout_stav_toolbar).pack(side="left", padx=(10, 0), pady=6)
        tk.Button(top, text=UIT("📈", T("Graf TOP stroje", "TOP-Graph")),
                  command=self.graf_top_stroje).pack(side="left", padx=(10, 0), pady=6)

        # Řazení
        tk.Label(top, text=T("Řadit:", "Sort:"), bg="#e9f2f7").pack(
            side="left", padx=(10, 0))

        self.sort_ui = tk.StringVar(value=T("číslo"))
        self.sort_ui.set(SORT_LABELS.get(self.sort_mode.get(), T("číslo")))

        self.sort_combo = ttk.Combobox(
            top,
            textvariable=self.sort_ui,
            state="readonly",
            values=list(SORT_KEYS.keys()),
            width=12
        )
        self.sort_combo.pack(side="left", padx=(4, 0))

        # Tooltip k roletce řazení (Sort)
        sort_tip = (
            T("Řazení dlaždic:", "Sortierung der Kacheln:") + "\n"
            + f"• {T('Číslo', 'Nr.')}: {T('podle čísla stroje', 'nach Maschinennummer')}\n"
            + f"• {T('Otevřené ↓', 'Offen ↓')}: {T('podle počtu otevřených poruch', 'nach Anzahl offener Störungen')}\n"
            + f"• {T('Poruchy 30d', 'Stör. 30T')}: {T('podle počtu poruch za 30 dní', 'nach Störungen der letzten 30 Tage')}\n"
            + f"• {T('Poruchy celkem', 'Stör. ges.')}: {T('podle počtu poruch celkem', 'nach Störungen gesamt')}\n"
            + f"• {T('Nejnovější otevřená', 'Letzte offen')}: {T('nejnovější otevřená porucha nahoře', 'letzte offene Störung oben')}"
        )
        # create_tooltip musí existovat (už ho v projektu máš)
        create_tooltip(self.sort_combo, sort_tip)

        def _on_sort_change(event=None):
            ui_value = self.sort_combo.get()
            key = SORT_KEYS.get(ui_value, "cislo")
            self.sort_mode.set(key)
            self.nakresli_mrizku()

        self.sort_combo.bind("<<ComboboxSelected>>", _on_sort_change)
        self.sort_combo.bind("<Return>", _on_sort_change)

        # ── Popisek dlaždice (vpravo) ────────────────────────────────────────────────
        rightgrp = tk.Frame(top, bg="#e9f2f7")
        rightgrp.pack(side="right", padx=(6, 10))

        tk.Label(rightgrp, text=T("Data Ma:", "Daten Ma:"),
                 bg="#e9f2f7").pack(side="left")

        # Popisek dlaždice: interně držíme klíče, v menu zobrazíme přeložené názvy
        pop = tk.OptionMenu(rightgrp, self.tile_field_ui, "")
        menu = pop["menu"]
        menu.delete(0, "end")

        for key, label in TILE_FIELD_LABELS.items():
            menu.add_command(
                label=label,
                command=lambda k=key, l=label: (
                    self.tile_field.set(k),       # interní klíč
                    self.tile_field_ui.set(l),    # UI text
                    self.nakresli_mrizku()
                )
            )

        pop.pack(side="left")

        # Rychlý skok na stroj (vpravo)
        srch_wrap = tk.Frame(top, bg="#e9f2f7")
        srch_wrap.pack(side="right", padx=(10, 10))
        tk.Label(srch_wrap, text=T("Stroj:", "Ma:"),
                 bg="#e9f2f7").pack(side="left")
        self.quick_go_var = tk.StringVar()
        quick_ent = tk.Entry(
            srch_wrap, textvariable=self.quick_go_var, width=6)
        quick_ent.pack(side="left")

        def _jump():
            raw = (self.quick_go_var.get() or "").strip()
            if not raw:
                return
            cand = raw.lstrip("0") or raw
            if cand in self.stroje:
                self.otevri_detail(cand)
            else:
                messagebox.showinfo(
                    T("Info", "Info"), f"{T('Stroj', 'Maschine')} {raw} {T('nenalezen', 'nicht gefunden')}.", parent=self)
            self.quick_go_var.set("")
        quick_ent.bind("<Return>", lambda e: (_jump(), "break"))

        # === STATUSBAR (pod lištou, aby se text nezkracoval) =========================
        statusbar = tk.Frame(self, bg="#eef5fb")
        statusbar.pack(fill="x")
        tk.Label(statusbar, textvariable=self.status, bg="#eef5fb",
                 font=("Segoe UI", 11)).pack(side="left", padx=10, pady=4)

        # === WARTUNG – EXPORT PANEL ================================================
        wartung_bar = tk.Frame(self, bg="#e9f2f7")
        wartung_bar.pack(fill="x", padx=10, pady=(4, 2))

        tk.Label(wartung_bar, text=T("Wartung export:", "Wartung Export:"), bg="#e9f2f7").pack(
            side="left", padx=(0, 6)
        )

        self.wartung_mode = tk.StringVar(value=T("≤ 30 dní", "≤ 30 Tage"))
        ttk.Combobox(
            wartung_bar,
            textvariable=self.wartung_mode,
            values=[T("prošlé", "überfällig"), T("≤ 30 dní", "≤ 30 Tage"), T(
                "vše s Wartung", "Alle mit Wartung")],
            state="readonly",
            width=12,
        ).pack(side="left", padx=(0, 6), pady=2)

        tk.Button(
            wartung_bar,
            text=T("Export Wartung", "Wartung export"),
            command=self.export_wartung_csv,
        ).pack(side="left", padx=(0, 6), pady=2)

        # ===== LEGENDA + FILTR KATEGORIÍ =====
        legend = tk.Frame(self, bg="#e9f2f7")
        legend.pack(fill="x", padx=10)
        tk.Label(legend, text=T("🟥 Elektrická", "🟥 Elektrisch"), bg="#f8c2c2",
                 padx=6).pack(side="left", padx=(0, 6))
        tk.Label(legend, text=T("🟦 Mechanická", "🟦 Mechanisch"), bg="#c2d4f8",
                 padx=6).pack(side="left", padx=6)
        tk.Label(legend, text=T("🟨 Jiná", "🟨 Sonstige"), bg="#f8f4c2",
                 padx=6).pack(side="left", padx=6)
        tk.Label(legend, text=T("🟩 OK", "🟩 OK"), bg="#c7f1d0",
                 padx=6).pack(side="left", padx=6)

        filtr = tk.Frame(self, bg="#e9f2f7")
        filtr.pack(fill="x", padx=10, pady=(0, 4))
        tk.Label(filtr, text=T("Zobrazit pouze:", "Nur anzeigen:"), bg="#e9f2f7",
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        self.filtr_kat = tk.StringVar(value="vse")

        def _set_f(cat):
            self.filtr_kat.set(cat)
            self.nakresli_mrizku()
        tk.Button(filtr, text=T("🟩 Vše", "🟩 Alle"),
                  command=lambda: _set_f("vse")).pack(side="left", padx=4)
        tk.Button(filtr, text=T("🟥 Elektrické", "🟥 Elektrisch"),
                  command=lambda: _set_f("elektricka")).pack(side="left", padx=4)
        tk.Button(filtr, text=T("🟦 Mechanické", "🟦 Mechanisch"),
                  command=lambda: _set_f("mechanicka")).pack(side="left", padx=4)
        tk.Button(filtr, text=T("🟨 Jiné", "🟨 Sonstige"),
                  command=lambda: _set_f("jina")).pack(side="left", padx=4)

        # skrolovatelná mřížka dlaždic
        wrap = tk.Frame(self, bg="#e9f2f7")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        self.canvas = tk.Canvas(wrap, bg="#e9f2f7", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(
            wrap, orient="vertical", command=self.canvas.yview)
        self.grid_frame = tk.Frame(self.canvas, bg="#e9f2f7")

        # A) Scrollregion podle obsahu (debounced)
        self.grid_frame.bind("<Configure>", self._on_grid_frame_configure)

        # B) Uložit ID vnořeného okna a držet stejnou šířku jako canvas
        self.win_id = self.canvas.create_window(
            (0, 0), window=self.grid_frame, anchor="nw")
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.configure(
            yscrollcommand=self.scrollbar.set, yscrollincrement=20)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel_linux)
        self.canvas.bind("<Button-5>", self._on_mousewheel_linux)

        # vykreslit
        self.nakresli_mrizku()
        self.bind("<Configure>", self.on_resize)

        # klávesy (numpad, čísla, Enter, N = nový stroj)
        self.bind("<Key>", self.on_key)
        self.bind_all("<Control-n>", lambda e: (
            self.last_selected and self.nova_porucha(self, self.stroje[self.last_selected])))
        self.bind_all("<Control-e>", lambda e: (
            self.last_selected and self.editovat_stroj_gui(self, self.last_selected)))
        # „H“ – historie (chráněno při psaní v Entry)
        self.bind_all("<Key-h>", lambda e: (self._focus_in_text_input() or (
            self.last_selected and self.historie_alarmu_gui(self, self.last_selected))))

        # při změně řazení překreslit
        self.sort_mode.trace_add("write", lambda *_: self.nakresli_mrizku())

        self.filter_only_problem = tk.BooleanVar(value=False)
        tk.Checkbutton(top, text=T("Defekte Ma", "Defekte Ma"), variable=self.filter_only_problem,
                       bg="#e9f2f7", command=self.nakresli_mrizku).pack(side="left", padx=(10, 0))

        tk.Button(top, text=f"📊 {T('Statistiky', 'Statistik')}",
                   command=self.statistiky_gui).pack(side="left", padx=(10, 0))

    # Pomoc: je fokus v textovém vstupu?

    def _focus_in_text_input(self):
        w = self.focus_get()
        return isinstance(w, (tk.Entry, tk.Text, tk.Spinbox))

    def _on_grid_frame_configure(self, event=None):
        if self._scrollregion_after_id is None:
            self._scrollregion_after_id = self.after_idle(self._flush_scrollregion)

    def _flush_scrollregion(self):
        self._scrollregion_after_id = None
        bbox = self.canvas.bbox("all")
        if bbox:
            self.canvas.configure(scrollregion=bbox)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.win_id, width=event.width)
        self._on_grid_frame_configure()

    def _on_mousewheel(self, event):
        self._wheel_accum += int(getattr(event, "delta", 0))
        if self._wheel_after_id is None:
            self._wheel_after_id = self.after(12, self._flush_mousewheel)
        return "break"

    def _on_mousewheel_linux(self, event):
        num = getattr(event, "num", None)
        if num == 4:
            self.canvas.yview_scroll(-3, "units")
        elif num == 5:
            self.canvas.yview_scroll(3, "units")
        return "break"

    def _flush_mousewheel(self):
        self._wheel_after_id = None
        total = self._wheel_remainder + (self._wheel_accum / 120.0)
        self._wheel_accum = 0
        if total >= 0:
            units = int(total // 1)
        else:
            units = -int((-total) // 1)
        self._wheel_remainder = total - units
        if units != 0:
            units = max(-12, min(12, units))
            self.canvas.yview_scroll(-units, "units")

    def _real_canvas_width(self):
        self.update_idletasks()
        w = self.canvas.winfo_width()
        if w < 200:
            w = self.winfo_width() - 48
        if w < 200:
            w = 1200
        return w

    def spocti_otevrene(self):
        counts = {}
        for p in self.poruchy:
            if p.get("stav") == "otevrena":
                c = p.get("cislo", "").strip()
                if c:
                    counts[c] = counts.get(c, 0) + 1
        return counts

    def _parse_dt(self, s):
        try:
            return datetime.strptime(s, "%Y-%m-%d %H:%M")
        except Exception:
            return None

    def spocti_poruchy(self, days=None):
        """Vrátí dict cislo -> počet poruch (všechny/posledních X dní)."""
        now = datetime.now()
        out = {}
        for p in self.poruchy:
            c = (p.get("cislo") or "").strip()
            if not c:
                continue
            if days:
                dt = self._parse_dt(p.get("cas", ""))
                if not dt or (now - dt).days > days:
                    continue
            out[c] = out.get(c, 0) + 1
        return out

    def statistiky_gui(self):
        # přepočítat z čerstvých dat
        self.poruchy = nacti_poruchy()
        open_counts = self.spocti_otevrene()
        cnt_30 = self.spocti_poruchy(days=30)
        cnt_all = self.spocti_poruchy(days=None)

        win = tk.Toplevel(self)
        win.title(T("Statistiky – TOP stroje", "Statistik – TOP Maschinen"))
        win.geometry("800x520")

        cols = ("cislo", "vyrobce", "typ", "otevrene", "za_30d", "celkem")

        HEAD = {
            "cislo":    T("Číslo", "Nr."),
            "vyrobce":  T("Výrobce", "Hersteller"),
            "typ":      T("Typ", "Typ"),
            "otevrene": T("Otevřené", "Offen"),
            "za_30d":   T("Za_30d", "30T"),
            "celkem":   T("Celkem", "Gesamt"),
        }

        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c, w in zip(cols, (70, 160, 260, 90, 90, 90)):
            tree.heading(c, text=HEAD.get(c, c))
            tree.column(c, width=w, stretch=True)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        # připrav data
        rows = []
        for c_str, stroj in self.stroje.items():
            rows.append((
                int(c_str),
                stroj.get("vyrobce", ""),
                stroj.get("typ", ""),
                open_counts.get(c_str, 0),
                cnt_30.get(c_str, 0),
                cnt_all.get(c_str, 0),
            ))
        # default: seřadit podle otevřených desc, pak 30d, pak celkem
        rows.sort(key=lambda r: (-r[3], -r[4], -r[5], r[0]))

        for r in rows:
            tree.insert("", "end", values=r)

        # Export CSV
        def export_csv():
            fname = filedialog.asksaveasfilename(parent=win, defaultextension=".csv",
                                                 initialfile="statistiky_stroje.csv",
                                                 filetypes=[("CSV", "*.csv")])
            if not fname:
                return
            with open(fname, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(cols)
                for iid in tree.get_children():
                    w.writerow(tree.item(iid)["values"])
            messagebox.showinfo(T("Export", "Export"), T(
                f"Uloženo: {fname}", f"Gespeichert: {fname}"), parent=win)

        tk.Button(win, text=T("Export CSV", "CSV export"),
                  command=export_csv).pack(pady=(0, 10))

    def _get_columns(self, sloupcu):
        if sloupcu is not None:
            return sloupcu
        try:
            w = self._real_canvas_width()
        except Exception:
            w = 1200
        tile_w = 220
        return max(3, min(12, max(1, w // tile_w)))

    def _get_visible_machine_numbers(self, counts_open, counts_30d, counts_all):
        # seznam čísel strojů (jen číslice)
        cisla = []
        for k in self.stroje.keys():
            ks = str(k).strip()
            if ks.isdigit():
                cisla.append(int(ks))

        # filtr: jen problémové
        if getattr(self, "filter_only_problem", tk.BooleanVar(value=False)).get():
            cisla = [c for c in cisla if counts_open.get(str(c), 0) > 0]

        # filtr: podle kategorie (poslední otevřená)
        if hasattr(self, "filtr_kat") and self.filtr_kat.get() != "vse":
            wanted = self.filtr_kat.get()

            def ok_cat(c):
                key = str(c)
                otevrene = [p for p in self.poruchy if p.get(
                    "cislo") == key and p.get("stav") == "otevrena"]
                if not otevrene:
                    return False

                def _key(p):
                    try:
                        return datetime.strptime(p.get("cas", ""), "%Y-%m-%d %H:%M")
                    except Exception:
                        return datetime.min

                posledni = sorted(otevrene, key=_key)[-1]
                return normalize_kategorie(posledni.get("kategorie", "")) == wanted

            cisla = [c for c in cisla if ok_cat(c)]


        # řazení
        # řazení (odděleno do helperu)
        cisla = self._apply_sort(cisla, counts_open, counts_30d, counts_all)
        return cisla

    def _normalize_tile_choice(self, choice: str) -> str:
        c = (choice or "cislo_only").strip().lower()
        if c == "typ":
            c = "seriove"
        if c in ("cislo", "nummer", "nummer ma", "nummer_ma", "nummerma", "nummer-ma"):
            c = "cislo_only"
        return c

    def _build_tile_subtitle(self, stroj: dict, choice: str) -> str:
        if choice == "cislo_only":
            return ""

        field = choice  # vyrobce | rok | spm | seriove
        val = (stroj.get(field) or "").strip()

        if (not val) and field == "vyrobce":
            val = (stroj.get("typ") or "").strip()

        if field == "rok":
            m = re.findall(r"\d{4}", val)
            val = m[0] if m else val

        max_len = 16
        if not val:
            return ""
        return f"\n{val[:max_len] + '…' if len(val) > max_len else val}"

    def _apply_wartung_border(self, widget, stroj: dict):
        wartung_dni = days_to_next_wartung(stroj)
        if wartung_dni is None:
            widget.config(highlightthickness=0)
            return

        if wartung_dni <= 0:
            widget.config(highlightbackground="#d40000",
                          highlightthickness=3)  # prošlá
        elif wartung_dni <= 30:
            widget.config(highlightbackground="#ffd000",
                          highlightthickness=2)  # blíží se
        else:
            widget.config(highlightthickness=0)

    def _build_tooltip(self, cislo: int, stroj: dict, open_count: int) -> str:
        dny = days_to_next_wartung(stroj)
        if dny is None:
            wart = ""
        elif dny <= 0:
            wart = f"\n{T('Wartung', 'Wartung')}: ❗ {T('PROŠLÁ', 'ÜBERFÄLLIG')}"
        else:
            wart = f"\n{T('Wartung', 'Wartung')}: {T('za', 'in')} {dny} {T('dní', 'Tagen')}"

        return (
            f"{T('Stroj', 'Maschine')}: {cislo}\n"
            f"{T('Výrobce', 'Hersteller')}: {stroj.get('vyrobce', '')}\n"
            f"{T('Typ', 'Typ')}: {stroj.get('typ', '')}\n"
            f"{T('Rok', 'Jahr')}: {stroj.get('rok', '')}\n"
            f"{T('SPM', 'SPM')}: {stroj.get('spm', '')}\n"
            f"{T('S/N', 'S/N')}: {stroj.get('seriove', '')}\n"
            f"{T('Stav', 'Status')}: {stav_ui(stroj.get('stav', ''))}\n"
            f"{T('Otevřené poruchy', 'Offene Störungen')}: {open_count}"
            f"{wart}"
        )

    def _last_open_dt(self, cislo: str):
        """Datetime poslední otevřené poruchy pro stroj (min pokud nic)."""
        last_dt = datetime.min
        for p in self.poruchy:
            if str(p.get("cislo")) != str(cislo):
                continue
            if p.get("stav") != "otevrena":
                continue
            try:
                dt = datetime.strptime(p.get("cas", ""), "%Y-%m-%d %H:%M")
            except Exception:
                dt = datetime.min
            if dt > last_dt:
                last_dt = dt
        return last_dt

    def _apply_sort(self, cisla, counts_open, counts_30d, counts_all):
        """Vrátí seřazený list čísel strojů podle self.sort_mode."""
        mode = (self.sort_mode.get() if hasattr(
            self, "sort_mode") else "cislo") or "cislo"

        if mode == "cislo":
            return sorted(cisla)

        if mode == "otevrene_desc":
            return sorted(cisla, key=lambda c: (-counts_open.get(str(c), 0), c))

        if mode == "poruchy_30d":
            return sorted(cisla, key=lambda c: (-counts_30d.get(str(c), 0), c))

        if mode == "poruchy_all":
            return sorted(cisla, key=lambda c: (-counts_all.get(str(c), 0), c))

        if mode == "last_open":
            # nejnovější otevřená porucha nahoře; stroje bez otevřené poruchy spadnou dolů
            return sorted(cisla, key=lambda c: (self._last_open_dt(str(c)), c), reverse=True)

        # fallback
        return sorted(cisla)

    def nakresli_mrizku(self, sloupcu=None):
        """Vykreslí mřížku dlaždic se stroji."""
        sloupcu = self._get_columns(sloupcu)
        self._last_draw_columns = sloupcu

        # vyčistit plochu
        for wdg in self.grid_frame.winfo_children():
            wdg.destroy()

        # data pro barvy/řazení
        counts_open = self.spocti_otevrene()
        counts_30d = self.spocti_poruchy(days=30)
        counts_all = self.spocti_poruchy(days=None)

        cisla = self._get_visible_machine_numbers(
            counts_open, counts_30d, counts_all)

        r, c = 0, 0
        for cislo in cisla:
            key = str(cislo)
            s = self.stroje.get(key, {})
            open_count = counts_open.get(key, 0)

            choice = self._normalize_tile_choice(
                self.tile_field.get() if hasattr(self, "tile_field") else "cislo_only")
            subtitle = self._build_tile_subtitle(s, choice)
            cnt_line = f"\n({open_count})" if open_count > 0 else ""

            tile = tk.Label(
                self.grid_frame,
                text=f"{cislo:02d}{subtitle}{cnt_line}",
                bd=1, relief="solid", width=12, height=4,
                font=("Segoe UI", 20, "bold"), fg="#0b1b2b",
                bg=barva_dlazdice(s.get("stav", "bezi"),
                                  open_count, key, self.poruchy),
            )

            self._apply_wartung_border(tile, s)
            tile.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")

            create_tooltip(tile, self._build_tooltip(cislo, s, open_count))

            tile.bind("<Button-1>", lambda e, num=key: (self._select(num)))
            tile.bind("<Double-1>", lambda e,
                      num=key: (self._select(num), self.otevri_detail(num)))
            tile.bind("<Button-3>", lambda e, num=key: self._tile_menu(num, e))

            c += 1
            if c >= sloupcu:
                c = 0
                r += 1

        # „+ Přidat stroj“
        add = tk.Label(
            self.grid_frame,
            bd=1, relief="solid", width=14, height=4,
            font=("Segoe UI", 14, "bold"),
            text=T("+ Přidat stroj", "+ Maschine hinzuf."),
            bg="#e0e0e0",
        )
        add.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
        add.bind("<Button-1>", lambda e: self.pridat_stroj_gui())
        add.bind("<Double-1>", lambda e: self.pridat_stroj_gui())

        self._on_grid_frame_configure()

        for i in range(sloupcu):
            self.grid_frame.grid_columnconfigure(i, weight=1)

    def _select(self, num: str):
        self.last_selected = num
        self.status.set(
            f"{T('Vybráno', 'Ausgewählt')}: {num} — {T('dvojklik pro akce', 'Doppelklick für Aktionen')}")

    def on_resize(self, event):
        if event.widget is self:
            next_cols = self._get_columns(None)
            if next_cols == self._last_draw_columns:
                return
            self._pending_columns = next_cols
            if self._resize_after_id is not None:
                self.after_cancel(self._resize_after_id)
            self._resize_after_id = self.after(80, self._flush_resize_redraw)

    def _flush_resize_redraw(self):
        self._resize_after_id = None
        cols = self._pending_columns
        self._pending_columns = None
        self.nakresli_mrizku(cols)

    # klávesy (globální zadávání čísla stroje, Enter = otevřít detail)
    def on_key(self, e: tk.Event):
        if self._focus_in_text_input():  # nespouštět při psaní do Entry/Text
            return
        ks = e.keysym
        if ks.startswith("KP_") and ks[3:].isdigit():
            ch = ks[3:]
            self.vstup_str += ch
            self.status.set(
                f"{T('Zadáno', 'Eingegeben')}: {self.vstup_str}  (Enter={T('potvrdit', 'bestätigen')}, Backspace={T('smazat', 'löschen')}, N={T('přidat stroj', 'Maschine hinzufügen')})")
            return
        if ks.isdigit():
            self.vstup_str += ks
            self.status.set(
                f"{T('Zadáno', 'Eingegeben')}: {self.vstup_str}  (Enter={T('potvrdit', 'bestätigen')}, Backspace={T('smazat', 'löschen')}, N={T('přidat stroj', 'Maschine hinzufügen')})")
            return
        if ks in ("BackSpace",):
            self.vstup_str = self.vstup_str[:-1]
            self.status.set(f"{T('Zadáno', 'Eingegeben')}: {self.vstup_str}")
            return
        if ks == "Return":
            cand = self.vstup_str.lstrip("0") or self.vstup_str
            if cand in self.stroje:
                self.otevri_detail(cand)
            elif self.vstup_str in self.stroje:
                self.otevri_detail(self.vstup_str)
            else:
                messagebox.showinfo(
                    T("Info", "Info"), f"{T('Stroj', 'Maschine')} {self.vstup_str or '…'} {T('nenalezen', 'nicht gefunden')}.", parent=self)
            self.vstup_str = ""
            self.status.set(
                T("Zadej číslo a Enter (nebo dvojklik). N = přidat stroj.",
                  "Nummer eingeben und Enter (oder Doppelklick). N = Maschine hinzufügen."))
            return
        if ks == "Escape":
            self.vstup_str = ""
            self.status.set(
                T("Zadej číslo a Enter (nebo dvojklik). N = přidat stroj.",
                  "Nummer eingeben und Enter (oder Doppelklick). N = Maschine hinzufügen."))
            return
        if ks.lower() == "n":
            self.pridat_stroj_gui()

    # Kontextové menu na dlaždici
    def _tile_menu(self, cislo: str, event=None):
        m = tk.Menu(self, tearoff=False)
        m.add_command(
            label=f"{T('Otevřít detail', 'Detail öffnen')} {cislo}", command=lambda: self.otevri_detail(cislo))
        m.add_command(label=T("Složka souborů…", "Dateiordner…"),
                      command=lambda: otevrit_slozku(slozka_stroje(cislo)))
        m.add_separator()
        stroj = self.stroje.get(cislo, {
                                "cislo": cislo, "typ": "", "vyrobce": "", "rok": "", "spm": "", "seriove": "", "stav": "bezi"})
        m.add_command(label=UIT("➕", T("Nová porucha", "Neue Störung")),
                      command=lambda: self.nova_porucha(self, stroj))
        m.add_command(label=UIT("✅", T("Uzavřít poruchu (podle alarmu)", "Störung schl. (nach Alarm)")),
                      command=lambda: self.uzavrit_poruchu_podle_alarmu(self, cislo))
        m.add_command(label=UIT("✏️", T("Editovat otevřenou poruchu", "Offene Störung bearbeiten")),
                      command=lambda: self.editovat_otevrenou_poruchu(self, cislo))
        m.add_separator()
        m.add_command(label=UIT("✏️", T("Editovat stroj…", "Maschine bearbeiten…")),
                      command=lambda: self.editovat_stroj_gui(self, cislo))
        m.add_command(label=UIT("🗑️", T("Smazat stroj…", "Maschine löschen…")),
                      command=lambda: self.smazat_stroj_gui(cislo))
        try:
            if event:
                m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()



class StrojeGrid(StrojeGrid):  # rozšíření
    def otevri_detail(self, cislo: str):
        stroj = self.stroje.get(cislo)
        if not stroj:
            messagebox.showwarning(
                T("Chyba", "Fehler"), f"{T('Stroj', 'Maschine')} {cislo} {T('nenalezen', 'nicht gefunden')}.", parent=self)
            return
        win = tk.Toplevel(self)
        win.title(
            f"{T('Stroj', 'Maschine')} {cislo} — {stroj.get('vyrobce', '')} {stroj.get('typ', '')}")
        win.geometry("820x600")
        win.after(0, lambda: self._detail_ui(win, cislo, stroj))

    def _detail_ui(self, win, cislo, stroj):
        win.lift()
        try:
            win.focus_force()
        except:
            pass

        header = tk.Frame(win, padx=10, pady=10)
        header.pack(fill="x")
        tk.Label(header, text=f"{T('Stroj', 'Maschine')} {cislo}", font=(
            "Segoe UI", 18, "bold")).pack(anchor="w")
        meta = []
        if stroj.get("vyrobce"):
            meta.append(stroj["vyrobce"])
        if stroj.get("typ"):
            meta.append(stroj["typ"])
        if stroj.get("rok"):
            meta.append(f"{T('rok', 'Jahr')} {stroj['rok']}")
        if stroj.get("spm"):
            meta.append(f"SPM {stroj['spm']}")
        if stroj.get("seriove"):
            meta.append(f"S/N {stroj['seriove']}")
        tk.Label(header, text=" · ".join(meta)).pack(anchor="w", pady=(2, 8))

        actions = tk.Frame(win, padx=10, pady=6)
        actions.pack(fill="x")

        tk.Button(
            actions,
            text=UIT("➕", T("Nová porucha", "Neue Störung")),
            command=lambda: self.nova_porucha(win, stroj),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        tk.Button(
            actions,
            text=UIT("✅", T("Uzavřít poruchu", "Störung schließen")),
            command=lambda: self.uzavrit_poruchu_podle_alarmu(win, cislo),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        tk.Button(
            actions,
            text=UIT("✓", T("Wartung dnes", "Wartung heute")),
            command=lambda: self.oznacit_wartung_dnes(win, cislo),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        tk.Button(
            actions,
            text=UIT("🧾", T("Export PDF", "PDF Export")),
            command=lambda: export_poruchy_pdf(win, cislo, self.stroje),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        tk.Button(
            actions,
            text=UIT("✏️", T("Editovat stroj", "Maschine bearbeiten")),
            command=lambda: self.editovat_stroj_gui(win, cislo),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        tk.Button(
            actions,
            text=UIT("📜", T("Historie", "Historie")),
            command=lambda: self.historie_alarmu_gui(win, cislo),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        tk.Button(
            actions,
            text=UIT("📂", T("Složka souborů…", "Dateiordner…")),
            command=lambda: otevrit_slozku(slozka_stroje(cislo)),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        srch = tk.Frame(win, padx=10, pady=6)
        srch.pack(fill="x")
        tk.Label(srch, text=T("Hledat řešení podle alarmu:",
                 "Lösung nach Alarm suchen:")).pack(anchor="w")
        alarm_var = tk.StringVar()
        ent = tk.Entry(srch, textvariable=alarm_var)
        ent.pack(side="left", fill="x", expand=True)
        ent.bind("<Return>", lambda e: self.hledat_reseni_gui(
            win, cislo, alarm_var.get()))
        tk.Button(srch, text=T("Hledat", "Suchen"), command=lambda: self.hledat_reseni_gui(
            win, cislo, alarm_var.get())).pack(side="left", padx=6)

        body = tk.Frame(win, padx=10, pady=10)
        body.pack(fill="both", expand=True)
        box = tk.Text(body, height=12, wrap="word")
        box.pack(fill="both", expand=True)
        otevrene = [p for p in self.poruchy if p.get(
            "cislo") == cislo and p.get("stav") == "otevrena"]
        if not otevrene:
            box.insert("1.0", T("Žádné otevřené poruchy.",
                       "Keine offenen Störungen."))
        else:
            lines = [f"[{p.get('cas')}] alarm {p.get('alarm')} | {kat_ui(p.get('kategorie', ''))} | {p.get('popis') or '-'}" for p in otevrene]

            box.insert("1.0", "\n".join(lines))
        box.configure(state="disabled")

    def hledat_reseni_gui(self, parent, cislo, alarm):
        alarm = (alarm or "").strip()
        if not alarm:
            return
        poruchy = nacti_poruchy()
        nalez = [p for p in poruchy if p["alarm"] == alarm and p.get("reseni")]
        text = [f"{T('Řešení pro alarm', 'Lösung für Alarm')} {alarm}:"]+[
            f"• {p['reseni']}" for p in nalez] if nalez else [T("Nic nenalezeno.", "Nichts gefunden.")]
        tip = self.sablony.get(alarm)
        if tip:
            text.append(f"{T('Šablona', 'Vorlage')}: {tip}")
        messagebox.showinfo(T("Výsledky", "Ergebnisse"),
                            "\n".join(text), parent=parent)

    def nova_porucha(self, parent, stroj):
        alarm = simpledialog.askstring(
            T("Nová porucha", "Neue Störung"), T("Alarm:", "Alarm:"), parent=parent)
        if not alarm:
            return

        kat = ask_kategorie_combobox(parent)   # ⬅️ nový dialog
        if kat is None:
            return

        popis = simpledialog.askstring(
            T("Nová porucha", "Neue Störung"), T("Popis:", "Beschreibung:"), parent=parent) or ""
        por = nacti_poruchy()
        pid = nove_id(por)
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        por.append({
            "id": pid,
            "cas": now,
            "cas_uzavreni": "",
            "cislo": stroj["cislo"],
            "typ": stroj.get("typ", ""),
            "alarm": alarm,
            "kategorie": kat,
            "popis": popis,
            "reseni": "",
            "stav": "otevrena"
        })
        uloz_poruchy(por)
        self.poruchy = por
        messagebox.showinfo(
            "OK", T("Porucha přidána.", "Störung hinzugefügt."), parent=parent)

    def uzavrit_otevrenou_poruchu(self, parent, cislo: str):
        por = nacti_poruchy()
        opened = [p for p in por if p.get("cislo") == str(cislo) and p.get("stav") =="otevrena"]

        if not opened:
            messagebox.showinfo(T("Uzavření", "Schließen"), T(
                "Žádná otevřená porucha u tohoto stroje.", "Keine offene Störung an dieser Maschine."), parent=parent)
            return

        # 🔽 NOVĚ: výběr konkrétní poruchy přes Combobox
        target = vyber_otevrenou_poruchu_combo(parent, opened)
        if target is None:
            return

        # zadání řešení
        reseni = simpledialog.askstring(T("Uzavřít", "Schließen"), f"{T('Řešení poruchy', 'Lösung Störung')} ({target.get('alarm', '')}):", parent=parent)
        if not reseni:
            return

        # zápis řešení + uzavření
        from datetime import datetime
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for p in por:
            if p.get("id") == target.get("id"):
                p["reseni"] = reseni
                p["stav"] = "uzavrena"
                p["cas_uzavreni"] = now
                break

        uloz_poruchy(por)
        self.poruchy = por
        self.nakresli_mrizku()
        messagebox.showinfo("OK", T("Porucha uzavřena.",
                            "Störung geschlossen."), parent=parent)

    def uzavrit_poruchu_podle_alarmu(self, parent, cislo: str):
        """Uzavření vybrané otevřené poruchy daného stroje (výběr přes Combobox)."""
        from datetime import datetime

        cislo = str(cislo)
        por = nacti_poruchy()

        opened = [
            p for p in por
            if p.get("cislo") == cislo and p.get("stav") == "otevrena"
        ]

        if not opened:
            messagebox.showinfo(
                T("Uzavření", "Schließen"),
                f"{T('Stroj', 'Maschine')} {cislo} {T('nemá žádné otevřené poruchy', 'hat keine offenen Störungen')}.",
                parent=parent
            )
            return

        # 1) vyber konkrétní poruchu (Combobox)
        target = vyber_otevrenou_poruchu_combo(parent, opened)
        if target is None:
            return  # uživatel kliknul Zrušit

        # 2) řešení poruchy
        reseni = simpledialog.askstring(
            T("Uzavřít", "Schließen"),
            f"{T('Řešení poruchy', 'Lösung Störung')} ({target.get('alarm', '')}):",
            parent=parent
        )
        if reseni is None or not reseni.strip():
            return

        # 3) operátor
        op = simpledialog.askstring(
            T("Operátor", "Operator"),
            T("Kdo poruchu uzavřel?", "Wer hat Störung geschlossen?"),
            initialvalue=self.operator,
            parent=parent,
        )
        if op is None or not op.strip():
            return
        self.operator = op.strip()

        now = datetime.now().strftime("%Y-%m-%d %H:%M")

        # 4) uložit změny zpět do seznamu poruch
        for p in por:
            if p.get("id") == target.get("id"):
                p["reseni"] = reseni
                p["stav"] = "uzavrena"
                p["cas_uzavreni"] = now
                p["operator_uzavrel"] = self.operator
                break

        uloz_poruchy(por)
        self.poruchy = por
        self.nakresli_mrizku()
        messagebox.showinfo("OK", T("Porucha uzavřena.",
                            "Störung geschlossen."), parent=parent)

        # pokud na stroji nezůstala žádná otevřená porucha → přepni stav na „bezi“
        open_left = any(p.get("cislo") == cislo and p.get(
            "stav") == "otevrena" for p in por)
        if not open_left and self.stroje.get(cislo, {}).get("stav") != "bezi":
            self.stroje[cislo]["stav"] = "bezi"
            uloz_stroje(self.stroje)

        # === Editovat stroj (všechny nové kolonky) ===

    def editovat_stroj_gui(self, parent, cislo):
        stroj = self.stroje.get(cislo)
        if not stroj:
            messagebox.showwarning(
                T("Editovat stroj", "Maschine bearbeiten"), f"{T('Stroj', 'Maschine')} {cislo} {T('nenalezen', 'nicht gefunden')}.", parent=parent)
            return

        vyrobce = simpledialog.askstring(
            T("Editovat stroj", "Maschine bearbeiten"), T("Výrobce:", "Hersteller:"), initialvalue=stroj.get("vyrobce", ""), parent=parent)
        if vyrobce is None:
            return
        typ = simpledialog.askstring(
            T("Editovat stroj", "Maschine bearbeiten"), T("Typ stroje:", "Maschinentyp:"), initialvalue=stroj.get("typ", ""), parent=parent)
        if typ is None:
            return

        rok = simpledialog.askstring(
            T("Editovat stroj", "Maschine bearbeiten"), T("Rok výroby (YYYY):", "Baujahr (YYYY):"), initialvalue=stroj.get("rok", ""), parent=parent)
        if rok is None:
            return
        rok = (rok or "").strip()
        m = re.findall(r"\d{4}", rok)
        rok = m[0] if m else ""

        spm = simpledialog.askstring(
            T("Editovat stroj", "Maschine bearbeiten"), T("SPM:", "SPM:"), initialvalue=stroj.get("spm", ""), parent=parent)
        if spm is None:
            return
        spm = (spm or "").strip()

        seriove = simpledialog.askstring(
            T("Editovat stroj", "Maschine bearbeiten"), T("Sériové číslo:", "Seriennummer:"), initialvalue=stroj.get("seriove", ""), parent=parent)
        if seriove is None:
            return
        seriove = (seriove or "").strip()

        # --- Stav: výběr z menu (interně ukládáme bezi/porucha) ---
        stav_key = normalize_stav(stroj.get("stav", "bezi"))
        stav_var = tk.StringVar(value=STAV_UI.get(stav_key, STAV_UI["bezi"]))

        dlg = tk.Toplevel(parent)
        dlg.title(T("Editovat stroj", "Maschine bearbeiten"))
        dlg.transient(parent)
        dlg.grab_set()

        frm = tk.Frame(dlg, padx=12, pady=10)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text=T("Stav", "Status") + \
                 ":").grid(row=0, column=0, sticky="w")

        stav_menu = tk.OptionMenu(frm, stav_var, "")
        m = stav_menu["menu"]
        m.delete(0, "end")

        for key, label in STAV_UI.items():
            m.add_command(
                label=label,
                command=lambda l=label: stav_var.set(l)
            )

        stav_menu.grid(row=0, column=1, sticky="we", padx=(8, 0))
        frm.grid_columnconfigure(1, weight=1)

        btns = tk.Frame(frm, pady=10)
        btns.grid(row=1, column=0, columnspan=2, sticky="e")

        result = {"ok": False}

        def _ok():
            result["ok"] = True
            dlg.destroy()

        def _cancel():
            dlg.destroy()

        tk.Button(btns, text=T("OK", "OK"), width=10,
                  command=_ok).pack(side="left", padx=(0, 6))
        tk.Button(btns, text=T("Zrušit", "Abbrechen"),
                  width=10, command=_cancel).pack(side="left")

        dlg.bind("<Return>", lambda e: _ok())
        dlg.bind("<Escape>", lambda e: _cancel())
               # --- vycentrovat nad parent (jinak to WM někdy hodí do rohu) ---
        dlg.update_idletasks()
        center_over(dlg, parent)

        dlg.wait_window()
        if not result["ok"]:
            return

        stav_ui = stav_var.get()
        stav = STAV_UI_REV.get(stav_ui, "bezi")  # => "bezi" / "porucha"

        stroj.update({
            "vyrobce": vyrobce or "",
            "typ": typ or "",
            "rok": rok,
            "spm": spm,
            "seriove": seriove,
            "stav": stav
        })
        uloz_stroje(self.stroje)
        self.nakresli_mrizku()
        messagebox.showinfo(
            T("Uloženo", "Gespeichert"), f"{T('Stroj', 'Maschine')} {cislo} {T('upraven', 'bearbeitet')}.", parent=parent)

    def oznacit_wartung_dnes(self, parent, cislo):
        """Označí Wartung na daném stroji jako provedenou dnes."""
        today = date.today().strftime("%Y-%m-%d")
        cislo = str(cislo)

        # načteme aktuální stroje z CSV (dict)
        stroje_dict = nacti_stroje()
        stroj = stroje_dict.get(cislo)

        if not stroj:
            messagebox.showerror(
                T("Wartung", "Wartung"),
                f"{T('Stroj', 'Maschine')} {cislo} {T('nebyl nalezen', 'nicht gefunden')}.",
                parent=parent,
            )
            return

        stroj["wartung_last"] = today
        if not stroj.get("wartung_interval"):
            stroj["wartung_interval"] = "180"

        # uložíme zpět
        uloz_stroje(stroje_dict)

        # aktualizujeme self.stroje a překreslíme mřížku
        self.stroje = stroje_dict
        self.nakresli_mrizku()

        messagebox.showinfo(
            T("Wartung", "Wartung"),
            f"{T('Na stroji', 'An Maschine')} {cislo} {T('byla zaznamenána Wartung', 'wurde Wartung erfasst')} ({today}).",
            parent=parent,
        )

    def historie_alarmu_gui(self, parent, cislo):
        por = [p for p in nacti_poruchy() if p["cislo"] == cislo]
        if not por:
            messagebox.showinfo(T("Historie", "Historie"), T(
                "Žádné záznamy.", "Keine Einträge."), parent=parent)
            return
        lines = [
            f"{p['cas']} | {p['alarm']} | {porucha_stav_ui(p.get('stav', ''))} | {p.get('reseni','')}"
            for p in por
        ]

        messagebox.showinfo(T("Historie", "Historie"),
                            "\n".join(lines), parent=parent)

    def backup_zip(self):
        em.backup_zip(self)

    def restore_zip(self):
        if not em.restore_zip(self):
            return
        self.stroje = nacti_stroje()
        self.poruchy = nacti_poruchy()
        self.sablony = nacti_sablony()
        self.nakresli_mrizku()

    # === Globální hledání poruch + export CSV ===
    def global_search_gui(self):
        win = tk.Toplevel(self)
        win.title(T("Hledání poruch", "Störungssuche"))
        win.geometry("950x600")

        # --- filtry ---
        frm = tk.Frame(win, padx=10, pady=6)
        frm.pack(fill="x")

        tk.Label(frm, text=T("Stroj:", "Ma:")).grid(
            row=0, column=0, sticky="w")
        var_cislo = tk.StringVar()
        tk.Entry(frm, textvariable=var_cislo, width=8).grid(
            row=0, column=1, padx=(4, 12)
        )

        tk.Label(frm, text=T("Alarm:", "Alarm:")).grid(
            row=0, column=2, sticky="w")
        var_alarm = tk.StringVar()
        tk.Entry(frm, textvariable=var_alarm, width=12).grid(
            row=0, column=3, padx=(4, 12)
        )

        tk.Label(frm, text=T("Stav:", "Status:")).grid(
            row=0, column=4, sticky="w")
        var_stav = tk.StringVar(value=T("vše", "alle"))
        ttk.Combobox(
            frm,
            textvariable=var_stav,
            width=12,
            values=[T("vše", "alle"), T("otevrena", "offen"),
                      T("uzavrena", "geschlossen")],
        ).grid(row=0, column=5, padx=(4, 12))

        tk.Label(frm, text=T("Text (popis/řešení):", "Text (Beschr./Lösung):")).grid(
            row=1, column=0, sticky="w"
        )
        var_text = tk.StringVar()
        tk.Entry(frm, textvariable=var_text, width=40).grid(
            row=1,
            column=1,
            columnspan=5,
            sticky="we",
            padx=(4, 12),
            pady=(4, 6),
        )

        # --- výsledky ---
        cols = ("id", "cas", "cislo", "typ", "alarm",
                "kategorie", "stav", "reseni")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        HEADERS = {
            "id":        T("ID", "ID"),
            "cas":       T("Čas", "Zeit"),
            "cislo":     T("Číslo", "Nr."),
            "typ":       T("Typ", "Typ"),
            "alarm":     T("Alarm", "Alarm"),
            "kategorie": T("Kategorie", "Kategorie"),
            "stav":      T("Stav", "Status"),
            "reseni":    T("Řešení", "Lösung"),
        }

        for c, w in zip(cols, (60, 120, 60, 160, 80, 120, 110, 260)):
            tree.heading(c, text=HEADERS.get(c, c.upper()))
            tree.column(c, width=w, stretch=True)

        tree.pack(fill="both", expand=True, padx=10, pady=6)

        def filtruj():
            # vždy načti aktuální data
            self.poruchy = nacti_poruchy()
            tree.delete(*tree.get_children())

            f_cislo = var_cislo.get().strip()
            f_alarm = var_alarm.get().strip()
            f_stav = var_stav.get().strip().lower()
            f_stav_ui = var_stav.get().strip().lower()

            # UI -> interní klíč
            if f_stav_ui in (T("vše", "alle").lower(), "vse", "alle"):
                f_stav = "vse"
            elif f_stav_ui in (T("otevrena", "offen").lower(), "otevrena", "offen"):
                f_stav = "otevrena"
            elif f_stav_ui in (T("uzavrena", "geschlossen").lower(), "uzavrena", "geschlossen"):
                f_stav = "uzavrena"
            else:
                f_stav = "vse"

            f_text = var_text.get().strip().lower()

            data = self.poruchy
            if f_cislo:
                data = [p for p in data if p.get("cislo") == f_cislo]
            if f_alarm:
                data = [p for p in data if p.get("alarm") == f_alarm]
            if f_stav in ("otevrena", "uzavrena"):
                data = [p for p in data if p.get("stav") == f_stav]

            if f_text:
                data = [
                    p
                    for p in data
                    if f_text
                    in (p.get("popis", "") + " " + p.get("reseni", "")).lower()
                ]

            # seřadit dle času
            def _key(p):
                try:
                    return datetime.strptime(p.get("cas", ""), "%Y-%m-%d %H:%M")
                except Exception:
                    return datetime.min

            data.sort(key=_key, reverse=True)

            # uložíme seznam právě zobrazených poruch pro detail na dvojklik
            win.zobrazene_poruchy = list(data)

            for p in data:
                tree.insert(
                    "",
                    "end",
                    values=(
                        p.get("id", ""),
                        p.get("cas", ""),
                        p.get("cislo", ""),
                        p.get("typ", ""),
                        p.get("alarm", ""),
                        kat_ui(p.get("kategorie", "")),
                        porucha_stav_ui(p.get("stav", "")),
                        (p.get("reseni", "") or ""),
                    ),
                )

        def export_csv():
            # vezmi to, co je zobrazené v tree
            rows = []
            for iid in tree.get_children():
                rows.append(tree.item(iid)["values"])
            if not rows:
                messagebox.showinfo(T("Export", "Export"), T(
                    "Není co exportovat.", "Nichts zu exportieren."), parent=win)
                return
            fname = filedialog.asksaveasfilename(
                parent=win,
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv")],
                initialfile="poruchy_filtr.csv",
            )
            if not fname:
                return
            with open(fname, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([c for c in cols])
                w.writerows(rows)
            messagebox.showinfo(
                T("Export", "Export"), f"{T('Uloženo', 'Gespeichert')}: {fname}", parent=win)

        # --- detail poruchy na dvojklik ---
        def ukaz_detail_poruchy(event):
            item_id = tree.focus()
            if not item_id:
                return

            values = tree.item(item_id, "values")
            if not values:
                return

            pid = str(values[0])  # první sloupec = ID
            por_list = getattr(win, "zobrazene_poruchy", [])
            por = next((p for p in por_list if str(
                p.get("id", "")) == pid), None)
            if not por:
                return

            txt = []
            txt.append(f"ID: {por.get('id', '')}")
            txt.append(
                f"{T('Stroj', 'Maschine')}: {por.get('cislo', '')}  {T('Typ', 'Typ')}: {por.get('typ', '')}")
            txt.append("")
            txt.append(f"{T('Čas', 'Zeit')}: {por.get('cas', '')}  {T('Stav', 'Status')}: {porucha_stav_ui(por.get('stav',''))}")
            txt.append(
                f"{T('Alarm', 'Alarm')}: {por.get('alarm', '')}  {T('Kategorie', 'Kategorie')}: {kat_ui(por.get('kategorie', ''))}")
            txt.append(T("Popis:", "Beschreibung:"))
            txt.append(por.get("popis", "") or "-")
            txt.append("")
            txt.append(T("Řešení:", "Lösung:"))
            txt.append(por.get("reseni", "") or "-")
            txt.append("")
            txt.append(
                f"{T('Operátor uzavřel', 'Operator geschlossen')}: {por.get('operator_uzavrel', '') or '-'}"
            )

            detail = tk.Toplevel(win)
            detail.title(f"{T('Detail poruchy', 'Störungsdetail')} {pid}")
            msg = tk.Message(detail, text="\n".join(txt), width=800)
            msg.pack(padx=10, pady=10)

        tree.bind("<Double-1>", ukaz_detail_poruchy)

        # --- tlačítka ---
        btns = tk.Frame(win, padx=10, pady=6)
        btns.pack(fill="x")
        tk.Button(btns, text=T("Hledat", "Suchen"),
                  command=filtruj).pack(side="left")
        tk.Button(btns, text=T("Export CSV", "CSV export"), command=export_csv).pack(
            side="left", padx=8
        )

        filtruj()  # první naplnění

    def export_wartung_csv(self):
        em.export_wartung_csv(self)

    def prepnout_stav_toolbar(self):
        """Tlačítko z horní lišty – přepnutí stavu stroje běží/porucha"""
        cislo = simpledialog.askstring(
            T("Přepnout stav", "Status wechseln"), T("Číslo stroje:", "Maschinen-Nr:"), parent=self)
        if not cislo:
            return
        try:
            key = str(int(str(cislo).strip()))
        except Exception:
            messagebox.showwarning(
                T("Chyba", "Fehler"), T("Neplatné číslo stroje.", "Ungültige Maschinen-Nr."), parent=self)
            return

        stroj = self.stroje.get(key)
        if not stroj:
            messagebox.showwarning(
                T("Chyba", "Fehler"), f"{T('Stroj', 'Maschine')} {key} {T('nenalezen', 'nicht gefunden')}.", parent=self)
            return

        # Přepnutí stavu: běží ↔ porucha
        cur = normalize_stav(stroj.get("stav", "bezi"))
        stroj["stav"] = "porucha" if cur == "bezi" else "bezi"

        uloz_stroje(self.stroje)
        self.nakresli_mrizku()
        self.status.set(f"{T('Stroj', 'Maschine')} {key}: {T('stav', 'Status')} → {stav_ui(stroj.get('stav', ''))}"
                        )

    def graf_top_stroje(self):
        try:
            import matplotlib
            matplotlib.use("TkAgg")
            import matplotlib.pyplot as plt
        except Exception:
            messagebox.showerror(
                T("Graf", "Graph"),
                T("Chybí knihovna matplotlib.\nNainstaluj: py -m pip install matplotlib",
                  "Matplotlib-Bibliothek fehlt.\nInstallieren: py -m pip install matplotlib"))
            return

        # data za 30 dní (fallback na vše)
        from collections import defaultdict, Counter
        self.poruchy = nacti_poruchy()
        now = datetime.now()

        def in_30(p):
            try:
                dt = datetime.strptime(p.get("cas", ""), "%Y-%m-%d %H:%M")
                return (now - dt).days <= 30
            except Exception:
                return False

        data_30 = [p for p in self.poruchy if in_30(p)]
        dataset = data_30 if data_30 else self.poruchy

        period_dates = []
        for p in dataset:
            try:
                period_dates.append(datetime.strptime(
                    p.get("cas", ""), "%Y-%m-%d %H:%M"))
            except Exception:
                pass

        by_machine_cat = defaultdict(lambda: Counter())
        for p in dataset:
            by_machine_cat[str(p.get("cislo"))][normalize_kategorie(
                p.get("kategorie"))] += 1

        if not by_machine_cat:
            messagebox.showinfo(T("Graf", "Graph"), T(
                "Zatím žádné poruchy k zobrazení.", "Noch keine Störungen zum Anzeigen."))
            return

        top = sorted(by_machine_cat.items(),
                     key=lambda kv: sum(kv[1].values()),
                     reverse=True)[:10]

        labels = [f"{cislo} {self.stroje.get(cislo, {}).get('vyrobce', '')}".strip()
                  for cislo, _ in top]

        E = [cnts.get("elektricka", 0) for _, cnts in top]
        M = [cnts.get("mechanicka", 0) for _, cnts in top]
        J = [cnts.get("jina", 0) for _, cnts in top]

        # --- NOVĚ: použijeme fig/ax a nastavíme prostor nahoře ---
        import numpy as np
        x = np.arange(len(labels))

        fig, ax = plt.subplots(figsize=(8, 3))

        p1 = ax.bar(x, E, label=T("Elektrická", "Elektrisch"),
                    color=COLORS["elektricka"])
        p2 = ax.bar(x, M, bottom=E, label=T(
            "Mechanická", "Mechanisch"), color=COLORS["mechanicka"])
        bottom_J = [e + m for e, m in zip(E, M)]
        p3 = ax.bar(x, J, bottom=bottom_J, label=T(
            "Jiná", "Sonstige"), color=COLORS["jina"])

        tot = [e + m + j for e, m, j in zip(E, M, J)]

        # maximální hodnota pro nastavení rozsahu osy
        max_tot = max(tot) if tot else 0
        if max_tot <= 0:
            max_tot = 1

        # trochu místa nad sloupci (30 % navíc)
        ax.set_ylim(0, max_tot * 1.3)

        # čísla přišpendlíme těsně nad sloupce
        for xi, y in zip(x, tot):
            ax.text(xi, y + max_tot * 0.03, str(y),
                    ha="center", va="bottom", fontsize=9)

        if period_dates:
            date_from = min(period_dates)
            date_to = max(period_dates)
            days_total = (date_to.date() - date_from.date()).days + 1
            period_label = (
                f"{date_from.strftime('%d.%m.%Y')} - "
                f"{date_to.strftime('%d.%m.%Y')}, {days_total} "
                f"{T('dní', 'Tage')}"
            )
        else:
            period_label = T("bez data", "ohne Datum")

        ax.legend()
        ax.set_title(
            f"{T('TOP problémové stroje', 'TOP Problem-Maschinen')} ({period_label})",
            pad=15
        )
        ax.set_xlabel(T("Stroj", "Maschine"))
        ax.set_ylabel(T("Počet poruch", "Anzahl Störungen"))
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=25, ha="right")

        fig.tight_layout()

        # nabídnout uložení PNG
        if messagebox.askyesno(T("Graf", "Graph"), T("Chceš uložit graf jako PNG?", "Graf als PNG speichern?")):
            from tkinter import filedialog
            fname = filedialog.asksaveasfilename(
                parent=self,
                defaultextension=".png",
                filetypes=[("PNG", "*.png")],
                initialfile="top_stroje.png",
            )
            if fname:
                fig.savefig(fname, dpi=150)

        plt.show()

    # === Smazání stroje (+ volitelné smazání poruch) ===
    def hromadne_uzavrit(self, parent):
        ids = bulk_uzavrit_dialog(parent, self.poruchy)
        if ids is None:
            return
        if not ids:
            messagebox.showinfo(T("Hromadné uzavření", "Sammelschließung"),
                                T("Nic nebylo vybráno.", "Nichts ausgewählt."), parent=parent)
            return
        n = 0
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for p in self.poruchy:
            if p.get("id") in ids and p.get("stav") == "otevrena":
                p["stav"] = "uzavrena"
                p["cas_uzavreni"] = now
                n += 1
        uloz_poruchy(self.poruchy)
        self.nakresli_mrizku()
        messagebox.showinfo(
            T("Hotovo", "Fertig"), f"{T('Uzavřeno', 'Geschlossen')} {n} {T('záznamů', 'Einträge')}.", parent=parent)

    def smazat_stroj_gui(self, cislo: str):
        if cislo not in self.stroje:
            messagebox.showwarning(T("Smazat stroj", "Maschine löschen"),
                                   f"{T('Stroj', 'Maschine')} {cislo} {T('nenalezen', 'nicht gefunden')}.", parent=self)
            return
        if messagebox.askyesno(T("Smazat stroj", "Maschine löschen"), f"{T('Opravdu smazat stroj', 'Maschine wirklich löschen')} {cislo}?", parent=self):
            del_por = messagebox.askyesno(T("Poruchy", "Störungen"), T(
                "Smazat i všechny poruchy tohoto stroje?", "Auch alle Störungen dieser Maschine löschen?"), parent=self)
            self.stroje.pop(cislo, None)
            uloz_stroje(self.stroje)
            if del_por:
                por = nacti_poruchy()
                por = [p for p in por if p.get("cislo") != cislo]
                uloz_poruchy(por)
                self.poruchy = por
            self.nakresli_mrizku()
            messagebox.showinfo(T("Hotovo", "Fertig"), T(
                "Stroj byl smazán.", "Maschine wurde gelöscht."), parent=self)

    def editovat_otevrenou_poruchu(self, parent, cislo: str):
        por = nacti_poruchy()
        opened = [p for p in por if p.get(
            "cislo") == cislo and p.get("stav") == "otevrena"]
        if not opened:
            messagebox.showinfo(T("Editace", "Bearbeiten"), T(
                "Žádná otevřená porucha u tohoto stroje.", "Keine offene Störung an dieser Maschine."), parent=parent)
            return

        # Když je více otevřených, vyber ID
        target = opened[0]
        if len(opened) > 1:
            ids = ", ".join(p.get("id", "?") for p in opened)
            sel = simpledialog.askstring(T("Výběr poruchy", "Störung auswählen"),
                                         f"{T('Otevřené ID', 'Offene ID')}: {ids}\n{T('Zadej ID k editaci', 'ID zur Bearbeitung eingeben')}:", parent=parent)
            if not sel:
                return
            found = [p for p in opened if p.get("id") == sel]
            if not found:
                messagebox.showwarning(
                    T("Editace", "Bearbeiten"), f"ID {sel} {T('nenalezeno', 'nicht gefunden')}.", parent=parent)
                return
            target = found[0]

        # Editace polí
        new_alarm = simpledialog.askstring(T("Editace", "Bearbeiten"), T(
            "Alarm:", "Alarm:"), initialvalue=target.get("alarm", ""), parent=parent)
        if new_alarm is None:
            return
        new_kat = ask_kategorie_combobox(parent)
        if new_kat is None:
            return
        new_pop = simpledialog.askstring(T("Editace", "Bearbeiten"), T(
            "Popis:", "Beschreibung:"), initialvalue=target.get("popis", ""), parent=parent)
        if new_pop is None:
            return

        # Uložit změny
        for p in por:
            if p.get("id") == target.get("id"):
                p["alarm"] = new_alarm
                p["kategorie"] = new_kat
                p["popis"] = new_pop
                break
        uloz_poruchy(por)
        self.poruchy = por
        messagebox.showinfo(T("Uloženo", "Gespeichert"),
                            f"{T('Porucha', 'Störung')} ID {target.get('id')} {T('upravena', 'bearbeitet')}.", parent=parent)

    def pridat_stroj_gui(self):
        """Dialog pro přidání nového stroje přes dlaždici + nebo klávesu N."""
        navrh = next_free_machine_number(self.stroje)
        cislo = simpledialog.askstring(T("Přidat stroj", "Maschine hinzufügen"),
                                       f"{T('Číslo stroje', 'Maschinen-Nr')} (Enter = {navrh}):",
            parent=self) or navrh

        cislo = str(cislo).strip().lstrip("0") or "0"
        if not cislo.isdigit():
            messagebox.showwarning(T("Přidat stroj", "Maschine hinzufügen"), T(
                "Číslo musí být celé číslo.", "Nummer muss ganzzahlig sein."), parent=self)
            return
        if cislo in self.stroje:
            messagebox.showwarning(T("Přidat stroj", "Maschine hinzufügen"),
                                   f"{T('Stroj', 'Maschine')} {cislo} {T('už existuje', 'existiert bereits')}.", parent=self)
            return

        vyrobce = simpledialog.askstring(T("Přidat stroj", "Maschine hinzufügen"), T(
            "Výrobce:", "Hersteller:"), parent=self)
        if vyrobce is None:
            return
        vyrobce = (vyrobce or "").strip()

        typ = simpledialog.askstring(T("Přidat stroj", "Maschine hinzufügen"), T(
            "Typ stroje:", "Maschinentyp:"), parent=self)
        if typ is None:
            return
        typ = (typ or "").strip()

        rok = simpledialog.askstring(T("Přidat stroj", "Maschine hinzufügen"), T(
            "Rok výroby (YYYY):", "Baujahr (YYYY):"), parent=self)
        if rok is None:
            return
        rok = (rok or "").strip()
        m = re.findall(r"\d{4}", rok)
        rok = m[0] if m else ""

        spm = simpledialog.askstring(
            T("Přidat stroj", "Maschine hinzufügen"), T("SPM:", "SPM:"), parent=self)
        if spm is None:
            return
        spm = (spm or "").strip()

        seriove = simpledialog.askstring(T("Přidat stroj", "Maschine hinzufügen"), T(
            "Sériové číslo:", "Seriennummer:"), parent=self)
        if seriove is None:
            return
        seriove = (seriove or "").strip()

        stav = simpledialog.askstring(T("Přidat stroj", "Maschine hinzufügen"), T(
            "Stav (b/p/běží/porucha):", "Status (l/s/läuft/Störung):"), parent=self)
        if stav is None:
            return
        stav = normalize_stav(stav)

        self.stroje[cislo] = {
            "cislo": cislo,
            "vyrobce": vyrobce,
            "typ": typ,
            "rok": rok,
            "spm": spm,
            "seriove": seriove,
            "stav": stav
        }
        uloz_stroje(self.stroje)
        self.nakresli_mrizku()
        messagebox.showinfo(T("Přidat stroj", "Maschine hinzufügen"),
                            f"{T('Stroj', 'Maschine')} {cislo} {T('uložen', 'gespeichert')}.", parent=self)

    # === Editace libovolné otevřené poruchy u stroje ===
    def _vyber_otevrenou_poruchu(self, parent, cislo: str):
        """Vrátí dict vybrané otevřené poruchy nebo None."""
        por = nacti_poruchy()
        otevrene = [p for p in por if p.get(
            "cislo") == cislo and p.get("stav") == "otevrena"]
        if not otevrene:
            messagebox.showinfo(T("Uzavřít poruchu", "Störung schließen"),
                                T("Žádné otevřené poruchy.", "Keine offenen Störungen."), parent=parent)
            return None

        # Když je jen jedna, rovnou ji vrátíme
        if len(otevrene) == 1:
            return otevrene[0]

        # Výběrové okno s filtrem Alarm
        win = tk.Toplevel(parent)
        win.title(
            f"{T('Vyber otevřenou poruchu', 'Offene Störung wählen')} – {T('stroj', 'Maschine')} {cislo}")
        win.geometry("720x420")
        win.transient(parent)
        win.grab_set()

        top = tk.Frame(win, padx=10, pady=6)
        top.pack(fill="x")
        tk.Label(top, text=T("Filtr alarmu:", "Alarm-Filter:")).pack(side="left")
        var_alarm = tk.StringVar()
        ent_alarm = tk.Entry(top, textvariable=var_alarm, width=12)
        ent_alarm.pack(side="left", padx=6)

        cols = ("id", "cas", "alarm", "kategorie", "popis")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=12)
        widths = (60, 120, 90, 100, 340)
        for c, w in zip(cols, widths):
            tree.heading(c, text=c.upper())
            tree.column(c, width=w, stretch=True)
        tree.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        def napln():
            tree.delete(*tree.get_children())
            f = var_alarm.get().strip()
            data = otevrene if not f else [
                p for p in otevrene if p.get("alarm", "").startswith(f)]
            # seřadit dle času

            def _key(p):
                try:
                    return datetime.strptime(p.get("cas", ""), "%Y-%m-%d %H:%M")
                except:
                    return datetime.min
            data.sort(key=_key, reverse=True)
            for p in data:
                tree.insert("", "end", values=(p.get("id", ""), p.get("cas", ""),
                                               p.get("alarm", ""), p.get(
                                                   "kategorie", ""),
                                               p.get("popis", "") or ""))

        def enter_filter(e): napln()
        ent_alarm.bind("<Return>", enter_filter)
        napln()

        vybrana = {"p": None}

        def potvrdit():
            sel = tree.focus()
            if not sel:
                messagebox.showwarning(
                    T("Výběr", "Auswahl"), T("Nejprve vyber řádek.", "Zuerst Zeile auswählen."), parent=win)
                return
            vid = tree.item(sel)["values"][0]  # ID
            found = [p for p in otevrene if p.get("id") == str(vid)]
            vybrana["p"] = found[0] if found else None
            win.destroy()

        btns = tk.Frame(win, padx=10, pady=6)
        btns.pack(fill="x")
        tk.Button(btns, text=T("Vybrat", "Auswählen"),
                  command=potvrdit).pack(side="left")
        tk.Button(btns, text=T("Zrušit", "Abbrechen"), command=win.destroy).pack(
            side="left", padx=8)
        win.wait_window()
        return vybrana["p"]

    def prepnout_stav_toolbar(self):
        # vezmi poslední vybraný stroj nebo si číslo vyžádej
        cislo = self.last_selected
        if not cislo:
            cislo = simpledialog.askstring(
                T("Přepnout stav", "Status wechseln"), T("Zadej číslo stroje:", "Maschinen-Nr eingeben:"), parent=self)
        if not cislo:
            return
        cislo = str(cislo).strip().lstrip("0") or "0"
        if cislo not in self.stroje:
            messagebox.showinfo(
                T("Info", "Info"), f"{T('Stroj', 'Maschine')} {cislo} {T('nenalezen', 'nicht gefunden')}.", parent=self)
            return

        s = self.stroje[cislo]
        cur = normalize_stav(s.get("stav", "bezi"))
        s["stav"] = "porucha" if cur == "bezi" else "bezi"
        uloz_stroje(self.stroje)
        self.nakresli_mrizku()


def vyber_otevrenou_poruchu_combo(parent, opened: list):
    """
    Vrátí vybraný záznam poruchy (dict) z listu 'opened' pomocí Comboboxu.
    Pokud uživatel zruší, vrací None.
    """
    if not opened:
        return None

    import tkinter as tk
    from tkinter import ttk

    top = tk.Toplevel(parent)
    top.title(T("Vyber poruchu", "Störung wählen"))
    top.transient(parent)
    top.grab_set()
    top.resizable(False, False)

    # hezké zobrazení: ID | datum | kategorie | alarm
    items = []
    for p in opened:
        kat_key = normalize_kategorie(p.get("kategorie", "") or "")
        kat_txt = kat_ui(kat_key)
        items.append(f"{p.get('id', '?')} | {p.get('cas','?')} | {kat_txt} | {p.get('alarm','?')}")

    tk.Label(top, text=T("Vyber poruchu k uzavření:", "Störung zum Schließen wählen:"),
             anchor="w").pack(padx=12, pady=(12, 6), fill="x")
    var = tk.StringVar(value=items[0])
    cb = ttk.Combobox(top, textvariable=var, values=items,
                      state="readonly", width=70)
    cb.current(0)
    cb.pack(padx=12, pady=6, fill="x")

    chosen = {"index": 0}  # box pro výsledek

    def _ok():
        try:
            chosen["index"] = items.index(var.get())
        except ValueError:
            chosen["index"] = 0
        top.destroy()

    def _cancel():
        chosen["index"] = None
        top.destroy()

    btns = tk.Frame(top)
    btns.pack(padx=12, pady=(8, 12), fill="x")
    ttk.Button(btns, text="OK", width=12, command=_ok).pack(
        side="left", padx=(0, 6))
    ttk.Button(btns, text=T("Zrušit", "Abbrechen"),
               width=12, command=_cancel).pack(side="left")

    top.bind("<Return>", lambda e: _ok())
    top.bind("<Escape>", lambda e: _cancel())

    # umístění uprostřed rodiče
    parent.update_idletasks()
    x = parent.winfo_rootx() + parent.winfo_width()//2 - 250
    y = parent.winfo_rooty() + parent.winfo_height()//2 - 60
    top.geometry(f"+{x}+{y}")

    top.wait_window()
    if chosen["index"] is None:
        return None
    return opened[chosen["index"]]


# ===== Main =====
if __name__ == "__main__":
    app = StrojeGrid()
    app.mainloop()
