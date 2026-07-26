#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import shutil
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import csv
import re
from datetime import date, datetime, timedelta
from pathlib import Path
import data_manager as dm
import export_manager as em
import audit_manager as audit
import audit_ui
import operator_manager
import operator_ui
import validation_ui
from app_logging import configure_logging, install_tk_exception_handler

APP_NAME = "SGM Evidence"
APP_VERSION = "1.2.0"


def bundled_resource(*parts: str) -> Path:
    """Vrátí cestu k assetu ve zdrojové i PyInstaller verzi."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base.joinpath(*parts)


# --- Přepínač emoji v UI (0 = vypnout emoji, 1 = zapnout) ---
USE_EMOJI = os.environ.get("SGM_USE_EMOJI", "1") != "0"


def UI(emoji_text: str, plain_text: str) -> str:
    return emoji_text if USE_EMOJI else plain_text
# --------------------------------------------------------------
# ===== SORT MODE (interní klíče vs. UI text) =====


# --- Jazyk UI (CZ/DE) ---
# Interní klíče (cislo, stav, kategorie, ...) se NIKDY nepřekládají.
T = dm.T


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


def install_text_context_menu(root: tk.Misc):
    """Globální kontextové menu pro Entry/ttk.Entry/Text."""
    menu = tk.Menu(root, tearoff=False)

    def _is_editable(widget) -> bool:
        try:
            state = str(widget.cget("state")).lower()
        except Exception:
            return True
        return state not in ("disabled", "readonly")

    def _show(event):
        widget = event.widget
        if not isinstance(widget, (tk.Entry, tk.Text, ttk.Entry)):
            return

        try:
            widget.focus_force()
        except Exception:
            pass

        editable = _is_editable(widget)
        menu.delete(0, "end")
        menu.add_command(
            label=T("Vyjmout", "Ausschneiden"),
            command=lambda w=widget: w.event_generate("<<Cut>>"),
            state=("normal" if editable else "disabled"),
        )
        menu.add_command(
            label=T("Kopírovat", "Kopieren"),
            command=lambda w=widget: w.event_generate("<<Copy>>"),
        )
        menu.add_command(
            label=T("Vložit", "Einfügen"),
            command=lambda w=widget: w.event_generate("<<Paste>>"),
            state=("normal" if editable else "disabled"),
        )
        menu.add_separator()

        def _select_all(w=widget):
            if isinstance(w, tk.Text):
                w.tag_add("sel", "1.0", "end-1c")
                w.mark_set("insert", "1.0")
                w.see("insert")
            else:
                w.selection_range(0, "end")
                w.icursor("end")

        menu.add_command(label=T("Vybrat vše", "Alles markieren"), command=_select_all)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    for seq in ("<Button-3>", "<Button-2>", "<Control-Button-1>"):
        root.bind_class("Entry", seq, _show, add="+")
        root.bind_class("TEntry", seq, _show, add="+")
        root.bind_class("Text", seq, _show, add="+")


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

class StrojeGridBase(tk.Tk):
    def __init__(self):
        super().__init__()
        install_text_context_menu(self)
        self.title(
            T(
                f"{APP_NAME} {APP_VERSION} – přehled strojů",
                f"{APP_NAME} {APP_VERSION} – Maschinenübersicht",
            )
        )
        try:
            self.iconbitmap(
                default=str(bundled_resource("assets", "sgm_jokey.ico"))
            )
        except tk.TclError:
            # Na některých ne-Windows Tk sestaveních není formát ICO podporován.
            pass
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
        # Jméno operátora se při skutečném spuštění ještě potvrdí v dialogu.
        self.operator = operator_manager.initial_operator()
        self.poruchy = nacti_poruchy()    # list
        self.sablony = nacti_sablony()    # dict alarm -> reseni
        # --- STAV UI (musí být připraveno před tvorbou widgetů) ---
        self.status = tk.StringVar(value=T("Zadej číslo a Enter (nebo dvojklik). N = přidat stroj.",
                                           "Nummer eingeben und Enter (oder Doppelklick). N = Maschine hinzufügen."))
        self.operator_display = tk.StringVar()
        self._refresh_operator_display()
        self.filter_only_problem = tk.BooleanVar(
            value=False)  # jen stroje s otevřenou poruchou
        self.show_archived = tk.BooleanVar(value=False)
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
        tk.Button(top, text=UIT("✓", T("Kontrola dat", "Datenprüf.")),
                  command=self.kontrola_dat_gui).pack(side="left", padx=(6, 0), pady=6)
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

        tk.Label(rightgrp, text=T("Údaj na dlaždici:", "Daten Ma:"),
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
        tk.Button(
            statusbar,
            text=T("Historie změn", "Änderungsverlauf"),
            command=lambda: audit_ui.open_audit_history(self),
        ).pack(side="right", padx=10, pady=2)
        tk.Button(
            statusbar,
            textvariable=self.operator_display,
            command=self.vybrat_operatora_gui,
        ).pack(side="right", padx=(4, 0), pady=2)

        # === WARTUNG – EXPORT PANEL ================================================
        wartung_bar = tk.Frame(self, bg="#e9f2f7")
        wartung_bar.pack(fill="x", padx=10, pady=(4, 2))

        tk.Label(wartung_bar, text=T("Export údržby:", "Wartung Export:"), bg="#e9f2f7").pack(
            side="left", padx=(0, 6)
        )

        self.wartung_mode = tk.StringVar(value=T("≤ 30 dní", "≤ 30 Tage"))
        ttk.Combobox(
            wartung_bar,
            textvariable=self.wartung_mode,
            values=[T("prošlé", "überfällig"), T("≤ 30 dní", "≤ 30 Tage"), T(
                "vše s údržbou", "Alle mit Wartung")],
            state="readonly",
            width=12,
        ).pack(side="left", padx=(0, 6), pady=2)

        tk.Button(
            wartung_bar,
            text=T("Export údržby", "Wartung export"),
            command=self.export_wartung_csv,
        ).pack(side="left", padx=(0, 6), pady=2)

        # ===== LEGENDA + FILTR KATEGORIÍ =====
        filtr = tk.Frame(self, bg="#e9f2f7")
        filtr.pack(fill="x", padx=10, pady=(0, 4))

        self._filter_buttons = {}
        filter_defs = [
            ("vse", "Alle", "#d9dde3", "#c7ccd3"),
            ("elektricka", "Elektrisch", "#f8c2c2", "#efaaaa"),
            ("mechanicka", "Mechanisch", "#c2d4f8", "#a9c2f2"),
            ("jina", "Sonstige", "#f8f4c2", "#efe99b"),
            ("ok", "OK", "#c7f1d0", "#aee6bb"),
        ]

        def _refresh_filter_buttons():
            active = self.filtr_kat.get()
            for key, btn in self._filter_buttons.items():
                is_active = key == active
                btn.configure(
                    relief=("sunken" if is_active else "raised"),
                    bd=(3 if is_active else 1),
                    font=("Segoe UI", 9, "bold" if is_active else "normal"),
                    highlightthickness=(1 if is_active else 0),
                )

        def _set_f(cat):
            self.filtr_kat.set(cat)
            _refresh_filter_buttons()
            self.nakresli_mrizku()

        for key, de_label, bg, active_bg in filter_defs:
            cz_label = (
                "Alle" if key == "vse" else
                "Elektrisch" if key == "elektricka" else
                "Mechanisch" if key == "mechanicka" else
                "Sonstige" if key == "jina" else
                "OK"
            )
            btn = tk.Button(
                filtr,
                text=T(cz_label, de_label),
                bg=bg,
                activebackground=active_bg,
                padx=10,
                pady=2,
                command=lambda c=key: _set_f(c),
            )
            btn.pack(side="left", padx=(0, 6))
            self._filter_buttons[key] = btn

        _refresh_filter_buttons()
        tk.Checkbutton(
            filtr,
            text=T("Zobrazit archivované", "Archivierte anzeigen"),
            variable=self.show_archived,
            bg="#e9f2f7",
            command=self.nakresli_mrizku,
        ).pack(side="right", padx=(10, 0))

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
        tk.Checkbutton(top, text=T("Poruchové stroje", "Defekte Ma"), variable=self.filter_only_problem,
                       bg="#e9f2f7", command=self.nakresli_mrizku).pack(side="left", padx=(10, 0))

        tk.Button(top, text=f"📊 {T('Statistiky', 'Statistik')}",
                   command=self.statistiky_gui).pack(side="left", padx=(10, 0))

    # Pomoc: je fokus v textovém vstupu?

    def _focus_in_text_input(self):
        w = self.focus_get()
        return isinstance(w, (tk.Entry, tk.Text, tk.Spinbox))

    def _refresh_operator_display(self):
        if hasattr(self, "operator_display"):
            self.operator_display.set(
                f"{T('Operátor', 'Operator')}: {self.operator or '—'}"
            )

    def vybrat_operatora_gui(self, required: bool = False) -> bool:
        settings = operator_manager.load_settings()
        selected = operator_ui.choose_operator(
            self,
            self.operator,
            settings["operators"],
            required=required,
        )
        if selected is None:
            return False
        try:
            operator_manager.remember_operator(selected)
        except Exception as exc:
            messagebox.showwarning(
                T("Operátor", "Operator"),
                T(
                    f"Operátor byl nastaven pouze pro tuto relaci.\n"
                    f"Nastavení se nepodařilo uložit:\n{exc}",
                    f"Der Operator wurde nur für diese Sitzung gesetzt.\n"
                    f"Die Einstellung konnte nicht gespeichert werden:\n{exc}",
                ),
                parent=self,
            )
        self.operator = selected
        self._refresh_operator_display()
        self.status.set(
            T(
                f"Aktivní operátor: {selected}",
                f"Aktiver Operator: {selected}",
            )
        )
        return True

    def _audit_event(
        self,
        action: str,
        entity_type: str,
        *,
        entity_id: str = "",
        machine_number: str = "",
        details: dict | None = None,
    ):
        try:
            return audit.record_event(
                action,
                entity_type,
                entity_id=entity_id,
                machine_number=machine_number,
                operator=getattr(self, "operator", ""),
                details=details,
            )
        except Exception:
            audit.LOGGER.exception("Auditní událost se nepodařila uložit")
            if hasattr(self, "status"):
                self.status.set(
                    T(
                        "Změna byla provedena, ale auditní záznam se nepodařilo uložit.",
                        "Änderung durchgeführt, Auditprotokoll konnte aber nicht gespeichert werden.",
                    )
                )
            return None

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
        win.title(T("Statistiky strojů", "Maschinenstatistik"))
        win.geometry("900x560")

        cols = ("cislo", "vyrobce", "typ", "otevrene", "za_30d", "celkem")

        HEAD = {
            "cislo":    T("Číslo", "Nr."),
            "vyrobce":  T("Výrobce", "Hersteller"),
            "typ":      T("Typ", "Typ"),
            "otevrene": T("Otevřené", "Offen"),
            "za_30d":   T("Za_30d", "30T"),
            "celkem":   T("Celkem", "Gesamt"),
        }

        style_name = "TopMachines.Treeview"
        style = ttk.Style(win)
        style.configure(
            style_name,
            rowheight=24,
            font=("Segoe UI", 10),
            background="#ffffff",
            fieldbackground="#ffffff",
        )
        style.configure(
            f"{style_name}.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#efefef",
            relief="flat",
        )

        outer = tk.Frame(win, padx=12, pady=12)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer,
            text=T(
                "Nejčastěji poruchové stroje podle počtu záznamů.",
                "Top Maschinen nach Anzahl der Störungen.",
            ),
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        filter_row = tk.Frame(outer)
        filter_row.pack(fill="x", pady=(0, 10))
        tk.Label(filter_row, text=T("Zobrazení:", "Ansicht:")).pack(side="left")
        view_options = [
            T("Všechny stroje", "Alle Maschinen"),
            T("Top 10 podle celkem", "Top 10 nach Gesamt"),
            T("Top 10 podle otevřených", "Top 10 nach Offen"),
            T("Top 10 za posledních 30 dní", "Top 10 letzte 30 Tage"),
        ]
        view_var = tk.StringVar(value=view_options[0])
        view_combo = ttk.Combobox(
            filter_row,
            textvariable=view_var,
            values=view_options,
            state="readonly",
            width=28,
        )
        view_combo.pack(side="left", padx=(8, 0))

        table_frame = tk.Frame(outer)
        table_frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(table_frame, columns=cols, show="headings", style=style_name)
        col_cfg = {
            "cislo": (60, "center", False),
            "vyrobce": (150, "w", False),
            "typ": (320, "w", True),
            "otevrene": (80, "center", False),
            "za_30d": (80, "center", False),
            "celkem": (90, "center", False),
        }
        for c in cols:
            width, anchor, stretch = col_cfg[c]
            tree.heading(c, text=HEAD.get(c, c), anchor="center" if c != "typ" else "w")
            tree.column(c, width=width, minwidth=width, anchor=anchor, stretch=stretch)

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        tree.tag_configure("odd", background="#ffffff")
        tree.tag_configure("even", background="#f6f6f6")
        tree.tag_configure("open_alert", foreground="#b00020")
        tree.tag_configure("top1", foreground="#0f5c2e", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("top2", foreground="#2f6b3d")
        tree.tag_configure("top3", foreground="#4c7b56")

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
        rows.sort(key=lambda r: (-r[3], -r[4], -r[5], r[0]))

        def render_rows():
            selected_view = view_var.get()
            if selected_view == view_options[1]:
                visible_rows = sorted(rows, key=lambda r: (-r[5], -r[3], -r[4], r[0]))[:10]
            elif selected_view == view_options[2]:
                visible_rows = sorted(rows, key=lambda r: (-r[3], -r[5], -r[4], r[0]))[:10]
            elif selected_view == view_options[3]:
                visible_rows = sorted(rows, key=lambda r: (-r[4], -r[3], -r[5], r[0]))[:10]
            else:
                visible_rows = list(rows)

            for iid in tree.get_children():
                tree.delete(iid)

            top_values = sorted({r[5] for r in visible_rows}, reverse=True)[:3]
            top_rank = {val: idx + 1 for idx, val in enumerate(top_values)}

            for idx, r in enumerate(visible_rows):
                tags = ["even" if idx % 2 else "odd"]
                if r[3] > 0:
                    tags.append("open_alert")
                rank = top_rank.get(r[5])
                if rank == 1:
                    tags.append("top1")
                elif rank == 2:
                    tags.append("top2")
                elif rank == 3:
                    tags.append("top3")
                tree.insert("", "end", values=r, tags=tuple(tags))

        view_combo.bind("<<ComboboxSelected>>", lambda e: render_rows())
        render_rows()

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

        btns = tk.Frame(outer)
        btns.pack(fill="x", pady=(12, 0))
        tk.Button(btns, text=T("Export CSV", "CSV exportieren"),
                  command=export_csv).pack(side="right")

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
        include_archived = self.show_archived.get() if hasattr(
            self, "show_archived"
        ) else False
        for k, machine in self.stroje.items():
            ks = str(k).strip()
            if ks.isdigit() and (
                include_archived or not dm.is_archived_machine(machine)
            ):
                cisla.append(int(ks))

        # filtr: jen problémové
        problem_filter = getattr(self, "filter_only_problem", None)
        if problem_filter is not None and problem_filter.get():
            cisla = [c for c in cisla if counts_open.get(str(c), 0) > 0]

        # filtr: podle kategorie (poslední otevřená)
        if hasattr(self, "filtr_kat") and self.filtr_kat.get() != "vse":
            wanted = self.filtr_kat.get()

            def ok_cat(c):
                key = str(c)
                otevrene = [p for p in self.poruchy if p.get(
                    "cislo") == key and p.get("stav") == "otevrena"]
                if wanted == "ok":
                    return not otevrene
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
        if dm.is_archived_machine(stroj):
            widget.config(highlightthickness=0)
            return
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
            wart = f"\n{T('Údržba', 'Wartung')}: ❗ {T('PROŠLÁ', 'ÜBERFÄLLIG')}"
        else:
            wart = f"\n{T('Údržba', 'Wartung')}: {T('za', 'in')} {dny} {T('dní', 'Tagen')}"

        archived = (
            f"\n{T('Archivováno', 'Archiviert')}: {T('ano', 'ja')}"
            if dm.is_archived_machine(stroj)
            else ""
        )
        return (
            f"{T('Stroj', 'Maschine')}: {cislo}\n"
            f"{T('Výrobce', 'Hersteller')}: {stroj.get('vyrobce', '')}\n"
            f"{T('Typ', 'Typ')}: {stroj.get('typ', '')}\n"
            f"{T('Rok', 'Jahr')}: {stroj.get('rok', '')}\n"
            f"{T('SPM', 'SPM')}: {stroj.get('spm', '')}\n"
            f"{T('S/N', 'S/N')}: {stroj.get('seriove', '')}\n"
            f"{T('Stav', 'Status')}: {stav_ui(stroj.get('stav', ''))}\n"
            f"{T('Otevřené poruchy', 'Offene Störungen')}: {open_count}"
            f"{archived}"
            f"{wart}"
        )

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
            return sorted(cisla, key=lambda c: (last_open_dt(self.poruchy, str(c)), c), reverse=True)

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
            archived = dm.is_archived_machine(s)
            archive_line = f"\n{T('ARCHIV', 'ARCHIV')}" if archived else ""

            tile = tk.Label(
                self.grid_frame,
                text=f"{cislo:02d}{subtitle}{cnt_line}{archive_line}",
                bd=1, relief="solid", width=12, height=4,
                font=("Segoe UI", 20, "bold"), fg="#0b1b2b",
                bg=(
                    "#c9c9c9"
                    if archived
                    else barva_dlazdice(
                        s.get("stav", "bezi"), open_count, key, self.poruchy
                    )
                ),
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
        m.add_command(
            label=T("Historie změn…", "Änderungsverlauf…"),
            command=lambda: audit_ui.open_audit_history(self, cislo),
        )
        m.add_separator()
        stroj = self.stroje.get(cislo, {
                                "cislo": cislo, "typ": "", "vyrobce": "", "rok": "", "spm": "", "seriove": "", "stav": "bezi"})
        if dm.is_archived_machine(stroj):
            m.add_command(
                label=T("Obnovit z archivu", "Aus Archiv wiederherstellen"),
                command=lambda: self.obnovit_stroj_z_archivu(cislo),
            )
            try:
                if event:
                    m.tk_popup(event.x_root, event.y_root)
            finally:
                m.grab_release()
            return
        m.add_command(label=UIT("➕", T("Nová porucha", "Neue Störung")),
                      command=lambda: self.nova_porucha(self, stroj))
        m.add_command(label=UIT("✅", T("Uzavřít poruchu (podle alarmu)", "Störung schl. (nach Alarm)")),
                      command=lambda: self.uzavrit_poruchu_podle_alarmu(self, cislo))
        m.add_command(label=UIT("✏️", T("Editovat otevřenou poruchu", "Offene Störung bearbeiten")),
                      command=lambda: self.editovat_otevrenou_poruchu(self, cislo))
        m.add_command(label=T("Opravit uzavřenou poruchu…",
                              "Geschlossene Störung korrigieren…"),
                      command=lambda: self.korigovat_uzavrenou_poruchu(self, cislo))
        m.add_separator()
        m.add_command(label=UIT("✏️", T("Editovat stroj…", "Maschine bearbeiten…")),
                      command=lambda: self.editovat_stroj_gui(self, cislo))
        m.add_command(label=UIT("📦", T("Archivovat stroj…", "Maschine archivieren…")),
                      command=lambda: self.archivovat_stroj_gui(cislo))
        try:
            if event:
                m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()



class StrojeGrid(StrojeGridBase):
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
            text=UIT("✓", T("Údržba dnes", "Wartung heute")),
            command=lambda: self.oznacit_wartung_dnes(win, cislo),
        ).pack(side="left", expand=True, fill="x", padx=6, pady=6)

        tk.Button(
            actions,
            text=UIT("🧾", T("Export PDF", "PDF Export")),
            command=lambda: self.export_poruchy_pdf_s_filtrem(win, cislo),
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

    def export_poruchy_pdf_s_filtrem(self, parent, cislo):
        poruchy_stroje = [p for p in self.poruchy if str(p.get("cislo")) == str(cislo)]
        if not poruchy_stroje:
            messagebox.showinfo(
                T("Export PDF", "PDF-Export"),
                T(f"Stroj {cislo} nemá žádné poruchy k exportu.", f"Maschine {cislo} hat keine Störungen zum Export."),
                parent=parent,
            )
            return

        poruchy_stroje = sorted(
            poruchy_stroje,
            key=lambda p: (p.get("cas") or "", str(p.get("id") or "")),
            reverse=True,
        )

        def _zkrat(text, max_len=45):
            text = (text or "").strip()
            if len(text) <= max_len:
                return text or "-"
            return text[: max_len - 3].rstrip() + "..."

        def _text_do_sloupce(p):
            stav = (p.get("stav") or "").strip().lower()
            if stav == "otevrena":
                return p.get("popis", "") or ""
            return p.get("reseni", "") or ""

        dlg = tk.Toplevel(parent)
        dlg.title(T("Výběr poruch pro PDF export", "Störungen für den PDF-Export auswählen"))
        dlg.transient(parent)
        dlg.grab_set()
        dlg.resizable(True, True)
        dlg.geometry("980x420")

        frm = tk.Frame(dlg, padx=12, pady=12)
        frm.pack(fill="both", expand=True)

        tk.Label(
            frm,
            text=T("Vyber jednu nebo více konkrétních poruch pro export PDF.", "Eine oder mehrere konkrete Störungen für den PDF-Export auswählen."),
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        tk.Label(
            frm,
            text=T(
                "Řádky jsou seřazeny od nejnovějších. Označ celý záznam, ne jen alarm nebo řešení.",
                "Einträge sind absteigend sortiert. Bitte immer den kompletten Eintrag auswählen.",
            ),
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))

        filter_row = tk.Frame(frm)
        filter_row.grid(row=2, column=0, sticky="w", pady=(0, 8))
        tk.Label(filter_row, text=T("Stav:", "Status:")).pack(side="left")

        stav_options = [
            (T("Všechny", "Alle"), None),
            (T("Jen otevřené", "Nur offen"), "otevrena"),
            (T("Jen uzavřené", "Nur geschlossen"), "uzavrena"),
        ]
        stav_map = {label: value for label, value in stav_options}
        stav_var = tk.StringVar(value=stav_options[0][0])
        stav_combo = ttk.Combobox(
            filter_row,
            textvariable=stav_var,
            values=[label for label, _ in stav_options],
            state="readonly",
            width=18,
        )
        stav_combo.pack(side="left", padx=(6, 0))

        cols = ("cas", "alarm", "stav", "text")
        tree_frame = tk.Frame(frm)
        tree_frame.grid(row=3, column=0, sticky="nsew")

        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended", height=12)
        tree.heading("cas", text=T("Datum/čas", "Datum/Zeit"))
        tree.heading("alarm", text=T("Alarm", "Alarm"))
        tree.heading("stav", text=T("Stav", "Status"))
        tree.heading("text", text=T("Text", "Text"))
        tree.column("cas", width=160, anchor="w", stretch=False)
        tree.column("alarm", width=130, anchor="w", stretch=False)
        tree.column("stav", width=110, anchor="w", stretch=False)
        tree.column("text", width=520, anchor="w", stretch=True)

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(3, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        result = {"selected_ids": None}
        poruchy_by_id = {
            str(p.get("id", "")).strip(): p
            for p in poruchy_stroje
            if str(p.get("id", "")).strip()
        }

        def zobraz_detail_poruchy(iid):
            por = poruchy_by_id.get(str(iid).strip())
            if not por:
                return

            detail = tk.Toplevel(dlg)
            detail.title(T("Detail poruchy", "Detail der Störung"))
            detail.transient(dlg)
            detail.resizable(False, False)

            win_w = 760
            win_h = 440
            detail.geometry(f"{win_w}x{win_h}")
            detail.update_idletasks()
            parent_x = dlg.winfo_rootx()
            parent_y = dlg.winfo_rooty()
            parent_w = dlg.winfo_width()
            parent_h = dlg.winfo_height()
            pos_x = parent_x + max((parent_w - win_w) // 2, 0)
            pos_y = parent_y + max((parent_h - win_h) // 2, 0)
            detail.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

            status_text = porucha_stav_ui(por.get("stav", "")) or "-"
            status_key = (por.get("stav") or "").strip().lower()
            status_color = "#b00020" if status_key == "otevrena" else "#1f7a1f" if status_key == "uzavrena" else "black"
            wrap = 560

            frm_detail = tk.Frame(detail, padx=14, pady=14)
            frm_detail.pack(fill="both", expand=True)
            frm_detail.grid_columnconfigure(1, weight=1)

            radky = [
                (T("Datum/čas", "Datum/Zeit"), por.get("cas", "") or "-", None),
                ("Alarm", por.get("alarm", "") or "-", None),
                (T("Stav", "Status"), status_text, status_color),
                (T("Popis", "Beschreibung"), por.get("popis", "") or "-", None),
                (T("Řešení", "Lösung"), por.get("reseni", "") or "-", None),
            ]

            for idx, (label, value, color) in enumerate(radky):
                tk.Label(
                    frm_detail,
                    text=f"{label}:",
                    font=("Segoe UI", 9, "bold"),
                    anchor="nw",
                ).grid(row=idx, column=0, sticky="nw", padx=(0, 12), pady=5)
                tk.Label(
                    frm_detail,
                    text=value,
                    justify="left",
                    anchor="w",
                    wraplength=wrap,
                    fg=color or "black",
                ).grid(row=idx, column=1, sticky="w", pady=5)

            detail_text = "\n".join([
                f"{T('Datum/čas', 'Datum/Zeit')}: {por.get('cas', '') or '-'}",
                f"Alarm: {por.get('alarm', '') or '-'}",
                f"{T('Stav', 'Status')}: {status_text}",
                f"{T('Popis', 'Beschreibung')}: {por.get('popis', '') or '-'}",
                f"{T('Řešení', 'Lösung')}: {por.get('reseni', '') or '-'}",
            ])

            def kopirovat():
                detail.clipboard_clear()
                detail.clipboard_append(detail_text)
                detail.update()
                messagebox.showinfo(
                    T("Informace", "Info"),
                    T(
                        "Zkopírováno do schránky.",
                        "In die Zwischenablage kopiert.",
                    ),
                    parent=detail,
                )

            btns = tk.Frame(frm_detail)
            btns.grid(row=len(radky), column=0, columnspan=2, sticky="e", pady=(16, 0))
            btn = tk.Button(
                btns, text=T("Zavřít", "Schließen"), command=detail.destroy
            )
            btn.pack(side="right")
            copy_btn = tk.Button(
                btns,
                text=T("Kopírovat do schránky", "In Zwischenablage kopieren"),
                command=kopirovat,
            )
            copy_btn.pack(side="right", padx=(0, 8))

            detail.bind("<Escape>", lambda e: detail.destroy())
            detail.bind("<Return>", lambda e: detail.destroy())
            detail.protocol("WM_DELETE_WINDOW", detail.destroy)
            detail.after(0, btn.focus_set)

        def naplnit_tree():
            selected_now = set(tree.selection())
            current_filter = stav_map.get(stav_var.get())
            for item in tree.get_children():
                tree.delete(item)

            for p in poruchy_stroje:
                iid = str(p.get("id", "")).strip()
                if not iid:
                    continue
                stav_key = (p.get("stav") or "").strip().lower()
                if current_filter and stav_key != current_filter:
                    continue
                tree.insert(
                    "",
                    "end",
                    iid=iid,
                    values=(
                        p.get("cas", "") or "",
                        p.get("alarm", "") or "",
                        porucha_stav_ui(p.get("stav", "")),
                        _zkrat(_text_do_sloupce(p)),
                    ),
                )
                if iid in selected_now:
                    tree.selection_add(iid)

        def potvrdit():
            selected = list(tree.selection())
            if not selected:
                messagebox.showwarning(
                    T("Export PDF", "PDF-Export"),
                    T("Nejdřív označ alespoň jeden záznam poruchy.", "Zuerst mindestens einen Störungseintrag auswählen."),
                    parent=dlg,
                )
                return
            result["selected_ids"] = selected
            dlg.destroy()

        def zrusit():
            dlg.destroy()

        def otevrit_detail_event(event):
            iid = tree.identify_row(event.y)
            if not iid:
                selection = tree.selection()
                iid = selection[0] if selection else ""
            if not iid:
                return
            tree.selection_add(iid)
            tree.focus(iid)
            zobraz_detail_poruchy(iid)

        stav_combo.bind("<<ComboboxSelected>>", lambda e: naplnit_tree())
        tree.bind("<Double-1>", otevrit_detail_event)
        naplnit_tree()

        btns = tk.Frame(frm)
        btns.grid(row=4, column=0, sticky="e", pady=(12, 0))
        tk.Button(btns, text=T("Storno", "Abbrechen"), command=zrusit).pack(side="right", padx=(6, 0))
        tk.Button(btns, text=T("Tisk PDF", "PDF exportieren"), command=potvrdit).pack(side="right")

        dlg.bind("<Return>", lambda e: potvrdit())
        dlg.bind("<Escape>", lambda e: zrusit())
        dlg.protocol("WM_DELETE_WINDOW", zrusit)
        dlg.after(0, tree.focus_set)
        parent.wait_window(dlg)

        if not result["selected_ids"]:
            return

        export_poruchy_pdf(
            parent,
            cislo,
            self.stroje,
            selected_ids=result["selected_ids"],
        )

    def nova_porucha(self, parent, stroj):
        if dm.is_archived_machine(stroj):
            messagebox.showwarning(
                T("Archivovaný stroj", "Archivierte Maschine"),
                T(
                    "K archivovanému stroji nelze přidat novou poruchu. "
                    "Nejprve jej obnovte z archivu.",
                    "Zu einer archivierten Maschine kann keine neue Störung "
                    "hinzugefügt werden. Stellen Sie sie zuerst wieder her.",
                ),
                parent=parent,
            )
            return
        alarm = simpledialog.askstring(
            T("Nová porucha", "Neue Störung"), T("Alarm:", "Alarm:"), parent=parent)
        if not alarm:
            return

        kat = ask_kategorie_combobox(parent)
        if kat is None:
            return

        def ask_beschreibung_dialog(parent_win):
            dlg = tk.Toplevel(parent_win)
            dlg.title(T("Nová porucha", "Neue Störung"))
            dlg.transient(parent_win)
            dlg.grab_set()
            dlg.resizable(True, True)

            win_w = 580
            win_h = 300
            dlg.geometry(f"{win_w}x{win_h}")
            dlg.update_idletasks()
            parent_x = parent_win.winfo_rootx()
            parent_y = parent_win.winfo_rooty()
            parent_w = parent_win.winfo_width()
            parent_h = parent_win.winfo_height()
            pos_x = parent_x + max((parent_w - win_w) // 2, 0)
            pos_y = parent_y + max((parent_h - win_h) // 2, 0)
            dlg.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

            frm = tk.Frame(dlg, padx=12, pady=12)
            frm.pack(fill="both", expand=True)
            frm.grid_columnconfigure(0, weight=1)
            frm.grid_rowconfigure(1, weight=1)

            tk.Label(frm, text=T("Popis:", "Beschreibung:")).grid(
                row=0, column=0, sticky="w", pady=(0, 6)
            )

            text_frame = tk.Frame(frm)
            text_frame.grid(row=1, column=0, sticky="nsew")
            text_frame.grid_columnconfigure(0, weight=1)
            text_frame.grid_rowconfigure(0, weight=1)

            txt = tk.Text(text_frame, height=7, wrap="word")
            scroll = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=scroll.set)
            txt.grid(row=0, column=0, sticky="nsew")
            scroll.grid(row=0, column=1, sticky="ns")

            result = {"value": None}

            def potvrdit():
                result["value"] = txt.get("1.0", "end-1c").strip()
                dlg.destroy()

            def zrusit():
                dlg.destroy()

            btns = tk.Frame(frm)
            btns.grid(row=2, column=0, sticky="e", pady=(12, 0))
            tk.Button(btns, text=T("Zrušit", "Abbrechen"), command=zrusit).pack(side="right")
            tk.Button(btns, text="OK", command=potvrdit, width=10).pack(side="right", padx=(0, 8))

            dlg.bind("<Escape>", lambda e: zrusit())
            dlg.bind("<Control-Return>", lambda e: potvrdit())
            dlg.protocol("WM_DELETE_WINDOW", zrusit)
            dlg.after(0, txt.focus_set)
            parent_win.wait_window(dlg)
            return result["value"]

        popis = ask_beschreibung_dialog(parent)
        if popis is None:
            return

        por = nacti_poruchy()
        pid = nove_id(por)
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        new_fault = {
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
        }
        por.append(new_fault)
        uloz_poruchy(por)
        self.poruchy = por
        self._audit_event(
            "fault_created",
            "fault",
            entity_id=pid,
            machine_number=stroj["cislo"],
            details={"fields": audit.changed_fields({}, new_fault)},
        )
        messagebox.showinfo(
            "OK", T("Porucha přidána.", "Störung hinzugefügt."), parent=parent)

    def _ask_multiline_text_dialog(self, parent_win, title: str, label: str, initial_text: str = "", width: int = 580, height: int = 300):
        dlg = tk.Toplevel(parent_win)
        dlg.title(title)
        dlg.transient(parent_win)
        dlg.grab_set()
        dlg.resizable(True, True)

        dlg.geometry(f"{width}x{height}")
        dlg.update_idletasks()
        parent_x = parent_win.winfo_rootx()
        parent_y = parent_win.winfo_rooty()
        parent_w = parent_win.winfo_width()
        parent_h = parent_win.winfo_height()
        pos_x = parent_x + max((parent_w - width) // 2, 0)
        pos_y = parent_y + max((parent_h - height) // 2, 0)
        dlg.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

        frm = tk.Frame(dlg, padx=12, pady=12)
        frm.pack(fill="both", expand=True)
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(1, weight=1)

        tk.Label(frm, text=label).grid(row=0, column=0, sticky="w", pady=(0, 6))

        text_frame = tk.Frame(frm)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)

        txt = tk.Text(text_frame, height=7, wrap="word")
        scroll = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        txt.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        if initial_text:
            txt.insert("1.0", initial_text)

        result = {"value": None}

        def potvrdit():
            result["value"] = txt.get("1.0", "end-1c").strip()
            dlg.destroy()

        def zrusit():
            dlg.destroy()

        btns = tk.Frame(frm)
        btns.grid(row=2, column=0, sticky="e", pady=(12, 0))
        tk.Button(btns, text=T("Zrušit", "Abbrechen"), command=zrusit).pack(side="right")
        tk.Button(btns, text="OK", command=potvrdit, width=10).pack(side="right", padx=(0, 8))

        dlg.bind("<Escape>", lambda e: zrusit())
        dlg.bind("<Control-Return>", lambda e: potvrdit())
        dlg.protocol("WM_DELETE_WINDOW", zrusit)
        dlg.after(0, txt.focus_set)
        parent_win.wait_window(dlg)
        return result["value"]

    def uzavrit_otevrenou_poruchu(self, parent, cislo: str):
        por = nacti_poruchy()
        opened = [p for p in por if p.get("cislo") == str(cislo) and p.get("stav") == "otevrena"]

        if not opened:
            messagebox.showinfo(T("Uzavření", "Schließen"), T(
                "Žádná otevřená porucha u tohoto stroje.", "Keine offene Störung an dieser Maschine."), parent=parent)
            return

        target = vyber_otevrenou_poruchu_combo(parent, opened)
        if target is None:
            return

        reseni = self._ask_multiline_text_dialog(
            parent,
            T("Uzavřít", "Schließen"),
            f"{T('Řešení poruchy', 'Lösung Störung')} ({target.get('alarm', '')}):",
        )
        if not reseni:
            return

        from datetime import datetime
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        before = dict(target)
        updated = target
        for p in por:
            if p.get("id") == target.get("id"):
                p["reseni"] = reseni
                p["stav"] = "uzavrena"
                p["cas_uzavreni"] = now
                updated = p
                break

        uloz_poruchy(por)
        self.poruchy = por
        self._audit_event(
            "fault_closed",
            "fault",
            entity_id=target.get("id", ""),
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, updated)},
        )
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

        target = vyber_otevrenou_poruchu_combo(parent, opened)
        if target is None:
            return

        reseni = self._ask_multiline_text_dialog(
            parent,
            T("Uzavřít", "Schließen"),
            f"{T('Řešení poruchy', 'Lösung Störung')} ({target.get('alarm', '')}):",
        )
        if reseni is None or not reseni.strip():
            return

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
        before = dict(target)
        updated = target
        for p in por:
            if p.get("id") == target.get("id"):
                p["reseni"] = reseni
                p["stav"] = "uzavrena"
                p["cas_uzavreni"] = now
                p["operator_uzavrel"] = self.operator
                updated = p
                break

        uloz_poruchy(por)
        self.poruchy = por
        self._audit_event(
            "fault_closed",
            "fault",
            entity_id=target.get("id", ""),
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, updated)},
        )
        self.nakresli_mrizku()
        messagebox.showinfo("OK", T("Porucha uzavřena.",
                            "Störung geschlossen."), parent=parent)

        open_left = any(p.get("cislo") == cislo and p.get(
            "stav") == "otevrena" for p in por)
        if not open_left and self.stroje.get(cislo, {}).get("stav") != "bezi":
            before_machine = dict(self.stroje[cislo])
            self.stroje[cislo]["stav"] = "bezi"
            uloz_stroje(self.stroje)
            self._audit_event(
                "machine_status_changed",
                "machine",
                entity_id=cislo,
                machine_number=cislo,
                details={
                    "fields": audit.changed_fields(
                        before_machine, self.stroje[cislo]
                    ),
                    "reason": "all_faults_closed",
                },
            )

    def editovat_stroj_gui(self, parent, cislo):
        stroj = self.stroje.get(cislo)
        if not stroj:
            messagebox.showwarning(
                T("Editovat stroj", "Maschine bearbeiten"), f"{T('Stroj', 'Maschine')} {cislo} {T('nenalezen', 'nicht gefunden')}.", parent=parent)
            return
        before = dict(stroj)

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
        self._audit_event(
            "machine_updated",
            "machine",
            entity_id=cislo,
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, stroj)},
        )
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
                T("Údržba", "Wartung"),
                f"{T('Stroj', 'Maschine')} {cislo} {T('nebyl nalezen', 'nicht gefunden')}.",
                parent=parent,
            )
            return

        if dm.is_archived_machine(stroj):
            messagebox.showwarning(
                T("Archivovaný stroj", "Archivierte Maschine"),
                T(
                    "U archivovaného stroje nelze zaznamenat údržbu.",
                    "Für eine archivierte Maschine kann keine Wartung erfasst werden.",
                ),
                parent=parent,
            )
            return
        before = dict(stroj)
        stroj["wartung_last"] = today
        if not stroj.get("wartung_interval"):
            stroj["wartung_interval"] = "180"

        # uložíme zpět
        uloz_stroje(stroje_dict)
        self._audit_event(
            "maintenance_recorded",
            "machine",
            entity_id=cislo,
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, stroj)},
        )

        # aktualizujeme self.stroje a překreslíme mřížku
        self.stroje = stroje_dict
        self.nakresli_mrizku()

        messagebox.showinfo(
            T("Údržba", "Wartung"),
            f"{T('Na stroji', 'An Maschine')} {cislo} {T('byla zaznamenána údržba', 'wurde Wartung erfasst')} ({today}).",
            parent=parent,
        )

    def historie_alarmu_gui(self, parent, cislo):
        por = [p for p in nacti_poruchy() if p["cislo"] == cislo]
        if not por:
            messagebox.showinfo(T("Historie", "Historie"), T(
                "Žádné záznamy.", "Keine Einträge."), parent=parent)
            return

        por = sorted(
            por,
            key=lambda p: (p.get("cas") or "", str(p.get("id") or "")),
            reverse=True,
        )
        poruchy_by_id = {
            str(p.get("id", "")).strip(): p
            for p in por
            if str(p.get("id", "")).strip()
        }

        def _zkrat(text, max_len=60):
            text = (text or "").strip()
            if len(text) <= max_len:
                return text or "-"
            return text[: max_len - 3].rstrip() + "..."

        def _text_do_sloupce(p):
            stav = (p.get("stav") or "").strip().lower()
            if stav == "otevrena":
                return p.get("popis", "") or ""
            return p.get("reseni", "") or ""

        def zobraz_detail_poruchy(iid):
            por_detail = poruchy_by_id.get(str(iid).strip())
            if not por_detail:
                return
            self._show_porucha_detail_dialog(win, por_detail)

        win = tk.Toplevel(parent)
        win.title(T("Historie poruch", "Störungshistorie"))
        win.transient(parent)
        win.grab_set()
        win.resizable(True, True)
        win.geometry("980x420")

        frm = tk.Frame(win, padx=12, pady=12)
        frm.pack(fill="both", expand=True)

        tk.Label(
            frm,
            text=T(
                "Všechny záznamy tohoto stroje. Nejnovější jsou první.",
                "Alle Einträge für diese Maschine. Neueste zuerst.",
            ),
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        cols = ("cas", "alarm", "stav", "text")
        tree_frame = tk.Frame(frm)
        tree_frame.grid(row=1, column=0, sticky="nsew")

        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=14)
        tree.heading("cas", text=T("Datum/čas", "Datum/Zeit"))
        tree.heading("alarm", text="Alarm")
        tree.heading("stav", text=T("Stav", "Status"))
        tree.heading("text", text=T("Text", "Text"))
        tree.column("cas", width=160, anchor="w", stretch=False)
        tree.column("alarm", width=130, anchor="w", stretch=False)
        tree.column("stav", width=110, anchor="w", stretch=False)
        tree.column("text", width=520, anchor="w", stretch=True)

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(1, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        for p in por:
            iid = str(p.get("id", "")).strip()
            if not iid:
                continue
            tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    p.get("cas", "") or "",
                    p.get("alarm", "") or "",
                    porucha_stav_ui(p.get("stav", "")),
                    _zkrat(_text_do_sloupce(p)),
                ),
            )

        def otevrit_detail_event(event):
            iid = tree.identify_row(event.y)
            if not iid:
                selection = tree.selection()
                iid = selection[0] if selection else ""
            if not iid:
                return
            tree.selection_set(iid)
            tree.focus(iid)
            zobraz_detail_poruchy(iid)

        tree.bind("<Double-1>", otevrit_detail_event)

        btns = tk.Frame(frm)
        btns.grid(row=2, column=0, sticky="e", pady=(12, 0))
        btn = tk.Button(btns, text=T("Zavřít", "Schließen"), command=win.destroy)
        btn.pack(side="right")

        win.bind("<Escape>", lambda e: win.destroy())
        win.bind("<Return>", lambda e: win.destroy())
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.after(0, tree.focus_set)

    def backup_zip(self):
        em.backup_zip(self)

    def restore_zip(self):
        if not em.restore_zip(self):
            return
        self.stroje = nacti_stroje()
        self.poruchy = nacti_poruchy()
        self.sablony = nacti_sablony()
        self.nakresli_mrizku()

    def kontrola_dat_gui(self):
        return validation_ui.open_data_validation_dialog(self)

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
            detail.geometry("620x420")

            frm_detail = tk.Frame(detail, padx=12, pady=12)
            frm_detail.pack(fill="both", expand=True)
            frm_detail.grid_columnconfigure(0, weight=1)
            frm_detail.grid_rowconfigure(0, weight=1)

            detail_text = "\n".join(txt)

            text_widget = tk.Text(frm_detail, wrap="word", height=16)
            scroll = ttk.Scrollbar(frm_detail, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scroll.set)
            text_widget.grid(row=0, column=0, sticky="nsew")
            scroll.grid(row=0, column=1, sticky="ns")
            text_widget.insert("1.0", detail_text)
            text_widget.configure(state="disabled")

            def copy_all():
                detail.clipboard_clear()
                detail.clipboard_append(detail_text)
                detail.update()
                messagebox.showinfo(
                    T("Informace", "Info"),
                    T(
                        "Zkopírováno do schránky.",
                        "In die Zwischenablage kopiert.",
                    ),
                    parent=detail,
                )

            def copy_selection(event=None):
                try:
                    selected = text_widget.get("sel.first", "sel.last")
                except tk.TclError:
                    selected = detail_text
                detail.clipboard_clear()
                detail.clipboard_append(selected)
                detail.update()
                return "break"

            def select_all(event=None):
                text_widget.tag_add("sel", "1.0", "end-1c")
                text_widget.mark_set("insert", "1.0")
                text_widget.see("insert")
                return "break"

            ctx = tk.Menu(detail, tearoff=False)
            ctx.add_command(label=T("Kopírovat", "Kopieren"), command=copy_selection)
            ctx.add_command(label=T("Vybrat vše", "Alles markieren"), command=select_all)

            def show_context(event):
                try:
                    ctx.tk_popup(event.x_root, event.y_root)
                finally:
                    ctx.grab_release()
                return "break"

            for seq in ("<Button-3>", "<Button-2>", "<Control-Button-1>"):
                text_widget.bind(seq, show_context)
            text_widget.bind("<Control-c>", copy_selection)

            btns = tk.Frame(frm_detail)
            btns.grid(row=1, column=0, columnspan=2, sticky="e", pady=(12, 0))
            tk.Button(
                btns, text=T("Zavřít", "Schließen"), command=detail.destroy
            ).pack(side="right")
            tk.Button(
                btns,
                text=T("Kopírovat do schránky", "In Zwischenablage kopieren"),
                command=copy_all,
            ).pack(side="right", padx=(0, 8))

            detail.bind("<Escape>", lambda e: detail.destroy())
            detail.after(0, text_widget.focus_set)

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

    def _porucha_foto_slozka(self, porucha: dict) -> Path:
        cislo = str(porucha.get("cislo", "")).strip() or "0"
        pid = str(porucha.get("id", "")).strip() or "0"
        return DATA_DIR / "photos" / f"stroj_{cislo}" / f"porucha_{pid}"

    def _porucha_foto_paths(self, porucha: dict) -> list[Path]:
        base = self._porucha_foto_slozka(porucha)
        if not base.is_dir():
            return []
        paths = []
        for ext in ("*.jpg", "*.jpeg", "*.png", "*.bmp", "*.webp"):
            paths.extend(base.glob(ext))
            paths.extend(base.glob(ext.upper()))
        return sorted(set(paths))

    def _otevrit_cestu_v_systemu(self, cesta: Path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(cesta))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{cesta}"')
            else:
                os.system(f'xdg-open "{cesta}"')
        except Exception as e:
            messagebox.showerror(
                T("Chyba", "Fehler"),
                f"{T('Nepodařilo se otevřít soubor', 'Datei konnte nicht geöffnet werden')}:\n{e}",
            )

    def _porucha_foto_hinzufuegen(self, parent, porucha: dict, on_change=None):
        src = filedialog.askopenfilename(
            parent=parent,
            title=T("Přidat fotografii", "Foto hinzufügen"),
            filetypes=[(T("Obrázky", "Bilder"), "*.jpg *.jpeg *.png *.bmp *.webp")],
        )
        if not src:
            return

        src_path = Path(src)
        suffix = src_path.suffix.lower()
        if suffix not in (".jpg", ".jpeg", ".png", ".bmp", ".webp"):
            suffix = ".jpg"

        dest_dir = self._porucha_foto_slozka(porucha)
        dest_dir.mkdir(parents=True, exist_ok=True)

        stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        idx = 1
        while True:
            dest_name = f"{stamp}_{idx:02d}{suffix}"
            dest_path = dest_dir / dest_name
            if not dest_path.exists():
                break
            idx += 1

        try:
            shutil.copy2(src_path, dest_path)
        except Exception as e:
            messagebox.showerror(T("Chyba", "Fehler"), f"{e}", parent=parent)
            return

        self._audit_event(
            "photo_added",
            "fault",
            entity_id=porucha.get("id", ""),
            machine_number=porucha.get("cislo", ""),
            details={"file": dest_name},
        )
        if on_change:
            on_change()
        messagebox.showinfo(
            T("Fotografie", "Foto"),
            T("Fotografie byla uložena.", "Foto gespeichert."),
            parent=parent,
        )

    def _porucha_fotos_oeffnen(self, parent, porucha: dict, on_change=None):
        photos = self._porucha_foto_paths(porucha)
        if not photos:
            messagebox.showinfo(
                T("Fotografie", "Foto"),
                T("Nejsou k dispozici žádné fotografie.", "Keine Fotos vorhanden."),
                parent=parent,
            )
            return

        win = tk.Toplevel(parent)
        win.title(T("Otevřít fotografie", "Fotos öffnen"))
        win.transient(parent)
        win.grab_set()
        win.resizable(True, True)
        win.geometry("520x320")

        frm = tk.Frame(win, padx=12, pady=12)
        frm.pack(fill="both", expand=True)
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(0, weight=1)

        listbox = tk.Listbox(frm)
        scroll = ttk.Scrollbar(frm, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scroll.set)
        listbox.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        for p in photos:
            listbox.insert("end", p.name)

        if photos:
            listbox.selection_set(0)

        def otevrit():
            sel = listbox.curselection()
            if not sel:
                return
            self._otevrit_cestu_v_systemu(photos[sel[0]])
            if on_change:
                on_change()

        btns = tk.Frame(frm)
        btns.grid(row=1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        tk.Button(btns, text=T("Zavřít", "Schließen"), command=win.destroy).pack(side="right")
        tk.Button(btns, text=T("Otevřít", "Öffnen"), command=otevrit).pack(side="right", padx=(0, 8))

        listbox.bind("<Double-1>", lambda e: otevrit())
        win.bind("<Escape>", lambda e: win.destroy())
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.after(0, listbox.focus_set)

    def _porucha_ordner_oeffnen(self, parent, porucha: dict):
        folder = self._porucha_foto_slozka(porucha)
        if not folder.is_dir():
            messagebox.showinfo(
                T("Složka", "Ordner"),
                T("Složka neexistuje.", "Kein Ordner vorhanden."),
                parent=parent,
            )
            return
        otevrit_slozku(folder)

    def _show_porucha_detail_dialog(self, parent, porucha: dict):
        detail = tk.Toplevel(parent)
        detail.title(T("Detail poruchy", "Detail der Störung"))
        detail.transient(parent)
        detail.resizable(False, False)

        win_w = 820
        win_h = 500
        detail.geometry(f"{win_w}x{win_h}")
        detail.update_idletasks()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        pos_x = parent_x + max((parent_w - win_w) // 2, 0)
        pos_y = parent_y + max((parent_h - win_h) // 2, 0)
        detail.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        status_text = porucha_stav_ui(porucha.get("stav", "")) or "-"
        status_key = (porucha.get("stav") or "").strip().lower()
        status_color = "#b00020" if status_key == "otevrena" else "#1f7a1f" if status_key == "uzavrena" else "black"
        wrap = 610

        frm_detail = tk.Frame(detail, padx=14, pady=14)
        frm_detail.pack(fill="both", expand=True)
        frm_detail.grid_columnconfigure(1, weight=1)

        foto_var = tk.StringVar()

        def refresh_photo_count():
            foto_var.set(f"Fotos: {len(self._porucha_foto_paths(porucha))}")

        radky = [
            (T("Datum/čas", "Datum/Zeit"), porucha.get("cas", "") or "-", None),
            ("Alarm", porucha.get("alarm", "") or "-", None),
            (T("Stav", "Status"), status_text, status_color),
            (T("Popis", "Beschreibung"), porucha.get("popis", "") or "-", None),
            (T("Řešení", "Lösung"), porucha.get("reseni", "") or "-", None),
        ]

        for idx, (label, value, color) in enumerate(radky):
            tk.Label(
                frm_detail,
                text=f"{label}:",
                font=("Segoe UI", 9, "bold"),
                anchor="nw",
            ).grid(row=idx, column=0, sticky="nw", padx=(0, 12), pady=5)
            tk.Label(
                frm_detail,
                text=value,
                justify="left",
                anchor="w",
                wraplength=wrap,
                fg=color or "black",
            ).grid(row=idx, column=1, sticky="w", pady=5)

        tk.Label(frm_detail, textvariable=foto_var, font=("Segoe UI", 9, "bold")).grid(
            row=len(radky), column=1, sticky="w", pady=(8, 2)
        )

        detail_text = "\n".join([
            f"{T('Datum/čas', 'Datum/Zeit')}: {porucha.get('cas', '') or '-'}",
            f"Alarm: {porucha.get('alarm', '') or '-'}",
            f"{T('Stav', 'Status')}: {status_text}",
            f"{T('Popis', 'Beschreibung')}: {porucha.get('popis', '') or '-'}",
            f"{T('Řešení', 'Lösung')}: {porucha.get('reseni', '') or '-'}",
        ])

        def kopirovat():
            detail.clipboard_clear()
            detail.clipboard_append(detail_text)
            detail.update()
            messagebox.showinfo(
                T("Informace", "Info"),
                T("Zkopírováno do schránky.", "In die Zwischenablage kopiert."),
                parent=detail,
            )

        photo_btns = tk.Frame(frm_detail)
        photo_btns.grid(row=len(radky) + 1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        tk.Button(
            photo_btns,
            text=T("Otevřít složku", "Ordner öffnen"),
            command=lambda: self._porucha_ordner_oeffnen(detail, porucha),
        ).pack(side="right")
        tk.Button(
            photo_btns,
            text=T("Otevřít fotografie", "Fotos öffnen"),
            command=lambda: self._porucha_fotos_oeffnen(detail, porucha, on_change=refresh_photo_count),
        ).pack(side="right", padx=(0, 8))
        tk.Button(
            photo_btns,
            text=T("Přidat fotografii", "Foto hinzufügen"),
            command=lambda: self._porucha_foto_hinzufuegen(detail, porucha, on_change=refresh_photo_count),
        ).pack(side="right", padx=(0, 8))

        btns = tk.Frame(frm_detail)
        btns.grid(row=len(radky) + 2, column=0, columnspan=2, sticky="e", pady=(16, 0))
        btn = tk.Button(
            btns, text=T("Zavřít", "Schließen"), command=detail.destroy
        )
        btn.pack(side="right")
        copy_btn = tk.Button(
            btns,
            text=T("Kopírovat do schránky", "In Zwischenablage kopieren"),
            command=kopirovat,
        )
        copy_btn.pack(side="right", padx=(0, 8))

        refresh_photo_count()
        detail.bind("<Escape>", lambda e: detail.destroy())
        detail.bind("<Return>", lambda e: detail.destroy())
        detail.protocol("WM_DELETE_WINDOW", detail.destroy)
        detail.after(0, btn.focus_set)

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

    # === Hromadné uzavření a bezpečná archivace strojů ===
    def hromadne_uzavrit(self, parent):
        ids = bulk_uzavrit_dialog(parent, self.poruchy)
        if ids is None:
            return
        if not ids:
            messagebox.showinfo(T("Hromadné uzavření", "Sammelschließung"),
                                T("Nic nebylo vybráno.", "Nichts ausgewählt."), parent=parent)
            return
        n = 0
        closed_ids = []
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for p in self.poruchy:
            if p.get("id") in ids and p.get("stav") == "otevrena":
                p["stav"] = "uzavrena"
                p["cas_uzavreni"] = now
                closed_ids.append(str(p.get("id", "")))
                n += 1
        uloz_poruchy(self.poruchy)
        if closed_ids:
            self._audit_event(
                "faults_bulk_closed",
                "fault",
                details={"fault_ids": closed_ids, "closed_at": now},
            )
        self.nakresli_mrizku()
        messagebox.showinfo(
            T("Hotovo", "Fertig"), f"{T('Uzavřeno', 'Geschlossen')} {n} {T('záznamů', 'Einträge')}.", parent=parent)

    def archivovat_stroj_gui(self, cislo: str):
        if cislo not in self.stroje:
            messagebox.showwarning(T("Archivovat stroj", "Maschine archivieren"),
                                   f"{T('Stroj', 'Maschine')} {cislo} {T('nenalezen', 'nicht gefunden')}.", parent=self)
            return
        stroj = self.stroje[cislo]
        if dm.is_archived_machine(stroj):
            return
        if not messagebox.askyesno(
            T("Archivovat stroj", "Maschine archivieren"),
            T(
                f"Archivovat stroj {cislo}?\n\n"
                "Poruchy, fotografie a dokumenty zůstanou zachované. "
                "Stroj lze kdykoli obnovit.",
                f"Maschine {cislo} archivieren?\n\n"
                "Störungen, Fotos und Dokumente bleiben erhalten. "
                "Die Maschine kann jederzeit wiederhergestellt werden.",
            ),
            parent=self,
        ):
            return
        before = dict(stroj)
        stroj["archivovan"] = "1"
        uloz_stroje(self.stroje)
        self._audit_event(
            "machine_archived",
            "machine",
            entity_id=cislo,
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, stroj)},
        )
        if self.last_selected == cislo:
            self.last_selected = None
        self.nakresli_mrizku()
        messagebox.showinfo(
            T("Hotovo", "Fertig"),
            T(
                "Stroj byl archivován; jeho historie a soubory byly zachovány.",
                "Maschine wurde archiviert; Verlauf und Dateien bleiben erhalten.",
            ),
            parent=self,
        )

    def obnovit_stroj_z_archivu(self, cislo: str):
        stroj = self.stroje.get(cislo)
        if not stroj or not dm.is_archived_machine(stroj):
            return
        before = dict(stroj)
        stroj["archivovan"] = "0"
        uloz_stroje(self.stroje)
        self._audit_event(
            "machine_restored",
            "machine",
            entity_id=cislo,
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, stroj)},
        )
        self.nakresli_mrizku()
        messagebox.showinfo(
            T("Hotovo", "Fertig"),
            T("Stroj byl obnoven z archivu.", "Maschine wurde aus dem Archiv wiederhergestellt."),
            parent=self,
        )

    def smazat_stroj_gui(self, cislo: str):
        """Zpětně kompatibilní název; stroje se již fyzicky nemažou."""
        return self.archivovat_stroj_gui(cislo)

    def editovat_otevrenou_poruchu(self, parent, cislo: str):
        por = nacti_poruchy()
        opened = [p for p in por if p.get("cislo") == cislo and p.get("stav") == "otevrena"]
        if not opened:
            messagebox.showinfo(T("Editace", "Bearbeiten"), T(
                "Žádná otevřená porucha u tohoto stroje.", "Keine offene Störung an dieser Maschine."), parent=parent)
            return

        def vyber_pro_editaci(parent_win, opened_items):
            opened_sorted = sorted(
                opened_items,
                key=lambda p: (p.get("cas") or "", str(p.get("id") or "")),
                reverse=True,
            )
            by_id = {str(p.get("id", "")).strip(): p for p in opened_sorted if str(p.get("id", "")).strip()}

            def _zkrat(text, max_len=60):
                text = (text or "").strip()
                if len(text) <= max_len:
                    return text or "-"
                return text[: max_len - 3].rstrip() + "..."

            dlg = tk.Toplevel(parent_win)
            dlg.title(T("Vybrat otevřenou poruchu", "Offene Störung auswählen"))
            dlg.transient(parent_win)
            dlg.grab_set()
            dlg.resizable(True, True)
            dlg.geometry("920x360")

            frm = tk.Frame(dlg, padx=12, pady=12)
            frm.pack(fill="both", expand=True)
            frm.grid_columnconfigure(0, weight=1)
            frm.grid_rowconfigure(1, weight=1)

            tk.Label(
                frm,
                text=T(
                    "Vyberte otevřenou poruchu k úpravě.",
                    "Bitte eine offene Störung zur Bearbeitung auswählen.",
                ),
                font=("Segoe UI", 10, "bold"),
            ).grid(row=0, column=0, sticky="w", pady=(0, 10))

            cols = ("id", "cas", "kategorie", "alarm", "popis")
            tree_frame = tk.Frame(frm)
            tree_frame.grid(row=1, column=0, sticky="nsew")
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)

            tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10)
            tree.heading("id", text="ID")
            tree.heading("cas", text=T("Datum/čas", "Datum/Zeit"))
            tree.heading("kategorie", text=T("Kategorie", "Kategorie"))
            tree.heading("alarm", text="Alarm")
            tree.heading("popis", text=T("Popis", "Beschreibung"))
            tree.column("id", width=60, anchor="center", stretch=False)
            tree.column("cas", width=150, anchor="w", stretch=False)
            tree.column("kategorie", width=120, anchor="w", stretch=False)
            tree.column("alarm", width=120, anchor="w", stretch=False)
            tree.column("popis", width=380, anchor="w", stretch=True)

            yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
            tree.grid(row=0, column=0, sticky="nsew")
            yscroll.grid(row=0, column=1, sticky="ns")
            xscroll.grid(row=1, column=0, sticky="ew")

            for p in opened_sorted:
                iid = str(p.get("id", "")).strip()
                if not iid:
                    continue
                kat_key = normalize_kategorie(p.get("kategorie", "") or "")
                tree.insert(
                    "",
                    "end",
                    iid=iid,
                    values=(
                        iid,
                        p.get("cas", "") or "",
                        kat_ui(kat_key),
                        p.get("alarm", "") or "",
                        _zkrat(p.get("popis", "") or ""),
                    ),
                )

            if opened_sorted:
                first_id = str(opened_sorted[0].get("id", "")).strip()
                if first_id:
                    tree.selection_set(first_id)
                    tree.focus(first_id)

            result = {"value": None}

            def potvrdit():
                selection = tree.selection()
                if not selection:
                    messagebox.showwarning(
                        T("Upozornění", "Warnung"),
                        T("Vyberte poruchu.", "Bitte eine Störung auswählen."),
                        parent=dlg,
                    )
                    return
                result["value"] = by_id.get(selection[0])
                dlg.destroy()

            def zrusit():
                dlg.destroy()

            def double_click(event):
                iid = tree.identify_row(event.y)
                if iid:
                    tree.selection_set(iid)
                    tree.focus(iid)
                potvrdit()

            btns = tk.Frame(frm)
            btns.grid(row=2, column=0, sticky="e", pady=(12, 0))
            ttk.Button(
                btns, text=T("Zrušit", "Abbrechen"), width=12, command=zrusit
            ).pack(side="right")
            ttk.Button(
                btns, text=T("Upravit", "Bearbeiten"), width=12, command=potvrdit
            ).pack(side="right", padx=(0, 8))

            tree.bind("<Double-1>", double_click)
            dlg.bind("<Return>", lambda e: potvrdit())
            dlg.bind("<Escape>", lambda e: zrusit())
            dlg.protocol("WM_DELETE_WINDOW", zrusit)
            dlg.after(0, tree.focus_set)
            parent_win.wait_window(dlg)
            return result["value"]

        target = vyber_pro_editaci(parent, opened)
        if target is None:
            return
        before = dict(target)

        new_alarm = simpledialog.askstring(T("Editace", "Bearbeiten"), T(
            "Alarm:", "Alarm:"), initialvalue=target.get("alarm", ""), parent=parent)
        if new_alarm is None:
            return
        new_kat = ask_kategorie_combobox(parent)
        if new_kat is None:
            return

        def ask_beschreibung_dialog(parent_win, initial_text=""):
            dlg = tk.Toplevel(parent_win)
            dlg.title(T("Upravit otevřenou poruchu", "Offene Störung bearbeiten"))
            dlg.transient(parent_win)
            dlg.grab_set()
            dlg.resizable(True, True)

            win_w = 580
            win_h = 300
            dlg.geometry(f"{win_w}x{win_h}")
            dlg.update_idletasks()
            parent_x = parent_win.winfo_rootx()
            parent_y = parent_win.winfo_rooty()
            parent_w = parent_win.winfo_width()
            parent_h = parent_win.winfo_height()
            pos_x = parent_x + max((parent_w - win_w) // 2, 0)
            pos_y = parent_y + max((parent_h - win_h) // 2, 0)
            dlg.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

            frm = tk.Frame(dlg, padx=12, pady=12)
            frm.pack(fill="both", expand=True)
            frm.grid_columnconfigure(0, weight=1)
            frm.grid_rowconfigure(1, weight=1)

            tk.Label(frm, text=T("Popis:", "Beschreibung:")).grid(
                row=0, column=0, sticky="w", pady=(0, 6)
            )

            text_frame = tk.Frame(frm)
            text_frame.grid(row=1, column=0, sticky="nsew")
            text_frame.grid_columnconfigure(0, weight=1)
            text_frame.grid_rowconfigure(0, weight=1)

            txt = tk.Text(text_frame, height=7, wrap="word")
            scroll = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=scroll.set)
            txt.grid(row=0, column=0, sticky="nsew")
            scroll.grid(row=0, column=1, sticky="ns")
            if initial_text:
                txt.insert("1.0", initial_text)

            result = {"value": None}

            def potvrdit():
                result["value"] = txt.get("1.0", "end-1c").strip()
                dlg.destroy()

            def zrusit():
                dlg.destroy()

            btns = tk.Frame(frm)
            btns.grid(row=2, column=0, sticky="e", pady=(12, 0))
            tk.Button(btns, text=T("Zrušit", "Abbrechen"), command=zrusit).pack(side="right")
            tk.Button(btns, text="OK", command=potvrdit, width=10).pack(side="right", padx=(0, 8))

            dlg.bind("<Escape>", lambda e: zrusit())
            dlg.bind("<Control-Return>", lambda e: potvrdit())
            dlg.protocol("WM_DELETE_WINDOW", zrusit)
            dlg.after(0, txt.focus_set)
            parent_win.wait_window(dlg)
            return result["value"]

        new_pop = ask_beschreibung_dialog(parent, target.get("popis", ""))
        if new_pop is None:
            return

        for p in por:
            if p.get("id") == target.get("id"):
                p["alarm"] = new_alarm
                p["kategorie"] = new_kat
                p["popis"] = new_pop
                break
        uloz_poruchy(por)
        self.poruchy = por
        self._audit_event(
            "fault_updated",
            "fault",
            entity_id=target.get("id", ""),
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, target)},
        )
        messagebox.showinfo(T("Uloženo", "Gespeichert"),
                            f"{T('Porucha', 'Störung')} ID {target.get('id')} {T('upravena', 'bearbeitet')}.", parent=parent)

    def korigovat_uzavrenou_poruchu(self, parent, cislo: str):
        def ask_continue():
            dlg = tk.Toplevel(parent)
            dlg.title(T("Upozornění", "Achtung"))
            dlg.transient(parent)
            dlg.grab_set()
            dlg.resizable(False, False)

            frm = tk.Frame(dlg, padx=16, pady=14)
            frm.pack(fill="both", expand=True)
            tk.Label(
                frm,
                text=T(
                    "Upravujete již uzavřenou poruchu.\n"
                    "Opravujte pouze chyby nebo neúplné záznamy.",
                    "Sie bearbeiten eine bereits geschlossene Störung.\n"
                    "Bitte nur Fehler oder unvollständige Einträge korrigieren.",
                ),
                justify="left",
                wraplength=460,
            ).pack(anchor="w", pady=(0, 14))

            result = {"ok": False}

            def weiter():
                result["ok"] = True
                dlg.destroy()

            def abbrechen():
                dlg.destroy()

            btns = tk.Frame(frm)
            btns.pack(anchor="e")
            ttk.Button(btns, text=T("Zrušit", "Abbrechen"), width=12, command=abbrechen).pack(side="right")
            ttk.Button(btns, text=T("Pokračovat", "Weiter"), width=12, command=weiter).pack(side="right", padx=(0, 8))

            dlg.bind("<Return>", lambda e: weiter())
            dlg.bind("<Escape>", lambda e: abbrechen())
            dlg.protocol("WM_DELETE_WINDOW", abbrechen)
            center_over(dlg, parent)
            parent.wait_window(dlg)
            return result["ok"]

        if not ask_continue():
            return

        por = nacti_poruchy()
        closed = [
            p for p in por
            if p.get("cislo") == str(cislo)
            and str(p.get("stav", "")).strip().lower() in ("uzavrena", "geschlossen", "g")
        ]
        if not closed:
            messagebox.showinfo(
                T("Vybrat uzavřenou poruchu", "Geschlossene Störung auswählen"),
                T(
                    "Tento stroj nemá žádnou uzavřenou poruchu.",
                    "Keine geschlossene Störung an dieser Maschine.",
                ),
                parent=parent,
            )
            return
        def vyber_uzavrenou(parent_win, closed_items):
            def _dt_key(p):
                for field in ("cas_uzavreni", "cas"):
                    value = (p.get(field) or "").strip()
                    if not value:
                        continue
                    try:
                        return datetime.strptime(value, "%Y-%m-%d %H:%M")
                    except Exception:
                        pass
                return datetime.min

            def _zkrat(text, max_len=70):
                text = (text or "").strip()
                if len(text) <= max_len:
                    return text or "-"
                return text[: max_len - 3].rstrip() + "..."

            closed_sorted = sorted(
                closed_items,
                key=lambda p: (_dt_key(p), str(p.get("id") or "")),
                reverse=True,
            )
            by_id = {str(p.get("id", "")).strip(): p for p in closed_sorted if str(p.get("id", "")).strip()}

            dlg = tk.Toplevel(parent_win)
            dlg.title(T("Vybrat uzavřenou poruchu", "Geschlossene Störung auswählen"))
            dlg.transient(parent_win)
            dlg.grab_set()
            dlg.resizable(True, True)
            dlg.geometry("1180x420")

            frm = tk.Frame(dlg, padx=12, pady=12)
            frm.pack(fill="both", expand=True)
            frm.grid_columnconfigure(0, weight=1)
            frm.grid_rowconfigure(1, weight=1)

            tk.Label(
                frm,
                text=T(
                    "Vyberte uzavřenou poruchu k opravě.",
                    "Bitte eine geschlossene Störung zur Korrektur auswählen.",
                ),
                font=("Segoe UI", 10, "bold"),
            ).grid(row=0, column=0, sticky="w", pady=(0, 10))

            cols = ("id", "cas", "alarm", "kategorie", "popis", "reseni", "cas_uzavreni")
            tree_frame = tk.Frame(frm)
            tree_frame.grid(row=1, column=0, sticky="nsew")
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)

            tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=12)
            headings = {
                "id": "ID",
                "cas": T("Datum/čas", "Datum/Zeit"),
                "alarm": "Alarm",
                "kategorie": T("Kategorie", "Kategorie"),
                "popis": T("Popis", "Beschreibung"),
                "reseni": T("Řešení", "Lösung"),
                "cas_uzavreni": T("Uzavřeno", "Geschlossen am"),
            }
            widths = {
                "id": 60,
                "cas": 145,
                "alarm": 120,
                "kategorie": 115,
                "popis": 280,
                "reseni": 280,
                "cas_uzavreni": 145,
            }
            for col in cols:
                tree.heading(col, text=headings[col])
                tree.column(col, width=widths[col], anchor=("center" if col == "id" else "w"), stretch=col in ("popis", "reseni"))

            yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
            tree.grid(row=0, column=0, sticky="nsew")
            yscroll.grid(row=0, column=1, sticky="ns")
            xscroll.grid(row=1, column=0, sticky="ew")

            for p in closed_sorted:
                iid = str(p.get("id", "")).strip()
                if not iid:
                    continue
                tree.insert(
                    "",
                    "end",
                    iid=iid,
                    values=(
                        iid,
                        p.get("cas", "") or "",
                        p.get("alarm", "") or "",
                        kat_ui(p.get("kategorie", "") or ""),
                        _zkrat(p.get("popis", "") or ""),
                        _zkrat(p.get("reseni", "") or ""),
                        p.get("cas_uzavreni", "") or "",
                    ),
                )

            if closed_sorted:
                first_id = str(closed_sorted[0].get("id", "")).strip()
                if first_id:
                    tree.selection_set(first_id)
                    tree.focus(first_id)

            result = {"value": None}

            def korrigieren():
                selection = tree.selection()
                if not selection:
                    messagebox.showwarning(
                        T("Upozornění", "Warnung"),
                        T("Vyberte poruchu.", "Bitte eine Störung auswählen."),
                        parent=dlg,
                    )
                    return
                result["value"] = by_id.get(selection[0])
                dlg.destroy()

            def abbrechen():
                dlg.destroy()

            def double_click(event):
                iid = tree.identify_row(event.y)
                if iid:
                    tree.selection_set(iid)
                    tree.focus(iid)
                korrigieren()

            btns = tk.Frame(frm)
            btns.grid(row=2, column=0, sticky="e", pady=(12, 0))
            ttk.Button(
                btns, text=T("Zrušit", "Abbrechen"), width=12, command=abbrechen
            ).pack(side="right")
            ttk.Button(
                btns, text=T("Opravit", "Korrigieren"), width=12, command=korrigieren
            ).pack(side="right", padx=(0, 8))

            tree.bind("<Double-1>", double_click)
            dlg.bind("<Return>", lambda e: korrigieren())
            dlg.bind("<Escape>", lambda e: abbrechen())
            dlg.protocol("WM_DELETE_WINDOW", abbrechen)
            dlg.after(0, tree.focus_set)
            center_over(dlg, parent_win)
            parent_win.wait_window(dlg)
            return result["value"]

        target = vyber_uzavrenou(parent, closed)
        if target is None:
            return
        before = dict(target)

        def edit_dialog(parent_win, porucha):
            dlg = tk.Toplevel(parent_win)
            dlg.title(T("Opravit uzavřenou poruchu", "Geschlossene Störung korrigieren"))
            dlg.transient(parent_win)
            dlg.grab_set()
            dlg.resizable(True, True)
            dlg.geometry("820x620")

            frm = tk.Frame(dlg, padx=12, pady=12)
            frm.pack(fill="both", expand=True)
            frm.grid_columnconfigure(1, weight=1)
            frm.grid_rowconfigure(4, weight=1)
            frm.grid_rowconfigure(5, weight=1)

            def add_label(row, text):
                tk.Label(frm, text=text).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)

            add_label(0, T("Stav", "Status"))
            tk.Label(frm, text=T("uzavřena", "geschlossen"), anchor="w").grid(
                row=0, column=1, sticky="ew", pady=4
            )

            add_label(1, "Alarm")
            alarm_var = tk.StringVar(value=porucha.get("alarm", "") or "")
            alarm_entry = ttk.Entry(frm, textvariable=alarm_var)
            alarm_entry.grid(row=1, column=1, sticky="ew", pady=4)

            add_label(2, T("Kategorie", "Kategorie"))
            kat_to_key = {
                T("Elektrická", "Elektrisch"): "elektricka",
                T("Mechanická", "Mechanisch"): "mechanicka",
                T("Jiná", "Sonstige"): "jina",
            }
            kat_values = list(kat_to_key)
            key_to_kat = {v: k for k, v in kat_to_key.items()}
            kat_var = tk.StringVar(value=key_to_kat.get(
                normalize_kategorie(porucha.get("kategorie", "")),
                T("Jiná", "Sonstige"),
            ))
            kat_cb = ttk.Combobox(frm, textvariable=kat_var, state="readonly", values=kat_values)
            kat_cb.grid(row=2, column=1, sticky="w", pady=4)

            add_label(3, T("Uzavřel operátor", "Operator geschlossen"))
            operator_var = tk.StringVar(value=porucha.get("operator_uzavrel", "") or "")
            ttk.Entry(frm, textvariable=operator_var).grid(row=3, column=1, sticky="ew", pady=4)

            def add_text(row, label, initial):
                add_label(row, label)
                box_frame = tk.Frame(frm)
                box_frame.grid(row=row, column=1, sticky="nsew", pady=4)
                box_frame.grid_rowconfigure(0, weight=1)
                box_frame.grid_columnconfigure(0, weight=1)
                txt = tk.Text(box_frame, height=7, wrap="word")
                scroll = ttk.Scrollbar(box_frame, orient="vertical", command=txt.yview)
                txt.configure(yscrollcommand=scroll.set)
                txt.grid(row=0, column=0, sticky="nsew")
                scroll.grid(row=0, column=1, sticky="ns")
                txt.insert("1.0", initial or "")
                return txt

            popis_txt = add_text(
                4, T("Popis", "Beschreibung"), porucha.get("popis", "") or ""
            )
            reseni_txt = add_text(
                5, T("Řešení", "Lösung"), porucha.get("reseni", "") or ""
            )

            result = {"value": None}

            def ask_save():
                confirm = tk.Toplevel(dlg)
                confirm.title(T("Uložit změny?", "Änderungen speichern?"))
                confirm.transient(dlg)
                confirm.grab_set()
                confirm.resizable(False, False)

                cfrm = tk.Frame(confirm, padx=16, pady=14)
                cfrm.pack(fill="both", expand=True)
                tk.Label(
                    cfrm,
                    text=T(
                        "Tato uzavřená porucha bude změněna.\n"
                        "Chcete změny uložit?",
                        "Diese geschlossene Störung wird geändert.\n"
                        "Möchten Sie die Änderungen speichern?",
                    ),
                    justify="left",
                    wraplength=420,
                ).pack(anchor="w", pady=(0, 14))

                answer = {"ok": False}

                def speichern():
                    answer["ok"] = True
                    confirm.destroy()

                def abbrechen_confirm():
                    confirm.destroy()

                cbtns = tk.Frame(cfrm)
                cbtns.pack(anchor="e")
                ttk.Button(
                    cbtns,
                    text=T("Zrušit", "Abbrechen"),
                    width=12,
                    command=abbrechen_confirm,
                ).pack(side="right")
                ttk.Button(
                    cbtns,
                    text=T("Uložit", "Speichern"),
                    width=12,
                    command=speichern,
                ).pack(side="right", padx=(0, 8))
                confirm.bind("<Return>", lambda e: speichern())
                confirm.bind("<Escape>", lambda e: abbrechen_confirm())
                confirm.protocol("WM_DELETE_WINDOW", abbrechen_confirm)
                center_over(confirm, dlg)
                dlg.wait_window(confirm)
                return answer["ok"]

            def speichern():
                if not ask_save():
                    return
                result["value"] = {
                    "alarm": alarm_var.get().strip(),
                    "kategorie": kat_to_key.get(kat_var.get(), "jina"),
                    "popis": popis_txt.get("1.0", "end-1c").strip(),
                    "reseni": reseni_txt.get("1.0", "end-1c").strip(),
                    "operator_uzavrel": operator_var.get().strip(),
                }
                dlg.destroy()

            def abbrechen():
                dlg.destroy()

            btns = tk.Frame(frm)
            btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(12, 0))
            ttk.Button(
                btns, text=T("Zrušit", "Abbrechen"), width=12, command=abbrechen
            ).pack(side="right")
            ttk.Button(
                btns, text=T("Uložit", "Speichern"), width=12, command=speichern
            ).pack(side="right", padx=(0, 8))

            dlg.bind("<Control-Return>", lambda e: speichern())
            dlg.bind("<Escape>", lambda e: abbrechen())
            dlg.protocol("WM_DELETE_WINDOW", abbrechen)
            dlg.after(0, alarm_entry.focus_set)
            center_over(dlg, parent_win)
            parent_win.wait_window(dlg)
            return result["value"]

        values = edit_dialog(parent, target)
        if values is None:
            return

        target_id = str(target.get("id", "")).strip()
        updated = False
        for p in por:
            if str(p.get("id", "")).strip() == target_id:
                p["alarm"] = values["alarm"]
                p["kategorie"] = values["kategorie"]
                p["popis"] = values["popis"]
                p["reseni"] = values["reseni"]
                p["operator_uzavrel"] = values["operator_uzavrel"]
                p["stav"] = "uzavrena"
                updated = True
                break

        if not updated:
            messagebox.showwarning(
                T("Upozornění", "Warnung"),
                T(
                    "Vybraná porucha nebyla nalezena.",
                    "Ausgewählte Störung wurde nicht gefunden.",
                ),
                parent=parent,
            )
            return

        uloz_poruchy(por)
        self.poruchy = por
        self._audit_event(
            "fault_updated",
            "fault",
            entity_id=target_id,
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, target)},
        )
        messagebox.showinfo(
            T("Uloženo", "Gespeichert"),
            T(
                f"Porucha ID {target_id} byla opravena.",
                f"Störung ID {target_id} wurde korrigiert.",
            ),
            parent=parent,
        )

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
            "stav": stav,
            "wartung_last": "",
            "wartung_interval": "180",
            "archivovan": "0",
        }
        uloz_stroje(self.stroje)
        self._audit_event(
            "machine_created",
            "machine",
            entity_id=cislo,
            machine_number=cislo,
            details={"fields": audit.changed_fields({}, self.stroje[cislo])},
        )
        self.nakresli_mrizku()
        messagebox.showinfo(T("Přidat stroj", "Maschine hinzufügen"),
                            f"{T('Stroj', 'Maschine')} {cislo} {T('uložen', 'gespeichert')}.", parent=self)

    # === Editace libovolné otevřené poruchy u stroje ===
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
        if dm.is_archived_machine(s):
            messagebox.showinfo(
                T("Archivovaný stroj", "Archivierte Maschine"),
                T(
                    "Stav archivovaného stroje nelze přepnout.",
                    "Der Status einer archivierten Maschine kann nicht geändert werden.",
                ),
                parent=self,
            )
            return
        before = dict(s)
        cur = normalize_stav(s.get("stav", "bezi"))
        s["stav"] = "porucha" if cur == "bezi" else "bezi"
        uloz_stroje(self.stroje)
        self._audit_event(
            "machine_status_changed",
            "machine",
            entity_id=cislo,
            machine_number=cislo,
            details={"fields": audit.changed_fields(before, s)},
        )
        self.nakresli_mrizku()


def vyber_otevrenou_poruchu_combo(parent, opened: list):
    """
    Vrátí vybraný záznam poruchy (dict) z listu 'opened' pomocí Treeview.
    Pokud uživatel zruší, vrací None.
    """
    if not opened:
        return None

    opened_sorted = sorted(
        opened,
        key=lambda p: (p.get("cas") or "", str(p.get("id") or "")),
        reverse=True,
    )
    by_id = {str(p.get("id", "")).strip(): p for p in opened_sorted if str(p.get("id", "")).strip()}

    def _zkrat(text, max_len=60):
        text = (text or "").strip()
        if len(text) <= max_len:
            return text or "-"
        return text[: max_len - 3].rstrip() + "..."

    top = tk.Toplevel(parent)
    top.title(T("Uzavřít poruchu", "Störung schließen"))
    top.transient(parent)
    top.grab_set()
    top.resizable(True, True)
    top.geometry("920x360")

    frm = tk.Frame(top, padx=12, pady=12)
    frm.pack(fill="both", expand=True)
    frm.grid_columnconfigure(0, weight=1)
    frm.grid_rowconfigure(1, weight=1)

    tk.Label(
        frm,
        text=T(
            "Vyberte otevřenou poruchu k uzavření.",
            "Bitte eine offene Störung zum Schließen auswählen.",
        ),
        font=("Segoe UI", 10, "bold"),
    ).grid(row=0, column=0, sticky="w", pady=(0, 10))

    cols = ("id", "cas", "kategorie", "alarm", "popis")
    tree_frame = tk.Frame(frm)
    tree_frame.grid(row=1, column=0, sticky="nsew")
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)

    tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10)
    tree.heading("id", text="ID")
    tree.heading("cas", text=T("Datum/čas", "Datum/Zeit"))
    tree.heading("kategorie", text=T("Kategorie", "Kategorie"))
    tree.heading("alarm", text="Alarm")
    tree.heading("popis", text=T("Popis", "Beschreibung"))
    tree.column("id", width=60, anchor="center", stretch=False)
    tree.column("cas", width=150, anchor="w", stretch=False)
    tree.column("kategorie", width=120, anchor="w", stretch=False)
    tree.column("alarm", width=120, anchor="w", stretch=False)
    tree.column("popis", width=380, anchor="w", stretch=True)

    yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    tree.grid(row=0, column=0, sticky="nsew")
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll.grid(row=1, column=0, sticky="ew")

    for p in opened_sorted:
        iid = str(p.get("id", "")).strip()
        if not iid:
            continue
        kat_key = normalize_kategorie(p.get("kategorie", "") or "")
        tree.insert(
            "",
            "end",
            iid=iid,
            values=(
                iid,
                p.get("cas", "") or "",
                kat_ui(kat_key),
                p.get("alarm", "") or "",
                _zkrat(p.get("popis", "") or ""),
            ),
        )

    if opened_sorted:
        first_id = str(opened_sorted[0].get("id", "")).strip()
        if first_id:
            tree.selection_set(first_id)
            tree.focus(first_id)

    chosen = {"value": None}

    def _ok():
        selection = tree.selection()
        if not selection:
            messagebox.showwarning(
                T("Upozornění", "Warnung"),
                T("Vyberte poruchu.", "Bitte eine Störung auswählen."),
                parent=top,
            )
            return
        chosen["value"] = by_id.get(selection[0])
        top.destroy()

    def _cancel():
        top.destroy()

    def _double_click(event):
        iid = tree.identify_row(event.y)
        if iid:
            tree.selection_set(iid)
            tree.focus(iid)
        _ok()

    btns = tk.Frame(frm)
    btns.grid(row=2, column=0, sticky="e", pady=(12, 0))
    ttk.Button(btns, text=T("Zrušit", "Abbrechen"), width=12, command=_cancel).pack(side="right")
    ttk.Button(btns, text=T("Uzavřít", "Schließen"), width=12, command=_ok).pack(side="right", padx=(0, 8))

    top.bind("<Return>", lambda e: _ok())
    top.bind("<Escape>", lambda e: _cancel())
    top.protocol("WM_DELETE_WINDOW", _cancel)
    tree.bind("<Double-1>", _double_click)
    top.after(0, tree.focus_set)
    parent.wait_window(top)
    return chosen["value"]


# ===== Main =====
if __name__ == "__main__":
    configure_logging(DATA_DIR)
    app = StrojeGrid()
    install_tk_exception_handler(app, T)
    if app.vybrat_operatora_gui(required=True):
        app.mainloop()
    else:
        app.destroy()
