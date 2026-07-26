#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

import data_manager as dm
import data_validator as dv
import export_manager as em


T = dm.T


def open_data_validation_dialog(app):
    """Otevře kontrolu dat a vrátí vytvořené okno."""
    win = tk.Toplevel(app)
    win.title(T("Kontrola dat", "Datenprüfung"))
    win.geometry("1040x620")
    win.minsize(820, 440)
    win.transient(app)

    outer = tk.Frame(win, padx=12, pady=12)
    outer.pack(fill="both", expand=True)
    outer.grid_columnconfigure(0, weight=1)
    outer.grid_rowconfigure(2, weight=1)

    tk.Label(
        outer,
        text=T("Kontrola konzistence dat", "Prüfung der Datenkonsistenz"),
        font=("Segoe UI", 14, "bold"),
    ).grid(row=0, column=0, sticky="w")

    summary_var = tk.StringVar()
    tk.Label(
        outer,
        textvariable=summary_var,
        anchor="w",
        justify="left",
    ).grid(row=1, column=0, sticky="ew", pady=(6, 10))

    table_frame = tk.Frame(outer)
    table_frame.grid(row=2, column=0, sticky="nsew")
    table_frame.grid_columnconfigure(0, weight=1)
    table_frame.grid_rowconfigure(0, weight=1)

    columns = ("severity", "file", "row", "message")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings")
    tree.heading("severity", text=T("Úroveň", "Stufe"))
    tree.heading("file", text=T("Soubor", "Datei"))
    tree.heading("row", text=T("Řádek", "Zeile"))
    tree.heading("message", text=T("Popis", "Beschreibung"))
    tree.column("severity", width=90, stretch=False, anchor="center")
    tree.column("file", width=145, stretch=False)
    tree.column("row", width=65, stretch=False, anchor="center")
    tree.column("message", width=690)
    tree.tag_configure("error", background="#ffd7d7")
    tree.tag_configure("warning", background="#fff2c7")
    tree.tag_configure("ok", background="#dff4df")

    yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    tree.grid(row=0, column=0, sticky="nsew")
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll.grid(row=1, column=0, sticky="ew")

    buttons = tk.Frame(outer)
    buttons.grid(row=3, column=0, sticky="ew", pady=(12, 0))
    current = {"report": None}

    def refresh():
        report = dv.validate_data(dm.DATA_DIR)
        current["report"] = report
        tree.delete(*tree.get_children())
        summary_var.set(
            T(
                f"Stroje: {report.machine_count}   Poruchy: {report.fault_count}   "
                f"Chyby: {report.errors}   Varování: {report.warnings}   "
                f"Bezpečně opravitelné: {report.fixable}",
                f"Maschinen: {report.machine_count}   Störungen: {report.fault_count}   "
                f"Fehler: {report.errors}   Warnungen: {report.warnings}   "
                f"Sicher korrigierbar: {report.fixable}",
            )
        )
        if not report.issues:
            tree.insert(
                "",
                "end",
                values=(
                    "OK",
                    "",
                    "",
                    T(
                        "Nebyly nalezeny žádné problémy.",
                        "Es wurden keine Probleme gefunden.",
                    ),
                ),
                tags=("ok",),
            )
        else:
            for issue in report.issues:
                level = (
                    T("Chyba", "Fehler")
                    if issue.severity == "error"
                    else T("Varování", "Warnung")
                )
                message = T(issue.message_cz, issue.message_de)
                if issue.fixable:
                    message += T(
                        "  [bezpečně opravitelné]",
                        "  [sicher korrigierbar]",
                    )
                tree.insert(
                    "",
                    "end",
                    values=(
                        level,
                        issue.filename,
                        issue.row or "",
                        message,
                    ),
                    tags=(issue.severity,),
                )
        repair_button.configure(
            state=("normal" if report.fixable else "disabled"),
            text=T(
                f"Bezpečně opravit ({report.fixable})",
                f"Sicher korrigieren ({report.fixable})",
            ),
        )

    def repair():
        report = current.get("report")
        if report is None or not report.fixable:
            return
        if not messagebox.askyesno(
            T("Bezpečná oprava dat", "Sichere Datenkorrektur"),
            T(
                f"Bude provedeno {report.fixable} jednoznačných normalizací.\n"
                "Duplicity, neplatné vazby ani nejasné hodnoty se automaticky "
                "nemění.\n\nPřed opravou se vytvoří úplná záloha. Pokračovat?",
                f"{report.fixable} eindeutige Normalisierungen werden durchgeführt.\n"
                "Duplikate, ungültige Verknüpfungen und unklare Werte werden "
                "nicht automatisch geändert.\n\nVorher wird eine vollständige "
                "Sicherung erstellt. Fortfahren?",
            ),
            parent=win,
        ):
            return

        backup_dir = dm.BASE_DIR / "backups"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_path = backup_dir / f"pred_opravou_{timestamp}.zip"
        try:
            em.create_backup_archive(backup_path, dm.DATA_DIR)
            repairs = dv.apply_safe_repairs(dm.DATA_DIR)
            app.stroje = dm.nacti_stroje()
            app.poruchy = dm.nacti_poruchy()
            app.sablony = dm.nacti_sablony()
            if repairs and hasattr(app, "_audit_event"):
                app._audit_event(
                    "data_safe_repairs",
                    "data",
                    details={
                        "backup": str(backup_path),
                        "repairs": [
                            {
                                "file": item.filename,
                                "row": item.row,
                                "field": item.field,
                                "old": item.old_value,
                                "new": item.new_value,
                            }
                            for item in repairs
                        ],
                    },
                )
            app.nakresli_mrizku()
            refresh()
        except Exception as exc:
            messagebox.showerror(
                T("Oprava dat", "Datenkorrektur"),
                T(
                    f"Oprava se nezdařila:\n{exc}",
                    f"Korrektur fehlgeschlagen:\n{exc}",
                ),
                parent=win,
            )
            return
        messagebox.showinfo(
            T("Oprava dat", "Datenkorrektur"),
            T(
                f"Provedeno změn: {len(repairs)}\nZáloha:\n{backup_path}",
                f"Änderungen: {len(repairs)}\nSicherung:\n{backup_path}",
            ),
            parent=win,
        )

    ttk.Button(
        buttons,
        text=T("Znovu zkontrolovat", "Erneut prüfen"),
        command=refresh,
    ).pack(side="left")
    repair_button = ttk.Button(buttons, command=repair)
    repair_button.pack(side="left", padx=(8, 0))
    ttk.Button(
        buttons,
        text=T("Zavřít", "Schließen"),
        command=win.destroy,
    ).pack(side="right")

    win.bind("<F5>", lambda _event: refresh())
    win.bind("<Escape>", lambda _event: win.destroy())
    refresh()
    return win
