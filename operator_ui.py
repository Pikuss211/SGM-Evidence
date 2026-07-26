#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import messagebox, ttk

import data_manager as dm


T = dm.T


def choose_operator(
    parent,
    current: str = "",
    operators: list[str] | tuple[str, ...] | None = None,
    *,
    required: bool = False,
) -> str | None:
    """Zobrazí modální výběr operátora a vrátí neprázdné jméno."""
    win = tk.Toplevel(parent)
    win.title(T("Výběr operátora", "Operator auswählen"))
    win.transient(parent)
    win.grab_set()
    win.resizable(False, False)

    frame = ttk.Frame(win, padding=16)
    frame.pack(fill="both", expand=True)
    ttk.Label(
        frame,
        text=T(
            "Zadejte nebo vyberte operátora pro tuto směnu.",
            "Operator für diese Schicht eingeben oder auswählen.",
        ),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

    ttk.Label(frame, text=T("Operátor:", "Operator:")).grid(
        row=1, column=0, sticky="w", padx=(0, 10)
    )
    value = tk.StringVar(value=current)
    combo = ttk.Combobox(
        frame,
        textvariable=value,
        values=list(operators or []),
        width=34,
    )
    combo.grid(row=1, column=1, sticky="ew")
    frame.grid_columnconfigure(1, weight=1)

    result = {"value": None}

    def confirm(_event=None):
        selected = " ".join(value.get().strip().split())
        if not selected:
            messagebox.showwarning(
                T("Operátor", "Operator"),
                T(
                    "Jméno operátora nesmí být prázdné.",
                    "Der Operatorname darf nicht leer sein.",
                ),
                parent=win,
            )
            return
        result["value"] = selected
        win.destroy()

    def cancel(_event=None):
        win.destroy()

    buttons = ttk.Frame(frame)
    buttons.grid(row=2, column=0, columnspan=2, sticky="e", pady=(14, 0))
    ttk.Button(
        buttons,
        text=(
            T("Ukončit program", "Programm beenden")
            if required
            else T("Zrušit", "Abbrechen")
        ),
        command=cancel,
    ).pack(side="right")
    ttk.Button(
        buttons, text=T("Potvrdit", "Bestätigen"), command=confirm
    ).pack(side="right", padx=(0, 8))

    win.bind("<Return>", confirm)
    win.bind("<Escape>", cancel)
    win.protocol("WM_DELETE_WINDOW", cancel)
    win.after(0, lambda: (combo.focus_set(), combo.selection_range(0, "end")))
    parent.wait_window(win)
    return result["value"]
