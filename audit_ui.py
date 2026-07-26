#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import tkinter as tk
from tkinter import messagebox, ttk

import audit_manager as audit
import data_manager as dm


T = dm.T
ACTION_LABELS = {
    "machine_created": ("Stroj vytvořen", "Maschine erstellt"),
    "machine_updated": ("Stroj upraven", "Maschine geändert"),
    "machine_status_changed": ("Stav stroje změněn", "Maschinenstatus geändert"),
    "machine_archived": ("Stroj archivován", "Maschine archiviert"),
    "machine_restored": ("Stroj obnoven", "Maschine wiederhergestellt"),
    "maintenance_recorded": ("Údržba zaznamenána", "Wartung erfasst"),
    "fault_created": ("Porucha vytvořena", "Störung erstellt"),
    "fault_updated": ("Porucha upravena", "Störung geändert"),
    "fault_closed": ("Porucha uzavřena", "Störung geschlossen"),
    "faults_bulk_closed": ("Poruchy hromadně uzavřeny", "Störungen gesammelt geschlossen"),
    "photo_added": ("Fotografie přidána", "Foto hinzugefügt"),
    "data_safe_repairs": ("Bezpečná oprava dat", "Sichere Datenkorrektur"),
}


def _action_label(action: str) -> str:
    labels = ACTION_LABELS.get(action)
    return T(*labels) if labels else action


def _details_text(raw: str) -> str:
    try:
        details = json.loads(raw or "{}")
    except json.JSONDecodeError:
        return raw or ""
    fields = details.get("fields")
    if isinstance(fields, dict):
        parts = []
        for field, values in fields.items():
            if isinstance(values, dict):
                parts.append(
                    f"{field}: {values.get('old', '')} → {values.get('new', '')}"
                )
        if parts:
            return "; ".join(parts)
    return "; ".join(f"{key}={value}" for key, value in details.items())


def open_audit_history(parent, machine_number: str | None = None):
    try:
        events = audit.read_events()
    except Exception as exc:
        messagebox.showerror(
            T("Historie změn", "Änderungsverlauf"),
            T(
                f"Auditní log nelze načíst:\n{exc}",
                f"Auditprotokoll kann nicht gelesen werden:\n{exc}",
            ),
            parent=parent,
        )
        return None
    if machine_number is not None:
        machine_number = str(machine_number)
        events = [
            event
            for event in events
            if str(event.get("machine_number", "")) == machine_number
        ]
    events.reverse()

    win = tk.Toplevel(parent)
    title = T("Historie změn", "Änderungsverlauf")
    if machine_number is not None:
        title += f" – {T('stroj', 'Maschine')} {machine_number}"
    win.title(title)
    win.geometry("1100x560")
    win.minsize(820, 400)
    win.transient(parent)

    frame = tk.Frame(win, padx=12, pady=12)
    frame.pack(fill="both", expand=True)
    frame.grid_rowconfigure(1, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    tk.Label(
        frame,
        text=T(
            f"Počet zaznamenaných změn: {len(events)}",
            f"Anzahl protokollierter Änderungen: {len(events)}",
        ),
        font=("Segoe UI", 11, "bold"),
    ).grid(row=0, column=0, sticky="w", pady=(0, 10))

    columns = ("timestamp", "operator", "action", "entity", "details")
    tree = ttk.Treeview(frame, columns=columns, show="headings")
    tree.heading("timestamp", text=T("Čas", "Zeit"))
    tree.heading("operator", text=T("Operátor", "Operator"))
    tree.heading("action", text=T("Akce", "Aktion"))
    tree.heading("entity", text=T("Záznam", "Eintrag"))
    tree.heading("details", text=T("Podrobnosti", "Details"))
    tree.column("timestamp", width=190, stretch=False)
    tree.column("operator", width=130, stretch=False)
    tree.column("action", width=190, stretch=False)
    tree.column("entity", width=120, stretch=False)
    tree.column("details", width=440)

    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.grid(row=1, column=0, sticky="nsew")
    scrollbar.grid(row=1, column=1, sticky="ns")

    for index, event in enumerate(events):
        entity = event.get("entity_type", "")
        if event.get("entity_id"):
            entity += f" {event['entity_id']}"
        tree.insert(
            "",
            "end",
            iid=str(index),
            values=(
                event.get("timestamp", ""),
                event.get("operator", ""),
                _action_label(event.get("action", "")),
                entity,
                _details_text(event.get("details", "")),
            ),
        )

    ttk.Button(
        frame, text=T("Zavřít", "Schließen"), command=win.destroy
    ).grid(row=2, column=0, sticky="e", pady=(12, 0))
    win.bind("<Escape>", lambda _event: win.destroy())
    return win
