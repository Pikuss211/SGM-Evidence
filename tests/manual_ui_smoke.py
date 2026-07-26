"""Read-only Tkinter smoke test. Run once per SGM_LANG value."""
import importlib.util
import json
import os
import tkinter as tk
from pathlib import Path
from tkinter import ttk


ROOT = Path(__file__).resolve().parents[1]
MAIN_PATH = ROOT / "SGM_v1.1-de_clean_fixed.py"


def _load_main():
    spec = importlib.util.spec_from_file_location("sgm_ui_smoke", MAIN_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _walk(widget):
    for child in widget.winfo_children():
        yield child
        yield from _walk(child)


def main():
    module = _load_main()
    language = os.environ.get("SGM_LANG", "de").strip().lower()
    expected = {
        "cz": {
            "title": "SGM Evidence 1.2.0 – přehled strojů",
            "audit": "Historie změn",
            "validation": "Kontrola dat",
            "operator": "Operátor:",
            "confirm": "Potvrdit",
        },
        "de": {
            "title": "SGM Evidence 1.2.0 – Maschinenübersicht",
            "audit": "Änderungsverlauf",
            "validation": "Datenprüfung",
            "operator": "Operator:",
            "confirm": "Bestätigen",
        },
    }.get(language)
    if expected is None:
        raise AssertionError(f"Nepodporovaný testovací jazyk: {language}")

    app = module.StrojeGrid()
    app.withdraw()
    try:
        app.update_idletasks()
        assert app.title() == expected["title"]
        assert app.operator_display.get().startswith(expected["operator"])
        assert len(app.stroje) > 0

        # Ověří skutečný modální dialog operátora bez uložení nastavení.
        def submit_operator():
            top_levels = [
                child
                for child in app.winfo_children()
                if isinstance(child, tk.Toplevel)
            ]
            assert top_levels
            dialog = top_levels[-1]
            combo = next(
                child for child in _walk(dialog) if isinstance(child, ttk.Combobox)
            )
            combo.set("Smoke Tester")
            confirm = next(
                child
                for child in _walk(dialog)
                if isinstance(child, ttk.Button)
                and child.cget("text") == expected["confirm"]
            )
            confirm.invoke()

        app.after(50, submit_operator)
        selected = module.operator_ui.choose_operator(
            app, app.operator, [app.operator], required=True
        )
        assert selected == "Smoke Tester"

        audit_window = module.audit_ui.open_audit_history(app)
        app.update_idletasks()
        assert audit_window.title() == expected["audit"]
        audit_window.destroy()

        validation_window = module.validation_ui.open_data_validation_dialog(app)
        app.update_idletasks()
        assert validation_window.title() == expected["validation"]
        validation_window.destroy()

        result = {
            "language": language,
            "machines": len(app.stroje),
            "faults": len(app.poruchy),
            "title": app.title(),
            "dialogs": ["operator", "audit", "validation"],
        }
        print(json.dumps(result, ensure_ascii=False))
    finally:
        app.destroy()


if __name__ == "__main__":
    main()
