#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import os
import shutil
from pathlib import Path

import data_manager as dm


SETTINGS_FILE = dm.DATA_DIR / "app_settings.json"
MAX_OPERATORS = 20


def _clean_operator(value) -> str:
    return " ".join(str(value or "").strip().split())


def _normalize_settings(raw) -> dict:
    if not isinstance(raw, dict):
        raise ValueError("Nastavení operátora musí být objekt JSON.")
    raw_operators = raw.get("operators", [])
    if not isinstance(raw_operators, (list, tuple)):
        raw_operators = []
    operators = []
    seen = set()
    for value in raw_operators:
        name = _clean_operator(value)
        folded = name.casefold()
        if name and folded not in seen:
            operators.append(name)
            seen.add(folded)
    last_operator = _clean_operator(raw.get("last_operator", ""))
    if last_operator:
        operators = [
            last_operator,
            *[
                item
                for item in operators
                if item.casefold() != last_operator.casefold()
            ],
        ]
    return {
        "last_operator": last_operator,
        "operators": operators[:MAX_OPERATORS],
    }


def _read_settings(path: Path) -> dict:
    with open(path, encoding="utf-8-sig") as source:
        return _normalize_settings(json.load(source))


def load_settings(settings_file: Path = SETTINGS_FILE) -> dict:
    """Načte nastavení; při poškození zkusí poslední záložní kopii."""
    settings_file = Path(settings_file)
    for candidate in (
        settings_file,
        settings_file.with_suffix(settings_file.suffix + ".bak"),
    ):
        if not candidate.exists():
            continue
        try:
            return _read_settings(candidate)
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError):
            continue
    return {"last_operator": "", "operators": []}


def initial_operator(
    settings_file: Path = SETTINGS_FILE, environ: dict | None = None
) -> str:
    settings = load_settings(settings_file)
    if settings["last_operator"]:
        return settings["last_operator"]
    environment = os.environ if environ is None else environ
    return _clean_operator(
        environment.get("USERNAME") or environment.get("USER") or ""
    )


def remember_operator(
    operator: str, settings_file: Path = SETTINGS_FILE
) -> dict:
    """Atomicky uloží vybraného operátora a seznam posledních uživatelů."""
    operator = _clean_operator(operator)
    if not operator:
        raise ValueError("Jméno operátora nesmí být prázdné.")

    settings_file = Path(settings_file)
    settings = load_settings(settings_file)
    operators = [
        operator,
        *[
            item
            for item in settings["operators"]
            if item.casefold() != operator.casefold()
        ],
    ][:MAX_OPERATORS]
    updated = {"last_operator": operator, "operators": operators}

    settings_file.parent.mkdir(parents=True, exist_ok=True)
    temp_path = settings_file.with_name(f".{settings_file.name}.tmp")
    try:
        with open(temp_path, "w", encoding="utf-8", newline="\n") as target:
            json.dump(updated, target, ensure_ascii=False, indent=2)
            target.write("\n")
            target.flush()
            os.fsync(target.fileno())

        # Zálohu aktualizujeme jen tehdy, když je současný hlavní soubor čitelný.
        if settings_file.exists():
            try:
                _read_settings(settings_file)
            except (OSError, UnicodeError, json.JSONDecodeError, ValueError):
                pass
            else:
                shutil.copy2(
                    settings_file,
                    settings_file.with_suffix(settings_file.suffix + ".bak"),
                )
        os.replace(temp_path, settings_file)
    finally:
        temp_path.unlink(missing_ok=True)
    return updated
