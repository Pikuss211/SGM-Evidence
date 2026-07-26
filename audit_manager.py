#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import csv
import json
import logging
from datetime import datetime
from pathlib import Path

import data_manager as dm


AUDIT_FIELDNAMES = [
    "timestamp",
    "operator",
    "action",
    "entity_type",
    "entity_id",
    "machine_number",
    "details",
]
AUDIT_FILE = dm.DATA_DIR / "audit_log.csv"
LOGGER = logging.getLogger("sgm.audit")


def read_events(audit_file: Path = AUDIT_FILE) -> list[dict]:
    audit_file = Path(audit_file)
    if not audit_file.exists() or audit_file.stat().st_size == 0:
        return []
    with open(audit_file, newline="", encoding="utf-8-sig") as source:
        reader = csv.DictReader(source)
        if not reader.fieldnames:
            return []
        missing = set(AUDIT_FIELDNAMES).difference(reader.fieldnames)
        if missing:
            raise ValueError(
                f"Auditní log nemá povinné sloupce: {', '.join(sorted(missing))}"
            )
        return list(reader)


def record_event(
    action: str,
    entity_type: str,
    *,
    entity_id: str = "",
    machine_number: str = "",
    operator: str = "",
    details: dict | None = None,
    audit_file: Path = AUDIT_FILE,
) -> dict:
    """Přidá auditní událost atomickým přepsáním malého CSV logu."""
    audit_file = Path(audit_file)
    rows = read_events(audit_file)
    event = {
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "operator": (operator or "").strip(),
        "action": (action or "").strip(),
        "entity_type": (entity_type or "").strip(),
        "entity_id": str(entity_id or "").strip(),
        "machine_number": str(machine_number or "").strip(),
        "details": json.dumps(
            details or {}, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        ),
    }
    rows.append(event)
    dm._atomic_csv_write(audit_file, AUDIT_FIELDNAMES, rows)
    return event


def changed_fields(
    before: dict | None,
    after: dict | None,
    fields: list[str] | tuple[str, ...] | None = None,
) -> dict:
    before = before or {}
    after = after or {}
    keys = fields if fields is not None else sorted(set(before) | set(after))
    changes = {}
    for key in keys:
        old = before.get(key, "")
        new = after.get(key, "")
        if old != new:
            changes[str(key)] = {"old": old, "new": new}
    return changes
