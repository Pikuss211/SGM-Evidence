#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import csv
import os
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import data_manager as dm


MACHINE_REQUIRED = {
    "cislo",
    "vyrobce",
    "typ",
    "rok",
    "spm",
    "seriove",
    "stav",
    "wartung_last",
    "wartung_interval",
}
FAULT_REQUIRED = {
    "id",
    "cislo",
    "stav",
    "cas",
    "cas_uzavreni",
    "kategorie",
    "alarm",
    "popis",
    "reseni",
}

MACHINE_STATUS_ALIASES = {
    "b": "bezi",
    "běží": "bezi",
    "bezi": "bezi",
    "l": "bezi",
    "läuft": "bezi",
    "laeuft": "bezi",
    "lauf": "bezi",
    "running": "bezi",
    "ok": "bezi",
    "p": "porucha",
    "porucha": "porucha",
    "s": "porucha",
    "störung": "porucha",
    "stoerung": "porucha",
    "fault": "porucha",
    "error": "porucha",
}
FAULT_STATUS_ALIASES = {
    "o": "otevrena",
    "offen": "otevrena",
    "otevrena": "otevrena",
    "otevřená": "otevrena",
    "g": "uzavrena",
    "geschlossen": "uzavrena",
    "uzavrena": "uzavrena",
    "uzavřená": "uzavrena",
}
CATEGORY_ALIASES = {
    "e": "elektricka",
    "elektricka": "elektricka",
    "elektrická": "elektricka",
    "electrical": "elektricka",
    "elektrisch": "elektricka",
    "m": "mechanicka",
    "mechanicka": "mechanicka",
    "mechanická": "mechanicka",
    "mechanical": "mechanicka",
    "mechanisch": "mechanicka",
    "j": "jina",
    "jina": "jina",
    "jiná": "jina",
    "other": "jina",
    "sonstige": "jina",
    "andere": "jina",
}
ARCHIVE_ALIASES = {
    "0": "0",
    "ne": "0",
    "false": "0",
    "no": "0",
    "nein": "0",
    "1": "1",
    "ano": "1",
    "true": "1",
    "yes": "1",
    "ja": "1",
    "archivovan": "1",
    "archiviert": "1",
}


@dataclass(frozen=True)
class DataIssue:
    severity: str
    code: str
    filename: str
    row: int | None
    message_cz: str
    message_de: str
    fixable: bool = False


@dataclass
class ValidationReport:
    issues: list[DataIssue] = field(default_factory=list)
    machine_count: int = 0
    fault_count: int = 0

    @property
    def errors(self) -> int:
        return sum(issue.severity == "error" for issue in self.issues)

    @property
    def warnings(self) -> int:
        return sum(issue.severity == "warning" for issue in self.issues)

    @property
    def fixable(self) -> int:
        return sum(issue.fixable for issue in self.issues)


@dataclass(frozen=True)
class DataRepair:
    filename: str
    row: int
    field: str
    old_value: str
    new_value: str


def _delimiter(sample: str) -> str:
    first_line = sample.splitlines()[0] if sample.splitlines() else ""
    return ";" if ";" in first_line and "," not in first_line else ","


def _read_csv(path: Path):
    with open(path, newline="", encoding="utf-8-sig") as source:
        sample = source.read(4096)
        source.seek(0)
        delimiter = _delimiter(sample)
        reader = csv.DictReader(source, delimiter=delimiter)
        return list(reader.fieldnames or []), list(reader)


def _parse_fault_datetime(value: str):
    value = (value or "").strip()
    if not value:
        return None, ""
    for fmt in (
        "%Y-%m-%d %H:%M",
        "%d.%m.%Y %H:%M",
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%Y/%m/%d %H:%M",
    ):
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed, parsed.strftime(dm.FMT)
        except ValueError:
            pass
    return None, value


def _parse_maintenance_date(value: str):
    value = (value or "").strip()
    if not value:
        return None, ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d"):
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed, parsed.strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None, value


def _issue(
    report: ValidationReport,
    severity: str,
    code: str,
    filename: str,
    row: int | None,
    cz: str,
    de: str,
    fixable: bool = False,
):
    report.issues.append(
        DataIssue(severity, code, filename, row, cz, de, fixable)
    )


def _read_for_validation(
    path: Path, filename: str, report: ValidationReport
) -> tuple[list[str], list[dict]]:
    if not path.exists():
        _issue(
            report,
            "error",
            "file_missing",
            filename,
            None,
            "Soubor chybí.",
            "Datei fehlt.",
        )
        return [], []
    try:
        return _read_csv(path)
    except (OSError, UnicodeError, csv.Error) as exc:
        _issue(
            report,
            "error",
            "file_unreadable",
            filename,
            None,
            f"Soubor nelze načíst: {exc}",
            f"Datei kann nicht gelesen werden: {exc}",
        )
        return [], []


def validate_data(data_dir: Path = dm.DATA_DIR) -> ValidationReport:
    data_dir = Path(data_dir)
    report = ValidationReport()
    machine_path = data_dir / "stroje.csv"
    fault_path = data_dir / "poruchy.csv"
    template_path = data_dir / "sablony_alarmu.csv"

    machine_fields, machines = _read_for_validation(
        machine_path, "stroje.csv", report
    )
    report.machine_count = len(machines)
    missing = MACHINE_REQUIRED.difference(machine_fields)
    if missing:
        _issue(
            report,
            "error",
            "machine_columns_missing",
            "stroje.csv",
            1,
            f"Chybí sloupce: {', '.join(sorted(missing))}",
            f"Spalten fehlen: {', '.join(sorted(missing))}",
        )

    machine_numbers = []
    for index, machine in enumerate(machines, start=2):
        if None in machine:
            _issue(
                report,
                "error",
                "machine_extra_values",
                "stroje.csv",
                index,
                "Řádek má více hodnot než hlavička.",
                "Zeile enthält mehr Werte als die Kopfzeile.",
            )
        number = (machine.get("cislo") or "").strip()
        if not number:
            _issue(
                report,
                "error",
                "machine_number_missing",
                "stroje.csv",
                index,
                "Chybí číslo stroje.",
                "Maschinennummer fehlt.",
            )
        else:
            machine_numbers.append(number)

        status = (machine.get("stav") or "").strip().lower()
        normalized_status = MACHINE_STATUS_ALIASES.get(status)
        if normalized_status is None:
            _issue(
                report,
                "error",
                "machine_status_invalid",
                "stroje.csv",
                index,
                f"Neplatný stav stroje: {machine.get('stav', '')!r}",
                f"Ungültiger Maschinenstatus: {machine.get('stav', '')!r}",
            )
        elif status != normalized_status:
            _issue(
                report,
                "warning",
                "machine_status_noncanonical",
                "stroje.csv",
                index,
                f"Stav {machine.get('stav')!r} bude sjednocen na {normalized_status!r}.",
                f"Status {machine.get('stav')!r} wird auf {normalized_status!r} vereinheitlicht.",
                True,
            )

        interval = (machine.get("wartung_interval") or "").strip()
        if not interval:
            _issue(
                report,
                "warning",
                "maintenance_interval_missing",
                "stroje.csv",
                index,
                "Chybí interval údržby; lze bezpečně doplnit 180 dní.",
                "Wartungsintervall fehlt; 180 Tage können sicher ergänzt werden.",
                True,
            )
        else:
            try:
                if int(interval) <= 0:
                    raise ValueError
            except ValueError:
                _issue(
                    report,
                    "warning",
                    "maintenance_interval_invalid",
                    "stroje.csv",
                    index,
                    f"Neplatný interval údržby: {interval!r}",
                    f"Ungültiges Wartungsintervall: {interval!r}",
                )

        maintenance = (machine.get("wartung_last") or "").strip()
        parsed, canonical = _parse_maintenance_date(maintenance)
        if maintenance and parsed is None:
            _issue(
                report,
                "warning",
                "maintenance_date_invalid",
                "stroje.csv",
                index,
                f"Neplatné datum poslední údržby: {maintenance!r}",
                f"Ungültiges Datum der letzten Wartung: {maintenance!r}",
            )
        elif maintenance and maintenance != canonical:
            _issue(
                report,
                "warning",
                "maintenance_date_noncanonical",
                "stroje.csv",
                index,
                f"Datum údržby bude sjednoceno na {canonical}.",
                f"Wartungsdatum wird auf {canonical} vereinheitlicht.",
                True,
            )

        if "archivovan" in machine_fields:
            archived = (machine.get("archivovan") or "").strip().lower()
            normalized_archived = ARCHIVE_ALIASES.get(archived)
            if not archived:
                _issue(
                    report,
                    "warning",
                    "machine_archive_flag_missing",
                    "stroje.csv",
                    index,
                    "Chybí příznak archivace; lze bezpečně doplnit 0.",
                    "Archivkennzeichen fehlt; 0 kann sicher ergänzt werden.",
                    True,
                )
            elif normalized_archived is None:
                _issue(
                    report,
                    "warning",
                    "machine_archive_flag_invalid",
                    "stroje.csv",
                    index,
                    f"Neplatný příznak archivace: {machine.get('archivovan')!r}",
                    f"Ungültiges Archivkennzeichen: {machine.get('archivovan')!r}",
                )
            elif archived != normalized_archived:
                _issue(
                    report,
                    "warning",
                    "machine_archive_flag_noncanonical",
                    "stroje.csv",
                    index,
                    f"Příznak archivace bude sjednocen na {normalized_archived}.",
                    f"Archivkennzeichen wird auf {normalized_archived} vereinheitlicht.",
                    True,
                )

    duplicate_machines = {
        value for value, count in Counter(machine_numbers).items() if count > 1
    }
    for index, machine in enumerate(machines, start=2):
        number = (machine.get("cislo") or "").strip()
        if number in duplicate_machines:
            _issue(
                report,
                "error",
                "machine_number_duplicate",
                "stroje.csv",
                index,
                f"Duplicitní číslo stroje: {number}",
                f"Doppelte Maschinennummer: {number}",
            )

    fault_fields, faults = _read_for_validation(fault_path, "poruchy.csv", report)
    report.fault_count = len(faults)
    missing = FAULT_REQUIRED.difference(fault_fields)
    if missing:
        _issue(
            report,
            "error",
            "fault_columns_missing",
            "poruchy.csv",
            1,
            f"Chybí sloupce: {', '.join(sorted(missing))}",
            f"Spalten fehlen: {', '.join(sorted(missing))}",
        )

    known_machines = set(machine_numbers)
    fault_ids = []
    for index, fault in enumerate(faults, start=2):
        if None in fault:
            _issue(
                report,
                "error",
                "fault_extra_values",
                "poruchy.csv",
                index,
                "Řádek má více hodnot než hlavička.",
                "Zeile enthält mehr Werte als die Kopfzeile.",
            )
        fault_id = (fault.get("id") or "").strip()
        if not fault_id:
            _issue(
                report,
                "error",
                "fault_id_missing",
                "poruchy.csv",
                index,
                "Chybí ID poruchy.",
                "Störungs-ID fehlt.",
            )
        else:
            fault_ids.append(fault_id)
            if not fault_id.isdigit():
                _issue(
                    report,
                    "warning",
                    "fault_id_nonnumeric",
                    "poruchy.csv",
                    index,
                    f"ID poruchy není číselné: {fault_id!r}",
                    f"Störungs-ID ist nicht numerisch: {fault_id!r}",
                )

        machine_number = (fault.get("cislo") or "").strip()
        if not machine_number:
            _issue(
                report,
                "error",
                "fault_machine_missing",
                "poruchy.csv",
                index,
                "Chybí číslo stroje.",
                "Maschinennummer fehlt.",
            )
        elif machine_number not in known_machines:
            _issue(
                report,
                "error",
                "fault_machine_unknown",
                "poruchy.csv",
                index,
                f"Porucha odkazuje na neexistující stroj {machine_number}.",
                f"Störung verweist auf unbekannte Maschine {machine_number}.",
            )

        status = (fault.get("stav") or "").strip().lower()
        normalized_status = FAULT_STATUS_ALIASES.get(status)
        if normalized_status is None:
            _issue(
                report,
                "error",
                "fault_status_invalid",
                "poruchy.csv",
                index,
                f"Neplatný stav poruchy: {fault.get('stav', '')!r}",
                f"Ungültiger Störungsstatus: {fault.get('stav', '')!r}",
            )
        elif status != normalized_status:
            _issue(
                report,
                "warning",
                "fault_status_noncanonical",
                "poruchy.csv",
                index,
                f"Stav bude sjednocen na {normalized_status!r}.",
                f"Status wird auf {normalized_status!r} vereinheitlicht.",
                True,
            )

        category = (fault.get("kategorie") or "").strip().lower()
        normalized_category = CATEGORY_ALIASES.get(category)
        if normalized_category is None:
            _issue(
                report,
                "warning",
                "fault_category_invalid",
                "poruchy.csv",
                index,
                f"Neznámá kategorie: {fault.get('kategorie', '')!r}",
                f"Unbekannte Kategorie: {fault.get('kategorie', '')!r}",
            )
        elif category != normalized_category:
            _issue(
                report,
                "warning",
                "fault_category_noncanonical",
                "poruchy.csv",
                index,
                f"Kategorie bude sjednocena na {normalized_category!r}.",
                f"Kategorie wird auf {normalized_category!r} vereinheitlicht.",
                True,
            )

        opened_text = (fault.get("cas") or "").strip()
        opened, opened_canonical = _parse_fault_datetime(opened_text)
        if not opened_text:
            _issue(
                report,
                "error",
                "fault_opened_missing",
                "poruchy.csv",
                index,
                "Chybí čas otevření poruchy.",
                "Öffnungszeit der Störung fehlt.",
            )
        elif opened is None:
            _issue(
                report,
                "error",
                "fault_opened_invalid",
                "poruchy.csv",
                index,
                f"Neplatný čas otevření: {opened_text!r}",
                f"Ungültige Öffnungszeit: {opened_text!r}",
            )
        elif opened_text != opened_canonical:
            _issue(
                report,
                "warning",
                "fault_opened_noncanonical",
                "poruchy.csv",
                index,
                f"Čas otevření bude sjednocen na {opened_canonical}.",
                f"Öffnungszeit wird auf {opened_canonical} vereinheitlicht.",
                True,
            )

        closed_text = (fault.get("cas_uzavreni") or "").strip()
        closed, closed_canonical = _parse_fault_datetime(closed_text)
        if closed_text and closed is None:
            _issue(
                report,
                "error",
                "fault_closed_invalid",
                "poruchy.csv",
                index,
                f"Neplatný čas uzavření: {closed_text!r}",
                f"Ungültige Schließzeit: {closed_text!r}",
            )
        elif closed_text and closed_text != closed_canonical:
            _issue(
                report,
                "warning",
                "fault_closed_noncanonical",
                "poruchy.csv",
                index,
                f"Čas uzavření bude sjednocen na {closed_canonical}.",
                f"Schließzeit wird auf {closed_canonical} vereinheitlicht.",
                True,
            )
        if normalized_status == "uzavrena" and not closed_text:
            _issue(
                report,
                "warning",
                "fault_closed_time_missing",
                "poruchy.csv",
                index,
                "Uzavřená porucha nemá čas uzavření.",
                "Geschlossene Störung hat keine Schließzeit.",
            )
        if opened is not None and closed is not None and closed < opened:
            _issue(
                report,
                "error",
                "fault_closed_before_opened",
                "poruchy.csv",
                index,
                "Čas uzavření je dříve než čas otevření.",
                "Schließzeit liegt vor der Öffnungszeit.",
            )

    duplicate_faults = {
        value for value, count in Counter(fault_ids).items() if count > 1
    }
    for index, fault in enumerate(faults, start=2):
        fault_id = (fault.get("id") or "").strip()
        if fault_id in duplicate_faults:
            _issue(
                report,
                "error",
                "fault_id_duplicate",
                "poruchy.csv",
                index,
                f"Duplicitní ID poruchy: {fault_id}",
                f"Doppelte Störungs-ID: {fault_id}",
            )

    if template_path.exists():
        template_fields, _ = _read_for_validation(
            template_path, "sablony_alarmu.csv", report
        )
        missing = {"alarm", "reseni"}.difference(template_fields)
        if missing:
            _issue(
                report,
                "warning",
                "template_columns_missing",
                "sablony_alarmu.csv",
                1,
                f"Chybí sloupce: {', '.join(sorted(missing))}",
                f"Spalten fehlen: {', '.join(sorted(missing))}",
            )

    return report


def apply_safe_repairs(data_dir: Path = dm.DATA_DIR) -> list[DataRepair]:
    """Provede pouze jednoznačné normalizace; neřeší duplicity ani vazby."""
    data_dir = Path(data_dir)
    repairs: list[DataRepair] = []

    machine_path = data_dir / "stroje.csv"
    machine_fields, machines = _read_csv(machine_path)
    archive_field_present = "archivovan" in machine_fields
    if not archive_field_present:
        machine_fields.append("archivovan")
    for index, machine in enumerate(machines, start=2):
        status = (machine.get("stav") or "").strip().lower()
        normalized = MACHINE_STATUS_ALIASES.get(status)
        if normalized is not None and status != normalized:
            repairs.append(
                DataRepair(
                    "stroje.csv",
                    index,
                    "stav",
                    machine.get("stav") or "",
                    normalized,
                )
            )
            machine["stav"] = normalized
        interval = (machine.get("wartung_interval") or "").strip()
        if not interval:
            repairs.append(
                DataRepair(
                    "stroje.csv", index, "wartung_interval", interval, "180"
                )
            )
            machine["wartung_interval"] = "180"
        maintenance = (machine.get("wartung_last") or "").strip()
        parsed, canonical = _parse_maintenance_date(maintenance)
        if parsed is not None and maintenance != canonical:
            repairs.append(
                DataRepair(
                    "stroje.csv",
                    index,
                    "wartung_last",
                    maintenance,
                    canonical,
                )
            )
            machine["wartung_last"] = canonical
        archived = (machine.get("archivovan") or "").strip().lower()
        normalized = ARCHIVE_ALIASES.get(archived)
        if not archived:
            normalized = "0"
        if not archive_field_present:
            machine["archivovan"] = normalized or "0"
        elif normalized is not None and archived != normalized:
            repairs.append(
                DataRepair(
                    "stroje.csv",
                    index,
                    "archivovan",
                    machine.get("archivovan") or "",
                    normalized,
                )
            )
            machine["archivovan"] = normalized

    fault_path = data_dir / "poruchy.csv"
    fault_fields, faults = _read_csv(fault_path)
    for index, fault in enumerate(faults, start=2):
        status = (fault.get("stav") or "").strip().lower()
        normalized = FAULT_STATUS_ALIASES.get(status)
        if normalized is not None and status != normalized:
            repairs.append(
                DataRepair(
                    "poruchy.csv",
                    index,
                    "stav",
                    fault.get("stav") or "",
                    normalized,
                )
            )
            fault["stav"] = normalized
        category = (fault.get("kategorie") or "").strip().lower()
        normalized = CATEGORY_ALIASES.get(category)
        if normalized is not None and category != normalized:
            repairs.append(
                DataRepair(
                    "poruchy.csv",
                    index,
                    "kategorie",
                    fault.get("kategorie") or "",
                    normalized,
                )
            )
            fault["kategorie"] = normalized
        for field_name in ("cas", "cas_uzavreni"):
            value = (fault.get(field_name) or "").strip()
            parsed, canonical = _parse_fault_datetime(value)
            if parsed is not None and value != canonical:
                repairs.append(
                    DataRepair(
                        "poruchy.csv",
                        index,
                        field_name,
                        value,
                        canonical,
                    )
                )
                fault[field_name] = canonical

    changed_machine = any(
        repair.filename == "stroje.csv" for repair in repairs
    )
    changed_fault = any(repair.filename == "poruchy.csv" for repair in repairs)
    originals = {
        path: path.read_bytes()
        for path, changed in (
            (machine_path, changed_machine),
            (fault_path, changed_fault),
        )
        if changed
    }
    try:
        if changed_machine:
            dm._atomic_csv_write(machine_path, machine_fields, machines)
        if changed_fault:
            dm._atomic_csv_write(fault_path, fault_fields, faults)
    except Exception:
        for path, content in originals.items():
            rollback_path = path.with_name(f".{path.name}.repair_rollback.tmp")
            try:
                with open(rollback_path, "wb") as target:
                    target.write(content)
                    target.flush()
                    os.fsync(target.fileno())
                os.replace(rollback_path, path)
            finally:
                rollback_path.unlink(missing_ok=True)
        raise
    return repairs
