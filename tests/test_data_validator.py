import csv
import unittest
from pathlib import Path
from unittest.mock import patch

import data_validator as dv


MACHINE_FIELDS = [
    "cislo",
    "vyrobce",
    "typ",
    "rok",
    "spm",
    "seriove",
    "stav",
    "wartung_last",
    "wartung_interval",
]
FAULT_FIELDS = [
    "alarm",
    "cas",
    "cas_uzavreni",
    "cislo",
    "id",
    "kategorie",
    "operator_uzavrel",
    "popis",
    "reseni",
    "stav",
    "typ",
]


def runtime_dir(name: str) -> Path:
    path = Path(__file__).parent / "_runtime" / name
    path.mkdir(parents=True, exist_ok=True)
    return path


def write_csv(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    with open(path, "w", newline="", encoding="utf-8") as target:
        writer = csv.DictWriter(target, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def valid_machine(**changes):
    row = {
        "cislo": "1",
        "vyrobce": "Arburg",
        "typ": "A",
        "rok": "2020",
        "spm": "",
        "seriove": "",
        "stav": "bezi",
        "wartung_last": "2026-01-01",
        "wartung_interval": "180",
    }
    row.update(changes)
    return row


def valid_fault(**changes):
    row = {
        "alarm": "A1",
        "cas": "2026-01-02 10:00",
        "cas_uzavreni": "",
        "cislo": "1",
        "id": "1",
        "kategorie": "elektricka",
        "operator_uzavrel": "",
        "popis": "Test",
        "reseni": "",
        "stav": "otevrena",
        "typ": "A",
    }
    row.update(changes)
    return row


class DataValidatorTests(unittest.TestCase):
    def test_clean_data_has_no_issues(self):
        data = runtime_dir("validator_clean")
        write_csv(data / "stroje.csv", MACHINE_FIELDS, [valid_machine()])
        write_csv(data / "poruchy.csv", FAULT_FIELDS, [valid_fault()])

        report = dv.validate_data(data)

        self.assertEqual(report.errors, 0)
        self.assertEqual(report.warnings, 0)
        self.assertEqual(report.fixable, 0)

    def test_validation_finds_duplicates_links_and_invalid_values(self):
        data = runtime_dir("validator_issues")
        write_csv(
            data / "stroje.csv",
            MACHINE_FIELDS,
            [
                valid_machine(
                    stav="běží",
                    wartung_last="31.02.2026",
                    wartung_interval="",
                ),
                valid_machine(typ="Duplicate"),
            ],
        )
        write_csv(
            data / "poruchy.csv",
            FAULT_FIELDS,
            [
                valid_fault(
                    cas="02.01.2026 10:00",
                    cislo="999",
                    kategorie="Elektrisch",
                    stav="offen",
                ),
                valid_fault(
                    cas="2026-01-03 10:00",
                    cas_uzavreni="2026-01-03 09:00",
                    stav="geschlossen",
                ),
            ],
        )

        report = dv.validate_data(data)
        codes = {issue.code for issue in report.issues}

        self.assertIn("machine_number_duplicate", codes)
        self.assertIn("maintenance_date_invalid", codes)
        self.assertIn("fault_id_duplicate", codes)
        self.assertIn("fault_machine_unknown", codes)
        self.assertIn("fault_closed_before_opened", codes)
        self.assertGreaterEqual(report.fixable, 5)

    def test_safe_repairs_only_normalize_unambiguous_values(self):
        data = runtime_dir("validator_repairs")
        write_csv(
            data / "stroje.csv",
            MACHINE_FIELDS,
            [
                valid_machine(
                    stav="běží",
                    wartung_last="01.02.2026",
                    wartung_interval="",
                )
            ],
        )
        write_csv(
            data / "poruchy.csv",
            FAULT_FIELDS,
            [
                valid_fault(
                    cas="02.02.2026 10:30",
                    kategorie="Elektrisch",
                    stav="offen",
                )
            ],
        )

        repairs = dv.apply_safe_repairs(data)

        self.assertEqual(len(repairs), 6)
        with open(
            data / "stroje.csv", newline="", encoding="utf-8"
        ) as source:
            machine = next(csv.DictReader(source))
        with open(
            data / "poruchy.csv", newline="", encoding="utf-8"
        ) as source:
            fault = next(csv.DictReader(source))
        self.assertEqual(machine["stav"], "bezi")
        self.assertEqual(machine["wartung_interval"], "180")
        self.assertEqual(machine["wartung_last"], "2026-02-01")
        self.assertEqual(fault["stav"], "otevrena")
        self.assertEqual(fault["kategorie"], "elektricka")
        self.assertEqual(fault["cas"], "2026-02-02 10:30")
        self.assertEqual(dv.validate_data(data).fixable, 0)

    def test_safe_repairs_roll_back_when_second_file_write_fails(self):
        data = runtime_dir("validator_rollback")
        write_csv(
            data / "stroje.csv",
            MACHINE_FIELDS,
            [valid_machine(stav="běží")],
        )
        write_csv(
            data / "poruchy.csv",
            FAULT_FIELDS,
            [valid_fault(stav="offen")],
        )
        original_machine = (data / "stroje.csv").read_bytes()
        original_fault = (data / "poruchy.csv").read_bytes()
        atomic_write = dv.dm._atomic_csv_write

        def fail_fault_write(path, fieldnames, rows):
            if Path(path).name == "poruchy.csv":
                raise OSError("simulovaná chyba zápisu")
            return atomic_write(path, fieldnames, rows)

        with patch.object(
            dv.dm, "_atomic_csv_write", side_effect=fail_fault_write
        ):
            with self.assertRaisesRegex(OSError, "simulovaná"):
                dv.apply_safe_repairs(data)

        self.assertEqual((data / "stroje.csv").read_bytes(), original_machine)
        self.assertEqual((data / "poruchy.csv").read_bytes(), original_fault)


if __name__ == "__main__":
    unittest.main()
