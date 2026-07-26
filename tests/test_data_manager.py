import csv
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch

import data_manager as dm


@contextmanager
def test_directory(name: str):
    path = Path(__file__).parent / "_runtime" / name
    path.mkdir(parents=True, exist_ok=True)
    yield str(path)


class AtomicCsvWriteTests(unittest.TestCase):
    def test_machine_save_is_atomic_and_keeps_previous_copy(self):
        with test_directory("data_machines") as temp:
            target = Path(temp) / "stroje.csv"
            with patch.object(dm, "SOUBOR_STROJE", target):
                dm.uloz_stroje(
                    {
                        "1": {
                            "vyrobce": "Arburg",
                            "typ": "A",
                            "rok": "2020",
                            "spm": "",
                            "seriove": "",
                            "stav": "bezi",
                            "wartung_last": "",
                            "wartung_interval": "180",
                        }
                    }
                )
                first_version = target.read_bytes()

                machines = dm.nacti_stroje()
                machines["1"]["typ"] = "B"
                dm.uloz_stroje(machines)

            self.assertEqual(
                target.with_suffix(".csv.bak").read_bytes(), first_version
            )
            with open(target, newline="", encoding="utf-8") as source:
                rows = list(csv.DictReader(source))
            self.assertEqual(rows[0]["typ"], "B")
            self.assertEqual(rows[0]["archivovan"], "0")
            self.assertFalse(list(Path(temp).glob("*.tmp")))

    def test_archived_machine_flag_accepts_legacy_and_canonical_values(self):
        self.assertTrue(dm.is_archived_machine({"archivovan": "1"}))
        self.assertTrue(dm.is_archived_machine({"archivovan": "ja"}))
        self.assertFalse(dm.is_archived_machine({"archivovan": "0"}))
        self.assertFalse(dm.is_archived_machine({}))

    def test_fault_save_preserves_additional_columns(self):
        with test_directory("data_faults") as temp:
            target = Path(temp) / "poruchy.csv"
            with patch.object(dm, "SOUBOR_PORUCHY", target):
                dm.uloz_poruchy(
                    [
                        {
                            "id": "1",
                            "cislo": "7",
                            "stav": "otevrena",
                            "alarm": "A1",
                            "vlastni_sloupec": "hodnota",
                        }
                    ]
                )
            with open(target, newline="", encoding="utf-8") as source:
                row = next(csv.DictReader(source))
            self.assertEqual(row["vlastni_sloupec"], "hodnota")


if __name__ == "__main__":
    unittest.main()
