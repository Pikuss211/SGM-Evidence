import unittest
import zipfile
from contextlib import contextmanager
from pathlib import Path

import export_manager as em


MACHINES = (
    "cislo,vyrobce,typ,rok,spm,seriove,stav,wartung_last,wartung_interval\n"
    "1,Arburg,A,2020,,,bezi,,180\n"
)
FAULTS = (
    "alarm,cas,cas_uzavreni,cislo,id,kategorie,operator_uzavrel,"
    "popis,reseni,stav,typ\n"
    "A1,2026-01-01 10:00,,1,1,elektricka,,Test,,otevrena,\n"
)


@contextmanager
def test_directory(name: str):
    path = Path(__file__).parent / "_runtime" / name
    path.mkdir(parents=True, exist_ok=True)
    yield str(path)


def prepare_data(path: Path, marker: str = "original") -> None:
    path.mkdir(parents=True, exist_ok=True)
    (path / "stroje.csv").write_text(MACHINES, encoding="utf-8")
    (path / "poruchy.csv").write_text(FAULTS, encoding="utf-8")
    photo = path / "photos" / "stroj_1" / "foto.jpg"
    photo.parent.mkdir(parents=True, exist_ok=True)
    photo.write_bytes(marker.encode("ascii"))


class BackupRestoreTests(unittest.TestCase):
    def test_complete_backup_contains_nested_data_and_excludes_logs(self):
        with test_directory("backup_complete") as temp:
            root = Path(temp)
            data = root / "data"
            prepare_data(data)
            logs = data / "logs"
            logs.mkdir(exist_ok=True)
            (logs / "sgm.log").write_text("diagnostika", encoding="utf-8")
            archive = root / "backup.zip"

            em.create_backup_archive(archive, data)
            info = em.inspect_backup_archive(archive)

            self.assertEqual(info["type"], "complete")
            with zipfile.ZipFile(archive) as zf:
                names = set(zf.namelist())
            self.assertIn("data/photos/stroj_1/foto.jpg", names)
            self.assertNotIn("data/logs/sgm.log", names)
            self.assertIn(em.BACKUP_MANIFEST, names)

    def test_restore_creates_safety_backup_and_restores_files(self):
        with test_directory("backup_restore") as temp:
            root = Path(temp)
            data = root / "data"
            backups = root / "safety"
            prepare_data(data, "from-backup")
            archive = root / "backup.zip"
            em.create_backup_archive(archive, data)

            (data / "stroje.csv").write_text(
                MACHINES.replace("Arburg", "Changed"), encoding="utf-8"
            )
            (data / "photos" / "stroj_1" / "foto.jpg").write_bytes(b"changed")
            safety = em.restore_backup_archive(archive, data, backups)

            self.assertIn("Arburg", (data / "stroje.csv").read_text("utf-8"))
            self.assertEqual(
                (data / "photos" / "stroj_1" / "foto.jpg").read_bytes(),
                b"from-backup",
            )
            self.assertTrue(safety.exists())
            safety_info = em.inspect_backup_archive(safety)
            self.assertEqual(safety_info["type"], "complete")

    def test_tampered_complete_backup_is_rejected(self):
        with test_directory("backup_tampered") as temp:
            root = Path(temp)
            data = root / "data"
            prepare_data(data)
            valid = root / "valid.zip"
            tampered = root / "tampered.zip"
            em.create_backup_archive(valid, data)

            with zipfile.ZipFile(valid) as source, zipfile.ZipFile(
                tampered, "w"
            ) as target:
                for info in source.infolist():
                    content = source.read(info.filename)
                    if info.filename == "data/stroje.csv":
                        content = content.replace(b"Arburg", b"Demag ")
                    target.writestr(info, content)

            with self.assertRaisesRegex(ValueError, "kontrolní součet"):
                em.inspect_backup_archive(tampered)

    def test_legacy_csv_backup_is_supported(self):
        with test_directory("backup_legacy") as temp:
            archive = Path(temp) / "legacy.zip"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("stroje.csv", MACHINES)
                zf.writestr("poruchy.csv", FAULTS)
            info = em.inspect_backup_archive(archive)
            self.assertEqual(info["type"], "legacy")
            self.assertEqual(info["file_count"], 2)

    def test_path_traversal_is_rejected(self):
        with test_directory("backup_unsafe") as temp:
            archive = Path(temp) / "unsafe.zip"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("../stroje.csv", MACHINES)
                zf.writestr("poruchy.csv", FAULTS)
            with self.assertRaisesRegex(ValueError, "Neplatná cesta"):
                em.inspect_backup_archive(archive)


if __name__ == "__main__":
    unittest.main()
