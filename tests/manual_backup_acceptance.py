"""Praktický test úplné zálohy a obnovy nad izolovanou kopií dat."""

import csv
import hashlib
import json
import shutil
import zipfile
from pathlib import Path

import data_manager as dm
import export_manager as em


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def csv_rows(path: Path) -> int:
    with open(path, newline="", encoding="utf-8-sig") as source:
        return sum(1 for _ in csv.DictReader(source))


def main() -> None:
    runtime_root = Path(__file__).parent / "_runtime"
    test_root = runtime_root / "acceptance"
    if test_root.parent.resolve() != runtime_root.resolve():
        raise RuntimeError("Neplatná cílová složka přejímacího testu.")
    if test_root.exists():
        shutil.rmtree(test_root)
    test_root.mkdir(parents=True)

    work_data = test_root / "work_data"
    archive = test_root / "complete_backup.zip"
    safety_dir = test_root / "safety"
    shutil.copytree(dm.DATA_DIR, work_data, ignore=shutil.ignore_patterns("logs"))

    em.create_backup_archive(archive, work_data)
    backup_info = em.inspect_backup_archive(archive)
    with zipfile.ZipFile(archive) as backup:
        manifest = json.loads(backup.read(em.BACKUP_MANIFEST))

    expected = {
        relative: metadata["sha256"]
        for relative, metadata in manifest["files"].items()
    }
    original_machines = (work_data / "stroje.csv").read_bytes()
    (work_data / "stroje.csv").write_bytes(original_machines + b"\n")

    safety_archive = em.restore_backup_archive(archive, work_data, safety_dir)
    restored = {
        relative: sha256(work_data / Path(*relative.split("/")))
        for relative in expected
    }
    mismatches = [
        relative
        for relative, expected_hash in expected.items()
        if restored.get(relative) != expected_hash
    ]
    if mismatches:
        raise AssertionError(
            "Po obnově nesouhlasí soubory: " + ", ".join(mismatches)
        )
    if (work_data / "stroje.csv").read_bytes() != original_machines:
        raise AssertionError("CSV stroje.csv se nevrátilo do původního stavu.")

    safety_info = em.inspect_backup_archive(safety_archive)
    result = {
        "machines": csv_rows(work_data / "stroje.csv"),
        "faults": csv_rows(work_data / "poruchy.csv"),
        "backup_files": backup_info["file_count"],
        "backup_bytes_uncompressed": backup_info["total_size"],
        "backup_zip": str(archive),
        "safety_backup": str(safety_archive),
        "safety_files": safety_info["file_count"],
        "verified_files": len(expected),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
