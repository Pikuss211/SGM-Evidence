#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import csv
import logging
import os
import shutil
import sys
from datetime import date, datetime, timedelta
from pathlib import Path


LANG = os.environ.get("SGM_LANG", "de").strip().lower()
FMT = "%Y-%m-%d %H:%M"


def T(cz: str, de: str | None = None) -> str:
    if LANG.upper() == "DE":
        return de if de is not None else cz
    return cz


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

SOUBOR_STROJE = DATA_DIR / "stroje.csv"
SOUBOR_PORUCHY = DATA_DIR / "poruchy.csv"
SOUBOR_SABLONY = DATA_DIR / "sablony_alarmu.csv"
PORUCHY_FIELDNAMES = [
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
LOGGER = logging.getLogger("sgm.data")


def _atomic_csv_write(path: Path, fieldnames: list[str], rows) -> None:
    """Zapíše CSV bezpečně přes dočasný soubor ve stejné složce."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f".{path.name}.tmp")
    try:
        with open(
            temp_path,
            "w",
            newline="",
            encoding="utf-8",
        ) as tmp:
            writer = csv.DictWriter(tmp, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
            tmp.flush()
            os.fsync(tmp.fileno())

        # Ověření, že dočasný soubor obsahuje očekávanou CSV hlavičku.
        with open(temp_path, newline="", encoding="utf-8") as check:
            actual_header = next(csv.reader(check), [])
        if actual_header != fieldnames:
            raise OSError(f"Neplatná CSV hlavička v {temp_path.name}")

        if path.exists():
            shutil.copy2(path, path.with_suffix(path.suffix + ".bak"))
        os.replace(temp_path, path)
    except Exception:
        LOGGER.exception("Bezpečné uložení CSV selhalo: %s", path)
        raise
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            LOGGER.warning(
                "Dočasný soubor nešel odstranit: %s", temp_path, exc_info=True
            )


def slozka_stroje(cislo: str) -> Path:
    p = DATA_DIR / "soubory" / str(cislo).zfill(2)
    p.mkdir(parents=True, exist_ok=True)
    return p


def nacti_stroje():
    if not SOUBOR_STROJE.exists():
        return {}

    with open(SOUBOR_STROJE, newline="", encoding="utf-8-sig") as f:
        sample = f.read(2048)
        f.seek(0)

        delimiter = ";" if ";" in sample and "," not in sample else ","
        r = csv.DictReader(f, delimiter=delimiter)

        stroje = {}
        for row in r:
            cislo = str(row.get("cislo", "")).strip()
            if not cislo:
                continue

            row.setdefault("wartung_last", "")
            row.setdefault("wartung_interval", "180")
            if not str(row.get("archivovan", "") or "").strip():
                row["archivovan"] = "0"

            stroje[cislo] = row
        return stroje


def uloz_stroje(stroje: dict):
    fieldnames = [
        "cislo",
        "vyrobce",
        "typ",
        "rok",
        "spm",
        "seriove",
        "stav",
        "wartung_last",
        "wartung_interval",
        "archivovan",
    ]

    rows = []
    for cislo, s in stroje.items():
        row = dict(s)
        row["cislo"] = cislo
        row["archivovan"] = "1" if is_archived_machine(row) else "0"
        rows.append(row)
    _atomic_csv_write(SOUBOR_STROJE, fieldnames, rows)


def nacti_poruchy():
    if not SOUBOR_PORUCHY.exists() or SOUBOR_PORUCHY.stat().st_size == 0:
        return []

    with open(SOUBOR_PORUCHY, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        if not r.fieldnames:
            return []

        poruchy = []
        for row in r:
            row["kategorie"] = normalize_kategorie(row.get("kategorie"))
            row["cas"] = normalize_dt(row.get("cas"))
            row["cas_uzavreni"] = normalize_dt(row.get("cas_uzavreni"))
            poruchy.append(row)

        return poruchy


def uloz_poruchy(poruchy: list):
    fieldnames = list(PORUCHY_FIELDNAMES)
    if poruchy:
        extra = sorted(
            {
                key
                for row in poruchy
                for key in row.keys()
                if key not in fieldnames
            }
        )
        fieldnames.extend(extra)
    elif SOUBOR_PORUCHY.exists():
        with open(SOUBOR_PORUCHY, newline="", encoding="utf-8") as f:
            existing = csv.reader(f)
            existing_header = next(existing, [])
        if existing_header:
            fieldnames = existing_header

    _atomic_csv_write(SOUBOR_PORUCHY, fieldnames, poruchy)


def nacti_sablony():
    if not SOUBOR_SABLONY.exists():
        return {}
    with open(SOUBOR_SABLONY, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        return {row["alarm"]: row["reseni"] for row in r if row.get("alarm")}


def normalize_dt(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    tried = ["%Y-%m-%d %H:%M", "%d.%m.%Y %H:%M",
             "%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d %H:%M"]
    for f in tried:
        try:
            return datetime.strptime(s, f).strftime(FMT)
        except Exception:
            pass
    return s


def normalize_stav(s: str) -> str:
    s = (s or "").strip().lower()

    if s in ("b", "běží", "bezi", "l", "läuft", "laeuft", "lauf", "running", "ok"):
        return "bezi"

    if s in ("p", "porucha", "s", "störung", "stoerung", "fault", "error"):
        return "porucha"

    return "bezi"


def stav_ui(value: str) -> str:
    key = normalize_stav(value)
    if key == "bezi":
        return T("běží", "läuft")
    if key == "porucha":
        return T("porucha", "Störung")
    return value or ""


def porucha_stav_ui(value: str) -> str:
    s = (value or "").strip().lower()
    if s in ("otevrena", "offen", "o"):
        return T("otevřená", "offen")
    if s in ("uzavrena", "geschlossen", "g"):
        return T("uzavřená", "geschlossen")
    return value or ""


def normalize_kategorie(s: str) -> str:
    s = (s or "").strip().lower()

    if s in ("e", "elektricka", "elektrická", "electrical", "elektrisch"):
        return "elektricka"

    if s in ("m", "mechanicka", "mechanická", "mechanical", "mechanisch"):
        return "mechanicka"

    if s in ("j", "jina", "jiná", "other", "sonstige", "andere"):
        return "jina"

    return "jina"


def kat_ui(kat: str) -> str:
    k = normalize_kategorie(kat)
    if k == "elektricka":
        return T("elektrická", "elektrisch")
    if k == "mechanicka":
        return T("mechanická", "mechanisch")
    return T("jiná", "sonstige")


def days_to_next_wartung(stroj: dict):
    s = (stroj.get("wartung_last") or "").strip()
    if not s:
        return None

    try:
        last = datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None

    try:
        interval = int(stroj.get("wartung_interval") or 180)
    except ValueError:
        interval = 180

    next_date = last + timedelta(days=interval)
    today = date.today()
    return (next_date - today).days


def last_open_dt(poruchy: list, cislo: str):
    best = None
    for p in poruchy:
        if str(p.get("cislo")) != str(cislo):
            continue
        if p.get("stav") != "otevrena":
            continue

        s = (p.get("cas") or "").strip()
        try:
            dt = datetime.strptime(s, "%Y-%m-%d %H:%M")
        except Exception:
            dt = None

        if dt is not None and (best is None or dt > best):
            best = dt

    return best if best is not None else datetime.min


COLORS = {
    "ok": "#c7f1d0",
    "elektricka": "#f8c2c2",
    "mechanicka": "#c2d4f8",
    "jina": "#f8f4c2",
}


def color_by_cat(cat: str) -> str:
    k = normalize_kategorie(cat)
    return COLORS.get(k, COLORS["ok"])


def nove_id(poruchy: list) -> str:
    nums = [int(p["id"]) for p in poruchy if str(p.get("id", "")).isdigit()]
    return str(max(nums) + 1 if nums else 1)


def last_open_issue(poruchy: list, cislo: str):
    opened = [p for p in poruchy if p.get("cislo") == str(cislo) and p.get("stav") == "otevrena"]
    if not opened:
        return None

    def _key(p):
        try:
            return datetime.strptime(p.get("cas", ""), "%Y-%m-%d %H:%M")
        except Exception:
            return datetime.min

    return sorted(opened, key=_key)[-1]


def next_free_machine_number(stroje: dict) -> str:
    used = set()
    for k in stroje.keys():
        s = str(k).strip()
        if s.isdigit():
            used.add(int(s))
    n = 1
    while n in used:
        n += 1
    return str(n)


def is_archived_machine(stroj: dict) -> bool:
    value = str(stroj.get("archivovan", "0") or "0").strip().lower()
    return value in ("1", "ano", "true", "yes", "ja", "archivovan", "archiviert")


def barva_dlazdice(stav: str, open_count: int, cislo: str, poruchy: list) -> str:
    if open_count <= 0 and normalize_stav(stav) == "bezi":
        return COLORS["ok"]
    last = last_open_issue(poruchy, cislo)
    return color_by_cat(last.get("kategorie") if last else "ok")
