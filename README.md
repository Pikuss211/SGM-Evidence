# SGM-Jokey

Maintenance helper for injection molding machines (SGM).

Current application release: **SGM Evidence 1.2.0**

The Windows executable and application window use the SGM-Jokey tools icon
stored in `assets/sgm_jokey.ico`.

## Features
- logging of machine faults
- electrical / mechanical classification
- quick search of previous repairs
- simple Tkinter interface
- export to CSV or PDF

## Technology
Python + Tkinter

## Project structure

- `SGM_v1.1-de_clean_fixed.py` – application window and machine/fault workflows;
- `data_manager.py` – CSV persistence and data-domain helpers;
- `data_validator.py` – consistency checks and conservative repairs;
- `validation_ui.py` – data-check dialog;
- `audit_manager.py` – persistent audit trail of application changes;
- `audit_ui.py` – global and per-machine change-history viewer;
- `operator_manager.py` – atomically persisted operator/session settings;
- `operator_ui.py` – operator selection dialog used at startup and shift change;
- `export_manager.py` – PDF/CSV export, backup and restore;
- `app_logging.py` – rotating runtime log and Tkinter exception handling.

## Data safety

- CSV files are saved atomically; the previous version remains next to the file
  with the `.bak` suffix.
- A backup ZIP contains the complete `data` directory including machine
  documents and photos. Diagnostic logs are intentionally excluded.
- Every new backup contains a manifest with file sizes and SHA-256 checksums.
- Restore validates the archive and required CSV columns before changing data.
- Before restore, the current data is automatically saved in the `backups`
  directory. Legacy backups containing only the three CSV files remain
  supported.
- Runtime errors are written to `data/logs/sgm.log` with automatic log rotation.

## Data validation

Use the `Kontrola dat / Datenprüfung` button in the main toolbar to check:

- missing or duplicate machine and fault IDs;
- faults linked to unknown machines;
- invalid states, categories, dates and maintenance intervals;
- missing required CSV files and columns.

The repair button only normalizes unambiguous values. It does not automatically
change duplicates, broken links or unclear values. A complete backup is created
in `backups` before every repair, and multi-file changes are rolled back if a
write fails.

## Audit history and machine archive

- New application changes are recorded in `data/audit_log.csv`, including the
  timestamp, operator, action, affected record and changed fields.
- At startup, the active operator must be confirmed. Use the operator button in
  the status bar when the shift changes; recent names are remembered.
- Use `Historie změn / Änderungsverlauf` for the complete history, or open a
  machine's context menu for its filtered history.
- Removing a machine now archives it instead of deleting it. Its faults, photos
  and documents stay intact.
- Archived machines are hidden by default. Enable `Zobrazit archivované /
  Archivierte anzeigen`, then use the context menu to restore one.
- Older `stroje.csv` files without the `archivovan` column remain compatible;
  existing machines are treated as active.

## Tests

Run the built-in test suite without touching production data:

```powershell
python -m unittest discover -s tests -t . -v
```

Run the practical backup/restore acceptance test on an isolated copy of the
current data:

```powershell
python -m tests.manual_backup_acceptance
```

Run the read-only Tkinter smoke test in both interface languages:

```powershell
$env:SGM_LANG='cz'; python -m tests.manual_ui_smoke
$env:SGM_LANG='de'; python -m tests.manual_ui_smoke
```

## Windows build

Install the pinned build dependencies:

```powershell
python -m pip install -r requirements.txt
```

Create the portable Windows executable from the checked build specification:

```powershell
python -m PyInstaller --noconfirm --clean SGM_Evidence.spec
```

Keep the writable `data` directory next to `SGM Evidence 1.2.0.exe`.
