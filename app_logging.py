#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def configure_logging(data_dir: Path) -> Path | None:
    """Zapne rotační log aplikace; selhání logu nesmí zabránit spuštění."""
    logger = logging.getLogger("sgm")
    if logger.handlers:
        handler = logger.handlers[0]
        return Path(handler.baseFilename) if hasattr(handler, "baseFilename") else None

    try:
        log_dir = Path(data_dir) / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "sgm.log"
        handler = RotatingFileHandler(
            log_path,
            maxBytes=1_000_000,
            backupCount=3,
            encoding="utf-8",
        )
        handler.setFormatter(
            logging.Formatter(
                "%(asctime)s %(levelname)s %(name)s: %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )
        logger.setLevel(logging.INFO)
        logger.addHandler(handler)
        logger.propagate = False
        logger.info("Aplikace spuštěna")
        return log_path
    except OSError:
        logging.getLogger(__name__).exception("Log aplikace nelze inicializovat")
        return None


def install_tk_exception_handler(root, translate=lambda cz, de: de):
    """Zapíše neošetřenou chybu Tkinter callbacku a informuje uživatele."""

    def report_callback_exception(exc_type, exc_value, traceback):
        logging.getLogger("sgm.ui").critical(
            "Neošetřená chyba uživatelského rozhraní",
            exc_info=(exc_type, exc_value, traceback),
        )
        try:
            from tkinter import messagebox

            messagebox.showerror(
                translate("Neočekávaná chyba", "Unerwarteter Fehler"),
                translate(
                    "Operace se nepodařila. Podrobnosti jsou v data/logs/sgm.log.",
                    "Vorgang fehlgeschlagen. Details stehen in data/logs/sgm.log.",
                ),
                parent=root,
            )
        except Exception:
            logging.getLogger("sgm.ui").exception(
                "Chybové hlášení uživateli se nepodařilo zobrazit"
            )

    root.report_callback_exception = report_callback_exception
