# -*- coding: utf-8 -*-
"""
Backup Excela PRIJE bilo kakvog upisa + automatsko brisanje starih backupa.
Pravilo projekta: NIKAD ne pisati u Excel bez prethodnog backupa.
"""
import shutil
from datetime import datetime
from pathlib import Path


def napravi_backup(excel_path, backup_dir, keep, logger):
    """Kopira Excel u backup mapu s datumom/vremenom u imenu i očisti višak.

    Ime: URA_BACKUP_2026-06-01_14h30.xlsm
    Vraća putanju napravljenog backupa.
    Diže iznimku ako kopiranje ne uspije (pozivatelj tada PREKIDA, ne dira Excel).
    """
    excel_path = Path(excel_path)
    backup_dir = Path(backup_dir)
    backup_dir.mkdir(parents=True, exist_ok=True)

    sada = datetime.now()
    ime = f"URA_BACKUP_{sada:%Y-%m-%d_%Hh%M}.xlsm"
    cilj = backup_dir / ime

    shutil.copy2(excel_path, cilj)  # copy2 čuva i datum izvorne datoteke
    logger.info("Backup napravljen: %s", cilj.name)

    _ocisti_stare(backup_dir, keep, logger)
    return cilj


def _ocisti_stare(backup_dir, keep, logger):
    """Zadrži samo 'keep' najnovijih backupa, ostale obriši."""
    if keep <= 0:
        return
    backupi = sorted(
        backup_dir.glob("URA_BACKUP_*.xlsm"),
        key=lambda p: p.stat().st_mtime,  # po vremenu izmjene, najstariji prvi
    )
    visak = backupi[:-keep]  # sve osim zadnjih 'keep'
    for f in visak:
        try:
            f.unlink()
            logger.info("Obrisan stari backup: %s", f.name)
        except OSError as e:
            logger.warning("Ne mogu obrisati stari backup %s: %s", f.name, e)
