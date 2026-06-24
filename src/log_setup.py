# -*- coding: utf-8 -*-
"""Postavljanje logiranja: piše i u konzolu (uživo) i u dnevnu log datoteku."""
import logging
import sys
from datetime import datetime
from pathlib import Path


def postavi_logging(log_dir):
    """Vrati logger koji istovremeno piše u:
      - konzolu (da korisnik uživo prati što se događa)
      - datoteku logs/ura_YYYY-MM-DD.log (za kasniju provjeru)
    """
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    datoteka = log_dir / f"ura_{datetime.now():%Y-%m-%d}.log"

    logger = logging.getLogger("ura")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()  # da se kod ponovnog pokretanja ne dupliraju poruke

    # 1) U datoteku: s vremenom i razinom (INFO/ERROR...)
    fh = logging.FileHandler(datoteka, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                      "%Y-%m-%d %H:%M:%S"))
    logger.addHandler(fh)

    # 2) U konzolu: čisto i čitljivo (bez tehničkih prefiksa)
    # Osiguraj da konzola podnosi hrvatske znakove (č, ć, š, ž, đ)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(ch)

    return logger
