# -*- coding: utf-8 -*-
"""
Središnja konfiguracija projekta — SVE putanje, sklopke i konstante na jednom mjestu.

VAŽNO (sigurnost):
  - Nikakve tajne (token, lozinke, OIB) NISU u kodu — čitaju se iz `.env`.
  - Zadano je DEMO način: aplikacija radi na potpuno izmišljenim podacima u
    `demo/data/` i NE treba ni server ni pravi API token.
  - Za pravi rad postavi `DEMO=0` i prave vrijednosti u `.env` (vidi `.env.example`).
"""
import os
from datetime import date
from pathlib import Path
from dotenv import load_dotenv

# Korijen projekta = mapa u kojoj se nalazi ovaj config.py
KORIJEN = Path(__file__).resolve().parent
load_dotenv(KORIJEN / ".env")


def _env(key, default=""):
    return os.getenv(key, default)


# =====================================================================
#  NAČIN RADA
# =====================================================================
# DEMO=1 (zadano): radi na IZMIŠLJENIM podacima u demo/data (bez servera i tokena).
# DEMO=0: pravi rad — putanje i tajne se čitaju iz .env.
DEMO = _env("DEMO", "1") == "1"

# True = radi na kopiji/demo knjizi; False = prava knjiga na serveru.
KORISTI_TEST_EXCEL = DEMO

# Skripta obrađuje SAMO račune/izvode OD ovog datuma (stari ručni zaostatak se ne dira).
OBRADJUJ_OD = date(2026, 5, 1)
# Zasebna granica za bankovne izvode (računi mogu ići od ranije, izvodi tek od ovog datuma).
IZVODI_OD = date(2026, 5, 1)

# True = samo ISPIŠI što bi se napravilo (ništa se ne sprema). False = stvarni rad.
SIMULACIJA = False

# Runtime zastavice — postavlja ih main.py tijekom rada (ne diraj ovdje).
PISI_EXCEL = False
STVARNO_DATOTEKE = False

# =====================================================================
#  TAJNE (iz .env — NIKAD upisivati vrijednosti u kod!)
# =====================================================================
PARRA_API_TOKEN = _env("PARRA_API_TOKEN")
PARRA_BUSINESS_ID = _env("PARRA_BUSINESS_ID", "00000")
PARRA_BASE_URL = _env("PARRA_BASE_URL", "https://api.parra.hr/v1.1")
PARRA_PAGE_SIZE = 50       # koliko računa po stranici
PARRA_MAX_STRANICA = 100   # sigurnosna granica

# Lozinka zaštite listova u Excelu putnih naloga (čita se iz .env).
PN_LOZINKA = _env("PN_LOZINKA")
# OIB tvrtke — koristi se za prepoznavanje bankovnih izvoda (iz .env; demo placeholder).
MOJ_OIB = _env("MOJ_OIB", "12345678901")

# =====================================================================
#  PREPOZNAVANJE / DEDUP
# =====================================================================
PDF_MATCH_SADRZAJ = True
DEDUP_TOL_IZNOS = 0.02
DEDUP_DANI_DUPLI = 7
DEDUP_DANI_NEIZVJESNO = 45

GLAVNI_SHEET = "UR"
HEADER_RED = 2   # zaglavlje u 2. redu; podaci od 3. reda

# =====================================================================
#  PUTANJE
# =====================================================================
_DEMO = KORIJEN / "demo" / "data"

if DEMO:
    # Sve pokazuje na izmišljeni skup podataka unutar projekta.
    EXCEL_PATH = str(_DEMO / "URA_demo.xlsx")
    FOLDER_PDF_FINAL = str(_DEMO / "pdf")
    FOLDER_PARRA = str(_DEMO / "Parra")
    FOLDER_DOWNLOADS = str(_DEMO / "Downloads")
    FOLDER_IZVODI_HPB = str(_DEMO / "Izvodi" / "HPB")
    FOLDER_IZVODI_RBA = str(_DEMO / "Izvodi" / "RBA")
    FOLDER_PN = str(_DEMO / "PN")
    FOLDER_PN_PRAVI = str(_DEMO / "PN")
    FOLDER_FOTKE = str(_DEMO / "fotke")
    TERENI_PATH = str(_DEMO / "tereni_demo.xlsx")
    BACKUP_DIR = str(_DEMO / "backup")
    STANJE_PATH = str(_DEMO / "stanje_demo.json")
else:
    # PRAVI rad — postavi u .env. Primjeri su GENERIČKI placeholderi (nisu stvarni).
    EXCEL_PATH = _env("EXCEL_PATH", r"\\SERVER\Financije\URA_2026.xlsm")
    FOLDER_PDF_FINAL = _env("FOLDER_PDF_FINAL", r"\\SERVER\Financije\Racuni\pdf")
    FOLDER_PARRA = _env("FOLDER_PARRA", r"\\SERVER\Financije\Racuni\Parra")
    FOLDER_DOWNLOADS = _env("FOLDER_DOWNLOADS", str(Path.home() / "Downloads"))
    FOLDER_IZVODI_HPB = _env("FOLDER_IZVODI_HPB", r"\\SERVER\Financije\Izvodi\HPB")
    FOLDER_IZVODI_RBA = _env("FOLDER_IZVODI_RBA", r"\\SERVER\Financije\Izvodi\RBA")
    FOLDER_PN_PRAVI = _env("FOLDER_PN", r"\\SERVER\Financije\PN")
    FOLDER_PN = FOLDER_PN_PRAVI
    FOLDER_FOTKE = _env("FOLDER_FOTKE", str(Path.home() / "Desktop" / "Racuni"))
    TERENI_PATH = _env("TERENI_PATH", r"\\SERVER\share\tereni.xlsx")
    BACKUP_DIR = _env("BACKUP_DIR", r"\\SERVER\Financije\Backup")
    STANJE_PATH = str(KORIJEN / "stanje.json")

LOG_DIR = str(KORIJEN / "logs")
BACKUP_KEEP = 20


def pn_excel_putanja(godina, mjesec):
    """Putanja do mjesečnog Excela putnih naloga u koji se PIŠE."""
    return os.path.join(FOLDER_PN, str(godina), f"pn {mjesec:02d}-{godina}.xlsx")


def pn_excel_pravi(godina, mjesec):
    """Putanja do mjesečnog Excela putnih naloga (izvor/predložak)."""
    return os.path.join(FOLDER_PN_PRAVI, str(godina), f"pn {mjesec:02d}-{godina}.xlsx")


# =====================================================================
#  VOZILA  (generički primjeri — porezni tretman: teretni=PDV 100%, osobni=PDV 50%)
# =====================================================================
VOZILA = [
    ("Kombi 1 (teretni)", "teretni"),
    ("Kombi 2 (teretni)", "teretni"),
    ("Dostavno 1 (teretni)", "teretni"),
    ("Osobno 1", "osobni"),
    ("Osobno 2", "osobni"),
    ("(nije vezano uz auto)", None),
]


def vozilo_porez(naziv):
    """Vrati 'osobni' / 'teretni' / None za zadani naziv vozila."""
    for ime, tip in VOZILA:
        if ime == naziv:
            return tip
    return None


# Mapa: naziv vozila u tablici terena (malim slovima) -> (puni naziv za nalog, ENC nadimak ili None).
VOZILO_MAP = {
    "kombi1":   ("Kombi 1  ZG 1000 AA", "kombi1"),
    "kombi2":   ("Kombi 2  ZG 2000 BB", "kombi2"),
    "dostavno": ("Dostavno 1  ZG 3000 CC", "dostavno"),
    "osobno1":  ("Osobno 1  ZG 4000 DD", None),
    "osobno2":  ("Osobno 2  ZG 5000 EE", None),
}


def vozilo_info(naziv):
    """Vrati (puni naziv vozila, ENC nadimak) za naziv iz tablice terena. Fallback: prva riječ."""
    s = (naziv or "").strip().lower()
    if s in VOZILO_MAP:
        return VOZILO_MAP[s]
    prva = s.split()[0] if s.split() else ""
    return VOZILO_MAP.get(prva, (naziv, None))
