# -*- coding: utf-8 -*-
"""
Auto-prijedlog VRSTE TROŠKA.

Pravilo (dogovoreno s korisnikom):
  1. Ako dobavljač u POVIJESTI (knjizi) UVIJEK ima istu vrstu → vrati tu vrstu.
  2. Ako VARIRA → pročitaj sa stavki računa i, ako je moguće, uskladi s jednom
     od vrsta koje je već koristio za tog dobavljača; inače pokušaj iz ključnih riječi.
  3. Ako se ne može pouzdano → vrati None (korisnik upiše).

Bolje PRAZNO nego krivo.
"""
import re
from collections import defaultdict, Counter
from src.utils import ocisti

# Gorivni brendovi (provjeravaju se kao ZASEBNA riječ — da 'kupovina' ne ulovi 'ina')
_FUEL = {"tifon", "ina", "petrol", "crodux", "lukoil", "shell", "mol", "eurosuper"}

# Ključne riječi (podniz u opisu) -> VRSTA TROŠKA. Samo dovoljno duge/sigurne osnove.
_KW = [
    ("čišćenj", "čišćenje"), ("ciscenj", "čišćenje"),
    ("osiguranj", "osiguranje"),
    ("cestarin", "cestarina"), ("autocest", "cestarina"),
    ("najamnin", "najam"), ("zakup", "najam"),
    ("opomen", "leasing"), ("leasing", "leasing"),
    ("električn", "električna energija"), ("elektricn", "električna energija"),
    ("tehnički pregled", "tehnički pregled"), ("tehnicki pregled", "tehnički pregled"),
    ("registracij", "registracija"),
    ("knjigovod", "knjigovodstvo"),
    ("noćenj", "noćenje-pn"), ("nocenj", "noćenje-pn"),
    ("prijevoz", "prijevoz"),
]

# KPD prefiksi -> VRSTA (Klasifikacija proizvoda po djelatnostima)
_KPD = [
    ("65.12", "osiguranje"), ("65.1", "osiguranje"),
    ("52.21", "cestarina"),
    ("81.2", "čišćenje"),
    ("77.11", "najam"),
    ("35.", "električna energija"),
    ("69.20", "knjigovodstvo"),
]


def nauci_iz_knjige(ws_val, header_red):
    """Iz knjige (data_only) nauči: dobavljač_clean -> Counter(vrsta)."""
    h = defaultdict(Counter)
    for i in range(header_red + 1, ws_val.max_row + 1):
        dob = ws_val.cell(row=i, column=5).value
        vr = ws_val.cell(row=i, column=8).value
        if dob and vr:
            v = str(vr).strip().lower()
            if v:
                h[ocisti(dob)][v] += 1
    return h


_DIJA = str.maketrans("čćžšđ", "cczsd")


def _fold(s):
    """Bez kvačica + mala slova (za usporedbu: 'perić' i 'peric' isto)."""
    return ocisti(s).translate(_DIJA)


def nauci_imena(ws, header_red):
    """Iz knjige nauči ustaljene nazive dobavljača: {bez_kvačica -> naziv (mala slova)}.
    Tako za 'PRIMJER d.o.o.' vratimo tvoj naziv 'primjer'."""
    imena = {}
    for i in range(header_red + 1, ws.max_row + 1):
        dob = ws.cell(row=i, column=5).value
        if dob:
            c = _fold(dob)
            if len(c) >= 4:
                imena.setdefault(c, str(dob).strip().lower())
    return imena


def kanonsko_ime(naziv, imena):
    """Vrati ustaljeni (kratki) naziv iz knjige; makni adresu (iza prvog zareza).
    Npr. 'PRIMJER d.o.o.' -> 'primjer'; 'PERO PERIĆ, ADRESA...' -> 'pero perić' (iz knjige)."""
    if not naziv:
        return naziv
    osnova = str(naziv).split(",")[0].strip()       # samo ime/firma, bez adrese
    nc = _fold(osnova)
    if not nc:
        return osnova.lower()
    if nc in imena:                                  # točno poklapanje (bez kvačica)
        return imena[nc]
    kand = [(len(c), ime) for c, ime in imena.items()
            if nc.startswith(c) or (len(c) >= 5 and c in nc)]
    if kand:
        kand.sort()                                  # najkraći ustaljeni naziv (npr. 'primjer')
        return kand[0][1]
    return osnova.lower()


def _nadji_povijest(dob_clean, history):
    if not dob_clean:
        return None
    if dob_clean in history:
        return history[dob_clean]
    for k in history:
        if dob_clean[:10] in k or k[:10] in dob_clean:
            return history[k]
    return None


def _iz_teksta(tekst, kpd=""):
    t = (tekst or "").lower()
    rijeci = set(re.findall(r"[a-zčćžšđ]+", t))
    # gorivo i ENC — kao ZASEBNE riječi
    if (rijeci & _FUEL) or any(w in t for w in ("gorivo", "dizel", "benzin", "nafta")):
        return "gorivo"
    if "enc" in rijeci:
        return "cestarina"
    for kw, vrsta in _KW:
        if kw in t:
            return vrsta
    for pref, vrsta in _KPD:
        if (kpd or "").startswith(pref):
            return vrsta
    return None


def predlozi_vrstu(dobavljac, history, stavke_tekst="", kpd=""):
    """Vrati predloženu VRSTU TROŠKA ili None (ostavi prazno)."""
    h = _nadji_povijest(ocisti(dobavljac), history)
    if h:
        if len(h) == 1:                       # uvijek ista vrsta -> sigurno
            return list(h)[0]
        # varira -> pokušaj prepoznati koju vrstu opisuje stavka (njegovim riječima)
        t = (stavke_tekst or "").lower()
        for vr, _ in h.most_common():
            rijeci = [w for w in vr.split() if len(w) > 3]
            if rijeci and all(w in t for w in rijeci):
                return vr
        # ne znamo iz stavke -> probaj ključne riječi, pa prazno
        return _iz_teksta(stavke_tekst, kpd)
    # nema povijesti -> ključne riječi / KPD
    return _iz_teksta(stavke_tekst, kpd)
