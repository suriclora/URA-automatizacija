# -*- coding: utf-8 -*-
"""
Razvrstavanje skinutih datoteka po mapama:
  - XML i Parrin vizualni PDF  ->  Parra/{Mjesec}/
  - originalni PDF dobavljača  ->  pdf/  kao "UR XXXX.pdf"  (+ linka se u Excel)

Parrine datoteke prepoznajemo po imenu (imaju obrazac koji originalni PDF nema).
"""
import re
import shutil
import zipfile
from pathlib import Path

from src.utils import ocisti


def _tokeni(s):
    """Rastavi na dijelove po znakovima koji nisu slovo/broj. '30/1/2026' -> ['30','1','2026']."""
    return [t.lower() for t in re.split(r"[^0-9A-Za-z]+", str(s or "")) if t]


def ime_odgovara_racunu(br_rac, datoteka):
    """Strogo podudaranje imena datoteke s brojem računa (izbjegava lažne pogotke).
    Pravi pogodak je kad:
      - očišćeno ime POČINJE očišćenim brojem računa (npr. '2760_1_1.pdf' za '2760/1/1'), ILI
      - SVI dijelovi broja računa postoje među dijelovima imena datoteke.
    """
    br_clean = ocisti(br_rac)
    if len(br_clean) <= 3:
        return False
    if datoteka["ime_clean"].startswith(br_clean):
        return True
    inv_tok = _tokeni(br_rac)
    file_tok = set(_tokeni(datoteka["ime"]))
    if inv_tok and all(t in file_tok for t in inv_tok):
        return True
    # Parrin vizualni PDF '(e) Račun <broj>' nosi SAMO jezgru broja (npr. '2007' za '2007/1/1').
    m = re.match(r"\(e\)\s*ra[čc]un\s*0*(\d{4,})", str(datoteka["ime"]), re.I)
    return bool(m and inv_tok and m.group(1) == inv_tok[0])

# Hrvatski nazivi mjeseci (kao u postojećim Parra podmapama: Veljača, Ožujak...)
HR_MJESECI = {
    1: "Siječanj", 2: "Veljača", 3: "Ožujak", 4: "Travanj", 5: "Svibanj", 6: "Lipanj",
    7: "Srpanj", 8: "Kolovoz", 9: "Rujan", 10: "Listopad", 11: "Studeni", 12: "Prosinac",
}

# Obrasci u imenu koje koristi Parra (XML i vizualni PDF)
_PARRA_UZORCI = ("ulaznieracun", "ulazniracun", "incomingeinvoice")


def je_parra_ime(ime):
    """True ako ime datoteke odgovara Parrinom obrascu (XML ili Parrin PDF).
    Uključuje i Parrin vizualni PDF '(e) Račun <broj>.PDF' — i on je Parrin, pa ide
    u Parra/{mjesec} arhivu (a kopija u pdf/ kao 'UR XXXX.pdf')."""
    n = ime.lower()
    if any(u in n for u in _PARRA_UZORCI):
        return True
    return bool(re.match(r"\(e\)\s*ra[čc]un", n))


def skeniraj_sve(folder):
    """Vrati listu svih .pdf i .xml datoteka u mapi — uključujući i one UNUTAR .zip arhiva.
    Svaki zapis: ime, putanja, ext, ime_clean, sadrzaj, zip, inner.
    Za samostalne datoteke zip=None; za one iz arhive zip=putanja_zipa, inner=ime_u_zipu.
    """
    out = []
    p = Path(folder)
    if not p.exists():
        return out
    for f in p.iterdir():
        suf = f.suffix.lower()
        if suf in (".pdf", ".xml"):
            out.append({
                "ime": f.name, "putanja": str(f), "ext": suf,
                "ime_clean": ocisti(f.name), "sadrzaj": None, "zip": None, "inner": None,
            })
        elif suf == ".zip":
            # Otvori ZIP i ubaci .xml/.pdf datoteke iz njega (npr. Parrin "Popis e-URA")
            try:
                with zipfile.ZipFile(str(f)) as z:
                    for inner in z.namelist():
                        ie = Path(inner).suffix.lower()
                        if ie in (".pdf", ".xml"):
                            base = Path(inner).name
                            out.append({
                                "ime": base, "putanja": None, "ext": ie,
                                "ime_clean": ocisti(base), "sadrzaj": None,
                                "zip": str(f), "inner": inner,
                            })
            except Exception:
                pass  # oštećen/nepodržan zip — preskoči
    return out


def mjesec_mapa(parra_base, datum, stvaraj=True):
    """Vrati Path do Parra/{HrvatskiMjesec} za zadani datum.
    Ako 'stvaraj' i mapa ne postoji — stvori je.
    """
    ime = HR_MJESECI.get(datum.month, f"{datum.month:02d}")
    mapa = Path(parra_base) / ime
    if stvaraj:
        mapa.mkdir(parents=True, exist_ok=True)
    return mapa


def nadji_datoteke_racuna(br_rac, datoteke, citaj_sadrzaj_fn=None):
    """Pronađi datoteke koje pripadaju zadanom (SIROVOM) broju računa i razvrstaj ih.
    Vraća dict: {'xml': [...], 'parra_pdf': [...], 'original_pdf': [...]}.

    - XML i Parrin PDF prepoznajemo po IMENU (Parrin obrazac + broj računa u imenu).
    - Originalni PDF dobavljača: prvo po imenu, pa po SADRŽAJU (ako je dana funkcija
      za čitanje teksta PDF-a).
    """
    rez = {"xml": [], "parra_pdf": [], "original_pdf": []}
    br_rac_clean = ocisti(br_rac)
    if len(br_rac_clean) <= 3:
        return rez  # prekratak broj — preriskantno

    for d in datoteke:
        ime_match = ime_odgovara_racunu(br_rac, d)

        if d["ext"] == ".xml":
            if ime_match:
                rez["xml"].append(d)
            continue

        # .pdf
        if je_parra_ime(d["ime"]):
            if ime_match:
                rez["parra_pdf"].append(d)
        else:
            # originalni PDF dobavljača
            if ime_match:
                rez["original_pdf"].append(d)
            elif citaj_sadrzaj_fn is not None and d.get("putanja"):
                if d["sadrzaj"] is None:
                    d["sadrzaj"] = ocisti(citaj_sadrzaj_fn(d["putanja"]))
                if d["sadrzaj"] and _sadrzaj_se_poklapa(br_rac, br_rac_clean, d["sadrzaj"]):
                    rez["original_pdf"].append(d)
    return rez


def _sadrzaj_se_poklapa(br_rac, br_rac_clean, sadrzaj):
    """Pojavljuje li se broj računa u TEKSTU PDF-a? Puni očišćeni broj, ili distinktivna
    jezgra (≥5 znamenki) — npr. '133627' iz '133627-1-1_N2026030180'."""
    if br_rac_clean and br_rac_clean in sadrzaj:
        return True
    for tok in re.findall(r"\d{5,}", str(br_rac or "")):
        if tok in sadrzaj:
            return True
    return False


def sveukupno_za_platiti(tekst):
    """OTP leasing obrok-račun: izvuci iznos iz retka
    'Sveukupno za platiti - NN. leasing obrok XXX,XX EUR' (pun obrok = glavnica+kamata).
    Parra šalje samo 'Ukupno za račun' (kamatu), pa je iznos s Parre krivo nizak.
    Vrati float ili None (ako fraza ne postoji — npr. opomena, ili nije OTP)."""
    if not tekst:
        return None
    m = re.search(r"sveukupno za platiti.*?leasing obrok\s*(\d[\d.]*,\d{2})\s*EUR",
                  tekst, re.I | re.S)
    if not m:
        return None
    try:
        return float(m.group(1).replace(".", "").replace(",", "."))
    except ValueError:
        return None


def spremi(d, ciljna_mapa, novo_ime=None):
    """Spremi datoteku u ciljnu mapu:
      - samostalnu (zip=None) PREMJESTI,
      - onu iz ZIP-a IZVUCI (original ostaje u zipu).
    Vrati novu putanju.
    """
    ciljna_mapa = Path(ciljna_mapa)
    ciljna_mapa.mkdir(parents=True, exist_ok=True)
    cilj = ciljna_mapa / (novo_ime if novo_ime else d["ime"])
    if d.get("zip"):
        with zipfile.ZipFile(d["zip"]) as z, z.open(d["inner"]) as src, open(cilj, "wb") as dst:
            shutil.copyfileobj(src, dst)
    else:
        shutil.move(d["putanja"], str(cilj))
    return str(cilj)
