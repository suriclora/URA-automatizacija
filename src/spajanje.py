# -*- coding: utf-8 -*-
"""
FAZA 4 — spajanje stavki bankovnog izvoda s računima u Excelu.

Za svaku stavku izvoda (samo Duguje/isplate; uplate nama se preskaču) odredi
prijedlog akcije po prioritetnom redoslijedu (korak D iz brifa):

  D1  smjer 'P' (uplata nama)            -> PRESKOČI (izlazni račun)
  D3  opis sadrži 'plaća'                -> NOVI UR0 red (VRSTA='plaća')
  D2  'pn' u opisu/pozivu                -> DNEVNICE (spoji na postojeće pn redove)
      porez/doprinos/proračun/HZZO       -> NOVI UR0 red (VRSTA='porez'/'zdravstveno…')
  D4  iznos+datum(+broj) -> 1 račun      -> SPOJI (auto)
  D5  HR00/HR99 + lokacija (PETROL…)     -> KARTICA: fuzzy, PITAJ
  D6  ostalo: fuzzy                      -> 1 kandidat: PITAJ; inače UR0 crveno

Ovo je logika ODLUČIVANJA; upis radi main (i to tek nakon backupa).
"""
import re
from datetime import datetime, date
from rapidfuzz import fuzz

from src.utils import ocisti

# Lokacije (kartično/gotovinsko plaćanje na licu mjesta) -> VRSTA hint
_GORIVO = ("petrol", "ina", "crodux", "tifon", "lukoil", "shell", "mol ", "europetrol")
_TRGOVINE = ("konzum", "kaufland", "lidl", "spar", "tommy", "plodine", "bipa", "dm ",
             "muller", "müller", "pevex", "bauhaus", "ikea", "metro", "eurospin", "ntl",
             "studenac", "tisak", "inovine", "ribola", "ktc", "metss")
_LOKACIJE = _GORIVO + _TRGOVINE

_PRAZAN_POZIV = ("hr00", "hr99")

# Prag sličnosti imena dobavljača (rapidfuzz partial_ratio) za potvrdu spoja uz broj
SIM_IME = 80
# Kod računa BEZ broja na izvodu (dućan/benzinska): koliko dana smije razlika datuma
# (datum računa = PRVI datum na izvodu) da se spoji po ime+datum+iznos
SPOJ_DANI = 10

# Poznati trgovci na karticama (DOBAVLJAČ se čita iz opisa, ne iz naziva = banka)
_TRGOVCI = ["tifon", "crodux", "lukoil", "petrol", "shell", "konzum", "kaufland", "lidl",
            "spar", "plodine", "tommy", "bipa", "muller", "pevex", "bauhaus", "ikea",
            "studenac", "eurospin", "ribola", "intersport", "ina", "dm", "ntl", "ktc"]


_TRG_SUM = {"int", "othr", "kupovina", "pm", "bp", "st", "p", "zg", "ac"}


_FUEL = {"tifon", "ina", "petrol", "crodux", "lukoil", "shell", "mol", "eurosuper"}


def je_gorivo(opis):
    """True ako opis sadrži gorivni brend kao ZASEBNU riječ (da 'kupovina' ne ulovi 'ina')."""
    o = (opis or "").lower()
    rijeci = set(re.findall(r"[a-zčćžšđ]+", o))
    return bool(rijeci & _FUEL) or any(w in o for w in ("gorivo", "dizel", "benzin", "nafta"))


def trgovac_iz_opisa(opis):
    """Vrati ime trgovca s kartičnog plaćanja iz opisa.
    1) poznati brend kao zasebna riječ ('tifon','ina'...) — čisto ime;
    2) inače generički: riječi iza 'KUPOVINA' (RBA) ili iza vodeće šifre (HPB).
    Tako 'ina' ne ulovi 'benzina', a '9505PETROL' i dalje da 'petrol'."""
    o = opis or ""
    rijeci = set(re.findall(r"[a-zčćžšđ]+", o.lower()))
    for t in _TRGOVCI:
        if t in rijeci:
            return t
    # generički: iza "KUPOVINA ..." do ponavljanja, ili makni vodeće znamenke (HPB šifra)
    m = re.search(r"kupovina\s+(.+?)(?:\s+kupovina|$)", o, re.I)
    kand = m.group(1) if m else re.sub(r"^\s*\d+", "", o)
    tok = [w for w in re.findall(r"[A-Za-zčćžšđ.&\-]+", kand) if w.lower() not in _TRG_SUM]
    return " ".join(tok[:2]).lower() if tok else None


def _datum(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _num(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def izgradi_indeks(ws, header_red):
    """Indeks postojećih redova (ws mora biti data_only):
      racuni: redovi s RAČUN BR (kandidati za spajanje), s iznosom/datumom
      dnevnice: {dobavljač_clean: [redovi VRSTA='dnevnice' bez izvatka]}
    """
    racuni, dnevnice = [], {}
    for i in range(header_red + 1, ws.max_row + 1):
        racun = ws.cell(row=i, column=2).value
        izvadak = ws.cell(row=i, column=3).value
        dob = ws.cell(row=i, column=5).value
        vrsta = (ws.cell(row=i, column=8).value or "")
        trosak = _num(ws.cell(row=i, column=10).value)
        pdv = _num(ws.cell(row=i, column=11).value)
        ukupno = _num(ws.cell(row=i, column=12).value)
        iznos = ukupno if ukupno is not None else (
            round((trosak or 0) + (pdv or 0), 2) if trosak is not None else None)
        red = {
            "row": i, "ur": ws.cell(row=i, column=1).value, "racun": racun,
            "dob": dob or "", "dob_clean": ocisti(dob), "iznos": iznos,
            "datum": _datum(ws.cell(row=i, column=4).value),
            "ima_izvadak": izvadak not in (None, ""), "vrsta": str(vrsta).strip().lower(),
        }
        if racun not in (None, ""):
            racuni.append(red)
        if "dnevnic" in red["vrsta"] and not red["ima_izvadak"]:
            dnevnice.setdefault(red["dob_clean"], []).append(red)
    return {"racuni": racuni, "dnevnice": dnevnice}


def _vrsta_iz_opisa(opis, poziv=""):
    """Pokušaj odrediti VRSTU TROŠKA iz opisa (gdje ne može — None, korisnik upisuje)."""
    o = (opis or "").lower()
    if any(k in o for k in ("hzzo", "zdravstv")):
        return "zdravstveno osiguranje"
    if any(k in o for k in ("porez", "doprinos", "proračun", "proracun", "mirovinsk")):
        return "porez"
    if re.search(r"\b(pla[cć]a|pla[cć]e)\b", o):   # cijela riječ, ne "plaćanje"
        return "plaća"
    if "prijevoz" in o:
        return "prijevoz"
    if any(k in o for k in _GORIVO):
        return "gorivo"
    return None


def _dani_razlika(di, s):
    """Najmanja razlika u danima između datuma računa (di) i bilo kojeg datuma na izvodu
    (datum računa/valute ILI datum plaćanja)."""
    najmanje = 999
    for sd in (s.get("datum"), s.get("datum_placanja")):
        if di and sd:
            najmanje = min(najmanje, abs((di - sd).days))
    return najmanje


def _kandidati_racun(s, indeks, ime, tol_iznos=0.02, dani=30):
    """Računi (bez izvatka) s istim iznosom unutar ±dani (po bilo kojem datumu izvoda);
    rangirani po sličnosti imena pa datumu. 'ime' = naziv s kojim uspoređujemo (kod
    kartice trgovac iz opisa, inače naziv s izvoda)."""
    out = []
    naziv_clean = ocisti(ime)
    for r in indeks["racuni"]:
        if r["ima_izvadak"] or r["iznos"] is None:
            continue
        if abs(r["iznos"] - s["iznos"]) > tol_iznos:
            continue
        dd = _dani_razlika(r["datum"], s)
        if dd > dani:
            continue
        sim = fuzz.partial_ratio(naziv_clean, r["dob_clean"]) if naziv_clean and r["dob_clean"] else 0
        out.append({**r, "dd": dd, "sim": sim})
    out.sort(key=lambda r: (-r["sim"], r["dd"]))
    return out


def _broj_se_poklapa(s, kandidat):
    """Poklapa li se broj računa s onim na izvodu?
    Na izvodu često piše JEZGRA broja (npr. 'Plaćanje po računu 133627-1-1'),
    a Parrin broj je dulji ('133627-1-1_N2026030180'). Zato uspoređujemo:
      - cijeli očišćeni broj u opisu/pozivu, ILI
      - bilo koji distinktivni brojčani dio (≥4 znamenke) broja računa.
    """
    blob = ocisti(s["opis"]) + ocisti(s["poziv_platitelja"]) + ocisti(s["poziv_primatelja"])
    rc = ocisti(kandidat["racun"])
    if len(rc) >= 4 and rc in blob:
        return True
    for tok in re.findall(r"\d{4,}", str(kandidat["racun"] or "")):
        if tok in blob:
            return True
    return False


def broj_obroka(tekst):
    """Izvuci BROJ LEASING OBROKA iz teksta (XML računa ili opis izvoda).
    Hvata 'NN. leasing obrok' i 'leasing obrok NN' (npr. '56. leasing obrok' -> 56)."""
    t = (tekst or "").lower()
    m = re.search(r"(\d{1,3})\s*\.\s*leasing\s*obrok", t)
    if m:
        return int(m.group(1))
    m = re.search(r"leasing\s*obrok[\s:.\-]*(\d{1,3})", t)
    if m:
        return int(m.group(1))
    return None


def klasificiraj_stavku(s, indeks):
    """Vrati prijedlog: {tip, akcija, vrsta, kandidati, poruka}."""
    if s["smjer"] == "P":
        return {"tip": "uplata", "akcija": "preskoči", "poruka": "uplata nama (izlazni)"}

    opis = (s["opis"] or "")
    naziv = (s["naziv"] or "")
    poziv = f"{s['poziv_platitelja']} {s['poziv_primatelja']}".lower()
    blob = (opis + " " + naziv + " " + poziv).lower()

    # zdravstveno osiguranje (HZZO / zavod za zdravstveno) — prije plaće
    if any(k in blob for k in ("hzzo", "zdravstv", "zavod za zdrav")):
        return {"tip": "zdravstveno", "akcija": "novi_ur0", "vrsta": "zdravstveno osiguranje"}

    # porez / doprinosi / državni proračun (prije plaće — npr. "Placa ... GOVI" državi)
    if any(k in blob for k in ("državni proračun", "drzavni proracun", "porez", "doprinos",
                               "mirovinsk", " govi")):
        return {"tip": "porez", "akcija": "novi_ur0", "vrsta": "porez"}

    # D3 — plaća (CIJELA riječ — da "plaćanje"/"placanje" NE okine)
    if re.search(r"\b(pla[cć]a|pla[cć]e|isplata\s+pla[cć]e|neto\s+pla[cć]a)\b", opis.lower()):
        return {"tip": "plaća", "akcija": "novi_ur0", "vrsta": "plaća"}

    # prijevoz (naknada djelatniku — UR0 red, bez računa)
    if re.search(r"\bprijevoz", opis.lower()):
        return {"tip": "prijevoz", "akcija": "novi_ur0", "vrsta": "prijevoz"}

    # D2 — dnevnice (pn)
    if re.search(r"\bpn\b|pn\s*\d", blob):
        return _dnevnice(s, indeks)

    # LEASING — spoji po BROJU RAČUNA ili BROJU OBROKA s izvoda, NEOVISNO o iznosu.
    # (Broj/obrok su jednoznačni; iznos zna varirati — kamata/glavnica/opomena.)
    if "leasing" in (naziv + " " + opis).lower():
        naz_clean = ocisti(naziv)
        # 1) PRVO broj obroka s izvoda == broj obroka iz XML-a računa (najpouzdanije za leasing)
        ob = broj_obroka(opis)
        if ob is not None:
            for c in indeks["racuni"]:
                if c["ima_izvadak"] or c.get("obrok") != ob:
                    continue
                sim = fuzz.partial_ratio(naz_clean, c["dob_clean"]) if (naz_clean and c["dob_clean"]) else 0
                if sim >= SIM_IME:
                    return {"tip": "leasing", "akcija": "spoji", "kandidati": [{**c, "sim": sim, "dd": 0}],
                            "poruka": f"leasing: obrok {ob} (iz XML-a)"}
        # 2) inače broj računa s izvoda == račun u knjizi (koji još čeka)
        for c in indeks["racuni"]:
            if c["ima_izvadak"]:
                continue
            sim = fuzz.partial_ratio(naz_clean, c["dob_clean"]) if (naz_clean and c["dob_clean"]) else 0
            if sim >= SIM_IME and _broj_se_poklapa(s, c):
                return {"tip": "leasing", "akcija": "spoji", "kandidati": [{**c, "sim": sim, "dd": 0}],
                        "poruka": f"leasing: broj računa s izvoda ({sim}%)"}
        # ni obrok ni broj se ne poklope -> nastavi normalno (najčešće UR0)

    # Plaćanje računa: kandidati = računi koji ČEKAJU (iznos + datum). Ime za usporedbu:
    # kod kartice trgovac iz opisa (na izvodu je banka), inače naziv s izvoda.
    kartica = bool(re.search(r"\bkupovina\b", opis, re.I)) or bool(re.match(r"^\d{4}\D", opis))
    ime_match = trgovac_iz_opisa(opis) if kartica else naziv
    kand = _kandidati_racun(s, indeks, ime_match)
    lokacija = kartica or any(loc in opis.lower() for loc in _LOKACIJE)
    vrsta_hint = "gorivo" if je_gorivo(opis) else None

    # 1) E-RAČUN: na izvodu PIŠE broj računa -> spoji ako se poklapa broj + ime.
    #    ŠIRI prozor (120 dana): broj+ime+iznos su jaki signal, a računi (npr. leasing
    #    obrok) znaju se platiti i 30+ dana kasnije, na dospijeće.
    for c in _kandidati_racun(s, indeks, ime_match, dani=120):
        if _broj_se_poklapa(s, c) and c["sim"] >= SIM_IME:
            return {"tip": "račun", "akcija": "spoji", "kandidati": [c],
                    "poruka": f"broj + ime ({c['sim']}%, {c['dd']}d)"}
    # 2) BEZ broja na izvodu: ime + iznos. Ako je JEDINSTVEN račun tog imena koji čeka -> spoji
    #    (npr. Telemach: datum na izvodu je datum plaćanja, ne datum računa).
    imen = [c for c in kand if c["sim"] >= SIM_IME]
    if len(imen) == 1:
        return {"tip": "račun", "akcija": "spoji", "kandidati": [imen[0]],
                "poruka": f"ime + iznos, jedinstven ({imen[0]['sim']}%)"}
    # više kandidata istog imena+iznosa -> uzmi onaj kojem datum (prvi/drugi) odgovara
    for c in imen:
        if c["dd"] <= SPOJ_DANI:
            return {"tip": "račun", "akcija": "spoji", "kandidati": [c],
                    "poruka": f"ime + datum + iznos ({c['sim']}%, {c['dd']}d)"}
    # Inače -> UR 0000 redak; spojit će se naknadno kad stigne račun
    return {"tip": "kartica" if lokacija else "nepovezano", "akcija": "novi_ur0",
            "vrsta": vrsta_hint, "za_pregled": True,
            "poruka": "nema para — UR 0000, spaja se naknadno"}


def _dnevnice(s, indeks):
    """Pokušaj spojiti uplatu dnevnica: zbroj pn redova istog djelatnika = iznos?"""
    naziv_clean = ocisti(s["naziv"])
    najbolji_kljuc, najbolji_sim = None, 0
    for kljuc in indeks["dnevnice"]:
        sim = fuzz.partial_ratio(naziv_clean, kljuc) if naziv_clean and kljuc else 0
        if sim > najbolji_sim:
            najbolji_sim, najbolji_kljuc = sim, kljuc
    if not najbolji_kljuc or najbolji_sim < 80:
        return {"tip": "dnevnice", "akcija": "ur0_crveno", "poruka": "djelatnik nije nađen među pn redovima"}
    redovi = indeks["dnevnice"][najbolji_kljuc]
    zbroj = round(sum(r["iznos"] for r in redovi if r["iznos"]), 2)
    if abs(zbroj - s["iznos"]) <= 0.02:
        return {"tip": "dnevnice", "akcija": "dnevnice_spoji", "kandidati": redovi,
                "poruka": f"{len(redovi)} pn redova ({najbolji_kljuc}) = {zbroj} €"}
    return {"tip": "dnevnice", "akcija": "pitaj", "kandidati": redovi,
            "poruka": f"zbroj pn ({zbroj} €) ≠ uplata ({s['iznos']} €) — provjeri"}
