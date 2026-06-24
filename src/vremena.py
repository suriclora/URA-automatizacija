# -*- coding: utf-8 -*-
"""
Određivanje vremena polaska/povratka za putni nalog (za izračun dnevnica u Excelu).

Pravila (korisnik):
  • Teren S ENC prolazom: polazak = 30 min prije prvog prolaska; smije se produžiti do
    MAX 1 h prije prvog i 1 h poslije zadnjeg prolaska — da se dosegne 8 h (0,5 dnevnice)
    ili 12 h (1 dnevnica).
  • Teren BEZ ENC prolaza: bili 12+ h na gradilištu → puna dnevnica. Vremena VARIRANA
    (ne svaki isti), polazak od 06:00 naviše, minute zaokružene na 5, trajanje > 12 h.

ENC CSV (HAC, ';' razdjelnik, utf-8-sig): stupci [1]=relacija, [3]=vrijeme ulaska,
[4]=vrijeme izlaska, [5]=naziv uređaja (nadimak vozila npr. 'jumpy','ford','opel').
"""
import csv
import hashlib
from datetime import datetime, timedelta, time


def _voz(s):
    """Normaliziraj naziv vozila na prvu riječ malim slovima ('Opel 2'->'opel')."""
    return (str(s or "").strip().lower().split() or [""])[0]


def _dt(s):
    try:
        return datetime.strptime(str(s).strip(), "%d.%m.%Y %H:%M:%S")
    except (ValueError, TypeError):
        return None


def enc_grupe(csv_path):
    """Vrati {(vozilo_norm, date): (najraniji_ulazak, najkasniji_izlazak)} iz ENC CSV-a."""
    grupe = {}
    try:
        raw = open(csv_path, encoding="utf-8-sig").read()
    except Exception:
        return grupe
    for r in list(csv.reader(raw.splitlines(), delimiter=";"))[1:]:
        if len(r) < 6:
            continue
        voz = _voz(r[5])
        ul, iz = _dt(r[3]), _dt(r[4])
        if not voz or not ul:
            continue
        iz = iz or ul
        k = (voz, ul.date())
        if k not in grupe:
            grupe[k] = [ul, iz]
        else:
            grupe[k][0] = min(grupe[k][0], ul)
            grupe[k][1] = max(grupe[k][1], iz)
    return {k: (v[0], v[1]) for k, v in grupe.items()}


def _broj_hr(s):
    """'3,00' / '25,30' -> float; None ako ne valja."""
    try:
        return round(float(str(s).replace(".", "").replace(",", ".")), 2)
    except (ValueError, TypeError, AttributeError):
        return None


def enc_prolasci(csv_path):
    """Svi pojedinačni ENC prolasci: lista {voz, datum(date), iznos(float), relacija}.
    Stupci: [1]relacija [3]ulazak [5]naziv uređaja(vozilo) [8]isplata(iznos)."""
    out = []
    try:
        raw = open(csv_path, encoding="utf-8-sig").read()
    except Exception:
        return out
    for r in list(csv.reader(raw.splitlines(), delimiter=";"))[1:]:
        if len(r) < 9:
            continue
        voz = _voz(r[5])
        ul = _dt(r[3])
        if not voz or not ul:
            continue
        out.append({"voz": voz, "datum": ul.date(), "iznos": _broj_hr(r[8]),
                    "relacija": (r[1] or "").strip()})
    return out


def prolasci_za_teren(enc_prolasci_lista, enc_nick, d1, d2):
    """Vrati ENC prolaske (datum, iznos) za vozilo (ENC nadimak) unutar [d1, d2]."""
    voz = _voz(enc_nick)
    return [p for p in enc_prolasci_lista if p["voz"] == voz and d1 <= p["datum"] <= d2]


def _round5(dtobj):
    """Zaokruži vrijeme na 5 minuta."""
    r = round(dtobj.minute / 5) * 5
    if r == 60:
        return (dtobj.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1))
    return dtobj.replace(minute=r, second=0, microsecond=0)


def _dnevnice(sati):
    if sati > 12:
        return 1.0
    if sati > 7.9:
        return 0.5
    return 0.0


def _enc_prozor(voz, d1, d2, enc):
    """Najraniji ulazak / najkasniji izlazak za vozilo unutar [d1, d2]; None ako nema."""
    voz = _voz(voz)
    A = B = None
    dan = d1
    while dan <= d2:
        if (voz, dan) in enc:
            ul, iz = enc[(voz, dan)]
            A = ul if A is None else min(A, ul)
            B = iz if B is None else max(B, iz)
        dan += timedelta(days=1)
    return A, B


def vremena(djelatnik, vozilo, polazak, povratak, enc):
    """Vrati (odlazak_datetime, povratak_datetime) za teren.
    Ako ima ENC prolaza -> po prolascima (+produženje za prag dnevnice); inače -> standardno 12h+."""
    seed = int(hashlib.md5(f"{djelatnik}|{polazak}".encode("utf-8")).hexdigest(), 16)
    A, B = _enc_prozor(vozilo, polazak, povratak, enc)
    if A and B:
        # produži polazak/povratak (do ±1h) da se dosegne 8h (0,5) ili 12h (1) dnevnica,
        # ali NE ravno na pragu — dodaj 5/10/15 min ako ima mjesta (da ne izgleda sumnjivo)
        raspon = (B - A).total_seconds() / 60.0     # minute od prvog ulaska do zadnjeg izlaska
        margin = 5 + 5 * (seed % 3)                  # 5 / 10 / 15 min varijacije
        maxdur = raspon + 120                        # najviše dostižno (±1h)
        if maxdur > 720:
            target = min(maxdur, 725 + margin)       # 12h+ -> 1 dnevnica (+margin, ne ravno na 12h)
        elif maxdur > 474:
            target = min(maxdur, 485 + margin)       # 8h+ -> 0,5 dnevnica (ako ima mjesta, ne ravno na 8h)
        else:
            target = raspon + 60                     # prag nedostižan -> osnova (0 dnevnica)
        total = min(120.0, max(60.0, target - raspon))
        before = min(60, max(30, int(round(total / 2))))
        after = int(min(60, max(30, round(total) - before)))
        return _round5(A - timedelta(minutes=before)), _round5(B + timedelta(minutes=after))

    # BEZ ENC-a: puna dnevnica (>12h), varirana vremena (deterministički po djelatniku+danu)
    dep_min = 5 * (seed % 12)                       # 0..55, korak 5  -> polazak 06:00..06:55
    dep = datetime.combine(polazak, time(6, 0)) + timedelta(minutes=dep_min)
    if polazak == povratak:
        traj = 12 * 60 + 5 + 5 * ((seed // 13) % 11)    # 12h05m .. 12h55m (>12h -> 1 dnevnica, max 13h)
        ret = dep + timedelta(minutes=traj)
    else:
        ret_min = 5 * ((seed // 7) % 24)            # povratak 18:00..19:55
        ret = datetime.combine(povratak, time(18, 0)) + timedelta(minutes=ret_min)
    return _round5(dep), _round5(ret)
