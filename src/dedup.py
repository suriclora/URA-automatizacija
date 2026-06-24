# -*- coding: utf-8 -*-
"""
Pametno prepoznavanje duplikata za e-račune s Parre.

Zašto ovaj modul postoji:
  Parra daje PUNI broj računa, a u Excelu su brojevi upisani RUČNO i često
  SKRAĆENO (npr. '751' umjesto '751/1/1'). Ni imena dobavljača se ne poklapaju
  (Excel: 'hrt', Parra: 'HRVATSKA RADIOTELEVIZIJA'). Zato se duplikat ne može
  pouzdano prepoznati po broju ni po imenu.

Rješenje: usporedba po IZNOSU (bruto = TROŠAK + PDV) i DATUMU (s tolerancijom),
jer se ti podaci poklapaju i kad je broj skraćen.
"""
from datetime import date, datetime
from src.utils import ocisti


def _num(v):
    """Pretvori u float ili vrati None (npr. ako je ćelija formula ili tekst)."""
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _datum_celije(v):
    """openpyxl datum -> date; inače None."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def izracunaj_iznos(ukupno, trosak, pdv):
    """Bruto iznos: prvo iz UKUPNO (izračunata vrijednost), pa rezerva TROŠAK+PDV.
    UKUPNO i PDV su u Excelu često formule — zato ovaj modul čita data_only.
    """
    u = _num(ukupno)
    if u is not None:
        return round(u, 2)
    t = _num(trosak)
    if t is not None:
        return round(t + (_num(pdv) or 0.0), 2)
    return None


def izgradi_indeks(ws, header_red):
    """Napravi popis postojećih redova: iznos, datum, UR, račun, dobavljač.
    'ws' MORA biti učitan s data_only=True (zbog formula UKUPNO/PDV).
    Koristi iter_rows jer je workbook u read_only načinu.
    Kolone (0-indeksirano u tuple): 0=UR, 1=RAČUN, 3=DATUM, 4=DOBAVLJAČ,
    9=TROŠAK, 10=PDV, 11=UKUPNO.
    """
    indeks = []
    for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if idx <= header_red:
            continue

        def g(n):
            return row[n] if len(row) > n else None

        iznos = izracunaj_iznos(g(11), g(9), g(10))
        if iznos is None:
            continue
        indeks.append({
            "iznos": iznos,
            "datum": _datum_celije(g(3)),
            "ur": g(0),
            "racun": g(1),
            "dob": ocisti(g(4)),
        })
    return indeks


def _parra_datum(r):
    ds = (r.get("documentDateInCet") or "")[:10]
    try:
        return datetime.strptime(ds, "%Y-%m-%d").date() if ds else None
    except ValueError:
        return None


def klasificiraj_racun(r, indeks, tol_iznos=0.02, dani_dupli=7, dani_neizvjesno=45):
    """Razvrstaj jedan Parrin račun:
      'duplikat'   - isti iznos i datum unutar ±dani_dupli (gotovo sigurno već upisan)
      'neizvjesno' - isti iznos, ali datum 8..dani_neizvjesno dana (možda mjesečni, pitati)
      'novo'       - nema sličnog iznosa (ili je vremenski jako daleko)
    Vraća (klasa, najblizi_kandidat_ili_None).
    """
    iznos = round(float(r.get("totalAmount", 0) or 0), 2)
    pdat = _parra_datum(r)

    kandidati = [k for k in indeks if abs(k["iznos"] - iznos) <= tol_iznos]
    if not kandidati:
        return "novo", None

    # nađi vremenski najbliži kandidat
    najblizi, najmanje = None, None
    for k in kandidati:
        dd = abs((k["datum"] - pdat).days) if (pdat and k["datum"]) else 999
        if najmanje is None or dd < najmanje:
            najmanje, najblizi = dd, k

    if najmanje <= dani_dupli:
        return "duplikat", najblizi
    if najmanje <= dani_neizvjesno:
        return "neizvjesno", najblizi
    return "novo", najblizi


def klasificiraj_sve(parra_racuni, indeks, **kwargs):
    """Razvrstaj sve račune u tri liste. Vraća dict s 'novo'/'neizvjesno'/'duplikat'."""
    klase = {"novo": [], "neizvjesno": [], "duplikat": []}
    for r in parra_racuni:
        klasa, kand = klasificiraj_racun(r, indeks, **kwargs)
        klase[klasa].append((r, kand))
    return klase
