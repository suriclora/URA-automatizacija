# -*- coding: utf-8 -*-
"""
Izvještaji (samo čitanje knjige, ništa se ne mijenja):

  racuni_bez_izvoda  — upisani RAČUNI koji NISU plaćeni (nema izvatka/plaćeno)
                       => "računima fali izvod"
  izvodi_bez_racuna  — UPLATE s izvoda (UR 0000) koje čekaju račun, BEZ onih koji
                       uvijek ostaju 0000 (plaća/porez/dnevnice/doprinosi/krediti/naknade…)
                       => "izvodima fale računi"
"""

# Vrste troška koje UVIJEK ostaju UR 0000 (nema vanjskog računa) — preskaču se u
# izvještaju "izvodi kojima fale računi". Podudaranje po ključnoj riječi (sadrži).
TRAJNE_UR0 = [
    "plaća", "placa", "placa e", "prijevoz", "dnevnice", "uskrsnjica", "uskrsnjica",
    "bolovanje", "porez", "doprinos", "mirovinsko", "zdravstveno", "kredit",
    "holding", "sudska", "naknada",  # bankovne naknade (za izvod/vođenje/dom pp)
]


def _vrsta_je_trajna(vrsta):
    v = (vrsta or "").strip().lower()
    return any(k in v for k in TRAJNE_UR0)


def _ur_broj(v):
    try:
        return int(str(v).split(".")[0])
    except (ValueError, TypeError):
        return None


def racuni_bez_izvoda(ws, header_red, od_datuma=None):
    """RAČUNI (UR>0, ima broj računa) koji NEMAJU izvadak ni plaćeno => fali im izvod.
    'ws' treba biti data_only (zbog UKUPNO formule). Vrati listu rječnika."""
    out = []
    for i in range(header_red + 1, ws.max_row + 1):
        ur = _ur_broj(ws.cell(row=i, column=1).value)
        racun = ws.cell(row=i, column=2).value
        if not ur or ur <= 0 or racun in (None, ""):
            continue
        izvadak = ws.cell(row=i, column=3).value
        placeno = ws.cell(row=i, column=16).value
        if izvadak not in (None, "") or placeno not in (None, ""):
            continue  # plaćeno/povezano — preskoči
        datum = ws.cell(row=i, column=4).value
        if od_datuma and datum and hasattr(datum, "date") and datum.date() < od_datuma:
            continue
        # UKUPNO je formula (=K+J) pa je u data_only None -> računamo TROŠAK+PDV
        neto = ws.cell(row=i, column=10).value or 0
        pdv = ws.cell(row=i, column=11).value or 0
        try:
            ukupno = round(float(neto) + float(pdv), 2)
        except (TypeError, ValueError):
            ukupno = None
        out.append({
            "red": i, "ur": ur, "racun": racun,
            "dobavljac": ws.cell(row=i, column=5).value,
            "datum": datum, "ukupno": ukupno,
        })
    return out


def brojke(ws, header_red, od_potvrda=None):
    """Brojevi za pločice na vrhu aplikacije (čita data_only knjigu):
      racuni        — ukupno upisanih računa (UR>0 + broj računa), CIJELA knjiga
      ceka_izvod    — računi koji nisu plaćeni (fali izvod), CIJELA knjiga
      treba_potvrdu — retci (OD 'od_potvrda', npr. datum starta) bez VRSTE (skripta nije znala)
      za_rucno      — uplate s izvoda bez računa (bez trajnih 0000), CIJELA knjiga
    """
    racuni = 0
    treba_potvrdu = 0
    for i in range(header_red + 1, ws.max_row + 1):
        ur = _ur_broj(ws.cell(row=i, column=1).value)
        racun = ws.cell(row=i, column=2).value
        if ur and ur > 0 and racun not in (None, ""):
            racuni += 1
            if ws.cell(row=i, column=8).value in (None, ""):   # nema vrste
                datum = ws.cell(row=i, column=4).value
                if not od_potvrda or (hasattr(datum, "date") and datum.date() >= od_potvrda):
                    treba_potvrdu += 1
    return {
        "racuni": racuni,
        "ceka_izvod": len(racuni_bez_izvoda(ws, header_red, None)),   # cijela knjiga
        "treba_potvrdu": treba_potvrdu,
        "za_rucno": len(izvodi_bez_racuna(ws, header_red, None)),     # cijela knjiga
    }


def racuni_bez_vrste(ws, header_red, od_datuma=None):
    """RAČUNI (UR>0, ima broj) kojima FALI VRSTA TROŠKA (skripta nije bila sigurna) =
    'treba potvrdu'. Isti kriterij kao u brojke(). Vrati listu rječnika."""
    out = []
    for i in range(header_red + 1, ws.max_row + 1):
        ur = _ur_broj(ws.cell(row=i, column=1).value)
        racun = ws.cell(row=i, column=2).value
        if not ur or ur <= 0 or racun in (None, ""):
            continue
        if ws.cell(row=i, column=8).value not in (None, ""):   # ima vrstu — u redu
            continue
        datum = ws.cell(row=i, column=4).value
        if od_datuma and not (hasattr(datum, "date") and datum.date() >= od_datuma):
            continue
        neto = ws.cell(row=i, column=10).value or 0
        pdv = ws.cell(row=i, column=11).value or 0
        try:
            ukupno = round(float(neto) + float(pdv), 2)
        except (TypeError, ValueError):
            ukupno = None
        out.append({"red": i, "ur": ur, "racun": racun,
                    "dobavljac": ws.cell(row=i, column=5).value,
                    "datum": datum, "ukupno": ukupno})
    return out


def izvodi_bez_racuna(ws, header_red, od_datuma=None):
    """UPLATE s izvoda (UR 0000, ima izvadak + plaćeno, NEMA broj računa) koje čekaju
    račun — BEZ trajnih (plaća/porez/dnevnice/…). Vrati listu rječnika."""
    out = []
    for i in range(header_red + 1, ws.max_row + 1):
        ur = _ur_broj(ws.cell(row=i, column=1).value)
        if ur != 0:
            continue
        if ws.cell(row=i, column=2).value not in (None, ""):   # ima broj računa
            continue
        izvadak = ws.cell(row=i, column=3).value
        placeno = ws.cell(row=i, column=16).value
        if izvadak in (None, "") or placeno in (None, ""):     # mora biti s izvoda i plaćeno
            continue
        vrsta = ws.cell(row=i, column=8).value
        if _vrsta_je_trajna(vrsta):                            # uvijek ostaje 0000 — preskoči
            continue
        datum = ws.cell(row=i, column=4).value
        if od_datuma and datum and hasattr(datum, "date") and datum.date() < od_datuma:
            continue
        out.append({
            "red": i, "izvadak": izvadak,
            "dobavljac": ws.cell(row=i, column=5).value,
            "datum": datum, "placeno": placeno, "vrsta": vrsta,
        })
    return out
