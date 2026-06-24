# -*- coding: utf-8 -*-
"""
FAZA 4 — obrada bankovnih izvoda: parsiranje, spajanje s računima i upis.

Tijek po izvodu (samo oni koji NISU već u knjizi — provjera po broju+banci):
  - preimenuj datoteku u "Izvod XX.pdf"
  - za svaku stavku (Duguje/isplata; uplate se preskaču): klasificiraj pa
      * spoji      -> na red računa upiši IZVADAK(C) + datum plać.(O) + PLAĆENO(P) + SREDSTVO(R)
      * dnevnice   -> isto na pn redove djelatnika
      * novi UR0   -> novi red (UR 0000) s datumom, dobavljačem, iznosom, izvadkom, vrstom, sredstvom

Upis radi tek kad 'pisi=True' (i nakon backupa). Inače samo prijavi što bi.
"""
import os
import re
import glob
from pathlib import Path

from src import izvodi, spajanje
from src import excel as excel_mod
from src import vrsta as vrsta_mod
from src.utils import ocisti

# Kolone (1-indeksirano)
UR, RACUN, IZVADAK, DATUM, DOBAVLJAC, VRSTA = 1, 2, 3, 4, 5, 8
TROSAK, PDV, UKUPNO, DAT_PLAC, PLACENO, DUGOVANJE, SREDSTVO = 10, 11, 12, 15, 16, 17, 18


def _bank_iz_sredstva(sr):
    s = (sr or "").lower()
    if s.startswith("hpb"):
        return "HPB"
    if s in ("virman", "debitna"):
        return "RBA"
    return None


def uneseni_izvodi(ws_val, header_red):
    """Skup (broj, banka) izvoda koji su VEĆ u knjizi (da se ne obrađuju ponovno)."""
    out = set()
    for i in range(header_red + 1, ws_val.max_row + 1):
        izv = ws_val.cell(row=i, column=IZVADAK).value
        if not isinstance(izv, (int, float)):
            continue
        bank = _bank_iz_sredstva(ws_val.cell(row=i, column=SREDSTVO).value)
        if bank:
            out.add((int(izv), bank))
    return out


def _je_kartica(s):
    """Kartično plaćanje: RBA opis sadrži 'KUPOVINA', HPB počinje šifrom (4 znamenke + trgovac)."""
    o = (s["opis"] or "")
    return bool(re.search(r"\bkupovina\b", o, re.I)) or bool(re.match(r"^\d{4}\D", o))


def _sredstvo(banka, kartica):
    if banka == "HPB":
        return "hpb kartica" if kartica else "hpb virman"
    return "debitna" if kartica else "virman"


def _spoji(ws, red, broj_izvoda, datum_pl, iznos, sredstvo, izvod_putanja=None):
    """Na POSTOJEĆI red računa upiši podatke o plaćanju.
    Prvo osiguraj ISPRAVAN stil ćelija (kod ručnih neplaćenih redova prazne ćelije
    nemaju stil), pa upiši. IZVADAK postaje hiperveza na PDF izvoda (plavo, bez podcrtavanja).
    """
    for col in (DAT_PLAC, PLACENO, SREDSTVO, IZVADAK):
        excel_mod.stil_celije(ws, red, col)
    ws.cell(row=red, column=DAT_PLAC).value = datum_pl
    ws.cell(row=red, column=PLACENO).value = iznos
    if not ws.cell(row=red, column=SREDSTVO).value:
        ws.cell(row=red, column=SREDSTVO).value = sredstvo
    ws.cell(row=red, column=IZVADAK).value = broj_izvoda
    # DUGOVANJE: ako ručni red nema formulu, postavi je (=L-P), uz ispravan stil
    dug = ws.cell(row=red, column=DUGOVANJE)
    if not (isinstance(dug.value, str) and dug.value.startswith("=")):
        excel_mod.stil_celije(ws, red, DUGOVANJE)
        dug.value = f"=L{red}-P{red}"
    if izvod_putanja:
        excel_mod.postavi_hyperlink(ws, red, IZVADAK, izvod_putanja)


def _novi_ur0(ws, red, stil_red, s, broj_izvoda, sredstvo, vrsta, dobavljac, izvod_putanja=None):
    """Novi UR 0000 red za stavku s izvoda (bez računa). Stil cijelog reda kopiran iz
    reda iznad (identičan izgled), UKUPNO/DUGOVANJE kao formule, IZVADAK kao hiperveza.
    DATUM (D) = datum računa (valuta); DAT_PLAC (O) = datum izvršenja (plaćanja)."""
    podaci = {
        UR: 0, IZVADAK: broj_izvoda, DATUM: s["datum"],
        DOBAVLJAC: (dobavljac or "").lower(), VRSTA: vrsta,
        TROSAK: s["iznos"], PDV: 0,
        DAT_PLAC: s.get("datum_placanja") or s["datum"],
        PLACENO: s["iznos"], SREDSTVO: sredstvo,
    }
    excel_mod.upisi_vrijednosti(ws, red, podaci, stil_red)  # stil + UKUPNO/DUGOVANJE formule
    if izvod_putanja:
        excel_mod.postavi_hyperlink(ws, red, IZVADAK, izvod_putanja)


def _broj_iz_imena(ime):
    m = re.search(r"Izvod\s+(\d+)", ime, re.I)
    return int(m.group(1)) if m else None


def _datum_izvoda(izv):
    """Datum izvoda iz stringa 'dd.mm.yyyy' (ili None)."""
    from datetime import datetime
    try:
        return datetime.strptime(izv.get("datum") or "", "%d.%m.%Y").date()
    except ValueError:
        return None


def obradi_izvode(ws, indeks, uneseni, folderi, logger, stat, pisi, premjesti, zadnji_red,
                  history=None, obradjuj_od=None, imena=None, stanje=None):
    """Obradi sve nove izvode iz zadanih mapa. Vrati (promijenjeno, novi_zadnji_red)."""
    red_pisi = zadnji_red + 1
    promijenjeno = False

    for banka, folder in folderi:
        for f in sorted(glob.glob(os.path.join(folder, "**", "*.pdf"), recursive=True)):
            ime = os.path.basename(f)
            broj = _broj_iz_imena(ime)
            # ako je već imenovan i u knjizi -> preskoči bez parsiranja
            if broj is not None and (broj, banka) in uneseni:
                continue
            try:
                izv = izvodi.parsiraj_izvod(f)
            except Exception as e:
                logger.error("   ⚠️ Ne mogu pročitati %s: %s", ime, e)
                continue
            if not izv or izv["banka"] != banka or not izv["broj"]:
                continue
            broj = int(izv["broj"])
            if (broj, banka) in uneseni:
                continue
            # DATUM STARTA: preskoči izvode prije starta (ručni zaostatak)
            di = _datum_izvoda(izv)
            if obradjuj_od and di and di < obradjuj_od:
                continue
            # VEĆ OBRAĐEN (zapamćen): ne obrađuj ponovno (čak i ako su redovi obrisani)
            kljuc = f"{banka}|{broj}"
            if stanje is not None and kljuc in stanje["izvodi"]:
                continue

            logger.info("--- %s Izvod %s (%s) — %s stavki ---", banka, broj, izv["datum"], len(izv["stavke"]))
            stat["izvoda_obradeno"] += 1

            # preimenovanje u "Izvod XX.pdf"
            zeljeno = f"Izvod {broj}.pdf"
            izvod_putanja = os.path.join(os.path.dirname(f), zeljeno)  # za hipervezu (IZVADAK)
            if ime != zeljeno:
                if pisi and premjesti:
                    try:
                        os.rename(f, os.path.join(os.path.dirname(f), zeljeno))
                        logger.info("   📄 preimenovano: %s → %s", ime, zeljeno)
                    except OSError as e:
                        logger.warning("   ne mogu preimenovati %s: %s", ime, e)
                else:
                    logger.info("   [prikaz] BIH preimenovao %s → %s", ime, zeljeno)

            for s in izv["stavke"]:
                k = spajanje.klasificiraj_stavku(s, indeks)
                akc = k["akcija"]
                kartica = _je_kartica(s)
                sred = _sredstvo(banka, kartica)
                datum_pl = s.get("datum_placanja") or s["datum"]   # stupac O (izvršenje)
                # DOBAVLJAČ: kod kartice trgovac iz opisa (TIFON/INA...), inače naziv s izvoda;
                # pa kanonsko ime iz knjige (jedan naziv, bez adrese — samo ime i prezime)
                trgovac = spajanje.trgovac_iz_opisa(s["opis"]) if kartica else None
                dobavljac = vrsta_mod.kanonsko_ime(trgovac or s["naziv"], imena or {})

                if akc == "preskoči":
                    stat["preskoceno_uplata"] += 1
                    continue

                if akc == "spoji":
                    c = k["kandidati"][0]
                    logger.info("   ✅ spoj [%s]: %s %.2f€ → UR%s %s",
                                k.get("poruka", ""), s["naziv"][:18], s["iznos"], c["ur"], c["dob"])
                    if pisi:
                        _spoji(ws, c["row"], broj, datum_pl, s["iznos"], sred, izvod_putanja)
                        c["ima_izvadak"] = True  # da se ne spoji dvaput
                        promijenjeno = True
                    stat["spojeno"] += 1

                elif akc == "dnevnice_spoji":
                    logger.info("   ✅ dnevnice: %s → %s pn redova", s["naziv"][:20], len(k["kandidati"]))
                    if pisi:
                        for r in k["kandidati"]:
                            _spoji(ws, r["row"], broj, datum_pl, r["iznos"], sred, izvod_putanja)
                            r["ima_izvadak"] = True
                        promijenjeno = True
                    stat["dnevnice"] += 1

                else:  # novi_ur0 (plaća/porez/prijevoz/kartica/nepovezano/dnevnice-bez-para)
                    vrsta = k.get("vrsta")
                    if not vrsta and history is not None:  # probaj iz povijesti dobavljača
                        vrsta = vrsta_mod.predlozi_vrstu(dobavljac, history, s["opis"] or "")
                    logger.info("   + UR0: %s %.2f€ (%s)", (dobavljac or "")[:20], s["iznos"], vrsta or "—")
                    if pisi:
                        _novi_ur0(ws, red_pisi, red_pisi - 1, s, broj, sred, vrsta, dobavljac, izvod_putanja)
                        promijenjeno = True
                        red_pisi += 1
                    stat["ur0_redova"] += 1

            if pisi and stanje is not None:
                stanje["izvodi"].add(kljuc)   # zapamti da je ovaj izvod obrađen

    return promijenjeno, red_pisi - 1
