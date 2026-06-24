# -*- coding: utf-8 -*-
"""
Parsiranje bankovnih izvoda (HPB i RBA) iz PDF-a.

Vraća, za svaki izvod: broj izvoda, datum, banka, i listu stavki.
Svaka stavka: red_br, iznos, smjer ('D'=duguje/isplata, 'P'=potražuje/uplata),
datum, naziv (platitelj/primatelj), racun (IBAN), opis, poziv_platitelja,
poziv_primatelja.

HPB: 5-redni format; smjer (Duguje/Potražuje) određujemo po VODORAVNOM položaju
iznosa na stranici (tekst ga ne piše riječima).
RBA: jednostavniji; smjer (D/P) piše direktno u retku stavke.
"""
import re
from datetime import date
import pdfplumber

import config

_IZNOS = r"\d{1,3}(?:\.\d{3})*,\d{2}"   # 2.365,61 ili 47,78


def _d(s):
    """'08.01.2026' -> date(2026,1,8); inače None."""
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", s or "")
    return date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


def _f(s):
    """'2.365,61' -> 2365.61"""
    return float(s.replace(".", "").replace(",", "."))


def detektiraj_banku(tekst):
    t = (tekst or "").upper()
    if config.MOJ_OIB in t or "IZVADAK O PROMJENAMA" in t or "HPBZHR2X" in t:
        return "HPB"
    if "BROJ IZVATKA" in t or "RZBHHR2X" in t or "IB IZVOD" in t or "RAIFFEISEN" in t:
        return "RBA"
    return None


def _grupiraj_linije(page, tol=3):
    """Grupiraj riječi stranice u linije (po y koordinati). Vrati listu:
    {'text': spojeni tekst, 'words': [riječi sortirane po x]}.
    """
    rijeci = sorted(page.extract_words(), key=lambda w: (w["top"], w["x0"]))
    linije = []
    trenutna = []
    zadnji_top = None
    for w in rijeci:
        if zadnji_top is None or abs(w["top"] - zadnji_top) <= tol:
            trenutna.append(w)
            zadnji_top = w["top"] if zadnji_top is None else zadnji_top
        else:
            linije.append(trenutna)
            trenutna = [w]
            zadnji_top = w["top"]
    if trenutna:
        linije.append(trenutna)
    out = []
    for ws in linije:
        ws = sorted(ws, key=lambda x: x["x0"])
        out.append({"text": " ".join(x["text"] for x in ws), "words": ws})
    return out


# ---------------------------------------------------------------------
#  HPB
# ---------------------------------------------------------------------

def parsiraj_hpb(pdf):
    broj = datum = None
    prag = None  # granica x između Duguje i Potražuje kolone
    stavke = []

    for page in pdf.pages:
        linije = _grupiraj_linije(page)

        for ln in linije:
            m = re.search(r"Izvadak broj/Datum:\s*(\d+)\s*/\s*(\d{2}\.\d{2}\.\d{4})", ln["text"])
            if m and broj is None:
                broj, datum = m.group(1), m.group(2)
            if prag is None:
                xi = xp = None
                for w in ln["words"]:
                    if w["text"].startswith("Isplata"):
                        xi = w["x1"]
                    if w["text"].startswith("Potražuje"):
                        xp = w["x0"]
                if xi and xp:
                    prag = (xi + xp) / 2

        for i, ln in enumerate(linije):
            # Redak stavke: "1 HR67 <poziv?> 2.365,61"
            m = re.match(rf"^(\d+)\s+(HR\d{{2}})\s+(?:(.*?)\s+)?({_IZNOS})$", ln["text"])
            if not m:
                continue
            redbr, model_pl, poziv_pl, iznos_s = m.groups()

            # položaj iznosa -> kolona (Duguje/Potražuje)
            amw = [w for w in ln["words"] if re.fullmatch(_IZNOS, w["text"])]
            x = amw[-1]["x0"] if amw else 0
            smjer = "D" if (prag is None or x < prag) else "P"

            # natrag: redak s IBAN-om (HR + 19 znamenki) + naziv; opis između
            naziv = racun = opis = ""
            for j in range(i - 1, max(i - 5, -1), -1):
                ma = re.match(r"^(HR\d{19})\s+(.*)$", linije[j]["text"])
                if ma:
                    racun, naziv = ma.group(1), ma.group(2)
                    opis = " ".join(linije[k]["text"] for k in range(j + 1, i))
                    break

            # poziv primatelja + datumi: redak ispod = datum VALUTE (datum računa),
            # redak ispod toga = datum IZVRŠENJA (datum plaćanja).
            poziv_pr = ""
            dat_valute = _d(datum)        # fallback: datum izvoda
            if i + 1 < len(linije):
                mp = re.match(r"^(\d{2}\.\d{2}\.\d{4})\.?\s+(HR\d{2})\s*(.*?)(?:\s+[\d.,]+)?$",
                              linije[i + 1]["text"])
                if mp:
                    dat_valute = _d(mp.group(1)) or dat_valute
                    poziv_pr = f"{mp.group(2)} {mp.group(3)}".strip()
            dat_izvrsenja = dat_valute
            if i + 2 < len(linije):
                md2 = re.match(r"^(\d{2}\.\d{2}\.\d{4})", linije[i + 2]["text"])
                if md2:
                    dat_izvrsenja = _d(md2.group(1)) or dat_izvrsenja

            stavke.append({
                "red_br": int(redbr), "iznos": _f(iznos_s), "smjer": smjer,
                "datum": dat_valute, "datum_placanja": dat_izvrsenja,
                "naziv": naziv.strip(), "racun": racun, "opis": opis.strip(),
                "poziv_platitelja": f"{model_pl} {poziv_pl or ''}".strip(),
                "poziv_primatelja": poziv_pr,
            })

    return {"banka": "HPB", "broj": broj, "datum": datum, "stavke": stavke}


# ---------------------------------------------------------------------
#  RBA
# ---------------------------------------------------------------------

def parsiraj_rba(pdf):
    broj = datum = None
    linije = []
    for page in pdf.pages:
        t = page.extract_text() or ""
        linije.extend(t.split("\n"))

    for ln in linije:
        m = re.search(r"Broj izvatka:\s*(\d+)", ln)
        if m and broj is None:
            broj = m.group(1)
        m = re.search(r"Datum:\s*(\d{2}\.\d{2}\.\d{4})", ln)
        if m and datum is None:
            datum = m.group(1)

    # Redak stavke: bilo kakve reference, pa "datum D/P iznos" na kraju.
    trans_re = re.compile(rf"^(.+?)\s+(\d{{2}}\.\d{{2}}\.\d{{4}})\s+([DP])\s+({_IZNOS})$")
    stop_re = re.compile(r"(Proknjiženo stanje|Ukupni promet|Rezervacije|Raspoloživo|"
                         r"SEKTOR|IZVADAK O|Početno stanje|broj naloga|Prijenos)")
    stavke = []
    i, n = 0, len(linije)
    while i < n:
        m = trans_re.match(linije[i].strip())
        if not m:
            i += 1
            continue
        _refs, dat, dp, izn = m.groups()

        # blok do iduće stavke ili kraja
        blok = []
        j = i + 1
        while j < n:
            s = linije[j].strip()
            if trans_re.match(s) or stop_re.search(s):
                break
            if s:
                blok.append(s)
            j += 1

        opis = racun = naziv = poziv_pl = poziv_pr = ""
        dat_valute = None
        if blok:
            mval = re.search(r"(\d{2}\.\d{2}\.\d{4})\.?\s*$", blok[0])  # datum valute (datum računa)
            if mval:
                dat_valute = _d(mval.group(1))
            opis = re.sub(r"\s*\d{2}\.\d{2}\.\d{4}\.?$", "", blok[0]).strip()
        for k, b in enumerate(blok):
            mr = re.match(r"^(HR\d{19,21})(?:\s+(.*))?$", b)
            if mr and not racun:
                racun = mr.group(1)
                poziv_pl = (mr.group(2) or "").strip()
                # naziv + poziv primatelja je obično sljedeći redak
                if k + 1 < len(blok):
                    mn = re.match(r"^(.+?)\s+(HR\d{2})(?:\s+([\w./-]+))?$", blok[k + 1])
                    if mn:
                        naziv = mn.group(1).strip()
                        poziv_pr = (mn.group(2) + (" " + mn.group(3) if mn.group(3) else "")).strip()
                    else:
                        naziv = blok[k + 1].strip()
                break

        stavke.append({
            "red_br": len(stavke) + 1, "iznos": _f(izn), "smjer": dp,
            "datum": dat_valute or _d(dat), "datum_placanja": _d(dat),
            "naziv": naziv, "racun": racun, "opis": opis,
            "poziv_platitelja": poziv_pl, "poziv_primatelja": poziv_pr,
        })
        i = j

    return {"banka": "RBA", "broj": broj, "datum": datum, "stavke": stavke}


def parsiraj_izvod(putanja):
    """Glavna funkcija: otvori PDF, prepoznaj banku, parsiraj. Vrati dict ili None."""
    with pdfplumber.open(putanja) as pdf:
        uvod = (pdf.pages[0].extract_text() or "")
        banka = detektiraj_banku(uvod)
        if banka == "HPB":
            return parsiraj_hpb(pdf)
        if banka == "RBA":
            return parsiraj_rba(pdf)
    return None
