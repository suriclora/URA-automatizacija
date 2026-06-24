# -*- coding: utf-8 -*-
"""
OCR fotkanih/skeniranih papirnatih računa (benzinska, dućan, restoran…).

OPREZ: OCR s fotki nije 100% pouzdan (kose crte u broju računa, rukopis se NE čita).
Zato ovo služi kao PREDPOPUNA — korisnik na ekranu vidi sliku računa i potvrdi/ispravi.
Iznose dodatno provjeravamo bruto-om s bankovnog izvoda.

Vadi: broj računa, datum, bruto (ukupno/za platiti), osnovica (neto), PDV, OIB.
"""
import io
import os
import re

import pytesseract
from PIL import Image, ImageOps

# Tesseract (instaliran preko winget), hrvatski model je u projektnom tessdata/
_TESS_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
_TESSDATA = os.path.join(os.path.dirname(os.path.dirname(__file__)), "tessdata")
if os.path.exists(_TESS_EXE):
    pytesseract.pytesseract.tesseract_cmd = _TESS_EXE
os.environ.setdefault("TESSDATA_PREFIX", _TESSDATA)


def _broj(s):
    """'1.234,56' / '85,60' -> float; None ako ne valja."""
    if not s:
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


# EasyOCR (deep-learning, puno bolji na mobilnim fotkama od Tesseracta).
# Reader se učita JEDNOM (start je spor + prvi put skine modele ~100 MB).
_READER = None


def _reader():
    global _READER
    if _READER is None:
        import easyocr
        _READER = easyocr.Reader(["hr"], gpu=False, verbose=False)
    return _READER


def _ocr_easy(img):
    """EasyOCR -> spojeni tekst (redovi posloženi po položaju, gore-dolje / lijevo-desno)."""
    import numpy as np
    res = _reader().readtext(np.array(img.convert("RGB")), detail=1, paragraph=False)
    res.sort(key=lambda r: (min(p[1] for p in r[0]), min(p[0] for p in r[0])))
    return "\n".join(str(r[1]) for r in res)


def _ocr(img):
    """Pročitaj tekst sa slike: EasyOCR (bolji); ako zakaže, Tesseract kao rezerva."""
    try:
        return _ocr_easy(img)
    except Exception:
        return pytesseract.image_to_string(img, lang="hrv")


def _najbolja_orijentacija(img):
    """Fotka zna biti zarotirana (bočno). Probaj 0/90/180/270 i uzmi tekst s NAJVIŠE
    prepoznatih znamenki i hrvatskih riječi (gruba, ali radi za račune)."""
    najbolji, ocjena = None, -1
    for ang in (0, 270):                      # uglavnom EXIF već uspravi; 270 hvata bočne
        slika = img if ang == 0 else img.rotate(ang, expand=True)
        t = _ocr(slika)
        bod = len(re.findall(r"\d", t)) + 5 * len(re.findall(
            r"(?i)\b(ukupno|porez|osnovica|ra[čc]un|datum|platiti|kartica|PDV)\b", t))
        if bod > ocjena:
            najbolji, ocjena = t, bod
        if ang == 0 and bod >= 8:
            break   # 0° je već dobar -> ne troši vrijeme na rotaciju (EasyOCR je sporiji)
    return najbolji or ""


def _ucitaj_sliku(path):
    """Učitaj sliku računa. Za .pdf renderiraj prvu stranicu (fitz)."""
    if path.lower().endswith(".pdf"):
        import fitz
        doc = fitz.open(path)
        pix = doc[0].get_pixmap(dpi=300)
        return Image.open(io.BytesIO(pix.tobytes("png")))
    return Image.open(path)


def _pdf_tekst(path):
    """Pročitaj UGRAĐENI tekst iz PDF-a (digitalni e-račun / PDF dobavljača) — točno, bez OCR-a.
    Vrati spojeni tekst svih stranica ili None ako PDF nema tekstualni sloj (skenirana slika)."""
    tekst = ""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            tekst = "\n".join((s.extract_text() or "") for s in pdf.pages)
    except Exception:
        try:
            from pypdf import PdfReader
            tekst = "\n".join((s.extract_text() or "") for s in PdfReader(path).pages)
        except Exception:
            return None
    # Skenirani PDF (samo slika) vrati malo/nimalo teksta -> tada radije OCR
    if len(re.sub(r"[^0-9A-Za-zČĆĐŠŽčćđšž]", "", tekst)) < 20:
        return None
    return tekst


def _priprema(path):
    img = _ucitaj_sliku(path)
    img = ImageOps.exif_transpose(img)        # poštuj EXIF rotaciju s mobitela
    img = img.convert("L")                    # sivo
    img = ImageOps.autocontrast(img)          # pojačaj kontrast (termalni papir)
    m = max(img.size)
    if m < 1600:                              # premale -> povećaj za bolji OCR
        f = 1600 / m
        img = img.resize((int(img.width * f), int(img.height * f)))
    elif m > 2000:                            # velike mobilne fotke -> blago smanji (brže, bez gubitka točnosti)
        f = 2000 / m
        img = img.resize((int(img.width * f), int(img.height * f)))
    return img


# rekapitulacija PDV-a: redak "STOPA  OSNOVICA  POREZ" (npr. '13,00  12,48  1,62')
_REKAP = re.compile(r"\b(0|5|13|25)(?:[.,]00)?\s+(\d{1,3}(?:[.,]\d{2}))\s+(\d{1,3}(?:[.,]\d{2}))")


def izvuci_polja(tekst):
    """Iz OCR teksta izvuci polja. Vrati dict (vrijednosti mogu biti None)."""
    t = tekst

    def naj(p, grp=1, flags=re.I):
        m = re.search(p, t, flags)
        return m.group(grp).strip() if m else None

    broj = naj(r"(?:R[-\s]?1\s*broj|Rn|ra[čc]un[a-z]*\s*(?:br\.?|broj)?)\s*[:.\s]\s*([0-9][\w\-/]+)")
    if broj and len(re.sub(r"[^0-9]", "", broj)) < 4:
        broj = None     # prekratko (npr. '15','05') = vjerojatno smeće -> radije prazno
    datum = naj(r"(\d{1,2}\.\d{1,2}\.\d{4})")
    oib_kupca = re.findall(r"\b(\d{11})\b", t)

    # bruto: 'ZA PLATITI', pa 'kartica ... X', pa 'UKUPNO'
    bruto = (naj(r"za\s*platiti\D{0,12}(\d[\d.]*,\d{2})")
             or naj(r"kartica\D{0,12}(\d[\d.]*,\d{2})")
             or naj(r"ukupno\s*€?\D{0,8}(\d[\d.]*,\d{2})"))

    # osnovica/PDV: prvo 'ukupno neto'/'ukupno porez' (INA), pa rekapitulacija (zbroj stopa)
    osnovica = naj(r"ukupn[oi]\s*neto\D{0,8}(\d[\d.]*,\d{2})")
    pdv = naj(r"ukupn[oi]\s*porez\D{0,8}(\d[\d.]*,\d{2})")
    if osnovica is None or pdv is None:
        rek = _REKAP.findall(t)
        if rek:
            so = round(sum(_broj(o) or 0 for _, o, _ in rek), 2)
            sp = round(sum(_broj(p) or 0 for _, _, p in rek), 2)
            osnovica = osnovica or (f"{so:.2f}".replace(".", ",") if so else None)
            pdv = pdv or (f"{sp:.2f}".replace(".", ",") if sp else None)

    return {
        "broj": broj,
        "datum": datum,
        "bruto": _broj(bruto),
        "osnovica": _broj(osnovica),
        "pdv": _broj(pdv),
        "oib": oib_kupca,
    }


def procitaj_racun(path):
    """Izvlačenje polja iz jednog računa. Vrati dict + 'tekst' (sirovi tekst) + 'izvor'.
    Digitalni PDF (s tekstualnim slojem) čita se DIREKTNO (točno, bez OCR-a);
    skenirani PDF i fotke idu na OCR."""
    tekst, izvor = None, "ocr"
    if path.lower().endswith(".pdf"):
        tekst = _pdf_tekst(path)
        if tekst:
            izvor = "pdf-tekst"
    if tekst is None:                              # fotka ili skenirani PDF -> OCR
        tekst = _najbolja_orijentacija(_priprema(path))
    polja = izvuci_polja(tekst)
    polja["tekst"] = tekst
    polja["datoteka"] = os.path.basename(path)
    polja["izvor"] = izvor
    return polja
