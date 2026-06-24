# -*- coding: utf-8 -*-
"""
Fotkani papirnati računi: pretvaranje slike u PDF + pomoć kod spajanja na uplatu s izvoda.
korisnik drži fotke privremeno u Desktop\\URA; skripta ih OCR-a, pretvori u PDF i spremi
u pdf/ kao 'UR XXXX.pdf'.
"""
import os
import glob

from PIL import Image, ImageOps

SLIKE_EXT = (".jpg", ".jpeg", ".png", ".pdf")


def popis_fotki(folder):
    """Vrati listu putanja fotki/skenova u mapi (jpg/png/pdf)."""
    out = []
    if not os.path.isdir(folder):
        return out
    for f in sorted(glob.glob(os.path.join(folder, "*"))):
        if f.lower().endswith(SLIKE_EXT):
            out.append(f)
    return out


def pretvori_u_pdf(slika_path, cilj_pdf):
    """Pretvori sliku (jpg/png) u PDF na cilj_pdf. Ako je već PDF — samo kopiraj."""
    os.makedirs(os.path.dirname(cilj_pdf), exist_ok=True)
    if slika_path.lower().endswith(".pdf"):
        import shutil
        shutil.copy2(slika_path, cilj_pdf)
        return cilj_pdf
    img = ImageOps.exif_transpose(Image.open(slika_path)).convert("RGB")
    img.save(cilj_pdf, "PDF", resolution=150.0)
    return cilj_pdf
