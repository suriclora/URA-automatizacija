# -*- coding: utf-8 -*-
"""
Pronalaženje PDF-a računa u mapama (Downloads i arhiva).

Popravak u odnosu na original: NE čitamo sadržaj SVIH PDF-ova unaprijed
(to je bilo sporo). Prvo tražimo po IMENU datoteke (brzo), a sadržaj PDF-a
čitamo tek ako po imenu ne nađemo (i zapamtimo pročitano da ne čitamo dvaput).
"""
from pathlib import Path
from pypdf import PdfReader

from src.utils import ocisti


def procitaj_tekst_pdf(putanja, max_stranica=3):
    """Vrati tekst prvih nekoliko stranica PDF-a (ili prazan string ako ne ide)."""
    try:
        reader = PdfReader(putanja)
        tekst = ""
        for page in reader.pages[:max_stranica]:
            tekst += page.extract_text() or ""
        return tekst
    except Exception:
        return ""


def skeniraj_imena(folder):
    """Vrati listu PDF-ova u mapi — SAMO imena i putanje (bez čitanja sadržaja).
    Brzo, jer ne otvara datoteke.
    """
    baza = []
    p = Path(folder)
    if p.exists():
        for f in p.iterdir():
            if f.suffix.lower() == ".pdf":
                baza.append({
                    "ime": f.name,
                    "putanja": str(f),
                    "ime_clean": ocisti(f.name),
                    "sadrzaj": None,  # čita se tek po potrebi (lazy)
                })
    return baza


def nadji_pdf(br_rac_clean, baze, sadrzaj_baze=("Downloads",)):
    """Pronađi PDF za zadani (očišćeni) broj računa.

    'baze' je lista parova (naziv_baze, lista_pdfova), npr.
        [("Downloads", baza_downloads), ("Arhiva", baza_arhiva)]
    Redoslijed je važan — prvo se pretražuje prva baza.

    'sadrzaj_baze' = nazivi baza u kojima SMIJEMO tražiti po sadržaju PDF-a.
    Po defaultu samo "Downloads": u arhivi NE tražimo po sadržaju jer to daje
    lažne pogotke (npr. novi HRT račun nađe stari HRT PDF jer dijele
    pretplatnički broj). U arhivi tražimo samo po imenu ("UR XXXX.pdf").

    Vraća (pdf_dict, naziv_baze) ili (None, None) ako ništa nije nađeno.
    """
    if len(br_rac_clean) <= 3:
        return None, None  # prekratak broj — preriskantno za podudaranje

    # 1) Brzo: traži po IMENU datoteke (u svim bazama)
    for naziv_baze, lista in baze:
        for pdf in lista:
            if br_rac_clean in pdf["ime_clean"]:
                return pdf, naziv_baze

    # 2) Sporije: traži po SADRŽAJU — SAMO u dopuštenim bazama (npr. Downloads)
    for naziv_baze, lista in baze:
        if naziv_baze not in sadrzaj_baze:
            continue
        for pdf in lista:
            if pdf["sadrzaj"] is None:
                pdf["sadrzaj"] = ocisti(procitaj_tekst_pdf(pdf["putanja"]))
            if pdf["sadrzaj"] and br_rac_clean in pdf["sadrzaj"]:
                return pdf, naziv_baze

    return None, None
