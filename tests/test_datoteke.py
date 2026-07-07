# -*- coding: utf-8 -*-
"""Testovi prepoznavanja datoteka (Parra imena, podudaranje s brojem računa)."""
from src import datoteke as D


def _dat(ime):
    """Minimalni zapis datoteke kakav vraća skeniraj_sve()."""
    from src.utils import ocisti
    return {"ime": ime, "ime_clean": ocisti(ime)}


# ---------- je_parra_ime ----------

def test_parra_xml_obrazac():
    assert D.je_parra_ime("UlazniERacun_12345.xml")


def test_parra_vizualni_pdf():
    assert D.je_parra_ime("(e) Račun 2007.PDF")
    assert D.je_parra_ime("(e) racun 315.pdf")


def test_obicna_datoteka_nije_parra():
    assert not D.je_parra_ime("IMG_20260611_134512.jpg")
    assert not D.je_parra_ime("izvod_55.pdf")


# ---------- ime_odgovara_racunu ----------

def test_ime_pocinje_brojem():
    assert D.ime_odgovara_racunu("2760/1/1", _dat("2760_1_1.pdf"))


def test_svi_dijelovi_broja_u_imenu():
    assert D.ime_odgovara_racunu("269-1-1", _dat("racun 269 1 1 svibanj.pdf"))


def test_parra_pdf_jezgra_broja():
    # Parrin vizualni PDF nosi samo jezgru broja ('2007' za '2007/1/1')
    assert D.ime_odgovara_racunu("2007/1/1", _dat("(e) Račun 2007.PDF"))


def test_krivi_broj_ne_odgovara():
    assert not D.ime_odgovara_racunu("2007/1/1", _dat("(e) Račun 9999.PDF"))


def test_prekratak_broj_odbijen():
    # brojevi od ≤3 znaka daju previše lažnih pogodaka -> uvijek False
    assert not D.ime_odgovara_racunu("12", _dat("12.pdf"))
