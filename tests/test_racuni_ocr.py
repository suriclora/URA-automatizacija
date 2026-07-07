# -*- coding: utf-8 -*-
"""Testovi čitanja polja s računa (parser OCR teksta) — bez pravog OCR-a."""
from src import racuni_ocr as R


# ---------- _broj: pretvaranje teksta u novčani iznos ----------

def test_broj_europski_zapis():
    assert R._broj("1.234,56") == 1234.56
    assert R._broj("85,60") == 85.60


def test_broj_tocka_kao_decimala():
    # neki računi (npr. cestarine) koriste točku umjesto zareza
    assert R._broj("11.60") == 11.60


def test_broj_ocr_razmaci():
    # OCR zna ubaciti razmake u broj
    assert R._broj("11 .60") == 11.60
    assert R._broj("2 . 32") == 2.32


def test_broj_tisucice_bez_decimala():
    assert R._broj("1.234") == 1234.0


def test_broj_nevaljalo():
    assert R._broj("") is None
    assert R._broj(None) is None
    assert R._broj("abc") is None


# ---------- izvuci_polja: izvlačenje broja/datuma/iznosa iz teksta ----------

def test_polja_trgovina_standard():
    t = "Racun broj: 2026-512\nDatum: 03.05.2026.\n25  40,00  10,00\nZA PLATITI:  50,00 EUR"
    p = R.izvuci_polja(t)
    assert p["broj"] == "2026-512"
    assert p["datum"] == "03.05.2026"
    assert p["bruto"] == 50.0
    assert p["osnovica"] == 40.0     # iz rekapitulacije (stopa/osnovica/porez)
    assert p["pdv"] == 10.0


def test_polja_cestarina_kose_crte_i_tocka():
    # cestarina: datum s kosim crtama, iznos s točkom i razmakom, 'Ukupan iznos'
    t = "Racuna: 14026210200161\nIzlaz 20/05/2026 1 1 : 40\nEUR 11 .60 Ukupan znos 25 PDV 00% EUR 2 . 32"
    p = R.izvuci_polja(t)
    assert p["datum"] == "20.05.2026"
    assert p["bruto"] == 11.60


def test_polja_rastrkan_ocr_najveci_iznos():
    # kad ključna riječ nije uz broj -> uzmi najveći novčani iznos (ukupno)
    t = "Racun broj: 123-4-1\nDatum: 15.05.2026\n80,00\n20,00\n100,00 EUR\nZA PLATITI:"
    p = R.izvuci_polja(t)
    assert p["bruto"] == 100.0
    # datum NE smije biti shvaćen kao iznos (15.05 = 1505 bi bila greška)
    assert p["datum"] == "15.05.2026"


def test_polja_odbaci_besmislen_iznos():
    # dugi ID brojevi ne smiju postati 'bruto'
    t = "Terminal 432366300,00\nnema kljucnih rijeci"
    p = R.izvuci_polja(t)
    assert p["bruto"] is None or p["bruto"] < 100000


def test_polja_prekratak_broj_racuna_odbacen():
    t = "Racun broj: 15\nDatum: 01.02.2026."
    p = R.izvuci_polja(t)
    assert p["broj"] is None      # '15' je vjerojatno OCR smeće -> radije prazno
