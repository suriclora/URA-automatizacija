# -*- coding: utf-8 -*-
"""Testovi pomoćnih funkcija iz main.py (UR iz naziva datoteke, hrvatski datumi)."""
from datetime import date

import main


# ---------- _ur_iz_imena: UR broj iz naziva PDF-a ----------

def test_ur_standardno_ime():
    assert main._ur_iz_imena("UR 0492.pdf") == 492


def test_ur_s_dodatkom():
    # ručno skenirani nalozi: 'UR 0492_PN 060.pdf'
    assert main._ur_iz_imena("UR 0492_PN 060.pdf") == 492
    assert main._ur_iz_imena("UR 0498_ PN 066.pdf") == 498


def test_ur_bez_razmaka():
    assert main._ur_iz_imena("UR0470.pdf") == 470


def test_ur_krivo_ime():
    assert main._ur_iz_imena("URA nesto.pdf") is None
    assert main._ur_iz_imena("racun 123.pdf") is None
    assert main._ur_iz_imena("UR 0492.txt") is None   # nije PDF


# ---------- _parse_datum_hr ----------

def test_datum_hr():
    assert main._parse_datum_hr("18.05.2026") == date(2026, 5, 18)


def test_datum_hr_nevaljao():
    assert main._parse_datum_hr("nije datum") is None
    assert main._parse_datum_hr("") is None
    assert main._parse_datum_hr(None) is None


# ---------- _nadji_uplatu: spajanje fotke s uplatom s izvoda ----------

def _uplate():
    from datetime import datetime
    return [
        {"red": 627, "placeno": 13.90, "datum": datetime(2026, 5, 20)},
        {"red": 630, "placeno": 99.82, "datum": datetime(2026, 5, 25)},
    ]


def test_uplata_iznos_i_datum():
    m = main._nadji_uplatu(13.90, date(2026, 5, 18), _uplate())
    assert m and m["red"] == 627


def test_uplata_nema_para():
    assert main._nadji_uplatu(50.00, date(2026, 5, 18), _uplate()) is None


def test_uplata_bez_datuma_spaja_po_iznosu():
    m = main._nadji_uplatu(99.82, None, _uplate())
    assert m and m["red"] == 630


def test_uplata_predaleko_datumski():
    # više od 14 dana razlike -> ne spajaj
    assert main._nadji_uplatu(13.90, date(2026, 7, 1), _uplate()) is None


def test_uplata_bez_brutoa():
    assert main._nadji_uplatu(None, date(2026, 5, 18), _uplate()) is None
