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
