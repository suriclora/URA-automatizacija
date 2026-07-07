# -*- coding: utf-8 -*-
"""Testovi povezivanja uplata s računima (leasing obroci)."""
from src import spajanje as S


def test_obrok_prije_rijeci():
    # XML računa: 'Sveukupno za platiti - 56. leasing obrok'
    assert S.broj_obroka("Sveukupno za platiti - 56. leasing obrok") == 56


def test_obrok_poslije_rijeci():
    # opis na izvodu: 'Leasing obrok 55'
    assert S.broj_obroka("Leasing obrok 55") == 55


def test_bez_obroka():
    assert S.broj_obroka("Plaćanje po računu 133627-1-1") is None
    assert S.broj_obroka("") is None
    assert S.broj_obroka(None) is None
