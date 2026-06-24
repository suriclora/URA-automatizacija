# -*- coding: utf-8 -*-
"""Male pomoćne funkcije koje koristi više modula."""


def ocisti(tekst):
    """Vrati samo slova i brojke, mala slova.
    Npr. 'UR-269/1-1' -> 'ur26911'. Koristi se za uspoređivanje
    brojeva računa s imenima/sadržajem PDF-ova (zanemaruje crtice, razmake...)."""
    if not tekst:
        return ""
    return "".join(filter(str.isalnum, str(tekst))).lower()
