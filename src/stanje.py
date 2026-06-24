# -*- coding: utf-8 -*-
"""
Trajno pamćenje OBRAĐENIH računa i izvoda (datoteka stanje.json).

Zašto: ako korisnik RUČNO obriše neki red (npr. leasing kamatu), skripta to NE smije
ponovno upisati u idućem pokretanju. Zato pamtimo što je obrađeno NEOVISNO o knjizi —
gledamo ovaj popis, ne samo trenutno stanje Excela.
"""
import json
from pathlib import Path


def ucitaj(path):
    """Vrati {'racuni': set(Parra id), 'izvodi': set('BANKA|broj'), 'putni': set('pn XXX')}."""
    p = Path(path)
    if p.exists():
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
            return {"racuni": set(d.get("racuni", [])), "izvodi": set(d.get("izvodi", [])),
                    "putni": set(d.get("putni", []))}
        except Exception:
            pass
    return {"racuni": set(), "izvodi": set(), "putni": set()}


def spremi(stanje, path):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    d = {"racuni": sorted(stanje["racuni"]), "izvodi": sorted(stanje["izvodi"]),
         "putni": sorted(stanje.get("putni", []))}
    p.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def kljuc_izvoda(banka, broj):
    return f"{banka}|{broj}"
