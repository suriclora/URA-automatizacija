# -*- coding: utf-8 -*-
"""
Komunikacija s Parra API-jem.

VAŽNO: dohvaćamo SAMO ULAZNE (dolazne) e-račune — one koje NAMA šalju
dobavljači. Endpoint je 'incoming-einvoices'. NIKADA ne diramo izlazne
(naše) račune.

Popravak u odnosu na original: prolazimo kroz SVE stranice (paginacija),
umjesto da gledamo samo prvih 50 i tražimo krhko "sidro".
"""
import os
import json

import requests

import config

# Sigurnosna provjera: smijemo zvati samo ulazni endpoint
_DOZVOLJENI_ENDPOINT = "incoming-einvoices"


def dohvati_ulazne_racune(token, business_id, base_url, logger,
                          page_size=50, max_stranica=100):
    """Vrati listu SVIH ulaznih e-računa s Parre (kroz sve stranice).

    DEMO način: bez pravog API-ja — čita izmišljene račune iz demo/data/parra_demo.json.

    Ako neka stranica ne uspije, zabilježi grešku i vrati ono što je dotad
    skupljeno (skripta NE smije pasti).
    """
    if getattr(config, "DEMO", False):
        put = os.path.join(str(config.KORIJEN), "demo", "data", "parra_demo.json")
        if os.path.exists(put):
            with open(put, encoding="utf-8") as f:
                demo = json.load(f)
            logger.info("   DEMO način: učitano %s izmišljenih računa iz parra_demo.json", len(demo))
            return demo
        logger.warning("   DEMO način, ali nema %s — vraćam prazno.", put)
        return []

    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    svi = []

    for page in range(1, max_stranica + 1):
        url = f"{base_url}/{business_id}/{_DOZVOLJENI_ENDPOINT}?page={page}&pageSize={page_size}"
        try:
            res = requests.get(url, headers=headers, timeout=30)
            res.raise_for_status()
        except requests.RequestException as e:
            logger.error("   Greška pri dohvaćanju s Parre (stranica %s): %s", page, e)
            break

        try:
            data = res.json().get("data", [])
        except ValueError:
            logger.error("   Parra je vratila neispravan odgovor (stranica %s).", page)
            break

        if not data:
            break

        svi.extend(data)
        logger.info("   Parra stranica %s: %s računa (ukupno %s)", page, len(data), len(svi))

        # Ako je stranica nepotpuna, to je zadnja stranica
        if len(data) < page_size:
            break
    else:
        logger.warning("   Dosegnut limit od %s stranica — ima li ih stvarno toliko?", max_stranica)

    return svi


def dohvati_detalj(token, business_id, base_url, invoice_id, logger=None, pokusaja=3):
    """Dohvati detalj jednog ulaznog e-računa (invoiceLines, KPD...). None ako ne uspije.
    Pokušava nekoliko puta (zbog povremenih grešaka API-ja)."""
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    url = f"{base_url}/{business_id}/incoming-einvoices/{invoice_id}"
    zadnja = None
    for _ in range(pokusaja):
        try:
            res = requests.get(url, headers=headers, timeout=30)
            res.raise_for_status()
            return res.json()
        except (requests.RequestException, ValueError) as e:
            zadnja = e
    if logger:
        logger.warning("   ne mogu dohvatiti detalj računa %s: %s", invoice_id, zadnja)
    return None


def filtriraj_nove(svi_racuni, postojeci_brojevi):
    """Zadrži samo račune kojih JOŠ NEMA u Excelu (dedup po broju računa).
    Sortira ih kronološki (po datumu dokumenta) da UR brojevi idu redom.
    """
    novi = []
    for r in svi_racuni:
        br = str(r.get("invoiceNumber", "")).strip()
        if br and br.lower() not in postojeci_brojevi:
            novi.append(r)

    # Kronološki redoslijed (datum + vrijeme) — stariji račun dobiva manji UR broj,
    # točno i unutar istog dana.
    novi.sort(key=lambda r: (r.get("documentDateInCet") or "")[:19])
    return novi
