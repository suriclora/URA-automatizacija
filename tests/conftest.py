# -*- coding: utf-8 -*-
"""Pytest konfiguracija: dodaj korijen projekta u sys.path da testovi vide src/ i main.py."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
