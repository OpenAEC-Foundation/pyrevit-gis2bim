# -*- coding: utf-8 -*-
"""Deurstaat - Aantallen Tellen.

Shim naar de gedeelde Kozijnstaat-Aantallen-logica met profile="deur".
Telt deur-instances (OST_Doors) per type (getekend + gespiegeld).

IronPython 2.7.
"""

__title__ = "Aantallen\nTellen"
__author__ = "3BM Bouwkunde"
__doc__ = "Tel deur-instances per type (getekend + gespiegeld)"

import os
import sys

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(SCRIPT_DIR))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import revit, forms

from kozijnstaat.shim import run_shared


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run_shared("KozijnstaatAantallen.pushbutton", "deur")
