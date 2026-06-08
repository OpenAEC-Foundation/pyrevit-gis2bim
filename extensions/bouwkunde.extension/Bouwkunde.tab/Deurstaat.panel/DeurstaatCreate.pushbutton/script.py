# -*- coding: utf-8 -*-
"""Deurstaat - Create.

Shim naar de gedeelde Kozijnstaat-Create-logica met profile="deur".
Plaatst alle unieke deur-types (OST_Doors) op een canvas-grid.

IronPython 2.7.
"""

__title__ = "Create\nDeurstaat"
__author__ = "3BM Bouwkunde"
__doc__ = "Plaats unieke deurtypes op een geselecteerde wand"

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
        run_shared("KozijnstaatCreate.pushbutton", "deur")
