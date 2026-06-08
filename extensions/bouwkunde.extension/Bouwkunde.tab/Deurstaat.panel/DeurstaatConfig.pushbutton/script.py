# -*- coding: utf-8 -*-
"""Deurstaat - Config.

Shim naar de gedeelde Kozijnstaat-Config-logica met profile="deur".
Bewerkt het deur-profiel (eigen user_config_deur.json). De gedeelde
form past labels/titels aan op het profiel en verbergt het glas-tag-veld.

IronPython 2.7.
"""

__title__ = "Config"
__author__ = "3BM Bouwkunde"
__doc__ = "Bewerk Deurstaat config (family namen, grid, offsets, refs)"

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
        run_shared("KozijnstaatConfig.pushbutton", "deur")
