# -*- coding: utf-8 -*-
"""Deurstaat - shim naar gedeelde Kozijnstaat-logica (profile="deur").

Roept KozijnstaatWindowTag.pushbutton.run("deur") aan via de gedeelde shim-loader.

IronPython 2.7.
"""

__title__ = "Deur\nTaggen"
__author__ = "3BM Bouwkunde"
__doc__ = "Plaats deur-tags onder elke deur-instance"

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
        run_shared("KozijnstaatWindowTag.pushbutton", "deur")
