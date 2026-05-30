# -*- coding: utf-8 -*-
"""Warmteverlies Grensvlak Wis — verwijder de grensvlak-DirectShapes.

Verwijdert alle DirectShapes met Comments-prefix "WV_BND" die door de
Grensvlak-check zijn aangemaakt.

IronPython 2.7 — geen f-strings, geen type hints.
"""

__title__ = "Grensvlak\nWis"
__author__ = "3BM Bouwkunde"
__doc__ = "Verwijder alle WV_BND grensvlak-DirectShapes uit het model"

import os
import sys

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import Transaction

# Lib pad toevoegen
sys.path.append(os.path.join(
    os.path.dirname(__file__), "..", "..", "..", "lib"
))

from warmteverlies.boundary_preview import clear_boundary_shapes


def run_grensvlak_wis(doc):
    """Verwijder alle WV_BND grensvlak-DirectShapes."""
    output = script.get_output()
    output.print_md("## Warmteverlies — Grensvlak Wis")

    count = 0
    t = Transaction(doc, "WV - Grensvlak-check wissen")
    t.Start()
    try:
        count = clear_boundary_shapes(doc)
        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(
            "Fout bij verwijderen grensvlakken:\n{0}".format(str(ex)),
            title="Fout",
        )
        return

    output.print_md(
        "Verwijderd: **{0}** WV_BND DirectShapes.".format(count)
    )


if __name__ == "__main__":
    doc = revit.doc
    if doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run_grensvlak_wis(doc)
