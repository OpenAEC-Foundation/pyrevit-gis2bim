# -*- coding: utf-8 -*-
"""Kozijnstaat - Glas Taggen.

Plaatst een tag (default family 'GEN_glas_v3') op alle glas-elementen
in de actieve view. Glas-elementen = FamilyInstances waarvan de
Family.Name 'glas' bevat (case-insensitive), inclusief sub-elements
van kozijn-families.

Vervangt 31_3BM_Kozijnstaat_tag_glas.dyn.
IronPython 2.7.
"""

__title__ = "Glas\nTaggen"
__author__ = "3BM Bouwkunde"
__doc__ = "Plaats glas-tags op alle glas-elementen in actieve view"

import os
import sys

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    FamilyInstance,
    BuiltInCategory,
)

from kozijnstaat.config import load_config
from kozijnstaat.tag_placer import (
    place_tag_with_family,
    get_bbox_center,
    offset_point,
)


GLAS_KEYWORD = "glas"


def _collect_glass_elements(doc, view_id):
    """Verzamel alle FamilyInstances met 'glas' in de family-naam.

    Kijkt in de actieve view, zonder categorie-restrictie omdat glas
    sub-elements van kozijnfamilies tot Windows of Generic Models kan
    behoren.
    """
    instances = (
        FilteredElementCollector(doc, view_id)
        .OfClass(FamilyInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    result = []
    for inst in instances:
        try:
            fam_name = inst.Symbol.Family.Name.lower()
        except Exception:
            continue
        if GLAS_KEYWORD in fam_name:
            result.append(inst)
    return result


def run():
    doc = revit.doc
    view = doc.ActiveView
    output = script.get_output()
    output.print_md("## Kozijnstaat - Glas Taggen")

    cfg = load_config()
    tag_family = cfg.get("glas_tag_family", "GEN_glas_v3")
    dx_mm = float(cfg.get("glas_tag_offset_x_mm", -500.0))
    dy_mm = float(cfg.get("glas_tag_offset_y_mm", 500.0))

    output.print_md(
        "Tag family: **{0}**, offset (x,y) = ({1}, {2}) mm"
        .format(tag_family, dx_mm, dy_mm)
    )

    elements = _collect_glass_elements(doc, view.Id)
    if not elements:
        forms.alert(
            "Geen glas-elementen gevonden in actieve view '{0}'.\n"
            "Zoekt naar family-naam bevattende '{1}'."
            .format(view.Name, GLAS_KEYWORD),
            title="Geen glas",
        )
        return

    output.print_md(
        "Gevonden: **{0}** glas-elementen".format(len(elements))
    )

    tx = Transaction(doc, "Kozijnstaat - Tag glas")
    tx.Start()
    try:
        n_placed = 0
        n_failed = 0
        for elem in elements:
            center = get_bbox_center(elem, view)
            if center is None:
                n_failed += 1
                continue

            loc = offset_point(center, dx_mm=dx_mm, dy_mm=dy_mm)
            tag = place_tag_with_family(
                doc, view, elem, loc, tag_family
            )
            if tag is not None:
                n_placed += 1
            else:
                n_failed += 1

        tx.Commit()

        output.print_md("---")
        output.print_md(
            "**Tags geplaatst:** {0}, **Mislukt:** {1}"
            .format(n_placed, n_failed)
        )

        forms.alert(
            "Klaar.\n{0} glas-tags geplaatst, {1} mislukt.".format(
                n_placed, n_failed
            ),
            title="Glas Taggen",
        )
    except Exception as ex:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        forms.alert("Fout:\n{0}".format(ex), title="Fout")
        output.print_md("**FOUT:** {0}".format(ex))


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run()
