# -*- coding: utf-8 -*-
"""Kozijnstaat - Maatvoering.

Plaatst per kozijn in de actieve view vier dimension-lines:
  - Horizontaal detail (via vakvulling_a1_l/r ... f1_l/r)
  - Verticaal detail (via vakvulling_a1_o/b ... a2_o/b)
  - Horizontaal hoofdmaat (Left-Right)
  - Verticaal hoofdmaat (Sill-Head)

Named references komen uit config.json (per-project aanpasbaar).

Vervangt 31_3BM_kozijnstaat_maatvoering.dyn.
IronPython 2.7.
"""

__title__ = "Maatvoeren"
__author__ = "3BM Bouwkunde"
__doc__ = "Plaats hoofd- en detailmaatvoering op kozijnen in actieve view"

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

from Autodesk.Revit.DB import Transaction, XYZ

from kozijnstaat.config import load_config
from kozijnstaat.family_collector import collect_window_instances
from kozijnstaat.reference_resolver import resolve_references
from kozijnstaat.dimension_creator import (
    get_instance_bbox,
    create_horizontal_dimension,
    create_vertical_dimension,
)

MM_TO_FT = 1.0 / 304.8

# Offsets van de dimension line t.o.v. de kozijn-BoundingBox (in mm)
DETAIL_H_OFFSET_MM = 250.0    # boven het kozijn
MAIN_H_OFFSET_MM = 500.0      # verder boven
DETAIL_V_OFFSET_MM = -250.0   # links van het kozijn
MAIN_V_OFFSET_MM = -500.0     # verder links


def _dim_at(doc, view, refs, direction, bbox, offset_mm):
    """Plaats dimension op offset t.o.v. bbox-min/max."""
    offset_ft = offset_mm * MM_TO_FT
    if direction == "horizontal":
        y = bbox.Max.Y + offset_ft
        return create_horizontal_dimension(doc, view, refs, y)
    return create_vertical_dimension(doc, view, refs,
                                     bbox.Min.X + offset_ft)


def run():
    doc = revit.doc
    view = doc.ActiveView
    output = script.get_output()
    output.print_md("## Kozijnstaat - Maatvoering")

    cfg = load_config()
    kozijn_family = cfg.get("kozijn_family", "3BM_kozijn")
    detail_h = list(cfg.get("detail_h_refs", []))
    detail_v = list(cfg.get("detail_v_refs", []))
    main_h = list(cfg.get("main_h_refs", []))
    main_v = list(cfg.get("main_v_refs", []))

    instances = collect_window_instances(
        doc, name_contains=kozijn_family, view_id=view.Id
    )
    if not instances:
        forms.alert(
            "Geen kozijnen '{0}' gevonden in actieve view '{1}'."
            .format(kozijn_family, view.Name),
            title="Geen kozijnen",
        )
        return

    output.print_md(
        "Kozijnen in view: **{0}**".format(len(instances))
    )

    tx = Transaction(doc, "Kozijnstaat - Plaats Maatvoering")
    tx.Start()
    try:
        n_dims = 0
        n_missing = 0
        missing_summary = {}

        for inst in instances:
            bbox = get_instance_bbox(inst, view)
            if bbox is None:
                continue

            # Detail horizontaal
            found_dh, missing_dh = resolve_references(inst, detail_h)
            if len(found_dh) >= 2:
                ordered_refs = [found_dh[n] for n in detail_h
                                if n in found_dh]
                if _dim_at(doc, view, ordered_refs, "horizontal",
                           bbox, DETAIL_H_OFFSET_MM):
                    n_dims += 1

            # Detail verticaal
            found_dv, missing_dv = resolve_references(inst, detail_v)
            if len(found_dv) >= 2:
                ordered_refs = [found_dv[n] for n in detail_v
                                if n in found_dv]
                if _dim_at(doc, view, ordered_refs, "vertical",
                           bbox, DETAIL_V_OFFSET_MM):
                    n_dims += 1

            # Hoofdmaat horizontaal
            found_mh, missing_mh = resolve_references(inst, main_h)
            if len(found_mh) >= 2:
                ordered_refs = [found_mh[n] for n in main_h
                                if n in found_mh]
                if _dim_at(doc, view, ordered_refs, "horizontal",
                           bbox, MAIN_H_OFFSET_MM):
                    n_dims += 1

            # Hoofdmaat verticaal
            found_mv, missing_mv = resolve_references(inst, main_v)
            if len(found_mv) >= 2:
                ordered_refs = [found_mv[n] for n in main_v
                                if n in found_mv]
                if _dim_at(doc, view, ordered_refs, "vertical",
                           bbox, MAIN_V_OFFSET_MM):
                    n_dims += 1

            # Verzamel missing references per kozijn
            for m in (missing_dh + missing_dv + missing_mh + missing_mv):
                missing_summary[m] = missing_summary.get(m, 0) + 1
                n_missing += 1

        tx.Commit()

        output.print_md("---")
        output.print_md(
            "**Dimension lines geplaatst:** {0}".format(n_dims)
        )
        if missing_summary:
            output.print_md(
                "**Ontbrekende references** ({0} totaal):".format(n_missing)
            )
            for name, cnt in sorted(
                missing_summary.items(), key=lambda x: -x[1]
            ):
                output.print_md("  - `{0}`: {1}x".format(name, cnt))

        forms.alert(
            "Klaar.\n{0} dimension lines geplaatst.\n"
            "{1} ontbrekende references (zie log).".format(
                n_dims, n_missing
            ),
            title="Maatvoering",
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
