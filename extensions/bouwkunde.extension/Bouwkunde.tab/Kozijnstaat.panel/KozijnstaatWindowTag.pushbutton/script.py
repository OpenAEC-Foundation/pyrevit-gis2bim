# -*- coding: utf-8 -*-
"""Kozijnstaat - Window Taggen.

Plaatst per kozijn in de actieve view een tag van een specifieke
family (default '31_TAG_wi_kozijnstaat_window'). De tag staat
horizontaal gecentreerd onder het kozijn, op een verticale offset
in view-coords.

IronPython 2.7.
"""

__title__ = "Window\nTaggen"
__author__ = "3BM Bouwkunde"
__doc__ = "Plaats window-tags onder elk kozijn in actieve view"

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
    XYZ,
)

from kozijnstaat.config import load_config
from kozijnstaat.family_collector import (
    collect_window_instances,
    get_symbol_height_mm,
)
from kozijnstaat.tag_placer import place_tag_with_family


MM_TO_FT = 1.0 / 304.8


def _safe_name(element):
    """Defensieve .Name-getter (IronPython 2.7 / Revit 2025 quirk)."""
    if element is None:
        return ""
    try:
        clr_type = element.GetType()
        prop = clr_type.GetProperty("Name")
        if prop is not None:
            v = prop.GetValue(element, None)
            if v is not None:
                return v
    except Exception:
        pass
    try:
        v = element.Name
        if v is not None:
            return v
    except Exception:
        pass
    return ""


def _tag_location_below_kozijn(instance, view, v_offset_mm):
    """Wereld-XYZ onder het kozijn, gecentreerd horizontaal.

    Anchor = bottom-center van kozijn (= placement-point voor hosted
    windows). Offset wordt toegepast langs view.UpDirection (negatief
    = onder de sill).
    """
    try:
        pt = instance.Location.Point
    except Exception:
        return None
    u = view.UpDirection
    v_ft = v_offset_mm * MM_TO_FT
    return XYZ(
        pt.X + u.X * v_ft,
        pt.Y + u.Y * v_ft,
        pt.Z + u.Z * v_ft,
    )


def run():
    doc = revit.doc
    view = doc.ActiveView
    output = script.get_output()
    output.print_md("## Kozijnstaat - Window Taggen")

    cfg = load_config()
    kozijn_family = cfg.get("kozijn_family", "3BM_kozijn")
    tag_family = cfg.get(
        "kozijn_tag_family", "31_TAG_wi_kozijnstaat_window"
    )
    v_offset_mm = float(cfg.get("kozijn_tag_v_offset_mm", -500.0))

    output.print_md(
        "Tag family: **{0}**, v-offset = **{1}** mm "
        "(negatief = onder kozijn)".format(tag_family, v_offset_mm)
    )

    instances = collect_window_instances(
        doc, name_contains=kozijn_family, view_id=view.Id,
    )
    if not instances:
        forms.alert(
            "Geen kozijnen '{0}' gevonden in actieve view '{1}'."
            .format(kozijn_family, _safe_name(view)),
            title="Geen kozijnen",
        )
        return

    # Filter tag-families uit (zelfde patroon als Aantallen)
    filtered = []
    excluded = 0
    for inst in instances:
        try:
            fam_name = _safe_name(inst.Symbol.Family)
        except Exception:
            fam_name = ""
        if "TAG" in fam_name.upper():
            excluded += 1
            continue
        filtered.append(inst)
    instances = filtered

    output.print_md(
        "Te taggen: **{0}** kozijnen (tag-families uitgesloten: {1})"
        .format(len(instances), excluded)
    )

    if not instances:
        forms.alert(
            "Geen kozijn-instances over om te taggen.",
            title="Niets te doen",
        )
        return

    tx = Transaction(doc, "Kozijnstaat - Tag kozijnen")
    tx.Start()
    try:
        n_placed = 0
        n_failed = 0
        for inst in instances:
            loc = _tag_location_below_kozijn(inst, view, v_offset_mm)
            if loc is None:
                n_failed += 1
                continue
            tag = place_tag_with_family(
                doc, view, inst, loc, tag_family,
            )
            if tag is not None:
                n_placed += 1
            else:
                n_failed += 1

        tx.Commit()

        output.print_md("---")
        output.print_md(
            "**Tags geplaatst:** {0}, **Mislukt:** {1}".format(
                n_placed, n_failed,
            )
        )
        forms.alert(
            "Klaar.\n{0} window-tags geplaatst, {1} mislukt.".format(
                n_placed, n_failed,
            ),
            title="Window Taggen",
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
