# -*- coding: utf-8 -*-
"""Kozijnstaat - Aantallen tellen.

Telt kozijn FamilyInstances per Type, bepaalt per instance de
handedness (Left/Right/LhR/RhR) en schrijft de totalen + gespiegeld-
aantal terug naar parameters op het FamilyType.

Vervangt 31_3BM_Kozijnstaat_aantallen_tellen.dyn.
IronPython 2.7.
"""

__title__ = "Aantallen\nTellen"
__author__ = "3BM Bouwkunde"
__doc__ = "Tel kozijnen per type + spiegeling en schrijf naar family types"

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

from Autodesk.Revit.DB import Transaction, BuiltInParameter

from kozijnstaat.config import load_config
from kozijnstaat.family_collector import (
    collect_window_instances,
    find_workset_id_by_name,
)
from kozijnstaat.handedness import classify_many, is_mirrored


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


def _ensure_param(symbol, param_name, type_name, output):
    """Check of de parameter bestaat en writable is op de FamilyType."""
    p = symbol.LookupParameter(param_name)
    if p is None:
        output.print_md(
            "  - *parameter '{0}' niet gevonden op type '{1}'*"
            .format(param_name, type_name)
        )
        return None
    if p.IsReadOnly:
        output.print_md(
            "  - *parameter '{0}' is read-only op type '{1}'*"
            .format(param_name, type_name)
        )
        return None
    return p


def run():
    doc = revit.doc
    output = script.get_output()
    output.print_md("## Kozijnstaat - Aantallen Tellen")

    cfg = load_config()
    name_filter = cfg.get("name_filter_contains") or ""
    p_aantal = cfg.get("param_aantal", "aantal_getekend")
    p_gespiegeld = cfg.get("param_aantal_gespiegeld", "aantal_gespiegeld")
    workset_name = cfg.get("kozijnstaat_workset_name", "") or ""
    exclude_ws_id = None
    if workset_name:
        exclude_ws_id = find_workset_id_by_name(doc, workset_name)

    output.print_md(
        "Filter: family-naam bevat **'{0}'** | canvas-workset "
        "uitgesloten: **'{1}'** (id={2})".format(
            name_filter, workset_name, exclude_ws_id,
        )
    )

    # 1. Verzamel instances
    instances = collect_window_instances(doc, name_contains=name_filter)
    if not instances:
        forms.alert(
            "Geen kozijnen gevonden met filter '{0}'.\n"
            "Pas de filter aan via de Config-knop of plaats kozijnen."
            .format(name_filter),
            title="Geen kozijnen",
        )
        return

    # Filter tag-families uit (bv. 31_TAG_wi_kozijnstaat_window) en
    # kozijnstaat-canvas-instances uit (workset-based) — die zijn
    # visualisaties van het model en mogen niet meetellen.
    filtered = []
    excluded_tags = 0
    excluded_canvas = 0
    for inst in instances:
        try:
            fam_name = _safe_name(inst.Symbol.Family)
        except Exception:
            fam_name = ""
        if "TAG" in fam_name.upper():
            excluded_tags += 1
            continue
        if exclude_ws_id is not None:
            try:
                pw = inst.get_Parameter(
                    BuiltInParameter.ELEM_PARTITION_PARAM
                )
                if (pw and pw.HasValue
                        and int(pw.AsInteger()) == exclude_ws_id):
                    excluded_canvas += 1
                    continue
            except Exception:
                pass
        filtered.append(inst)
    instances = filtered

    output.print_md(
        "Gevonden: **{0}** kozijn-instances "
        "(tag-families uitgesloten: {1}, canvas-instances "
        "uitgesloten: {2})".format(
            len(instances), excluded_tags, excluded_canvas,
        )
    )

    if not instances:
        forms.alert(
            "Alleen tag-families gevonden, geen echte kozijnen.",
            title="Geen kozijnen",
        )
        return

    # 2. Groepeer per FamilySymbol (= Type)
    by_symbol = {}
    for inst in instances:
        try:
            sid = inst.Symbol.Id.IntegerValue
        except Exception:
            try:
                sid = inst.Symbol.Id.Value
            except Exception:
                continue
        by_symbol.setdefault(sid, []).append(inst)

    output.print_md(
        "Unieke types: **{0}**".format(len(by_symbol))
    )

    # 3. Per type: classificeer + schrijf parameters
    tx = Transaction(doc, "Kozijnstaat - Aantallen tellen")
    tx.Start()
    try:
        total_rows = 0
        total_skipped = 0
        output.print_md("---")
        output.print_md(
            "| Type | Totaal | Basis | Gespiegeld |"
        )
        output.print_md("|---|---:|---:|---:|")

        for sid, insts in by_symbol.items():
            symbol = insts[0].Symbol
            type_name = _safe_name(symbol) or "?"

            buckets = classify_many(insts)
            mirrored_count = sum(
                len(v) for k, v in buckets.items() if is_mirrored(k)
            )
            total = len(insts)
            basis = total - mirrored_count

            # Schrijf naar parameters — aantal_getekend = basis (niet-
            # gespiegeld), aantal_gespiegeld = mirrored. Som = total.
            p_a = _ensure_param(symbol, p_aantal, type_name, output)
            p_g = _ensure_param(symbol, p_gespiegeld, type_name, output)
            if p_a is None or p_g is None:
                total_skipped += 1
                continue

            try:
                p_a.Set(basis)
                p_g.Set(mirrored_count)
                total_rows += 1
            except Exception as ex:
                output.print_md(
                    "  - *Fout bij type '{0}': {1}*".format(type_name, ex)
                )
                total_skipped += 1
                continue

            output.print_md(
                "| {0} | {1} | {2} | {3} |".format(
                    type_name, total, basis, mirrored_count
                )
            )

        tx.Commit()

        output.print_md("---")
        output.print_md(
            "**Bijgewerkt:** {0} types, **Overgeslagen:** {1}"
            .format(total_rows, total_skipped)
        )
        forms.alert(
            "Klaar.\n{0} types bijgewerkt, {1} overgeslagen."
            .format(total_rows, total_skipped),
            title="Aantallen Tellen",
        )
    except Exception as ex:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        forms.alert(
            "Fout:\n{0}".format(ex),
            title="Aantallen Tellen",
        )
        output.print_md("**FOUT:** {0}".format(ex))


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run()
