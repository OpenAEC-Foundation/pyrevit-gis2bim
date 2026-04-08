# -*- coding: utf-8 -*-
"""Hernoemt family types naar merk_bBREEDTE_hHOOGTE formaat.

Enkel type: interactieve dialoog met voorgestelde naam.
Meerdere types: opeenvolgende nummering (prefix + volgnummer).
"""
__title__ = "Type\nHernoem"
__author__ = "3BM Bouwkunde"
__doc__ = (
    "Hernoemt family types naar merk_bBREEDTE_hHOOGTE formaat.\n\n"
    "1 type: dialoog met voorgestelde naam.\n"
    "Meerdere types: opeenvolgend (prefix+nr_bBREEDTE_hHOOGTE)."
)

import os
import sys

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Element,
    FamilyInstance,
    FamilySymbol,
    StorageType,
)
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, revit, script

sys.path.append(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "lib")
)
from bm_logger import get_logger

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()
log = get_logger("TypeRename")

# =============================================================================
# CONFIGURATIE
# =============================================================================

FEET_TO_MM = 304.8

MERK_PARAMS = [
    "Merk", "merk", "Manufacturer", "Fabrikant", "Leverancier",
]
WIDTH_PARAMS = [
    "Breedte", "breedte", "Width", "width",
    "Kozijn_breedte", "kozijn_breedte",
    "Rough Width", "Dagmaat breedte", "b", "B",
]
HEIGHT_PARAMS = [
    "Hoogte", "hoogte", "Height", "height",
    "Kozijn_hoogte", "kozijn_hoogte",
    "Rough Height", "Dagmaat hoogte", "h", "H",
]

WIDTH_KEYWORDS = ["breedte", "width"]
HEIGHT_KEYWORDS = ["hoogte", "height"]


# =============================================================================
# PARAMETER HELPERS
# =============================================================================

def find_string_param(el, names):
    """Zoek een string parameter op naam, return eerste match."""
    for name in names:
        p = el.LookupParameter(name)
        if p and p.HasValue and p.StorageType == StorageType.String:
            v = p.AsString()
            if v and v.strip():
                return v.strip()
    try:
        p = el.get_Parameter(BuiltInParameter.ALL_MODEL_MANUFACTURER)
        if p and p.HasValue:
            v = p.AsString()
            if v and v.strip():
                return v.strip()
    except Exception:
        pass
    return None


def find_dim_param_mm(el, names, keywords=None):
    """Zoek een dimensie parameter, return waarde in mm (int)."""
    for name in names:
        p = el.LookupParameter(name)
        if p and p.HasValue and p.StorageType == StorageType.Double:
            v = p.AsDouble()
            if v > 0:
                return int(round(v * FEET_TO_MM))
    if keywords:
        for p in el.Parameters:
            pname = p.Definition.Name.lower()
            if p.HasValue and p.StorageType == StorageType.Double:
                for kw in keywords:
                    if kw in pname:
                        v = p.AsDouble()
                        if v > 0:
                            return int(round(v * FEET_TO_MM))
    return None


def get_element_type(el):
    """Haal het FamilySymbol / ElementType op uit een element."""
    if isinstance(el, FamilySymbol):
        return el
    try:
        type_id = el.GetTypeId()
        if type_id and type_id.IntegerValue > 0:
            return doc.GetElement(type_id)
    except Exception:
        pass
    return None


def read_type_data(el_type):
    """Lees merk, breedte, hoogte van een type. Return dict."""
    current_name = Element.Name.__get__(el_type)
    try:
        family_name = el_type.FamilyName
    except Exception:
        family_name = "System Family"

    merk = find_string_param(el_type, MERK_PARAMS)
    breedte = find_dim_param_mm(el_type, WIDTH_PARAMS, WIDTH_KEYWORDS)
    hoogte = find_dim_param_mm(el_type, HEIGHT_PARAMS, HEIGHT_KEYWORDS)

    return {
        "type": el_type,
        "current_name": current_name,
        "family_name": family_name,
        "merk": merk,
        "breedte": breedte,
        "hoogte": hoogte,
    }


def build_single_name(data):
    """Bouw voorgestelde naam voor enkel type."""
    parts = []
    if data["merk"]:
        parts.append(data["merk"])
    if data["breedte"] is not None:
        parts.append("b{}".format(data["breedte"]))
    if data["hoogte"] is not None:
        parts.append("h{}".format(data["hoogte"]))
    return "_".join(parts) if parts else data["current_name"]


def build_batch_name(prefix, nummer, data, pad):
    """Bouw naam voor batch: prefix+NR_bBREEDTE_hHOOGTE."""
    parts = ["{}{:0{}d}".format(prefix, nummer, pad)]
    if data["breedte"] is not None:
        parts.append("b{}".format(data["breedte"]))
    if data["hoogte"] is not None:
        parts.append("h{}".format(data["hoogte"]))
    return "_".join(parts)


# =============================================================================
# SELECTIE OPHALEN — alle unieke types
# =============================================================================

log.info("TypeRename gestart")

types_dict = {}  # {ElementId: el_type}
selected_ids = uidoc.Selection.GetElementIds()

if selected_ids and selected_ids.Count > 0:
    for eid in selected_ids:
        el = doc.GetElement(eid)
        if el:
            el_type = get_element_type(el)
            if el_type:
                tid = el_type.Id
                if tid not in types_dict:
                    types_dict[tid] = el_type

# Niets geselecteerd -> laat gebruiker klikken
if not types_dict:
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Klik op een element om het type te hernoemen"
        )
        el = doc.GetElement(ref.ElementId)
        el_type = get_element_type(el)
        if el_type:
            types_dict[el_type.Id] = el_type
    except Exception:
        script.exit()

if not types_dict:
    forms.alert(
        "Kon geen family type(s) bepalen uit de selectie.",
        title="Type Hernoemen",
        exitscript=True,
    )

type_list = list(types_dict.values())
log.info("{} unieke types gevonden".format(len(type_list)))


# =============================================================================
# ENKEL TYPE — interactieve dialoog
# =============================================================================

if len(type_list) == 1:
    data = read_type_data(type_list[0])
    proposed = build_single_name(data)

    info = (
        "Family: {fam}\n"
        "Huidig: {cur}\n\n"
        "Gevonden:\n"
        "  Merk: {merk}\n"
        "  Breedte: {b}\n"
        "  Hoogte: {h}"
    ).format(
        fam=data["family_name"],
        cur=data["current_name"],
        merk=data["merk"] or "(niet gevonden)",
        b="{} mm".format(data["breedte"]) if data["breedte"] else "(niet gevonden)",
        h="{} mm".format(data["hoogte"]) if data["hoogte"] else "(niet gevonden)",
    )

    new_name = forms.ask_for_string(
        prompt=info,
        title="Type Hernoemen \u2014 merk_bBREEDTE_hHOOGTE",
        default=proposed,
    )

    if not new_name:
        script.exit()
    if new_name == data["current_name"]:
        forms.alert("Naam is ongewijzigd.", title="Type Hernoemen")
        script.exit()

    with revit.Transaction("Type hernoemen naar {}".format(new_name)):
        try:
            data["type"].Name = new_name
            log.info("Hernoemd: {} -> {}".format(data["current_name"], new_name))
        except Exception as ex:
            log.error("Fout: {}".format(str(ex)))
            forms.alert(
                "Fout bij hernoemen:\n{}".format(str(ex)),
                title="Type Hernoemen",
            )
            log.finalize(success=False)
            script.exit()

    output.print_md("# Type Hernoemen")
    output.print_md("**{}** \u2192 **{}**".format(data["current_name"], new_name))
    output.print_md("*Family: {}*".format(data["family_name"]))
    log.finalize(success=True)
    script.exit()


# =============================================================================
# MEERDERE TYPES — batch met opeenvolgende nummering
# =============================================================================

# Lees data voor alle types
all_data = [read_type_data(t) for t in type_list]

# Sorteer op breedte (oplopend), dan hoogte (oplopend)
all_data.sort(key=lambda d: (d["breedte"] or 0, d["hoogte"] or 0))

# Vraag prefix en startnummer
prefix = forms.ask_for_string(
    prompt="{} types geselecteerd.\n\n"
           "Voer een prefix in voor de opeenvolgende nummering.\n"
           "Resultaat: PREFIX01_bBREEDTE_hHOOGTE".format(len(all_data)),
    title="Batch Hernoemen \u2014 Prefix",
    default="K",
)
if not prefix:
    script.exit()

start_str = forms.ask_for_string(
    prompt="Startnummer voor de reeks:",
    title="Batch Hernoemen \u2014 Startnummer",
    default="1",
)
if not start_str:
    script.exit()

try:
    start_nr = int(start_str)
except ValueError:
    forms.alert("Ongeldig startnummer.", title="Type Hernoemen", exitscript=True)

# Bepaal zero-padding breedte
end_nr = start_nr + len(all_data) - 1
pad = max(2, len(str(end_nr)))

# Bouw voorgestelde namen
proposals = []
for i, data in enumerate(all_data):
    nummer = start_nr + i
    new_name = build_batch_name(prefix, nummer, data, pad)
    proposals.append((data, new_name))

# Preview tonen
preview_lines = [
    "{} types worden hernoemd:\n".format(len(proposals)),
]
for data, new_name in proposals:
    preview_lines.append(
        "  {} \u2192 {}".format(data["current_name"], new_name)
    )

if not forms.alert(
    "\n".join(preview_lines),
    title="Batch Hernoemen \u2014 Bevestiging",
    yes=True,
    no=True,
):
    script.exit()

# Hernoemen in een transactie
renamed = 0
errors = []

with revit.Transaction("Batch type hernoemen ({} types)".format(len(proposals))):
    for data, new_name in proposals:
        try:
            data["type"].Name = new_name
            renamed += 1
            log.info("Hernoemd: {} -> {}".format(data["current_name"], new_name))
        except Exception as ex:
            errors.append((data["current_name"], new_name, str(ex)))
            log.error("Fout bij {}: {}".format(data["current_name"], str(ex)))

# Rapport
output.print_md("# Batch Type Hernoemen")
output.print_md("---")
output.print_md(
    "**Hernoemd: {}** | Fouten: {}".format(renamed, len(errors))
)
output.print_md("")

if renamed > 0:
    output.print_md("## Hernoemd")
    for data, new_name in proposals:
        cur = data["current_name"]
        if cur not in [e[0] for e in errors]:
            output.print_md("- **{}** \u2192 **{}**".format(cur, new_name))

if errors:
    output.print_md("## Fouten")
    for cur, target, err in errors:
        output.print_md("- {} \u2192 {}: `{}`".format(cur, target, err[:80]))

output.print_md("---")
output.print_md(
    "*Gesorteerd op breedte/hoogte, prefix: {}, nummering: {:0{}d}-{:0{}d}*".format(
        prefix, start_nr, pad, end_nr, pad
    )
)

log.finalize(success=renamed > 0)
