# -*- coding: utf-8 -*-
"""Verzamel unieke kozijn FamilySymbols (types) uit het actieve document."""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    FamilySymbol,
    FamilyInstance,
)


def _name(symbol):
    """Family symbol naam ophalen (IronPython-safe)."""
    try:
        return symbol.Family.Name
    except Exception:
        return ""


def collect_window_symbols(doc, name_contains=None):
    """Alle FamilySymbols in categorie Windows.

    Args:
        doc: Revit Document
        name_contains: optioneel - alleen families waarvan de
            Family.Name dit substring bevat (case-insensitive)

    Returns:
        list[FamilySymbol] gesorteerd op family-naam + type-naam
    """
    symbols = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Windows)
        .OfClass(FamilySymbol)
        .WhereElementIsElementType()
        .ToElements()
    )

    filtered = []
    needle = (name_contains or "").lower()
    for s in symbols:
        fam_name = _name(s)
        if needle and needle not in fam_name.lower():
            continue
        filtered.append(s)

    def sort_key(s):
        try:
            type_name = s.Name
        except Exception:
            type_name = ""
        return (_name(s), type_name)

    filtered.sort(key=sort_key)
    return filtered


def collect_window_instances(doc, name_contains=None, view_id=None):
    """Alle FamilyInstances in Windows-categorie.

    Args:
        doc: Revit Document
        name_contains: optioneel - alleen instances waarvan de
            Family.Name dit substring bevat (case-insensitive)
        view_id: optioneel - alleen in deze view zichtbaar

    Returns:
        list[FamilyInstance]
    """
    if view_id is not None:
        collector = FilteredElementCollector(doc, view_id)
    else:
        collector = FilteredElementCollector(doc)

    instances = (
        collector
        .OfCategory(BuiltInCategory.OST_Windows)
        .OfClass(FamilyInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if not name_contains:
        return list(instances)

    needle = name_contains.lower()
    result = []
    for inst in instances:
        try:
            fam_name = inst.Symbol.Family.Name
        except Exception:
            continue
        if needle in fam_name.lower():
            result.append(inst)
    return result


def get_symbol_width_mm(symbol):
    """Haal de breedte van een window FamilySymbol op in mm.

    Zoekt in deze volgorde:
      1. Built-in FAMILY_WIDTH_PARAM
      2. Parameter met exacte naam 'Width'
      3. Parameter 'Breedte' of 'kozijn_breedte'

    Returns:
        float (mm) of 0.0 als niet gevonden
    """
    from Autodesk.Revit.DB import BuiltInParameter

    try:
        p = symbol.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM)
        if p and p.HasValue:
            return p.AsDouble() * 304.8  # feet -> mm
    except Exception:
        pass

    for name in ("Width", "Breedte", "kozijn_breedte"):
        try:
            p = symbol.LookupParameter(name)
            if p and p.HasValue:
                return p.AsDouble() * 304.8
        except Exception:
            continue
    return 0.0


def get_symbol_height_mm(symbol):
    """Hoogte van window FamilySymbol in mm - zelfde strategie als breedte."""
    from Autodesk.Revit.DB import BuiltInParameter

    try:
        p = symbol.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM)
        if p and p.HasValue:
            return p.AsDouble() * 304.8
    except Exception:
        pass

    for name in ("Height", "Hoogte", "kozijn_hoogte"):
        try:
            p = symbol.LookupParameter(name)
            if p and p.HasValue:
                return p.AsDouble() * 304.8
        except Exception:
            continue
    return 0.0
