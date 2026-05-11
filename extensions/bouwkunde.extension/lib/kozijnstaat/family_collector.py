# -*- coding: utf-8 -*-
"""Verzamel unieke kozijn FamilySymbols (types) uit het actieve document."""

import re

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    FamilySymbol,
    FamilyInstance,
)

_NATURAL_SPLIT_RE = re.compile(r"(\d+)")


def _natural_key(value):
    """Tuple voor natural sort (KZ-2 voor KZ-10).

    Returnt een tuple van (kind, value) paren — kind=0 voor int, kind=1
    voor lowercased text — zodat ints en strings nooit direct met elkaar
    vergeleken worden.
    """
    if not value:
        return ((1, u""),)
    parts = _NATURAL_SPLIT_RE.split(value)
    out = []
    for p in parts:
        if not p:
            continue
        if p.isdigit():
            out.append((0, int(p)))
        else:
            out.append((1, p.lower()))
    return tuple(out) if out else ((1, u""),)


def _read_string_param(symbol, param_name):
    """Lees een string-parameter op een FamilySymbol; lege string bij absentie."""
    if not param_name:
        return u""
    try:
        p = symbol.LookupParameter(param_name)
    except Exception:
        return u""
    if p is None or not p.HasValue:
        return u""
    try:
        v = p.AsString()
        if v:
            return v
    except Exception:
        pass
    try:
        v = p.AsValueString()
        if v:
            return v
    except Exception:
        pass
    return u""

try:
    from kozijnstaat import logger as _log
except Exception:
    _log = None


def _try_log(level, msg):
    if _log is None:
        return
    try:
        getattr(_log, level)(msg)
    except Exception:
        pass


def _name(symbol):
    """Family symbol naam ophalen (IronPython-safe)."""
    try:
        return symbol.Family.Name
    except Exception:
        return ""


def collect_window_symbols(doc, name_contains=None, merk_param="merk"):
    """Alle FamilySymbols in categorie Windows.

    Args:
        doc: Revit Document
        name_contains: optioneel - alleen families waarvan de
            Family.Name dit substring bevat (case-insensitive)
        merk_param: type-parameter waarop gesorteerd wordt (natural sort).
            Leeg/None = fallback op family-naam + type-naam.

    Returns:
        list[FamilySymbol] gesorteerd op (merk_param value, family-naam,
        type-naam). Symbols zonder merk-waarde komen achteraan.
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
            type_name = u""
        merk = _read_string_param(s, merk_param) if merk_param else u""
        has_merk = 0 if merk else 1
        return (
            has_merk,
            _natural_key(merk),
            _natural_key(_name(s)),
            _natural_key(type_name),
        )

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


def _id_int(element):
    """Stabiel int uit een ElementId (Revit 2023 .Value, eerder .IntegerValue)."""
    if element is None:
        return None
    try:
        eid = element.Id
    except Exception:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(eid, attr)
            if callable(v):
                v = v()
            return int(v)
        except Exception:
            continue
    return None


def find_workset_id_by_name(doc, name):
    """Zoek de int-id van een UserWorkset op naam (case-sensitive).

    Returns:
        int (Workset.Id.IntegerValue), of None bij niet gevonden / niet
        workshared.
    """
    if not name:
        return None
    try:
        from Autodesk.Revit.DB import (
            FilteredWorksetCollector, WorksetKind,
        )
    except Exception:
        return None
    try:
        wsc = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    except Exception:
        return None
    for ws in wsc:
        try:
            if ws.Name == name:
                return int(ws.Id.IntegerValue)
        except Exception:
            continue
    return None


def _instance_workset_id(inst):
    """Lees de workset-id (int) van een instance via ELEM_PARTITION_PARAM."""
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        p = inst.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        if p and p.HasValue:
            return int(p.AsInteger())
    except Exception:
        pass
    return None


def find_first_instance_per_symbol(doc, symbols, exclude_workset_id=None):
    """Voor elke FamilySymbol de eerst-gevonden FamilyInstance in doc.

    Wordt door KozijnstaatCreate gebruikt om instance-parameters
    (bv. 'stelkozijn'/'sparing_type') te kopieren van de model-instance
    naar de canvas-instance. Bij projecten met meerdere instances per
    type krijgt de canvas de waarde van de eerste gevonden instance.

    Args:
        doc: Revit Document
        symbols: iterable van FamilySymbol — beperkt de zoekruimte
        exclude_workset_id: int — instances op deze workset overslaan.
            Gebruikt om canvas-instances (kozijnstaat-workset) uit te
            sluiten als source — anders kopieert een tweede run de
            defaults van een vorige run.

    Returns:
        dict[int, FamilyInstance] — keyed by symbol.Id int-waarde.
        Types zonder model-instance (buiten de exclude-workset) komen
        niet in de dict.
    """
    target_ids = set()
    for s in symbols:
        sid = _id_int(s)
        if sid is not None:
            target_ids.add(sid)

    result = {}
    if not target_ids:
        return result

    instances = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Windows)
        .OfClass(FamilyInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    for inst in instances:
        try:
            sym = inst.Symbol
        except Exception:
            continue
        sid = _id_int(sym)
        if sid is None or sid not in target_ids:
            continue
        if sid in result:
            continue
        if exclude_workset_id is not None:
            ws_id = _instance_workset_id(inst)
            if ws_id is not None and ws_id == exclude_workset_id:
                continue
        result[sid] = inst
    return result


def _safe_name(element):
    """Robuuste Name-getter, werkt rond IronPython 2.7 quirks."""
    if element is None:
        return u""
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
    return u""


def _symbol_label(symbol):
    fam = u""
    try:
        fam = _safe_name(symbol.Family)
    except Exception:
        fam = u""
    typ = _safe_name(symbol)
    if fam or typ:
        return u"{0} : {1}".format(fam or u"?", typ or u"?")
    return u"<onbekend>"


def _read_length_param_mm(symbol, param_name, label, dim_kind):
    """Lees een parameter en converteer naar mm.

    Logt diagnostische info: StorageType, AsDouble, AsValueString.
    Probeert eerst feet→mm conversie (Length parameter, default Revit
    intern). Bij waardes buiten sanity (kozijnen 100..6000 mm) wordt
    geprobeerd de raw waarde als mm te interpreteren (Number param).

    Args:
        symbol: FamilySymbol
        param_name: parameter naam (string)
        label: voor logging (bv. 'fam : type')
        dim_kind: 'width' / 'height' (alleen voor logging)

    Returns:
        float mm, of None bij ontbreken / onbruikbaar
    """
    from Autodesk.Revit.DB import StorageType

    try:
        p = symbol.LookupParameter(param_name)
    except Exception as ex:
        _try_log(
            "warn",
            u"{0}: {1} LookupParameter('{2}') raised {3}".format(
                dim_kind, label, param_name, ex,
            ),
        )
        return None

    if p is None:
        _try_log(
            "info",
            u"{0}: {1} param '{2}' niet aanwezig".format(
                dim_kind, label, param_name,
            ),
        )
        return None
    if not p.HasValue:
        _try_log(
            "info",
            u"{0}: {1} param '{2}' aanwezig maar leeg".format(
                dim_kind, label, param_name,
            ),
        )
        return None

    try:
        storage = p.StorageType
    except Exception:
        storage = None

    raw = 0.0
    try:
        if storage == StorageType.Double:
            raw = p.AsDouble()
        elif storage == StorageType.Integer:
            raw = float(p.AsInteger())
    except Exception as ex:
        _try_log(
            "warn",
            u"{0}: {1} param '{2}' AsDouble raised {3}".format(
                dim_kind, label, param_name, ex,
            ),
        )
        return None

    val_str = u""
    try:
        v = p.AsValueString()
        if v is not None:
            val_str = v
    except Exception:
        pass

    mm_as_length = raw * 304.8
    _try_log(
        "info",
        u"{0}: {1} '{2}' storage={3} raw={4} valueString='{5}' "
        u"(=> {6:.0f}mm if length-ft)".format(
            dim_kind, label, param_name, storage, raw,
            val_str, mm_as_length,
        ),
    )

    # Sanity: kozijn 100..6000mm
    if 100.0 <= mm_as_length <= 6000.0:
        return mm_as_length

    # Mogelijk een Number-param met mm-waarde direct erin
    if 100.0 <= raw <= 6000.0:
        _try_log(
            "warn",
            u"{0}: {1} '{2}' length-conversie buiten sanity "
            u"({3:.0f}mm); raw={4} valt binnen mm-range, "
            u"interpreteer als raw-mm".format(
                dim_kind, label, param_name,
                mm_as_length, raw,
            ),
        )
        return raw

    _try_log(
        "warn",
        u"{0}: {1} '{2}' raw={3} buiten alle sanity ranges, skip".format(
            dim_kind, label, param_name, raw,
        ),
    )
    return None


def get_symbol_width_mm(symbol):
    """Haal de breedte van een window FamilySymbol op in mm.

    Zoekt in deze volgorde:
      1. Custom param 'kozijn_breedte' (3BM-conventie)
      2. Built-in FAMILY_WIDTH_PARAM
      3. Parameters 'Width' / 'Breedte'

    Returns:
        float (mm) of 0.0 als niet gevonden
    """
    from Autodesk.Revit.DB import BuiltInParameter

    label = _symbol_label(symbol)

    for name in ("kozijn_breedte", "Breedte", "Width"):
        mm = _read_length_param_mm(symbol, name, label, "width")
        if mm is not None:
            return mm

    try:
        p = symbol.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM)
        if p and p.HasValue:
            ft = p.AsDouble()
            mm = ft * 304.8
            _try_log(
                "info",
                u"width: {0} via FAMILY_WIDTH_PARAM = "
                u"{1:.1f} ft -> {2:.0f} mm".format(label, ft, mm),
            )
            return mm
        else:
            _try_log(
                "info",
                u"width: {0} FAMILY_WIDTH_PARAM niet gevonden/leeg"
                .format(label),
            )
    except Exception as ex:
        _try_log(
            "warn",
            u"width: {0} FAMILY_WIDTH_PARAM raised {1}"
            .format(label, ex),
        )

    _try_log("warn", u"width: {0} -> 0.0 (geen bron)".format(label))
    return 0.0


def get_symbol_height_mm(symbol):
    """Hoogte van window FamilySymbol in mm.

    Zoekt 'kozijn_hoogte' (3BM) → FAMILY_HEIGHT_PARAM → 'Height'/'Hoogte'.
    """
    from Autodesk.Revit.DB import BuiltInParameter

    label = _symbol_label(symbol)

    for name in ("kozijn_hoogte", "Hoogte", "Height"):
        mm = _read_length_param_mm(symbol, name, label, "height")
        if mm is not None:
            return mm

    try:
        p = symbol.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM)
        if p and p.HasValue:
            ft = p.AsDouble()
            mm = ft * 304.8
            _try_log(
                "info",
                u"height: {0} via FAMILY_HEIGHT_PARAM = "
                u"{1:.1f} ft -> {2:.0f} mm".format(label, ft, mm),
            )
            return mm
        else:
            _try_log(
                "info",
                u"height: {0} FAMILY_HEIGHT_PARAM niet gevonden/leeg"
                .format(label),
            )
    except Exception as ex:
        _try_log(
            "warn",
            u"height: {0} FAMILY_HEIGHT_PARAM raised {1}"
            .format(label, ex),
        )

    _try_log("warn", u"height: {0} -> 0.0 (geen bron)".format(label))
    return 0.0
