# -*- coding: utf-8 -*-
"""Legend-view builder voor Kozijnstaat (POC).

Revit API beperking
-------------------
Legend-views en LegendComponents kunnen NIET from-scratch worden
aangemaakt via de Revit API. Workaround: het document moet één
handmatig aangemaakte 'template' Legend-view bevatten met daarin
minstens één LegendComponent. Die wordt gedupliceerd en de
LegendComponent wordt 3x gekloond + omgezet via:
  - BuiltInParameter.LEGEND_COMPONENT      (FamilyType id)
  - BuiltInParameter.LEGEND_COMPONENT_VIEW (Plan / Front / Right)
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    View,
    ViewType,
    ViewDuplicateOption,
    ElementTransformUtils,
    ElementId,
    SubTransaction,
    Transaction,
    Transform,
    XYZ,
)
from System.Collections.Generic import List


# Chars die Revit niet accepteert in view-naam
_FORBIDDEN_VIEW_NAME_CHARS = u"<>:\"/\\|?*{}[];"


def _sanitize_view_name(name):
    """Strip Revit-verboden chars uit een view-naam."""
    if not name:
        return u"Kozijn legend"
    out = []
    for ch in unicode(name):
        if ch in _FORBIDDEN_VIEW_NAME_CHARS:
            out.append(u"_")
        else:
            out.append(ch)
    cleaned = u"".join(out).strip()
    return cleaned if cleaned else u"Kozijn legend"

try:
    from kozijnstaat import logger as _logmod
    _log = _logmod.make_logger("kozijnstaat_legend.log")
except Exception:
    _log = None


# LEGEND_COMPONENT_VIEW canonical integer-mapping (Revit 2020+)
# Bij andere Revit-versies kunnen deze indexen afwijken — pas in dat
# geval `LEGEND_VIEW_*` aan op basis van de raw waarde die we loggen
# uit de template-component.
LEGEND_VIEW_PLAN = 0
LEGEND_VIEW_FRONT = 2
LEGEND_VIEW_RIGHT = 5


def _log_info(msg):
    if _log is None:
        return
    try:
        _log.info(msg)
    except Exception:
        pass


def _log_warn(msg):
    if _log is None:
        return
    try:
        _log.warn(msg)
    except Exception:
        pass


def _id_int(eid):
    """ElementId -> int, robuust over Revit-versies."""
    if eid is None:
        return -1
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(eid, attr)
            if callable(v):
                v = v()
            return int(v)
        except Exception:
            continue
    return -1


def find_template_legend(doc, name):
    """Zoek een Legend-view op naam (case-insensitive).

    Returns:
        View of None
    """
    if not name:
        return None
    needle = name.lower()
    for v in FilteredElementCollector(doc).OfClass(View).ToElements():
        try:
            if v.IsTemplate:
                continue
            if v.ViewType != ViewType.Legend:
                continue
            if v.Name.lower() == needle:
                return v
        except Exception:
            continue
    return None


def list_legend_views(doc):
    """Returneer alle Legend-views in document (voor diagnose)."""
    out = []
    for v in FilteredElementCollector(doc).OfClass(View).ToElements():
        try:
            if not v.IsTemplate and v.ViewType == ViewType.Legend:
                out.append(v)
        except Exception:
            continue
    return out


def find_first_legend_component(doc, legend_view):
    """Eerste LegendComponent (FamilyInstance-like) in de legend-view."""
    instances = (
        FilteredElementCollector(doc, legend_view.Id)
        .OfCategory(BuiltInCategory.OST_LegendComponents)
        .ToElements()
    )
    return instances[0] if len(instances) > 0 else None


def _get_component_location_xyz(component):
    """Best-effort: lees positie van een LegendComponent.

    LegendComponents zijn FamilyInstance-achtig met Location.Point.
    Returns XYZ(0,0,0) als location niet leesbaar is.
    """
    try:
        loc = component.Location
        if loc is not None and hasattr(loc, "Point"):
            return loc.Point
    except Exception:
        pass
    return XYZ(0.0, 0.0, 0.0)


def _safe_set_int(element, bip, value):
    """Set integer-parameter, log bij falen."""
    try:
        p = element.get_Parameter(bip)
        if p is None:
            _log_warn(u"param {0} niet aanwezig".format(bip))
            return False
        if p.IsReadOnly:
            _log_warn(u"param {0} is read-only".format(bip))
            return False
        p.Set(value)
        return True
    except Exception as ex:
        _log_warn(u"set {0}={1} raised {2}".format(bip, value, ex))
        return False


def _safe_set_elementid(element, bip, eid):
    try:
        p = element.get_Parameter(bip)
        if p is None:
            _log_warn(u"param {0} niet aanwezig".format(bip))
            return False
        if p.IsReadOnly:
            _log_warn(u"param {0} is read-only".format(bip))
            return False
        p.Set(eid)
        return True
    except Exception as ex:
        _log_warn(u"set {0}=id raised {1}".format(bip, ex))
        return False


def build_kozijn_legend(
    doc,
    symbol,
    template_legend,
    source_component,
    new_name,
    scale=20,
    spacing_ft=3.0,
    uidoc=None,
):
    """Bouw een nieuwe Legend-view voor 1 kozijn met Plan/Front/Right.

    Layout:
        [Plan         ]
        [Front][Right]

    Args:
        doc: Revit Document. **Niet** binnen een outer Transaction
            aanroepen — de builder maakt zelf separate transactions
            per stap (init + 1 per view-placement). Reden: Revit
            internal errors tijdens `CopyElements` vergiftigen een
            outer transaction zodanig dat zelfs SubTransaction.RollBack
            de eindcommit niet meer kan redden.
        symbol: FamilySymbol van het kozijn-type
        template_legend: View (Legend) — wordt gedupliceerd
        source_component: LegendComponent uit template_legend — wordt
            3x gekopieerd naar de nieuwe legend
        new_name: naam voor de nieuwe Legend-view
        scale: view-scale denominator (1:20 → 20)
        spacing_ft: horizontale/verticale afstand tussen views in feet

    Returns:
        dict met new_legend, placed (lijst van ElementId), skipped (lijst)
    """
    _log_info(
        u"builder: enter — symbol_id={0} template_id={1} template_name='{2}' "
        u"new_name='{3}' scale=1:{4} spacing={5}ft".format(
            _id_int(symbol.Id), _id_int(template_legend.Id),
            template_legend.Name, new_name, scale, spacing_ft,
        )
    )

    # 1. Duplicate template — eigen Transaction
    _log_info(u"builder.tx_init: start (Duplicate + Name + Scale + cleanup)")
    tx_init = Transaction(doc, "Kozijnstaat Legend - init")
    tx_init.Start()
    try:
        _log_info(u"builder.1: template.Duplicate(Duplicate)")
        new_id = template_legend.Duplicate(ViewDuplicateOption.Duplicate)
        _log_info(u"builder.1 OK — new_id={0}".format(_id_int(new_id)))
    except Exception as ex:
        _log_warn(u"builder.1 FAILED — Duplicate raised {0}: {1}".format(
            type(ex).__name__, ex,
        ))
        try:
            tx_init.RollBack()
        except Exception:
            pass
        raise

    if _id_int(new_id) <= 0:
        _log_warn(
            u"builder.1: Duplicate returned invalid id ({0}) — abort. "
            u"Mogelijk staat de template-view in een phase/area waar "
            u"duplicate niet werkt, of is de view-naam '{1}' al bezet."
            .format(_id_int(new_id), new_name)
        )
        try:
            tx_init.RollBack()
        except Exception:
            pass
        raise Exception(
            "Duplicate gaf invalid ElementId terug — "
            "view-naam '{0}' is mogelijk al in gebruik.".format(new_name)
        )

    try:
        new_legend = doc.GetElement(new_id)
        _log_info(u"builder.1b: GetElement -> {0}".format(
            "OK" if new_legend is not None else "None",
        ))
    except Exception as ex:
        _log_warn(u"builder.1b FAILED — GetElement raised {0}: {1}".format(
            type(ex).__name__, ex,
        ))
        raise

    safe_name = _sanitize_view_name(new_name)
    if safe_name != new_name:
        _log_info(u"builder.1c: sanitized name '{0}' -> '{1}'".format(
            new_name, safe_name,
        ))
    try:
        new_legend.Name = safe_name
        _log_info(u"builder.1c: set Name='{0}' OK".format(safe_name))
    except Exception as ex:
        _log_warn(u"builder.1c FAILED — set Name='{0}' raised {1}: {2}".format(
            safe_name, type(ex).__name__, ex,
        ))
    try:
        new_legend.Scale = int(scale)
        _log_info(u"builder.1d: set Scale={0} OK".format(scale))
    except Exception as ex:
        _log_warn(u"builder.1d: set scale {0} raised {1}: {2}".format(
            scale, type(ex).__name__, ex,
        ))

    # 2. Verwijder bestaande LegendComponents in de duplicaat (Duplicate
    #    kopieert ze mee). We bouwen vers op vanuit source_component.
    try:
        existing_ids = (
            FilteredElementCollector(doc, new_id)
            .OfCategory(BuiltInCategory.OST_LegendComponents)
            .ToElementIds()
        )
        _log_info(u"builder.2: found {0} existing components in duplicate".format(
            existing_ids.Count if existing_ids else 0,
        ))
    except Exception as ex:
        _log_warn(u"builder.2: collector raised {0}: {1}".format(
            type(ex).__name__, ex,
        ))
        existing_ids = None

    if existing_ids and existing_ids.Count > 0:
        try:
            doc.Delete(existing_ids)
            _log_info(u"builder.2: deleted existing components")
        except Exception as ex:
            _log_warn(u"builder.2: delete raised {0}: {1}".format(
                type(ex).__name__, ex,
            ))

    # Commit init transaction — view bestaat nu fysiek met juiste
    # naam/schaal en is leeg van components. Eventuele errors bij de
    # placements hierna kunnen deze view dus niet meer kapot maken.
    try:
        tx_init.Commit()
        _log_info(u"builder.tx_init: commit OK")
    except Exception as ex:
        _log_warn(u"builder.tx_init: commit FAILED {0}: {1}".format(
            type(ex).__name__, ex,
        ))
        try:
            if tx_init.HasStarted() and not tx_init.HasEnded():
                tx_init.RollBack()
        except Exception:
            pass
        raise

    # Maak new_legend de actieve view — Revit's CopyElements naar
    # een Legend-view in een nieuwe Transaction faalt soms met
    # "internal error" als de target view nog niet "loaded" is in UI.
    if uidoc is not None:
        try:
            uidoc.ActiveView = new_legend
            _log_info(u"builder.tx_init+: ActiveView -> new_legend OK")
        except Exception as ex:
            _log_warn(u"builder.tx_init+: kon ActiveView niet zetten: {0}: {1}".format(
                type(ex).__name__, ex,
            ))

    # 3. Bepaal target-posities (in feet, view-space)
    s = float(spacing_ft)
    placement = [
        (u"Plan",  XYZ(0.0, 0.0, 0.0),  LEGEND_VIEW_PLAN),
        (u"Front", XYZ(0.0, -s, 0.0),   LEGEND_VIEW_FRONT),
        (u"Right", XYZ(s,   -s, 0.0),   LEGEND_VIEW_RIGHT),
    ]

    source_xyz = _get_component_location_xyz(source_component)
    _log_info(
        u"source component loc=({0:.2f},{1:.2f},{2:.2f})ft".format(
            source_xyz.X, source_xyz.Y, source_xyz.Z,
        )
    )

    # Diagnostiek: log de RAW view-mode integer van de source component
    # zodat we voor deze Revit-versie weten welke modes valide zijn.
    try:
        p_view = source_component.get_Parameter(
            BuiltInParameter.LEGEND_COMPONENT_VIEW,
        )
        if p_view is not None:
            _log_info(u"source component LEGEND_COMPONENT_VIEW raw={0}".format(
                p_view.AsInteger(),
            ))
        p_sym = source_component.get_Parameter(
            BuiltInParameter.LEGEND_COMPONENT,
        )
        if p_sym is not None:
            _log_info(u"source component LEGEND_COMPONENT raw_id={0}".format(
                _id_int(p_sym.AsElementId()),
            ))
    except Exception as ex:
        _log_warn(u"diag: read source params raised {0}: {1}".format(
            type(ex).__name__, ex,
        ))

    placed = []
    skipped = []
    source_id_list = List[ElementId]()
    source_id_list.Add(source_component.Id)

    for label, target_xyz, view_mode in placement:
        # Per-view ECHTE Transaction (geen SubTransaction). Reden:
        # Revit "internal error" tijdens CopyElements zet een outer
        # transaction in een corrupt state die ook na SubTransaction
        # .RollBack() niet meer commit-baar is. Door elk view-placement
        # in zijn eigen Transaction te zetten kan een faal alleen die
        # ene placement schaden — de eerder gecommitte views blijven.
        tx_place = Transaction(
            doc, "Kozijnstaat Legend - {0}".format(label),
        )
        try:
            tx_place.Start()
            _log_info(u"{0}.a: tx.Start OK — translation source=({1:.2f},{2:.2f}) target=({3:.2f},{4:.2f})".format(
                label, source_xyz.X, source_xyz.Y, target_xyz.X, target_xyz.Y,
            ))
            translation = Transform.CreateTranslation(target_xyz - source_xyz)
            _log_info(u"{0}.b: CreateTranslation OK — calling CopyElements".format(label))
            try:
                copied = ElementTransformUtils.CopyElements(
                    template_legend,
                    source_id_list,
                    new_legend,
                    translation,
                    None,
                )
                _log_info(u"{0}.c: CopyElements returned count={1}".format(
                    label, copied.Count if copied is not None else -1,
                ))
            except Exception as cex:
                _log_warn(u"{0}.c FAILED — CopyElements raised {1}: {2}".format(
                    label, type(cex).__name__, cex,
                ))
                try:
                    tx_place.RollBack()
                except Exception:
                    pass
                skipped.append(label)
                continue

            if copied is None or copied.Count == 0:
                _log_warn(u"{0}.c: CopyElements returned empty".format(label))
                tx_place.RollBack()
                skipped.append(label)
                continue
            new_eid = None
            for eid in copied:
                new_eid = eid
                break
            _log_info(u"{0}.d: new component id={1} — fetching element".format(
                label, _id_int(new_eid),
            ))
            new_comp = doc.GetElement(new_eid)
            _log_info(u"{0}.e: GetElement OK — setting LEGEND_COMPONENT (symbol)".format(label))

            ok_sym = _safe_set_elementid(
                new_comp, BuiltInParameter.LEGEND_COMPONENT, symbol.Id,
            )
            _log_info(u"{0}.f: set symbol -> {1} — setting LEGEND_COMPONENT_VIEW={2}".format(
                label, ok_sym, view_mode,
            ))
            ok_view = _safe_set_int(
                new_comp, BuiltInParameter.LEGEND_COMPONENT_VIEW, view_mode,
            )
            _log_info(u"{0}.g: set view-mode -> {1}".format(label, ok_view))

            if not ok_sym or not ok_view:
                _log_warn(
                    u"{0}: rollback — ok_sym={1} ok_view={2}".format(
                        label, ok_sym, ok_view,
                    )
                )
                tx_place.RollBack()
                skipped.append(label)
                continue

            _log_info(u"{0}.h: committing tx".format(label))
            tx_place.Commit()
            _log_info(
                u"{0}.i: COMMITTED — placed id={1} target=({2:.2f},{3:.2f}) "
                u"symbol_ok={4} view_ok={5} view_mode={6}".format(
                    label, _id_int(new_eid),
                    target_xyz.X, target_xyz.Y,
                    ok_sym, ok_view, view_mode,
                )
            )
            placed.append(new_eid)
        except Exception as ex:
            _log_warn(u"{0}: outer placement raised {1}: {2}".format(
                label, type(ex).__name__, ex,
            ))
            try:
                if tx_place.HasStarted() and not tx_place.HasEnded():
                    tx_place.RollBack()
            except Exception as rb_ex:
                _log_warn(u"{0}: rollback raised {1}: {2}".format(
                    label, type(rb_ex).__name__, rb_ex,
                ))
            skipped.append(label)

    return {
        "new_legend": new_legend,
        "placed": placed,
        "skipped": skipped,
    }
