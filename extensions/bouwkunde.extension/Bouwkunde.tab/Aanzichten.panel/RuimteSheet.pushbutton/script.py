# -*- coding: utf-8 -*-
"""Ruimte Sheet - genereer aanzichten + plattegrond + plafond op sheet.

Workflow:
  1. Gebruiker selecteert een ruimte (Room).
  2. UI vraagt templates (elevation/floor/ceiling), titleblock, sheet#
     en sheetnaam.
  3. Per wand in de boundary van de ruimte wordt een loodrecht
     aanzicht gegenereerd. Werkt ook met L-vormige of grillige
     plattegronden — elk segment krijgt z'n eigen aanzicht.
  4. Plattegrond en plafondaanzicht van het level worden
     gedupliceerd en gecropt op de room bounding box.
  5. Alle views worden op een nieuwe sheet geplaatst in een grid.

Schaal voor alle views: 1:20.

IronPython 2.7 / Revit API.
"""

__title__ = "Ruimte\nSheet"
__author__ = "3BM Bouwkunde"
__doc__ = "Genereer aanzichten + plattegrond + plafond van een ruimte op sheet"

import os
import sys
import math
import datetime
import traceback as _tb_mod

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)


# Early-log
_EARLY_LOG_DIR = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")),
    "3bm_exchange",
)
_EARLY_LOG_FILE = os.path.join(_EARLY_LOG_DIR, "ruimte_sheet_debug.log")


def _log(msg):
    try:
        if not os.path.isdir(_EARLY_LOG_DIR):
            os.makedirs(_EARLY_LOG_DIR)
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f = open(_EARLY_LOG_FILE, "a")
        try:
            f.write("[{0}] {1}\n".format(ts, msg))
        finally:
            f.close()
    except Exception:
        pass


def _log_exc(msg):
    try:
        tb = _tb_mod.format_exc()
    except Exception:
        tb = "<no traceback>"
    _log("{0}\n{1}".format(msg, tb))


_log("===== SCRIPT LOAD =====")

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    XYZ,
    Line,
    BoundingBoxXYZ,
    ElementId,
    ElementTransformUtils,
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    FamilySymbol,
    Wall,
    ElevationMarker,
    ViewFamilyType,
    ViewFamily,
    ViewSheet,
    Viewport,
    View,
    ViewPlan,
    ViewType,
    ViewDuplicateOption,
    SpatialElementBoundaryOptions,
    SpatialElementBoundaryLocation,
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException


SCALE = 20
MARKER_OFFSET_FT = 3.0   # afstand van wand tot marker, in ft (~915mm)
WALL_MARGIN_FT = 1.0     # extra ruimte links/rechts/onder/boven crop
ROOM_MARGIN_FT = 2.0     # margin rond room voor plattegrond/plafond


# ---------------------------------------------------------------- helpers

def _name(elem):
    if elem is None:
        return ""
    try:
        t = elem.GetType().GetProperty("Name")
        if t is not None:
            v = t.GetValue(elem, None)
            if v is not None:
                return v
    except Exception:
        pass
    try:
        return elem.Name or ""
    except Exception:
        return ""


def _id_int(elem):
    if elem is None:
        return None
    try:
        eid = elem.Id
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


class _RoomFilter(ISelectionFilter):
    def AllowElement(self, e):
        try:
            return isinstance(e, Room)
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


def _pick_room(uidoc):
    # Eerst kijken of selectie al een room bevat
    try:
        sel_ids = list(uidoc.Selection.GetElementIds())
        if len(sel_ids) == 1:
            e = uidoc.Document.GetElement(sel_ids[0])
            if isinstance(e, Room):
                return e
    except Exception:
        pass
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, _RoomFilter(),
            "Selecteer een ruimte",
        )
        return uidoc.Document.GetElement(ref.ElementId)
    except OperationCanceledException:
        return None


def _get_view_family_type(doc, vf):
    """Eerste matching ViewFamilyType voor gegeven ViewFamily."""
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        try:
            if vft.ViewFamily == vf:
                return vft
        except Exception:
            continue
    return None


def _collect_templates(doc, view_type_set):
    """Geef list views die templates zijn EN matchen op ViewType set.

    view_type_set: set/tuple van ViewType waarden.
    """
    out = []
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if not v.IsTemplate:
                continue
            if v.ViewType in view_type_set:
                out.append(v)
        except Exception:
            continue
    out.sort(key=lambda x: _name(x).lower())
    return out


def _collect_titleblocks(doc):
    out = []
    for fs in FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_TitleBlocks) \
            .OfClass(FamilySymbol):
        out.append(fs)
    out.sort(key=lambda fs: (_name(fs.Family), _name(fs)))
    return out


def _select_optional(label, items, name_func):
    """Toon SelectFromList met items + '(geen)' optie. Returns elem of None."""
    if not items:
        return None
    name_to_elem = {}
    display = ["(geen template)"]
    for it in items:
        nm = name_func(it)
        # bij dubbele namen suffix met id
        if nm in name_to_elem:
            nm = u"{0} [{1}]".format(nm, _id_int(it))
        name_to_elem[nm] = it
        display.append(nm)
    pick = forms.SelectFromList.show(
        display, title=label, button_name="Kies", multiselect=False,
    )
    if pick is None or pick == "(geen template)":
        return None
    return name_to_elem.get(pick)


def _select_required(label, items, name_func):
    if not items:
        forms.alert(
            "Geen items beschikbaar voor: {0}".format(label),
            title="Fout",
        )
        return None
    name_to_elem = {}
    display = []
    for it in items:
        nm = name_func(it)
        if nm in name_to_elem:
            nm = u"{0} [{1}]".format(nm, _id_int(it))
        name_to_elem[nm] = it
        display.append(nm)
    pick = forms.SelectFromList.show(
        display, title=label, button_name="Kies", multiselect=False,
    )
    if pick is None:
        return None
    return name_to_elem.get(pick)


def _tb_display(fs):
    fam = ""
    try:
        fam = _name(fs.Family)
    except Exception:
        pass
    return u"{0} : {1}".format(fam, _name(fs))


# ---------------------------------------------------------------- boundary

def _get_wall_segments(room):
    """Geef list van (wall, curve_segment) voor elk wand-grensstuk.

    Curve_segment is de DEEL-curve die deze ruimte begrenst (kan
    korter zijn dan de hele wand). Werkt voor L-vormig / grillig.
    """
    opts = SpatialElementBoundaryOptions()
    opts.SpatialElementBoundaryLocation = (
        SpatialElementBoundaryLocation.Finish
    )
    loops = room.GetBoundarySegments(opts)
    if loops is None:
        return []
    doc = room.Document
    out = []
    for loop in loops:
        for seg in loop:
            try:
                eid = seg.ElementId
            except Exception:
                continue
            if eid is None or _id_int_from_id(eid) <= 0:
                continue
            elem = doc.GetElement(eid)
            if not isinstance(elem, Wall):
                continue
            try:
                curve = seg.GetCurve()
            except Exception:
                continue
            if curve is None:
                continue
            out.append((elem, curve))
    return out


def _id_int_from_id(eid):
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(eid, attr)
            if callable(v):
                v = v()
            return int(v)
        except Exception:
            continue
    return -1


# ---------------------------------------------------------------- elevation

def _room_center_xy(room):
    try:
        loc = room.Location
        p = loc.Point
        return XYZ(p.X, p.Y, p.Z)
    except Exception:
        pass
    bb = room.get_BoundingBox(None)
    if bb is None:
        return XYZ.Zero
    return XYZ(
        (bb.Min.X + bb.Max.X) / 2.0,
        (bb.Min.Y + bb.Max.Y) / 2.0,
        (bb.Min.Z + bb.Max.Z) / 2.0,
    )


def _wall_height_ft(wall):
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        if p and p.HasValue:
            return p.AsDouble()
    except Exception:
        pass
    return 10.0


def _create_wall_elevation(
    doc, plan_view_id, vft_elev_id, wall, curve_seg,
    room_center, level_elev_ft, scale, template_id,
):
    """Maak 1 binnen-elevation loodrecht op het wand-segment.

    - marker geplaatst aan room-zijde van de wand, op `MARKER_OFFSET_FT`
      van de wand af.
    - marker gedraaid zodat index 0 NAAR de wand kijkt.
    - crop ingesteld op segment-lengte (+ margin) horizontaal,
      en wandhoogte (+ margin) verticaal.
    """
    p0 = curve_seg.GetEndPoint(0)
    p1 = curve_seg.GetEndPoint(1)
    seg_len = p0.DistanceTo(p1)
    if seg_len < 1e-6:
        return None
    seg_dir = XYZ(
        (p1.X - p0.X) / seg_len,
        (p1.Y - p0.Y) / seg_len,
        0.0,
    )
    seg_mid = XYZ(
        (p0.X + p1.X) / 2.0,
        (p0.Y + p1.Y) / 2.0,
        (p0.Z + p1.Z) / 2.0,
    )

    # twee perpendiculairen in XY-vlak
    perp_a = XYZ(-seg_dir.Y, seg_dir.X, 0.0)
    to_center = XYZ(
        room_center.X - seg_mid.X,
        room_center.Y - seg_mid.Y,
        0.0,
    )
    if perp_a.DotProduct(to_center) >= 0:
        room_normal = perp_a   # wijst INTO room
    else:
        room_normal = XYZ(seg_dir.Y, -seg_dir.X, 0.0)

    # Marker in de room geplaatst, marker view direction = -room_normal
    marker_pt = XYZ(
        seg_mid.X + room_normal.X * MARKER_OFFSET_FT,
        seg_mid.Y + room_normal.Y * MARKER_OFFSET_FT,
        level_elev_ft + 1.0,
    )

    marker = ElevationMarker.CreateElevationMarker(
        doc, vft_elev_id, marker_pt, scale,
    )

    # Default index 0 view direction is +X richting (geverifieerd via
    # rotation=0). We willen view-direction = -room_normal, dus
    # roteer marker met angle = atan2(-room_normal.Y, -room_normal.X).
    view_dir = XYZ(-room_normal.X, -room_normal.Y, 0.0)
    angle = math.atan2(view_dir.Y, view_dir.X)
    if abs(angle) > 1e-9:
        axis = Line.CreateBound(
            marker_pt,
            XYZ(marker_pt.X, marker_pt.Y, marker_pt.Z + 1.0),
        )
        try:
            ElementTransformUtils.RotateElement(
                doc, marker.Id, axis, angle,
            )
        except Exception as ex:
            _log("RotateElement marker failed: {0}".format(ex))

    elev = marker.CreateElevation(doc, plan_view_id, 0)
    if elev is None:
        return None

    # Naam zetten
    try:
        wall_name = _name(wall) or "Wand"
        wall_id = _id_int(wall)
        new_name = u"Aanzicht {0} - id{1}".format(wall_name, wall_id)
        elev.Name = _unique_view_name(doc, new_name)
    except Exception as ex:
        _log("rename elevation failed: {0}".format(ex))

    # Schaal
    try:
        elev.Scale = scale
    except Exception as ex:
        _log("set scale failed: {0}".format(ex))

    # Template (na rename + scale; template kan scale overschrijven)
    if template_id is not None:
        try:
            elev.ViewTemplateId = template_id
        except Exception as ex:
            _log("apply elevation template failed: {0}".format(ex))

    # Crop op segment+wand-hoogte. Werkt in view-locale coords:
    # X = view RightDirection (langs wand), Y = view UpDirection (omhoog),
    # Z = view ViewDirection (uit view, dus -Z is naar wand).
    try:
        wall_h = _wall_height_ft(wall)
        half_len = seg_len / 2.0 + WALL_MARGIN_FT
        bb = BoundingBoxXYZ()
        bb.Min = XYZ(-half_len, -WALL_MARGIN_FT, -1.0)
        bb.Max = XYZ(half_len, wall_h + WALL_MARGIN_FT, 5.0)
        elev.CropBox = bb
        elev.CropBoxActive = True
        elev.CropBoxVisible = True
    except Exception as ex:
        _log("set elevation cropbox failed: {0}".format(ex))

    return elev


# ---------------------------------------------------------------- plan/rcp

def _unique_view_name(doc, base):
    """Forceer unieke view-name. Append _2, _3, ..."""
    existing = set()
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            existing.add(_name(v))
        except Exception:
            pass
    if base not in existing:
        return base
    i = 2
    while True:
        cand = u"{0} ({1})".format(base, i)
        if cand not in existing:
            return cand
        i += 1


def _find_source_plan_for_level(doc, level_id, view_type):
    """Vind een bestaande ViewPlan op level_id met juiste ViewType."""
    lvl_int = _id_int_from_id(level_id)
    for v in FilteredElementCollector(doc).OfClass(ViewPlan):
        try:
            if v.IsTemplate:
                continue
            if v.ViewType != view_type:
                continue
            if v.GenLevel is None:
                continue
            if _id_int(v.GenLevel) == lvl_int:
                return v
        except Exception:
            continue
    return None


def _create_room_plan(
    doc, room, level_id, view_type, scale, template_id, name_suffix,
):
    src = _find_source_plan_for_level(doc, level_id, view_type)
    if src is None:
        return None
    try:
        new_id = src.Duplicate(ViewDuplicateOption.Duplicate)
    except Exception as ex:
        _log("duplicate plan failed ({0}): {1}".format(view_type, ex))
        return None
    new_view = doc.GetElement(new_id)
    if new_view is None:
        return None

    room_name = _name(room) or "Ruimte"
    try:
        new_view.Name = _unique_view_name(
            doc, u"{0} - {1}".format(room_name, name_suffix),
        )
    except Exception as ex:
        _log("rename plan failed: {0}".format(ex))

    try:
        new_view.Scale = scale
    except Exception as ex:
        _log("set plan scale failed: {0}".format(ex))

    if template_id is not None:
        try:
            new_view.ViewTemplateId = template_id
        except Exception as ex:
            _log("apply plan template failed: {0}".format(ex))

    # Crop op room bounding box
    bb = room.get_BoundingBox(None)
    if bb is not None:
        try:
            crop = BoundingBoxXYZ()
            crop.Min = XYZ(
                bb.Min.X - ROOM_MARGIN_FT,
                bb.Min.Y - ROOM_MARGIN_FT,
                bb.Min.Z,
            )
            crop.Max = XYZ(
                bb.Max.X + ROOM_MARGIN_FT,
                bb.Max.Y + ROOM_MARGIN_FT,
                bb.Max.Z,
            )
            new_view.CropBox = crop
            new_view.CropBoxActive = True
            new_view.CropBoxVisible = True
        except Exception as ex:
            _log("set plan cropbox failed: {0}".format(ex))

    return new_view


# ---------------------------------------------------------------- sheet

def _create_sheet(doc, titleblock_type_id, number, name):
    try:
        if not titleblock_type_id:
            sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
        else:
            sheet = ViewSheet.Create(doc, titleblock_type_id)
    except Exception as ex:
        _log("ViewSheet.Create failed: {0}".format(ex))
        return None
    try:
        sheet.SheetNumber = number
    except Exception as ex:
        _log("set sheet number '{0}' failed: {1}".format(number, ex))
    try:
        sheet.Name = name
    except Exception as ex:
        _log("set sheet name '{0}' failed: {1}".format(name, ex))
    return sheet


def _viewport_size(viewport):
    """Width, Height van viewport outline (ft, sheet-coords)."""
    try:
        outline = viewport.GetBoxOutline()
        mn = outline.MinimumPoint
        mx = outline.MaximumPoint
        return (mx.X - mn.X, mx.Y - mn.Y)
    except Exception:
        return (1.0, 1.0)


def _place_views_on_sheet(doc, sheet, views, margin_ft=0.15):
    """Plaats views in een grid op de sheet, met margins.

    Strategie: simpele rij-gebaseerde packer. Probeert binnen sheet-
    grenzen te blijven; valt terug op overlap als sheet te klein is.
    """
    # Sheet-bbox via outline van titleblock (indien aanwezig).
    sheet_w = 30.0
    sheet_h = 22.0  # default A1 landscape ~ 33x23 inch -> ft, ruim
    try:
        # Probeer een TitleBlock instance te vinden voor bbox
        col = FilteredElementCollector(doc, sheet.Id) \
            .OfCategory(BuiltInCategory.OST_TitleBlocks) \
            .WhereElementIsNotElementType()
        for tb in col:
            bb = tb.get_BoundingBox(sheet)
            if bb is not None:
                sheet_w = bb.Max.X - bb.Min.X
                sheet_h = bb.Max.Y - bb.Min.Y
                break
    except Exception:
        pass

    placed = []
    failed = []
    cursor_x = margin_ft
    cursor_y = sheet_h - margin_ft
    row_h = 0.0

    for view in views:
        if view is None:
            continue
        # Tijdelijk midden plaatsen om viewport size te bepalen
        center = XYZ(cursor_x + 1.0, cursor_y - 1.0, 0.0)
        try:
            vp = Viewport.Create(doc, sheet.Id, view.Id, center)
        except Exception as ex:
            _log("Viewport.Create failed voor view '{0}': {1}".format(
                _name(view), ex,
            ))
            failed.append(view)
            continue
        w, h = _viewport_size(vp)
        # Nieuw row als view niet meer past horizontaal
        if cursor_x + w > sheet_w - margin_ft and cursor_x > margin_ft:
            cursor_x = margin_ft
            cursor_y -= row_h + margin_ft
            row_h = 0.0
        # Eindpositie: midden van viewport
        new_center = XYZ(
            cursor_x + w / 2.0,
            cursor_y - h / 2.0,
            0.0,
        )
        try:
            # Verplaats viewport
            delta = XYZ(
                new_center.X - center.X,
                new_center.Y - center.Y,
                0.0,
            )
            ElementTransformUtils.MoveElement(doc, vp.Id, delta)
        except Exception as ex:
            _log("move viewport failed: {0}".format(ex))

        placed.append(vp)
        cursor_x += w + margin_ft
        if h > row_h:
            row_h = h

    return placed, failed


# ---------------------------------------------------------------- run

def run():
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()
    output.print_md("## Ruimte Sheet")
    output.print_md(
        "*Debug-log: `{0}`*".format(_EARLY_LOG_FILE)
    )

    # 1. Ruimte selecteren
    room = _pick_room(uidoc)
    if room is None:
        output.print_md("*Geen ruimte geselecteerd.*")
        return

    room_name = _name(room) or "Ruimte"
    room_number = ""
    try:
        p = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        if p and p.HasValue:
            room_number = p.AsString() or ""
    except Exception:
        pass
    output.print_md(
        "Ruimte: **{0}** (nr {1})".format(room_name, room_number)
    )

    # 2. Wanden ophalen
    wall_segs = _get_wall_segments(room)
    output.print_md(
        "Wand-segmenten gevonden: **{0}**".format(len(wall_segs))
    )
    if not wall_segs:
        forms.alert(
            "Geen wand-segmenten in de boundary van deze ruimte.",
            title="Geen wanden",
        )
        return

    # 3. UI - templates + titleblock + sheet info
    elev_tpls = _collect_templates(
        doc, set([ViewType.Elevation]),
    )
    plan_tpls = _collect_templates(
        doc, set([ViewType.FloorPlan]),
    )
    rcp_tpls = _collect_templates(
        doc, set([ViewType.CeilingPlan]),
    )
    titleblocks = _collect_titleblocks(doc)

    output.print_md(
        "Templates beschikbaar - elevation:{0} floor:{1} ceiling:{2}"
        .format(len(elev_tpls), len(plan_tpls), len(rcp_tpls))
    )

    tpl_elev = _select_optional(
        "Kies template voor AANZICHTEN", elev_tpls, _name,
    )
    tpl_plan = _select_optional(
        "Kies template voor PLATTEGROND", plan_tpls, _name,
    )
    tpl_rcp = _select_optional(
        "Kies template voor PLAFOND", rcp_tpls, _name,
    )

    tb_id = ElementId.InvalidElementId
    if titleblocks:
        tb = _select_required(
            "Kies titleblock voor sheet", titleblocks, _tb_display,
        )
        if tb is None:
            output.print_md("*Geen titleblock gekozen - afgebroken.*")
            return
        tb_id = tb.Id

    # Sheet number + name via prompt
    default_number = "fase_601"
    sheet_number = forms.ask_for_string(
        default=default_number,
        prompt="Sheet nummer (format: fase_6xx):",
        title="Ruimte Sheet - nummer",
    )
    if not sheet_number:
        return
    sheet_name = forms.ask_for_string(
        default=room_name,
        prompt="Sheet naam:",
        title="Ruimte Sheet - naam",
    )
    if not sheet_name:
        return

    # 4. ViewFamilyType voor elevations
    vft_elev = _get_view_family_type(doc, ViewFamily.Elevation)
    if vft_elev is None:
        forms.alert(
            "Geen ViewFamilyType voor Elevation gevonden.",
            title="Fout",
        )
        return

    # 5. Level info voor marker Z + bron voor plattegrond/plafond
    level_id = None
    try:
        level_id = room.LevelId
    except Exception:
        pass
    if level_id is None or _id_int_from_id(level_id) <= 0:
        forms.alert(
            "Ruimte heeft geen geldig level.",
            title="Fout",
        )
        return
    level = doc.GetElement(level_id)
    level_elev = 0.0
    try:
        level_elev = level.Elevation
    except Exception:
        pass

    # 6. Vind een floor-plan op dit level om als ownerPlan voor markers
    # te gebruiken (vereist door CreateElevation).
    owner_plan = _find_source_plan_for_level(
        doc, level_id, ViewType.FloorPlan,
    )
    if owner_plan is None:
        forms.alert(
            "Geen FloorPlan gevonden voor level '{0}' - nodig als "
            "owner voor elevation markers.".format(_name(level)),
            title="Fout",
        )
        return

    room_center = _room_center_xy(room)
    tpl_elev_id = tpl_elev.Id if tpl_elev is not None else None
    tpl_plan_id = tpl_plan.Id if tpl_plan is not None else None
    tpl_rcp_id = tpl_rcp.Id if tpl_rcp is not None else None

    # 7. Alles in 1 transaction-group voor rollback bij fout
    tg = TransactionGroup(doc, "Ruimte Sheet - {0}".format(room_name))
    tg.Start()
    elevations = []
    plan_view = None
    rcp_view = None
    sheet = None
    placed_count = 0
    try:
        # 7a. Elevations
        tx = Transaction(doc, "Genereer aanzichten")
        tx.Start()
        try:
            for wall, curve_seg in wall_segs:
                try:
                    ev = _create_wall_elevation(
                        doc, owner_plan.Id, vft_elev.Id,
                        wall, curve_seg, room_center,
                        level_elev, SCALE, tpl_elev_id,
                    )
                    if ev is not None:
                        elevations.append(ev)
                except Exception as ex:
                    _log_exc("elevation failed wall id={0}: {1}".format(
                        _id_int(wall), ex,
                    ))
            tx.Commit()
        except Exception:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
            raise

        # 7b. Plattegrond
        tx = Transaction(doc, "Genereer plattegrond")
        tx.Start()
        try:
            plan_view = _create_room_plan(
                doc, room, level_id, ViewType.FloorPlan,
                SCALE, tpl_plan_id, "plattegrond",
            )
            tx.Commit()
        except Exception:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
            raise

        # 7c. Plafond
        tx = Transaction(doc, "Genereer plafond")
        tx.Start()
        try:
            rcp_view = _create_room_plan(
                doc, room, level_id, ViewType.CeilingPlan,
                SCALE, tpl_rcp_id, "plafond",
            )
            tx.Commit()
        except Exception:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
            raise

        # 7d. Sheet
        tx = Transaction(doc, "Maak sheet")
        tx.Start()
        try:
            sheet = _create_sheet(doc, tb_id, sheet_number, sheet_name)
            tx.Commit()
        except Exception:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
            raise
        if sheet is None:
            raise Exception("Sheet kon niet aangemaakt worden.")

        # 7e. Plaats views op sheet
        views_to_place = []
        if plan_view is not None:
            views_to_place.append(plan_view)
        if rcp_view is not None:
            views_to_place.append(rcp_view)
        views_to_place.extend(elevations)

        tx = Transaction(doc, "Plaats viewports")
        tx.Start()
        try:
            placed, failed = _place_views_on_sheet(doc, sheet, views_to_place)
            placed_count = len(placed)
            tx.Commit()
            if failed:
                _log("viewports failed: {0}".format(
                    [_name(v) for v in failed]
                ))
        except Exception:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
            raise

        tg.Assimilate()
    except Exception as ex:
        if tg.HasStarted() and not tg.HasEnded():
            tg.RollBack()
        _log_exc("Ruimte Sheet failed")
        forms.alert(
            "Genereren mislukt:\n{0}: {1}".format(
                type(ex).__name__, ex,
            ),
            title="Fout",
        )
        return

    # 8. Rapport
    output.print_md("### Resultaat")
    output.print_md("- Aanzichten: **{0}**".format(len(elevations)))
    output.print_md(
        "- Plattegrond: **{0}**".format(
            _name(plan_view) if plan_view else "(geen)"
        )
    )
    output.print_md(
        "- Plafond: **{0}**".format(
            _name(rcp_view) if rcp_view else "(geen)"
        )
    )
    output.print_md(
        "- Sheet: **{0} - {1}**".format(sheet.SheetNumber, sheet.Name)
    )
    output.print_md("- Viewports geplaatst: **{0}**".format(placed_count))

    # Klikbare link naar sheet
    try:
        output.print_md(
            "[Open sheet: {0}]({1})".format(
                sheet.SheetNumber,
                output.linkify(sheet.Id),
            )
        )
    except Exception:
        pass


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        try:
            run()
        except Exception as ex:
            _log_exc("run() crashed")
            try:
                tb_text = _tb_mod.format_exc()
            except Exception:
                tb_text = "<no traceback>"
            forms.alert(
                "Onverwachte fout:\n{0}: {1}\n\n{2}\n\nLog: {3}".format(
                    type(ex).__name__, ex, tb_text, _EARLY_LOG_FILE,
                ),
                title="Fout",
            )
