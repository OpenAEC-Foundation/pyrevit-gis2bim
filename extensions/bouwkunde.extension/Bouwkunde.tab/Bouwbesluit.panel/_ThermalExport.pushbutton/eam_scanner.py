# -*- coding: utf-8 -*-
"""EnergyAnalysisDetailModel scanner voor thermische schil export.

Scant het Revit model via de EAM API (SecondLevelBoundaries) en bouwt
een data-dict op met rooms, constructions, openings en open_connections.

IronPython 2.7 — geen f-strings, geen type hints.
"""
import math
import re

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FailureProcessingResult,
    IFailuresPreprocessor,
    RevitLinkInstance,
    Transaction,
    XYZ,
)
from Autodesk.Revit.DB.Analysis import (
    EnergyAnalysisDetailModel,
    EnergyAnalysisDetailModelOptions,
    EnergyAnalysisDetailModelTier,
    EnergyModelType,
)


class _WarningSwallower(IFailuresPreprocessor):
    """Onderdrukt alle warnings tijdens EAM creatie.

    EnergyAnalysisDetailModel.Create() genereert warnings voor
    onbegrensde rooms, ontbrekende daken, etc. Deze blokkeren de
    transactie als ze niet afgevangen worden.
    """

    def PreprocessFailures(self, failuresAccessor):
        failures = failuresAccessor.GetFailureMessages()
        for failure in failures:
            failuresAccessor.DeleteWarning(failure)
        return FailureProcessingResult.Continue

import sys
import os
import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "..", "lib"))

# Simpele file logger (IronPython logging module buffert niet goed)
_log_dir = os.path.join(os.path.dirname(__file__), "logs")
if not os.path.exists(_log_dir):
    os.makedirs(_log_dir)
_log_path = os.path.join(_log_dir, "eam_scanner.log")
_log_file = None


def _log(msg):
    """Schrijf direct naar logbestand (flush per regel)."""
    global _log_file
    try:
        if _log_file is None:
            _log_file = open(_log_path, "w")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        _log_file.write("{0} {1}\n".format(ts, msg))
        _log_file.flush()
    except Exception:
        pass


def _close_log():
    """Sluit het logbestand."""
    global _log_file
    try:
        if _log_file is not None:
            _log_file.close()
            _log_file = None
    except Exception:
        pass


from warmteverlies.unit_utils import (
    internal_to_sqm,
    internal_to_meters,
    internal_to_mm,
    get_param_value,
)
from warmteverlies.constants import FEET_TO_M, REVIT_LAMBDA_DIVISOR


# =============================================================================
# Compass helpers
# =============================================================================
_COMPASS_DIRS = [
    ("N", 0), ("NE", 45), ("E", 90), ("SE", 135),
    ("S", 180), ("SW", 225), ("W", 270), ("NW", 315),
]


def _angle_to_compass(angle_deg):
    """Converteer een hoek (0=Noord, CW) naar kompasrichting."""
    angle_deg = angle_deg % 360
    best = "N"
    best_diff = 999
    for name, ref in _COMPASS_DIRS:
        diff = abs(angle_deg - ref)
        if diff > 180:
            diff = 360 - diff
        if diff < best_diff:
            best_diff = diff
            best = name
    return best


def _normal_to_compass(normal):
    """Bereken kompasrichting uit een surface normal XYZ vector.

    Revit: X = Oost, Y = Noord. We berekenen de hoek vanaf Noord (Y-as), CW.
    """
    angle_rad = math.atan2(normal.X, normal.Y)
    angle_deg = math.degrees(angle_rad)
    if angle_deg < 0:
        angle_deg += 360
    return _angle_to_compass(angle_deg)


# =============================================================================
# Orientation helpers
# =============================================================================
def _classify_orientation(normal):
    """Classificeer de orientatie van een vlak op basis van z-component.

    Returns:
        tuple: (orientation_str, is_vertical)
    """
    z = normal.Z
    if z > 0.7:
        return "roof", False  # wijst omhoog = dak of plafond
    elif z < -0.7:
        return "floor", False  # wijst omlaag = vloer
    else:
        return "wall", True


def _refine_orientation(orientation, adj_type):
    """Verfijn roof/floor op basis van adjacency.

    Dak naar buiten = roof, dak naar andere ruimte = ceiling.
    """
    if orientation == "roof" and adj_type in ("room", "unheated"):
        return "ceiling"
    return orientation


# =============================================================================
# Polyloop area
# =============================================================================
def _polyloop_area_sqft(polyloop):
    """Bereken oppervlakte van een polyloop in square feet (Shoelace 3D)."""
    pts = list(polyloop.GetPoints())
    if len(pts) < 3:
        return 0.0

    # Bereken normaal via cross product van eerste twee edges
    v1 = XYZ(pts[1].X - pts[0].X, pts[1].Y - pts[0].Y, pts[1].Z - pts[0].Z)
    v2 = XYZ(pts[2].X - pts[0].X, pts[2].Y - pts[0].Y, pts[2].Z - pts[0].Z)
    normal = XYZ(
        v1.Y * v2.Z - v1.Z * v2.Y,
        v1.Z * v2.X - v1.X * v2.Z,
        v1.X * v2.Y - v1.Y * v2.X,
    )
    length = math.sqrt(normal.X ** 2 + normal.Y ** 2 + normal.Z ** 2)
    if length < 1e-10:
        return 0.0
    normal = XYZ(normal.X / length, normal.Y / length, normal.Z / length)

    # 3D polygon area via Newell's method
    area = 0.0
    n = len(pts)
    for i in range(n):
        j = (i + 1) % n
        cross = XYZ(
            pts[i].Y * pts[j].Z - pts[i].Z * pts[j].Y,
            pts[i].Z * pts[j].X - pts[i].X * pts[j].Z,
            pts[i].X * pts[j].Y - pts[i].Y * pts[j].X,
        )
        area += normal.X * cross.X + normal.Y * cross.Y + normal.Z * cross.Z

    return abs(area) / 2.0


def _polyloop_to_2d(polyloop):
    """Converteer polyloop naar 2D punten in meters [[x, y], ...]."""
    pts = list(polyloop.GetPoints())
    result = []
    for p in pts:
        result.append([round(p.X * FEET_TO_M, 4), round(p.Y * FEET_TO_M, 4)])
    return result


# =============================================================================
# Linked document helpers
# =============================================================================
def _get_linked_docs(doc):
    """Verzamel alle geladen linked documents.

    Returns:
        list[Document]: Linked Revit documents
    """
    linked_docs = []
    try:
        collector = (
            FilteredElementCollector(doc)
            .OfClass(RevitLinkInstance)
        )
        for link_inst in collector:
            try:
                linked_doc = link_inst.GetLinkDocument()
                if linked_doc is not None:
                    linked_docs.append(linked_doc)
            except Exception:
                pass
    except Exception:
        pass
    return linked_docs


def _resolve_element(doc, element_id, linked_docs):
    """Resolve een element: eerst in host doc, dan in linked docs.

    Args:
        doc: Host Revit Document
        element_id: ElementId om te resolven
        linked_docs: list[Document] van linked models

    Returns:
        tuple: (element, element_doc) of (None, None)
    """
    # Eerst in host doc
    element = doc.GetElement(element_id)
    if element is not None:
        cat_name = ""
        try:
            cat_name = element.Category.Name if element.Category else "no-cat"
        except Exception:
            cat_name = type(element).__name__
        _log("  _resolve: id={0} -> host doc, cat={1}".format(
            element_id.IntegerValue, cat_name))
        return element, doc

    # Probeer elk linked document
    for linked_doc in linked_docs:
        try:
            element = linked_doc.GetElement(element_id)
            if element is not None:
                cat_name = ""
                try:
                    cat_name = element.Category.Name if element.Category else "no-cat"
                except Exception:
                    cat_name = type(element).__name__
                _log("  _resolve: id={0} -> linked doc '{1}', cat={2}".format(
                    element_id.IntegerValue, linked_doc.Title, cat_name))
                return element, linked_doc
        except Exception:
            pass

    _log("  _resolve: id={0} -> NOT FOUND in host or {1} linked docs".format(
        element_id.IntegerValue, len(linked_docs)))
    return None, None


# =============================================================================
# Compound structure layer extraction
# =============================================================================
def _extract_layers(doc, source_element, element_doc=None):
    """Extraheer laagopbouw uit een Revit element's CompoundStructure.

    Args:
        doc: Revit host Document
        source_element: Revit Element (Wall/Floor/Roof)
        element_doc: Document waarin het element leeft (linked of host).
                     Indien None wordt doc gebruikt.

    Returns:
        list[dict]: Lagen van interieur naar exterieur met material, dikte, type, lambda.
    """
    if source_element is None:
        _log("  _extract_layers: source_element is None")
        return []

    if element_doc is None:
        element_doc = doc

    try:
        type_id = source_element.GetTypeId()
        _log("  _extract_layers: element type={0}, TypeId={1}".format(
            type(source_element).__name__, type_id.IntegerValue if type_id else "None"))
        elem_type = element_doc.GetElement(type_id)
        if elem_type is None:
            _log("  _extract_layers: elem_type is None")
            return []

        _log("  _extract_layers: elem_type={0}".format(type(elem_type).__name__))
        compound = None
        try:
            compound = elem_type.GetCompoundStructure()
        except Exception as ex:
            _log("  _extract_layers: GetCompoundStructure() error: {0}".format(ex))
        if compound is None:
            _log("  _extract_layers: compound is None")
            return []

        raw_layers = compound.GetLayers()
        if not raw_layers or raw_layers.Count == 0:
            _log("  _extract_layers: no layers in compound")
            return []
        _log("  _extract_layers: {0} layers found".format(raw_layers.Count))
    except Exception as ex:
        _log("  _extract_layers: EXCEPTION: {0}".format(ex))
        return []

    result = []
    cumulative_mm = 0.0

    for layer in raw_layers:
        thickness_ft = layer.Width
        thickness_mm = round(internal_to_mm(thickness_ft), 1)

        # Materiaal ophalen
        mat_name = "Onbekend"
        lambda_val = None
        layer_type = "solid"

        mat_id = layer.MaterialId
        if mat_id is not None and mat_id != ElementId.InvalidElementId:
            material = element_doc.GetElement(mat_id)
            if material is not None:
                mat_name = material.Name or "Onbekend"
                lambda_val = _get_lambda(material)

        # Air gap detectie via MaterialFunctionAssignment
        try:
            func_assignment = layer.Function
            # MaterialFunctionAssignment enum: 0=Structure, 1=Substrate,
            # 2=Insulation, 3=Finish1, 4=Finish2, 5=Membrane,
            # 6=StructuralDeck, 7=ThermalAirGap (indien beschikbaar)
            func_name = str(func_assignment)
            if "Air" in func_name or "Thermal" in func_name:
                layer_type = "air_gap"
        except Exception:
            pass

        layer_dict = {
            "material": mat_name,
            "thickness_mm": thickness_mm,
            "distance_from_interior_mm": round(cumulative_mm, 1),
            "type": layer_type,
        }
        if lambda_val is not None and lambda_val > 0:
            layer_dict["lambda"] = round(lambda_val, 4)

        result.append(layer_dict)
        cumulative_mm += thickness_mm

    return result


def _get_lambda(material):
    """Haal thermische geleidbaarheid (lambda) uit een Revit Material."""
    try:
        thermal_asset_id = material.ThermalAssetId
        if (thermal_asset_id is None
                or thermal_asset_id == ElementId.InvalidElementId):
            return None

        doc = material.Document
        prop_set = doc.GetElement(thermal_asset_id)
        if prop_set is None:
            return None

        thermal_asset = prop_set.GetThermalAsset()
        if thermal_asset is None:
            return None

        # Revit ThermalConductivity is in BTU*in/(hr*ft2*degF)
        # Conversie naar W/(m*K): / 6.93347
        raw = thermal_asset.ThermalConductivity
        if raw is None or raw <= 0:
            return None
        return raw / REVIT_LAMBDA_DIVISOR
    except Exception:
        return None


# =============================================================================
# Fallback: layer lookup by type name
# =============================================================================
_type_layers_cache = {}  # type_name -> layers list


def _find_layers_by_type_name(doc, type_name, orientation, linked_docs):
    """Zoek laagopbouw via wandtype naam als OriginatingElementId faalt.

    Doorzoekt alle Wall/Floor/Roof types in host + linked docs.
    Resultaten worden gecacht per type_name.
    """
    if type_name in _type_layers_cache:
        return _type_layers_cache[type_name]

    # Bepaal welke categorieën te doorzoeken
    categories = []
    if orientation == "wall":
        categories.append(BuiltInCategory.OST_Walls)
    elif orientation == "floor":
        categories.append(BuiltInCategory.OST_Floors)
    elif orientation in ("roof", "ceiling"):
        categories.append(BuiltInCategory.OST_Roofs)
        categories.append(BuiltInCategory.OST_Floors)
    else:
        categories.append(BuiltInCategory.OST_Walls)
        categories.append(BuiltInCategory.OST_Floors)
        categories.append(BuiltInCategory.OST_Roofs)

    all_docs = [doc] + (linked_docs or [])

    for search_doc in all_docs:
        for cat in categories:
            try:
                collector = (
                    FilteredElementCollector(search_doc)
                    .OfCategory(cat)
                    .WhereElementIsElementType()
                )
                for elem_type in collector:
                    try:
                        tn_param = elem_type.get_Parameter(
                            BuiltInParameter.ALL_MODEL_TYPE_NAME
                        )
                        tn = ""
                        if tn_param and tn_param.HasValue:
                            tn = tn_param.AsString() or ""

                        # Match: exacte naam of type_name bevat tn
                        if tn and (tn == type_name or tn in type_name
                                   or type_name in tn):
                            # Maak tijdelijk element-achtig object
                            layers = _extract_layers_from_type(
                                search_doc, elem_type
                            )
                            if layers:
                                _log("  Fallback match: '{0}' -> {1} layers"
                                     " (via {2})".format(
                                         type_name, len(layers),
                                         search_doc.Title))
                                _type_layers_cache[type_name] = layers
                                return layers
                    except Exception:
                        pass
            except Exception:
                pass

    _type_layers_cache[type_name] = []
    return []


def _extract_layers_from_type(element_doc, elem_type):
    """Extraheer laagopbouw direct uit een WallType/FloorType/RoofType."""
    try:
        compound = elem_type.GetCompoundStructure()
        if compound is None:
            return []

        raw_layers = compound.GetLayers()
        if not raw_layers or raw_layers.Count == 0:
            return []
    except Exception:
        return []

    result = []
    cumulative_mm = 0.0

    for layer in raw_layers:
        thickness_ft = layer.Width
        thickness_mm = round(internal_to_mm(thickness_ft), 1)

        mat_name = "Onbekend"
        lambda_val = None
        layer_type = "solid"

        mat_id = layer.MaterialId
        if mat_id is not None and mat_id != ElementId.InvalidElementId:
            material = element_doc.GetElement(mat_id)
            if material is not None:
                mat_name = material.Name or "Onbekend"
                lambda_val = _get_lambda(material)

        try:
            func_name = str(layer.Function)
            if "Air" in func_name or "Thermal" in func_name:
                layer_type = "air_gap"
        except Exception:
            pass

        layer_dict = {
            "material": mat_name,
            "thickness_mm": thickness_mm,
            "distance_from_interior_mm": round(cumulative_mm, 1),
            "type": layer_type,
        }
        if lambda_val is not None and lambda_val > 0:
            layer_dict["lambda"] = round(lambda_val, 4)

        result.append(layer_dict)
        cumulative_mm += thickness_mm

    return result


# =============================================================================
# Opening extraction from analytical openings
# =============================================================================
def _extract_eam_opening(doc, opening, constr_id, opening_idx, linked_docs=None):
    """Extraheer opening data uit een EnergyAnalysisOpening.

    Args:
        doc: Revit host Document
        opening: EnergyAnalysisOpening
        constr_id: Construction ID string
        opening_idx: Volgnummer
        linked_docs: list[Document] van linked models

    Returns:
        dict conform thermal-import schema Opening definitie
    """
    if linked_docs is None:
        linked_docs = []
    opening_dict = {
        "id": "opening-{0}".format(opening_idx),
        "construction_id": constr_id,
    }

    # Type bepalen
    open_type = str(opening.OpeningType)
    if "Window" in open_type or "Skylight" in open_type:
        opening_dict["type"] = "window"
    elif "Door" in open_type:
        opening_dict["type"] = "door"
    else:
        opening_dict["type"] = "window"  # default

    # Afmetingen uit polyloop
    try:
        polyloop = opening.GetPolyloop()
        pts = list(polyloop.GetPoints())
        if len(pts) >= 3:
            # Bounding box benadering
            xs = [p.X for p in pts]
            ys = [p.Y for p in pts]
            zs = [p.Z for p in pts]
            dx = max(xs) - min(xs)
            dy = max(ys) - min(ys)
            dz = max(zs) - min(zs)

            # Breedte = max van dx, dy; hoogte = dz (voor wanden)
            width_ft = max(dx, dy)
            height_ft = dz if dz > 0.01 else max(dx, dy)
            if dz < 0.01:
                # Horizontale opening (daklicht) — gebruik dx en dy
                width_ft = dx
                height_ft = dy

            opening_dict["width_mm"] = round(width_ft * FEET_TO_M * 1000, 0)
            opening_dict["height_mm"] = round(height_ft * FEET_TO_M * 1000, 0)
        else:
            opening_dict["width_mm"] = 1000
            opening_dict["height_mm"] = 1000
    except Exception:
        opening_dict["width_mm"] = 1000
        opening_dict["height_mm"] = 1000

    # Sill height
    try:
        sill_param = None
        origin_elem, origin_doc = _resolve_element(
            doc, opening.OriginatingElementId, linked_docs
        )
        if origin_elem is not None:
            sill_param = get_param_value(origin_elem, "Sill Height")
            if sill_param is None:
                sill_param = get_param_value(origin_elem, "Rough Sill Height")
        if sill_param is not None and sill_param > 0:
            opening_dict["sill_height_mm"] = round(
                internal_to_mm(sill_param), 0
            )
    except Exception:
        pass

    # Revit element reference
    try:
        orig_id = opening.OriginatingElementId
        if orig_id is not None and orig_id != ElementId.InvalidElementId:
            opening_dict["revit_element_id"] = orig_id.IntegerValue
            orig_elem, orig_doc = _resolve_element(
                doc, orig_id, linked_docs
            )
            if orig_elem is not None:
                try:
                    elem_type = orig_doc.GetElement(
                        orig_elem.GetTypeId()
                    )
                    if elem_type is not None:
                        family = getattr(elem_type, "FamilyName", "")
                        tname = elem_type.get_Parameter(
                            BuiltInParameter.ALL_MODEL_TYPE_NAME
                        )
                        if tname and tname.HasValue:
                            opening_dict["revit_type_name"] = "{0}: {1}".format(
                                family, tname.AsString()
                            )
                        elif family:
                            opening_dict["revit_type_name"] = family
                except Exception:
                    pass
    except Exception:
        pass

    return opening_dict


# =============================================================================
# Outside room detection
# =============================================================================
_OUTSIDE_PATTERNS = [
    re.compile(r"^buiten$", re.IGNORECASE),
    re.compile(r"^room$", re.IGNORECASE),
    re.compile(r"^terras", re.IGNORECASE),
    re.compile(r"^balkon", re.IGNORECASE),
    re.compile(r"^tuin\b", re.IGNORECASE),
    re.compile(r"^dakterras", re.IGNORECASE),
    re.compile(r"^galerij", re.IGNORECASE),
]


def _is_likely_outside(name, area_m2):
    """Check of een ruimte waarschijnlijk buiten/onverwarmd is."""
    if not name:
        return False
    for pat in _OUTSIDE_PATTERNS:
        if pat.search(name):
            return True
    # Extreem grote ruimtes (>200m2) met generieke namen
    if area_m2 > 200 and name.strip().lower() in ("buiten", "room", "space"):
        return True
    return False


# =============================================================================
# Room mapping: EnergyAnalysisSpace -> Revit Room
# =============================================================================
def _build_room_map(doc, eam, output):
    """Bouw een mapping van EnergyAnalysisSpace ID -> room data dict.

    Returns:
        tuple: (room_map, rooms_list, room_id_counter)
            room_map: dict van space_id -> room dict
            rooms_list: list van alle room dicts
    """
    rooms_list = []
    room_map = {}  # space_id -> room dict
    room_counter = 0

    spaces = eam.GetAnalyticalSpaces()
    if output:
        output.print_md("Analytische spaces gevonden: **{0}**".format(
            len(spaces) if spaces else 0
        ))

    if not spaces:
        return room_map, rooms_list

    for space in spaces:
        # GetAnalyticalSpaces() retourneert EnergyAnalysisSpace objecten
        # niet ElementIds — gebruik .Id om de ElementId te krijgen
        space_elem = doc.GetElement(space.Id)
        if space_elem is None:
            space_elem = space  # gebruik het object zelf als fallback

        space_name = "Space"
        try:
            space_name = space_elem.SpaceName or "Space"
        except Exception:
            pass
        space_int_id = space.Id.IntegerValue

        # Probeer gekoppelde Revit Room te vinden via CADObjectUniqueId
        revit_room = None
        revit_id = None
        try:
            cad_uid = space_elem.CADObjectUniqueId
            if cad_uid:
                # CADObjectUniqueId is vaak de UniqueId van het Room element
                collector = (
                    FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                )
                for room in collector:
                    if room.UniqueId == cad_uid:
                        revit_room = room
                        revit_id = room.Id.IntegerValue
                        break
        except Exception:
            pass

        # Room data opbouwen
        room_id = "room-{0}".format(room_counter)
        room_counter += 1

        area_m2 = 0.0
        height_m = 0.0
        level_name = ""
        boundary_polygon = None

        if revit_room is not None:
            area_m2 = internal_to_sqm(revit_room.Area) if revit_room.Area > 0 else 0.0
            from warmteverlies.unit_utils import get_room_height
            height_m = get_room_height(revit_room)

            level = doc.GetElement(revit_room.LevelId)
            level_name = level.Name if level else ""

            name = ""
            name_param = revit_room.get_Parameter(BuiltInParameter.ROOM_NAME)
            if name_param and name_param.HasValue:
                name = name_param.AsString() or ""
            if name:
                space_name = name
        else:
            # Fallback: probeer area uit de space zelf
            try:
                area_m2 = internal_to_sqm(space_elem.Area) if space_elem.Area > 0 else 0.0
            except Exception:
                pass

        # Outside detectie
        room_type = "heated"
        if _is_likely_outside(space_name, area_m2):
            room_type = "outside"
            _log("Room '{0}' gemarkeerd als outside".format(space_name))

        room_dict = {
            "id": room_id,
            "name": space_name,
            "type": room_type,
            "level": level_name,
            "area_m2": round(area_m2, 2),
            "height_m": round(height_m, 2),
            "volume_m3": round(area_m2 * height_m, 2),
        }
        _log("Room: {0} ({1}) type={2}, area={3:.1f} m2, revit_id={4}".format(
            space_name, room_id, room_type, area_m2, revit_id))
        if revit_id is not None:
            room_dict["revit_id"] = revit_id

        rooms_list.append(room_dict)
        room_map[space_int_id] = room_dict

    return room_map, rooms_list


# =============================================================================
# Pseudo-ruimtes
# =============================================================================
_PSEUDO_OUTSIDE = {
    "id": "room-outside",
    "name": "Buiten",
    "type": "outside",
    "level": "",
    "area_m2": 0.0,
    "height_m": 0.0,
    "volume_m3": 0.0,
}

_PSEUDO_GROUND = {
    "id": "room-ground",
    "name": "Grond",
    "type": "ground",
    "level": "",
    "area_m2": 0.0,
    "height_m": 0.0,
    "volume_m3": 0.0,
}


# =============================================================================
# Construction consolidation
# =============================================================================
def _consolidate_constructions(constructions, openings):
    """Consolideer individuele EAM surfaces naar constructie-typen.

    Groepeert op (room_a, room_b, revit_type_name, orientation) en
    sommeert oppervlaktes. Openings worden mee-verplaatst.

    Args:
        constructions: list[dict] van individuele surface constructies
        openings: list[dict] van openings (met construction_id referentie)

    Returns:
        tuple: (consolidated_constructions, updated_openings)
    """
    if not constructions:
        return constructions, openings

    # Bouw lookup: oude constr_id -> openings
    opening_lookup = {}
    for o in openings:
        cid = o.get("construction_id", "")
        if cid not in opening_lookup:
            opening_lookup[cid] = []
        opening_lookup[cid].append(o)

    # Groepeer op sleutel
    groups = {}  # key -> list[constr_dict]
    for c in constructions:
        key = (
            c.get("room_a", ""),
            c.get("room_b", ""),
            c.get("revit_type_name", ""),
            c.get("orientation", ""),
            c.get("compass", ""),
        )
        if key not in groups:
            groups[key] = []
        groups[key].append(c)

    # Consolideer groepen
    consolidated = []
    new_openings = []
    constr_idx = 0
    opening_idx = 0

    for key, group in groups.items():
        new_id = "constr-{0}".format(constr_idx)
        constr_idx += 1

        # Oppervlakte sommeren
        total_area = sum(c.get("gross_area_m2", 0.0) for c in group)

        # Neem layers/metadata van het eerste element met layers
        base = group[0]
        for c in group:
            if c.get("layers"):
                base = c
                break

        merged = {
            "id": new_id,
            "room_a": base.get("room_a", ""),
            "room_b": base.get("room_b", ""),
            "orientation": base.get("orientation", ""),
            "gross_area_m2": round(total_area, 2),
        }
        if base.get("compass"):
            merged["compass"] = base["compass"]
        if base.get("layers"):
            merged["layers"] = base["layers"]
        if base.get("revit_element_id") is not None:
            merged["revit_element_id"] = base["revit_element_id"]
        if base.get("revit_type_name"):
            merged["revit_type_name"] = base["revit_type_name"]
        if len(group) > 1:
            merged["surface_count"] = len(group)

        consolidated.append(merged)

        # Openings van alle surfaces in de groep verplaatsen
        for c in group:
            old_id = c.get("id", "")
            for o in opening_lookup.get(old_id, []):
                new_o = dict(o)
                new_o["id"] = "opening-{0}".format(opening_idx)
                new_o["construction_id"] = new_id
                new_openings.append(new_o)
                opening_idx += 1

    _log("Consolidatie: {0} surfaces -> {1} constructie-typen".format(
        len(constructions), len(consolidated)))
    return consolidated, new_openings


# =============================================================================
# Main scan function
# =============================================================================
def scan_thermal_shell(doc, output=None):
    """Scan de thermische schil via de EnergyAnalysisDetailModel API.

    Args:
        doc: Revit Document
        output: pyrevit script output (optioneel, voor voortgangsmeldingen)

    Returns:
        dict met keys: rooms, constructions, openings, open_connections
        of None bij fout
    """
    _log("=== EAM scan gestart ===")
    if output:
        output.print_md("### EAM Scanner")
        output.print_md("EnergyAnalysisDetailModel aanmaken...")

    # ------------------------------------------------------------------
    # 1. EAM aanmaken
    # ------------------------------------------------------------------
    options = EnergyAnalysisDetailModelOptions()
    options.Tier = EnergyAnalysisDetailModelTier.SecondLevelBoundaries
    options.EnergyModelType = EnergyModelType.SpatialElement

    eam = None
    trans = Transaction(doc, "Create EAM for Thermal Export")
    try:
        failure_opts = trans.GetFailureHandlingOptions()
        failure_opts.SetFailuresPreprocessor(_WarningSwallower())
        trans.SetFailureHandlingOptions(failure_opts)
        trans.Start()
        eam = EnergyAnalysisDetailModel.Create(doc, options)
        trans.Commit()
    except Exception as ex:
        if trans.HasStarted():
            trans.RollBack()
        if output:
            output.print_md("**FOUT:** Kan EAM niet aanmaken: {0}".format(str(ex)))
        return None

    if eam is None:
        if output:
            output.print_md("**FOUT:** EAM is None na Create()")
        return None

    # ------------------------------------------------------------------
    # 2. Rooms opbouwen vanuit analytische spaces
    # ------------------------------------------------------------------
    room_map, rooms_list = _build_room_map(doc, eam, output)

    if not rooms_list:
        if output:
            output.print_md("**Waarschuwing:** Geen analytische spaces gevonden.")

    # Pseudo-ruimtes toevoegen
    has_outside = False
    has_ground = False

    # Linked documents ophalen voor element resolution
    linked_docs = _get_linked_docs(doc)
    if output and linked_docs:
        output.print_md("Linked models: **{0}**".format(len(linked_docs)))

    # ------------------------------------------------------------------
    # 3. Surfaces doorlopen -> constructions + openings
    # ------------------------------------------------------------------
    constructions = []
    openings = []
    constr_counter = 0
    opening_counter = 0

    surfaces = eam.GetAnalyticalSurfaces()
    if output:
        output.print_md("Analytische surfaces: **{0}**".format(
            len(surfaces) if surfaces else 0
        ))

    if surfaces:
        for surface in surfaces:
            # GetAnalyticalSurfaces() retourneert EnergyAnalysisSurface
            # objecten — direct bruikbaar, geen doc.GetElement() nodig

            # Polyloop en area
            try:
                polyloop = surface.GetPolyloop()
                area_sqft = _polyloop_area_sqft(polyloop)
                area_m2 = area_sqft * 0.09290304
            except Exception:
                continue

            if area_m2 < 0.25:
                continue

            # Normaal vector
            try:
                normal = surface.Normal
            except Exception:
                normal = XYZ(0, 0, 1)

            # Orientatie
            orientation, is_vertical = _classify_orientation(normal)

            # Room A (de space waar dit surface bij hoort)
            room_a_dict = None
            try:
                analytical_space = surface.GetAnalyticalSpace()
                if analytical_space is not None:
                    # Kan EnergyAnalysisSpace object of ElementId zijn
                    try:
                        space_int_id = analytical_space.Id.IntegerValue
                    except AttributeError:
                        space_int_id = analytical_space.IntegerValue
                    room_a_dict = room_map.get(space_int_id)
            except Exception:
                pass

            if room_a_dict is None:
                # Surface zonder space — overslaan
                continue

            # Room B (adjacent space)
            room_b_dict = None
            adj_type = "outside"  # default
            try:
                adj_space = surface.GetAdjacentAnalyticalSpace()
                if adj_space is not None:
                    try:
                        adj_int_id = adj_space.Id.IntegerValue
                    except AttributeError:
                        adj_int_id = adj_space.IntegerValue
                    if adj_int_id > 0:
                        room_b_dict = room_map.get(adj_int_id)
                        if room_b_dict is not None:
                            adj_type = "room"
            except Exception:
                pass

            if room_b_dict is None:
                # Geen adjacent space: buiten of grond
                if orientation == "floor":
                    room_b_dict = _PSEUDO_GROUND
                    adj_type = "ground"
                    has_ground = True
                else:
                    room_b_dict = _PSEUDO_OUTSIDE
                    adj_type = "outside"
                    has_outside = True

            # Verfijn orientatie
            orientation = _refine_orientation(orientation, adj_type)

            # Compass richting voor wanden
            compass = None
            if is_vertical:
                compass = _normal_to_compass(normal)

            # Source element voor laagopbouw (host of linked doc)
            source_element = None
            element_doc = doc
            revit_element_id = None
            revit_type_name = None

            # Methode 1: OriginatingElementId (gooit vaak exception)
            try:
                orig_id = surface.OriginatingElementId
                _log("Surface {0}: OriginatingElementId={1}".format(
                    constr_counter,
                    orig_id.IntegerValue if orig_id else "None"))
                if orig_id is not None and orig_id != ElementId.InvalidElementId:
                    source_element, element_doc = _resolve_element(
                        doc, orig_id, linked_docs
                    )
                    revit_element_id = orig_id.IntegerValue
            except Exception as ex:
                _log("Surface {0}: OriginatingElementId EXCEPTION: {1}".format(
                    constr_counter, ex))

            # Methode 2: OriginatingElementDescription als fallback naam
            if revit_type_name is None:
                try:
                    desc = surface.OriginatingElementDescription
                    if desc:
                        revit_type_name = desc
                        _log("Surface {0}: Description='{1}'".format(
                            constr_counter, desc))
                except Exception:
                    pass

            # Methode 3: SurfaceName als laatste fallback
            if revit_type_name is None:
                try:
                    sname = surface.SurfaceName
                    if sname:
                        revit_type_name = sname
                        _log("Surface {0}: SurfaceName='{1}'".format(
                            constr_counter, sname))
                except Exception:
                    pass

            # Type naam uit element (als OriginatingElementId werkte)
            if source_element is not None and revit_type_name is None:
                try:
                    et = element_doc.GetElement(
                        source_element.GetTypeId()
                    )
                    if et is not None:
                        tn = et.get_Parameter(
                            BuiltInParameter.ALL_MODEL_TYPE_NAME
                        )
                        if tn and tn.HasValue:
                            revit_type_name = tn.AsString()
                except Exception:
                    pass

            # Laagopbouw extraheren
            layers = _extract_layers(doc, source_element, element_doc)

            # Fallback: als geen layers via element, zoek op type naam
            if not layers and revit_type_name:
                layers = _find_layers_by_type_name(
                    doc, revit_type_name, orientation, linked_docs
                )
            is_linked = element_doc is not doc
            _log("Surface {0}: orient={1}, area={2:.2f} m2, elem={3}, "
                 "linked={4}, layers={5}, type={6}".format(
                     constr_counter, orientation, area_m2,
                     revit_element_id, is_linked, len(layers),
                     revit_type_name or "?"))

            # Construction dict
            constr_id = "constr-{0}".format(constr_counter)
            constr_counter += 1

            constr_dict = {
                "id": constr_id,
                "room_a": room_a_dict["id"],
                "room_b": room_b_dict["id"],
                "orientation": orientation,
                "gross_area_m2": round(area_m2, 2),
            }
            if compass:
                constr_dict["compass"] = compass
            if layers:
                constr_dict["layers"] = layers
            if revit_element_id is not None:
                constr_dict["revit_element_id"] = revit_element_id
            if revit_type_name:
                constr_dict["revit_type_name"] = revit_type_name

            constructions.append(constr_dict)

            # Openings van dit surface
            try:
                surface_openings = surface.GetAnalyticalOpenings()
                if surface_openings:
                    for open_obj in surface_openings:
                        # GetAnalyticalOpenings() retourneert
                        # EnergyAnalysisOpening objecten
                        opening_dict = _extract_eam_opening(
                            doc, open_obj, constr_id, opening_counter,
                            linked_docs
                        )
                        openings.append(opening_dict)
                        opening_counter += 1
            except Exception:
                pass

    # ------------------------------------------------------------------
    # 4. Pseudo-ruimtes toevoegen als ze gebruikt worden
    # ------------------------------------------------------------------
    if has_outside:
        rooms_list.append(dict(_PSEUDO_OUTSIDE))
    if has_ground:
        rooms_list.append(dict(_PSEUDO_GROUND))

    # ------------------------------------------------------------------
    # 5. EAM NIET opruimen — bewust laten staan voor visuele inspectie
    # ------------------------------------------------------------------
    # TODO: Zet onderstaand blok weer aan na debugging
    # try:
    #     trans2 = Transaction(doc, "Delete EAM")
    #     failure_opts2 = trans2.GetFailureHandlingOptions()
    #     failure_opts2.SetFailuresPreprocessor(_WarningSwallower())
    #     trans2.SetFailureHandlingOptions(failure_opts2)
    #     trans2.Start()
    #     EnergyAnalysisDetailModel.Destroy(doc)
    #     trans2.Commit()
    # except Exception:
    #     if trans2.HasStarted():
    #         trans2.RollBack()
    _log("EAM NIET opgeruimd — blijft zichtbaar voor inspectie")

    # ------------------------------------------------------------------
    # 6. Consolideer constructies (305 surfaces -> ~10 typen)
    # ------------------------------------------------------------------
    _log("Pre-consolidatie: {0} surfaces, {1} openings".format(
        len(constructions), len(openings)))
    constructions, openings = _consolidate_constructions(
        constructions, openings
    )

    # ------------------------------------------------------------------
    # 7. Resultaat
    # ------------------------------------------------------------------
    result = {
        "rooms": rooms_list,
        "constructions": constructions,
        "openings": openings,
        "open_connections": [],
    }

    # Log samenvatting
    with_layers = sum(1 for c in constructions if c.get("layers"))
    _log("Scan compleet: {0} rooms, {1} constructies ({2} met layers), "
         "{3} openings".format(
             len(rooms_list), len(constructions), with_layers,
             len(openings)))

    if output:
        output.print_md(
            "Scan compleet: **{0}** rooms, **{1}** constructies "
            "(**{2}** met laagopbouw), **{3}** openings".format(
                len(rooms_list), len(constructions),
                with_layers, len(openings)
            )
        )
        output.print_md("Log: `{0}`".format(_log_path))

    _close_log()
    return result
