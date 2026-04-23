# -*- coding: utf-8 -*-
"""Raycast boundary scanner voor warmteverlies export.

Scant room boundaries via ReferenceIntersector voor cross-model
laagopbouw detectie, inclusief spouw en terminal conditie.

IronPython 2.7 -- geen f-strings, geen type hints, geen walrus operator.
"""
import math

from Autodesk.Revit.DB import (
    XYZ,
    Line,
    ReferenceIntersector,
    FindReferenceTarget,
    SpatialElementGeometryCalculator,
    SpatialElementBoundaryOptions,
    SpatialElementBoundaryLocation,
    ElementId,
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    RevitLinkInstance,
    View3D,
    UV,
)

from warmteverlies.constants import (
    FEET_TO_M,
    RAY_HEIGHT_STEP_M,
    RAY_MAX_DIST_M,
    MIN_CAVITY_MM,
    MAX_CAVITY_MM,
    CONSTRUCTION_CATEGORIES,
    OPENING_CATEGORIES,
    IGNORE_CATEGORIES,
    WATER_MATERIAL_KEYWORDS,
    GROUND_MATERIAL_KEYWORDS,
    DEBUG_OPENINGS,
)
from warmteverlies.unit_utils import internal_to_meters, internal_to_mm


# =========================================================================
# Publieke API
# =========================================================================

def scan_room_boundaries(doc, room, view3d, all_rooms):
    """Scan alle grensvlakken van een room via raycasting.

    Gebruikt SpatialElementGeometryCalculator voor face-detectie en
    ReferenceIntersector voor laagopbouw scanning inclusief linked models.

    Args:
        doc: Revit host Document
        room: Revit Room element
        view3d: Revit View3D zonder section box (voor ReferenceIntersector)
        all_rooms: Lijst van room data dicts (voor adjacent detectie)

    Returns:
        dict met keys: constructions, openings, open_connections
    """
    result = {
        "constructions": [],
        "openings": [],
        "open_connections": [],
    }

    try:
        intersector = _create_intersector(doc, view3d)
    except Exception:
        return result

    # Phase resolution: view phase heeft voorrang, want
    # room.CreatedPhaseId retourneert in sommige modellen
    # ElementId.InvalidElementId (-1) terwijl rooms wel in de
    # actieve phase bestaan. Fallbacks: room.CreatedPhaseId,
    # laatste phase in document, anders None.
    phase = _resolve_phase(doc, room, view3d)

    # Room center in feet (Revit internal)
    try:
        room_pt = room.Location.Point
        room_center_xy = XYZ(room_pt.X, room_pt.Y, 0.0)
    except Exception:
        return result

    # Room Z-range
    try:
        room_bb = room.get_BoundingBox(None)
        room_z_min_m = room_bb.Min.Z * FEET_TO_M
        room_z_max_m = room_bb.Max.Z * FEET_TO_M
    except Exception:
        room_z_min_m = 0.0
        room_z_max_m = 2.6

    # Boundary faces via SEGC of fallback
    faces = _get_boundary_faces(doc, room, room_z_min_m, room_z_max_m)

    # Room lookup voor adjacent detectie
    room_lookup = {}
    for rd in all_rooms:
        room_lookup[rd["element_id"]] = rd

    # Bug E.4: room-wide dedupe set voor openings. SEGC splitst een
    # woonboot-wand soms in ~9 sub-faces. Zonder een gedeelde `seen`
    # set zou elke sub-face-call hetzelfde raam opnieuw oppikken, wat
    # tot N-voudige aftrek en negatieve net_area leidt. Een lokale
    # set per room is genoeg -- per room scannen we onafhankelijke
    # boundary faces zonder overlap tussen rooms.
    seen_openings = set()

    for face_info in faces:
        normal = face_info["normal"]
        position_type = face_info["position_type"]
        area_m2 = face_info["area_m2"]
        z_min_m = face_info["z_min_m"]
        z_max_m = face_info["z_max_m"]

        if position_type == "wall":
            _scan_wall_face(
                doc, intersector, room_center_xy, normal,
                z_min_m, z_max_m, area_m2, room, all_rooms,
                result, phase, seen_openings, view3d,
            )
        elif position_type == "floor":
            _scan_horizontal_face(
                doc, intersector, room_center_xy,
                XYZ(0.0, 0.0, -1.0), z_min_m, area_m2,
                "floor", room, all_rooms, result, phase,
            )
        elif position_type == "ceiling":
            _scan_horizontal_face(
                doc, intersector, room_center_xy,
                XYZ(0.0, 0.0, 1.0), z_max_m, area_m2,
                "ceiling", room, all_rooms, result, phase,
            )

    return result


# =========================================================================
# Phase resolution
# =========================================================================

def _resolve_phase(doc, room, view3d):
    """Bepaal de actieve phase voor room-point lookups.

    Volgorde:
    1. View3D.VIEW_PHASE parameter
    2. room.CreatedPhaseId (kan InvalidElementId zijn in sommige
       modellen, dan skippen)
    3. Laatste phase in doc.Phases
    4. None als alles faalt

    Args:
        doc: Revit host Document
        room: Revit Room element
        view3d: Revit View3D element

    Returns:
        Phase object of None
    """
    # 1. View phase
    try:
        phase_param = view3d.get_Parameter(BuiltInParameter.VIEW_PHASE)
        if phase_param is not None:
            view_phase_id = phase_param.AsElementId()
            if (view_phase_id is not None
                    and view_phase_id != ElementId.InvalidElementId):
                view_phase = doc.GetElement(view_phase_id)
                if view_phase is not None:
                    return view_phase
    except Exception:
        pass

    # 2. Room created phase
    try:
        room_phase_id = room.CreatedPhaseId
        if (room_phase_id is not None
                and room_phase_id != ElementId.InvalidElementId):
            room_phase = doc.GetElement(room_phase_id)
            if room_phase is not None:
                return room_phase
    except Exception:
        pass

    # 3. Laatste phase in document
    try:
        if doc.Phases is not None and doc.Phases.Size > 0:
            return doc.Phases.get_Item(doc.Phases.Size - 1)
    except Exception:
        pass

    return None


# =========================================================================
# Intersector setup
# =========================================================================

def _create_intersector(doc, view3d):
    """Maak een ReferenceIntersector aan met linked model support.

    Args:
        doc: Revit host Document
        view3d: View3D element

    Returns:
        ReferenceIntersector geconfigureerd voor linked models
    """
    intersector = ReferenceIntersector(view3d)
    intersector.FindReferencesInRevitLinks = True
    intersector.TargetType = FindReferenceTarget.Face
    return intersector


# =========================================================================
# Boundary face detectie
# =========================================================================

def _get_boundary_faces(doc, room, room_z_min_m, room_z_max_m):
    """Verkrijg boundary faces via SEGC met bbox fallback.

    Args:
        doc: Revit Document
        room: Revit Room element
        room_z_min_m: Room onderkant in meters
        room_z_max_m: Room bovenkant in meters

    Returns:
        list of dicts met normal, position_type, area_m2, z_min_m, z_max_m
    """
    try:
        return _get_faces_from_segc(doc, room)
    except Exception:
        return _get_faces_from_bbox(room, room_z_min_m, room_z_max_m)


def _get_faces_from_segc(doc, room):
    """Verkrijg faces via SpatialElementGeometryCalculator.

    Args:
        doc: Revit Document
        room: Revit Room element

    Returns:
        list of face info dicts
    """
    calculator = SpatialElementGeometryCalculator(doc)
    seg_result = calculator.CalculateSpatialElementGeometry(room)
    spatial_solid = seg_result.GetGeometry()

    faces = []
    for face in spatial_solid.Faces:
        sub_faces = seg_result.GetBoundaryFaceInfo(face)

        # Fallback: horizontale faces (ceiling/floor) zonder bounding
        # element leveren geen sub_faces. Raycast vanuit de room vindt
        # alsnog de bovengelegen/ondergelegen constructie, maar alleen
        # als we de face meegeven. Walls zonder sub-boundary blijven
        # geskipt — die zijn onbetrouwbaar zonder bounding element.
        if not sub_faces or sub_faces.Count == 0:
            try:
                parent_area_m2 = face.Area * FEET_TO_M * FEET_TO_M
            except Exception:
                parent_area_m2 = 0.0

            if parent_area_m2 < 0.05:
                continue

            try:
                parent_normal = _get_face_normal(face)
            except Exception:
                continue

            # Alleen horizontale faces (|Z| > 0.7) zonder bounding
            # element meenemen — verticale wanden zonder bounding
            # element zijn niet betrouwbaar genoeg om te scannen.
            if abs(parent_normal.Z) <= 0.7:
                continue

            parent_position_type = _classify_normal(parent_normal)
            parent_z_min_m, parent_z_max_m = _get_face_z_range(face)

            faces.append({
                "normal": parent_normal,
                "position_type": parent_position_type,
                "area_m2": parent_area_m2,
                "z_min_m": parent_z_min_m,
                "z_max_m": parent_z_max_m,
            })
            continue

        for sub_face in sub_faces:
            sub_geom = sub_face.GetSubface()
            area_sqft = sub_geom.Area
            area_m2 = area_sqft * FEET_TO_M * FEET_TO_M

            if area_m2 < 0.05:
                continue

            normal = _get_face_normal(sub_geom)
            position_type = _classify_normal(normal)
            z_min_m, z_max_m = _get_face_z_range(sub_geom)

            faces.append({
                "normal": normal,
                "position_type": position_type,
                "area_m2": area_m2,
                "z_min_m": z_min_m,
                "z_max_m": z_max_m,
            })

    return faces


def _get_faces_from_bbox(room, room_z_min_m, room_z_max_m):
    """Fallback: genereer faces uit room bounding box.

    Creert 4 cardinale wanden + vloer + plafond.

    Args:
        room: Revit Room element
        room_z_min_m: Room onderkant in meters
        room_z_max_m: Room bovenkant in meters

    Returns:
        list of face info dicts
    """
    faces = []

    try:
        bb = room.get_BoundingBox(None)
        if bb is None:
            return faces
    except Exception:
        return faces

    dx_m = (bb.Max.X - bb.Min.X) * FEET_TO_M
    dy_m = (bb.Max.Y - bb.Min.Y) * FEET_TO_M
    height_m = room_z_max_m - room_z_min_m

    if height_m <= 0:
        height_m = 2.6

    # 4 cardinale richtingen
    cardinals = [
        (XYZ(1.0, 0.0, 0.0), dy_m * height_m),   # Oost
        (XYZ(-1.0, 0.0, 0.0), dy_m * height_m),   # West
        (XYZ(0.0, 1.0, 0.0), dx_m * height_m),    # Noord
        (XYZ(0.0, -1.0, 0.0), dx_m * height_m),   # Zuid
    ]
    for normal, area in cardinals:
        faces.append({
            "normal": normal,
            "position_type": "wall",
            "area_m2": area,
            "z_min_m": room_z_min_m,
            "z_max_m": room_z_max_m,
        })

    # Vloer
    floor_area = dx_m * dy_m
    faces.append({
        "normal": XYZ(0.0, 0.0, -1.0),
        "position_type": "floor",
        "area_m2": floor_area,
        "z_min_m": room_z_min_m,
        "z_max_m": room_z_min_m,
    })

    # Plafond
    faces.append({
        "normal": XYZ(0.0, 0.0, 1.0),
        "position_type": "ceiling",
        "area_m2": floor_area,
        "z_min_m": room_z_max_m,
        "z_max_m": room_z_max_m,
    })

    return faces


# =========================================================================
# Face geometry helpers
# =========================================================================

def _get_face_z_range(face):
    """Bepaal het Z-bereik van een face in meters.

    Loopt alle vertices van de getrianguleerde face.

    Args:
        face: Revit Face object

    Returns:
        tuple: (z_min_m, z_max_m)
    """
    try:
        mesh = face.Triangulate()
        if mesh is None or mesh.NumTriangles == 0:
            return (0.0, 0.0)

        z_min = float("inf")
        z_max = float("-inf")

        for i in range(mesh.NumTriangles):
            triangle = mesh.get_Triangle(i)
            for j in range(3):
                vertex = triangle.get_Vertex(j)
                z_m = vertex.Z * FEET_TO_M
                if z_m < z_min:
                    z_min = z_m
                if z_m > z_max:
                    z_max = z_m

        if z_min == float("inf"):
            return (0.0, 0.0)

        return (z_min, z_max)

    except Exception:
        return (0.0, 0.0)


def _get_face_normal(face):
    """Bepaal de gemiddelde normal van een face.

    Args:
        face: Revit Face object

    Returns:
        XYZ normal vector
    """
    try:
        bbox = face.GetBoundingBox()
        mid_u = (bbox.Min.U + bbox.Max.U) / 2.0
        mid_v = (bbox.Min.V + bbox.Max.V) / 2.0
        normal = face.ComputeNormal(UV(mid_u, mid_v))
        return normal
    except Exception:
        return XYZ(0, 0, 0)


def _classify_normal(normal):
    """Classificeer face normal als wall, floor of ceiling.

    Args:
        normal: XYZ normal vector

    Returns:
        str: "ceiling", "floor" of "wall"
    """
    z = normal.Z
    if z > 0.7:
        return "ceiling"
    elif z < -0.7:
        return "floor"
    return "wall"


# =========================================================================
# Wall face scanning
# =========================================================================

def _scan_wall_face(doc, intersector, room_center_xy, normal,
                    z_min_m, z_max_m, face_area_m2, room,
                    all_rooms, result, phase, seen_openings, view3d):
    """Scan een verticale (wand) face op meerdere hoogtes.

    Cast rays per RAY_HEIGHT_STEP_M hoogte, groepeert in zones en
    voegt constructies en openings toe aan result.

    Args:
        doc: Revit host Document
        intersector: ReferenceIntersector
        room_center_xy: XYZ room center (Z=0)
        normal: XYZ face normal (horizontaal)
        z_min_m: Onderkant face in meters
        z_max_m: Bovenkant face in meters
        face_area_m2: Face oppervlak in m2
        room: Revit Room element
        all_rooms: Lijst van room data dicts
        result: Resultaat dict om aan te vullen
        phase: Revit Phase voor GetRoomAtPoint, of None
        seen_openings: Set met element_id's die al in deze room zijn
            verwerkt. Gedeeld over alle wall-faces van de room zodat
            SEGC sub-face splitsing niet tot duplicate openings leidt.
        view3d: Revit View3D (raycast-view). Wordt doorgegeven aan
            `_collect_openings_from_hits` -> `_get_opening_z_range_m`
            voor view-projected bbox van linked openings (Bug E.5b).
    """
    direction = _normal_to_compass(normal)
    face_height_m = z_max_m - z_min_m
    if face_height_m <= 0:
        return

    # Cast rays per hoogtestap
    stacks_by_height = {}
    opening_hits_all = []
    empty_heights = []

    # Horizontale ray richting uit face normal (identiek aan
    # _cast_ray_at_height zodat het exit punt correct ligt)
    try:
        ray_direction = XYZ(normal.X, normal.Y, 0.0).Normalize()
    except Exception:
        ray_direction = None

    z = z_min_m + (RAY_HEIGHT_STEP_M / 2.0)
    while z < z_max_m:
        hits, opening_hits = _cast_ray_at_height(
            intersector, room_center_xy, z, normal, doc
        )
        opening_hits_all.extend(opening_hits)

        if not hits:
            empty_heights.append(z)
            z += RAY_HEIGHT_STEP_M
            continue

        # Ray origin zoals gebruikt in _cast_ray_at_height
        ray_origin = XYZ(
            room_center_xy.X,
            room_center_xy.Y,
            z / FEET_TO_M,
        )

        layers, terminal = _hits_to_layer_stack(
            hits, doc, room, all_rooms, ray_origin, ray_direction,
            phase,
        )
        stacks_by_height[z] = (layers, terminal)

        z += RAY_HEIGHT_STEP_M

    # _detect_open_connections gedeactiveerd sinds Bug F fix (Optie C hybride):
    # openings komen nu via Wall.FindInserts() met curtain-panel expansion.
    # De legacy raycast-empty-heights detectie dupliceert + dubbeltelt.
    # Behoud voor referentie, activeer opnieuw alleen als er rooms zijn zonder
    # FindInserts-dekking (bv open doorgangen ZONDER wall-element).
    # open_conns = _detect_open_connections(
    #     empty_heights, z_min_m, z_max_m, direction
    # )
    # result["open_connections"].extend(open_conns)

    if not stacks_by_height:
        return

    # Zones: groepeer opeenvolgende gelijke stacks
    zones = _detect_zones(stacks_by_height)

    # Face breedte schatten uit area en hoogte
    face_width_m = face_area_m2 / face_height_m if face_height_m > 0 else 1.0

    for zone in zones:
        zone_height = zone["z_max"] - zone["z_min"]
        zone_area = face_width_m * zone_height

        construction = {
            "direction": direction,
            "position_type": "wall",
            "area_m2": round(zone_area, 2),
            "z_min_m": round(zone["z_min"], 3),
            "z_max_m": round(zone["z_max"], 3),
            "layers": zone["layers"],
            "terminal_type": zone["terminal_type"],
        }
        result["constructions"].append(construction)

    # Openings via Wall.FindInserts() (hybride optie C - Bug F fix)
    wall_openings = _collect_openings_from_boundary_walls(
        doc, room, normal, z_min_m, z_max_m, direction,
        seen_openings, view3d, face_area_m2
    )

    # Wijs elke opening toe aan de verticale sub-zone waarin hij valt.
    # Dit voorkomt dat een raam op gevel-hoogte wordt afgetrokken van
    # een smalle beton-strook onderaan de wand. Elk opening krijgt een
    # reference naar de `layers`/`terminal_type` van de winnende zone
    # zodat de thermal JSON builder (of backend) de opening aan de
    # juiste geconsolideerde construction kan koppelen.
    for opening in wall_openings:
        zone_idx = _assign_opening_to_zone(opening, zones)
        if zone_idx >= 0:
            winning_zone = zones[zone_idx]
            opening["zone_index"] = zone_idx
            opening["zone_terminal_type"] = winning_zone.get(
                "terminal_type", "outside"
            )
            opening["zone_layer_fingerprint"] = _make_zone_fingerprint(
                winning_zone.get("layers", [])
            )
            opening["zone_z_min_m"] = round(
                winning_zone.get("z_min", 0.0), 3
            )
            opening["zone_z_max_m"] = round(
                winning_zone.get("z_max", 0.0), 3
            )

    result["openings"].extend(wall_openings)


def _make_zone_fingerprint(layers):
    """Maak een stabiele fingerprint van een zone layer stack.

    Gebruikt `(material, thickness_mm)` tuples zodat de downstream JSON
    builder kan matchen op dezelfde sleutel die `_build_constructions`
    hanteert. IronPython 2.7 tuples zijn hashable en serialiseerbaar als
    lijst van lijsten in JSON.

    Args:
        layers: lijst van layer dicts

    Returns:
        list van [material, thickness_mm] paren
    """
    parts = []
    for layer in layers or []:
        material = layer.get("material", "")
        if layer.get("type") == "air_gap":
            material = "__air_gap__"
        thickness = layer.get("thickness_mm", 0)
        parts.append([material, thickness])
    return parts


# =========================================================================
# Horizontal face scanning (floor/ceiling)
# =========================================================================

def _scan_horizontal_face(doc, intersector, room_center_xy, ray_dir,
                          z_m, area_m2, position_type, room,
                          all_rooms, result, phase):
    """Scan een horizontaal vlak (vloer/plafond) met een enkele ray.

    Args:
        doc: Revit host Document
        intersector: ReferenceIntersector
        room_center_xy: XYZ room center
        ray_dir: XYZ ray richting (Z+ of Z-)
        z_m: Hoogte in meters
        area_m2: Oppervlak in m2
        position_type: "floor" of "ceiling"
        room: Revit Room element
        all_rooms: Lijst van room data dicts
        result: Resultaat dict
        phase: Revit Phase voor GetRoomAtPoint, of None
    """
    try:
        origin = XYZ(
            room_center_xy.X,
            room_center_xy.Y,
            z_m / FEET_TO_M,
        )

        refs = intersector.Find(origin, ray_dir)
        if refs is None:
            return

        hits = _filter_ray_hits(refs, doc)
        if not hits:
            return

        layers, terminal = _hits_to_layer_stack(
            hits, doc, room, all_rooms, origin, ray_dir, phase,
        )

        if position_type == "floor":
            dir_str = "floor"
        else:
            dir_str = "ceiling"

        construction = {
            "direction": dir_str,
            "position_type": position_type,
            "area_m2": round(area_m2, 2),
            "z_min_m": round(z_m, 3),
            "z_max_m": round(z_m, 3),
            "layers": layers,
            "terminal_type": terminal,
        }
        result["constructions"].append(construction)

    except Exception:
        pass


# =========================================================================
# Ray casting
# =========================================================================

def _cast_ray_at_height(intersector, origin_xy, z_m, face_normal, doc):
    """Cast een ray vanuit het room center op hoogte z in face-normal richting.

    Args:
        intersector: ReferenceIntersector
        origin_xy: XYZ room center (Z wordt genegeerd)
        z_m: Hoogte in meters
        face_normal: XYZ face normal (horizontaal)
        doc: Revit host Document

    Returns:
        tuple: (construction_hits, opening_hits) -- beide lists of dicts
    """
    construction_hits = []
    opening_hits = []

    try:
        origin = XYZ(
            origin_xy.X,
            origin_xy.Y,
            z_m / FEET_TO_M,
        )

        ray_direction = XYZ(face_normal.X, face_normal.Y, 0.0).Normalize()

        refs = intersector.Find(origin, ray_direction)
        if refs is None:
            return (construction_hits, opening_hits)

        for ref_with_ctx in refs:
            try:
                distance_ft = ref_with_ctx.Proximity
                distance_m = distance_ft * FEET_TO_M

                if distance_m > RAY_MAX_DIST_M:
                    continue

                reference = ref_with_ctx.GetReference()
                hit_info = _resolve_hit_element(reference, doc)
                if hit_info is None:
                    continue

                cat_id = hit_info["category_id"]

                # Skip ongewenste categorieen
                if cat_id in IGNORE_CATEGORIES:
                    continue

                # Openings apart opslaan
                if cat_id in OPENING_CATEGORIES:
                    hit_info["distance_m"] = distance_m
                    opening_hits.append(hit_info)
                    continue

                # Alleen constructie-categorieen doorlaten
                if cat_id not in CONSTRUCTION_CATEGORIES:
                    # Toposolid check (category id kan varieren per versie)
                    cat_name = hit_info.get("category_name", "")
                    if "topo" not in cat_name.lower():
                        continue

                hit_info["distance_m"] = distance_m
                construction_hits.append(hit_info)

            except Exception:
                continue

    except Exception:
        pass

    # Sorteer op afstand
    construction_hits.sort(key=lambda h: h["distance_m"])
    return (construction_hits, opening_hits)


def _filter_ray_hits(refs, doc):
    """Filter ray hit resultaten voor constructie-elementen.

    Gebruikt dezelfde logica als _cast_ray_at_height maar voor
    verticale rays (vloer/plafond) waar geen opening detectie nodig is.

    Args:
        refs: ReferenceWithContext collection
        doc: Revit host Document

    Returns:
        list of hit info dicts gesorteerd op afstand
    """
    hits = []

    for ref_with_ctx in refs:
        try:
            distance_ft = ref_with_ctx.Proximity
            distance_m = distance_ft * FEET_TO_M

            if distance_m > RAY_MAX_DIST_M:
                continue

            reference = ref_with_ctx.GetReference()
            hit_info = _resolve_hit_element(reference, doc)
            if hit_info is None:
                continue

            cat_id = hit_info["category_id"]

            if cat_id in IGNORE_CATEGORIES:
                continue
            if cat_id in OPENING_CATEGORIES:
                continue

            if cat_id not in CONSTRUCTION_CATEGORIES:
                cat_name = hit_info.get("category_name", "")
                if "topo" not in cat_name.lower():
                    continue

            hit_info["distance_m"] = distance_m
            hits.append(hit_info)

        except Exception:
            continue

    hits.sort(key=lambda h: h["distance_m"])
    return hits


def _resolve_hit_element(reference, doc):
    """Resolve een Reference naar element info (host of linked).

    Args:
        reference: Revit Reference object
        doc: Revit host Document

    Returns:
        dict met element info of None
    """
    try:
        linked_elem_id = reference.LinkedElementId
        elem_id = reference.ElementId

        is_linked = (
            linked_elem_id is not None
            and linked_elem_id != ElementId.InvalidElementId
        )

        if is_linked:
            # Element in linked model
            link_instance = doc.GetElement(elem_id)
            if link_instance is None:
                return None

            try:
                linked_doc = link_instance.GetLinkDocument()
            except Exception:
                return None

            if linked_doc is None:
                return None

            element = linked_doc.GetElement(linked_elem_id)
            if element is None:
                return None

            element_doc = linked_doc
            element_id_int = linked_elem_id.IntegerValue
        else:
            # Element in host document
            element = doc.GetElement(elem_id)
            if element is None:
                return None

            element_doc = doc
            element_id_int = elem_id.IntegerValue
            is_linked = False

        # Categorie info
        cat_id = 0
        cat_name = ""
        cat = element.Category
        if cat is not None:
            cat_id = cat.Id.IntegerValue
            cat_name = cat.Name or ""

        # Materiaal naam uit element
        material_name = _get_element_material_name(element, element_doc)

        return {
            "element": element,
            "element_doc": element_doc,
            "element_id": element_id_int,
            "category_id": cat_id,
            "category_name": cat_name,
            "material_name": material_name,
            "is_linked": is_linked,
        }

    except Exception:
        return None


def _get_element_material_name(element, element_doc):
    """Haal de primaire materiaal naam op van een element.

    Probeert eerst het materiaal van de buitenste face, dan
    het structurele materiaal.

    Args:
        element: Revit Element
        element_doc: Document waarin element leeft

    Returns:
        str: Materiaal naam of "Onbekend"
    """
    try:
        # Probeer materiaal IDs
        mat_ids = element.GetMaterialIds(False)
        if mat_ids and mat_ids.Count > 0:
            for mat_id in mat_ids:
                mat = element_doc.GetElement(mat_id)
                if mat is not None and mat.Name:
                    return mat.Name
    except Exception:
        pass

    try:
        # Probeer structural material parameter
        struct_mat_id = element.StructuralMaterialId
        if (struct_mat_id is not None
                and struct_mat_id != ElementId.InvalidElementId):
            mat = element_doc.GetElement(struct_mat_id)
            if mat is not None and mat.Name:
                return mat.Name
    except Exception:
        pass

    return "Onbekend"


# =========================================================================
# Layer stack building
# =========================================================================

def _hits_to_layer_stack(hits, doc, room, all_rooms, ray_origin,
                         ray_direction, phase):
    """Converteer ray hits naar een layer stack met terminal detectie.

    Groepeert hits per element, berekent laagdiktes en detecteert
    luchtspouwen tussen opeenvolgende elementen.

    Args:
        hits: list of hit info dicts, gesorteerd op distance_m
        doc: Revit host Document
        room: Revit Room element (voor adjacent detectie)
        all_rooms: Lijst van room data dicts
        ray_origin: XYZ waarvandaan de ray startte (Revit internal feet)
        ray_direction: XYZ eenheidsvector van de ray richting
        phase: Revit Phase voor GetRoomAtPoint, of None

    Returns:
        tuple: (layers_list, terminal_type_string)
    """
    if not hits:
        return ([], "outside")

    # Groepeer hits per element_id: bepaal enter/exit distance
    element_groups = {}
    element_order = []

    for hit in hits:
        eid = hit["element_id"]
        dist = hit["distance_m"]

        if eid not in element_groups:
            element_groups[eid] = {
                "enter": dist,
                "exit": dist,
                "hit": hit,
            }
            element_order.append(eid)
        else:
            grp = element_groups[eid]
            if dist < grp["enter"]:
                grp["enter"] = dist
            if dist > grp["exit"]:
                grp["exit"] = dist

    # Sorteer op enter distance
    element_order.sort(key=lambda eid: element_groups[eid]["enter"])

    layers = []
    terminal = "outside"
    last_exit = 0.0

    for idx, eid in enumerate(element_order):
        grp = element_groups[eid]
        hit = grp["hit"]
        enter = grp["enter"]
        exit_d = grp["exit"]

        if enter > RAY_MAX_DIST_M:
            break

        # Element dikte: minimaal 1mm (een face hit geeft enter==exit)
        thickness_mm = max(1, int(round((exit_d - enter) * 1000.0)))

        # Fallback: als element een wand/vloer is, probeer Width parameter
        if thickness_mm <= 1:
            fallback_thick = _get_element_thickness_mm(
                hit["element"], hit["element_doc"]
            )
            if fallback_thick > 0:
                thickness_mm = fallback_thick

        # Gap met vorig element
        if idx > 0:
            gap_mm = int(round((enter - last_exit) * 1000.0))

            if gap_mm > MAX_CAVITY_MM:
                # Te grote gap -- beschouw vorige lagen als compleet
                break
            elif gap_mm >= MIN_CAVITY_MM:
                layers.append({
                    "type": "air_gap",
                    "thickness_mm": gap_mm,
                })

        # Laag toevoegen
        lambda_val = _get_hit_lambda(hit)
        layer = {
            "material": hit["material_name"],
            "thickness_mm": thickness_mm,
            "is_linked": hit["is_linked"],
        }
        if lambda_val is not None:
            layer["lambda"] = round(lambda_val, 4)

        layers.append(layer)

        # Update tracking
        last_exit = max(exit_d, enter + (thickness_mm / 1000.0))

        # Terminal detectie op het laatste element
        if idx == len(element_order) - 1:
            terminal = _detect_terminal(
                hit, doc, room, all_rooms, last_exit,
                ray_origin, ray_direction, phase,
            )

    return (layers, terminal)


def _get_element_thickness_mm(element, element_doc):
    """Probeer de dikte van een element op te halen via Width parameter.

    Args:
        element: Revit Element
        element_doc: Document van het element

    Returns:
        int: dikte in mm, of 0 als niet bepaald
    """
    try:
        # Wanden: Width property
        width = getattr(element, "Width", None)
        if width is not None and width > 0:
            return int(round(width * FEET_TO_M * 1000.0))
    except Exception:
        pass

    try:
        # Vloeren/daken: compound structure totale dikte
        elem_type = element_doc.GetElement(element.GetTypeId())
        if elem_type is not None:
            compound = elem_type.GetCompoundStructure()
            if compound is not None:
                total_width = 0.0
                for layer in compound.GetLayers():
                    total_width += layer.Width
                if total_width > 0:
                    return int(round(total_width * FEET_TO_M * 1000.0))
    except Exception:
        pass

    return 0


def _get_hit_lambda(hit):
    """Haal de thermische geleidbaarheid van een hit element.

    Probeert de CompoundStructure lagen, dan het ThermalAsset.

    Args:
        hit: Hit info dict

    Returns:
        float lambda in W/(m*K) of None
    """
    element = hit["element"]
    element_doc = hit["element_doc"]

    try:
        # Via CompoundStructure: gewogen gemiddelde
        elem_type = element_doc.GetElement(element.GetTypeId())
        if elem_type is not None:
            compound = elem_type.GetCompoundStructure()
            if compound is not None:
                layers = compound.GetLayers()
                if layers and layers.Count > 0:
                    total_r = 0.0
                    total_d = 0.0
                    for layer in layers:
                        d = layer.Width * FEET_TO_M
                        if d <= 0:
                            continue
                        mat_id = layer.MaterialId
                        if (mat_id is None
                                or mat_id == ElementId.InvalidElementId):
                            continue
                        mat = element_doc.GetElement(mat_id)
                        if mat is None:
                            continue
                        lam = _get_material_lambda(mat, element_doc)
                        if lam is not None and lam > 0:
                            total_r += d / lam
                            total_d += d

                    if total_d > 0 and total_r > 0:
                        # Equivalent lambda = totale dikte / totale R
                        return total_d / total_r
    except Exception:
        pass

    # Fallback: eerste materiaal met thermal asset
    try:
        mat_ids = element.GetMaterialIds(False)
        if mat_ids:
            for mat_id in mat_ids:
                mat = element_doc.GetElement(mat_id)
                if mat is not None:
                    lam = _get_material_lambda(mat, element_doc)
                    if lam is not None:
                        return lam
    except Exception:
        pass

    return None


def _get_material_lambda(material, material_doc):
    """Haal lambda [W/(m*K)] uit een Revit Material.

    Args:
        material: Revit Material
        material_doc: Document van het materiaal

    Returns:
        float lambda of None
    """
    try:
        thermal_asset_id = material.ThermalAssetId
        if (thermal_asset_id is None
                or thermal_asset_id == ElementId.InvalidElementId):
            return None

        prop_set = material_doc.GetElement(thermal_asset_id)
        if prop_set is None:
            return None

        thermal_asset = prop_set.GetThermalAsset()
        if thermal_asset is None:
            return None

        # Revit ThermalConductivity is BTU*in/(hr*ft2*degF)
        # Conversie naar W/(m*K): / 6.93347
        raw = thermal_asset.ThermalConductivity
        if raw is not None and raw > 0:
            return raw / 6.93347

    except Exception:
        pass

    return None


# =========================================================================
# Terminal detectie
# =========================================================================

def _detect_terminal(hit, doc, room, all_rooms, exit_distance_m,
                     ray_origin, ray_direction, phase):
    """Bepaal de terminal conditie voorbij het laatste constructie-element.

    Checkt Toposolid materiaal keywords, dan adjacent room, dan outside.

    Args:
        hit: Laatste hit info dict
        doc: Revit host Document
        room: Huidige Room
        all_rooms: Alle room data dicts
        exit_distance_m: Afstand tot exit punt in meters
        ray_origin: XYZ waarvandaan de ray startte (Revit internal feet)
        ray_direction: XYZ eenheidsvector van de ray richting
        phase: Revit Phase voor GetRoomAtPoint, of None

    Returns:
        str: "water", "ground", "outside" of room element_id als int
    """
    # Check Toposolid materiaal
    cat_name = hit.get("category_name", "")
    mat_name = hit.get("material_name", "")

    if "topo" in cat_name.lower() or "topo" in mat_name.lower():
        mat_lower = mat_name.lower()
        for keyword in WATER_MATERIAL_KEYWORDS:
            if keyword in mat_lower:
                return "water"
        for keyword in GROUND_MATERIAL_KEYWORDS:
            if keyword in mat_lower:
                return "ground"
        # Topo zonder herkenbaar materiaal -> ground
        return "ground"

    # Check adjacent room
    try:
        adjacent_room_id = _find_room_at_exit(
            doc, room, ray_origin, ray_direction, exit_distance_m,
            phase,
        )
        if adjacent_room_id is not None:
            return adjacent_room_id
    except Exception:
        pass

    return "outside"


def _find_room_at_exit(doc, current_room, ray_origin, ray_direction,
                       exit_distance_m, phase):
    """Zoek een room voorbij het exit punt van de laag-stack.

    Gebruikt Document.GetRoomAtPoint op een XYZ punt net voorbij de
    laatste hit van de ray, in de meegegeven phase.

    Args:
        doc: Revit host Document
        current_room: Huidige Room (voor exclude)
        ray_origin: XYZ waarvandaan de ray startte (Revit internal feet)
        ray_direction: XYZ eenheidsvector van de ray richting
        exit_distance_m: Afstand van origin tot exit punt in meters
        phase: Revit Phase object (verplicht voor GetRoomAtPoint); als
            None dan wordt er geen lookup gedaan

    Returns:
        int room element_id van de gevonden room, of None
    """
    try:
        # Exit punt = origin + direction * (afstand + kleine buffer)
        # Buffer van 0.1 m voorkomt dat we precies op de face landen.
        buffer_m = 0.1
        total_dist_ft = (exit_distance_m + buffer_m) / FEET_TO_M
        exit_point = XYZ(
            ray_origin.X + ray_direction.X * total_dist_ft,
            ray_origin.Y + ray_direction.Y * total_dist_ft,
            ray_origin.Z + ray_direction.Z * total_dist_ft,
        )

        # Zonder phase kunnen we geen betrouwbare lookup doen.
        # GetRoomAtPoint zonder phase retourneert in modellen waar
        # rooms enkel in een named phase bestaan altijd None.
        if phase is None:
            return None

        found_room = doc.GetRoomAtPoint(exit_point, phase)
        if found_room is None:
            return None

        found_id = found_room.Id.IntegerValue
        current_id = current_room.Id.IntegerValue

        # Zelfde room = geen adjacent (geen zelfgrens terugmelden)
        if found_id == current_id:
            return None

        return found_id
    except Exception:
        return None


# =========================================================================
# Zone detectie
# =========================================================================

def _detect_zones(stacks_by_height):
    """Groepeer opeenvolgende ray-hoogtes met gelijke stacks in zones.

    Args:
        stacks_by_height: dict van {z_height: (layers, terminal)}

    Returns:
        list of zone dicts met z_min, z_max, layers, terminal_type
    """
    if not stacks_by_height:
        return []

    sorted_heights = sorted(stacks_by_height.keys())
    zones = []
    current_zone = None

    for z in sorted_heights:
        layers, terminal = stacks_by_height[z]

        if current_zone is None:
            current_zone = {
                "z_min": z - (RAY_HEIGHT_STEP_M / 2.0),
                "z_max": z + (RAY_HEIGHT_STEP_M / 2.0),
                "layers": layers,
                "terminal_type": terminal,
            }
        elif _stacks_are_equal(current_zone["layers"], layers):
            current_zone["z_max"] = z + (RAY_HEIGHT_STEP_M / 2.0)
            # Terminal overschrijven als concreter
            if terminal != "outside":
                current_zone["terminal_type"] = terminal
        else:
            zones.append(current_zone)
            current_zone = {
                "z_min": z - (RAY_HEIGHT_STEP_M / 2.0),
                "z_max": z + (RAY_HEIGHT_STEP_M / 2.0),
                "layers": layers,
                "terminal_type": terminal,
            }

    if current_zone is not None:
        zones.append(current_zone)

    return zones


def _stacks_are_equal(stack_a, stack_b):
    """Vergelijk twee layer stacks op gelijkheid.

    Gelijk als zelfde aantal layers, zelfde materiaal per laag,
    en dikte binnen 5mm tolerantie.

    Args:
        stack_a: list of layer dicts
        stack_b: list of layer dicts

    Returns:
        bool
    """
    if len(stack_a) != len(stack_b):
        return False

    for i in range(len(stack_a)):
        a = stack_a[i]
        b = stack_b[i]

        # Air gap check
        a_is_gap = a.get("type") == "air_gap"
        b_is_gap = b.get("type") == "air_gap"

        if a_is_gap != b_is_gap:
            return False

        if a_is_gap and b_is_gap:
            if abs(a.get("thickness_mm", 0) - b.get("thickness_mm", 0)) > 5:
                return False
            continue

        # Material check
        if a.get("material", "") != b.get("material", ""):
            return False

        # Dikte tolerantie 5mm
        if abs(a.get("thickness_mm", 0) - b.get("thickness_mm", 0)) > 5:
            return False

    return True


# =========================================================================
# Open connection detectie
# =========================================================================

def _detect_open_connections(empty_heights, z_min_m, z_max_m, direction):
    """Detecteer open connections waar rays niets raken.

    Een open connection is een aaneengesloten reeks lege ray-hoogtes
    die minimaal 0.5m hoog is.

    Args:
        empty_heights: Lijst van Z-waarden (m) waar geen hits waren
        z_min_m: Face onderkant
        z_max_m: Face bovenkant
        direction: Kompasrichting string

    Returns:
        list of open connection dicts
    """
    if not empty_heights:
        return []

    sorted_heights = sorted(empty_heights)
    connections = []
    streak_start = sorted_heights[0]
    streak_end = sorted_heights[0]

    for i in range(1, len(sorted_heights)):
        z = sorted_heights[i]
        # Aaneengesloten als binnen 1.5x de stapgrootte
        if z - streak_end <= RAY_HEIGHT_STEP_M * 1.5:
            streak_end = z
        else:
            height = (streak_end - streak_start) + RAY_HEIGHT_STEP_M
            if height >= 0.5:
                connections.append({
                    "direction": direction,
                    "width_m": round(1.0, 2),  # Breedte onbekend bij ray
                    "height_m": round(height, 3),
                })
            streak_start = z
            streak_end = z

    # Laatste streak
    height = (streak_end - streak_start) + RAY_HEIGHT_STEP_M
    if height >= 0.5:
        connections.append({
            "direction": direction,
            "width_m": round(1.0, 2),
            "height_m": round(height, 3),
        })

    return connections


# =========================================================================
# Kompasrichting
# =========================================================================

def _normal_to_compass(normal):
    """Converteer een XYZ normal naar kompasrichting string.

    Gebruikt atan2 op de X,Y componenten.

    Args:
        normal: XYZ vector

    Returns:
        str: N, NE, E, SE, S, SW, W of NW
    """
    angle = math.atan2(normal.Y, normal.X)
    # Converteer van radialen naar graden (0 = oost, 90 = noord)
    degrees = math.degrees(angle)

    # Normaliseer naar 0-360
    if degrees < 0:
        degrees += 360

    # Compass mapping (Revit: X=oost, Y=noord)
    # 0=E, 45=NE, 90=N, 135=NW, 180=W, 225=SW, 270=S, 315=SE
    compass_ranges = [
        (337.5, 360.0, "E"),
        (0.0, 22.5, "E"),
        (22.5, 67.5, "NE"),
        (67.5, 112.5, "N"),
        (112.5, 157.5, "NW"),
        (157.5, 202.5, "W"),
        (202.5, 247.5, "SW"),
        (247.5, 292.5, "S"),
        (292.5, 337.5, "SE"),
    ]

    for low, high, direction in compass_ranges:
        if low <= degrees < high:
            return direction

    return "N"


# =========================================================================
# Opening extractie via Wall.FindInserts() (Bug F fix - hybride optie C)
# =========================================================================

def _collect_openings_from_boundary_walls(doc, room, face_normal, z_min_m,
                                          z_max_m, wall_direction,
                                          seen_openings, view3d,
                                          face_area_m2):
    """Verzamel openings via Wall.FindInserts() voor boundary walls.

    Bug F fix: vervang raycast-gebaseerde opening detectie door Wall.FindInserts()
    om multi-panel curtain walls en geroteerde wanden volledig te scannen.
    Raycast mist openings buiten de ene ray-lijn, FindInserts vindt alle inserts.

    NOTE: DEBUG_OPENINGS=True bypasst overlap-filter voor diagnose.

    Args:
        doc: Revit host Document
        room: Revit Room element
        face_normal: XYZ normale vector van de boundary face
        z_min_m: Onderkant face in meters
        z_max_m: Bovenkant face in meters
        wall_direction: Kompasrichting string ("N", "E", etc.)
        seen_openings: Set met element_id's die al verwerkt zijn (room-wide)
        view3d: Revit View3D voor z-range bepaling
        face_area_m2: Oppervlak van de boundary face voor filter

    Returns:
        list: opening dicts compatible met downstream zone assignment
    """
    openings = []

    # Vind alle walls die bijdragen aan deze room boundary via SEGC
    boundary_walls = _find_boundary_walls_for_face(
        doc, room, face_normal, z_min_m, z_max_m
    )

    if DEBUG_OPENINGS:
        print("DEBUG_OPENINGS: room {} direction {} found {} boundary walls".format(
            room.Id.IntegerValue, wall_direction, len(boundary_walls)
        ))
        wall_ids = [wi["wall"].Id.IntegerValue for wi in boundary_walls]
        linked_flags = [wi["is_linked"] for wi in boundary_walls]
        print("  wall_ids={} linked={}".format(wall_ids, linked_flags))

    for wall_info in boundary_walls:
        wall = wall_info["wall"]
        wall_doc = wall_info["wall_doc"]
        is_linked = wall_info["is_linked"]

        if DEBUG_OPENINGS:
            wall_cls = type(wall).__name__
            wall_cat = "unknown"
            try:
                if wall.Category is not None:
                    wall_cat = wall.Category.Name
            except:
                pass
            wall_kind = "unknown"
            try:
                wall_kind = str(wall.WallType.Kind) if hasattr(wall, "WallType") and wall.WallType is not None else "no-WallType"
            except:
                wall_kind = "kind-error"
            print("  Wall {} class={} cat={} kind={} linked={}".format(
                wall.Id.IntegerValue, wall_cls, wall_cat, wall_kind, is_linked
            ))

        # FindInserts met correcte parameters voor alle opening types
        try:
            # FindInserts(includeShadows, includeEmbeddedWalls,
            #             includeSharedEmbeddedInserts, includeWallOpenings)
            # IronPython 2.7 kan kwargs niet resolven op .NET overloads → positional.
            insert_ids = wall.FindInserts(False, True, False, True)
        except Exception as ex:
            if DEBUG_OPENINGS:
                print("    FindInserts EXCEPTION on Wall {}: {}".format(
                    wall.Id.IntegerValue, str(ex)
                ))
            continue

        if not insert_ids or insert_ids.Count == 0:
            if DEBUG_OPENINGS:
                print("  Wall {} has 0 inserts".format(wall.Id.IntegerValue))
            continue

        if DEBUG_OPENINGS:
            print("  Wall {} has {} inserts".format(
                wall.Id.IntegerValue, insert_ids.Count
            ))

        # Process elke insert
        for insert_id in insert_ids:
            element_id_int = insert_id.IntegerValue

            # Dedupe check (room-wide)
            if element_id_int in seen_openings:
                if DEBUG_OPENINGS:
                    print("    insert {} DROP seen".format(element_id_int))
                continue
            seen_openings.add(element_id_int)

            element = wall_doc.GetElement(insert_id)
            if element is None:
                if DEBUG_OPENINGS:
                    print("    insert {} DROP GetElement=None".format(element_id_int))
                continue

            # Host guard (zelfde logica als raycast variant)
            try:
                host = getattr(element, "Host", None)
                if host is None or host.Id.IntegerValue <= 0:
                    if DEBUG_OPENINGS:
                        print("    insert {} DROP no-host".format(element_id_int))
                    continue
            except Exception:
                if DEBUG_OPENINGS:
                    print("    insert {} DROP no-host".format(element_id_int))
                continue

            # Embedded curtain wall → expandeer naar individuele panels
            is_embedded_wall = False
            try:
                if element.Category is not None and element.Category.Id.IntegerValue == -2000011:
                    is_embedded_wall = True
            except Exception:
                pass

            if is_embedded_wall:
                # Check of het een curtain wall is (heeft CurtainGrid)
                curtain_grid = None
                try:
                    curtain_grid = element.CurtainGrid
                except Exception:
                    curtain_grid = None

                if curtain_grid is not None:
                    try:
                        panel_ids = curtain_grid.GetPanelIds()
                    except Exception:
                        panel_ids = []

                    if DEBUG_OPENINGS:
                        print("    insert {} EXPAND curtain grid -> {} panels".format(
                            element_id_int, len(panel_ids)
                        ))

                    for panel_id in panel_ids:
                        panel_int_id = panel_id.IntegerValue
                        if panel_int_id in seen_openings:
                            if DEBUG_OPENINGS:
                                print("      panel {} DROP seen".format(panel_int_id))
                            continue
                        seen_openings.add(panel_int_id)

                        panel = wall_doc.GetElement(panel_id)
                        if panel is None:
                            if DEBUG_OPENINGS:
                                print("      panel {} DROP GetElement=None".format(panel_int_id))
                            continue

                        panel_type = _classify_insert_opening_type(panel)
                        if panel_type is None:
                            if DEBUG_OPENINGS:
                                # Log category name for diagnosis
                                cat_name = "unknown"
                                try:
                                    if panel.Category is not None:
                                        cat_name = panel.Category.Name
                                except:
                                    pass
                                print("      panel {} DROP classify=None cat={}".format(
                                    panel_int_id, cat_name
                                ))
                            continue

                        p_width_mm, p_height_mm = _get_opening_dimensions_mm(panel, wall_doc)
                        p_u_value = _get_opening_u_value(panel, wall_doc, panel_type)
                        p_z_min_m, p_z_max_m = _get_opening_z_range_m(
                            panel, wall_doc, is_linked, view3d, doc
                        )
                        p_revit_type_name = _get_opening_type_name(panel, wall_doc)

                        if DEBUG_OPENINGS:
                            print("      panel {} KEEP type={} dims={}x{}".format(
                                panel_int_id, panel_type, p_width_mm, p_height_mm
                            ))

                        openings.append({
                            "type": panel_type,
                            "width_mm": p_width_mm,
                            "height_mm": p_height_mm,
                            "wall_direction": wall_direction,
                            "u_value": p_u_value,
                            "revit_element_id": panel_int_id,
                            "revit_type_name": p_revit_type_name,
                            "is_linked": is_linked,
                            "z_min_m": p_z_min_m,
                            "z_max_m": p_z_max_m,
                        })

                    continue  # skip de normale insert-processing voor de curtain-wall-host zelf
                else:
                    # Embedded wall zonder curtain grid → geen opening, skip
                    if DEBUG_OPENINGS:
                        print("    insert {} DROP embedded-wall (no curtain grid)".format(
                            element_id_int
                        ))
                    continue

            # Geometric relevantie: alleen inserts binnen face bbox
            if not _insert_overlaps_face(element, wall_doc, is_linked,
                                       face_normal, z_min_m, z_max_m,
                                       face_area_m2, view3d, doc):
                if DEBUG_OPENINGS:
                    print("    insert {} DROP overlaps=False".format(element_id_int))
                continue

            # Extract opening properties
            opening_type = _classify_insert_opening_type(element)
            if opening_type is None:
                if DEBUG_OPENINGS:
                    cat_name = "unknown"
                    try:
                        if element.Category is not None:
                            cat_name = element.Category.Name
                    except:
                        pass
                    print("    insert {} DROP classify=None cat={}".format(
                        element_id_int, cat_name
                    ))
                continue

            width_mm, height_mm = _get_opening_dimensions_mm(element, wall_doc)
            u_value = _get_opening_u_value(element, wall_doc, opening_type)

            # Z-range voor zone assignment
            z_min_opening_m, z_max_opening_m = _get_opening_z_range_m(
                element, wall_doc, is_linked, view3d, doc
            )

            # Type name voor thermal JSON builder
            revit_type_name = _get_opening_type_name(element, wall_doc)

            if DEBUG_OPENINGS:
                print("    insert {} KEEP type={} dims={}x{}".format(
                    element_id_int, opening_type, width_mm, height_mm
                ))

            openings.append({
                "type": opening_type,
                "width_mm": width_mm,
                "height_mm": height_mm,
                "wall_direction": wall_direction,
                "u_value": u_value,
                "revit_element_id": element_id_int,
                "revit_type_name": revit_type_name,
                "is_linked": is_linked,
                "z_min_m": z_min_opening_m,
                "z_max_m": z_max_opening_m,
            })

    return openings


def _find_boundary_walls_for_face(doc, room, face_normal, z_min_m, z_max_m):
    """Vind alle Wall elementen die bijdragen aan een room boundary face.

    Gebruikt SpatialElementGeometryCalculator om boundary sub-faces te krijgen,
    en extract dan de onderliggende Wall elementen (host + linked).

    Args:
        doc: Revit host Document
        room: Revit Room element
        face_normal: XYZ normale van de gezochte face
        z_min_m: Onderkant face in meters
        z_max_m: Bovenkant face in meters

    Returns:
        list: dicts met 'wall', 'wall_doc', 'is_linked'
    """
    walls = []

    try:
        calculator = SpatialElementGeometryCalculator(doc)
        seg_result = calculator.CalculateSpatialElementGeometry(room)
        spatial_solid = seg_result.GetGeometry()

        for face in spatial_solid.Faces:
            # Check if dit de juiste face is (normale + z-range match)
            face_normal_calc = _get_face_normal(face)
            if not _normals_are_similar(face_normal, face_normal_calc):
                continue

            face_z_min, face_z_max = _get_face_z_range(face)
            if not _z_ranges_overlap(z_min_m, z_max_m, face_z_min, face_z_max):
                continue

            # Haal sub-faces en extract wall elements
            sub_faces = seg_result.GetBoundaryFaceInfo(face)
            if sub_faces and sub_faces.Count > 0:
                for sub_face in sub_faces:
                    element = sub_face.SpatialBoundaryElement
                    if element is None:
                        continue

                    # Check of het een Wall is
                    if element.Category is None:
                        continue
                    cat_id = element.Category.Id.IntegerValue
                    if cat_id != -2000011:  # OST_Walls
                        continue

                    # Host vs linked detectie
                    if hasattr(element, 'Document'):
                        wall_doc = element.Document
                        is_linked = (wall_doc.Title != doc.Title)
                    else:
                        wall_doc = doc
                        is_linked = False

                    walls.append({
                        "wall": element,
                        "wall_doc": wall_doc,
                        "is_linked": is_linked,
                    })

    except Exception:
        # Fallback: probeer walls via room boundaries (legacy methode)
        try:
            walls.extend(_find_walls_via_room_boundaries(doc, room, face_normal))
        except Exception:
            pass

    # Dedup walls op wall.Id.IntegerValue - SEGC kan duplicaten opleveren
    seen_wall_ids = set()
    deduped = []
    for wi in walls:
        wid = wi["wall"].Id.IntegerValue
        if wid in seen_wall_ids:
            continue
        seen_wall_ids.add(wid)
        deduped.append(wi)

    return deduped


def _find_walls_via_room_boundaries(doc, room, face_normal):
    """Fallback: vind walls via Room.GetBoundarySegments().

    Minder robuust dan SEGC maar werkt als fallback voor edge cases.
    """
    walls = []

    try:
        options = SpatialElementBoundaryOptions()
        options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

        boundary_segments = room.GetBoundarySegments(options)
        if not boundary_segments:
            return walls

        for segment_loop in boundary_segments:
            for segment in segment_loop:
                element = doc.GetElement(segment.ElementId)
                if element is None:
                    continue
                if element.Category is None:
                    continue
                if element.Category.Id.IntegerValue != -2000011:  # OST_Walls
                    continue

                walls.append({
                    "wall": element,
                    "wall_doc": doc,
                    "is_linked": False,
                })

    except Exception:
        pass

    return walls


def _normals_are_similar(normal_a, normal_b, tolerance=0.1):
    """Check of twee normals in dezelfde richting wijzen (binnen tolerantie)."""
    try:
        dot_product = normal_a.DotProduct(normal_b)
        return dot_product > (1.0 - tolerance)
    except Exception:
        return False


def _z_ranges_overlap(z1_min, z1_max, z2_min, z2_max, tolerance=0.1):
    """Check of twee z-ranges overlappen (binnen tolerantie)."""
    return not (z1_max + tolerance < z2_min or z2_max + tolerance < z1_min)


def _insert_overlaps_face(element, element_doc, is_linked, face_normal,
                         z_min_m, z_max_m, face_area_m2, view3d, host_doc):
    """Check of een insert geometrisch relevant is voor een boundary face.

    Gebruikt bounding box overlap check. Voor kleine faces (< 1 m2) worden
    alle inserts geaccepteerd om geen openings te missen.

    Args:
        element: Revit FamilyInstance (insert)
        element_doc: Document van element (host of linked)
        is_linked: True als element uit linked model komt
        face_normal: XYZ normale van boundary face
        z_min_m: Onderkant face in meters
        z_max_m: Bovenkant face in meters
        face_area_m2: Oppervlak van face voor filter-drempel
        view3d: View3D voor bbox lookup
        host_doc: Host document voor link transforms

    Returns:
        bool: True als insert relevant is voor deze face
    """
    # Voor kleine faces: alles doorlaten (voorkoming van false negatives)
    if face_area_m2 < 1.0:
        return True

    try:
        # Haal insert z-range op
        ins_z_min_m, ins_z_max_m = _get_opening_z_range_m(
            element, element_doc, is_linked, view3d, host_doc
        )

        if ins_z_min_m is None or ins_z_max_m is None:
            # Geen z-range beschikbaar -> doorlaten (conservatief)
            return True

        # Z-overlap check
        return _z_ranges_overlap(z_min_m, z_max_m, ins_z_min_m, ins_z_max_m)

    except Exception:
        # Bij falen: doorlaten (conservatief)
        return True


def _classify_insert_opening_type(element):
    """Classificeer een insert element als opening type.

    Args:
        element: Revit FamilyInstance

    Returns:
        str: "window", "door", "curtain_wall" of None
    """
    try:
        cat = element.Category
        if cat is None:
            return None

        cat_id = cat.Id.IntegerValue
        if cat_id == -2000014:  # OST_Windows
            return "window"
        elif cat_id == -2000023:  # OST_Doors
            return "door"
        elif cat_id == -2000170:  # OST_CurtainWallPanels
            return "curtain_wall"
        elif cat_id == -2000151:  # OST_GenericModel
            # NL-bouwkunde families (NLRS etc.) zijn vaak Generic Models.
            # Als FindInserts() dit element retourneert, heeft het een gat
            # in de wand gemaakt → per definitie een opening.
            # Naam-heuristiek om type te raden.
            name = ""
            try:
                if element.Name is not None:
                    name = element.Name
            except Exception:
                pass
            # Probeer ook Symbol/FamilyName voor FamilyInstance
            try:
                if hasattr(element, "Symbol") and element.Symbol is not None:
                    sym_name = element.Symbol.Family.Name if element.Symbol.Family else ""
                    if sym_name:
                        name = name + " " + sym_name
            except Exception:
                pass
            lower = name.lower()
            if "deur" in lower or "door" in lower:
                return "door"
            elif "paneel" in lower or "panel" in lower:
                return "curtain_wall"
            # Default: gevelraam
            return "window"

    except Exception:
        pass

    return None


# =========================================================================
# Opening extractie uit ray hits (legacy - niet meer gebruikt na Bug F fix)
# =========================================================================

def _collect_openings_from_hits(doc, opening_hits, wall_direction,
                                seen_openings, view3d):
    """Verzamel openings uit ray hits die in OPENING_CATEGORIES vallen.

    Dedupliceert sub-instances van composite window families (bv.
    `31_WI_raam_1` met genest `BU draai-val rechts` + `WI_subvak`). Die
    sub-elementen classificeren allemaal als OST_Windows maar hebben geen
    echte `Host` wand en zouden anders dubbel/drievoudig meetellen naast
    de parent insert.

    Bug E.4: dedupe set is nu room-wide i.p.v. per-face. SEGC splitst
    een woonboot-wand in meerdere sub-faces en elke sub-face-call gebruikte
    voorheen een nieuwe lokale `seen` set, waardoor hetzelfde raam N keer
    werd opgepikt. `seen_openings` wordt door de caller in
    `scan_room_boundaries` per room aangemaakt en aan elke face-call
    doorgegeven.

    Bug E.5b: de z-range lookup krijgt nu `view3d` + host-doc mee zodat
    linked doors / curtain panels ook een bruikbare z-range terugleveren
    (anders valt `_assign_opening_to_zone` terug op `_largest_zone_index`,
    fout voor woonboot-gevels met waterstrook-zones).

    Voegt per opening een z-range (`z_min_m`, `z_max_m`) toe aan de dict
    zodat de aanroeper (`_scan_wall_face`) de opening kan attribueren
    aan de juiste verticale sub-zone.

    Args:
        doc: Revit host Document
        opening_hits: Lijst van opening hit dicts
        wall_direction: Kompasrichting van de wand
        seen_openings: Set met al verwerkte element_id's (room-wide).
            Wordt in-place bijgewerkt door deze functie.
        view3d: Revit View3D voor view-projected bbox lookup (tier 1
            van `_get_opening_z_range_m`).

    Returns:
        list of opening dicts met `z_min_m` / `z_max_m` gevuld
    """
    if not opening_hits:
        return []

    openings = []

    for hit in opening_hits:
        eid = hit["element_id"]
        if eid in seen_openings:
            continue
        seen_openings.add(eid)

        element = hit["element"]
        element_doc = hit["element_doc"]
        cat_id = hit["category_id"]

        # Skip sub-instances van composite window-families zonder echte
        # Host. Dit voorkomt triple-counting bij families als
        # `31_WI_raam_1` waarbij elk genest sub-element ook als een
        # OST_Windows family instance terugkomt uit de ray hits. Zelfde
        # guard als in `opening_extractor.extract_openings`.
        try:
            host = getattr(element, "Host", None)
            if host is None:
                continue
            host_id = host.Id.IntegerValue
            if host_id <= 0:
                continue
        except Exception:
            continue

        # Classificeer type
        if cat_id in OPENING_CATEGORIES:
            opening_type = OPENING_CATEGORIES[cat_id]
        else:
            opening_type = "window"

        # Afmetingen
        width_mm, height_mm = _get_opening_dimensions_mm(
            element, element_doc
        )

        # U-waarde
        u_value = _get_opening_u_value(element, element_doc, opening_type)

        # Z-range van de opening bounding box (in meters). Wordt
        # gebruikt door `_assign_opening_to_zone` om de opening aan de
        # juiste verticale sub-zone toe te wijzen. Bug E.5b: geef
        # view3d + host_doc door zodat linked elementen ook een
        # bruikbare z-range teruggeven (tier 1/2 fallback).
        z_min_m, z_max_m = _get_opening_z_range_m(
            element, element_doc, hit["is_linked"], view3d, doc
        )

        # Bug E.4: metadata voor downstream JSON builder. De builder
        # (`thermal_json_builder._build_openings`) verwacht
        # `revit_element_id` en `revit_type_name`; de scanner schreef
        # voorheen `element_id` en gaf geen type-naam door, waardoor
        # beide velden `None` bleven in de export.
        revit_type_name = _get_opening_type_name(element, element_doc)

        openings.append({
            "type": opening_type,
            "width_mm": width_mm,
            "height_mm": height_mm,
            "wall_direction": wall_direction,
            "u_value": u_value,
            "revit_element_id": eid,
            "revit_type_name": revit_type_name,
            "is_linked": hit["is_linked"],
            "z_min_m": z_min_m,
            "z_max_m": z_max_m,
        })

    return openings


def _get_opening_type_name(element, element_doc):
    """Haal de type-naam op van een opening FamilyInstance.

    Zelfde format als `opening_extractor._get_type_name`:
    "<FamilyName>: <TypeName>" met fallback op alleen FamilyName of
    lege string bij falen. Gebruikt ALL_MODEL_TYPE_NAME builtin param
    zodat we geen extra `Element` import nodig hebben.

    Args:
        element: Revit FamilyInstance
        element_doc: Document van het element (host of linked)

    Returns:
        str: type-naam of lege string bij falen
    """
    try:
        etype_id = element.GetTypeId()
        if etype_id is None or etype_id == ElementId.InvalidElementId:
            return ""
        elem_type = element_doc.GetElement(etype_id)
        if elem_type is None:
            return ""
        family_name = getattr(elem_type, "FamilyName", "") or ""
        type_param = elem_type.get_Parameter(
            BuiltInParameter.ALL_MODEL_TYPE_NAME
        )
        if type_param is not None and type_param.HasValue:
            type_name = type_param.AsString() or ""
            if family_name and type_name:
                return "{0}: {1}".format(family_name, type_name)
            if type_name:
                return type_name
        return family_name
    except Exception:
        return ""


def _get_link_transform(host_doc, linked_doc):
    """Vind de RevitLinkInstance die bij linked_doc hoort en geef Transform.

    Zoekt in `host_doc` naar een `RevitLinkInstance` met een link-document
    waarvan de `Title` overeenkomt met `linked_doc.Title` en retourneert
    `GetTotalTransform()` zodat de caller een linked-element coordinate
    naar host-coordinaten kan vertalen.

    Args:
        host_doc: Revit host Document
        linked_doc: Revit linked Document

    Returns:
        Transform object of None als geen match gevonden
    """
    if linked_doc is None or host_doc is None:
        return None
    try:
        collector = FilteredElementCollector(host_doc).OfClass(
            RevitLinkInstance
        )
        for link_inst in collector:
            try:
                ld = link_inst.GetLinkDocument()
                if ld is not None and ld.Title == linked_doc.Title:
                    return link_inst.GetTotalTransform()
            except Exception:
                continue
    except Exception:
        pass
    return None


def _get_opening_z_range_m(element, element_doc, is_linked, view3d,
                           host_doc):
    """Robuuste Z-range extractie voor openings in host + linked context.

    Bug E.5b: `element.get_BoundingBox(None)` retourneert `None` voor
    linked doors / curtain panels, waardoor de caller terugviel op de
    `_largest_zone_index` heuristiek -- fout voor woonboot-gevels met
    waterstrook-zones. Drie-tier fallback:

    - Tier 1: `get_BoundingBox(view3d)` -- werkt voor host + linked
      elementen omdat de view-context de link-transform toepast.
    - Tier 2: model-bbox + expliciete link-transform Z-offset.
    - Tier 3: `LocationPoint` + type-height parameter (DOOR_HEIGHT etc).

    Args:
        element: Revit FamilyInstance (window / door / curtain panel)
        element_doc: Document waarin element leeft (host of linked)
        is_linked: True als element uit een linked model komt
        view3d: Revit View3D voor view-projected bbox (mag None zijn)
        host_doc: Revit host Document (voor link-transform lookup)

    Returns:
        tuple: (z_min_m, z_max_m) in meters, of (None, None)
    """
    # Tier 1 -- view-projected bbox (werkt voor host + linked)
    if view3d is not None:
        try:
            bb = element.get_BoundingBox(view3d)
            if bb is not None:
                return (bb.Min.Z * FEET_TO_M, bb.Max.Z * FEET_TO_M)
        except Exception:
            pass

    # Tier 2 -- model bbox + link-transform Z-offset
    try:
        bb = element.get_BoundingBox(None)
        if bb is not None:
            z_min_ft = bb.Min.Z
            z_max_ft = bb.Max.Z
            if is_linked and element_doc is not None and host_doc is not None:
                link_transform = _get_link_transform(host_doc, element_doc)
                if link_transform is not None:
                    # Alleen Z-translatie is relevant voor heat loss zones
                    # (links zijn meestal niet om een horizontale as
                    # geroteerd, dus origin-Z volstaat).
                    origin_z = link_transform.Origin.Z
                    z_min_ft += origin_z
                    z_max_ft += origin_z
            return (z_min_ft * FEET_TO_M, z_max_ft * FEET_TO_M)
    except Exception:
        pass

    # Tier 3 -- LocationPoint + type-height parameter
    try:
        loc = getattr(element, "Location", None)
        if loc is not None and hasattr(loc, "Point") and loc.Point is not None:
            z_center_ft = loc.Point.Z
            if is_linked and element_doc is not None and host_doc is not None:
                link_transform = _get_link_transform(host_doc, element_doc)
                if link_transform is not None:
                    z_center_ft += link_transform.Origin.Z

            # Probeer hoogte uit type-parameters
            h_ft = 0.0
            try:
                etype_id = element.GetTypeId()
                if (etype_id is not None
                        and etype_id != ElementId.InvalidElementId
                        and element_doc is not None):
                    etype = element_doc.GetElement(etype_id)
                    if etype is not None:
                        for bip in (
                            BuiltInParameter.GENERIC_HEIGHT,
                            BuiltInParameter.DOOR_HEIGHT,
                            BuiltInParameter.WINDOW_HEIGHT,
                            BuiltInParameter.FAMILY_HEIGHT_PARAM,
                        ):
                            try:
                                p = etype.get_Parameter(bip)
                                if p is not None and p.HasValue:
                                    v = p.AsDouble()
                                    if v > 0:
                                        h_ft = v
                                        break
                            except Exception:
                                continue
            except Exception:
                pass

            if h_ft > 0:
                z_min_m = (z_center_ft - h_ft / 2.0) * FEET_TO_M
                z_max_m = (z_center_ft + h_ft / 2.0) * FEET_TO_M
                return (z_min_m, z_max_m)
    except Exception:
        pass

    return (None, None)


def _assign_opening_to_zone(opening, zones):
    """Bepaal in welke verticale sub-zone een opening thuis hoort.

    De toewijzing gebruikt de grootste verticale overlap tussen de
    opening `[z_min_m, z_max_m]` en de zone `[z_min, z_max]`. Valt bij
    nul overlap terug op de zone waarin het z-center van de opening
    ligt. Als ook dat niet lukt -- of de opening heeft geen z-range --
    wordt de zone met grootste hoogte (= dominante constructielaag)
    gekozen.

    Bug E.4: de oude legacy fallback retourneerde `0` (eerste zone), wat
    bij woonboten de laagste zone (`z1_water_buiten_20gr`) betekent. Door
    openings zonder z-info daar automatisch op te laten landen ontstaan
    negatieve net_area's op de waterstrook terwijl de ramen op de gevel
    boven water horen. Fallback op grootste zone-hoogte is veilig.

    Args:
        opening: opening dict met optionele `z_min_m` / `z_max_m`
        zones: lijst zone dicts met `z_min` / `z_max` (in meters)

    Returns:
        int: index in `zones`, of -1 als er geen zones zijn
    """
    if not zones:
        return -1

    z_min = opening.get("z_min_m")
    z_max = opening.get("z_max_m")

    # Zonder z-info: fallback op zone met grootste hoogte (zie docstring).
    if z_min is None or z_max is None:
        return _largest_zone_index(zones)

    # Primair: grootste overlap in meters.
    best_idx = -1
    best_overlap = 0.0

    for idx, zone in enumerate(zones):
        zone_min = zone.get("z_min", 0.0)
        zone_max = zone.get("z_max", 0.0)

        overlap_lo = z_min if z_min > zone_min else zone_min
        overlap_hi = z_max if z_max < zone_max else zone_max
        overlap = overlap_hi - overlap_lo

        if overlap > best_overlap:
            best_overlap = overlap
            best_idx = idx

    if best_idx >= 0:
        return best_idx

    # Fallback: zone die het z-center van de opening bevat.
    z_center = (z_min + z_max) / 2.0
    for idx, zone in enumerate(zones):
        zone_min = zone.get("z_min", 0.0)
        zone_max = zone.get("z_max", 0.0)
        if zone_min <= z_center <= zone_max:
            return idx

    # Laatste redmiddel: zone met grootste hoogte i.p.v. index 0.
    return _largest_zone_index(zones)


def _largest_zone_index(zones):
    """Geef de index van de zone met de grootste verticale hoogte.

    Gebruikt als veilige fallback voor `_assign_opening_to_zone` wanneer
    de opening geen bruikbare z-range heeft of geen overlap met een zone
    oplevert. De grootste zone is vrijwel altijd de dominante
    constructielaag (bv. de hoofdgevel boven een smalle waterstrook).

    Args:
        zones: lijst zone dicts met `z_min` / `z_max` (in meters)

    Returns:
        int: index in `zones`, of 0 als alle zones hoogte 0 hebben
    """
    best_idx = 0
    best_height = -1.0
    for idx in range(len(zones)):
        zone = zones[idx]
        height = zone.get("z_max", 0.0) - zone.get("z_min", 0.0)
        if height > best_height:
            best_height = height
            best_idx = idx
    return best_idx


def _get_opening_dimensions_mm(element, element_doc):
    """Bepaal afmetingen van een opening in millimeters.

    Args:
        element: Revit FamilyInstance
        element_doc: Document van het element

    Returns:
        tuple: (width_mm, height_mm)
    """
    width_mm = 1000
    height_mm = 1500

    try:
        from warmteverlies.unit_utils import get_param_value

        # Instance parameters
        for w_name in ["Width", "Rough Width", "Breedte"]:
            w = get_param_value(element, w_name)
            if w is not None and w > 0:
                width_mm = int(round(w * FEET_TO_M * 1000.0))
                break

        for h_name in ["Height", "Rough Height", "Hoogte"]:
            h = get_param_value(element, h_name)
            if h is not None and h > 0:
                height_mm = int(round(h * FEET_TO_M * 1000.0))
                break

        # Type parameters fallback
        try:
            elem_type = element_doc.GetElement(element.GetTypeId())
            if elem_type:
                if width_mm == 1000:
                    for w_name in ["Width", "Rough Width", "Breedte"]:
                        w = get_param_value(elem_type, w_name)
                        if w is not None and w > 0:
                            width_mm = int(round(w * FEET_TO_M * 1000.0))
                            break

                if height_mm == 1500:
                    for h_name in ["Height", "Rough Height", "Hoogte"]:
                        h = get_param_value(elem_type, h_name)
                        if h is not None and h > 0:
                            height_mm = int(round(h * FEET_TO_M * 1000.0))
                            break
        except Exception:
            pass

    except Exception:
        pass

    return (width_mm, height_mm)


def _get_opening_u_value(element, element_doc, opening_type):
    """Bepaal de U-waarde van een opening.

    Args:
        element: Revit FamilyInstance
        element_doc: Document van het element
        opening_type: "window", "door" of "curtain_wall"

    Returns:
        float U-waarde in W/(m2*K)
    """
    try:
        from warmteverlies.unit_utils import get_param_value
        from warmteverlies.constants import DEFAULT_U_VALUES

        for param_name in [
            "Heat Transfer Coefficient (U)",
            "Warmtedoorgangscoefficient (U)",
            "Thermal Transmittance",
        ]:
            u = get_param_value(element, param_name)
            if u is not None and u > 0:
                return u

        # Type parameter
        try:
            elem_type = element_doc.GetElement(element.GetTypeId())
            if elem_type:
                for param_name in [
                    "Heat Transfer Coefficient (U)",
                    "Warmtedoorgangscoefficient (U)",
                    "Thermal Transmittance",
                ]:
                    u = get_param_value(elem_type, param_name)
                    if u is not None and u > 0:
                        return u
        except Exception:
            pass

        if opening_type == "window":
            return DEFAULT_U_VALUES["window"]
        return DEFAULT_U_VALUES["door_exterior"]

    except Exception:
        return 1.6
