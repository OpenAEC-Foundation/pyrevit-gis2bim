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
                result, phase,
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
                    all_rooms, result, phase):
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

    # Detecteer open connections (grote lege gebieden)
    open_conns = _detect_open_connections(
        empty_heights, z_min_m, z_max_m, direction
    )
    result["open_connections"].extend(open_conns)

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

    # Openings uit ray hits
    wall_openings = _collect_openings_from_hits(
        doc, opening_hits_all, direction
    )
    result["openings"].extend(wall_openings)


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
# Opening extractie uit ray hits
# =========================================================================

def _collect_openings_from_hits(doc, opening_hits, wall_direction):
    """Verzamel openings uit ray hits die in OPENING_CATEGORIES vallen.

    Args:
        doc: Revit host Document
        opening_hits: Lijst van opening hit dicts
        wall_direction: Kompasrichting van de wand

    Returns:
        list of opening dicts
    """
    if not opening_hits:
        return []

    # Deduplicate per element_id
    seen = set()
    openings = []

    for hit in opening_hits:
        eid = hit["element_id"]
        if eid in seen:
            continue
        seen.add(eid)

        element = hit["element"]
        element_doc = hit["element_doc"]
        cat_id = hit["category_id"]

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

        openings.append({
            "type": opening_type,
            "width_mm": width_mm,
            "height_mm": height_mm,
            "wall_direction": wall_direction,
            "u_value": u_value,
            "element_id": eid,
            "is_linked": hit["is_linked"],
        })

    return openings


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
