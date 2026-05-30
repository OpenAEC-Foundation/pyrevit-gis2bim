# -*- coding: utf-8 -*-
"""Visuele controle van SEGC-grensvlakken via gekleurde DirectShapes.

Rendert per ruimte de grensvlakken die de warmteverlies-exporter geometrisch
ziet (uit SpatialElementGeometryCalculator) als gekleurde Generic Models.
Dit is een visuele controle vooraf, voordat de JSON-export wordt gedraaid.

Kleurcode:
- dak/plafond (top)  = rood
- wand (verticaal)   = geel
- vloer (bot)        = groen
- openingen/vlies    = blauw

Alle DirectShapes krijgen Comments-prefix "WV_BND" voor cleanup.

IronPython 2.7 — geen f-strings, geen type hints.
"""
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    SpatialElementBoundaryOptions,
    SpatialElementGeometryCalculator,
    UV,
    XYZ,
    Color,
    ElementId,
    BuiltInCategory,
    DirectShape,
    Material,
    Options,
    GeometryInstance,
    Solid,
    TessellatedShapeBuilder,
    TessellatedFace,
    TessellatedShapeBuilderTarget,
    TessellatedShapeBuilderFallback,
    FilteredElementCollector,
)

from warmteverlies.unit_utils import internal_to_sqm

# =============================================================================
# Constanten
# =============================================================================
COMMENTS_PREFIX = "WV_BND"

# Materiaal-namen + RGB
MAT_TOP = ("WV_BND_TOP", (200, 60, 60))    # dak/plafond = rood
MAT_WALL = ("WV_BND_WALL", (225, 200, 80))  # wand = geel
MAT_BOT = ("WV_BND_BOT", (70, 175, 80))     # vloer = groen
MAT_OPEN = ("WV_BND_OPEN", (0, 170, 255))   # openingen = blauw

MATERIAL_TRANSPARENCY = 15

# Horizontaal-drempel voor face-normal Z
HORIZ_NORMAL_Z = 0.7

# Opening-offset langs wand-normaal richting ruimtecentrum (mm -> feet)
OPENING_OFFSET_MM = 20.0
MM_TO_FEET = 1.0 / 304.8

OPENING_CATEGORIES = ("Doors", "Windows")


# =============================================================================
# Materiaal-management
# =============================================================================
def ensure_materials(doc):
    """Maak de WV_BND materialen aan (indien afwezig) of hergebruik op naam.

    Args:
        doc: Revit Document

    Returns:
        dict: {naam: ElementId} voor de 4 materialen
    """
    # Bestaande materialen op naam indexeren
    existing = {}
    collector = FilteredElementCollector(doc).OfClass(Material)
    for mat in collector:
        existing[mat.Name] = mat.Id

    result = {}
    for name, rgb in (MAT_TOP, MAT_WALL, MAT_BOT, MAT_OPEN):
        if name in existing:
            result[name] = existing[name]
            continue
        mat_id = Material.Create(doc, name)
        mat = doc.GetElement(mat_id)
        mat.Color = Color(rgb[0], rgb[1], rgb[2])
        mat.Transparency = MATERIAL_TRANSPARENCY
        result[name] = mat_id
    return result


# =============================================================================
# Geometrie helpers
# =============================================================================
def _face_normal(face):
    """Bepaal de normal van een face met fallback.

    Args:
        face: Revit Face

    Returns:
        XYZ of None
    """
    try:
        return face.ComputeNormal(UV(0.5, 0.5))
    except Exception:
        pass
    try:
        return face.FaceNormal
    except Exception:
        return None


def _valid_host_ids(seg_result, face):
    """Verzamel geldige HostElementId integers voor een face.

    Args:
        seg_result: SpatialElementGeometryResults
        face: Revit Face

    Returns:
        list[int]: host element id integers (> 0)
    """
    ids = []
    try:
        sub_faces = seg_result.GetBoundaryFaceInfo(face)
    except Exception:
        return ids
    for sf in sub_faces:
        try:
            sbe = sf.SpatialBoundaryElement
            if sbe is None:
                continue
            hid = sbe.HostElementId
            if hid is not None and hid.IntegerValue > 0:
                ids.append(hid.IntegerValue)
        except Exception:
            continue
    return ids


def _is_curtain_wall(element):
    """Controleer of een element een curtain wall is (heeft CurtainGrid)."""
    try:
        return element.CurtainGrid is not None
    except Exception:
        return False


def _build_directshape_from_triangles(doc, triangles, material_id, comment):
    """Bouw een DirectShape uit een lijst driehoeken (elk 3 XYZ).

    Args:
        doc: Revit Document
        triangles: list van (v0, v1, v2) XYZ-tuples
        material_id: ElementId van het materiaal
        comment: string voor de Comments-parameter (WV_BND prefix)

    Returns:
        DirectShape of None
    """
    if not triangles:
        return None

    builder = TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    added = 0
    for tri in triangles:
        try:
            pts = List[XYZ]()
            pts.Add(tri[0])
            pts.Add(tri[1])
            pts.Add(tri[2])
            builder.AddFace(TessellatedFace(pts, material_id))
            added += 1
        except Exception:
            continue
    if added == 0:
        return None

    builder.CloseConnectedFaceSet()
    builder.Target = TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = TessellatedShapeBuilderFallback.Mesh
    builder.Build()

    geom = builder.GetBuildResult().GetGeometricalObjects()
    if geom.Count == 0:
        return None

    ds = DirectShape.CreateElement(
        doc, ElementId(BuiltInCategory.OST_GenericModel)
    )
    ds.SetShape(geom)
    try:
        comment_param = ds.LookupParameter("Comments")
        if comment_param is not None:
            comment_param.Set(comment)
    except Exception:
        pass
    return ds


def _face_to_triangles(face):
    """Trianguleer een face naar een lijst (v0, v1, v2) tuples.

    Args:
        face: Revit Face

    Returns:
        list van XYZ-tuples
    """
    triangles = []
    try:
        mesh = face.Triangulate()
    except Exception:
        return triangles
    if mesh is None:
        return triangles
    count = mesh.NumTriangles
    for i in range(count):
        try:
            tri = mesh.get_Triangle(i)
            triangles.append((
                tri.get_Vertex(0),
                tri.get_Vertex(1),
                tri.get_Vertex(2),
            ))
        except Exception:
            continue
    return triangles


# =============================================================================
# Openingen (deuren / ramen)
# =============================================================================
def _collect_solids_from_geom(geom_element, out_solids):
    """Verzamel Solids uit een geometry element (incl. instance geometry).

    Args:
        geom_element: GeometryElement
        out_solids: list om te vullen
    """
    if geom_element is None:
        return
    for obj in geom_element:
        if isinstance(obj, Solid):
            if obj.Volume > 0:
                out_solids.append(obj)
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
                _collect_solids_from_geom(inst_geom, out_solids)
            except Exception:
                continue


def _opening_vertices(insert):
    """Verzamel alle geometrie-vertices van een opening (door/window).

    Args:
        insert: FamilyInstance (deur of raam)

    Returns:
        list[XYZ]
    """
    vertices = []
    solids = []
    opt = Options()
    try:
        geom = insert.get_Geometry(opt)
    except Exception:
        geom = None
    _collect_solids_from_geom(geom, solids)

    for solid in solids:
        for face in solid.Faces:
            for tri in _face_to_triangles(face):
                vertices.append(tri[0])
                vertices.append(tri[1])
                vertices.append(tri[2])
    return vertices


def _opening_direction(insert, host_wall):
    """Bepaal de horizontale richting langs de wand voor de opening.

    Args:
        insert: FamilyInstance
        host_wall: Wall element

    Returns:
        (dx, dy) genormaliseerde richting, of None
    """
    # Probeer HandOrientation
    try:
        hand = insert.HandOrientation
        if hand is not None:
            length = (hand.X * hand.X + hand.Y * hand.Y) ** 0.5
            if length > 1e-9:
                return (hand.X / length, hand.Y / length)
    except Exception:
        pass

    # Fallback: host LocationCurve-richting
    try:
        loc = host_wall.Location
        curve = loc.Curve
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        length = (dx * dx + dy * dy) ** 0.5
        if length > 1e-9:
            return (dx / length, dy / length)
    except Exception:
        pass
    return None


def _build_opening_triangles(insert, host_wall, room_center):
    """Bouw de 2 driehoeken voor een opening-rechthoek in het wandvlak.

    Args:
        insert: FamilyInstance (deur/raam)
        host_wall: Wall element
        room_center: XYZ ruimtecentrum (voor offset-richting)

    Returns:
        list van (v0, v1, v2) tuples (2 driehoeken) of []
    """
    direction = _opening_direction(insert, host_wall)
    if direction is None:
        return []
    dx, dy = direction

    vertices = _opening_vertices(insert)
    if not vertices:
        return []

    # Centroid (cx, cy)
    n = len(vertices)
    cx = sum(v.X for v in vertices) / n
    cy = sum(v.Y for v in vertices) / n

    # z-bereik
    zmin = min(v.Z for v in vertices)
    zmax = max(v.Z for v in vertices)

    # s-bereik langs richting
    s_vals = [v.X * dx + v.Y * dy for v in vertices]
    smin = min(s_vals)
    smax = max(s_vals)

    base_s = cx * dx + cy * dy

    def corner(s_abs, z):
        return XYZ(
            cx + dx * (s_abs - base_s),
            cy + dy * (s_abs - base_s),
            z,
        )

    c0 = corner(smin, zmin)
    c1 = corner(smax, zmin)
    c2 = corner(smax, zmax)
    c3 = corner(smin, zmax)

    # Wand-normaal (nx, ny) = (-dy, dx); offset richting ruimtecentrum
    nx = -dy
    ny = dx
    to_center_x = room_center.X - cx
    to_center_y = room_center.Y - cy
    if (nx * to_center_x + ny * to_center_y) < 0:
        nx = -nx
        ny = -ny

    offset = OPENING_OFFSET_MM * MM_TO_FEET
    ox = nx * offset
    oy = ny * offset

    def shift(p):
        return XYZ(p.X + ox, p.Y + oy, p.Z)

    c0 = shift(c0)
    c1 = shift(c1)
    c2 = shift(c2)
    c3 = shift(c3)

    return [
        (c0, c1, c2),
        (c0, c2, c3),
    ]


def _room_center(room_element):
    """Bepaal een ruimtecentrum (XYZ) voor offset-richting.

    Args:
        room_element: Revit Room

    Returns:
        XYZ
    """
    try:
        loc = room_element.Location
        if loc is not None:
            pt = loc.Point
            if pt is not None:
                return pt
    except Exception:
        pass
    return XYZ(0.0, 0.0, 0.0)


# =============================================================================
# Hoofdrenderfunctie
# =============================================================================
def render_room_boundaries(doc, rooms, material_ids, params, output=None):
    """Render de SEGC-grensvlakken voor alle (verwarmde) ruimten.

    Args:
        doc: Revit Document
        rooms: list[dict] room-data (uit collect_rooms + map_all_rooms)
        material_ids: dict {naam: ElementId} (uit ensure_materials)
        params: dict met keys:
            min_face_area_m2 (float)
            show_openings (bool)
            hide_hostless_slivers (bool)
            heated_only (bool)
        output: pyRevit output object (optioneel) voor voortgang

    Returns:
        dict met telling-statistieken
    """
    min_area = params.get("min_face_area_m2", 0.10)
    show_openings = params.get("show_openings", True)
    hide_slivers = params.get("hide_hostless_slivers", True)
    heated_only = params.get("heated_only", True)

    stats = {
        "rooms_processed": 0,
        "rooms_skipped": 0,
        "rooms_failed": 0,
        "top": 0,
        "bot": 0,
        "wall": 0,
        "open": 0,
        "slivers_hidden": 0,
        "openings": 0,
        "faces_failed": 0,
    }

    opt = SpatialElementBoundaryOptions()

    # Globale dedup van openingen over alle ruimten heen
    rendered_openings = set()

    for room_data in rooms:
        name = room_data.get("name", "") or ""
        if "buiten" in name.lower():
            stats["rooms_skipped"] += 1
            continue
        if heated_only and not room_data.get("is_heated"):
            stats["rooms_skipped"] += 1
            continue

        room_element = room_data.get("element")
        room_number = room_data.get("number", "?")
        if room_element is None:
            stats["rooms_skipped"] += 1
            continue

        try:
            calc = SpatialElementGeometryCalculator(doc, opt)
            seg_result = calc.CalculateSpatialElementGeometry(room_element)
            solid = seg_result.GetGeometry()
        except Exception:
            stats["rooms_failed"] += 1
            continue

        if solid is None:
            stats["rooms_failed"] += 1
            continue

        center = _room_center(room_element)

        # Wanden waarvoor we openingen moeten zoeken (niet-curtain hosts)
        host_wall_ids = set()

        for face in solid.Faces:
            try:
                normal = _face_normal(face)
                if normal is None:
                    stats["faces_failed"] += 1
                    continue

                area_m2 = internal_to_sqm(face.Area)
                if area_m2 < min_area:
                    continue

                horiz = abs(normal.Z) > HORIZ_NORMAL_Z

                hosts = _valid_host_ids(seg_result, face)

                if horiz:
                    if normal.Z > 0:
                        mat_name = MAT_TOP[0]
                        orient = "top"
                    else:
                        mat_name = MAT_BOT[0]
                        orient = "bot"
                else:
                    # Verticaal
                    if not hosts:
                        if hide_slivers:
                            stats["slivers_hidden"] += 1
                            continue
                        mat_name = MAT_WALL[0]
                        orient = "wall"
                    else:
                        # Check curtain wall onder de hosts
                        is_curtain = False
                        for hid in hosts:
                            host_el = doc.GetElement(ElementId(hid))
                            if host_el is not None and _is_curtain_wall(host_el):
                                is_curtain = True
                                break
                        if is_curtain:
                            mat_name = MAT_OPEN[0]
                            orient = "vliesgevel"
                        else:
                            mat_name = MAT_WALL[0]
                            orient = "wall"
                            for hid in hosts:
                                host_el = doc.GetElement(ElementId(hid))
                                if host_el is not None:
                                    host_wall_ids.add(hid)

                triangles = _face_to_triangles(face)
                comment = "{0} {1} {2}".format(
                    COMMENTS_PREFIX, room_number, orient
                )
                ds = _build_directshape_from_triangles(
                    doc, triangles, material_ids[mat_name], comment
                )
                if ds is not None:
                    if orient == "top":
                        stats["top"] += 1
                    elif orient == "bot":
                        stats["bot"] += 1
                    elif orient == "vliesgevel":
                        stats["open"] += 1
                    else:
                        stats["wall"] += 1
            except Exception:
                stats["faces_failed"] += 1
                continue

        # --- Openingen (deuren/ramen) per niet-curtain host-wand ---
        if show_openings:
            for wid in host_wall_ids:
                try:
                    wall = doc.GetElement(ElementId(wid))
                    if wall is None:
                        continue
                    inserts = wall.FindInserts(True, False, True, True)
                except Exception:
                    continue

                for ins_id in inserts:
                    try:
                        ins_key = ins_id.IntegerValue
                        if ins_key in rendered_openings:
                            continue
                        insert = doc.GetElement(ins_id)
                        if insert is None:
                            continue
                        cat = insert.Category
                        if cat is None or cat.Name not in OPENING_CATEGORIES:
                            continue

                        tris = _build_opening_triangles(insert, wall, center)
                        if not tris:
                            continue
                        comment = "{0} {1} opening".format(
                            COMMENTS_PREFIX, room_number
                        )
                        ds = _build_directshape_from_triangles(
                            doc, tris, material_ids[MAT_OPEN[0]], comment
                        )
                        if ds is not None:
                            rendered_openings.add(ins_key)
                            stats["openings"] += 1
                    except Exception:
                        continue

        stats["rooms_processed"] += 1

        if output is not None:
            try:
                output.print_md(
                    "- Ruimte **{0}** ({1}) verwerkt".format(
                        room_number, name
                    )
                )
            except Exception:
                pass

    return stats


# =============================================================================
# Cleanup
# =============================================================================
def clear_boundary_shapes(doc):
    """Verwijder alle DirectShapes met Comments-prefix WV_BND.

    Args:
        doc: Revit Document

    Returns:
        int: aantal verwijderde elementen
    """
    collector = (
        FilteredElementCollector(doc)
        .OfClass(DirectShape)
        .WhereElementIsNotElementType()
    )
    to_delete = []
    for ds in collector:
        try:
            param = ds.LookupParameter("Comments")
            if param is None or not param.HasValue:
                continue
            value = param.AsString()
            if value and value.startswith(COMMENTS_PREFIX):
                to_delete.append(ds.Id)
        except Exception:
            continue

    count = 0
    for eid in to_delete:
        try:
            doc.Delete(eid)
            count += 1
        except Exception:
            continue
    return count
