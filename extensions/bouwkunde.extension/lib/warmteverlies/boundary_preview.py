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
import os

from System.Collections.Generic import List, IList

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
    View3D,
    ViewFamilyType,
    ViewFamily,
)

from warmteverlies.unit_utils import internal_to_sqm
from warmteverlies.constants import (
    RAY_MAX_DIST_M,
    CONSTRUCTION_CATEGORIES,
    IGNORE_CATEGORIES,
    FEET_TO_M,
)
from warmteverlies.raycast_scanner import (
    _create_intersector,
    _resolve_hit_element,
)

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

# Offset van de blauwe opening-rechthoek IN het gat (tegen z-fighting)
HOLE_OPENING_OFFSET_MM = 15.0

# Clamp-marge (ft, intern) waarmee het gat NAAR BINNEN wordt geknepen,
# zodat een gat altijd binnen de outer-loop van het wandvlak valt.
HOLE_CLAMP_MARGIN_FT = 0.01

# Square-feet -> square-meters (face.Area is intern in ft^2)
SQFT_TO_M2 = 0.092903

OPENING_CATEGORIES = ("Doors", "Windows")

# Meter -> feet (interne Revit-eenheid)
METER_TO_FEET = 1.0 / 0.3048

# Adjacency-probe: offsets (meter) naar buiten langs de face-normaal, door de
# gestapelde constructie-dikte heen, om de buurruimte te vinden.
ADJ_PROBE_OFFSETS_M = (0.10, 0.20, 0.35, 0.50, 0.75, 1.0)

# Inwaartse offset (meter) om de normaal-richting te ijken op de bron-ruimte.
ADJ_INWARD_CHECK_M = 0.15

# Type-stapel: marge (meter) bovenop de buurruimte-cutoff, zodat de verre
# wandzijde-laag net wel meekomt (de probe vindt de buur iets achter de wand).
TYPE_STACK_CUTOFF_MARGIN_M = 0.02

# Type-stapel fallback (echte buitenlucht): max luchtspleet (meter) tussen twee
# opeenvolgende constructie-hits; groter gat = einde wandpakket -> stoppen.
TYPE_STACK_MAX_AIRGAP_M = 0.05

# Ruimte-naam-patronen die buitenlucht voorstellen: een door de probe gevonden
# ruimte met zo'n naam telt als exterior, niet als unheated_space.
# (case-insensitive substring-match, whitespace gestript)
OUTDOOR_ROOM_NAME_PATTERNS = ("buiten",)

# =============================================================================
# Shared parameters (adjacency-info op de DirectShapes)
# =============================================================================
# Groep in het shared-parameter-bestand
SHARED_PARAM_GROUP = "Berekeningen"

# (naam, type_key)  type_key in {"text", "number"}
WV_PARAM_DEFS = (
    ("warmteverlies_ruimte", "text"),
    ("warmteverlies_naar_ruimte", "text"),
    ("warmteverlies_grenstype", "text"),
    ("warmteverlies_orientatie", "text"),
    ("warmteverlies_oppervlak_m2", "number"),
    ("warmteverlies_host_type", "text"),
    ("warmteverlies_type_stapel", "text"),
)

# orient (intern) -> NL label voor warmteverlies_orientatie
ORIENT_LABEL = {
    "top": "dak",
    "bot": "vloer",
    "wall": "wand",
    "vliesgevel": "opening",
    "open": "opening",
}

# Type-parameter (Yes/No) die een constructie-Type als afwerklaag markeert.
# Type-binding op Walls/Floors/Roofs/Ceilings. Afwerklagen worden uit de
# type-stapel (raycast) gefilterd.
AFWERKLAAG_PARAM_NAME = "warmteverlies_afwerklaag"

# Type Comments-substring waarop bootstrap een Type als afwerklaag markeert.
AFWERKLAAG_COMMENT_PATTERN = "afwerk"

# Module-flag: parameters al gebonden in deze sessie (idempotent fast-path)
_wv_params_created = False
_afwerklaag_param_created = False


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


def _is_glass_wall(doc, element):
    """Geldt dit wand-element als vliesgevel/opening voor warmteverlies?

    True voor (a) echte curtain walls (CurtainGrid) en (b) gewone wanden met
    een glas-compound, herkend aan de 3BM-typenaamconventie ('glas' in de
    Type-naam, bv. 28_CWA_glas20mm). Beide tellen als opening: overwegend glas
    draagt thermisch als raam, niet als dichte wand.
    """
    try:
        if _is_curtain_wall(element):
            return True
    except Exception:
        pass
    try:
        tn = _element_type_name(doc, element) or ""
        if "glas" in tn.lower():
            return True
    except Exception:
        pass
    return False


# OST_CurtainWallPanels (glas-paneel) en OST_GenericModel (eigen WV-shapes)
CURTAIN_PANEL_CAT_ID = -2000170
GENERIC_MODEL_CAT_ID = -2000151


def _curtain_glass_behind(doc, intersector, centroid, outward, max_dist_m=0.4):
    """True als het eerste echte element achter dit vlak een curtain-wall-paneel
    (glas) is. Detecteert een vliesgevel die NIET de room-bounding host is maar
    er vlak achter zit — bv. een bewust geplaatste room-separation-line op de
    gevel met de curtain wall erachter. Negeert eigen WV-shapes (Generic Models)
    en de IGNORE-categorieen.
    """
    if intersector is None or centroid is None or outward is None:
        return False
    back = 0.05 * METER_TO_FEET
    origin = XYZ(
        centroid.X - outward.X * back,
        centroid.Y - outward.Y * back,
        centroid.Z - outward.Z * back,
    )
    try:
        refs = intersector.Find(origin, outward)
    except Exception:
        return False
    if refs is None:
        return False
    best_prox = None
    best_cat = None
    for rwc in refs:
        try:
            prox_m = rwc.Proximity * FEET_TO_M
            if prox_m > max_dist_m:
                continue
            hit = _resolve_hit_element(rwc.GetReference(), doc)
            if hit is None:
                continue
            cat_id = hit.get("category_id")
            if cat_id in IGNORE_CATEGORIES:
                continue
            if cat_id == GENERIC_MODEL_CAT_ID:
                continue
            if best_prox is None or prox_m < best_prox:
                best_prox = prox_m
                best_cat = cat_id
        except Exception:
            continue
    return best_prox is not None and best_cat == CURTAIN_PANEL_CAT_ID


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
# Multi-loop face (wand met gat)
# =============================================================================
def _face_outer_loop(face):
    """Haal de buitencontour-punten van een face op.

    Args:
        face: Revit Face

    Returns:
        list[XYZ] of [] bij faal
    """
    pts = []
    try:
        loops = face.GetEdgesAsCurveLoops()
    except Exception:
        return pts
    if loops is None or loops.Count == 0:
        return pts
    # Eerste loop = buitencontour
    outer = loops[0]
    for curve in outer:
        try:
            pts.append(curve.GetEndPoint(0))
        except Exception:
            continue
    return pts


def _project_on_plane(p, origin, normal):
    """Projecteer punt p loodrecht op het vlak (origin, normal).

    p_proj = p - n * dot(p - o, n)
    """
    vx = p.X - origin.X
    vy = p.Y - origin.Y
    vz = p.Z - origin.Z
    d = vx * normal.X + vy * normal.Y + vz * normal.Z
    return XYZ(
        p.X - normal.X * d,
        p.Y - normal.Y * d,
        p.Z - normal.Z * d,
    )


def _build_directshape_face_with_holes(
    doc, outer_pts, holes_corners, material_id, normal, comment
):
    """Bouw een DirectShape: één face met gaten via multi-loop TessellatedFace.

    Args:
        doc: Revit Document
        outer_pts: list[XYZ] buitencontour
        holes_corners: list van list[XYZ] (per gat 4 hoeken, al geprojecteerd)
        material_id: ElementId van het wand-materiaal
        normal: XYZ face-normal (voor winding-bepaling)
        comment: string voor Comments-parameter

    Returns:
        DirectShape of None (None => caller moet fallback renderen)
    """
    if not outer_pts or len(outer_pts) < 3:
        return None

    builder = TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)

    loops = List[IList[XYZ]]()

    outer = List[XYZ]()
    for p in outer_pts:
        outer.Add(p)
    loops.Add(outer)

    # Gaten met tegengestelde winding t.o.v. de buitencontour
    for corners in holes_corners:
        if len(corners) < 3:
            continue
        inner = List[XYZ]()
        # Omgekeerde volgorde => tegengestelde winding (gat)
        for p in reversed(corners):
            inner.Add(p)
        loops.Add(inner)

    try:
        builder.AddFace(TessellatedFace(loops, material_id))
    except Exception:
        return None

    builder.CloseConnectedFaceSet()
    builder.Target = TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = TessellatedShapeBuilderFallback.Mesh
    try:
        builder.Build()
        geom = builder.GetBuildResult().GetGeometricalObjects()
    except Exception:
        return None
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


def _face_direction(normal):
    """Bepaal de horizontale richting (langs het wandvlak) uit de face-normal.

    dir = normalize(-n.Y, n.X) — loodrecht op de normaal, in het horizontale vlak.

    Args:
        normal: XYZ face-normal

    Returns:
        (dirx, diry) of None bij (bijna) horizontale normaal
    """
    dl = (normal.X * normal.X + normal.Y * normal.Y) ** 0.5
    if dl < 1e-9:
        return None
    return (-normal.Y / dl, normal.X / dl)


def _match_opening_to_face(opening, dirx, diry, base_s, smin, smax,
                           zmin, zmax, o0):
    """Match een opening aan een wandvlak via range-overlap + clamp.

    De opening-geometrie-vertices worden geprojecteerd op de FACE-richting
    (niet HandOrientation). Bij overlap wordt het gat geclampt tot binnen de
    face-grenzen (marge HOLE_CLAMP_MARGIN_FT naar binnen) zodat het gat altijd
    binnen de outer-loop valt.

    Args:
        opening: dict uit _opening_rect (vereist "vertices")
        dirx, diry: horizontale FACE-richting
        base_s: referentie-s van outer-loop oorsprong (o0)
        smin, smax: face s-range (langs dir)
        zmin, zmax: face z-range
        o0: XYZ oorsprong van de outer-loop (voor terug-projectie)

    Returns:
        (inner_corners list[XYZ], hole_area_m2 float) bij match, anders None
    """
    vertices = opening.get("vertices")
    if not vertices:
        return None

    s_vals = [v.X * dirx + v.Y * diry for v in vertices]
    os0 = min(s_vals)
    os1 = max(s_vals)
    oz0 = min(v.Z for v in vertices)
    oz1 = max(v.Z for v in vertices)

    # MATCH = range-overlap (strikte overlap, geen rand-aanraking)
    if not (os1 > smin and os0 < smax and oz1 > zmin and oz0 < zmax):
        return None

    # CLAMP tot face-grenzen (marge naar binnen)
    cs0 = max(os0, smin + HOLE_CLAMP_MARGIN_FT)
    cs1 = min(os1, smax - HOLE_CLAMP_MARGIN_FT)
    cz0 = max(oz0, zmin + HOLE_CLAMP_MARGIN_FT)
    cz1 = min(oz1, zmax - HOLE_CLAMP_MARGIN_FT)

    if not (cs1 > cs0 and cz1 > cz0):
        return None

    def pt(s, z):
        return XYZ(
            o0.X + dirx * (s - base_s),
            o0.Y + diry * (s - base_s),
            z,
        )

    # Hoekvolgorde (cs0,cz0),(cs0,cz1),(cs1,cz1),(cs1,cz0).
    # _build_directshape_face_with_holes draait deze om (reversed) voor
    # de tegengestelde winding t.o.v. de outer-loop.
    corners = [
        pt(cs0, cz0),
        pt(cs0, cz1),
        pt(cs1, cz1),
        pt(cs1, cz0),
    ]
    hole_area_m2 = (cs1 - cs0) * (cz1 - cz0) * SQFT_TO_M2
    return (corners, hole_area_m2)


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


def _opening_rect(insert, host_wall):
    """Bepaal de rechthoek-data van een opening in het wandvlak.

    Bouwt een platte rechthoek (4 hoeken) uit de bounding-extent van de
    opening-geometrie, geprojecteerd op het wandvlak. Onafhankelijk van
    ruimtecentrum of offset (dat doet de caller).

    Args:
        insert: FamilyInstance (deur/raam)
        host_wall: Wall element

    Returns:
        dict met keys:
            corners: list[XYZ] (4 hoeken, volgorde c0,c1,c2,c3)
            center: XYZ middelpunt
            area_m2: float oppervlak in m2
            direction: (dx, dy) horizontale wandrichting
        of None bij faal
    """
    direction = _opening_direction(insert, host_wall)
    if direction is None:
        return None
    dx, dy = direction

    vertices = _opening_vertices(insert)
    if not vertices:
        return None

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

    width_ft = smax - smin
    height_ft = zmax - zmin
    area_m2 = internal_to_sqm(width_ft * height_ft)

    center = XYZ((c0.X + c2.X) * 0.5, (c0.Y + c2.Y) * 0.5,
                 (zmin + zmax) * 0.5)

    return {
        "corners": [c0, c1, c2, c3],
        "center": center,
        "area_m2": area_m2,
        "direction": (dx, dy),
        "vertices": vertices,
    }


def _offset_corners_to_center(corners, direction, room_center, offset_mm):
    """Verschuif rechthoek-hoeken langs de wandnormaal richting ruimtecentrum.

    Args:
        corners: list[XYZ] (4 hoeken)
        direction: (dx, dy) horizontale wandrichting
        room_center: XYZ ruimtecentrum
        offset_mm: offset in mm

    Returns:
        list[XYZ] verschoven hoeken
    """
    dx, dy = direction
    nx = -dy
    ny = dx
    cx = sum(p.X for p in corners) / len(corners)
    cy = sum(p.Y for p in corners) / len(corners)
    to_center_x = room_center.X - cx
    to_center_y = room_center.Y - cy
    if (nx * to_center_x + ny * to_center_y) < 0:
        nx = -nx
        ny = -ny
    offset = offset_mm * MM_TO_FEET
    ox = nx * offset
    oy = ny * offset
    return [XYZ(p.X + ox, p.Y + oy, p.Z) for p in corners]


def _corners_to_triangles(corners):
    """Trianguleer een rechthoek (4 hoeken) naar 2 driehoeken."""
    if len(corners) < 4:
        return []
    c0, c1, c2, c3 = corners[0], corners[1], corners[2], corners[3]
    return [
        (c0, c1, c2),
        (c0, c2, c3),
    ]


def _build_opening_triangles(insert, host_wall, room_center):
    """Bouw de 2 driehoeken voor een opening-rechthoek in het wandvlak.

    Fallback-render: losse blauwe rechthoek met offset richting ruimtecentrum
    (gebruikt wanneer de opening niet in een wandvlak gesneden kon worden).

    Args:
        insert: FamilyInstance (deur/raam)
        host_wall: Wall element
        room_center: XYZ ruimtecentrum (voor offset-richting)

    Returns:
        list van (v0, v1, v2) tuples (2 driehoeken) of []
    """
    rect = _opening_rect(insert, host_wall)
    if rect is None:
        return []
    shifted = _offset_corners_to_center(
        rect["corners"], rect["direction"], room_center, OPENING_OFFSET_MM
    )
    return _corners_to_triangles(shifted)


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
# Shared-parameter aanmaken / binden
# =============================================================================
def _wv_create_param_options(name, type_key):
    """Maak ExternalDefinitionCreationOptions (version-safe, R2025).

    Args:
        name: parameternaam
        type_key: "text" of "number"

    Returns:
        ExternalDefinitionCreationOptions
    """
    from Autodesk.Revit.DB import ExternalDefinitionCreationOptions

    try:
        from Autodesk.Revit.DB import SpecTypeId
        if type_key == "number":
            return ExternalDefinitionCreationOptions(name, SpecTypeId.Number)
        if type_key == "yesno":
            return ExternalDefinitionCreationOptions(
                name, SpecTypeId.Boolean.YesNo
            )
        return ExternalDefinitionCreationOptions(name, SpecTypeId.String.Text)
    except (ImportError, AttributeError):
        pass

    from Autodesk.Revit.DB import ParameterType as RevitPT
    if type_key == "number":
        return ExternalDefinitionCreationOptions(name, RevitPT.Number)
    if type_key == "yesno":
        return ExternalDefinitionCreationOptions(name, RevitPT.YesNo)
    return ExternalDefinitionCreationOptions(name, RevitPT.Text)


def _wv_bind_parameter(doc, binding_map, definition, binding):
    """Bind een parameter aan het document (version-safe, R2025).

    Properties-palette groep: Data (PG_DATA / GroupTypeId.Data).
    """
    try:
        from Autodesk.Revit.DB import GroupTypeId
        binding_map.Insert(definition, binding, GroupTypeId.Data)
        return
    except (ImportError, AttributeError):
        pass

    try:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        binding_map.Insert(
            definition, binding, BuiltInParameterGroup.PG_DATA
        )
        return
    except Exception:
        pass

    binding_map.Insert(definition, binding)


def ensure_warmteverlies_parameters(doc):
    """Zorg dat de 6 warmteverlies_ shared parameters bestaan op GenericModel.

    Maakt (idempotent) deze instance shared parameters aan, gebonden aan
    de categorie Generic Models (OST_GenericModel — DirectShapes):
        warmteverlies_ruimte        (text)
        warmteverlies_naar_ruimte   (text)
        warmteverlies_grenstype     (text)
        warmteverlies_orientatie    (text)
        warmteverlies_oppervlak_m2  (number)
        warmteverlies_host_type     (text)

    Groep in het shared-parameter-bestand: "Berekeningen".
    Properties-palette groep: Data.

    MOET binnen een open Transaction worden aangeroepen (BindingMap.Insert
    + doc.Regenerate vereisen dit). render_room_boundaries draait al binnen
    de transactie van de pushbutton, dus deze functie opent géén eigen
    transactie maar bindt direct in de ambient transactie.

    Args:
        doc: Revit Document

    Returns:
        bool: True bij succes
    """
    global _wv_params_created
    if _wv_params_created:
        return True

    from Autodesk.Revit.DB import BuiltInCategory as _BIC

    app = doc.Application

    # Welke parameternamen zijn al gebonden?
    existing = set()
    binding_map = doc.ParameterBindings
    it = binding_map.ForwardIterator()
    while it.MoveNext():
        try:
            existing.add(it.Key.Name)
        except Exception:
            continue

    needed = [(n, t) for n, t in WV_PARAM_DEFS if n not in existing]
    if not needed:
        _wv_params_created = True
        return True

    original_spf = ""
    try:
        original_spf = app.SharedParametersFilename
    except Exception:
        pass

    try:
        temp_dir = os.environ.get(
            "TEMP",
            os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp"),
        )
        temp_path = os.path.join(temp_dir, "Warmteverlies_SharedParams.txt")

        if not os.path.exists(temp_path):
            f = open(temp_path, "w")
            f.close()

        app.SharedParametersFilename = temp_path
        def_file = app.OpenSharedParameterFile()

        # Zoek of maak de groep "Berekeningen"
        group = None
        for g in def_file.Groups:
            if g.Name == SHARED_PARAM_GROUP:
                group = g
                break
        if group is None:
            group = def_file.Groups.Create(SHARED_PARAM_GROUP)

        # Categorie-set: Generic Models
        cat_set = app.Create.NewCategorySet()
        gm_cat = doc.Settings.Categories.get_Item(_BIC.OST_GenericModel)
        cat_set.Insert(gm_cat)

        for param_name, type_key in needed:
            definition = None
            for d in group.Definitions:
                if d.Name == param_name:
                    definition = d
                    break
            if definition is None:
                opts = _wv_create_param_options(param_name, type_key)
                definition = group.Definitions.Create(opts)

            binding = app.Create.NewInstanceBinding(cat_set)
            _wv_bind_parameter(doc, binding_map, definition, binding)

        doc.Regenerate()
        _wv_params_created = True
        return True
    except Exception:
        return False
    finally:
        try:
            if original_spf:
                app.SharedParametersFilename = original_spf
        except Exception:
            pass


def ensure_afwerklaag_parameter(doc):
    """Zorg dat de Yes/No Type-parameter warmteverlies_afwerklaag bestaat.

    Type-binding (NewTypeBinding) op Walls, Floors, Roofs, Ceilings — het is
    een eigenschap van het constructie-Type, niet van de instance. Groep
    "Berekeningen" in het shared-parameter-bestand, palette-groep Data.

    Idempotent (BindingMap-check). MOET binnen een open Transaction draaien
    (zelfde ambient transactie als ensure_warmteverlies_parameters).

    Args:
        doc: Revit Document

    Returns:
        bool: True bij succes
    """
    global _afwerklaag_param_created
    if _afwerklaag_param_created:
        return True

    from Autodesk.Revit.DB import BuiltInCategory as _BIC

    app = doc.Application

    # Al gebonden?
    binding_map = doc.ParameterBindings
    it = binding_map.ForwardIterator()
    while it.MoveNext():
        try:
            if it.Key.Name == AFWERKLAAG_PARAM_NAME:
                _afwerklaag_param_created = True
                return True
        except Exception:
            continue

    original_spf = ""
    try:
        original_spf = app.SharedParametersFilename
    except Exception:
        pass

    try:
        temp_dir = os.environ.get(
            "TEMP",
            os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp"),
        )
        temp_path = os.path.join(temp_dir, "Warmteverlies_SharedParams.txt")
        if not os.path.exists(temp_path):
            f = open(temp_path, "w")
            f.close()

        app.SharedParametersFilename = temp_path
        def_file = app.OpenSharedParameterFile()

        group = None
        for g in def_file.Groups:
            if g.Name == SHARED_PARAM_GROUP:
                group = g
                break
        if group is None:
            group = def_file.Groups.Create(SHARED_PARAM_GROUP)

        # Categorie-set: Walls, Floors, Roofs, Ceilings
        cat_set = app.Create.NewCategorySet()
        for bic in (_BIC.OST_Walls, _BIC.OST_Floors,
                    _BIC.OST_Roofs, _BIC.OST_Ceilings):
            try:
                cat = doc.Settings.Categories.get_Item(bic)
                if cat is not None:
                    cat_set.Insert(cat)
            except Exception:
                continue

        definition = None
        for d in group.Definitions:
            if d.Name == AFWERKLAAG_PARAM_NAME:
                definition = d
                break
        if definition is None:
            opts = _wv_create_param_options(AFWERKLAAG_PARAM_NAME, "yesno")
            definition = group.Definitions.Create(opts)

        # TYPE-binding (eigenschap van het constructie-Type)
        binding = app.Create.NewTypeBinding(cat_set)
        _wv_bind_parameter(doc, binding_map, definition, binding)

        doc.Regenerate()
        _afwerklaag_param_created = True
        return True
    except Exception:
        return False
    finally:
        try:
            if original_spf:
                app.SharedParametersFilename = original_spf
        except Exception:
            pass


def bootstrap_afwerklaag_from_comments(doc):
    """Markeer constructie-Types als afwerklaag o.b.v. Type Comments.

    Voor elk Wall/Floor/Roof/Ceiling-TYPE waarvan Type Comments
    (ALL_MODEL_TYPE_COMMENTS) de substring "afwerk" bevat (case-insensitive)
    wordt warmteverlies_afwerklaag = 1 (Yes) gezet.

    Alleen Yes zetten, NOOIT unsetten — handmatig aangevinkte Types blijven
    behouden, de param wordt de filter-bron. Idempotent. MOET binnen een open
    Transaction draaien.

    Args:
        doc: Revit Document

    Returns:
        int: aantal Types dat (nieuw) op Yes is gezet
    """
    from Autodesk.Revit.DB import (
        BuiltInParameter, BuiltInCategory as _BIC, ElementType,
        FilteredElementCollector as _FEC,
    )

    type_cats = (
        _BIC.OST_Walls, _BIC.OST_Floors, _BIC.OST_Roofs, _BIC.OST_Ceilings,
    )

    count = 0
    seen = set()
    for bic in type_cats:
        try:
            collector = (
                _FEC(doc)
                .OfCategory(bic)
                .WhereElementIsElementType()
            )
        except Exception:
            continue

        for et in collector:
            try:
                tid = et.Id.IntegerValue
                if tid in seen:
                    continue
                seen.add(tid)

                cp = et.get_Parameter(
                    BuiltInParameter.ALL_MODEL_TYPE_COMMENTS
                )
                if cp is None or not cp.HasValue:
                    continue
                comment = cp.AsString()
                if not comment:
                    continue
                if AFWERKLAAG_COMMENT_PATTERN not in comment.lower():
                    continue

                ap = et.LookupParameter(AFWERKLAAG_PARAM_NAME)
                if ap is None or ap.IsReadOnly:
                    continue
                # Alleen Yes zetten als nog niet Yes (nooit unsetten)
                if ap.AsInteger() != 1:
                    ap.Set(1)
                    count += 1
            except Exception:
                continue

    return count


# =============================================================================
# Adjacency-bepaling per DirectShape
# =============================================================================
def _innermost_host_type_name(doc, host_ids):
    """Bepaal de Type-naam van het innerste host-element.

    Het innerste host-element is het eerste geldige element in de stapel
    (SEGC levert sub_faces van binnen naar buiten). Eén Type, niet de
    hele stapel.

    Args:
        doc: Revit Document
        host_ids: iterable van host element-id integers

    Returns:
        str: Type-naam, of "" als onbekend

    Let op: `ElementType.Name` gooit in de IronPython-2.7.12-engine van Revit
    een exception ("Name"). Gebruik daarom SYMBOL_NAME_PARAM, met
    `Element.Name.__get__(...)` als fallback (die descriptor-vorm werkt wél
    in deze engine).
    """
    from Autodesk.Revit.DB import BuiltInParameter, Element

    for hid in host_ids:
        try:
            host_el = doc.GetElement(ElementId(hid))
            if host_el is None:
                continue
            type_id = host_el.GetTypeId()
            if type_id is None or type_id.IntegerValue <= 0:
                continue
            type_el = doc.GetElement(type_id)
            if type_el is None:
                continue

            # 1) SYMBOL_NAME_PARAM (werkt voor WallType/FloorType/etc.)
            try:
                p = type_el.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if p is not None and p.HasValue:
                    val = p.AsString()
                    if val:
                        return val
            except Exception:
                pass

            # 2) Fallback: Element.Name descriptor-get
            try:
                val = Element.Name.__get__(type_el)
                if val:
                    return val
            except Exception:
                pass
        except Exception:
            continue
    return ""


def _room_phase(doc, room_element):
    """Haal de fase op waarin een ruimte staat (vereist voor GetRoomAtPoint).

    Args:
        doc: Revit Document
        room_element: Revit Room

    Returns:
        Phase-element of None
    """
    from Autodesk.Revit.DB import BuiltInParameter
    try:
        p = room_element.get_Parameter(BuiltInParameter.ROOM_PHASE)
        if p is None:
            return None
        phase_id = p.AsElementId()
        if phase_id is None or phase_id.IntegerValue <= 0:
            return None
        return doc.GetElement(phase_id)
    except Exception:
        return None


def _face_centroid(face, outer_pts):
    """Bepaal het zwaartepunt van een face.

    Eerste keus: gemiddelde van de outer-loop punten. Fallback: gemiddelde
    van de triangle-vertices (voor faces zonder bruikbare outer-loop).

    Args:
        face: Revit Face
        outer_pts: list[XYZ] outer-loop punten (mag leeg/None zijn)

    Returns:
        XYZ of None
    """
    pts = outer_pts
    if not pts:
        pts = []
        for tri in _face_to_triangles(face):
            pts.append(tri[0])
            pts.append(tri[1])
            pts.append(tri[2])
    if not pts:
        return None
    n = len(pts)
    sx = sum(p.X for p in pts)
    sy = sum(p.Y for p in pts)
    sz = sum(p.Z for p in pts)
    return XYZ(sx / n, sy / n, sz / n)


def _room_at(doc, xyz, phase):
    """GetRoomAtPoint met fase (zonder fase altijd None in deze modellen)."""
    if xyz is None or phase is None:
        return None
    try:
        return doc.GetRoomAtPoint(xyz, phase)
    except Exception:
        return None


def _outward_normal_at_face(doc, room_eid, face, normal, outer_pts, phase):
    """Bepaal zwaartepunt + geborgde naar-buiten-normaal van een grensvlak.

    De SEGC-normaal wijst doorgaans weg van het room-volume, maar dat wordt
    geverifieerd: een punt iets NAAR BINNEN (tegen de normaal in) moet de
    bron-ruimte opleveren. Zo niet, wordt de richting omgedraaid. Hergebruikt
    door zowel de adjacency-probe als de type-stapel-raycast (DRY).

    Args:
        doc: Revit Document
        room_eid: element-id (int) van de bron-ruimte
        face: Revit Face
        normal: XYZ SEGC face-normaal
        outer_pts: list[XYZ] outer-loop (voor zwaartepunt)
        phase: Phase-element van de bron-ruimte

    Returns:
        (centroid_XYZ, outward_normal_XYZ) of (None, None)
    """
    centroid = _face_centroid(face, outer_pts)
    if centroid is None:
        return (None, None)

    outx, outy, outz = normal.X, normal.Y, normal.Z

    # Richting borgen kan alleen met een fase (GetRoomAtPoint vereist die).
    if phase is not None:
        inward = ADJ_INWARD_CHECK_M * METER_TO_FEET
        p_in = XYZ(
            centroid.X - outx * inward,
            centroid.Y - outy * inward,
            centroid.Z - outz * inward,
        )
        room_in = _room_at(doc, p_in, phase)
        inward_hits_source = (
            room_in is not None and room_in.Id.IntegerValue == room_eid
        )
        if not inward_hits_source:
            p_in2 = XYZ(
                centroid.X + outx * inward,
                centroid.Y + outy * inward,
                centroid.Z + outz * inward,
            )
            room_in2 = _room_at(doc, p_in2, phase)
            if (room_in2 is not None
                    and room_in2.Id.IntegerValue == room_eid):
                # Normaal wees naar binnen -> omdraaien (nu naar buiten)
                outx, outy, outz = -outx, -outy, -outz

    return (centroid, XYZ(outx, outy, outz))


def _resolve_adjacency(doc, room_element, room_eid, face, normal, outer_pts,
                       all_rooms, heated_room_ids, phase, outward=None,
                       centroid=None):
    """Bepaal grenstype + buurruimte-label via geometrische probe.

    Vanuit het zwaartepunt van het grensvlak wordt naar buiten gestapt langs
    de (naar buiten wijzende) face-normaal. De eerste ruimte != bron-ruimte
    die GetRoomAtPoint oplevert is de echte buur aan DIT vlak — werkt voor
    wanden, vloeren (ruimte eronder) en daken/plafonds (ruimte erboven).

    Args:
        doc: Revit Document
        room_element: Revit Room (bron-ruimte)
        room_eid: element-id (int) van de bron-ruimte
        face: Revit Face
        normal: XYZ face-normaal (SEGC: wijst weg van het room-volume)
        outer_pts: list[XYZ] outer-loop (voor zwaartepunt)
        all_rooms: alle room-data dicts (voor heated-label fallback, niet vereist)
        heated_room_ids: set van verwarmde room element-ids
        phase: Phase-element van de bron-ruimte
        outward: XYZ geborgde naar-buiten-normaal (optioneel; anders berekend)
        centroid: XYZ face-zwaartepunt (optioneel; anders berekend)

    Returns:
        (grenstype, naar_ruimte_label, cutoff_m)
            grenstype in {exterior, adjacent_room, unheated_space}
            (ground wordt door de caller bepaald voor vloeren op maaiveld)
            cutoff_m: offset (meter) waarop een ruimte != bron is gevonden
                      (de verre wandzijde); None bij echte buitenlucht zonder
                      gevonden buurruimte. Wordt door de type-stapel-raycast
                      gebruikt om hits voorbij de wand af te kappen.
    """
    if centroid is None or outward is None:
        centroid, outward = _outward_normal_at_face(
            doc, room_eid, face, normal, outer_pts, phase
        )
    if centroid is None or outward is None or phase is None:
        return ("exterior", "BUITEN", None)

    outx, outy, outz = outward.X, outward.Y, outward.Z

    # Naar buiten stappen door de constructie-dikte heen
    for off_m in ADJ_PROBE_OFFSETS_M:
        off = off_m * METER_TO_FEET
        p = XYZ(
            centroid.X + outx * off,
            centroid.Y + outy * off,
            centroid.Z + outz * off,
        )
        nbr = _room_at(doc, p, phase)
        if nbr is None:
            continue
        nbr_eid = nbr.Id.IntegerValue
        if nbr_eid == room_eid:
            # Nog in de bron-ruimte (dunne sliver) -> verder stappen
            continue

        nbr_name = _get_room_name_safe(nbr)

        # Buitenlucht-ruimte (bv. "Buiten") telt als exterior, niet als
        # unheated_space. Buitenlucht is het eindpunt van de probe — maar de
        # offset is wel een geldige cutoff (verre wandzijde).
        nbr_name_norm = nbr_name.strip().lower()
        if any(pat in nbr_name_norm for pat in OUTDOOR_ROOM_NAME_PATTERNS):
            return ("exterior", "BUITEN", off_m)

        adj_label = "{0} {1}".format(
            _get_room_number_safe(nbr),
            nbr_name,
        ).strip()
        if nbr_eid in heated_room_ids:
            return ("adjacent_room", adj_label, off_m)
        return ("unheated_space", "ONVERWARMD:" + adj_label, off_m)

    # Niets gevonden -> echte buitenlucht, geen buurruimte-afstand beschikbaar
    return ("exterior", "BUITEN", None)


def _get_room_number_safe(room):
    """Room-nummer (BuiltInParameter), met fallback op element-id."""
    from Autodesk.Revit.DB import BuiltInParameter
    try:
        p = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        if p is not None and p.HasValue:
            v = p.AsString()
            if v:
                return v
    except Exception:
        pass
    return str(room.Id.IntegerValue)


def _get_room_name_safe(room):
    """Room-naam (BuiltInParameter), met lege fallback."""
    from Autodesk.Revit.DB import BuiltInParameter
    try:
        p = room.get_Parameter(BuiltInParameter.ROOM_NAME)
        if p is not None and p.HasValue:
            v = p.AsString()
            if v:
                return v
    except Exception:
        pass
    return ""


# =============================================================================
# Type-stapel (raycast door de constructie heen)
# =============================================================================
def _element_type_name(element_doc, element):
    """Type-naam van een element (host of linked), engine-robuust.

    Gebruikt SYMBOL_NAME_PARAM (`.Name` faalt in de Revit-IronPython-engine),
    met Element.Name.__get__ als fallback.

    Args:
        element_doc: Document waarin het element leeft (host of link)
        element: Revit Element

    Returns:
        str: Type-naam of "" als onbekend
    """
    from Autodesk.Revit.DB import BuiltInParameter, Element
    try:
        type_id = element.GetTypeId()
        if type_id is None or type_id.IntegerValue <= 0:
            return ""
        type_el = element_doc.GetElement(type_id)
        if type_el is None:
            return ""
        try:
            p = type_el.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p is not None and p.HasValue:
                v = p.AsString()
                if v:
                    return v
        except Exception:
            pass
        try:
            v = Element.Name.__get__(type_el)
            if v:
                return v
        except Exception:
            pass
    except Exception:
        pass
    return ""


def _is_afwerklaag_type(element_doc, element):
    """Is het Type van dit element gemarkeerd als afwerklaag (Yes/No)?

    Leest warmteverlies_afwerklaag op het Type. Param afwezig (bv. in een
    linked doc dat de param niet kent) of niet Yes -> False (= structureel,
    wordt niet uit de stapel gefilterd).

    Args:
        element_doc: Document waarin het element leeft
        element: Revit Element

    Returns:
        bool
    """
    try:
        type_id = element.GetTypeId()
        if type_id is None or type_id.IntegerValue <= 0:
            return False
        type_el = element_doc.GetElement(type_id)
        if type_el is None:
            return False
        p = type_el.LookupParameter(AFWERKLAAG_PARAM_NAME)
        if p is None:
            return False
        return p.AsInteger() == 1
    except Exception:
        return False


def _type_stack_for_face(doc, intersector, centroid, outward, cutoff_m=None):
    """Bepaal de geordende type-stapel die een straal naar buiten passeert.

    Schiet een straal vanuit (iets binnen) het face-zwaartepunt langs de
    naar-buiten-normaal en verzamelt de constructie-element-Types op afstand.

    De ray levert per element TWEE faces (voor- + achtervlak), dus de
    luchtspleet-detectie is ELEMENT-bewust: hits worden per element_id
    gegroepeerd tot een span [near, far]. Een dik element (bv. 120mm PIR)
    is één span, geen valse spleet. De luchtspleet wordt gemeten als
    `volgend_element.near - huidig_element.far` (echte spouw tussen twee
    elementen).

    Afkapping (cruciaal — anders schiet de straal de buurruimte in):
    - luchtspleet: break wanneer de inter-element-spleet > TYPE_STACK_MAX_AIRGAP_M
      (verre wandzijde / eerste echte spouw of ruimte-overgang).
    - `cutoff_m` (buurruimte-afstand): neem geen element mee waarvan `near`
      voorbij cutoff_m (+ marge) ligt.

    Afwerklaag-elementen worden volledig overgeslagen; de spleet wordt dan
    gemeten tussen de overgebleven STRUCTURELE buur-elementen.

    Daarna pas worden opeenvolgende duplicaten samengevoegd.

    Args:
        doc: Revit host Document
        intersector: ReferenceIntersector (uit _create_intersector)
        centroid: XYZ face-zwaartepunt
        outward: XYZ geborgde naar-buiten-normaal (zelfde als de probe gebruikte)
        cutoff_m: buurruimte-afstand (meter) of None

    Returns:
        str: "Type1 > Type2 > ..." of "" als geen constructie geraakt
    """
    if intersector is None or centroid is None or outward is None:
        return ""

    # Iets naar binnen starten zodat de eerste (binnenste) laag meekomt.
    back = 0.05 * METER_TO_FEET
    origin = XYZ(
        centroid.X - outward.X * back,
        centroid.Y - outward.Y * back,
        centroid.Z - outward.Z * back,
    )

    try:
        refs = intersector.Find(origin, outward)
    except Exception:
        return ""
    if refs is None:
        return ""

    # Bovengrens: buurruimte-cutoff (+marge) of de globale RAY_MAX_DIST_M.
    if cutoff_m is not None:
        max_dist_m = cutoff_m + TYPE_STACK_CUTOFF_MARGIN_M
    else:
        max_dist_m = RAY_MAX_DIST_M

    # --- Hits per element_id groeperen tot spans [near, far] ---
    # element_id -> dict(near, far, tname, afwerk). De ray levert per element
    # twee faces (voor/achter); door op element_id te groeperen wordt een dik
    # element één span i.p.v. twee hits die een valse luchtspleet vormen.
    elem_spans = {}
    for rwc in refs:
        try:
            prox_m = rwc.Proximity * FEET_TO_M
            if prox_m > max_dist_m:
                continue
            hit = _resolve_hit_element(rwc.GetReference(), doc)
            if hit is None:
                continue
            cat_id = hit.get("category_id")
            # Eigen WV_BND Generic Models + meubels etc. wegfilteren
            if cat_id in IGNORE_CATEGORIES:
                continue
            # Alleen echte constructie (Wall/Floor/Roof/Ceiling); filtert
            # openingen automatisch weg (niet in CONSTRUCTION_CATEGORIES).
            if cat_id not in CONSTRUCTION_CATEGORIES:
                continue

            eid = hit.get("element_id")
            if eid is None:
                continue

            existing = elem_spans.get(eid)
            if existing is not None:
                # Span uitbreiden (tweede face van hetzelfde element)
                if prox_m < existing["near"]:
                    existing["near"] = prox_m
                if prox_m > existing["far"]:
                    existing["far"] = prox_m
                continue

            afwerk = _is_afwerklaag_type(
                hit.get("element_doc"), hit.get("element")
            )
            tname = _element_type_name(
                hit.get("element_doc"), hit.get("element")
            )
            elem_spans[eid] = {
                "near": prox_m,
                "far": prox_m,
                "tname": tname,
                "afwerk": afwerk,
            }
        except Exception:
            continue

    if not elem_spans:
        return ""

    # Elementen op near-afstand sorteren
    elements = sorted(elem_spans.values(), key=lambda e: e["near"])

    cutoff_bound = None
    if cutoff_m is not None:
        cutoff_bound = cutoff_m + TYPE_STACK_CUTOFF_MARGIN_M

    # --- Truncatie op element-niveau ---
    # Luchtspleet = volgend_element.near - huidig_element.far (alleen tussen
    # STRUCTURELE elementen gemeten; afwerklaag-elementen worden overgeslagen
    # zonder de spleet kunstmatig te vergroten).
    # Cutoff = element meenemen zolang near <= cutoff_bound.
    ordered = []
    prev_far = None
    for el in elements:
        # Bovengrens buurruimte-cutoff (op begin van het element)
        if cutoff_bound is not None and el["near"] > cutoff_bound:
            break

        # Afwerklaag: element overslaan; prev_far NIET bijwerken, zodat de
        # spleet straks tussen de structurele buren wordt gemeten.
        if el["afwerk"]:
            continue

        # Luchtspleet t.o.v. de vorige STRUCTURELE laag. De air-gap-break geldt
        # alleen wanneer er GEEN cutoff bekend is (cutoff_m is None): met een
        # cutoff_m (buur/exterieur-afstand) bewaakt cutoff_bound de assemblage-
        # grens al, en is een grote spleet intern (tapse dakspie, ventilatie-
        # spouw, gevelbekleding op afstand) — die mag de stapel niet afkappen.
        if prev_far is not None and cutoff_m is None:
            gap = el["near"] - prev_far
            if gap > TYPE_STACK_MAX_AIRGAP_M:
                break

        tname = el["tname"]
        if tname:
            # Opeenvolgende duplicaten samenvoegen
            if not ordered or ordered[-1] != tname:
                ordered.append(tname)

        prev_far = el["far"]

    return " > ".join(ordered)


def _face_sample_points(face, n=3):
    """Genereer tot n*n sample-punten verdeeld over het vlak (binnen de
    face-grenzen). Gebruikt om de type-stapel-raycast vanuit meerdere punten
    te schieten i.p.v. alleen het zwaartepunt — dat laatste kan op een
    deur/raam-opening vallen waardoor de straal de wand mist.
    """
    pts = []
    try:
        bb = face.GetBoundingBox()
        umin = bb.Min.U
        vmin = bb.Min.V
        umax = bb.Max.U
        vmax = bb.Max.V
        for i in range(n):
            for j in range(n):
                u = umin + (umax - umin) * (i + 0.5) / n
                v = vmin + (vmax - vmin) * (j + 0.5) / n
                uv = UV(u, v)
                try:
                    if face.IsInside(uv):
                        pts.append(face.Evaluate(uv))
                except Exception:
                    continue
    except Exception:
        pass
    return pts


def _type_stack_for_face_sampled(doc, intersector, face, outward, cutoff_m,
                                 fallback_centroid=None):
    """Bepaal de type-stapel via meerdere sample-punten over het vlak.

    Schiet de type-stapel-raycast vanuit een raster van punten op het vlak en
    kiest de meest-voorkomende niet-lege stapel (meerderheid). Punten die op
    een deur/raam-opening vallen (en dus de buurruimte in schieten) zijn in de
    minderheid en worden zo weggestemd. Valt terug op het zwaartepunt als er
    geen sample-punten binnen het vlak liggen.
    """
    samples = _face_sample_points(face)
    if fallback_centroid is not None:
        samples.append(fallback_centroid)
    counts = {}
    for pt in samples:
        try:
            s = _type_stack_for_face(doc, intersector, pt, outward, cutoff_m)
        except Exception:
            s = ""
        if s:
            counts[s] = counts.get(s, 0) + 1
    if not counts:
        return ""
    # Meest voorkomende; gelijke stand -> meeste lagen (meeste '>').
    best = sorted(
        counts.items(),
        key=lambda kv: (kv[1], kv[0].count(">")),
        reverse=True,
    )
    return best[0][0]


def _first_3d_view_family_type_id(doc):
    """Vind het eerste ViewFamilyType voor 3D-views (voor de temp-view).

    Returns:
        ElementId of None
    """
    try:
        for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
            try:
                if vft.ViewFamily == ViewFamily.ThreeDimensional:
                    return vft.Id
            except Exception:
                continue
    except Exception:
        pass
    return None


def _create_temp_3d_view(doc):
    """Maak een tijdelijke schone isometrische 3D-view (zonder section box).

    Cruciaal: ReferenceIntersector geeft 0 hits in een view met actieve
    section box -> we maken een verse view zonder box.

    Returns:
        View3D of None
    """
    vft_id = _first_3d_view_family_type_id(doc)
    if vft_id is None:
        return None
    try:
        v3d = View3D.CreateIsometric(doc, vft_id)
        try:
            v3d.IsSectionBoxActive = False
        except Exception:
            pass
        return v3d
    except Exception:
        return None


def _set_wv_params(ds, ruimte, naar_ruimte, grenstype, orient_label,
                   area_m2, host_type, type_stapel=""):
    """Zet de 7 warmteverlies_ parameters op een DirectShape (best-effort).

    Number-param met float, text-params met string. Ontbrekende param wordt
    netjes overgeslagen (geen crash).

    Args:
        ds: DirectShape
        ruimte: str bron-ruimte label
        naar_ruimte: str buurruimte / BUITEN / GROND / ONVERWARMD:...
        grenstype: str classificatie
        orient_label: str dak/wand/vloer/opening
        area_m2: float netto oppervlak in m2
        host_type: str host Type-naam (innerste)
        type_stapel: str geordende type-stapel ("T1 > T2 > ...")
    """
    if ds is None:
        return

    text_vals = (
        ("warmteverlies_ruimte", ruimte),
        ("warmteverlies_naar_ruimte", naar_ruimte),
        ("warmteverlies_grenstype", grenstype),
        ("warmteverlies_orientatie", orient_label),
        ("warmteverlies_host_type", host_type),
        ("warmteverlies_type_stapel", type_stapel),
    )
    for pname, pval in text_vals:
        try:
            p = ds.LookupParameter(pname)
            if p is not None and not p.IsReadOnly:
                p.Set(pval if pval is not None else "")
        except Exception:
            continue

    try:
        p = ds.LookupParameter("warmteverlies_oppervlak_m2")
        if p is not None and not p.IsReadOnly:
            p.Set(round(float(area_m2), 2))
    except Exception:
        pass


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
        "netto_wall_m2": 0.0,
        "holes_cut": 0,
        "type_stack_failed": 0,
    }

    # --- Shared parameters borgen (binnen ambient transactie van pushbutton) ---
    ensure_warmteverlies_parameters(doc)
    # Yes/No Type-param + bootstrap afwerklagen uit Type Comments
    ensure_afwerklaag_parameter(doc)
    bootstrap_afwerklaag_from_comments(doc)

    # --- Verwarmde-ruimte set (zelfde predicate als heated_only-filter) ---
    heated_room_ids = set(
        rd["element_id"] for rd in rooms if rd.get("is_heated")
    )

    # --- Temp 3D-view + intersector voor type-stapel-raycast (één keer) ---
    # Schone view zonder section box: ReferenceIntersector geeft anders 0 hits.
    temp_view = _create_temp_3d_view(doc)
    intersector = None
    if temp_view is not None:
        try:
            intersector = _create_intersector(doc, temp_view)
        except Exception:
            intersector = None

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

        # Bron-ruimte label + element-id voor adjacency-parameters
        room_eid = room_data.get("element_id")
        room_label = "{0} {1}".format(room_number, name).strip()

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

        # Fase van deze ruimte (vereist voor GetRoomAtPoint in de adjacency-probe)
        room_phase = _room_phase(doc, room_element)

        # Wanden waarvoor we openingen moeten zoeken (niet-curtain hosts)
        host_wall_ids = set()

        # Niet-curtain wandvlakken die we pas renderen NA opening-detectie,
        # zodat we openingen er als gat uit kunnen snijden.
        deferred_walls = []

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
                        # Check vliesgevel/glaswand onder de hosts
                        is_curtain = False
                        for hid in hosts:
                            host_el = doc.GetElement(ElementId(hid))
                            if host_el is not None and _is_glass_wall(doc, host_el):
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

                # Niet-curtain wandvlakken uitstellen (gat-snijden later)
                if orient == "wall":
                    deferred_walls.append({
                        "face": face,
                        "normal": normal,
                        "area_m2": area_m2,
                        "outer_pts": _face_outer_loop(face),
                        "host_ids": set(hosts),
                    })
                    continue

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

                    # --- Zwaartepunt + naar-buiten-normaal éénmaal bepalen
                    # (gedeeld door adjacency-probe en type-stapel-raycast) ---
                    f_outer = _face_outer_loop(face)
                    f_centroid, f_outward = _outward_normal_at_face(
                        doc, room_eid, face, normal, f_outer, room_phase
                    )

                    # --- Adjacency-parameters zetten (geometrische probe) ---
                    grenstype, naar_ruimte, cutoff_m = _resolve_adjacency(
                        doc, room_element, room_eid, face, normal,
                        f_outer, rooms, heated_room_ids, room_phase,
                        outward=f_outward, centroid=f_centroid,
                    )
                    # Vloer op maaiveld -> GROND (zelfde drempel als
                    # adjacent_detector: level_elevation_m < 0.5). Alleen als
                    # de probe geen buurruimte vond (terugval exterior).
                    if (orient == "bot" and grenstype == "exterior"
                            and room_data.get("level_elevation_m", 0.0) < 0.5):
                        grenstype = "ground"
                        naar_ruimte = "GROND"

                    # --- Type-stapel via raycast (afgekapt op buur-cutoff) ---
                    type_stapel = ""
                    try:
                        type_stapel = _type_stack_for_face(
                            doc, intersector, f_centroid, f_outward, cutoff_m
                        )
                    except Exception:
                        stats["type_stack_failed"] += 1

                    _set_wv_params(
                        ds,
                        room_label,
                        naar_ruimte,
                        grenstype,
                        ORIENT_LABEL.get(orient, orient),
                        area_m2,
                        _innermost_host_type_name(doc, hosts),
                        type_stapel,
                    )
            except Exception:
                stats["faces_failed"] += 1
                continue

        # --- Openingen verzamelen voor niet-curtain host-wanden ---
        # key = insert IntegerValue, value = dict(rect-data + wall + insert)
        room_openings = {}
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
                        if ins_key in room_openings:
                            continue
                        insert = doc.GetElement(ins_id)
                        if insert is None:
                            continue
                        cat = insert.Category
                        if cat is None or cat.Name not in OPENING_CATEGORIES:
                            continue
                        rect = _opening_rect(insert, wall)
                        if rect is None:
                            continue
                        rect["insert"] = insert
                        rect["wall"] = wall
                        rect["host_wall_id"] = wid
                        room_openings[ins_key] = rect
                    except Exception:
                        continue

        # --- Wandvlakken renderen met gaten uit gematchte openingen ---
        consumed = set()
        for wd in deferred_walls:
            try:
                face = wd["face"]
                normal = wd["normal"]
                outer_pts = wd["outer_pts"]
                wall_host_ids = wd.get("host_ids", set())

                comment = "{0} {1} wall".format(COMMENTS_PREFIX, room_number)

                # Zwaartepunt + naar-buiten-normaal éénmaal (adjacency + stapel)
                w_centroid, w_outward = _outward_normal_at_face(
                    doc, room_eid, face, normal, outer_pts, room_phase
                )

                # Adjacency (geometrische probe) + host-type éénmaal per wandvlak
                wall_grenstype, wall_naar, wall_cutoff_m = _resolve_adjacency(
                    doc, room_element, room_eid, face, normal, outer_pts,
                    rooms, heated_room_ids, room_phase,
                    outward=w_outward, centroid=w_centroid,
                )
                wall_host_type = _innermost_host_type_name(doc, wall_host_ids)

                # Type-stapel via raycast (afgekapt op buur-cutoff, 1x per wand)
                wall_type_stapel = ""
                try:
                    wall_type_stapel = _type_stack_for_face_sampled(
                        doc, intersector, face, w_outward, wall_cutoff_m,
                        w_centroid
                    )
                except Exception:
                    stats["type_stack_failed"] += 1

                # Vliesgevel achter een room-separation-line: lege stapel +
                # curtain-paneel vlak erachter -> dit vlak IS de glas-gevel,
                # niet een dichte wand. Render als blauwe opening en ga door.
                if not wall_type_stapel and _curtain_glass_behind(
                    doc, intersector, w_centroid, w_outward
                ):
                    triangles = _face_to_triangles(face)
                    gds = _build_directshape_from_triangles(
                        doc, triangles, material_ids[MAT_OPEN[0]],
                        "{0} {1} vlies".format(COMMENTS_PREFIX, room_number)
                    )
                    if gds is not None:
                        stats["open"] += 1
                        _set_wv_params(
                            gds, room_label, "BUITEN", "exterior",
                            ORIENT_LABEL.get("vliesgevel", "opening"),
                            wd["area_m2"], wall_host_type, "vliesgevel",
                        )
                    continue

                # Horizontale FACE-richting + face s/z-ranges uit outer-loop
                fdir = _face_direction(normal)
                if outer_pts and fdir is not None:
                    dirx, diry = fdir
                    o0 = outer_pts[0]
                    base_s = o0.X * dirx + o0.Y * diry
                    s_vals = [p.X * dirx + p.Y * diry for p in outer_pts]
                    smin = min(s_vals)
                    smax = max(s_vals)
                    zmin = min(p.Z for p in outer_pts)
                    zmax = max(p.Z for p in outer_pts)

                    # Match openingen: host-wall-id-koppeling + range-overlap
                    matched = []   # (ins_key, inner_corners, hole_area_m2, rect)
                    for ins_key, rect in room_openings.items():
                        if ins_key in consumed:
                            continue
                        # Koppel via host-wall-id (insert hoort bij wall die ook
                        # host van DEZE face is) — voorkomt false matches op
                        # verre segmenten van lange wanden.
                        if rect.get("host_wall_id") not in wall_host_ids:
                            continue
                        hit = _match_opening_to_face(
                            rect, dirx, diry, base_s,
                            smin, smax, zmin, zmax, o0
                        )
                        if hit is not None:
                            matched.append((ins_key, hit[0], hit[1], rect))
                else:
                    matched = []

                holes = [m[1] for m in matched]

                ds = None
                if outer_pts and holes:
                    ds = _build_directshape_face_with_holes(
                        doc, outer_pts, holes,
                        material_ids[MAT_WALL[0]], normal, comment
                    )

                if ds is None:
                    # Fallback: volle gele wand zonder gat (altijd zichtbaar)
                    triangles = _face_to_triangles(face)
                    ds = _build_directshape_from_triangles(
                        doc, triangles, material_ids[MAT_WALL[0]], comment
                    )
                    # Geen gaten gesneden -> bruto area, geen netto-aftrek
                    if ds is not None:
                        stats["wall"] += 1
                        stats["netto_wall_m2"] += wd["area_m2"]
                        _set_wv_params(
                            ds, room_label, wall_naar, wall_grenstype,
                            ORIENT_LABEL.get("wall", "wand"),
                            wd["area_m2"], wall_host_type, wall_type_stapel,
                        )
                    continue

                # Gat(en) succesvol gesneden
                stats["wall"] += 1
                opening_area = sum(m[2] for m in matched)
                netto = wd["area_m2"] - opening_area
                if netto < 0.0:
                    netto = 0.0
                stats["netto_wall_m2"] += netto

                _set_wv_params(
                    ds, room_label, wall_naar, wall_grenstype,
                    ORIENT_LABEL.get("wall", "wand"),
                    netto, wall_host_type, wall_type_stapel,
                )

                # Render gematchte openingen als blauwe rechthoek IN het gat
                for ins_key, inner_corners, hole_area, rect in matched:
                    consumed.add(ins_key)
                    stats["holes_cut"] += 1
                    shifted = _offset_corners_to_center(
                        inner_corners, rect["direction"], center,
                        HOLE_OPENING_OFFSET_MM
                    )
                    tris = _corners_to_triangles(shifted)
                    open_comment = "{0} {1} opening".format(
                        COMMENTS_PREFIX, room_number
                    )
                    ods = _build_directshape_from_triangles(
                        doc, tris, material_ids[MAT_OPEN[0]], open_comment
                    )
                    if ods is not None:
                        rendered_openings.add(ins_key)
                        stats["openings"] += 1
                        # Opening grenst altijd aan buiten (deur/raam in wand)
                        _set_wv_params(
                            ods, room_label, "BUITEN", "exterior",
                            ORIENT_LABEL.get("open", "opening"),
                            hole_area, wall_host_type,
                            _element_type_name(doc, rect["insert"]),
                        )
            except Exception:
                stats["faces_failed"] += 1
                continue

        # --- Niet-gematchte openingen: fallback losse blauwe rechthoek ---
        if show_openings:
            for ins_key, rect in room_openings.items():
                if ins_key in consumed:
                    continue
                if ins_key in rendered_openings:
                    continue
                try:
                    tris = _build_opening_triangles(
                        rect["insert"], rect["wall"], center
                    )
                    if not tris:
                        continue
                    comment = "{0} {1} opening".format(
                        COMMENTS_PREFIX, room_number
                    )
                    ods = _build_directshape_from_triangles(
                        doc, tris, material_ids[MAT_OPEN[0]], comment
                    )
                    if ods is not None:
                        rendered_openings.add(ins_key)
                        stats["openings"] += 1
                        host_ids = []
                        wid = rect.get("host_wall_id")
                        if wid is not None:
                            host_ids.append(wid)
                        _set_wv_params(
                            ods, room_label, "BUITEN", "exterior",
                            ORIENT_LABEL.get("open", "opening"),
                            rect.get("area_m2", 0.0),
                            _innermost_host_type_name(doc, host_ids),
                            _element_type_name(doc, rect["insert"]),
                        )
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

    # --- Temp 3D-view opruimen (geen "3D View N"-rommel in het project) ---
    if temp_view is not None:
        try:
            doc.Delete(temp_view.Id)
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
