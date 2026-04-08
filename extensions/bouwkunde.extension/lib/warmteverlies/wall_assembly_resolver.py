# -*- coding: utf-8 -*-
"""Wall Assembly Resolver — detecteert gestapelde wanden voor U-waarde.

In Revit worden buitenwanden vaak als aparte lagen getekend:
binnenblad + isolatie + buitenblad. De SpatialElementGeometryCalculator
rapporteert alleen de wand die direct aan de ruimte grenst.

Deze module detecteert parallelle wanden die samen een constructie-
assemblage vormen en retourneert ze als geordende lijst [binnen -> buiten].

IronPython 2.7 — geen f-strings, geen type hints.
"""
import math

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Wall,
    Line,
    XYZ,
)

from warmteverlies.unit_utils import internal_to_meters
from warmteverlies.constants import (
    MAX_ASSEMBLY_GAP_M,
    PARALLEL_COS_TOLERANCE,
    MIN_OVERLAP_FRACTION,
    MAX_ASSEMBLY_DEPTH,
    FEET_TO_M,
)


def collect_all_walls(doc):
    """Verzamel alle wanden met pre-berekende geometrie.

    Roep dit eenmalig aan per export. Filtert curtain walls en
    wanden zonder LocationCurve.

    Args:
        doc: Revit Document

    Returns:
        list[dict]: Per wand: element, midpoint, direction, normal,
                    width_m, bbox, element_id
    """
    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
    )

    walls_data = []
    for wall in collector:
        if not isinstance(wall, Wall):
            continue

        # Curtain walls overslaan — geen CompoundStructure
        try:
            wall_type = doc.GetElement(wall.GetTypeId())
            if wall_type and wall_type.Kind.ToString() == "Curtain":
                continue
        except Exception:
            pass

        loc = wall.Location
        if loc is None:
            continue

        try:
            curve = loc.Curve
        except Exception:
            continue

        if curve is None:
            continue

        # Alleen rechte wanden (Line) in V1
        if not isinstance(curve, Line):
            continue

        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)

        midpoint = XYZ(
            (start.X + end.X) / 2.0,
            (start.Y + end.Y) / 2.0,
            (start.Z + end.Z) / 2.0,
        )

        direction = XYZ(
            end.X - start.X,
            end.Y - start.Y,
            end.Z - start.Z,
        )
        dir_length = math.sqrt(
            direction.X ** 2 + direction.Y ** 2 + direction.Z ** 2
        )
        if dir_length < 1e-9:
            continue

        direction = XYZ(
            direction.X / dir_length,
            direction.Y / dir_length,
            direction.Z / dir_length,
        )

        # Normal: 90 graden rotatie in XY-vlak (rechterhand)
        normal = XYZ(-direction.Y, direction.X, 0.0)

        width_m = internal_to_meters(wall.Width)
        wall_length_m = internal_to_meters(dir_length)

        bbox = wall.get_BoundingBox(None)

        walls_data.append({
            "element": wall,
            "element_id": wall.Id.IntegerValue,
            "midpoint": midpoint,
            "direction": direction,
            "normal": normal,
            "width_m": width_m,
            "length_m": wall_length_m,
            "curve_start": start,
            "curve_end": end,
            "bbox": bbox,
        })

    return walls_data


def resolve_wall_assembly(doc, boundary_wall, face_normal, all_walls_data):
    """Detecteer parallelle wanden die samen een assemblage vormen.

    Begint bij de room-begrenzende wand en stapt in de face_normal
    richting naar buiten. Zoekt telkens de dichtstbijzijnde parallelle
    wand binnen de gap-tolerantie.

    Args:
        doc: Revit Document
        boundary_wall: Revit Wall element (room-begrenzend)
        face_normal: tuple (x, y, z) — richting naar buiten
        all_walls_data: Resultaat van collect_all_walls()

    Returns:
        list[Wall]: Geordende assemblage [binnen -> buiten],
                    altijd minimaal 1 element
    """
    if boundary_wall is None:
        return []

    boundary_id = boundary_wall.Id.IntegerValue
    normal_vec = XYZ(face_normal[0], face_normal[1], face_normal[2])

    # Normaliseer face normal (XY-vlak)
    normal_xy_len = math.sqrt(normal_vec.X ** 2 + normal_vec.Y ** 2)
    if normal_xy_len < 1e-9:
        # Verticale normal (vloer/plafond) — geen assembly detectie
        return [boundary_wall]

    search_dir = XYZ(
        normal_vec.X / normal_xy_len,
        normal_vec.Y / normal_xy_len,
        0.0,
    )

    # Zoek de boundary wall data
    current_data = None
    for wd in all_walls_data:
        if wd["element_id"] == boundary_id:
            current_data = wd
            break

    if current_data is None:
        return [boundary_wall]

    assembly = [boundary_wall]
    used_ids = {boundary_id}
    current = current_data

    for _ in range(MAX_ASSEMBLY_DEPTH - 1):
        candidate = _find_next_wall(
            current, search_dir, all_walls_data, used_ids
        )

        if candidate is None:
            break

        assembly.append(candidate["element"])
        used_ids.add(candidate["element_id"])
        current = candidate

    return assembly


def _find_next_wall(current_data, search_dir, all_walls_data, used_ids):
    """Zoek de dichtstbijzijnde parallelle wand in de zoekrichting.

    Args:
        current_data: dict met huidige wandgegevens
        search_dir: XYZ zoekrichting (genormaliseerd, XY-vlak)
        all_walls_data: Alle wanden
        used_ids: Set van al gebruikte element IDs

    Returns:
        dict of None: Beste kandidaat wanddata
    """
    current_mid = current_data["midpoint"]
    current_dir = current_data["direction"]
    current_half_width = current_data["width_m"] / 2.0

    # Buitenvlak van huidige wand
    outer_point = XYZ(
        current_mid.X + search_dir.X * (current_half_width / FEET_TO_M),
        current_mid.Y + search_dir.Y * (current_half_width / FEET_TO_M),
        current_mid.Z,
    )

    max_gap_ft = MAX_ASSEMBLY_GAP_M / FEET_TO_M
    best_candidate = None
    best_distance = float("inf")

    for wd in all_walls_data:
        if wd["element_id"] in used_ids:
            continue

        # Snelle bbox pre-filter
        if not _bboxes_overlap_xy(current_data, wd, max_gap_ft):
            continue

        # Parallel check: abs(dot product) van richtingen >= tolerantie
        dot = abs(
            current_dir.X * wd["direction"].X
            + current_dir.Y * wd["direction"].Y
        )
        if dot < PARALLEL_COS_TOLERANCE:
            continue

        # Richting check: kandidaat moet in de zoekrichting liggen
        delta = XYZ(
            wd["midpoint"].X - current_mid.X,
            wd["midpoint"].Y - current_mid.Y,
            0.0,
        )
        dot_direction = (
            delta.X * search_dir.X + delta.Y * search_dir.Y
        )
        if dot_direction < 0:
            # Kandidaat ligt achter de huidige wand
            continue

        # Gap check: afstand tussen buitenvlak huidige en binnenvlak kandidaat
        cand_half_width = wd["width_m"] / 2.0
        cand_inner_point = XYZ(
            wd["midpoint"].X - search_dir.X * (cand_half_width / FEET_TO_M),
            wd["midpoint"].Y - search_dir.Y * (cand_half_width / FEET_TO_M),
            wd["midpoint"].Z,
        )

        gap_vec = XYZ(
            cand_inner_point.X - outer_point.X,
            cand_inner_point.Y - outer_point.Y,
            0.0,
        )
        gap_dist_ft = (
            gap_vec.X * search_dir.X + gap_vec.Y * search_dir.Y
        )

        # Gap moet positief (of net negatief bij overlappende wanden) zijn
        # en kleiner dan max gap
        if gap_dist_ft < -0.5 / FEET_TO_M:
            # Te veel overlap — waarschijnlijk zelfde wand of doorkruising
            continue

        gap_dist_m = gap_dist_ft * FEET_TO_M
        if gap_dist_m > MAX_ASSEMBLY_GAP_M:
            continue

        # Overlap check: projectie van kandidaat op huidige wandlijn
        overlap = _compute_overlap_fraction(current_data, wd)
        if overlap < MIN_OVERLAP_FRACTION:
            continue

        # Kies dichtstbijzijnde
        abs_dist = abs(gap_dist_ft)
        if abs_dist < best_distance:
            best_distance = abs_dist
            best_candidate = wd

    return best_candidate


def _bboxes_overlap_xy(data_a, data_b, margin_ft):
    """Snelle pre-filter: bounding boxes overlappen in XY + marge.

    Args:
        data_a: Wanddata dict met bbox
        data_b: Wanddata dict met bbox
        margin_ft: Extra marge in feet

    Returns:
        bool: True als bboxes overlappen
    """
    bbox_a = data_a.get("bbox")
    bbox_b = data_b.get("bbox")

    if bbox_a is None or bbox_b is None:
        # Geen bbox beschikbaar — niet uitsluiten
        return True

    if (bbox_a.Max.X + margin_ft < bbox_b.Min.X
            or bbox_b.Max.X + margin_ft < bbox_a.Min.X):
        return False

    if (bbox_a.Max.Y + margin_ft < bbox_b.Min.Y
            or bbox_b.Max.Y + margin_ft < bbox_a.Min.Y):
        return False

    # Verticale overlap ook checken (Z)
    if bbox_a.Max.Z < bbox_b.Min.Z or bbox_b.Max.Z < bbox_a.Min.Z:
        return False

    return True


def _compute_overlap_fraction(data_a, data_b):
    """Bereken de overlappende fractie van twee wanden.

    Projecteert de wandlijnen op de richting van wand A en berekent
    de overlap als fractie van de kortste wand.

    Args:
        data_a: Wanddata dict
        data_b: Wanddata dict

    Returns:
        float: Overlap fractie (0.0 - 1.0)
    """
    dir_a = data_a["direction"]

    # Projecteer start/eind van beide wanden op de as van wand A
    a_start_proj = (
        data_a["curve_start"].X * dir_a.X
        + data_a["curve_start"].Y * dir_a.Y
    )
    a_end_proj = (
        data_a["curve_end"].X * dir_a.X
        + data_a["curve_end"].Y * dir_a.Y
    )
    b_start_proj = (
        data_b["curve_start"].X * dir_a.X
        + data_b["curve_start"].Y * dir_a.Y
    )
    b_end_proj = (
        data_b["curve_end"].X * dir_a.X
        + data_b["curve_end"].Y * dir_a.Y
    )

    a_min = min(a_start_proj, a_end_proj)
    a_max = max(a_start_proj, a_end_proj)
    b_min = min(b_start_proj, b_end_proj)
    b_max = max(b_start_proj, b_end_proj)

    overlap_start = max(a_min, b_min)
    overlap_end = min(a_max, b_max)

    if overlap_end <= overlap_start:
        return 0.0

    overlap_length = overlap_end - overlap_start
    shorter_length = min(a_max - a_min, b_max - b_min)

    if shorter_length <= 0:
        return 0.0

    return overlap_length / shorter_length
