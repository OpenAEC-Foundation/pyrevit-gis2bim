# -*- coding: utf-8 -*-
"""Grid layout berekening voor kozijn-plaatsing op een surface.

Vervangt de Surface.PointAtParameter + urev/vrev logica uit
31_3BM_Kozijnstaat_create.dyn.
"""

from Autodesk.Revit.DB import XYZ

MM_TO_FT = 1.0 / 304.8


def compute_grid_points(origin, u_dir, v_dir, u_length, v_length,
                        cols, rows, item_count):
    """Bereken plaatsingspunten voor een grid op een vlak.

    Het vlak wordt beschreven door een origin + twee orthogonale
    richtingsvectoren (u = horizontaal, v = verticaal). Punten
    worden per rij van boven naar onder uitgedeeld en per rij van
    links naar rechts gevuld.

    Args:
        origin: XYZ linker-onder-hoek van het canvas (feet)
        u_dir: XYZ horizontaal normalized richtingsvector
        v_dir: XYZ verticaal normalized richtingsvector
        u_length: float, breedte canvas in feet
        v_length: float, hoogte canvas in feet
        cols: int, aantal kolommen
        rows: int, aantal rijen
        item_count: int, aantal items om te plaatsen (kan < cols*rows)

    Returns:
        list[XYZ] van lengte min(item_count, cols*rows)
    """
    if cols <= 0 or rows <= 0:
        return []

    u_step = u_length / float(cols)
    v_step = v_length / float(rows)

    points = []
    # Rijen van boven (row=0) naar onder (row=rows-1)
    for row in range(rows):
        v_param = v_length - (row + 0.5) * v_step  # midden van de rij
        for col in range(cols):
            if len(points) >= item_count:
                return points
            u_param = (col + 0.5) * u_step  # midden van de kolom
            pt = XYZ(
                origin.X + u_dir.X * u_param + v_dir.X * v_param,
                origin.Y + u_dir.Y * u_param + v_dir.Y * v_param,
                origin.Z + u_dir.Z * u_param + v_dir.Z * v_param,
            )
            points.append(pt)
    return points


def compute_tag_points(placement_points, tag_offset_ft):
    """Bepaal tag-posities als Z-offset t.o.v. plaatsingspunten.

    Args:
        placement_points: list[XYZ]
        tag_offset_ft: float offset in feet (negatief = onder)

    Returns:
        list[XYZ]
    """
    return [
        XYZ(p.X, p.Y, p.Z + tag_offset_ft)
        for p in placement_points
    ]


def estimate_canvas_size_mm(widths_mm, heights_mm, cols, rows,
                            padding_mm=500.0):
    """Schat minimum canvas-afmetingen voor een grid van kozijnen.

    Args:
        widths_mm: list[float] breedtes van te plaatsen kozijnen
        heights_mm: list[float] hoogtes
        cols, rows: grid-dimensies
        padding_mm: marge tussen cellen

    Returns:
        tuple (breedte_mm, hoogte_mm)
    """
    if not widths_mm:
        widths_mm = [1000.0]
    if not heights_mm:
        heights_mm = [2000.0]

    max_w = max(widths_mm) + padding_mm
    max_h = max(heights_mm) + padding_mm
    return (max_w * cols, max_h * rows)
