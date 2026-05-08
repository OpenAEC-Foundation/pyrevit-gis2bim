# -*- coding: utf-8 -*-
"""Dimension creator - native RevitAPI dimension helpers.

Vervangt Genius Loci's Dimension.ByElementAndReferences en archi-lab's
dimension types filtering. Plaatst horizontale en verticale dimension
lines langs een set references.
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    DimensionType,
    SpotDimensionType,
    Dimension,
    Line,
    XYZ,
    ReferenceArray,
    BoundingBoxXYZ,
)

try:
    from kozijnstaat.family_collector import (
        get_symbol_width_mm,
        get_symbol_height_mm,
    )
except Exception:
    def get_symbol_width_mm(_s):
        return 0.0

    def get_symbol_height_mm(_s):
        return 0.0

MM_TO_FT = 1.0 / 304.8


def list_project_dimension_types(doc):
    """Haal alle reguliere DimensionTypes (exclusief SpotDimensionTypes) op.

    Vervangt de 41-regel Konrad Sobon Python node in de Dynamo scripts.

    Returns:
        list[DimensionType]
    """
    dim_types = list(
        FilteredElementCollector(doc)
        .OfClass(DimensionType)
        .ToElements()
    )
    spot_types = list(
        FilteredElementCollector(doc)
        .OfClass(SpotDimensionType)
        .ToElements()
    )
    spot_ids = set(st.Id.IntegerValue for st in spot_types)

    return [
        dt for dt in dim_types
        if dt.Id.IntegerValue not in spot_ids
    ]


def find_dimension_type(doc, name):
    """Zoek een DimensionType op naam."""
    for dt in list_project_dimension_types(doc):
        try:
            if dt.Name == name:
                return dt
        except Exception:
            continue
    return None


def get_instance_bbox(instance, view):
    """Haal BoundingBox van een FamilyInstance in een view.

    Returns:
        BoundingBoxXYZ of None
    """
    try:
        return instance.get_BoundingBox(view)
    except Exception:
        return None


def create_horizontal_dimension(doc, view, references, y_position_ft,
                                dimension_type=None):
    """Maak een horizontale dimension line op hoogte y_position_ft.

    Args:
        doc: Document
        view: View
        references: list[Reference] (minimaal 2)
        y_position_ft: float (Y-coord van de dimension line in feet)
        dimension_type: DimensionType (optioneel)

    Returns:
        Dimension of None bij fout
    """
    if len(references) < 2:
        return None

    ref_array = ReferenceArray()
    for r in references:
        ref_array.Append(r)

    # Horizontale line op y_position_ft
    p1 = XYZ(-1000.0, y_position_ft, 0.0)
    p2 = XYZ(1000.0, y_position_ft, 0.0)
    dim_line = Line.CreateBound(p1, p2)

    try:
        if dimension_type is not None:
            return doc.Create.NewDimension(view, dim_line, ref_array,
                                           dimension_type)
        return doc.Create.NewDimension(view, dim_line, ref_array)
    except Exception:
        return None


def create_vertical_dimension(doc, view, references, x_position_ft,
                              dimension_type=None):
    """Maak een verticale dimension line op x_position_ft."""
    if len(references) < 2:
        return None

    ref_array = ReferenceArray()
    for r in references:
        ref_array.Append(r)

    p1 = XYZ(x_position_ft, -1000.0, 0.0)
    p2 = XYZ(x_position_ft, 1000.0, 0.0)
    dim_line = Line.CreateBound(p1, p2)

    try:
        if dimension_type is not None:
            return doc.Create.NewDimension(view, dim_line, ref_array,
                                           dimension_type)
        return doc.Create.NewDimension(view, dim_line, ref_array)
    except Exception:
        return None


def dimension_at_offset(doc, view, references, direction,
                        origin_ft, offset_mm, dimension_type=None):
    """LEGACY — gebruikt wereld-XY-plane, klopt niet voor elevations.

    Bewaard voor backward compat. Nieuwe code gebruikt
    create_dim_at_bbox_side().
    """
    offset_ft = offset_mm * MM_TO_FT
    if direction == "horizontal":
        y = origin_ft.Y + offset_ft
        return create_horizontal_dimension(doc, view, references, y,
                                           dimension_type)
    return create_vertical_dimension(doc, view, references,
                                     origin_ft.X + offset_ft,
                                     dimension_type)


def _bbox_center(bbox):
    return XYZ(
        (bbox.Min.X + bbox.Max.X) / 2.0,
        (bbox.Min.Y + bbox.Max.Y) / 2.0,
        (bbox.Min.Z + bbox.Max.Z) / 2.0,
    )


def _bbox_corners(bbox):
    corners = []
    for cx in (bbox.Min.X, bbox.Max.X):
        for cy in (bbox.Min.Y, bbox.Max.Y):
            for cz in (bbox.Min.Z, bbox.Max.Z):
                corners.append(XYZ(cx, cy, cz))
    return corners


def _max_projection(corners, center, direction):
    """Max signed projection van (corner - center) op direction."""
    best = None
    for c in corners:
        dx = c.X - center.X
        dy = c.Y - center.Y
        dz = c.Z - center.Z
        proj = (dx * direction.X
                + dy * direction.Y
                + dz * direction.Z)
        if best is None or proj > best:
            best = proj
    return best or 0.0


def create_dim_at_kozijn_side(doc, view, references, instance, side,
                              offset_mm, dimension_type=None,
                              line_half_length_ft=100.0):
    """Plaats dimension line op offset van een specifieke kozijn-instance.

    Gebruikt de werkelijke kozijn-afmetingen uit de type-parameters
    (`kozijn_breedte` / `kozijn_hoogte`) en de placement-locatie van de
    instance — NIET de view-bounding-box (die bevat ook reference-lines
    en is daardoor veel groter dan het zichtbare kozijn).

    Aannames:
      - Kozijn placement-point = bottom-center van de opening
      - Wand staat verticaal (kozijn-hoogte langs wereld-Z)
      - instance.HandOrientation wijst langs de wand
        (= horizontale axis van het kozijn in elevation)

    Args:
        side: 'top' / 'bottom' / 'left' / 'right'
        offset_mm: afstand van kozijn-rand tot dim-line in mm
    """
    if len(references) < 2 or instance is None:
        return None

    sym = instance.Symbol
    w_mm = get_symbol_width_mm(sym) or 1000.0
    h_mm = get_symbol_height_mm(sym) or 2000.0
    w_ft = w_mm * MM_TO_FT
    h_ft = h_mm * MM_TO_FT
    offset_ft = offset_mm * MM_TO_FT

    try:
        pt = instance.Location.Point
    except Exception:
        return None

    try:
        hand = instance.HandOrientation
    except Exception:
        hand = view.RightDirection

    up = XYZ(0.0, 0.0, 1.0)

    if side == "top":
        line_center = XYZ(
            pt.X, pt.Y, pt.Z + h_ft + offset_ft,
        )
        line_dir = hand
    elif side == "bottom":
        line_center = XYZ(
            pt.X, pt.Y, pt.Z - offset_ft,
        )
        line_dir = hand
    elif side == "left":
        edge = w_ft / 2.0 + offset_ft
        line_center = XYZ(
            pt.X - hand.X * edge,
            pt.Y - hand.Y * edge,
            pt.Z + h_ft / 2.0,
        )
        line_dir = up
    elif side == "right":
        edge = w_ft / 2.0 + offset_ft
        line_center = XYZ(
            pt.X + hand.X * edge,
            pt.Y + hand.Y * edge,
            pt.Z + h_ft / 2.0,
        )
        line_dir = up
    else:
        return None

    p1 = XYZ(
        line_center.X - line_dir.X * line_half_length_ft,
        line_center.Y - line_dir.Y * line_half_length_ft,
        line_center.Z - line_dir.Z * line_half_length_ft,
    )
    p2 = XYZ(
        line_center.X + line_dir.X * line_half_length_ft,
        line_center.Y + line_dir.Y * line_half_length_ft,
        line_center.Z + line_dir.Z * line_half_length_ft,
    )
    dim_line = Line.CreateBound(p1, p2)

    ref_array = ReferenceArray()
    for r in references:
        ref_array.Append(r)

    try:
        if dimension_type is not None:
            return doc.Create.NewDimension(
                view, dim_line, ref_array, dimension_type,
            )
        return doc.Create.NewDimension(view, dim_line, ref_array)
    except Exception:
        return None


def create_dim_at_bbox_side(doc, view, references, bbox, side,
                            offset_mm, dimension_type=None,
                            line_half_length_ft=1000.0):
    """Plaats dimension line buiten een bbox-zijde, met offset.

    Berekent in view-coords (UpDirection / RightDirection) waardoor
    het werkt op elevations, secties EN plans.

    Args:
        side: 'top' / 'bottom' / 'left' / 'right'
        offset_mm: afstand van bbox-rand tot dimension line in mm
        line_half_length_ft: halve lengte van de dimension line
            (Revit clipt automatisch op de references)
    """
    if len(references) < 2 or bbox is None:
        return None

    offset_ft = offset_mm * MM_TO_FT
    up = view.UpDirection
    right = view.RightDirection

    if side == "top":
        offset_dir = up
        line_dir = right
    elif side == "bottom":
        offset_dir = XYZ(-up.X, -up.Y, -up.Z)
        line_dir = right
    elif side == "right":
        offset_dir = right
        line_dir = up
    elif side == "left":
        offset_dir = XYZ(-right.X, -right.Y, -right.Z)
        line_dir = up
    else:
        return None

    center = _bbox_center(bbox)
    corners = _bbox_corners(bbox)
    edge_dist = _max_projection(corners, center, offset_dir)
    total = edge_dist + offset_ft

    line_center = XYZ(
        center.X + offset_dir.X * total,
        center.Y + offset_dir.Y * total,
        center.Z + offset_dir.Z * total,
    )

    p1 = XYZ(
        line_center.X - line_dir.X * line_half_length_ft,
        line_center.Y - line_dir.Y * line_half_length_ft,
        line_center.Z - line_dir.Z * line_half_length_ft,
    )
    p2 = XYZ(
        line_center.X + line_dir.X * line_half_length_ft,
        line_center.Y + line_dir.Y * line_half_length_ft,
        line_center.Z + line_dir.Z * line_half_length_ft,
    )
    dim_line = Line.CreateBound(p1, p2)

    ref_array = ReferenceArray()
    for r in references:
        ref_array.Append(r)

    try:
        if dimension_type is not None:
            return doc.Create.NewDimension(
                view, dim_line, ref_array, dimension_type,
            )
        return doc.Create.NewDimension(view, dim_line, ref_array)
    except Exception:
        return None
