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
    """Plaats een dimension line op een offset t.o.v. een origin-punt.

    Args:
        direction: 'horizontal' of 'vertical'
        origin_ft: XYZ referentiepunt (b.v. BoundingBox min of max)
        offset_mm: float, positief of negatief
    """
    offset_ft = offset_mm * MM_TO_FT
    if direction == "horizontal":
        y = origin_ft.Y + offset_ft
        return create_horizontal_dimension(doc, view, references, y,
                                           dimension_type)
    return create_vertical_dimension(doc, view, references,
                                     origin_ft.X + offset_ft,
                                     dimension_type)
