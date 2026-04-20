# -*- coding: utf-8 -*-
"""Tag placer - plaats tags op kozijnen en glas-elementen.

Vervangt Revit.Elements.Tag.ByElementAndLocation uit de Dynamo scripts.
Gebruikt native IndependentTag.Create.
"""

from Autodesk.Revit.DB import (
    IndependentTag,
    TagMode,
    TagOrientation,
    Reference,
    XYZ,
    BuiltInCategory,
    FilteredElementCollector,
    FamilySymbol,
)

MM_TO_FT = 1.0 / 304.8


def place_tag(doc, view, element, location_xyz,
              tag_mode=None, tag_orientation=None,
              leader=False):
    """Plaats een tag op een element op de gegeven locatie.

    Args:
        doc: Document
        view: View (moet tag toelaten)
        element: het te taggen element
        location_xyz: XYZ plaatsingspunt in feet
        tag_mode: TagMode enum (default TM_ADDBY_CATEGORY)
        tag_orientation: TagOrientation (default Horizontal)
        leader: bool

    Returns:
        IndependentTag of None bij fout
    """
    if tag_mode is None:
        tag_mode = TagMode.TM_ADDBY_CATEGORY
    if tag_orientation is None:
        tag_orientation = TagOrientation.Horizontal

    try:
        ref = Reference(element)
        tag = IndependentTag.Create(
            doc, view.Id, ref, leader, tag_mode, tag_orientation,
            location_xyz,
        )
        return tag
    except Exception:
        return None


def place_tag_with_family(doc, view, element, location_xyz,
                          tag_family_name, leader=False):
    """Plaats een tag en forceer daarna het gewenste tag-type.

    Revit's IndependentTag.Create kent geen 'type' argument; we plaatsen
    eerst met de default category-tag en veranderen daarna het type
    naar het gewenste tag_family_name.

    Args:
        tag_family_name: naam van de te gebruiken tag FamilySymbol

    Returns:
        IndependentTag of None
    """
    tag = place_tag(doc, view, element, location_xyz, leader=leader)
    if tag is None:
        return None

    tag_type = _find_tag_type_by_family(doc, tag_family_name,
                                        element.Category.Id)
    if tag_type is not None:
        try:
            tag.ChangeTypeId(tag_type.Id)
        except Exception:
            pass
    return tag


def _find_tag_type_by_family(doc, family_name, category_id):
    """Zoek een tag FamilySymbol op family-naam binnen een categorie."""
    symbols = (
        FilteredElementCollector(doc)
        .OfClass(FamilySymbol)
        .ToElements()
    )
    for s in symbols:
        try:
            if s.Family.Name != family_name:
                continue
            if s.Category is None:
                continue
            if s.Category.Id.IntegerValue != category_id.IntegerValue:
                continue
            return s
        except Exception:
            continue
    return None


def get_bbox_center(element, view):
    """Center-point van een element's BoundingBox in de gegeven view.

    Returns:
        XYZ of None
    """
    try:
        bbox = element.get_BoundingBox(view)
        if bbox is None:
            return None
        return XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
    except Exception:
        return None


def offset_point(point, dx_mm=0.0, dy_mm=0.0, dz_mm=0.0):
    """XYZ-offset in mm (veilige helper)."""
    return XYZ(
        point.X + dx_mm * MM_TO_FT,
        point.Y + dy_mm * MM_TO_FT,
        point.Z + dz_mm * MM_TO_FT,
    )
