# -*- coding: utf-8 -*-
"""Resolve boundary elements uit host doc of linked models.

SpatialElementGeometryCalculator retourneert SpatialBoundaryElement
met LinkInstanceId + HostElementId per sub-face. Deze module abstraheert
het resolven naar het juiste element ongeacht of het in de host of
een linked model zit.

IronPython 2.7 — geen f-strings, geen type hints.
"""
from Autodesk.Revit.DB import ElementId


def resolve_boundary_element(doc, sub_face):
    """Resolve host element van een sub-face (host of linked model).

    Args:
        doc: Revit host Document
        sub_face: SpatialBoundaryFaceInfo object

    Returns:
        dict met keys:
            host_element: Revit Element (of None)
            host_doc: Revit Document (host of linked)
            host_element_id: int (of None)
            link_instance_id: int (of None)
            host_category: str (of None)
    """
    result = {
        "host_element": None,
        "host_doc": doc,
        "host_element_id": None,
        "link_instance_id": None,
        "host_category": None,
    }

    sbe = sub_face.SpatialBoundaryElement
    if sbe is None:
        return result

    link_id = sbe.LinkInstanceId
    host_id = sbe.HostElementId

    if host_id is None or host_id == ElementId.InvalidElementId:
        return result

    if link_id is not None and link_id != ElementId.InvalidElementId:
        # Element zit in een linked model
        link_instance = doc.GetElement(link_id)
        if link_instance is None:
            return result

        try:
            linked_doc = link_instance.GetLinkDocument()
        except Exception:
            linked_doc = None

        if linked_doc is None:
            return result

        host_element = linked_doc.GetElement(host_id)
        result["host_element"] = host_element
        result["host_doc"] = linked_doc
        result["host_element_id"] = host_id.IntegerValue
        result["link_instance_id"] = link_id.IntegerValue
        if host_element is not None:
            result["host_category"] = _get_category_name(host_element)
    else:
        # Element zit in het host document
        host_element = doc.GetElement(host_id)
        result["host_element"] = host_element
        result["host_element_id"] = host_id.IntegerValue
        if host_element is not None:
            result["host_category"] = _get_category_name(host_element)

    return result


def _get_category_name(element):
    """Haal de categorienaam op: Wall, Floor, Roof, Ceiling."""
    if element is None:
        return None

    cat = element.Category
    if cat is None:
        return None

    cat_id = cat.Id.IntegerValue
    if cat_id == -2000011:
        return "Wall"
    elif cat_id == -2000032:
        return "Floor"
    elif cat_id == -2000035:
        return "Roof"
    elif cat_id == -2000038:
        return "Ceiling"
    return cat.Name
