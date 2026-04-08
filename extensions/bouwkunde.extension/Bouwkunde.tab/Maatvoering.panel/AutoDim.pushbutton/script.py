# -*- coding: utf-8 -*-
"""AutoDim - Automatische Maatvoering

Plaats automatisch maatvoering langs een getekende Detail Line.
Detecteert wanden, grids, kolommen en totaalmaten.

Auteur: 3BM Bouwkunde
Versie: 2.1.0 - Linked model ondersteuning + meerdere maatlijn types
"""

import clr
import sys
import os
import math

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from bm_logger import get_logger
log = get_logger("AutoDim")
log.info("AutoDim v2.0")

from wpf_template import WPFWindow, Huisstijl

try:
    clr.AddReference('RevitAPI')
    clr.AddReference('RevitAPIUI')
except:
    pass

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from System.Collections.Generic import List

from pyrevit import revit, forms

# GEEN doc/uidoc/active_view = revit.doc hier! 
# Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
uidoc = None
active_view = None



# =============================================================================
# SELECTION FILTER
# =============================================================================
class DetailLineFilter(ISelectionFilter):
    def AllowElement(self, element):
        if isinstance(element, DetailLine):
            return True
        if hasattr(element, 'Category') and element.Category:
            return element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Lines)
        return False
    
    def AllowReference(self, ref, pos):
        return True


# =============================================================================
# HELPERS
# =============================================================================
def get_line_from_element(element):
    if hasattr(element, 'GeometryCurve'):
        curve = element.GeometryCurve
        if isinstance(curve, Line):
            return curve
    if hasattr(element, 'Location') and hasattr(element.Location, 'Curve'):
        curve = element.Location.Curve
        if isinstance(curve, Line):
            return curve
    return None


def curves_intersect_2d(l1s, l1e, l2s, l2e):
    """Snelle 2D lijn intersectie."""
    x1, y1, x2, y2 = l1s.X, l1s.Y, l1e.X, l1e.Y
    x3, y3, x4, y4 = l2s.X, l2s.Y, l2e.X, l2e.Y
    
    denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    if abs(denom) < 1e-10:
        return False, 0, 0
    
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom
    u = -((x1 - x2) * (y1 - y3) - (y1 - y2) * (x1 - x3)) / denom
    
    if 0 <= t <= 1 and 0 <= u <= 1:
        return True, x1 + t * (x2 - x1), y1 + t * (y2 - y1)
    return False, 0, 0


def point_to_line_distance_2d(px, py, l1x, l1y, l2x, l2y):
    """Afstand van punt tot lijn segment in 2D."""
    dx = l2x - l1x
    dy = l2y - l1y
    length_sq = dx * dx + dy * dy
    
    if length_sq < 1e-10:
        return math.sqrt((px - l1x)**2 + (py - l1y)**2)
    
    t = max(0, min(1, ((px - l1x) * dx + (py - l1y) * dy) / length_sq))
    proj_x = l1x + t * dx
    proj_y = l1y + t * dy
    
    return math.sqrt((px - proj_x)**2 + (py - proj_y)**2)


def get_dimension_types():
    """Haal alle dimension types op."""
    collector = FilteredElementCollector(doc).OfClass(DimensionType)
    types = []
    for dt in collector:
        try:
            if dt.StyleType == DimensionStyleType.Linear:
                name = dt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if name:
                    types.append((dt.Id, name.AsString()))
        except:
            pass
    types.sort(key=lambda x: x[1])
    return types


def find_default_dim_type_index(dim_types):
    """Zoek index van eerste type met '1.8' in de naam."""
    for i, (dt_id, dt_name) in enumerate(dim_types):
        if "1.8" in dt_name:
            return i
    return 0 if dim_types else -1


def get_view_cut_plane_height(view):
    """Haal de cutplane hoogte van een ViewPlan op."""
    try:
        if not isinstance(view, ViewPlan):
            return None, None
        
        view_range = view.GetViewRange()
        cut_plane_offset = view_range.GetOffset(PlanViewPlane.CutPlane)
        bottom_offset = view_range.GetOffset(PlanViewPlane.BottomClipPlane)
        top_offset = view_range.GetOffset(PlanViewPlane.TopClipPlane)
        
        base_elevation = 0
        if view.GenLevel:
            base_elevation = view.GenLevel.Elevation
        
        bottom_height = base_elevation + bottom_offset
        top_height = base_elevation + top_offset
        
        return bottom_height, top_height
    except:
        return None, None


def element_visible_at_cut_plane(element, bottom_height, top_height):
    """Check of een element zichtbaar is op de cutplane hoogte."""
    if bottom_height is None or top_height is None:
        return True
    
    try:
        bbox = element.get_BoundingBox(None)
        if not bbox:
            return True
        
        elem_bottom = bbox.Min.Z
        elem_top = bbox.Max.Z
        
        return elem_bottom < top_height and elem_top > bottom_height
    except:
        return True


# =============================================================================
# REFERENCE GETTERS
# =============================================================================
def get_wall_face_references(wall, measure_line, view):
    """Haal references naar de buitenste faces van een wand."""
    references = []
    
    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    line_dir = (line_end - line_start).Normalize()
    
    wall_loc = wall.Location
    if not isinstance(wall_loc, LocationCurve):
        return []
    
    wall_curve = wall_loc.Curve
    ws = wall_curve.GetEndPoint(0)
    we = wall_curve.GetEndPoint(1)
    
    intersects, ix, iy = curves_intersect_2d(line_start, line_end, ws, we)
    if not intersects:
        return []
    
    t_intersect = (ix - line_start.X) * line_dir.X + (iy - line_start.Y) * line_dir.Y
    
    wall_dir = (we - ws).Normalize()
    wall_normal = XYZ(-wall_dir.Y, wall_dir.X, 0)
    
    options = Options()
    options.ComputeReferences = True
    options.View = view
    
    geom = wall.get_Geometry(options)
    if not geom:
        return []
    
    for geom_obj in geom:
        if not isinstance(geom_obj, Solid) or geom_obj.Volume <= 0:
            continue
        
        for face in geom_obj.Faces:
            if not isinstance(face, PlanarFace):
                continue
            
            fn = face.FaceNormal
            if abs(fn.Z) > 0.1:
                continue
            
            dot = abs(fn.X * wall_normal.X + fn.Y * wall_normal.Y)
            if dot < 0.9:
                continue
            
            ref = face.Reference
            if not ref:
                continue
            
            fo = face.Origin
            ft = (fo.X - line_start.X) * line_dir.X + (fo.Y - line_start.Y) * line_dir.Y
            
            if abs(ft - t_intersect) < 2.0:
                references.append((ref, ft))
    
    if len(references) >= 2:
        references.sort(key=lambda x: x[1])
        return [references[0], references[-1]]
    
    return references


def get_grid_references(measure_line, view):
    """Haal grid references."""
    references = []
    
    grids = FilteredElementCollector(doc, view.Id).OfClass(Grid).ToElements()
    
    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    line_dir = (line_end - line_start).Normalize()
    
    for grid in grids:
        gc = grid.Curve
        if not gc:
            continue
        
        gs, ge = gc.GetEndPoint(0), gc.GetEndPoint(1)
        intersects, ix, iy = curves_intersect_2d(line_start, line_end, gs, ge)
        
        if intersects:
            t = (ix - line_start.X) * line_dir.X + (iy - line_start.Y) * line_dir.Y
            ref = Reference(grid)
            references.append((ref, t))
    
    references.sort(key=lambda x: x[1])
    return references


def get_column_face_references(column, measure_line, view):
    """Haal references naar de buitenste faces van een kolom."""
    references = []
    
    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    line_dir = (line_end - line_start).Normalize()
    
    # Kolom locatie
    col_loc = column.Location
    if not col_loc:
        return []
    
    if hasattr(col_loc, 'Point'):
        col_point = col_loc.Point
    else:
        return []
    
    # Check of kolom dichtbij de lijn ligt
    dist = point_to_line_distance_2d(
        col_point.X, col_point.Y,
        line_start.X, line_start.Y,
        line_end.X, line_end.Y
    )
    
    # Kolom moet binnen ~1m van de lijn liggen
    if dist > 3.28:  # ~1m in feet
        return []
    
    # T-positie van kolom langs de lijn
    t_col = (col_point.X - line_start.X) * line_dir.X + (col_point.Y - line_start.Y) * line_dir.Y
    
    options = Options()
    options.ComputeReferences = True
    options.View = view
    
    geom = column.get_Geometry(options)
    if not geom:
        return []
    
    for geom_obj in geom:
        if isinstance(geom_obj, GeometryInstance):
            geom_obj = geom_obj.GetInstanceGeometry()
        
        solids = []
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            solids.append(geom_obj)
        elif hasattr(geom_obj, 'GetEnumerator'):
            for g in geom_obj:
                if isinstance(g, Solid) and g.Volume > 0:
                    solids.append(g)
        
        for solid in solids:
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                
                fn = face.FaceNormal
                if abs(fn.Z) > 0.1:
                    continue
                
                # Face moet loodrecht op de maatlijn staan
                dot = abs(fn.X * line_dir.X + fn.Y * line_dir.Y)
                if dot < 0.7:
                    continue
                
                ref = face.Reference
                if not ref:
                    continue
                
                fo = face.Origin
                ft = (fo.X - line_start.X) * line_dir.X + (fo.Y - line_start.Y) * line_dir.Y
                
                references.append((ref, ft))
    
    if len(references) >= 2:
        references.sort(key=lambda x: x[1])
        return [references[0], references[-1]]
    
    return references


# =============================================================================
# FIND CROSSING ELEMENTS
# =============================================================================
def find_crossing_walls(measure_line, walls, bottom_height, top_height, min_thickness):
    """Vind kruisende wanden binnen view range."""
    crossing = []
    
    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    
    for wall in walls:
        if not element_visible_at_cut_plane(wall, bottom_height, top_height):
            continue
        
        wl = wall.Location
        if not isinstance(wl, LocationCurve):
            continue
        
        wc = wl.Curve
        ws, we = wc.GetEndPoint(0), wc.GetEndPoint(1)
        
        intersects, ix, iy = curves_intersect_2d(line_start, line_end, ws, we)
        
        if intersects:
            wt = doc.GetElement(wall.GetTypeId())
            thickness = 0
            
            if wt:
                p = wt.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
                if p:
                    thickness = p.AsDouble() * 304.8
            
            if thickness >= min_thickness:
                crossing.append(wall)
    
    return crossing


def find_crossing_columns(measure_line, columns, bottom_height, top_height):
    """Vind kolommen nabij de maatlijn."""
    crossing = []
    
    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    
    for column in columns:
        if not element_visible_at_cut_plane(column, bottom_height, top_height):
            continue
        
        col_loc = column.Location
        if not col_loc or not hasattr(col_loc, 'Point'):
            continue
        
        col_point = col_loc.Point
        
        dist = point_to_line_distance_2d(
            col_point.X, col_point.Y,
            line_start.X, line_start.Y,
            line_end.X, line_end.Y
        )
        
        if dist < 3.28:  # ~1m
            crossing.append(column)
    
    return crossing


# =============================================================================
# LINKED MODELS
# =============================================================================
def get_linked_models(view):
    """Verzamel alle geladen RevitLinkInstances met document + transform.

    Returns:
        list of dict: Elk met 'instance', 'doc', 'transform' keys.
    """
    links = []
    link_instances = (
        FilteredElementCollector(doc, view.Id)
        .OfClass(RevitLinkInstance)
        .ToElements()
    )

    for link_inst in link_instances:
        link_doc = link_inst.GetLinkDocument()
        if not link_doc:
            continue
        transform = link_inst.GetTotalTransform()
        links.append({
            'instance': link_inst,
            'doc': link_doc,
            'transform': transform,
        })
        log.info("Linked model: {} ({} elements)".format(
            link_doc.Title,
            FilteredElementCollector(link_doc).WhereElementIsNotElementType()
                .GetElementCount()
        ))
    return links


def convert_linked_reference(face_ref, link_doc, link_instance):
    """Converteer een face reference uit een linked doc naar host-compatible.

    Args:
        face_ref: Reference naar een face in het linked document.
        link_doc: Het linked Document object.
        link_instance: De RevitLinkInstance.

    Returns:
        Reference: Host-compatible reference, of None bij falen.
    """
    try:
        stable = face_ref.ConvertToStableRepresentation(link_doc)
        host_stable = "{}:{}".format(link_instance.UniqueId, stable)
        return Reference.ParseFromStableRepresentation(doc, host_stable)
    except Exception as ex:
        log.debug("convert_linked_reference failed: {}".format(ex))
        return None


def get_linked_wall_face_references(
    wall, measure_line, view, link_doc, link_instance, transform
):
    """Variant van get_wall_face_references() voor linked wanden.

    Transformeert linked coördinaten naar host space en converteert
    face references via stable representation.
    """
    references = []

    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    line_dir = (line_end - line_start).Normalize()

    wall_loc = wall.Location
    if not isinstance(wall_loc, LocationCurve):
        return []

    wall_curve = wall_loc.Curve
    ws = transform.OfPoint(wall_curve.GetEndPoint(0))
    we = transform.OfPoint(wall_curve.GetEndPoint(1))

    intersects, ix, iy = curves_intersect_2d(line_start, line_end, ws, we)
    if not intersects:
        return []

    t_intersect = (
        (ix - line_start.X) * line_dir.X
        + (iy - line_start.Y) * line_dir.Y
    )

    wall_dir = (we - ws).Normalize()
    wall_normal = XYZ(-wall_dir.Y, wall_dir.X, 0)

    options = Options()
    options.ComputeReferences = True
    options.View = view

    geom = wall.get_Geometry(options)
    if not geom:
        return []

    for geom_obj in geom:
        if not isinstance(geom_obj, Solid) or geom_obj.Volume <= 0:
            continue

        for face in geom_obj.Faces:
            if not isinstance(face, PlanarFace):
                continue

            fn_local = face.FaceNormal
            fn = transform.OfVector(fn_local)

            if abs(fn.Z) > 0.1:
                continue

            dot = abs(fn.X * wall_normal.X + fn.Y * wall_normal.Y)
            if dot < 0.9:
                continue

            ref = face.Reference
            if not ref:
                continue

            host_ref = convert_linked_reference(ref, link_doc, link_instance)
            if not host_ref:
                continue

            fo = transform.OfPoint(face.Origin)
            ft = (
                (fo.X - line_start.X) * line_dir.X
                + (fo.Y - line_start.Y) * line_dir.Y
            )

            if abs(ft - t_intersect) < 2.0:
                references.append((host_ref, ft))

    if len(references) >= 2:
        references.sort(key=lambda x: x[1])
        return [references[0], references[-1]]

    return references


def get_linked_column_face_references(
    column, measure_line, view, link_doc, link_instance, transform
):
    """Variant van get_column_face_references() voor linked kolommen."""
    references = []

    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    line_dir = (line_end - line_start).Normalize()

    col_loc = column.Location
    if not col_loc or not hasattr(col_loc, 'Point'):
        return []

    col_point = transform.OfPoint(col_loc.Point)

    dist = point_to_line_distance_2d(
        col_point.X, col_point.Y,
        line_start.X, line_start.Y,
        line_end.X, line_end.Y
    )
    if dist > 3.28:
        return []

    options = Options()
    options.ComputeReferences = True
    options.View = view

    geom = column.get_Geometry(options)
    if not geom:
        return []

    for geom_obj in geom:
        if isinstance(geom_obj, GeometryInstance):
            geom_obj = geom_obj.GetInstanceGeometry()

        solids = []
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            solids.append(geom_obj)
        elif hasattr(geom_obj, 'GetEnumerator'):
            for g in geom_obj:
                if isinstance(g, Solid) and g.Volume > 0:
                    solids.append(g)

        for solid in solids:
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue

                fn_local = face.FaceNormal
                fn = transform.OfVector(fn_local)

                if abs(fn.Z) > 0.1:
                    continue

                dot = abs(fn.X * line_dir.X + fn.Y * line_dir.Y)
                if dot < 0.7:
                    continue

                ref = face.Reference
                if not ref:
                    continue

                host_ref = convert_linked_reference(
                    ref, link_doc, link_instance
                )
                if not host_ref:
                    continue

                fo = transform.OfPoint(face.Origin)
                ft = (
                    (fo.X - line_start.X) * line_dir.X
                    + (fo.Y - line_start.Y) * line_dir.Y
                )

                references.append((host_ref, ft))

    if len(references) >= 2:
        references.sort(key=lambda x: x[1])
        return [references[0], references[-1]]

    return references


def find_crossing_linked_walls(
    measure_line, link_info, bottom_height, top_height, min_thickness
):
    """Vind kruisende wanden in een linked model.

    Args:
        measure_line: De maatlijn.
        link_info: Dict met 'instance', 'doc', 'transform'.
        bottom_height: Onderkant view range (host coords).
        top_height: Bovenkant view range (host coords).
        min_thickness: Minimum wanddikte in mm.

    Returns:
        list of tuples: (wall, link_info) voor elke kruisende wand.
    """
    crossing = []
    link_doc = link_info['doc']
    transform = link_info['transform']

    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)

    walls = (
        FilteredElementCollector(link_doc)
        .OfClass(Wall)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for wall in walls:
        # View range check met transform offset
        if bottom_height is not None and top_height is not None:
            try:
                bbox = wall.get_BoundingBox(None)
                if bbox:
                    elem_bottom = transform.OfPoint(bbox.Min).Z
                    elem_top = transform.OfPoint(bbox.Max).Z
                    if not (elem_bottom < top_height
                            and elem_top > bottom_height):
                        continue
            except:
                pass

        wl = wall.Location
        if not isinstance(wl, LocationCurve):
            continue

        wc = wl.Curve
        ws = transform.OfPoint(wc.GetEndPoint(0))
        we = transform.OfPoint(wc.GetEndPoint(1))

        intersects, ix, iy = curves_intersect_2d(
            line_start, line_end, ws, we
        )

        if intersects:
            wt = link_doc.GetElement(wall.GetTypeId())
            thickness = 0
            if wt:
                p = wt.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
                if p:
                    thickness = p.AsDouble() * 304.8

            if thickness >= min_thickness:
                crossing.append((wall, link_info))

    return crossing


def find_crossing_linked_columns(
    measure_line, link_info, bottom_height, top_height
):
    """Vind kolommen nabij de maatlijn in een linked model.

    Returns:
        list of tuples: (column, link_info) voor elke nabije kolom.
    """
    crossing = []
    link_doc = link_info['doc']
    transform = link_info['transform']

    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)

    columns = list(
        FilteredElementCollector(link_doc)
        .OfCategory(BuiltInCategory.OST_StructuralColumns)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    columns.extend(list(
        FilteredElementCollector(link_doc)
        .OfCategory(BuiltInCategory.OST_Columns)
        .WhereElementIsNotElementType()
        .ToElements()
    ))

    for column in columns:
        if bottom_height is not None and top_height is not None:
            try:
                bbox = column.get_BoundingBox(None)
                if bbox:
                    elem_bottom = transform.OfPoint(bbox.Min).Z
                    elem_top = transform.OfPoint(bbox.Max).Z
                    if not (elem_bottom < top_height
                            and elem_top > bottom_height):
                        continue
            except:
                pass

        col_loc = column.Location
        if not col_loc or not hasattr(col_loc, 'Point'):
            continue

        col_point = transform.OfPoint(col_loc.Point)
        dist = point_to_line_distance_2d(
            col_point.X, col_point.Y,
            line_start.X, line_start.Y,
            line_end.X, line_end.Y
        )

        if dist < 3.28:
            crossing.append((column, link_info))

    return crossing


# =============================================================================
# DIMENSION CREATION
# =============================================================================
def create_dimension_from_refs(refs, measure_line, view, dim_type_id, offset=0):
    """Maak een dimension van een lijst references met optionele offset."""
    if len(refs) < 2:
        return None
    
    refs.sort(key=lambda x: x[1])
    
    ref_array = ReferenceArray()
    for ref, t in refs:
        ref_array.Append(ref)
    
    # Bereken offset loodrecht op de lijn
    line_start = measure_line.GetEndPoint(0)
    line_end = measure_line.GetEndPoint(1)
    line_dir = (line_end - line_start).Normalize()
    
    # Loodrechte richting (90 graden in XY vlak)
    perp_dir = XYZ(-line_dir.Y, line_dir.X, 0)
    
    # Offset in feet (500mm = 1.64 feet)
    offset_feet = offset * 1.64042  # offset is aantal x 500mm
    
    # Nieuwe lijn met offset
    new_start = XYZ(
        line_start.X + perp_dir.X * offset_feet,
        line_start.Y + perp_dir.Y * offset_feet,
        line_start.Z
    )
    new_end = XYZ(
        line_end.X + perp_dir.X * offset_feet,
        line_end.Y + perp_dir.Y * offset_feet,
        line_end.Z
    )
    
    dim_line = Line.CreateBound(new_start, new_end)
    
    try:
        if dim_type_id:
            dim_type = doc.GetElement(dim_type_id)
            dim = doc.Create.NewDimension(view, dim_line, ref_array, dim_type)
        else:
            dim = doc.Create.NewDimension(view, dim_line, ref_array)
        
        if dim and dim.Segments and dim.Segments.Size > 0:
            return dim
        return None
    except:
        return None


def create_dimensions(
    measure_line, options, view, walls, columns,
    bottom_height, top_height, linked_models=None
):
    """Maak alle gevraagde dimensions met offset voor meerdere lijnen.

    Args:
        linked_models: Optionele lijst van link_info dicts van
            get_linked_models(). Wordt gebruikt als include_linked=True.
    """
    created_dims = []
    dim_type_id = options.get('dim_type_id')
    include_linked = options.get('include_linked', False)
    offset_count = 0

    # Verzamel alle refs per type
    grid_refs = []
    wall_refs = []
    column_refs = []

    # GRIDS
    if options['include_grids']:
        grid_refs = get_grid_references(measure_line, view)
        log.info("Grid refs: {}".format(len(grid_refs)))

        if len(grid_refs) >= 2:
            dim = create_dimension_from_refs(
                grid_refs, measure_line, view, dim_type_id, offset_count
            )
            if dim:
                created_dims.append(('grids', dim))
                offset_count += 1

    # WANDEN — host
    if options['include_walls']:
        crossing_walls = find_crossing_walls(
            measure_line, walls, bottom_height, top_height,
            options['min_thickness']
        )
        log.info("Crossing walls (host): {}".format(len(crossing_walls)))

        for wall in crossing_walls:
            refs = get_wall_face_references(wall, measure_line, view)
            wall_refs.extend(refs)

        # WANDEN — linked
        if include_linked and linked_models:
            for link_info in linked_models:
                linked_walls = find_crossing_linked_walls(
                    measure_line, link_info,
                    bottom_height, top_height,
                    options['min_thickness']
                )
                log.info("Crossing walls (linked {}): {}".format(
                    link_info['doc'].Title, len(linked_walls)
                ))
                for wall, li in linked_walls:
                    refs = get_linked_wall_face_references(
                        wall, measure_line, view,
                        li['doc'], li['instance'], li['transform']
                    )
                    wall_refs.extend(refs)

        log.info("Wall refs total: {}".format(len(wall_refs)))

        if len(wall_refs) >= 2:
            dim = create_dimension_from_refs(
                wall_refs, measure_line, view, dim_type_id, offset_count
            )
            if dim:
                created_dims.append(('walls', dim))
                offset_count += 1

    # KOLOMMEN — host
    if options['include_columns']:
        crossing_cols = find_crossing_columns(
            measure_line, columns, bottom_height, top_height
        )
        log.info("Crossing columns (host): {}".format(len(crossing_cols)))

        for col in crossing_cols:
            refs = get_column_face_references(col, measure_line, view)
            column_refs.extend(refs)

        # KOLOMMEN — linked
        if include_linked and linked_models:
            for link_info in linked_models:
                linked_cols = find_crossing_linked_columns(
                    measure_line, link_info,
                    bottom_height, top_height
                )
                log.info("Crossing columns (linked {}): {}".format(
                    link_info['doc'].Title, len(linked_cols)
                ))
                for col, li in linked_cols:
                    refs = get_linked_column_face_references(
                        col, measure_line, view,
                        li['doc'], li['instance'], li['transform']
                    )
                    column_refs.extend(refs)

        log.info("Column refs total: {}".format(len(column_refs)))

        if len(column_refs) >= 2:
            dim = create_dimension_from_refs(
                column_refs, measure_line, view, dim_type_id, offset_count
            )
            if dim:
                created_dims.append(('columns', dim))
                offset_count += 1

    # TOTAAL (buitenste refs van alles)
    if options['include_total']:
        all_refs = []
        all_refs.extend(grid_refs)
        all_refs.extend(wall_refs)
        all_refs.extend(column_refs)

        log.info("Total refs before filter: {}".format(len(all_refs)))

        if len(all_refs) >= 2:
            all_refs.sort(key=lambda x: x[1])
            total_refs = [all_refs[0], all_refs[-1]]

            dim = create_dimension_from_refs(
                total_refs, measure_line, view, dim_type_id, offset_count
            )
            if dim:
                created_dims.append(('total', dim))
                offset_count += 1

    return created_dims


# =============================================================================
# UI
# =============================================================================
class AutoDimWindow(WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        super(AutoDimWindow, self).__init__(xaml_file, "AutoDim", width=450, height=None)
        self.options = None
        self.dim_types = get_dimension_types()
        self._populate_dim_types()
        self._bind_events()

    def _populate_dim_types(self):
        """Vul dimension type combobox"""
        for dt_id, dt_name in self.dim_types:
            self.cmb_type.Items.Add(dt_name)

        default_idx = find_default_dim_type_index(self.dim_types)
        if default_idx >= 0:
            self.cmb_type.SelectedIndex = default_idx

    def _bind_events(self):
        """Bind button events"""
        if self.btn_start:
            self.btn_start.Click += self._on_run

    def _on_run(self, sender, args):
        dim_type_id = None
        if self.cmb_type.SelectedIndex >= 0 and self.cmb_type.SelectedIndex < len(self.dim_types):
            dim_type_id = self.dim_types[self.cmb_type.SelectedIndex][0]

        try:
            min_thickness = float(self.txt_thickness.Text)
        except (ValueError, TypeError):
            min_thickness = 20.0

        self.options = {
            'include_grids': self.chk_grids.IsChecked == True,
            'include_total': self.chk_total.IsChecked == True,
            'include_walls': self.chk_walls.IsChecked == True,
            'include_columns': self.chk_columns.IsChecked == True,
            'include_linked': self.chk_linked.IsChecked == True,
            'min_thickness': min_thickness,
            'dim_type_id': dim_type_id
        }
        self.close_ok()


# =============================================================================
# MAIN
# =============================================================================
def main():
    global doc, uidoc, active_view
    
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    uidoc = revit.uidoc
    
    if not doc:
        forms.alert("Open eerst een Revit project.", title="AutoDim")
        return
    
    active_view = doc.ActiveView
    
    log.log_revit_info()
    log.section("Main")

    if not isinstance(active_view, ViewPlan):
        forms.alert("Alleen in plattegronden.", title="AutoDim")
        return
    
    # Haal view range voor cutplane filter
    bottom_height, top_height = get_view_cut_plane_height(active_view)
    log.info("View range: {:.2f} - {:.2f}".format(
        bottom_height or 0, top_height or 0
    ))
    
    # Toon UI
    window = AutoDimWindow()
    if not window.show_dialog():
        return

    options = window.options
    log.log_options(options)
    
    # Check of er iets geselecteerd is
    if not any([options['include_grids'], options['include_total'], 
                options['include_walls'], options['include_columns']]):
        forms.alert("Selecteer minimaal één optie.", title="AutoDim")
        return
    
    # Haal elementen in view — host document
    walls = list(
        FilteredElementCollector(doc, active_view.Id)
        .OfClass(Wall).ToElements()
    )
    log.info("Walls in view (host): {}".format(len(walls)))

    columns = list(
        FilteredElementCollector(doc, active_view.Id)
        .OfCategory(BuiltInCategory.OST_StructuralColumns)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    columns.extend(list(
        FilteredElementCollector(doc, active_view.Id)
        .OfCategory(BuiltInCategory.OST_Columns)
        .WhereElementIsNotElementType()
        .ToElements()
    ))
    log.info("Columns in view (host): {}".format(len(columns)))

    # Linked models (alleen laden als checkbox actief)
    linked_models = []
    if options.get('include_linked', False):
        linked_models = get_linked_models(active_view)
        log.info("Linked models loaded: {}".format(len(linked_models)))
    
    # Loop tot ESC
    dim_count = 0
    line_filter = DetailLineFilter()
    
    while True:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                line_filter,
                "Selecteer Detail Line (ESC om te stoppen)"
            )
            
            element = doc.GetElement(ref.ElementId)
            measure_line = get_line_from_element(element)
            
            if not measure_line:
                continue
            
            log.info("Line selected: {:.2f}m".format(measure_line.Length * 0.3048))
            
            # Maak dimensions
            with revit.Transaction("AutoDim"):
                created = create_dimensions(
                    measure_line, options, active_view,
                    walls, columns, bottom_height, top_height,
                    linked_models
                )
                
                for dim_type, dim in created:
                    dim_count += 1
                    log.info("Created {} dim: {}".format(dim_type, dim.Id.IntegerValue))
        
        except OperationCanceledException:
            log.info("Selection cancelled by user")
            break
        except Exception as ex:
            if "cancel" in str(ex).lower() or "aborted" in str(ex).lower():
                break
            log.exception("Error during selection")
            break
    
    # Klaar
    if dim_count > 0:
        forms.alert("{} maatlijn(en) geplaatst!".format(dim_count), title="AutoDim")
    
    log.finalize(True, "{} dimensions created".format(dim_count))


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log.exception("Error")
        log.finalize(False)
        raise
