# -*- coding: utf-8 -*-
"""IFC Kozijn Analyzer v3.9 - Complete workflow met plaatsing, orientatie check & maat vergelijking.

Gebaseerd op werkende logica uit Project 2964 Ponec de Winter.

Features:
- Scan IFC links voor Windows + Doors met K-type codes
- Vergelijk met geladen Revit families
- MAAT VERGELIJKING: IFC maat vs Revit family type (oranje bij mismatch)
- Plaats hulpkruisjes voor verificatie
- Plaats ontbrekende kozijnen met CORRECTE orientatie (CreateByOriginAndBasis)
- Install Depth parameter instellen bij plaatsing
- Update Install Depth voor bestaande kozijnen
- CHECK ORIENTATIE: vergelijk Revit FacingOrientation met IFC normal
- Automatisch herplaatsen van verkeerd georiënteerde kozijnen
- Plaats 3D tekst labels bij geplaatste kozijnen
- Export naar CSV

Formule voor True Center:
    X,Y = Transform.Origin + (half_width × BasisX) + (voorvlak_offset × BasisY)
    Z = Geometry Min Z (NIET BBox!)

v3.9 VERBETERING:
- IFC Maat kolom in vergelijking tabel
- Oranje highlighting bij maat mismatch (>15mm tolerantie)
- Revit family type afmetingen worden opgehaald via Width/Height parameters
"""
__title__ = "IFC Kozijn\nAnalyzer"
__author__ = "3BM Bouwkunde"
__doc__ = "Analyseert K-type kozijnen (Windows + Doors) in IFC links en plaatst ze met correcte orientatie"

# pyRevit imports
from pyrevit import revit, DB, script

# System imports
import sys
import os
import re
import math
from datetime import datetime

# Voeg lib toe aan path
lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

from bm_logger import get_logger

log = get_logger("IFCKozijnAnalyzer")

# UI imports
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Application, Panel, TextBox, AnchorStyles, BorderStyle,
    DockStyle, Padding, ScrollBars, DataGridViewCellStyle
)
from System.Drawing import (
    Point, Size, Color, Font, FontStyle, Graphics, Pen, SolidBrush,
    Rectangle, StringFormat, StringAlignment, ContentAlignment, Drawing2D
)
from System import Array

from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl

# GEEN doc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
FEET_TO_MM = 304.8

# ==============================================================================
# 3D TEKST LABELS CONFIGURATIE
# ==============================================================================

TEXT_LABEL_FAMILY = "31_GM_3D_txt_descr."
TEXT_OFFSET_NORMAL_MM = 100.0
TEXT_OFFSET_Z_MM = 100.0


# ==============================================================================
# IFC SCANNING FUNCTIES
# ==============================================================================

def get_ifc_links():
    """Haal alle RevitLinkInstances op die IFC bevatten."""
    collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.RevitLinkInstance)\
        .ToElements()
    
    ifc_links = []
    for link in collector:
        link_doc = link.GetLinkDocument()
        if link_doc:
            name = link_doc.Title or ""
            if ".ifc" in name.lower():
                link_transform = link.GetTotalTransform()
                ifc_links.append((link, link_doc, name, link_transform))
    return ifc_links


def get_openings_from_link(link_doc):
    """Haal alle Window EN Door elementen uit een gelinkt document."""
    elements = []
    
    windows = DB.FilteredElementCollector(link_doc)\
        .OfCategory(DB.BuiltInCategory.OST_Windows)\
        .WhereElementIsNotElementType()\
        .ToElements()
    for w in windows:
        elements.append((w, 'Window'))
    
    doors = DB.FilteredElementCollector(link_doc)\
        .OfCategory(DB.BuiltInCategory.OST_Doors)\
        .WhereElementIsNotElementType()\
        .ToElements()
    for d in doors:
        elements.append((d, 'Door'))
    
    return elements


def extract_k_type(ifc_name):
    """Extraheer K-type code uit IfcName string."""
    if not ifc_name:
        return None
    
    patterns = [
        r'(SK-K\d+(?:-[a-z]+)?)',
        r'(K\d+-\d+(?:sp)?)',
        r'(K\d+-\d+-\d+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, ifc_name, re.IGNORECASE)
        if match:
            return match.group(1).upper()
    return None


def get_param_value(element, param_names):
    """Haal parameter waarde op."""
    if isinstance(param_names, str):
        param_names = [param_names]
    
    for pname in param_names:
        param = element.LookupParameter(pname)
        if param and param.HasValue:
            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
    return None


def get_ifc_name(element):
    """Haal IfcName parameter op."""
    result = get_param_value(element, ["IfcName", "IFC Name", "Name"])
    if result:
        return result
    try:
        return DB.Element.Name.__get__(element)
    except:
        return None


def get_geometry_min_z(element):
    """Haal de echte minimum Z uit geometry."""
    opt = DB.Options()
    geom = element.get_Geometry(opt)
    
    min_z = None
    
    if geom:
        for geom_obj in geom:
            if isinstance(geom_obj, DB.GeometryInstance):
                inst_geom = geom_obj.GetInstanceGeometry()
                for ig in inst_geom:
                    if isinstance(ig, DB.Solid) and ig.Volume > 0:
                        for edge in ig.Edges:
                            curve = edge.AsCurve()
                            pt0 = curve.GetEndPoint(0)
                            pt1 = curve.GetEndPoint(1)
                            for pt in [pt0, pt1]:
                                if min_z is None or pt.Z < min_z:
                                    min_z = pt.Z
    
    return min_z


def get_front_face_offset(element, transform, basisY):
    """Bepaal offset van Transform.Origin naar voorvlak."""
    opt = DB.Options()
    geom = element.get_Geometry(opt)
    origin = transform.Origin
    
    for g in geom:
        if isinstance(g, DB.GeometryInstance):
            instance_geom = g.GetInstanceGeometry()
            max_dist = None
            
            for ig in instance_geom:
                if isinstance(ig, DB.Solid) and ig.Volume > 0:
                    for face in ig.Faces:
                        if isinstance(face, DB.PlanarFace):
                            normal = face.FaceNormal
                            dot = abs(normal.X * basisY.X + normal.Y * basisY.Y + normal.Z * basisY.Z)
                            if dot > 0.9:
                                bbox = face.GetBoundingBox()
                                center_uv = DB.UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2)
                                center_pt = face.Evaluate(center_uv)
                                diff = center_pt - origin
                                dist = diff.X * basisY.X + diff.Y * basisY.Y
                                if max_dist is None or dist > max_dist:
                                    max_dist = dist
            
            return max_dist if max_dist else 0.0
    return 0.0


def get_element_transform_data(element, link_transform, category='Window'):
    """Haal transform data op via GeometryInstance.Transform.
    
    Returns: dict met positie, normal, BasisX, afmetingen, rotate_180
    """
    result = {
        'x_mm': 0, 'y_mm': 0, 'z_mm': 0,
        'normal_x': 0, 'normal_y': 1, 'normal_z': 0,
        'basis_x_x': 1, 'basis_x_y': 0, 'basis_x_z': 0,
        'width_mm': 0, 'height_mm': 0,
        'rotate_180': 0
    }
    
    opt = DB.Options()
    geom = element.get_Geometry(opt)
    
    if not geom:
        return result
    
    for geom_obj in geom:
        if isinstance(geom_obj, DB.GeometryInstance):
            transform = geom_obj.Transform
            origin = transform.Origin
            basisX = transform.BasisX
            basisY = transform.BasisY
            
            # Afmetingen
            if category == 'Door':
                width_param = element.LookupParameter("Qto_DoorBaseQuantities.Width")
                height_param = element.LookupParameter("Qto_DoorBaseQuantities.Height")
            else:
                width_param = element.LookupParameter("Qto_WindowBaseQuantities.Width")
                height_param = element.LookupParameter("Qto_WindowBaseQuantities.Height")
            
            width_ft = width_param.AsDouble() if width_param and width_param.HasValue else 0
            height_ft = height_param.AsDouble() if height_param and height_param.HasValue else 0
            
            if width_ft == 0:
                w = get_param_value(element, ["Width", "Breedte", "Overall Width"])
                if w and isinstance(w, float):
                    width_ft = w
            if height_ft == 0:
                h = get_param_value(element, ["Height", "Hoogte", "Overall Height"])
                if h and isinstance(h, float):
                    height_ft = h
            
            result['width_mm'] = width_ft * FEET_TO_MM
            result['height_mm'] = height_ft * FEET_TO_MM
            
            half_width_ft = width_ft / 2.0
            front_offset = get_front_face_offset(element, transform, basisY)
            
            true_center_x = origin.X + (half_width_ft * basisX.X) + (front_offset * basisY.X)
            true_center_y = origin.Y + (half_width_ft * basisX.Y) + (front_offset * basisY.Y)
            
            min_z = get_geometry_min_z(element)
            true_center_z = min_z if min_z is not None else origin.Z
            
            local_point = DB.XYZ(true_center_x, true_center_y, true_center_z)
            transformed_point = link_transform.OfPoint(local_point)
            
            result['x_mm'] = transformed_point.X * FEET_TO_MM
            result['y_mm'] = transformed_point.Y * FEET_TO_MM
            result['z_mm'] = transformed_point.Z * FEET_TO_MM
            
            # Normal = BasisY getransformeerd
            transformed_normal = link_transform.OfVector(basisY)
            result['normal_x'] = round(transformed_normal.X, 6)
            result['normal_y'] = round(transformed_normal.Y, 6)
            result['normal_z'] = round(transformed_normal.Z, 6)
            
            # BasisX getransformeerd - CRUCIAAL voor correcte plaatsing
            transformed_basisX = link_transform.OfVector(basisX)
            result['basis_x_x'] = round(transformed_basisX.X, 6)
            result['basis_x_y'] = round(transformed_basisX.Y, 6)
            result['basis_x_z'] = round(transformed_basisX.Z, 6)
            
            result['rotate_180'] = 1 if transformed_normal.Y > 0.1 else 0
            
            break
    
    return result


def scan_ifc_openings(filter_prefix=""):
    """Scan alle IFC links en verzamel kozijn data."""
    ifc_links = get_ifc_links()
    all_elements = []
    mark_counter = 100
    
    for link, link_doc, link_name, link_transform in ifc_links:
        openings = get_openings_from_link(link_doc)
        
        for elem, category in openings:
            ifc_name = get_ifc_name(elem)
            k_type = extract_k_type(ifc_name)
            
            if not k_type:
                continue
            
            if filter_prefix and not k_type.upper().startswith(filter_prefix.upper()):
                continue
            
            data = get_element_transform_data(elem, link_transform, category)
            
            mark = get_param_value(elem, ["Mark", "IfcTag", "Tag"])
            if not mark:
                mark = mark_counter
                mark_counter += 1
            
            elem_data = {
                'ifc_id': elem.Id.IntegerValue,
                'type_name': k_type,
                'category': category,
                'mark': mark,
                'level': '',
                'x_mm': data['x_mm'],
                'y_mm': data['y_mm'],
                'z_mm': data['z_mm'],
                'normal_x': data['normal_x'],
                'normal_y': data['normal_y'],
                'normal_z': data['normal_z'],
                'basis_x_x': data['basis_x_x'],
                'basis_x_y': data['basis_x_y'],
                'basis_x_z': data['basis_x_z'],
                'width_mm': data['width_mm'],
                'height_mm': data['height_mm'],
                'rotate_180': data['rotate_180'],
                'link_name': link_name
            }
            all_elements.append(elem_data)
    
    return all_elements


# ==============================================================================
# REVIT FAMILY SCANNING
# ==============================================================================

def get_family_symbol_dimensions(symbol):
    """Haal afmetingen (Width, Height) uit een FamilySymbol.
    
    Probeert verschillende parameter namen en built-in parameters.
    Returns: (width_mm, height_mm) of (0, 0) als niet gevonden.
    """
    width_mm = 0
    height_mm = 0
    
    # Probeer verschillende width parameter namen
    width_params = ["Width", "Breedte", "Overall Width", "Rough Width", 
                    "Frame Width", "Window Width", "Door Width"]
    for pname in width_params:
        param = symbol.LookupParameter(pname)
        if param and param.HasValue and param.StorageType == DB.StorageType.Double:
            width_mm = param.AsDouble() * FEET_TO_MM
            break
    
    # Probeer built-in parameter als fallback
    if width_mm == 0:
        try:
            param = symbol.get_Parameter(DB.BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM)
            if param and param.HasValue:
                width_mm = param.AsDouble() * FEET_TO_MM
        except:
            pass
    
    # Probeer verschillende height parameter namen
    height_params = ["Height", "Hoogte", "Overall Height", "Rough Height",
                     "Frame Height", "Window Height", "Door Height"]
    for pname in height_params:
        param = symbol.LookupParameter(pname)
        if param and param.HasValue and param.StorageType == DB.StorageType.Double:
            height_mm = param.AsDouble() * FEET_TO_MM
            break
    
    # Probeer built-in parameter als fallback
    if height_mm == 0:
        try:
            param = symbol.get_Parameter(DB.BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM)
            if param and param.HasValue:
                height_mm = param.AsDouble() * FEET_TO_MM
        except:
            pass
    
    return (round(width_mm), round(height_mm))


def get_loaded_family_types():
    """Haal alle geladen Window en Door family types op met K-type codes."""
    family_types = {}
    
    for cat in [DB.BuiltInCategory.OST_Windows, DB.BuiltInCategory.OST_Doors]:
        symbols = DB.FilteredElementCollector(doc)\
            .OfClass(DB.FamilySymbol)\
            .OfCategory(cat)\
            .ToElements()
        
        category_name = 'Window' if cat == DB.BuiltInCategory.OST_Windows else 'Door'
        
        for symbol in symbols:
            try:
                family_name = symbol.FamilyName or ""
                type_name = DB.Element.Name.__get__(symbol)
                full_name = "{}: {}".format(family_name, type_name)
                
                k_code = extract_k_type(full_name) or extract_k_type(type_name) or extract_k_type(family_name)
                
                if k_code and k_code not in family_types:
                    # Haal afmetingen op
                    revit_width, revit_height = get_family_symbol_dimensions(symbol)
                    
                    family_types[k_code] = {
                        'family': family_name,
                        'type': type_name,
                        'symbol_id': symbol.Id.IntegerValue,
                        'is_active': symbol.IsActive,
                        'category': category_name,
                        'width_mm': revit_width,
                        'height_mm': revit_height
                    }
            except:
                continue
    
    return family_types


def get_placed_revit_instances():
    """Haal alle geplaatste Window en Door instances op."""
    placed = {}
    
    for cat in [DB.BuiltInCategory.OST_Windows, DB.BuiltInCategory.OST_Doors]:
        elements = DB.FilteredElementCollector(doc)\
            .OfCategory(cat)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        category_name = 'Window' if cat == DB.BuiltInCategory.OST_Windows else 'Door'
        
        for elem in elements:
            try:
                if elem.Document.Title != doc.Title:
                    continue
                
                type_id = elem.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    symbol = doc.GetElement(type_id)
                    if symbol:
                        family_name = symbol.FamilyName or ""
                        type_name = DB.Element.Name.__get__(symbol)
                        full_name = "{}: {}".format(family_name, type_name)
                        
                        k_code = extract_k_type(full_name) or extract_k_type(type_name)
                        
                        if k_code:
                            if k_code not in placed:
                                placed[k_code] = []
                            placed[k_code].append({
                                'element_id': elem.Id.IntegerValue,
                                'element': elem,
                                'category': category_name
                            })
            except:
                continue
    
    return placed


def export_placement_csv(data, output_path):
    """Exporteer data naar CSV."""
    header = "ifc_id,type_name,category,mark,level,x_mm,y_mm,z_mm,normal_x,normal_y,normal_z,width_mm,height_mm,rotate_180"
    lines = [header]
    
    for item in data:
        line = "{},{},{},{},{},{},{},{},{},{},{},{},{},{}".format(
            item.get('ifc_id', ''),
            item.get('type_name', ''),
            item.get('category', ''),
            item.get('mark', ''),
            item.get('level', ''),
            round(item.get('x_mm', 0), 1),
            round(item.get('y_mm', 0), 1),
            round(item.get('z_mm', 0), 1),
            item.get('normal_x', 0),
            item.get('normal_y', 0),
            item.get('normal_z', 0),
            round(item.get('width_mm', 0), 1),
            round(item.get('height_mm', 0), 1),
            item.get('rotate_180', 0)
        )
        lines.append(line)
    
    with open(output_path, 'w') as f:
        f.write('\n'.join(lines))
    return len(lines) - 1


# ==============================================================================
# HULPKRUISJES
# ==============================================================================

def create_helper_cross(doc, item, h_len_mm=500, v_len_mm=300):
    """Maak 3D kruis marker op kozijnpositie."""
    x_ft = item['x_mm'] / FEET_TO_MM
    y_ft = item['y_mm'] / FEET_TO_MM
    z_ft = item['z_mm'] / FEET_TO_MM
    h_half = (h_len_mm / 2.0) / FEET_TO_MM
    v_len = v_len_mm / FEET_TO_MM
    
    center = DB.XYZ(x_ft, y_ft, z_ft)
    
    nx = item.get('normal_x', 0)
    ny = item.get('normal_y', 1)
    nz = item.get('normal_z', 0)
    normal = DB.XYZ(nx, ny, nz).Normalize()
    
    up = DB.XYZ(0, 0, 1)
    
    perp = normal.CrossProduct(up)
    if perp.GetLength() < 0.001:
        perp = DB.XYZ(1, 0, 0)
    else:
        perp = perp.Normalize()
    
    l1_start = DB.XYZ(center.X - normal.X * h_half * 0.3, center.Y - normal.Y * h_half * 0.3, center.Z)
    l1_end = DB.XYZ(center.X + normal.X * h_half, center.Y + normal.Y * h_half, center.Z)
    
    l2_start = DB.XYZ(center.X - perp.X * h_half, center.Y - perp.Y * h_half, center.Z)
    l2_end = DB.XYZ(center.X + perp.X * h_half, center.Y + perp.Y * h_half, center.Z)
    
    l3_start = center
    l3_end = DB.XYZ(center.X, center.Y, center.Z + v_len)
    
    h_plane = DB.Plane.CreateByNormalAndOrigin(up, center)
    h_sketch = DB.SketchPlane.Create(doc, h_plane)
    
    v_plane = DB.Plane.CreateByNormalAndOrigin(normal, center)
    v_sketch = DB.SketchPlane.Create(doc, v_plane)
    
    doc.Create.NewModelCurve(DB.Line.CreateBound(l1_start, l1_end), h_sketch)
    doc.Create.NewModelCurve(DB.Line.CreateBound(l2_start, l2_end), h_sketch)
    doc.Create.NewModelCurve(DB.Line.CreateBound(l3_start, l3_end), v_sketch)


def place_helper_crosses(data, h_len_mm=500, v_len_mm=300):
    """Plaats hulpkruisjes voor alle items."""
    placed = 0
    failed = 0
    errors = []
    
    with revit.Transaction("Hulpkruisjes {}x".format(len(data))):
        for item in data:
            try:
                create_helper_cross(doc, item, h_len_mm, v_len_mm)
                placed += 1
            except Exception as e:
                failed += 1
                if len(errors) < 5:
                    errors.append("Mark {}: {}".format(item.get('mark', '?'), str(e)[:50]))
    
    return placed, failed, errors


# ==============================================================================
# KOZIJN PLAATSING - VERBETERD MET CreateByOriginAndBasis
# ==============================================================================

def place_single_window(doc, symbol, item, install_depth_mm=None):
    """Plaats één kozijn met CORRECTE WorkPlane orientatie.
    
    v3.8: Directe BasisX berekening uit IFC normal.
    Formule: BasisX = (-normal_y, normal_x, 0) zodat BasisX × BasisY = normal
    """
    try:
        # Coordinaten
        x_ft = item['x_mm'] / FEET_TO_MM
        y_ft = item['y_mm'] / FEET_TO_MM
        z_ft = item['z_mm'] / FEET_TO_MM
        point = DB.XYZ(x_ft, y_ft, z_ft)
        
        # Normal vector (facing direction) uit IFC
        nx = item.get('normal_x', 0)
        ny = item.get('normal_y', 1)
        nz = item.get('normal_z', 0)
        
        # Voor verticale kozijnen: BasisY = Z-up
        plane_basisY = DB.XYZ(0, 0, 1)
        
        # BasisX berekenen zodat: BasisX × BasisY = normal
        # (bx, by, 0) × (0, 0, 1) = (by, -bx, 0)
        # We willen (nx, ny, 0), dus: by = nx, -bx = ny → bx = -ny, by = nx
        plane_basisX = DB.XYZ(-ny, nx, 0)
        
        # Normaliseer (zou al genormaliseerd moeten zijn maar voor zekerheid)
        length = math.sqrt(plane_basisX.X**2 + plane_basisX.Y**2)
        if length > 0.001:
            plane_basisX = DB.XYZ(plane_basisX.X / length, plane_basisX.Y / length, 0)
        else:
            # Fallback voor verticale normals
            plane_basisX = DB.XYZ(1, 0, 0)
        
        # Maak WorkPlane
        plane = DB.Plane.CreateByOriginAndBasis(point, plane_basisX, plane_basisY)
        sketch = DB.SketchPlane.Create(doc, plane)
        
        # Plaats family instance
        instance = doc.Create.NewFamilyInstance(
            point, symbol, sketch,
            DB.Structure.StructuralType.NonStructural
        )
        
        # GEEN 180° rotatie meer - de BasisX berekening zorgt voor correcte facing
        
        # Zet Mark
        mark = item.get('mark', '')
        if mark:
            mark_param = instance.LookupParameter("Mark")
            if mark_param and not mark_param.IsReadOnly:
                mark_param.Set(str(mark))
        
        # Zet Install Depth (instance of type parameter)
        if install_depth_mm is not None:
            depth_ft = install_depth_mm / FEET_TO_MM
            depth_param = instance.LookupParameter("Install Depth")
            if depth_param and not depth_param.IsReadOnly:
                depth_param.Set(depth_ft)
            else:
                type_id = instance.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    elem_type = doc.GetElement(type_id)
                    if elem_type:
                        type_depth_param = elem_type.LookupParameter("Install Depth")
                        if type_depth_param and not type_depth_param.IsReadOnly:
                            type_depth_param.Set(depth_ft)
        
        return (True, instance.Id.IntegerValue)
    
    except Exception as e:
        return (False, str(e))


def place_windows_batch(data, install_depth_mm=None):
    """Plaats meerdere kozijnen in één transaction."""
    placed = 0
    failed = 0
    errors = []
    
    by_symbol = {}
    for item in data:
        sid = item.get('symbol_id')
        if sid:
            if sid not in by_symbol:
                by_symbol[sid] = []
            by_symbol[sid].append(item)
    
    with revit.Transaction("Plaats {} kozijnen".format(len(data))):
        for symbol_id, items in by_symbol.items():
            symbol = doc.GetElement(DB.ElementId(symbol_id))
            if not symbol:
                for item in items:
                    failed += 1
                    if len(errors) < 10:
                        errors.append("Mark {}: Symbol niet gevonden".format(item.get('mark', '?')))
                continue
            
            if not symbol.IsActive:
                symbol.Activate()
                doc.Regenerate()
            
            for item in items:
                success, result = place_single_window(doc, symbol, item, install_depth_mm)
                if success:
                    placed += 1
                else:
                    failed += 1
                    if len(errors) < 10:
                        errors.append("Mark {}: {}".format(item.get('mark', '?'), result[:50]))
    
    return placed, failed, errors


def update_install_depth_batch(element_ids, depth_mm):
    """Update Install Depth voor een lijst van element IDs."""
    updated = 0
    failed = 0
    depth_ft = depth_mm / FEET_TO_MM
    updated_type_ids = set()
    
    with revit.Transaction("Update Install Depth {}x".format(len(element_ids))):
        for elem_id in element_ids:
            try:
                element = doc.GetElement(DB.ElementId(elem_id))
                if not element:
                    failed += 1
                    continue
                
                depth_param = element.LookupParameter("Install Depth")
                if depth_param and not depth_param.IsReadOnly:
                    depth_param.Set(depth_ft)
                    updated += 1
                    continue
                
                type_id = element.GetTypeId()
                if not type_id or type_id == DB.ElementId.InvalidElementId:
                    failed += 1
                    continue
                
                if type_id.IntegerValue in updated_type_ids:
                    updated += 1
                    continue
                
                elem_type = doc.GetElement(type_id)
                if not elem_type:
                    failed += 1
                    continue
                
                type_depth_param = elem_type.LookupParameter("Install Depth")
                if type_depth_param and not type_depth_param.IsReadOnly:
                    type_depth_param.Set(depth_ft)
                    updated_type_ids.add(type_id.IntegerValue)
                    updated += 1
                else:
                    failed += 1
                    
            except:
                failed += 1
    
    return updated, failed, len(updated_type_ids)


# ==============================================================================
# ORIENTATIE VERIFICATIE & CORRECTIE (NIEUW v3.6)
# ==============================================================================

def check_orientation_mismatch(ifc_data):
    """Vergelijk Revit FacingOrientation met IFC normal voor geplaatste kozijnen.
    
    Returns dict met 'correct', 'wrong', 'not_found' lijsten.
    """
    result = {'correct': [], 'wrong': [], 'not_found': []}
    
    # Verzamel alle geplaatste kozijnen per afgeronde positie
    placed_by_pos = {}
    
    for cat in [DB.BuiltInCategory.OST_Windows, DB.BuiltInCategory.OST_Doors]:
        elements = DB.FilteredElementCollector(doc)\
            .OfCategory(cat)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        for elem in elements:
            try:
                if elem.Document.Title != doc.Title:
                    continue
                
                loc = elem.Location
                if not hasattr(loc, 'Point'):
                    continue
                
                point = loc.Point
                facing = elem.FacingOrientation
                
                # Rond positie af voor matching (binnen 100mm)
                pos_key = (
                    round(point.X * FEET_TO_MM / 100) * 100,
                    round(point.Y * FEET_TO_MM / 100) * 100,
                    round(point.Z * FEET_TO_MM / 100) * 100
                )
                
                if pos_key not in placed_by_pos:
                    placed_by_pos[pos_key] = []
                
                placed_by_pos[pos_key].append({
                    'element_id': elem.Id.IntegerValue,
                    'element': elem,
                    'facing_x': facing.X,
                    'facing_y': facing.Y,
                    'facing_z': facing.Z,
                    'point': point
                })
            except:
                continue
    
    # Match IFC items met Revit elements
    for ifc_item in ifc_data:
        pos_key = (
            round(ifc_item['x_mm'] / 100) * 100,
            round(ifc_item['y_mm'] / 100) * 100,
            round(ifc_item['z_mm'] / 100) * 100
        )
        
        if pos_key not in placed_by_pos:
            result['not_found'].append(ifc_item)
            continue
        
        best_match = None
        best_dot = -2
        
        for placed in placed_by_pos[pos_key]:
            dot = (placed['facing_x'] * ifc_item['normal_x'] + 
                   placed['facing_y'] * ifc_item['normal_y'] +
                   placed['facing_z'] * ifc_item['normal_z'])
            
            if dot > best_dot:
                best_dot = dot
                best_match = placed
        
        if best_match:
            match_data = {
                'ifc_item': ifc_item,
                'revit_element': best_match,
                'dot_product': best_dot,
                'angle_degrees': math.degrees(math.acos(min(1, max(-1, best_dot))))
            }
            
            # dot > 0.85 = hoek < ~30° = correct
            if best_dot > 0.85:
                result['correct'].append(match_data)
            else:
                result['wrong'].append(match_data)
        else:
            result['not_found'].append(ifc_item)
    
    return result


def reorient_window(element, ifc_item):
    """Herplaats een kozijn met correcte orientatie.
    
    Gebruikt dezelfde BasisX berekening als place_single_window.
    """
    try:
        symbol_id = element.GetTypeId()
        symbol = doc.GetElement(symbol_id)
        
        mark_param = element.LookupParameter("Mark")
        mark = mark_param.AsString() if mark_param and mark_param.HasValue else ""
        
        install_depth_ft = None
        depth_param = element.LookupParameter("Install Depth")
        if depth_param and depth_param.HasValue:
            install_depth_ft = depth_param.AsDouble()
        else:
            elem_type = doc.GetElement(symbol_id)
            if elem_type:
                type_depth_param = elem_type.LookupParameter("Install Depth")
                if type_depth_param and type_depth_param.HasValue:
                    install_depth_ft = type_depth_param.AsDouble()
        
        # Verwijder origineel
        doc.Delete(element.Id)
        
        # Plaats opnieuw met IFC orientatie
        item_data = dict(ifc_item)
        item_data['mark'] = mark
        
        install_depth_mm = install_depth_ft * FEET_TO_MM if install_depth_ft else None
        success, result = place_single_window(doc, symbol, item_data, install_depth_mm)
        
        return (success, result)
    
    except Exception as e:
        return (False, str(e))


def reorient_windows_batch(wrong_items):
    """Herplaats meerdere verkeerd georiënteerde kozijnen."""
    corrected = 0
    failed = 0
    errors = []
    
    with revit.Transaction("Heroriënteer {} kozijnen".format(len(wrong_items))):
        for item in wrong_items:
            ifc_data = item['ifc_item']
            revit_data = item['revit_element']
            element = revit_data['element']
            
            success, result = reorient_window(element, ifc_data)
            
            if success:
                corrected += 1
            else:
                failed += 1
                if len(errors) < 10:
                    mark = ifc_data.get('mark', '?')
                    errors.append("Mark {}: {}".format(mark, str(result)[:50]))
    
    return corrected, failed, errors


# ==============================================================================
# 3D TEKST LABELS
# ==============================================================================

def get_text_label_symbol():
    """Zoek de 3D tekst family symbol."""
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        if s.FamilyName == TEXT_LABEL_FAMILY:
            return s
    return None


def get_default_level():
    """Haal een default level."""
    for lvl in DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements():
        if "00" in lvl.Name or "begane" in lvl.Name.lower():
            return lvl
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements())
    return levels[0] if levels else None


def place_text_label_at_window(doc, window, symbol, level):
    """Plaats een 3D tekst label bij een kozijn."""
    offset_normal = TEXT_OFFSET_NORMAL_MM / FEET_TO_MM
    offset_z = TEXT_OFFSET_Z_MM / FEET_TO_MM
    
    loc = window.Location
    if not hasattr(loc, 'Point'):
        return None
    
    point = loc.Point
    facing = window.FacingOrientation
    
    type_id = window.GetTypeId()
    if type_id and type_id != DB.ElementId.InvalidElementId:
        symbol_elem = doc.GetElement(type_id)
        if symbol_elem:
            type_name = DB.Element.Name.__get__(symbol_elem)
            k_type = extract_k_type(type_name) or type_name
        else:
            k_type = "?"
    else:
        k_type = "?"
    
    text_x = point.X + facing.X * offset_normal
    text_y = point.Y + facing.Y * offset_normal
    text_z = point.Z + offset_z
    
    text_point = DB.XYZ(text_x, text_y, text_z)
    
    instance = doc.Create.NewFamilyInstance(
        text_point, symbol, level,
        DB.Structure.StructuralType.NonStructural
    )
    
    merk_param = instance.LookupParameter("merk")
    if merk_param and not merk_param.IsReadOnly:
        merk_param.Set(str(k_type))
    
    angle = math.atan2(facing.Y, facing.X) + math.pi / 2
    axis = DB.Line.CreateBound(text_point, DB.XYZ(text_point.X, text_point.Y, text_point.Z + 1))
    DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)
    
    return instance


def delete_existing_labels():
    """Verwijder alle bestaande tekst labels."""
    instances = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).ToElements()
    
    deleted = 0
    ids_to_delete = []
    
    for inst in instances:
        try:
            if inst.Symbol.FamilyName == TEXT_LABEL_FAMILY:
                ids_to_delete.append(inst.Id)
        except:
            continue
    
    if ids_to_delete:
        with revit.Transaction("Verwijder {} labels".format(len(ids_to_delete))):
            for elem_id in ids_to_delete:
                try:
                    doc.Delete(elem_id)
                    deleted += 1
                except:
                    pass
    
    return deleted


def place_labels_for_placed_kozijns(k_types_filter=None):
    """Plaats 3D tekst labels bij alle geplaatste kozijnen."""
    symbol = get_text_label_symbol()
    if not symbol:
        return (0, 0, ["Family '{}' niet gevonden!".format(TEXT_LABEL_FAMILY)])
    
    level = get_default_level()
    if not level:
        return (0, 0, ["Geen level gevonden!"])
    
    placed = 0
    failed = 0
    errors = []
    
    with revit.Transaction("Plaats 3D labels"):
        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()
        
        for cat in [DB.BuiltInCategory.OST_Windows, DB.BuiltInCategory.OST_Doors]:
            elements = DB.FilteredElementCollector(doc)\
                .OfCategory(cat)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            for elem in elements:
                try:
                    if elem.Document.Title != doc.Title:
                        continue
                    
                    if k_types_filter:
                        type_id = elem.GetTypeId()
                        if type_id and type_id != DB.ElementId.InvalidElementId:
                            sym = doc.GetElement(type_id)
                            if sym:
                                type_name = DB.Element.Name.__get__(sym)
                                k_type = extract_k_type(type_name)
                                if k_type and k_type not in k_types_filter:
                                    continue
                    
                    instance = place_text_label_at_window(doc, elem, symbol, level)
                    if instance:
                        placed += 1
                    else:
                        failed += 1
                except Exception as e:
                    failed += 1
                    if len(errors) < 5:
                        errors.append(str(e)[:50])
    
    return placed, failed, errors


# ==============================================================================
# PREVIEW PANEL
# ==============================================================================

class PreviewPanel(Panel):
    """Panel dat kozijnen tekent als plattegrond preview."""
    
    def __init__(self):
        self.data = []
        self.selected_index = -1
        self.BackColor = Color.White
        self.BorderStyle = BorderStyle.FixedSingle
        self.DoubleBuffered = True
        
        self.type_colors = {
            'K': Color.FromArgb(69, 182, 168),
            'SK': Color.FromArgb(239, 189, 117),
        }
        self.default_color = Color.FromArgb(160, 160, 180)
    
    def set_data(self, data):
        self.data = data or []
        self.Invalidate()
    
    def set_selected(self, index):
        self.selected_index = index
        self.Invalidate()
    
    def get_type_color(self, type_name):
        if not type_name:
            return self.default_color
        for prefix, color in self.type_colors.items():
            if type_name.startswith(prefix):
                return color
        return self.default_color
    
    def OnPaint(self, e):
        g = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.Clear(Color.White)
        
        if not self.data:
            font = Font("Segoe UI", 10)
            brush = SolidBrush(Huisstijl.TEXT_SECONDARY)
            text = "Geen kozijnen geladen"
            sf = StringFormat()
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            rect = Rectangle(0, 0, int(self.Width), int(self.Height))
            g.DrawString(text, font, brush, rect, sf)
            return
        
        min_x = min(w['x_mm'] for w in self.data)
        max_x = max(w['x_mm'] for w in self.data)
        min_y = min(w['y_mm'] for w in self.data)
        max_y = max(w['y_mm'] for w in self.data)
        
        padding = 40
        width = max_x - min_x
        height = max_y - min_y
        
        if width == 0:
            width = 1000
        if height == 0:
            height = 1000
        
        scale_x = (self.Width - 2 * padding) / width if width > 0 else 1
        scale_y = (self.Height - 2 * padding) / height if height > 0 else 1
        scale = min(scale_x, scale_y)
        
        offset_x = padding + (self.Width - 2 * padding - width * scale) / 2
        offset_y = padding + (self.Height - 2 * padding - height * scale) / 2
        
        grid_pen = Pen(Color.FromArgb(240, 240, 240), 1)
        for i in range(0, int(self.Width), 50):
            g.DrawLine(grid_pen, i, 0, i, int(self.Height))
        for i in range(0, int(self.Height), 50):
            g.DrawLine(grid_pen, 0, i, int(self.Width), i)
        
        for i, item in enumerate(self.data):
            x = offset_x + (item['x_mm'] - min_x) * scale
            y = offset_y + (max_y - item['y_mm']) * scale
            
            color = self.get_type_color(item['type_name'])
            
            w = max(6, item['width_mm'] * scale * 0.1)
            h = max(6, item['height_mm'] * scale * 0.1)
            
            if i == self.selected_index:
                highlight_pen = Pen(Huisstijl.MAGENTA, 3)
                g.DrawRectangle(highlight_pen, int(x - w/2 - 2), int(y - h/2 - 2), int(w + 4), int(h + 4))
            
            brush = SolidBrush(color)
            pen = Pen(Color.FromArgb(50, 50, 50), 1)
            
            if item.get('category') == 'Door':
                points = Array[Point]([
                    Point(int(x), int(y - h/2)),
                    Point(int(x + w/2), int(y)),
                    Point(int(x), int(y + h/2)),
                    Point(int(x - w/2), int(y)),
                ])
                g.FillPolygon(brush, points)
                g.DrawPolygon(pen, points)
            else:
                rect = Rectangle(int(x - w/2), int(y - h/2), int(w), int(h))
                g.FillRectangle(brush, rect)
                g.DrawRectangle(pen, rect)
            
            nx = item['normal_x']
            ny = item['normal_y']
            if abs(nx) > 0.01 or abs(ny) > 0.01:
                arrow_pen = Pen(Huisstijl.PEACH, 2)
                g.DrawLine(arrow_pen, int(x), int(y), int(x + nx * 10), int(y - ny * 10))
        
        self._draw_legend(g)
    
    def _draw_legend(self, g):
        x = 10
        y = int(self.Height) - 80
        
        font = Font("Segoe UI", 8)
        brush = SolidBrush(Huisstijl.TEXT_PRIMARY)
        
        rect_k = Rectangle(x, y, 12, 12)
        g.FillRectangle(SolidBrush(self.type_colors['K']), rect_k)
        g.DrawString("K-type (window)", font, brush, x + 18, y - 2)
        
        rect_sk = Rectangle(x, y + 18, 12, 12)
        g.FillRectangle(SolidBrush(self.type_colors['SK']), rect_sk)
        g.DrawString("SK-type (window)", font, brush, x + 18, y + 16)
        
        points = Array[Point]([
            Point(x + 6, y + 36),
            Point(x + 12, y + 42),
            Point(x + 6, y + 48),
            Point(x, y + 42),
        ])
        g.FillPolygon(SolidBrush(self.type_colors['K']), points)
        g.DrawString("= Door", font, brush, x + 18, y + 36)


# ==============================================================================
# VERGELIJKING DIALOG
# ==============================================================================

class CompareDialog(BaseForm):
    """Dialog voor vergelijking IFC vs Revit."""
    
    def __init__(self, ifc_data, filter_prefix):
        super(CompareDialog, self).__init__(
            "Vergelijking IFC vs Revit",
            width=960, height=680
        )
        self.ifc_data = ifc_data
        self.filter_prefix = filter_prefix
        self.comparison_data = []
        self.placeable_data = []
        self._setup_ui()
        self._do_comparison()
    
    def _setup_ui(self):
        margin = int(DPIScaler.scale(15))
        y = margin
        
        self.lbl_info = UIFactory.create_label("Vergelijking wordt uitgevoerd...", bold=True)
        self.lbl_info.Location = Point(margin, y)
        self.pnl_content.Controls.Add(self.lbl_info)
        y += int(DPIScaler.scale(35))
        
        columns = [
            ("type", "K-Type", 90),
            ("cat", "Cat", 50),
            ("ifc_maat", "IFC Maat", 85),
            ("ifc", "In IFC", 55),
            ("revit_type", "Family", 85),
            ("revit_placed", "Geplaatst", 65),
            ("to_place", "Te plaatsen", 75),
            ("status", "Status", 100),
            ("symbol_id", "ID", 60),
        ]
        
        self.grid = UIFactory.create_datagridview(columns, 910, 380)
        self.grid.Location = Point(margin, y)
        self.pnl_content.Controls.Add(self.grid)
        y += int(DPIScaler.scale(400))
        
        # Install Depth invoer
        lbl_depth = UIFactory.create_label("Install Depth (mm):", bold=True)
        lbl_depth.Location = Point(margin, y)
        self.pnl_content.Controls.Add(lbl_depth)
        
        self.txt_install_depth = UIFactory.create_textbox(width=80)
        self.txt_install_depth.Text = "0"
        self.txt_install_depth.Location = Point(margin + int(DPIScaler.scale(130)), y - 3)
        self.pnl_content.Controls.Add(self.txt_install_depth)
        
        self.lbl_depth_hint = UIFactory.create_label("← Voor nieuwe kozijnen", color=Huisstijl.TEXT_SECONDARY)
        self.lbl_depth_hint.Location = Point(margin + int(DPIScaler.scale(220)), y)
        self.pnl_content.Controls.Add(self.lbl_depth_hint)
        
        # Footer buttons
        self.add_footer_button("Plaats Ontbrekende", 'primary', self._on_place_missing, width=150)
        self.add_footer_button("Check Orientatie", 'warning', self._on_check_orientation, width=130)
        self.add_footer_button("Plaats Labels", 'secondary', self._on_place_labels, width=110)
        self.add_footer_button("Export CSV", 'secondary', self._on_export_placeable, width=100)
        self.add_footer_button("Update Depth", 'secondary', self._on_update_existing_depth, width=100)
    
    def _do_comparison(self):
        family_types = get_loaded_family_types()
        placed_instances = get_placed_revit_instances()
        
        ifc_counts = {}
        for item in self.ifc_data:
            k_type = item['type_name']
            if k_type not in ifc_counts:
                ifc_counts[k_type] = {'count': 0, 'category': item['category'], 'items': [], 
                                      'width_mm': 0, 'height_mm': 0}
            ifc_counts[k_type]['count'] += 1
            ifc_counts[k_type]['items'].append(item)
            # Neem de eerste niet-nul maat als representatief
            if ifc_counts[k_type]['width_mm'] == 0 and item.get('width_mm', 0) > 0:
                ifc_counts[k_type]['width_mm'] = round(item['width_mm'])
            if ifc_counts[k_type]['height_mm'] == 0 and item.get('height_mm', 0) > 0:
                ifc_counts[k_type]['height_mm'] = round(item['height_mm'])
        
        all_types = set(ifc_counts.keys()) | set(family_types.keys())
        
        self.comparison_data = []
        self.placeable_data = []
        
        stats = {'total_ifc': 0, 'has_family': 0, 'already_placed': 0, 'to_place': 0, 'missing_family': 0}
        
        # Tolerantie voor maat vergelijking (mm)
        SIZE_TOLERANCE = 15
        
        for k_type in sorted(all_types):
            ifc_count = ifc_counts.get(k_type, {}).get('count', 0)
            ifc_category = ifc_counts.get(k_type, {}).get('category', '-')
            ifc_items = ifc_counts.get(k_type, {}).get('items', [])
            ifc_width = ifc_counts.get(k_type, {}).get('width_mm', 0)
            ifc_height = ifc_counts.get(k_type, {}).get('height_mm', 0)
            
            has_family = k_type in family_types
            family_info = family_types.get(k_type, {})
            revit_width = family_info.get('width_mm', 0)
            revit_height = family_info.get('height_mm', 0)
            
            # Check size mismatch
            size_mismatch = False
            if has_family and ifc_count > 0 and ifc_width > 0 and ifc_height > 0:
                if revit_width > 0 and revit_height > 0:
                    width_diff = abs(ifc_width - revit_width)
                    height_diff = abs(ifc_height - revit_height)
                    if width_diff > SIZE_TOLERANCE or height_diff > SIZE_TOLERANCE:
                        size_mismatch = True
            
            # Format IFC maat string
            if ifc_width > 0 and ifc_height > 0:
                ifc_maat_str = "{}x{}".format(int(ifc_width), int(ifc_height))
            else:
                ifc_maat_str = "-"
            
            placed_count = len(placed_instances.get(k_type, []))
            to_place = max(0, ifc_count - placed_count) if has_family else 0
            
            if ifc_count == 0:
                status = "Alleen Revit"
            elif not has_family:
                status = "Family mist!"
                stats['missing_family'] += ifc_count
            elif placed_count >= ifc_count:
                status = "Compleet"
                stats['already_placed'] += ifc_count
            elif placed_count > 0:
                status = "Deels"
                stats['already_placed'] += placed_count
                stats['to_place'] += to_place
            else:
                status = "Te plaatsen"
                stats['to_place'] += to_place
            
            stats['total_ifc'] += ifc_count
            if has_family and ifc_count > 0:
                stats['has_family'] += ifc_count
            
            row_data = {
                'type': k_type,
                'category': ifc_category,
                'ifc_count': ifc_count,
                'ifc_maat': ifc_maat_str,
                'ifc_width': ifc_width,
                'ifc_height': ifc_height,
                'revit_width': revit_width,
                'revit_height': revit_height,
                'size_mismatch': size_mismatch,
                'has_family': has_family,
                'family_name': family_info.get('family', '-')[:15] if has_family else '-',
                'placed_count': placed_count,
                'to_place': to_place,
                'status': status,
                'symbol_id': family_info.get('symbol_id', '') if has_family else '',
                'items': ifc_items
            }
            
            self.comparison_data.append(row_data)
            
            if to_place > 0 and has_family:
                for item in ifc_items[:to_place]:
                    item_copy = dict(item)
                    item_copy['symbol_id'] = family_info.get('symbol_id', '')
                    self.placeable_data.append(item_copy)
        
        self.grid.Rows.Clear()
        for row in self.comparison_data:
            row_idx = self.grid.Rows.Add(
                row['type'],
                row['category'][0] if row['category'] != '-' else '-',
                row['ifc_maat'],
                str(row['ifc_count']) if row['ifc_count'] > 0 else '-',
                "Ja" if row['has_family'] else "Nee",
                str(row['placed_count']) if row['placed_count'] > 0 else '-',
                str(row['to_place']) if row['to_place'] > 0 else '-',
                row['status'],
                str(row['symbol_id']) if row['symbol_id'] else '-'
            )
            
            # Maat kolom oranje kleuren bij mismatch (kolom index 2)
            if row['size_mismatch']:
                self.grid.Rows[row_idx].Cells[2].Style.BackColor = Huisstijl.YELLOW
                self.grid.Rows[row_idx].Cells[2].Style.ForeColor = Huisstijl.VIOLET
                self.grid.Rows[row_idx].Cells[2].Style.Font = Font("Segoe UI", 9, FontStyle.Bold)
            
            # Status kolom kleuren
            if row['status'] == "Family mist!":
                self.grid.Rows[row_idx].Cells[7].Style.ForeColor = Huisstijl.PEACH
                self.grid.Rows[row_idx].Cells[7].Style.Font = Font("Segoe UI", 9, FontStyle.Bold)
            elif row['status'] == "Compleet":
                self.grid.Rows[row_idx].Cells[7].Style.ForeColor = Huisstijl.TEAL
            elif row['status'] == "Te plaatsen":
                self.grid.Rows[row_idx].Cells[7].Style.ForeColor = Huisstijl.VIOLET
                self.grid.Rows[row_idx].Cells[7].Style.Font = Font("Segoe UI", 9, FontStyle.Bold)
        
        # Tel size mismatches
        size_mismatch_count = sum(1 for r in self.comparison_data if r['size_mismatch'])
        
        info_text = "IFC: {} | Family: {} | Geplaatst: {} | Te plaatsen: {} | Mist: {}".format(
            stats['total_ifc'], stats['has_family'], stats['already_placed'], stats['to_place'], stats['missing_family'])
        if size_mismatch_count > 0:
            info_text += " | Maat afwijking: {}".format(size_mismatch_count)
        self.lbl_info.Text = info_text
        
        self.set_subtitle("{} te plaatsen".format(len(self.placeable_data)))
    
    def _get_install_depth_mm(self):
        try:
            value = float(self.txt_install_depth.Text.strip())
            if value < 0 or value > 500:
                self.show_warning("Install Depth moet tussen 0 en 500 mm zijn.")
                return None
            return value
        except:
            self.show_warning("Ongeldige Install Depth waarde.")
            return None
    
    def _on_update_existing_depth(self, sender, args):
        depth_mm = self._get_install_depth_mm()
        if depth_mm is None:
            return
        
        placed_instances = get_placed_revit_instances()
        types_in_comparison = set(row['type'] for row in self.comparison_data if row['ifc_count'] > 0)
        
        element_ids = []
        for k_type, instances in placed_instances.items():
            if k_type in types_in_comparison:
                for inst in instances:
                    element_ids.append(inst['element_id'])
        
        if not element_ids:
            self.show_warning("Geen geplaatste kozijnen gevonden.")
            return
        
        if not self.ask_confirm("Install Depth updaten naar {} mm voor {} kozijnen?".format(depth_mm, len(element_ids))):
            return
        
        updated, failed, types_count = update_install_depth_batch(element_ids, depth_mm)
        
        msg = "Install Depth geüpdatet: {} kozijnen ({} types)".format(updated, types_count)
        if failed > 0:
            msg += "\nMislukt: {}".format(failed)
        
        self.show_info(msg) if failed == 0 else self.show_warning(msg)
    
    def _on_check_orientation(self, sender, args):
        """Check orientatie van geplaatste kozijnen tegen IFC data."""
        if not self.ifc_data:
            self.show_warning("Geen IFC data beschikbaar.")
            return
        
        result = check_orientation_mismatch(self.ifc_data)
        
        correct_count = len(result['correct'])
        wrong_count = len(result['wrong'])
        not_found_count = len(result['not_found'])
        
        if wrong_count == 0:
            self.show_info(
                "Orientatie Check:\n\n"
                "✓ Correct: {}\n"
                "○ Niet gevonden: {}\n\n"
                "Alle geplaatste kozijnen OK!".format(correct_count, not_found_count))
            return
        
        wrong_details = []
        for item in result['wrong'][:10]:
            mark = item['ifc_item'].get('mark', '?')
            angle = item['angle_degrees']
            wrong_details.append("  Mark {}: {:.0f}° af".format(mark, angle))
        
        if len(result['wrong']) > 10:
            wrong_details.append("  ... en {} meer".format(len(result['wrong']) - 10))
        
        if self.ask_confirm(
            "Orientatie Check:\n\n"
            "✓ Correct: {}\n"
            "✗ Verkeerd: {}\n"
            "○ Niet gevonden: {}\n\n"
            "Verkeerde:\n{}\n\n"
            "Herplaatsen met correcte orientatie?".format(
                correct_count, wrong_count, not_found_count, "\n".join(wrong_details))):
            
            corrected, failed, errors = reorient_windows_batch(result['wrong'])
            
            msg = "Gecorrigeerd: {}".format(corrected)
            if failed > 0:
                msg += "\nMislukt: {}".format(failed)
                if errors:
                    msg += "\n\n" + "\n".join(errors[:5])
            
            self.show_info(msg) if failed == 0 else self.show_warning(msg)
            self._do_comparison()
    
    def _on_place_missing(self, sender, args):
        if not self.placeable_data:
            self.show_warning("Geen ontbrekende kozijnen.")
            return
        
        depth_mm = self._get_install_depth_mm()
        if depth_mm is None:
            return
        
        type_counts = {}
        for item in self.placeable_data:
            t = item['type_name']
            type_counts[t] = type_counts.get(t, 0) + 1
        
        type_summary = "\n".join(["  {} x {}".format(c, t) for t, c in sorted(type_counts.items())])
        
        if not self.ask_confirm(
            "{} kozijnen plaatsen?\n\n{}\n\nInstall Depth: {} mm".format(
                len(self.placeable_data), type_summary, depth_mm)):
            return
        
        placed, failed, errors = place_windows_batch(self.placeable_data, install_depth_mm=depth_mm)
        
        msg = "Geplaatst: {} (Depth: {} mm)".format(placed, depth_mm)
        if failed > 0:
            msg += "\nMislukt: {}".format(failed)
            if errors:
                msg += "\n\n" + "\n".join(errors[:5])
        
        self.show_info(msg) if failed == 0 else self.show_warning(msg)
        self._do_comparison()
    
    def _on_place_labels(self, sender, args):
        symbol = get_text_label_symbol()
        if not symbol:
            self.show_error("Family '{}' niet gevonden!".format(TEXT_LABEL_FAMILY))
            return
        
        placed_instances = get_placed_revit_instances()
        types_in_comparison = set(row['type'] for row in self.comparison_data if row['ifc_count'] > 0)
        
        total_placed = sum(len(instances) for k_type, instances in placed_instances.items() if k_type in types_in_comparison)
        
        if total_placed == 0:
            self.show_warning("Geen geplaatste kozijnen gevonden.")
            return
        
        if not self.ask_confirm("Labels plaatsen bij {} kozijnen?".format(total_placed)):
            return
        
        placed, failed, errors = place_labels_for_placed_kozijns(types_in_comparison)
        
        msg = "Labels geplaatst: {}".format(placed)
        if failed > 0:
            msg += "\nMislukt: {}".format(failed)
        
        self.show_info(msg) if failed == 0 else self.show_warning(msg)
    
    def _on_export_placeable(self, sender, args):
        if not self.placeable_data:
            self.show_warning("Geen data om te exporteren.")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = "kozijnen_te_plaatsen_{}.csv".format(timestamp)
        
        save_path = self.save_file_dialog(filter="CSV (*.csv)|*.csv", filename=default_name)
        
        if save_path:
            try:
                count = export_placement_csv(self.placeable_data, save_path)
                self.show_info("Geëxporteerd: {} elementen".format(count))
                os.startfile(os.path.dirname(save_path))
            except Exception as ex:
                self.show_error("Export fout: {}".format(str(ex)))


# ==============================================================================
# HOOFD FORM
# ==============================================================================

class IFCKozijnAnalyzerForm(BaseForm):
    """Hoofdformulier voor IFC Kozijn Analyzer."""
    
    def __init__(self):
        super(IFCKozijnAnalyzerForm, self).__init__(
            "IFC Kozijn Analyzer v3.9", 
            width=1150, height=700
        )
        self.all_data = []
        self.filtered_data = []
        self._setup_ui()
        self._load_data()
    
    def _setup_ui(self):
        margin = int(DPIScaler.scale(15))
        left_width = int(DPIScaler.scale(600))
        y = margin
        
        lbl_filter = UIFactory.create_label("Filter prefix:", bold=True)
        lbl_filter.Location = Point(margin, y)
        self.pnl_content.Controls.Add(lbl_filter)
        
        self.txt_filter = UIFactory.create_textbox(width=100)
        self.txt_filter.Text = "K"
        self.txt_filter.Location = Point(margin + int(DPIScaler.scale(100)), y - 3)
        self.txt_filter.TextChanged += self._on_filter_changed
        self.pnl_content.Controls.Add(self.txt_filter)
        
        self.lbl_count = UIFactory.create_label("0 elementen", color=Huisstijl.TEXT_SECONDARY)
        self.lbl_count.Location = Point(margin + int(DPIScaler.scale(220)), y)
        self.pnl_content.Controls.Add(self.lbl_count)
        
        y += int(DPIScaler.scale(40))
        
        columns = [
            ("type", "Type", 80),
            ("cat", "Cat", 55),
            ("mark", "Mark", 55),
            ("x", "X (mm)", 75),
            ("y", "Y (mm)", 75),
            ("z", "Z (mm)", 65),
            ("width", "B (mm)", 65),
            ("height", "H (mm)", 65),
        ]
        
        self.grid = UIFactory.create_datagridview(columns, 580, 420)
        self.grid.Location = Point(margin, y)
        self.grid.SelectionChanged += self._on_grid_selection_changed
        self.pnl_content.Controls.Add(self.grid)
        
        preview_x = left_width + margin
        preview_width = int(DPIScaler.scale(480))
        preview_height = int(DPIScaler.scale(460))
        
        lbl_preview = UIFactory.create_label("Plattegrond Preview:", bold=True)
        lbl_preview.Location = Point(preview_x, margin)
        self.pnl_content.Controls.Add(lbl_preview)
        
        self.preview = PreviewPanel()
        self.preview.Location = Point(preview_x, margin + int(DPIScaler.scale(30)))
        self.preview.Size = Size(preview_width, preview_height)
        self.pnl_content.Controls.Add(self.preview)
        
        # Footer buttons
        self.add_footer_button("Kruisjes", 'warning', self._on_place_crosses_click, width=90)
        self.add_footer_button("Verwijder Labels", 'danger', self._on_delete_labels_click, width=120)
        self.add_footer_button("Labels", 'warning', self._on_place_labels_click, width=80)
        self.add_footer_button("Export CSV", 'primary', self._on_export_click, width=100)
        self.add_footer_button("Vergelijk Revit", 'secondary', self._on_compare_click, width=120)
        self.add_footer_button("Vernieuwen", 'secondary', self._on_refresh_click, width=100)
        self.add_footer_button("?", 'icon', self._on_help_click, width=40)
        
        self.set_subtitle("Laadt IFC data...")
    
    def _load_data(self):
        try:
            self.all_data = scan_ifc_openings("")
            self._apply_filter()
        except Exception as ex:
            self.show_error("Fout bij laden: {}".format(str(ex)))
    
    def _apply_filter(self):
        filter_text = self.txt_filter.Text.strip()
        
        if filter_text:
            self.filtered_data = [w for w in self.all_data if w['type_name'].upper().startswith(filter_text.upper())]
        else:
            self.filtered_data = list(self.all_data)
        
        self.grid.Rows.Clear()
        for item in self.filtered_data:
            self.grid.Rows.Add(
                item['type_name'],
                item['category'][0],
                str(item['mark']),
                "{:.0f}".format(item['x_mm']),
                "{:.0f}".format(item['y_mm']),
                "{:.0f}".format(item['z_mm']),
                "{:.0f}".format(item['width_mm']),
                "{:.0f}".format(item['height_mm'])
            )
        
        self.preview.set_data(self.filtered_data)
        
        total = len(self.all_data)
        filtered = len(self.filtered_data)
        win_count = sum(1 for x in self.filtered_data if x['category'] == 'Window')
        door_count = sum(1 for x in self.filtered_data if x['category'] == 'Door')
        
        self.lbl_count.Text = "{} van {} ({} W, {} D)".format(filtered, total, win_count, door_count)
        self.set_subtitle("{} elementen".format(filtered))
    
    def _on_filter_changed(self, sender, args):
        self._apply_filter()
    
    def _on_grid_selection_changed(self, sender, args):
        if self.grid.SelectedRows.Count > 0:
            self.preview.set_selected(self.grid.SelectedRows[0].Index)
        else:
            self.preview.set_selected(-1)
    
    def _on_refresh_click(self, sender, args):
        self.set_subtitle("Vernieuwen...")
        self._load_data()
    
    def _on_compare_click(self, sender, args):
        if not self.filtered_data:
            self.show_warning("Geen data om te vergelijken.")
            return
        
        dialog = CompareDialog(self.filtered_data, self.txt_filter.Text.strip())
        dialog.ShowDialog()
    
    def _on_place_crosses_click(self, sender, args):
        if not self.filtered_data:
            self.show_warning("Geen elementen.")
            return
        
        if not self.ask_confirm("{} hulpkruisjes plaatsen?".format(len(self.filtered_data))):
            return
        
        placed, failed, errors = place_helper_crosses(self.filtered_data)
        
        msg = "Geplaatst: {} kruisjes".format(placed)
        if failed > 0:
            msg += "\nMislukt: {}".format(failed)
        
        self.show_info(msg) if failed == 0 else self.show_warning(msg)
    
    def _on_place_labels_click(self, sender, args):
        symbol = get_text_label_symbol()
        if not symbol:
            self.show_error("Family '{}' niet gevonden!".format(TEXT_LABEL_FAMILY))
            return
        
        k_types_filter = set(item['type_name'] for item in self.filtered_data)
        placed_instances = get_placed_revit_instances()
        
        total_placed = sum(len(instances) for k_type, instances in placed_instances.items() if k_type in k_types_filter)
        
        if total_placed == 0:
            self.show_warning("Geen geplaatste kozijnen gevonden.")
            return
        
        if not self.ask_confirm("Labels plaatsen bij {} kozijnen?".format(total_placed)):
            return
        
        placed, failed, errors = place_labels_for_placed_kozijns(k_types_filter)
        
        msg = "Labels geplaatst: {}".format(placed)
        if failed > 0:
            msg += "\nMislukt: {}".format(failed)
        
        self.show_info(msg) if failed == 0 else self.show_warning(msg)
    
    def _on_delete_labels_click(self, sender, args):
        symbol = get_text_label_symbol()
        if not symbol:
            self.show_warning("Family '{}' niet in project.".format(TEXT_LABEL_FAMILY))
            return
        
        instances = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).ToElements()
        label_count = sum(1 for inst in instances if inst.Symbol.FamilyName == TEXT_LABEL_FAMILY)
        
        if label_count == 0:
            self.show_info("Geen labels om te verwijderen.")
            return
        
        if not self.ask_confirm("{} labels verwijderen?".format(label_count)):
            return
        
        deleted = delete_existing_labels()
        self.show_info("Verwijderd: {}".format(deleted))
    
    def _on_export_click(self, sender, args):
        if not self.filtered_data:
            self.show_warning("Geen data.")
            return
        
        filter_text = self.txt_filter.Text.strip() or "all"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = "kozijnen_{}_placement_{}.csv".format(filter_text, timestamp)
        
        save_path = self.save_file_dialog(filter="CSV (*.csv)|*.csv", filename=default_name)
        
        if save_path:
            try:
                count = export_placement_csv(self.filtered_data, save_path)
                self.show_info("Geëxporteerd: {} elementen".format(count))
                os.startfile(os.path.dirname(save_path))
            except Exception as ex:
                self.show_error("Export fout: {}".format(str(ex)))


    def _on_help_click(self, sender, args):
        """Toon help informatie en credits in 3BM huisstijl."""
        dialog = HelpDialog()
        dialog.ShowDialog()


class HelpDialog(BaseForm):
    """Help en credits dialog in 3BM huisstijl."""
    
    def __init__(self):
        super(HelpDialog, self).__init__(
            "Help - IFC Kozijn Analyzer v3.9",
            width=600, height=700
        )
        self.set_subtitle("Documentatie & Credits")
        self._setup_ui()
    
    def _setup_ui(self):
        from System.Windows.Forms import RichTextBox, ScrollBars, BorderStyle
        
        margin = int(DPIScaler.scale(15))
        y = margin
        
        # Workflow sectie
        lbl_workflow = UIFactory.create_label("WORKFLOW", bold=True, color=Huisstijl.TEAL)
        lbl_workflow.Location = Point(margin, y)
        self.pnl_content.Controls.Add(lbl_workflow)
        y += int(DPIScaler.scale(25))
        
        workflow_text = """1. Filter op K-type prefix (bijv. "K" of "SK-K")
2. Bekijk IFC kozijnen in grid en preview
3. Klik "Vergelijk Revit" om te zien wat ontbreekt
4. Plaats ontbrekende kozijnen met één klik"""
        
        lbl_wf = UIFactory.create_label(workflow_text)
        lbl_wf.Location = Point(margin + 10, y)
        lbl_wf.Size = Size(int(DPIScaler.scale(550)), int(DPIScaler.scale(70)))
        self.pnl_content.Controls.Add(lbl_wf)
        y += int(DPIScaler.scale(80))
        
        # Functies sectie
        lbl_func = UIFactory.create_label("FUNCTIES", bold=True, color=Huisstijl.TEAL)
        lbl_func.Location = Point(margin, y)
        self.pnl_content.Controls.Add(lbl_func)
        y += int(DPIScaler.scale(25))
        
        functions = [
            ("Kruisjes", "Plaats 3D markers op IFC posities (verificatie)"),
            ("Vergelijk Revit", "Toon welke families beschikbaar/geplaatst zijn"),
            ("Plaats Ontbrekende", "Automatisch plaatsen met correcte orientatie"),
            ("Check Orientatie", "Detecteer en corrigeer verkeerd geplaatste kozijnen"),
            ("Labels", "Plaats 3D tekst labels met K-type codes"),
            ("Export CSV", "Exporteer placement data"),
        ]
        
        for name, desc in functions:
            lbl_name = UIFactory.create_label("• " + name, bold=True)
            lbl_name.Location = Point(margin + 10, y)
            lbl_name.Size = Size(int(DPIScaler.scale(140)), int(DPIScaler.scale(20)))
            self.pnl_content.Controls.Add(lbl_name)
            
            lbl_desc = UIFactory.create_label(desc, color=Huisstijl.TEXT_SECONDARY)
            lbl_desc.Location = Point(margin + 150, y)
            lbl_desc.Size = Size(int(DPIScaler.scale(400)), int(DPIScaler.scale(20)))
            self.pnl_content.Controls.Add(lbl_desc)
            y += int(DPIScaler.scale(22))
        
        y += int(DPIScaler.scale(15))
        
        # Tips sectie
        lbl_tips = UIFactory.create_label("TIPS", bold=True, color=Huisstijl.YELLOW)
        lbl_tips.Location = Point(margin, y)
        self.pnl_content.Controls.Add(lbl_tips)
        y += int(DPIScaler.scale(25))
        
        tips_text = """• Controleer eerst met kruisjes of IFC posities kloppen
• Check Orientatie na plaatsing voor verificatie
• Labels helpen bij visuele controle in 3D views
• Install Depth werkt met instance én type parameters"""
        
        lbl_t = UIFactory.create_label(tips_text)
        lbl_t.Location = Point(margin + 10, y)
        lbl_t.Size = Size(int(DPIScaler.scale(550)), int(DPIScaler.scale(75)))
        self.pnl_content.Controls.Add(lbl_t)
        y += int(DPIScaler.scale(90))
        
        # Divider
        pnl_div = Panel()
        pnl_div.BackColor = Huisstijl.TEAL
        pnl_div.Location = Point(margin, y)
        pnl_div.Size = Size(int(DPIScaler.scale(550)), 2)
        self.pnl_content.Controls.Add(pnl_div)
        y += int(DPIScaler.scale(20))
        
        # Credits sectie
        lbl_credits = UIFactory.create_label("CREDITS", bold=True, color=Huisstijl.VIOLET)
        lbl_credits.Location = Point(margin, y)
        self.pnl_content.Controls.Add(lbl_credits)
        y += int(DPIScaler.scale(30))
        
        # 3BM Logo tekst
        lbl_3bm = UIFactory.create_label("3BM Bouwkunde", bold=True, font_size=14)
        lbl_3bm.ForeColor = Huisstijl.VIOLET
        lbl_3bm.Location = Point(margin + 10, y)
        self.pnl_content.Controls.Add(lbl_3bm)
        y += int(DPIScaler.scale(22))
        
        lbl_3bm_sub = UIFactory.create_label("Bouwfysica • Constructies • BIM", color=Huisstijl.TEXT_SECONDARY)
        lbl_3bm_sub.Location = Point(margin + 10, y)
        self.pnl_content.Controls.Add(lbl_3bm_sub)
        y += int(DPIScaler.scale(18))
        
        lbl_www = UIFactory.create_label("www.3bm.nl", color=Huisstijl.TEAL)
        lbl_www.Location = Point(margin + 10, y)
        self.pnl_content.Controls.Add(lbl_www)
        y += int(DPIScaler.scale(35))
        
        # Claude credit
        lbl_claude = UIFactory.create_label("AI-Assisted Development", bold=True, font_size=11)
        lbl_claude.Location = Point(margin + 10, y)
        self.pnl_content.Controls.Add(lbl_claude)
        y += int(DPIScaler.scale(20))
        
        lbl_claude_sub = UIFactory.create_label("Powered by Claude (Anthropic)", color=Huisstijl.TEXT_SECONDARY)
        lbl_claude_sub.Location = Point(margin + 10, y)
        self.pnl_content.Controls.Add(lbl_claude_sub)
        y += int(DPIScaler.scale(35))
        
        # Project reference
        lbl_ref = UIFactory.create_label("Gebaseerd op workflows uit Project 2964", color=Huisstijl.TEXT_SECONDARY, font_size=9)
        lbl_ref.Location = Point(margin + 10, y)
        self.pnl_content.Controls.Add(lbl_ref)
        y += int(DPIScaler.scale(30))
        
        # Copyright
        lbl_copy = UIFactory.create_label("© 2026 3BM Bouwkunde", color=Huisstijl.TEXT_SECONDARY, font_size=9)
        lbl_copy.Location = Point(margin + 10, y)
        self.pnl_content.Controls.Add(lbl_copy)
        
        # Sluit knop
        self.add_footer_button("Sluiten", 'primary', lambda s, e: self.Close(), width=100)


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    """Entry point."""
    global doc
    
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    
    if not doc:
        from pyrevit import forms
        forms.alert("Open eerst een Revit project.", title="IFC Kozijn Analyzer")
        return
    
    ifc_links = get_ifc_links()
    if not ifc_links:
        from pyrevit import forms
        forms.alert("Geen IFC links gevonden in dit model.", exitscript=True)
    
    form = IFCKozijnAnalyzerForm()
    form.ShowDialog()


if __name__ == "__main__":
    main()
