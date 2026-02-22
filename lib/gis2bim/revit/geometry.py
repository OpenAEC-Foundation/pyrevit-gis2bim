# -*- coding: utf-8 -*-
"""
Revit Geometry Creation
=======================

Functies voor het aanmaken van Revit geometry elementen vanuit GIS data.
Ondersteunt Model Lines, Text Notes en coordinaat transformaties.

Gebruik:
    from gis2bim.revit.geometry import create_model_lines, create_text_note

    # Maak perceelgrens lijnen
    lines = create_model_lines(doc, polylines, origin_rd=(155000, 463000))
"""

import math
import os

def _geom_log(msg):
    """Log naar bestand voor debugging."""
    print(msg)
    try:
        log_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
                os.path.abspath(__file__))))),
            "logs", "WFS_debug.log"
        )
        with open(log_path, "a") as f:
            f.write(msg + "\n")
    except Exception:
        pass

# Revit context check
IN_REVIT = False

try:
    from Autodesk.Revit.DB import (
        Transaction,
        XYZ,
        Line,
        Plane,
        SketchPlane,
        FilteredElementCollector,
        BuiltInCategory,
        TextNoteType,
        TextNote,
        ModelCurve,
        ElementId,
        Element,
        FilledRegion,
        FilledRegionType,
        CurveLoop,
        Material,
        Color,
        Options,
    )
    IN_REVIT = True
except ImportError:
    pass


# =============================================================================
# Constanten
# =============================================================================

FEET_TO_METER = 0.3048
METER_TO_FEET = 1.0 / FEET_TO_METER


# =============================================================================
# Coordinaat Transformaties
# =============================================================================

def rd_to_revit_xyz(rd_x, rd_y, origin_rd_x, origin_rd_y, z=0.0):
    """
    Converteer RD coordinaten naar Revit XYZ (lokaal, in feet).

    Args:
        rd_x: RD X coordinaat in meters
        rd_y: RD Y coordinaat in meters
        origin_rd_x: RD X van project origin (Survey Point)
        origin_rd_y: RD Y van project origin (Survey Point)
        z: Hoogte in meters (default 0)

    Returns:
        XYZ object in Revit internal units (feet)
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    # Relatief tov origin, geconverteerd naar feet
    local_x = (rd_x - origin_rd_x) * METER_TO_FEET
    local_y = (rd_y - origin_rd_y) * METER_TO_FEET
    local_z = z * METER_TO_FEET

    return XYZ(local_x, local_y, local_z)


def meters_to_feet(meters):
    """Converteer meters naar feet."""
    return meters * METER_TO_FEET


def feet_to_meters(feet):
    """Converteer feet naar meters."""
    return feet * FEET_TO_METER


# =============================================================================
# Model Lines
# =============================================================================

def create_model_lines(doc, polylines, origin_rd, line_style=None, z_offset=0.0):
    """
    Maak Model Lines in Revit van polylines.

    Args:
        doc: Revit document
        polylines: Lijst van polylines, elke polyline is lijst van (x, y) tuples
        origin_rd: Tuple (rd_x, rd_y) van project origin
        line_style: Optioneel GraphicsStyle element voor lijnstijl
        z_offset: Hoogte offset in meters

    Returns:
        Lijst van aangemaakte ModelCurve ElementIds
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    origin_x, origin_y = origin_rd
    created_ids = []
    skipped_short = 0
    errors = 0

    _geom_log("GIS2BIM geometry: create_model_lines start")
    _geom_log("  polylines: {0}, origin: ({1}, {2})".format(
        len(polylines), origin_x, origin_y))

    # Haal of maak SketchPlane op Z=0
    sketch_plane = _get_or_create_sketch_plane(doc, z_offset)
    _geom_log("  SketchPlane: {0}".format(sketch_plane.Id))

    for pi, polyline in enumerate(polylines):
        if len(polyline) < 2:
            continue

        # Log eerste polyline als voorbeeld
        if pi == 0:
            _geom_log("  Eerste polyline: {0} punten".format(len(polyline)))
            _geom_log("    pt[0]: {0}".format(polyline[0]))
            _geom_log("    pt[1]: {0}".format(polyline[1]))
            xyz_test = rd_to_revit_xyz(
                polyline[0][0], polyline[0][1], origin_x, origin_y, z_offset)
            _geom_log("    xyz[0]: ({0:.2f}, {1:.2f}, {2:.2f})".format(
                xyz_test.X, xyz_test.Y, xyz_test.Z))

        # Maak lijnsegmenten
        for i in range(len(polyline) - 1):
            pt1 = polyline[i]
            pt2 = polyline[i + 1]

            xyz1 = rd_to_revit_xyz(pt1[0], pt1[1], origin_x, origin_y, z_offset)
            xyz2 = rd_to_revit_xyz(pt2[0], pt2[1], origin_x, origin_y, z_offset)

            # Skip zeer korte lijnen (< 1mm in Revit)
            dist = xyz1.DistanceTo(xyz2)
            if dist < 0.003:  # ~1mm in feet
                skipped_short += 1
                continue

            try:
                line = Line.CreateBound(xyz1, xyz2)
                model_curve = doc.Create.NewModelCurve(line, sketch_plane)

                if line_style is not None:
                    model_curve.LineStyle = line_style

                created_ids.append(model_curve.Id)

            except Exception as e:
                errors += 1
                if errors <= 3:
                    _geom_log("Fout bij aanmaken lijn: {0}".format(e))

    _geom_log("  Resultaat: {0} lijnen, {1} te kort, {2} errors".format(
        len(created_ids), skipped_short, errors))
    return created_ids


def create_model_lines_from_features(doc, features, origin_rd, line_style=None, z_offset=0.0):
    """
    Maak Model Lines van WFSFeature objecten.

    Args:
        doc: Revit document
        features: Lijst van WFSFeature objecten met polygon geometry
        origin_rd: Tuple (rd_x, rd_y) van project origin
        line_style: Optioneel GraphicsStyle element
        z_offset: Hoogte offset in meters

    Returns:
        Tuple van (aantal_lijnen, aantal_features)
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    polylines = []

    for feature in features:
        geom = feature.geometry
        geom_type = feature.geometry_type

        if geom_type == "polygon" and geom:
            # Sluit polygon door eerste punt toe te voegen
            ring = list(geom)
            if ring and ring[0] != ring[-1]:
                ring.append(ring[0])
            polylines.append(ring)

        elif geom_type == "multipolygon" and geom:
            # Meerdere polygonen
            for polygon in geom:
                ring = list(polygon)
                if ring and ring[0] != ring[-1]:
                    ring.append(ring[0])
                polylines.append(ring)

        elif geom_type == "line" and geom:
            polylines.append(geom)

        elif geom_type == "multiline" and geom:
            polylines.extend(geom)

    ids = create_model_lines(doc, polylines, origin_rd, line_style, z_offset)
    return (ids, len(features))


def _get_or_create_sketch_plane(doc, z_offset=0.0):
    """Haal bestaande of maak nieuwe SketchPlane op niveau Z."""
    z_feet = z_offset * METER_TO_FEET

    # Maak vlak op Z niveau
    origin = XYZ(0, 0, z_feet)
    normal = XYZ(0, 0, 1)
    plane = Plane.CreateByNormalAndOrigin(normal, origin)

    return SketchPlane.Create(doc, plane)


# =============================================================================
# Text Notes
# =============================================================================

def create_text_notes(doc, view, annotations, origin_rd, text_type=None):
    """
    Maak Text Notes in Revit voor annotaties.

    Args:
        doc: Revit document
        view: View waarin de text notes worden geplaatst
        annotations: Lijst van dicts met {text, x, y, rotation}
        origin_rd: Tuple (rd_x, rd_y) van project origin
        text_type: Optioneel TextNoteType element

    Returns:
        Lijst van aangemaakte TextNote ElementIds
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    origin_x, origin_y = origin_rd
    created_ids = []
    errors = 0

    _geom_log("GIS2BIM geometry: create_text_notes start")
    _geom_log("  annotations: {0}, origin: ({1}, {2})".format(
        len(annotations), origin_x, origin_y))
    _geom_log("  view: {0} (Id={1})".format(view.Name, view.Id))

    # Haal default text type als niet opgegeven
    if text_type is None:
        text_type = _get_default_text_type(doc)

    if text_type is None:
        _geom_log("FOUT: Geen TextNoteType beschikbaar in document!")
        return []

    try:
        type_name = Element.Name.__get__(text_type)
    except Exception:
        type_name = "ID={0}".format(text_type.Id.IntegerValue)
    _geom_log("  TextNoteType: {0} (Id={1})".format(type_name, text_type.Id))

    for i, ann in enumerate(annotations):
        text = ann.get("text", "")
        if not text:
            continue

        x = ann.get("x", 0)
        y = ann.get("y", 0)
        rotation = ann.get("rotation", 0)

        xyz = rd_to_revit_xyz(x, y, origin_x, origin_y)

        if i == 0:
            _geom_log("  Eerste annotatie: text={0}, rd=({1}, {2}), xyz=({3:.2f}, {4:.2f})".format(
                repr(text), x, y, xyz.X, xyz.Y))

        try:
            text_note = TextNote.Create(
                doc,
                view.Id,
                xyz,
                text,
                text_type.Id
            )

            # Roteer als nodig
            if abs(rotation) > 0.1:
                _rotate_text_note(doc, text_note, xyz, rotation)

            created_ids.append(text_note.Id)

        except Exception as e:
            errors += 1
            if errors <= 3:
                _geom_log("Fout bij text note '{0}': {1}".format(text, e))

    _geom_log("  Resultaat: {0} text notes, {1} errors".format(
        len(created_ids), errors))
    return created_ids


def create_text_notes_from_features(doc, view, features, origin_rd, text_type=None):
    """
    Maak Text Notes van WFSFeature objecten met labels.

    Args:
        doc: Revit document
        view: View voor plaatsing
        features: Lijst van WFSFeature objecten met label info
        origin_rd: Tuple (rd_x, rd_y) van project origin
        text_type: Optioneel TextNoteType element

    Returns:
        Tuple van (aantal_notes, aantal_features_met_label)
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    annotations = []
    features_with_label = 0

    for feature in features:
        if not feature.label:
            continue

        features_with_label += 1

        # Bepaal positie
        if feature.label_position:
            x, y = feature.label_position
        elif feature.geometry_type == "point" and feature.geometry:
            x, y = feature.geometry
        elif feature.geometry_type == "polygon" and feature.geometry:
            # Gebruik centroid van polygon
            x, y = _calculate_centroid(feature.geometry)
        else:
            continue

        annotations.append({
            "text": feature.label,
            "x": x,
            "y": y,
            "rotation": feature.label_rotation
        })

    ids = create_text_notes(doc, view, annotations, origin_rd, text_type)
    return (ids, features_with_label)


def _get_default_text_type(doc):
    """Haal eerste beschikbare TextNoteType op."""
    collector = FilteredElementCollector(doc)
    types = collector.OfClass(TextNoteType).ToElements()

    if types:
        return types[0]
    return None


def _create_text_note_options(rotation):
    """Maak TextNoteOptions (indien beschikbaar in Revit versie)."""
    try:
        from Autodesk.Revit.DB import TextNoteOptions
        options = TextNoteOptions()
        options.Rotation = math.radians(rotation)
        return options
    except ImportError:
        return None


def _rotate_text_note(doc, text_note, position, rotation_deg):
    """Roteer een text note rond zijn positie."""
    try:
        from Autodesk.Revit.DB import ElementTransformUtils

        axis = Line.CreateBound(
            position,
            XYZ(position.X, position.Y, position.Z + 1)
        )

        ElementTransformUtils.RotateElement(
            doc,
            text_note.Id,
            axis,
            math.radians(rotation_deg)
        )
    except Exception as e:
        print("Kon text note niet roteren: {0}".format(e))


def _calculate_centroid(polygon):
    """Bereken centroid van een polygon."""
    if not polygon:
        return (0, 0)

    sum_x = sum(p[0] for p in polygon)
    sum_y = sum(p[1] for p in polygon)
    n = len(polygon)

    return (sum_x / n, sum_y / n)


# =============================================================================
# Line Styles
# =============================================================================

def get_line_style(doc, style_name):
    """
    Haal een lijnstijl op via naam.

    Args:
        doc: Revit document
        style_name: Naam van de lijnstijl (bijv. "<Thin Lines>")

    Returns:
        GraphicsStyle element of None
    """
    if not IN_REVIT:
        return None

    try:
        from Autodesk.Revit.DB import GraphicsStyle

        collector = FilteredElementCollector(doc)
        styles = collector.OfClass(GraphicsStyle).ToElements()

        for style in styles:
            if style.Name == style_name:
                return style

        return None

    except Exception as e:
        print("Fout bij ophalen lijnstijl: {0}".format(e))
        return None


def get_available_line_styles(doc):
    """
    Haal alle beschikbare lijnstijlen op.

    Returns:
        Lijst van (naam, element) tuples
    """
    if not IN_REVIT:
        return []

    try:
        from Autodesk.Revit.DB import GraphicsStyle, GraphicsStyleType

        collector = FilteredElementCollector(doc)
        styles = collector.OfClass(GraphicsStyle).ToElements()

        result = []
        for style in styles:
            # Filter op curve styles
            if style.GraphicsStyleCategory is not None:
                cat = style.GraphicsStyleCategory
                if cat.CategoryType == CategoryType.Model:
                    result.append((style.Name, style))

        return result

    except Exception:
        return []


# NEN-standaard KLIC kleurenschema
KLIC_COLORS = {
    "electricity": (255, 0, 0),       # Rood
    "telecom":     (0, 128, 0),       # Groen
    "gas":         (255, 255, 0),     # Geel
    "water":       (0, 0, 255),       # Blauw
    "sewer":       (139, 90, 43),     # Bruin
    "duct":        (128, 128, 128),   # Grijs
    "other":       (200, 200, 200),   # Lichtgrijs
}


def create_or_get_line_style(doc, name, rgb, line_weight=3):
    """
    Maak of haal een lijnstijl subcategorie onder OST_Lines.

    Args:
        doc: Revit document
        name: Naam (bijv. "KLIC_Elektriciteitskabel")
        rgb: Tuple (R, G, B)
        line_weight: Lijndikte (default 3)

    Returns:
        GraphicsStyle element, of None bij fout
    """
    if not IN_REVIT:
        return None

    from Autodesk.Revit.DB import (
        GraphicsStyle, GraphicsStyleType, Color as RevitColor,
        LinePatternElement
    )

    try:
        lines_cat = doc.Settings.Categories.get_Item(
            BuiltInCategory.OST_Lines)

        # Zoek bestaande subcategorie
        for subcat in lines_cat.SubCategories:
            if subcat.Name == name:
                return subcat.GetGraphicsStyle(
                    GraphicsStyleType.Projection)

        # Maak nieuwe subcategorie
        new_cat = doc.Settings.Categories.NewSubcategory(lines_cat, name)

        # Stel kleur in
        new_cat.LineColor = RevitColor(rgb[0], rgb[1], rgb[2])

        # Stel lijndikte in
        new_cat.SetLineWeight(line_weight, GraphicsStyleType.Projection)

        # Stel lijnpatroon in (solid/doorgetrokken)
        solid_pattern = LinePatternElement.GetSolidPatternId()
        new_cat.SetLinePatternId(solid_pattern, GraphicsStyleType.Projection)

        return new_cat.GetGraphicsStyle(GraphicsStyleType.Projection)

    except Exception as e:
        _geom_log("create_or_get_line_style fout ({0}): {1}".format(name, e))
        return None


# =============================================================================
# Text Types
# =============================================================================

def get_text_type(doc, type_name):
    """
    Haal een TextNoteType op via naam.

    Args:
        doc: Revit document
        type_name: Naam van het text type

    Returns:
        TextNoteType element of None
    """
    if not IN_REVIT:
        return None

    try:
        collector = FilteredElementCollector(doc)
        types = collector.OfClass(TextNoteType).ToElements()

        for text_type in types:
            if text_type.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME
            ).AsString() == type_name:
                return text_type

        return None

    except Exception:
        return None


def get_available_text_types(doc):
    """
    Haal alle beschikbare TextNoteTypes op.

    Returns:
        Lijst van (naam, element) tuples
    """
    if not IN_REVIT:
        return []

    try:
        from Autodesk.Revit.DB import BuiltInParameter

        collector = FilteredElementCollector(doc)
        types = collector.OfClass(TextNoteType).ToElements()

        result = []
        for text_type in types:
            try:
                name_param = text_type.get_Parameter(
                    BuiltInParameter.ALL_MODEL_TYPE_NAME
                )
                if name_param:
                    result.append((name_param.AsString(), text_type))
            except Exception:
                pass

        return result

    except Exception:
        return []


# =============================================================================
# Filled Regions
# =============================================================================

def create_filled_regions(doc, view, polygons, origin_rd, filled_type=None, z_offset=0.0):
    """
    Maak Filled Regions in Revit van polygonen.

    Args:
        doc: Revit document
        view: View waarin de filled regions worden geplaatst
        polygons: Lijst van polygonen, elke polygon is lijst van (x, y) tuples
        origin_rd: Tuple (rd_x, rd_y) van project origin
        filled_type: Optioneel FilledRegionType element
        z_offset: Hoogte offset in meters

    Returns:
        Lijst van aangemaakte FilledRegion ElementIds
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    from System.Collections.Generic import List as GenericList

    origin_x, origin_y = origin_rd
    created_ids = []
    skipped = 0
    errors = 0

    _geom_log("GIS2BIM geometry: create_filled_regions start")
    _geom_log("  polygons: {0}, origin: ({1}, {2})".format(
        len(polygons), origin_x, origin_y))

    # Haal default filled region type als niet opgegeven
    if filled_type is None:
        filled_type = _get_default_filled_region_type(doc)

    if filled_type is None:
        _geom_log("FOUT: Geen FilledRegionType beschikbaar!")
        return []

    _geom_log("  FilledRegionType Id: {0}".format(filled_type.Id))

    for pi, polygon in enumerate(polygons):
        if len(polygon) < 3:
            skipped += 1
            continue

        try:
            # Converteer punten naar Revit XYZ
            raw_points = []
            for pt in polygon:
                xyz = rd_to_revit_xyz(pt[0], pt[1], origin_x, origin_y, z_offset)
                raw_points.append(xyz)

            # Filter opeenvolgende duplicaten
            points = [raw_points[0]]
            for j in range(1, len(raw_points)):
                if raw_points[j].DistanceTo(points[-1]) >= 0.003:
                    points.append(raw_points[j])

            # Verwijder sluitpunt als het al dicht bij eerste punt ligt
            if len(points) > 1 and points[0].DistanceTo(points[-1]) < 0.003:
                points = points[:-1]

            # Te weinig unieke punten voor een vlak
            if len(points) < 3:
                skipped += 1
                continue

            # Log eerste polygon als voorbeeld
            if pi == 0:
                _geom_log("  Eerste polygon: {0} punten".format(len(points)))
                _geom_log("    pt[0]: ({0:.2f}, {1:.2f})".format(
                    points[0].X, points[0].Y))

            # Bouw gesloten CurveLoop
            curve_loop = CurveLoop()
            for i in range(len(points)):
                p1 = points[i]
                p2 = points[(i + 1) % len(points)]
                line = Line.CreateBound(p1, p2)
                curve_loop.Append(line)

            # Maak lijst van CurveLoops
            loops = GenericList[CurveLoop]()
            loops.Add(curve_loop)

            # Maak FilledRegion
            filled_region = FilledRegion.Create(
                doc, filled_type.Id, view.Id, loops
            )
            created_ids.append(filled_region.Id)

        except Exception as e:
            errors += 1
            if errors <= 3:
                _geom_log("Fout bij filled region: {0}".format(e))

    _geom_log("  Resultaat: {0} filled regions, {1} overgeslagen, {2} errors".format(
        len(created_ids), skipped, errors))
    return created_ids


def create_filled_regions_from_features(doc, view, features, origin_rd,
                                        filled_type=None, z_offset=0.0):
    """
    Maak Filled Regions van WFSFeature objecten met polygon geometry.

    Args:
        doc: Revit document
        view: View waarin de filled regions worden geplaatst
        features: Lijst van WFSFeature objecten met polygon geometry
        origin_rd: Tuple (rd_x, rd_y) van project origin
        filled_type: Optioneel FilledRegionType element
        z_offset: Hoogte offset in meters

    Returns:
        Tuple van (aantal_regions, aantal_features)
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    polygons = []

    for feature in features:
        geom = feature.geometry
        geom_type = feature.geometry_type

        if geom_type == "polygon" and geom:
            polygons.append(geom)
        elif geom_type == "multipolygon" and geom:
            polygons.extend(geom)

    ids = create_filled_regions(doc, view, polygons, origin_rd, filled_type, z_offset)
    return (ids, len(features))


def _get_default_filled_region_type(doc):
    """Haal eerste beschikbare FilledRegionType op."""
    try:
        collector = FilteredElementCollector(doc)
        types = collector.OfClass(FilledRegionType).ToElements()
        if types:
            return types[0]
    except Exception:
        pass
    return None


def get_filled_region_type(doc, type_name):
    """
    Haal een FilledRegionType op via naam.

    Args:
        doc: Revit document
        type_name: Naam van het filled region type

    Returns:
        FilledRegionType element of None
    """
    if not IN_REVIT:
        return None

    try:
        collector = FilteredElementCollector(doc)
        types = collector.OfClass(FilledRegionType).ToElements()

        for frt in types:
            try:
                name = Element.Name.__get__(frt)
                if name == type_name:
                    return frt
            except Exception:
                pass

        return None
    except Exception:
        return None


# =============================================================================
# 3D Geometry - Cilinders (voor sonderingen/boringen)
# =============================================================================

def create_cylinder_solid(center_x_feet, center_y_feet, z_bottom_feet,
                          radius_feet, height_feet, segments=24):
    """
    Maak een cilindrische Solid op de opgegeven positie.

    Gebruikt een polygon-benadering (24-gon) voor maximale
    compatibiliteit met alle Revit versies en IronPython.

    Args:
        center_x_feet: X positie in feet (Revit internal units)
        center_y_feet: Y positie in feet
        z_bottom_feet: Z positie van de onderkant in feet
        radius_feet: Straal van de cilinder in feet
        height_feet: Hoogte van de cilinder in feet
        segments: Aantal segmenten voor de cirkel (default 24)

    Returns:
        Solid object
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    from Autodesk.Revit.DB import GeometryCreationUtilities
    from System.Collections.Generic import List as GenericList

    # Bouw circulair profiel als polygon (betrouwbaarder dan Arc)
    points = []
    for i in range(segments):
        angle = 2.0 * math.pi * i / segments
        px = center_x_feet + radius_feet * math.cos(angle)
        py = center_y_feet + radius_feet * math.sin(angle)
        points.append(XYZ(px, py, z_bottom_feet))

    profile = CurveLoop()
    for i in range(segments):
        p1 = points[i]
        p2 = points[(i + 1) % segments]
        profile.Append(Line.CreateBound(p1, p2))

    loops = GenericList[CurveLoop]()
    loops.Add(profile)

    direction = XYZ(0, 0, 1)
    solid = GeometryCreationUtilities.CreateExtrusionGeometry(
        loops, direction, height_feet
    )

    return solid


def create_directshape(doc, solids, name=None):
    """
    Maak een DirectShape element van een lijst Solids.

    Args:
        doc: Revit document
        solids: Lijst van Solid objecten
        name: Optionele naam voor het element

    Returns:
        DirectShape element
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    from Autodesk.Revit.DB import DirectShape, GeometryObject
    from System.Collections.Generic import List as GenericList

    ds = DirectShape.CreateElement(
        doc, ElementId(BuiltInCategory.OST_GenericModel)
    )

    shape_list = GenericList[GeometryObject]()
    for solid in solids:
        shape_list.Add(solid)
    ds.SetShape(shape_list)

    if name:
        try:
            ds.SetName(name)
        except Exception:
            pass

    return ds


def set_element_color(doc, view, element_id, rgb):
    """
    Stel element kleur in via OverrideGraphicSettings.

    Args:
        doc: Revit document
        view: Revit View object
        element_id: ElementId van het element
        rgb: Tuple (R, G, B) met kleurwaarden 0-255
    """
    if not IN_REVIT:
        return

    from Autodesk.Revit.DB import (
        OverrideGraphicSettings, Color, FillPatternElement
    )

    # Zoek solid fill pattern
    solid_fill_id = None
    patterns = FilteredElementCollector(doc).OfClass(
        FillPatternElement
    ).ToElements()
    for pat in patterns:
        try:
            if pat.GetFillPattern().IsSolidFill:
                solid_fill_id = pat.Id
                break
        except Exception:
            continue

    if solid_fill_id is None:
        return

    color = Color(rgb[0], rgb[1], rgb[2])
    ogs = OverrideGraphicSettings()

    try:
        # Revit 2019+
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
        ogs.SetSurfaceBackgroundPatternColor(color)
        ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
    except AttributeError:
        # Oudere Revit versies
        try:
            ogs.SetProjectionFillColor(color)
            ogs.SetProjectionFillPatternId(solid_fill_id)
        except Exception:
            pass

    view.SetElementOverrides(element_id, ogs)


# =============================================================================
# Materials (voor consistente kleuren in alle views)
# =============================================================================

# Sessie-cache: voorkomt herhaalde FilteredElementCollector queries
_material_cache = {}


def get_or_create_material(doc, name, rgb):
    """
    Zoek of maak een Material met de opgegeven naam en kleur.

    Gebruikt een sessie-cache om herhaalde collector-queries te vermijden.

    Args:
        doc: Revit document
        name: Naam voor het materiaal (bijv. "BRO - Klei")
        rgb: Tuple (R, G, B) met kleurwaarden 0-255

    Returns:
        ElementId van het materiaal
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    # Check cache
    cache_key = (name, rgb)
    if cache_key in _material_cache:
        # Verifieer dat gecached materiaal nog bestaat
        cached_id = _material_cache[cache_key]
        if doc.GetElement(cached_id) is not None:
            return cached_id

    # Zoek bestaand materiaal op naam
    collector = FilteredElementCollector(doc).OfClass(Material)
    for mat in collector:
        try:
            mat_name = Element.Name.__get__(mat)
        except Exception:
            mat_name = ""
        if mat_name == name:
            _material_cache[cache_key] = mat.Id
            # Werk kleur bij als nodig
            try:
                mat.Color = Color(rgb[0], rgb[1], rgb[2])
            except Exception:
                pass
            return mat.Id

    # Maak nieuw materiaal
    mat_id = Material.Create(doc, name)
    mat = doc.GetElement(mat_id)
    mat.Color = Color(rgb[0], rgb[1], rgb[2])

    # Maak materiaal niet-transparant
    try:
        mat.Transparency = 0
    except Exception:
        pass

    _material_cache[cache_key] = mat_id
    return mat_id


def set_element_material(doc, element, material_id):
    """
    Wijs materiaal toe aan een DirectShape element via doc.Paint().

    Paint() werkt op individuele vlakken en zorgt ervoor dat het materiaal
    zichtbaar is in alle views, 3D, en exports.

    Args:
        doc: Revit document
        element: DirectShape element
        material_id: ElementId van het materiaal

    Returns:
        True als succesvol, False bij fout
    """
    if not IN_REVIT:
        return False

    try:
        opts = Options()
        geom = element.get_Geometry(opts)
        if geom is None:
            return False

        face_index = 0
        for geom_obj in geom:
            solid = None
            # DirectShape geometry kan Solid of GeometryInstance zijn
            try:
                from Autodesk.Revit.DB import Solid, GeometryInstance
                if isinstance(geom_obj, Solid):
                    solid = geom_obj
                elif isinstance(geom_obj, GeometryInstance):
                    for inst_obj in geom_obj.GetInstanceGeometry():
                        if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                            for face in inst_obj.Faces:
                                doc.Paint(element.Id, face, material_id)
                                face_index += 1
                    continue
            except ImportError:
                pass

            if solid is None or solid.Volume <= 0:
                continue

            for face in solid.Faces:
                doc.Paint(element.Id, face, material_id)
                face_index += 1

        return face_index > 0

    except Exception as e:
        _geom_log("set_element_material fout: {0}".format(e))
        return False


def create_textured_material(doc, name, image_path, width_m, height_m):
    """
    Maak of update een materiaal met een bitmap textuur.

    Gebruikt de Revit AppearanceAsset API (Revit 2019+) om een
    diffuse bitmap in te stellen op een Generic materiaal.
    Valt terug op een groen materiaal als textuur setup mislukt.

    Args:
        doc: Revit document
        name: Naam voor het materiaal (bijv. "AHN - Luchtfoto")
        image_path: Absoluut pad naar afbeelding (JPG/PNG)
        width_m: Breedte van de afbeelding in meters (real-world)
        height_m: Hoogte van de afbeelding in meters (real-world)

    Returns:
        ElementId van het materiaal
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")

    from Autodesk.Revit.DB import AppearanceAssetElement

    # Zoek bestaand materiaal op naam
    existing_id = _find_material_by_name(doc, name)
    if existing_id:
        mat = doc.GetElement(existing_id)
        # Probeer textuur te updaten
        if mat.AppearanceAssetId and mat.AppearanceAssetId.IntegerValue != -1:
            _update_bitmap_texture(
                doc, mat.AppearanceAssetId,
                image_path, width_m, height_m)
        return existing_id

    # Maak nieuw materiaal
    mat_id = Material.Create(doc, name)
    mat = doc.GetElement(mat_id)
    mat.Color = Color(150, 180, 140)  # Lichtgroen fallback
    mat.Transparency = 0

    # Probeer textuur in te stellen
    aae_id = _setup_bitmap_appearance(
        doc, name, image_path, width_m, height_m)
    if aae_id:
        mat.AppearanceAssetId = aae_id

    return mat_id


def _find_material_by_name(doc, name):
    """Zoek een materiaal op naam.

    Returns:
        ElementId of None
    """
    collector = FilteredElementCollector(doc).OfClass(Material)
    for mat in collector:
        try:
            mat_name = Element.Name.__get__(mat)
        except Exception:
            mat_name = ""
        if mat_name == name:
            return mat.Id
    return None


def _setup_bitmap_appearance(doc, name, image_path, width_m, height_m):
    """Maak een AppearanceAssetElement met bitmap textuur.

    Zoekt een bestaand Generic appearance asset om te dupliceren.
    Als er geen bestaat, maakt er een aan vanuit de Revit asset library.

    Args:
        doc: Revit document
        name: Asset naam
        image_path: Pad naar bitmap
        width_m: Breedte in meters
        height_m: Hoogte in meters

    Returns:
        ElementId van het AppearanceAssetElement, of None bij fout
    """
    from Autodesk.Revit.DB import AppearanceAssetElement

    try:
        # Stap 1: Zoek een Generic appearance asset om te dupliceren
        source_aae = _find_generic_appearance_asset(doc)

        if source_aae is None:
            # Probeer vanuit Revit library
            source_aae = _create_generic_asset_from_library(doc)

        if source_aae is None:
            _geom_log("Geen Generic appearance asset gevonden")
            return None

        # Stap 2: Dupliceer het asset
        asset_name = "GIS2BIM {0}".format(name)
        new_aae = source_aae.Duplicate(asset_name)

        # Stap 3: Stel bitmap in
        _set_bitmap_on_asset(doc, new_aae.Id, image_path, width_m, height_m)

        return new_aae.Id

    except Exception as e:
        _geom_log("_setup_bitmap_appearance fout: {0}".format(e))
        return None


def _find_generic_appearance_asset(doc):
    """Zoek een bestaand AppearanceAssetElement met Generic schema.

    Returns:
        AppearanceAssetElement of None
    """
    from Autodesk.Revit.DB import AppearanceAssetElement

    collector = FilteredElementCollector(doc).OfClass(AppearanceAssetElement)
    for aae in collector:
        try:
            asset = aae.GetRenderingAsset()
            # Check of het Generic schema heeft (via generic_diffuse property)
            prop = asset.FindByName("generic_diffuse")
            if prop is not None:
                return aae
        except Exception:
            continue
    return None


def _create_generic_asset_from_library(doc):
    """Maak een Generic AppearanceAssetElement vanuit de Revit library.

    Valt terug op het dupliceren van het eerste beschikbare asset.

    Returns:
        AppearanceAssetElement of None
    """
    from Autodesk.Revit.DB import AppearanceAssetElement

    try:
        # Revit 2019+: haal assets uit de library
        from Autodesk.Revit.DB.Visual import AssetType
        lib_assets = doc.Application.GetAssets(AssetType.Appearance)

        for asset in lib_assets:
            # Zoek het "Generic" schema
            prop = asset.FindByName("generic_diffuse")
            if prop is not None:
                aae = AppearanceAssetElement.Create(
                    doc, "GIS2BIM_Generic_Base", asset)
                return aae

    except Exception as e:
        _geom_log("Library asset lookup: {0}".format(e))

    # Fallback: dupliceer eerste beschikbare AAE
    try:
        collector = FilteredElementCollector(doc).OfClass(
            AppearanceAssetElement)
        for aae in collector:
            return aae
    except Exception:
        pass

    return None


def _set_bitmap_on_asset(doc, aae_id, image_path, width_m, height_m):
    """Stel bitmap textuur in op een AppearanceAssetElement.

    Args:
        doc: Revit document
        aae_id: ElementId van het AppearanceAssetElement
        image_path: Pad naar bitmap
        width_m: Breedte in meters
        height_m: Hoogte in meters
    """
    try:
        from Autodesk.Revit.DB.Visual import AppearanceAssetEditScope

        scope = AppearanceAssetEditScope(doc)
        editable = scope.Start(aae_id)

        # Zoek diffuse color property
        diffuse = editable.FindByName("generic_diffuse")
        if diffuse is None:
            scope.Cancel()
            return

        # Voeg bitmap connectie toe als die er niet is
        if diffuse.NumberOfConnectedProperties == 0:
            try:
                diffuse.AddConnectedAsset("UnifiedBitmap")
            except Exception as e:
                _geom_log("AddConnectedAsset fout: {0}".format(e))
                scope.Cancel()
                return

        # Haal de bitmap asset op
        bitmap = diffuse.GetConnectedProperty(0)
        if bitmap is None:
            scope.Cancel()
            return

        # Stel bitmap pad in
        path_prop = bitmap.FindByName("unifiedbitmap_Bitmap")
        if path_prop:
            path_prop.Value = image_path

        # Stel real-world schaal in (meters -> feet)
        width_feet = width_m * METER_TO_FEET
        height_feet = height_m * METER_TO_FEET

        _set_asset_double(bitmap, "texture_RealWorldScaleX", width_feet)
        _set_asset_double(bitmap, "texture_RealWorldScaleY", height_feet)

        # Offset: verschuif zodat de textuur gecentreerd is rond origin
        offset_x_feet = -(width_m / 2.0) * METER_TO_FEET
        offset_y_feet = -(height_m / 2.0) * METER_TO_FEET
        _set_asset_double(bitmap, "texture_RealWorldOffsetX", offset_x_feet)
        _set_asset_double(bitmap, "texture_RealWorldOffsetY", offset_y_feet)

        scope.Commit(True)

    except Exception as e:
        _geom_log("_set_bitmap_on_asset fout: {0}".format(e))
        try:
            scope.Cancel()
        except Exception:
            pass


def _update_bitmap_texture(doc, aae_id, image_path, width_m, height_m):
    """Update de bitmap textuur van een bestaand AppearanceAssetElement."""
    try:
        from Autodesk.Revit.DB.Visual import AppearanceAssetEditScope

        scope = AppearanceAssetEditScope(doc)
        editable = scope.Start(aae_id)

        diffuse = editable.FindByName("generic_diffuse")
        if diffuse and diffuse.NumberOfConnectedProperties > 0:
            bitmap = diffuse.GetConnectedProperty(0)
            if bitmap:
                path_prop = bitmap.FindByName("unifiedbitmap_Bitmap")
                if path_prop:
                    path_prop.Value = image_path

                width_feet = width_m * METER_TO_FEET
                height_feet = height_m * METER_TO_FEET
                _set_asset_double(bitmap, "texture_RealWorldScaleX", width_feet)
                _set_asset_double(bitmap, "texture_RealWorldScaleY", height_feet)

                offset_x_feet = -(width_m / 2.0) * METER_TO_FEET
                offset_y_feet = -(height_m / 2.0) * METER_TO_FEET
                _set_asset_double(
                    bitmap, "texture_RealWorldOffsetX", offset_x_feet)
                _set_asset_double(
                    bitmap, "texture_RealWorldOffsetY", offset_y_feet)

        scope.Commit(True)

    except Exception as e:
        _geom_log("_update_bitmap_texture fout: {0}".format(e))
        try:
            scope.Cancel()
        except Exception:
            pass


def _set_asset_double(asset, prop_name, value):
    """Stel een double property in op een asset (robuust).

    Probeert verschillende property types (Double, Distance, Float).
    """
    prop = asset.FindByName(prop_name)
    if prop is None:
        return

    try:
        prop.Value = value
    except Exception:
        try:
            # Sommige properties zijn Distance in plaats van Double
            from Autodesk.Revit.DB.Visual import AssetPropertyDouble
            if isinstance(prop, AssetPropertyDouble):
                prop.Value = value
        except Exception:
            pass


# =============================================================================
# Project Parameters (voor BRO data op DirectShapes)
# =============================================================================

_bro_params_created = False


def ensure_bro_parameters(doc):
    """
    Zorg dat BRO parameters bestaan op GenericModel categorie.

    Maakt de volgende instance parameters aan als ze nog niet bestaan:
    - CPT_diepte (length): Diepte van CPT sondering in meters
    - boring_diepte (length): Totale diepte van boring in meters
    - boring_materiaal (text): Grondsoort van de laag

    Args:
        doc: Revit document

    Returns:
        True als succesvol
    """
    global _bro_params_created
    if not IN_REVIT:
        return False

    if _bro_params_created:
        return True

    from Autodesk.Revit.DB import ExternalDefinitionCreationOptions

    app = doc.Application

    # Parameter definities: (naam, type_key)
    param_defs = [
        ("CPT_diepte", "length"),
        ("boring_diepte", "length"),
        ("boring_materiaal", "text"),
        ("CPT_qc", "number"),
        ("CPT_classificatie", "text"),
        ("BRO_url", "text"),
    ]

    # Check welke al bestaan
    existing = set()
    bm = doc.ParameterBindings
    it = bm.ForwardIterator()
    while it.MoveNext():
        existing.add(it.Key.Name)

    needed = [(n, t) for n, t in param_defs if n not in existing]
    if not needed:
        _bro_params_created = True
        return True

    # Bewaar origineel shared parameter file
    original_spf = ""
    try:
        original_spf = app.SharedParametersFilename
    except Exception:
        pass

    try:
        temp_dir = os.environ.get(
            "TEMP", os.path.join(
                os.path.expanduser("~"), "AppData", "Local", "Temp"))
        temp_path = os.path.join(temp_dir, "GIS2BIM_SharedParams.txt")

        if not os.path.exists(temp_path):
            with open(temp_path, 'w') as f:
                pass

        app.SharedParametersFilename = temp_path
        def_file = app.OpenSharedParameterFile()

        # Zoek of maak groep
        group = None
        for g in def_file.Groups:
            if g.Name == "GIS2BIM":
                group = g
                break
        if group is None:
            group = def_file.Groups.Create("GIS2BIM")

        # Categorie set: GenericModel
        cat_set = app.Create.NewCategorySet()
        gm_cat = doc.Settings.Categories.get_Item(
            BuiltInCategory.OST_GenericModel)
        cat_set.Insert(gm_cat)

        for param_name, type_key in needed:
            # Zoek bestaande definitie in groep
            definition = None
            for d in group.Definitions:
                if d.Name == param_name:
                    definition = d
                    break

            if definition is None:
                opts = _create_param_options(param_name, type_key)
                definition = group.Definitions.Create(opts)

            # Bind als instance parameter
            binding = app.Create.NewInstanceBinding(cat_set)
            _bind_parameter(doc, bm, definition, binding)

        # Regenerate zodat LookupParameter werkt op nieuwe elementen
        doc.Regenerate()
        _bro_params_created = True
        return True

    except Exception as e:
        _geom_log("ensure_bro_parameters fout: {0}".format(e))
        return False
    finally:
        try:
            if original_spf:
                app.SharedParametersFilename = original_spf
        except Exception:
            pass


def _create_param_options(name, type_key):
    """Maak ExternalDefinitionCreationOptions (version-safe)."""
    from Autodesk.Revit.DB import ExternalDefinitionCreationOptions

    try:
        # Revit 2023+: ForgeTypeId
        from Autodesk.Revit.DB import SpecTypeId
        if type_key == "length":
            return ExternalDefinitionCreationOptions(
                name, SpecTypeId.Length)
        elif type_key == "number":
            return ExternalDefinitionCreationOptions(
                name, SpecTypeId.Number)
        else:
            return ExternalDefinitionCreationOptions(
                name, SpecTypeId.String.Text)
    except (ImportError, AttributeError):
        pass

    # Revit 2022 en eerder: ParameterType
    from Autodesk.Revit.DB import ParameterType as RevitPT
    if type_key == "length":
        return ExternalDefinitionCreationOptions(name, RevitPT.Length)
    elif type_key == "number":
        return ExternalDefinitionCreationOptions(name, RevitPT.Number)
    else:
        return ExternalDefinitionCreationOptions(name, RevitPT.Text)


def _bind_parameter(doc, binding_map, definition, binding):
    """Bind parameter aan document (version-safe)."""
    try:
        from Autodesk.Revit.DB import GroupTypeId
        binding_map.Insert(definition, binding, GroupTypeId.Data)
        return
    except (ImportError, AttributeError):
        pass

    try:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        binding_map.Insert(
            definition, binding, BuiltInParameterGroup.PG_DATA)
        return
    except Exception:
        pass

    binding_map.Insert(definition, binding)


def set_element_parameter(element, param_name, value):
    """
    Stel een parameter waarde in op een element.

    Voor Length parameters: waarde in Revit internal units (feet).
    Voor Text parameters: string waarde.

    Args:
        element: Revit element
        param_name: Naam van de parameter
        value: Waarde om in te stellen

    Returns:
        True als succesvol
    """
    if not IN_REVIT:
        return False

    try:
        from Autodesk.Revit.DB import StorageType

        param = element.LookupParameter(param_name)
        if param is None or param.IsReadOnly:
            return False

        if param.StorageType == StorageType.Double:
            param.Set(float(value))
        elif param.StorageType == StorageType.String:
            param.Set(str(value))
        elif param.StorageType == StorageType.Integer:
            param.Set(int(value))
        else:
            return False

        return True

    except Exception as e:
        _geom_log("set_element_parameter fout ({0}): {1}".format(
            param_name, e))
        return False
