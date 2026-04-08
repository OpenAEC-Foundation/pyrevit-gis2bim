# -*- coding: utf-8 -*-
"""
Heatmap Visualisatie - GebiedsAnalyse
======================================

Revit SpatialFieldManager wrapper voor het visualiseren van
gebiedsanalyse scores als een heatmap op een Floor element.

Workflow:
    1. Maak een Floor element als analyse-oppervlak
    2. Haal SpatialFieldManager voor de actieve view
    3. Registreer AnalysisResultSchema
    4. Zet UV-punten + scores op het bovenvlak (PlanarFace, normal.Z ~ 1)
    5. Stel AnalysisDisplayStyle in met kleurstops

Gebruik:
    from gis2bim.analysis.heatmap import create_and_apply_heatmap
"""

import math

IMPORT_ERRORS = []

try:
    from Autodesk.Revit.DB import (
        XYZ, UV,
        Transaction, TransactionGroup,
        FilteredElementCollector,
        Floor, FloorType, Level,
        CurveLoop, Line,
        Options, Solid,
        PlanarFace,
    )
except ImportError as e:
    IMPORT_ERRORS.append("Revit.DB: {0}".format(e))

try:
    from Autodesk.Revit.DB.Analysis import (
        SpatialFieldManager,
        AnalysisResultSchema,
        AnalysisDisplayStyle,
        AnalysisDisplayColoredSurfaceSettings,
        AnalysisDisplayColorSettings,
        AnalysisDisplayLegendSettings,
        FieldDomainPointsByUV,
        FieldValues,
        ValueAtPoint,
    )
except ImportError as e:
    IMPORT_ERRORS.append("Revit.DB.Analysis: {0}".format(e))

try:
    from System.Collections.Generic import List as GenericList
except ImportError as e:
    IMPORT_ERRORS.append("System.Collections: {0}".format(e))

IN_REVIT = len(IMPORT_ERRORS) == 0

# Conversie constanten
METER_TO_FEET = 1.0 / 0.3048


def create_floor_surface(doc, center_rd, origin_rd, size_m, level=None, log=None):
    """Maak een Floor element als analyse-oppervlak.

    Args:
        doc: Revit Document
        center_rd: Tuple (rd_x, rd_y) van het grid-centrum
        origin_rd: Tuple (rd_x, rd_y) van de project-origin
        size_m: Grootte van het grid in meters
        level: Revit Level (of None voor laagste level)
        log: Optionele log functie

    Returns:
        Floor element, of None bij fout
    """
    if log is None:
        log = lambda msg: None

    if level is None:
        levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
        if not levels:
            log("Geen levels gevonden")
            return None
        level = sorted(levels, key=lambda lv: lv.Elevation)[0]

    # Floor type ophalen
    floor_types = list(FilteredElementCollector(doc).OfClass(FloorType).ToElements())
    if not floor_types:
        log("Geen FloorType gevonden")
        return None
    floor_type = floor_types[0]

    # Bereken hoekpunten relatief aan project-origin, in feet
    ox, oy = origin_rd
    cx, cy = center_rd
    half = size_m / 2.0

    corners = [
        XYZ((cx - half - ox) * METER_TO_FEET, (cy - half - oy) * METER_TO_FEET, 0),
        XYZ((cx + half - ox) * METER_TO_FEET, (cy - half - oy) * METER_TO_FEET, 0),
        XYZ((cx + half - ox) * METER_TO_FEET, (cy + half - oy) * METER_TO_FEET, 0),
        XYZ((cx - half - ox) * METER_TO_FEET, (cy + half - oy) * METER_TO_FEET, 0),
    ]

    profile = CurveLoop()
    for i in range(4):
        line = Line.CreateBound(corners[i], corners[(i + 1) % 4])
        profile.Append(line)

    profiles = GenericList[CurveLoop]()
    profiles.Add(profile)

    try:
        floor = Floor.Create(doc, profiles, floor_type.Id, level.Id)
    except Exception as e:
        log("Fout bij Floor.Create: {0}".format(e))
        return None

    log("Floor aangemaakt: {0}x{0}m".format(int(size_m)))
    return floor


def _get_top_face(floor, log=None):
    """Haal het bovenvlak (PlanarFace met normal.Z ~ 1) van een Floor.

    Probeert eerst standaard Options, dan met IncludeNonVisibleObjects
    als fallback. Itereert recursief door alle geometry objecten inclusief
    GeometryInstance nesting.

    Args:
        floor: Revit Floor element
        log: Optionele log functie

    Returns:
        PlanarFace of None
    """
    if log is None:
        log = lambda msg: None

    def _find_top_face_in_solid(solid):
        """Zoek bovenvlak in een Solid."""
        if solid.Faces.Size == 0:
            return None
        for face in solid.Faces:
            if isinstance(face, PlanarFace):
                normal = face.FaceNormal
                if abs(normal.Z - 1.0) < 0.01:
                    return face
        return None

    def _search_geometry(geom_element):
        """Recursief zoeken door geometry hierarchy."""
        for geom_obj in geom_element:
            if isinstance(geom_obj, Solid) and geom_obj.Faces.Size > 0:
                result = _find_top_face_in_solid(geom_obj)
                if result is not None:
                    return result
            elif hasattr(geom_obj, "GetInstanceGeometry"):
                sub_geom = geom_obj.GetInstanceGeometry()
                if sub_geom is not None:
                    result = _search_geometry(sub_geom)
                    if result is not None:
                        return result
        return None

    # Poging 1: standaard Options
    opt = Options()
    opt.ComputeReferences = True
    geom = floor.get_Geometry(opt)
    if geom is not None:
        result = _search_geometry(geom)
        if result is not None:
            return result

    # Poging 2: met IncludeNonVisibleObjects (vangt gevallen waar
    # geometry na commit nog niet volledig zichtbaar is)
    log("Top face niet gevonden met standaard Options, fallback poging")
    opt2 = Options()
    opt2.ComputeReferences = True
    opt2.IncludeNonVisibleObjects = True
    geom2 = floor.get_Geometry(opt2)
    if geom2 is not None:
        result = _search_geometry(geom2)
        if result is not None:
            return result

    return None


def _setup_display_style(doc, sfm, schema_index, log=None):
    """Stel het kleurenschema in voor de heatmap.

    Kleurstops van laag (rood) naar hoog (groen):
        Min:  (232, 89, 60)  - Rood
        Max:  (29, 117, 24)  - Groen
    Revit interpoleert automatisch tussen min en max.

    Args:
        doc: Revit Document
        sfm: SpatialFieldManager
        schema_index: Index van het geregistreerde schema
        log: Optionele log functie
    """
    if log is None:
        log = lambda msg: None

    from Autodesk.Revit.DB import Color as RevitColor

    try:
        color_settings = AnalysisDisplayColorSettings()
        color_settings.MinColor = RevitColor(232, 89, 60)   # Rood (laag)
        color_settings.MaxColor = RevitColor(29, 117, 24)   # Groen (hoog)

        surface_settings = AnalysisDisplayColoredSurfaceSettings()

        legend_settings = AnalysisDisplayLegendSettings()
        legend_settings.ShowLegend = True

        # Verwijder bestaande style met dezelfde naam
        style_name = "GebiedsAnalyse Heatmap"
        collector = FilteredElementCollector(doc)
        existing = collector.OfClass(AnalysisDisplayStyle).ToElements()

        for existing_style in existing:
            try:
                if existing_style.Name == style_name:
                    doc.Delete(existing_style.Id)
                    break
            except Exception:
                pass

        new_style = AnalysisDisplayStyle.CreateAnalysisDisplayStyle(
            doc, style_name, surface_settings, color_settings, legend_settings
        )

        view = doc.ActiveView
        view.AnalysisDisplayStyleId = new_style.Id

    except Exception as e:
        log("Display style fout (niet kritisch): {0}".format(e))


def apply_heatmap(doc, floor, grid_points, scores, grid_size_m, n_rows, n_cols,
                  origin_rd, center_rd, log=None):
    """Pas heatmap scores toe op een Floor via SpatialFieldManager.

    Args:
        doc: Revit Document
        floor: Floor element (analyse-oppervlak)
        grid_points: Lijst van grid dicts met 'lat', 'lon', 'row', 'col'
        scores: Lijst van floats (score per gridpunt)
        grid_size_m: Grootte van het grid in meters
        n_rows: Aantal rijen in het grid
        n_cols: Aantal kolommen in het grid
        origin_rd: Tuple (rd_x, rd_y) van project origin
        center_rd: Tuple (rd_x, rd_y) van grid centrum
        log: Optionele log functie

    Returns:
        True bij succes, False bij fout
    """
    if log is None:
        log = lambda msg: None

    view = doc.ActiveView

    # Haal bovenvlak van de floor
    top_face = _get_top_face(floor, log=log)
    if top_face is None:
        log("Kon bovenvlak van Floor niet vinden")
        return False

    face_ref = top_face.Reference
    if face_ref is None:
        log("top_face.Reference is None")
        return False

    # SpatialFieldManager ophalen of aanmaken
    sfm = SpatialFieldManager.GetSpatialFieldManager(view)
    if sfm is None:
        sfm = SpatialFieldManager.CreateSpatialFieldManager(view, 1)

    # Schema registreren
    schema_name = "GebiedsAnalyse Score"
    schema_index = -1

    existing_schemas = sfm.GetRegisteredResults()
    for sid in existing_schemas:
        schema = sfm.GetResultSchema(sid)
        if schema.Name == schema_name:
            schema_index = sid
            break

    if schema_index == -1:
        schema = AnalysisResultSchema(schema_name, "Voorzieningenscore per gridpunt")
        schema_index = sfm.RegisterResult(schema)

    # UV-punten op het face zetten
    ox, oy = origin_rd
    cx, cy = center_rd
    half = grid_size_m / 2.0

    from System import Double as SysDouble

    uv_points = GenericList[UV]()
    score_values = []

    for i, point in enumerate(grid_points):
        row = point["row"]
        col = point["col"]

        dx_m = -half + col * (grid_size_m / float(n_cols - 1)) if n_cols > 1 else 0
        dy_m = -half + row * (grid_size_m / float(n_rows - 1)) if n_rows > 1 else 0

        x_ft = (cx + dx_m - ox) * METER_TO_FEET
        y_ft = (cy + dy_m - oy) * METER_TO_FEET

        try:
            result = top_face.Project(XYZ(x_ft, y_ft, top_face.Origin.Z))
            if result is not None:
                uv_points.Add(result.UVPoint)
                score_values.append(float(scores[i]))
        except Exception:
            pass

    if uv_points.Count == 0:
        log("Geen UV-punten konden worden geprojecteerd op het face")
        return False

    log("{0} van {1} punten op face geplaatst".format(uv_points.Count, len(grid_points)))

    # Maak FieldDomainPointsByUV en FieldValues via ValueAtPoint
    domain_points = FieldDomainPointsByUV(uv_points)

    val_at_points = GenericList[ValueAtPoint]()
    for score in score_values:
        dbl_list = GenericList[SysDouble]()
        dbl_list.Add(SysDouble(score))
        val_at_points.Add(ValueAtPoint(dbl_list))
    field_values = FieldValues(val_at_points)

    # Update het face met de waarden
    primitive_id = sfm.AddSpatialFieldPrimitive(face_ref)
    sfm.UpdateSpatialFieldPrimitive(
        primitive_id, domain_points, field_values, schema_index)

    # Display style instellen
    _setup_display_style(doc, sfm, schema_index, log=log)

    log("Heatmap toegepast: {0} punten".format(uv_points.Count))
    return True


def create_and_apply_heatmap(doc, grid_points, scores, grid_size_m, n_rows, n_cols,
                             origin_rd, center_rd, log=None):
    """Convenience functie: maak floor + pas heatmap toe in een transactie.

    Args:
        doc: Revit Document
        grid_points: Lijst van grid dicts
        scores: Lijst van score floats
        grid_size_m: Grid grootte in meters
        n_rows, n_cols: Grid dimensies
        origin_rd: Tuple (rd_x, rd_y) project origin
        center_rd: Tuple (rd_x, rd_y) grid centrum
        log: Optionele log functie

    Returns:
        Dict met resultaat info, of None bij fout
    """
    if log is None:
        log = lambda msg: None

    if IMPORT_ERRORS:
        log("Heatmap import errors: {0}".format("; ".join(IMPORT_ERRORS)))
        return None

    tg = TransactionGroup(doc, "GebiedsAnalyse Heatmap")
    tg.Start()

    try:
        # Stap 1: Floor aanmaken
        t1 = Transaction(doc, "GebiedsAnalyse - Floor aanmaken")
        t1.Start()

        floor = create_floor_surface(doc, center_rd, origin_rd, grid_size_m, log=log)
        if floor is None:
            t1.RollBack()
            tg.RollBack()
            return None

        t1.Commit()
        # t1.Commit() regenereert automatisch

        # Stap 2: Heatmap toepassen
        t2 = Transaction(doc, "GebiedsAnalyse - Heatmap toepassen")
        t2.Start()

        success = apply_heatmap(
            doc, floor, grid_points, scores,
            grid_size_m, n_rows, n_cols,
            origin_rd, center_rd, log=log
        )

        if not success:
            t2.RollBack()
            tg.RollBack()
            return None

        t2.Commit()
        tg.Assimilate()

        # Statistieken
        valid_scores = [s for s in scores if s > 0]
        return {
            "floor_id": floor.Id.IntegerValue,
            "point_count": len(grid_points),
            "avg_score": sum(scores) / len(scores) if scores else 0,
            "max_score": max(scores) if scores else 0,
            "min_score": min(scores) if scores else 0,
            "nonzero_count": len(valid_scores),
        }

    except Exception as e:
        log("Fout bij heatmap: {0}".format(e))
        try:
            tg.RollBack()
        except Exception:
            pass
        return None
