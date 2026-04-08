# -*- coding: UTF-8 -*-
"""
Modification Module for Revit MCP
Batch updates, sheet creation, view placement, wall creation, family placement.
"""

from pyrevit import routes, revit, DB
from Autodesk.Revit.DB import Transaction
import logging
import json

logger = logging.getLogger(__name__)


def safe_str(value):
    """Safely convert value to ASCII-safe string for JSON serialization."""
    if value is None:
        return None
    try:
        return str(value).encode('ascii', 'replace').decode('ascii')
    except:
        return str(type(value))


def get_element_name(elem):
    """Get element name safely."""
    try:
        if hasattr(elem, 'Name') and elem.Name:
            return elem.Name
        return "Unnamed"
    except:
        return "Unknown"


def register_modification_routes(api):
    """Register modification routes with the API."""

    @api.route("/batch_update/", methods=["POST"])
    def batch_update(doc, request):
        """
        Batch update parameters on multiple elements.

        Expected data format:
        {
            "updates": [
                {"element_id": 12345, "parameters": {"Mark": "A-101", "Comments": "Updated"}},
                {"element_id": 12346, "parameters": {"Mark": "A-102"}}
            ]
        }
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            updates = data.get("updates", [])

            if not updates:
                return routes.make_response(data={"error": "updates list required"}, status=400)

            results = {"success": [], "failed": []}

            with Transaction(doc, "MCP Batch Update") as t:
                t.Start()
                for update in updates:
                    elem_id = update.get("element_id")
                    params = update.get("parameters", {})

                    if not elem_id:
                        results["failed"].append({"element_id": None, "error": "Missing element_id"})
                        continue

                    element = doc.GetElement(DB.ElementId(int(elem_id)))
                    if not element:
                        results["failed"].append({"element_id": elem_id, "error": "Element not found"})
                        continue

                    # Get type for type parameters
                    elem_type = None
                    type_id = element.GetTypeId()
                    if type_id and type_id.IntegerValue > 0:
                        elem_type = doc.GetElement(type_id)

                    elem_success = []
                    elem_failed = []

                    for param_name, new_value in params.items():
                        try:
                            param = element.LookupParameter(param_name)
                            if not param and elem_type:
                                param = elem_type.LookupParameter(param_name)

                            if not param:
                                elem_failed.append({"parameter": param_name, "error": "Not found"})
                                continue

                            if param.IsReadOnly:
                                elem_failed.append({"parameter": param_name, "error": "Read-only"})
                                continue

                            success = False
                            if param.StorageType == DB.StorageType.String:
                                success = param.Set(str(new_value) if new_value is not None else "")
                            elif param.StorageType == DB.StorageType.Integer:
                                if isinstance(new_value, bool):
                                    success = param.Set(1 if new_value else 0)
                                elif str(new_value).lower() in ("yes", "true", "1"):
                                    success = param.Set(1)
                                elif str(new_value).lower() in ("no", "false", "0"):
                                    success = param.Set(0)
                                else:
                                    success = param.Set(int(new_value))
                            elif param.StorageType == DB.StorageType.Double:
                                success = param.Set(float(new_value))
                            elif param.StorageType == DB.StorageType.ElementId:
                                success = param.Set(DB.ElementId(int(new_value)))

                            if success:
                                elem_success.append(param_name)
                            else:
                                elem_failed.append({"parameter": param_name, "error": "Set failed"})
                        except Exception as e:
                            elem_failed.append({"parameter": param_name, "error": str(e)})

                    if elem_failed:
                        results["failed"].append({
                            "element_id": elem_id,
                            "success": elem_success,
                            "failed": elem_failed
                        })
                    else:
                        results["success"].append({
                            "element_id": elem_id,
                            "parameters_updated": elem_success
                        })
                t.Commit()

            return routes.make_response(data={
                "status": "success",
                "total_elements": len(updates),
                "success_count": len(results["success"]),
                "failed_count": len(results["failed"]),
                "results": results
            })

        except Exception as e:
            logger.error("Batch update failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/create_sheet/", methods=["POST"])
    def create_sheet(doc, request):
        """
        Create a new sheet.

        Expected data:
        {
            "sheet_number": "A-101",
            "sheet_name": "Floor Plan - Level 1",
            "title_block_name": "A1 metric" (optional - uses first available if not specified)
        }
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            sheet_number = data.get("sheet_number")
            sheet_name = data.get("sheet_name", "New Sheet")
            title_block_name = data.get("title_block_name")

            if not sheet_number:
                return routes.make_response(data={"error": "sheet_number required"}, status=400)

            # Find title block
            title_block_id = DB.ElementId.InvalidElementId
            title_blocks = DB.FilteredElementCollector(doc).OfCategory(
                DB.BuiltInCategory.OST_TitleBlocks
            ).WhereElementIsElementType().ToElements()

            if title_block_name:
                for tb in title_blocks:
                    if title_block_name.lower() in get_element_name(tb).lower():
                        title_block_id = tb.Id
                        break
                if title_block_id == DB.ElementId.InvalidElementId:
                    return routes.make_response(
                        data={"error": "Title block '{}' not found".format(title_block_name)},
                        status=404
                    )
            elif title_blocks:
                title_block_id = title_blocks[0].Id

            with Transaction(doc, "MCP Create Sheet") as t:
                t.Start()
                new_sheet = DB.ViewSheet.Create(doc, title_block_id)
                new_sheet.SheetNumber = sheet_number
                new_sheet.Name = sheet_name
                t.Commit()

            return routes.make_response(data={
                "status": "success",
                "sheet_id": new_sheet.Id.IntegerValue,
                "sheet_number": sheet_number,
                "sheet_name": sheet_name
            })

        except Exception as e:
            logger.error("Create sheet failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/place_view_on_sheet/", methods=["POST"])
    def place_view_on_sheet(doc, request):
        """
        Place a view on a sheet.

        Expected data:
        {
            "sheet_id": 12345,
            "view_id": 67890,
            "x": 0.5,  (position in feet from sheet origin, optional)
            "y": 0.5   (position in feet from sheet origin, optional)
        }
        OR use names:
        {
            "sheet_number": "A-101",
            "view_name": "Level 1"
        }
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data

            # Get sheet
            sheet = None
            sheet_id = data.get("sheet_id")
            sheet_number = data.get("sheet_number")

            if sheet_id:
                sheet = doc.GetElement(DB.ElementId(int(sheet_id)))
            elif sheet_number:
                sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()
                for s in sheets:
                    if s.SheetNumber == sheet_number:
                        sheet = s
                        break

            if not sheet:
                return routes.make_response(data={"error": "Sheet not found"}, status=404)

            # Get view
            view = None
            view_id = data.get("view_id")
            view_name = data.get("view_name")

            if view_id:
                view = doc.GetElement(DB.ElementId(int(view_id)))
            elif view_name:
                views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
                for v in views:
                    if hasattr(v, 'Name') and v.Name == view_name:
                        view = v
                        break

            if not view:
                return routes.make_response(data={"error": "View not found"}, status=404)

            # Check if view can be placed
            if not DB.Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
                return routes.make_response(
                    data={"error": "View cannot be placed on sheet (already placed or is a template)"},
                    status=400
                )

            # Position (default to center-ish of sheet)
            x = data.get("x", 1.0)  # feet
            y = data.get("y", 0.75)  # feet
            location = DB.XYZ(float(x), float(y), 0)

            with Transaction(doc, "MCP Place View on Sheet") as t:
                t.Start()
                viewport = DB.Viewport.Create(doc, sheet.Id, view.Id, location)
                t.Commit()

            return routes.make_response(data={
                "status": "success",
                "viewport_id": viewport.Id.IntegerValue,
                "sheet_id": sheet.Id.IntegerValue,
                "sheet_number": sheet.SheetNumber,
                "view_id": view.Id.IntegerValue,
                "view_name": safe_str(view.Name)
            })

        except Exception as e:
            logger.error("Place view on sheet failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/create_walls_at_lines/", methods=["POST"])
    def create_walls_at_lines(doc, request):
        """
        Create walls at model line locations.

        Expected data:
        {
            "line_ids": [12345, 12346],  (element IDs of model lines/detail lines)
            "wall_type_name": "Generic - 200mm",  (optional, uses first available)
            "level_name": "Level 1",  (optional, uses first level)
            "height": 3000,  (wall height in mm, default 3000)
            "structural": false  (optional)
        }
        OR create walls from coordinates:
        {
            "lines": [
                {"start": {"x": 0, "y": 0}, "end": {"x": 5000, "y": 0}},
                {"start": {"x": 5000, "y": 0}, "end": {"x": 5000, "y": 3000}}
            ],
            "level_name": "Level 1",
            "height": 3000
        }
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data

            line_ids = data.get("line_ids", [])
            lines_data = data.get("lines", [])
            wall_type_name = data.get("wall_type_name")
            level_name = data.get("level_name")
            height_mm = data.get("height", 3000)
            structural = data.get("structural", False)

            # Convert height to feet
            height_feet = height_mm / 304.8

            # Find wall type
            wall_type = None
            wall_types = DB.FilteredElementCollector(doc).OfClass(DB.WallType).ToElements()

            if wall_type_name:
                for wt in wall_types:
                    if wall_type_name.lower() in get_element_name(wt).lower():
                        wall_type = wt
                        break
            if not wall_type and wall_types:
                # Use first basic wall type
                for wt in wall_types:
                    if wt.Kind == DB.WallKind.Basic:
                        wall_type = wt
                        break
                if not wall_type:
                    wall_type = wall_types[0]

            if not wall_type:
                return routes.make_response(data={"error": "No wall types available"}, status=404)

            # Find level
            level = None
            levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()

            if level_name:
                for lvl in levels:
                    if level_name.lower() in get_element_name(lvl).lower():
                        level = lvl
                        break
            if not level and levels:
                level = sorted(levels, key=lambda l: l.Elevation)[0]

            if not level:
                return routes.make_response(data={"error": "No levels available"}, status=404)

            # Collect lines to create walls from
            wall_lines = []

            # From existing line elements
            for line_id in line_ids:
                elem = doc.GetElement(DB.ElementId(int(line_id)))
                if elem:
                    try:
                        if hasattr(elem, 'GeometryCurve'):
                            wall_lines.append(elem.GeometryCurve)
                        elif hasattr(elem, 'Location') and hasattr(elem.Location, 'Curve'):
                            wall_lines.append(elem.Location.Curve)
                    except:
                        pass

            # From coordinate data
            for line_data in lines_data:
                try:
                    start = line_data.get("start", {})
                    end = line_data.get("end", {})
                    # Convert mm to feet
                    start_pt = DB.XYZ(
                        start.get("x", 0) / 304.8,
                        start.get("y", 0) / 304.8,
                        start.get("z", 0) / 304.8
                    )
                    end_pt = DB.XYZ(
                        end.get("x", 0) / 304.8,
                        end.get("y", 0) / 304.8,
                        end.get("z", 0) / 304.8
                    )
                    line = DB.Line.CreateBound(start_pt, end_pt)
                    wall_lines.append(line)
                except Exception as e:
                    logger.warning("Could not create line: {}".format(str(e)))

            if not wall_lines:
                return routes.make_response(
                    data={"error": "No valid lines found to create walls"},
                    status=400
                )

            results = {"created": [], "failed": []}

            with Transaction(doc, "MCP Create Walls at Lines") as t:
                t.Start()
                for i, line in enumerate(wall_lines):
                    try:
                        wall = DB.Wall.Create(
                            doc,
                            line,
                            wall_type.Id,
                            level.Id,
                            height_feet,
                            0,  # offset
                            False,  # flip
                            structural
                        )
                        results["created"].append({
                            "wall_id": wall.Id.IntegerValue,
                            "line_index": i
                        })
                    except Exception as e:
                        results["failed"].append({
                            "line_index": i,
                            "error": str(e)
                        })
                t.Commit()

            return routes.make_response(data={
                "status": "success",
                "walls_created": len(results["created"]),
                "walls_failed": len(results["failed"]),
                "wall_type": safe_str(get_element_name(wall_type)),
                "level": safe_str(get_element_name(level)),
                "height_mm": height_mm,
                "results": results
            })

        except Exception as e:
            logger.error("Create walls failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/batch_family_placement/", methods=["POST"])
    def batch_family_placement(doc, request):
        """
        Place multiple family instances at once.

        Expected data:
        {
            "family_name": "Chair",
            "type_name": "Standard",  (optional)
            "placements": [
                {"x": 1000, "y": 2000, "z": 0, "rotation": 0},
                {"x": 3000, "y": 2000, "z": 0, "rotation": 90},
                {"x": 5000, "y": 2000, "z": 0}
            ],
            "level_name": "Level 1"  (optional)
        }
        Coordinates in mm, rotation in degrees.
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data

            family_name = data.get("family_name")
            type_name = data.get("type_name")
            placements = data.get("placements", [])
            level_name = data.get("level_name")

            if not family_name:
                return routes.make_response(data={"error": "family_name required"}, status=400)

            if not placements:
                return routes.make_response(data={"error": "placements list required"}, status=400)

            # Find family symbol
            family_symbol = None
            symbols = DB.FilteredElementCollector(doc).OfClass(
                DB.FamilySymbol
            ).WhereElementIsElementType().ToElements()

            for sym in symbols:
                sym_name = get_element_name(sym)
                fam_name = sym.Family.Name if sym.Family else ""

                if family_name.lower() in fam_name.lower():
                    if type_name:
                        if type_name.lower() in sym_name.lower():
                            family_symbol = sym
                            break
                    else:
                        family_symbol = sym
                        break

            if not family_symbol:
                return routes.make_response(
                    data={"error": "Family '{}' not found".format(family_name)},
                    status=404
                )

            # Find level
            level = None
            if level_name:
                levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
                for lvl in levels:
                    if level_name.lower() in get_element_name(lvl).lower():
                        level = lvl
                        break

            results = {"placed": [], "failed": []}

            with Transaction(doc, "MCP Batch Family Placement") as t:
                t.Start()

                # Activate symbol if needed
                if not family_symbol.IsActive:
                    family_symbol.Activate()
                    doc.Regenerate()

                for i, placement in enumerate(placements):
                    try:
                        # Get coordinates in feet
                        x = placement.get("x", 0) / 304.8
                        y = placement.get("y", 0) / 304.8
                        z = placement.get("z", 0) / 304.8
                        rotation_deg = placement.get("rotation", 0)

                        location = DB.XYZ(x, y, z)

                        # Create instance
                        if level:
                            instance = doc.Create.NewFamilyInstance(
                                location,
                                family_symbol,
                                level,
                                DB.Structure.StructuralType.NonStructural
                            )
                        else:
                            instance = doc.Create.NewFamilyInstance(
                                location,
                                family_symbol,
                                DB.Structure.StructuralType.NonStructural
                            )

                        # Apply rotation if specified
                        if rotation_deg != 0:
                            import math
                            rotation_rad = math.radians(rotation_deg)
                            axis = DB.Line.CreateBound(location, DB.XYZ(x, y, z + 1))
                            DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, rotation_rad)

                        results["placed"].append({
                            "instance_id": instance.Id.IntegerValue,
                            "placement_index": i,
                            "location_mm": {"x": placement.get("x", 0), "y": placement.get("y", 0)}
                        })

                    except Exception as e:
                        results["failed"].append({
                            "placement_index": i,
                            "error": str(e)
                        })

                t.Commit()

            return routes.make_response(data={
                "status": "success",
                "family": safe_str(family_symbol.Family.Name),
                "type": safe_str(get_element_name(family_symbol)),
                "placed_count": len(results["placed"]),
                "failed_count": len(results["failed"]),
                "results": results
            })

        except Exception as e:
            logger.error("Batch family placement failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/coordinate_converter/", methods=["POST"])
    def coordinate_converter(doc, request):
        """
        Convert coordinates between different systems and units.

        Expected data:
        {
            "coordinates": {"x": 1000, "y": 2000, "z": 0},
            "from_unit": "mm",  (mm, m, ft, in)
            "to_unit": "ft",
            "from_system": "internal",  (internal, project, shared, survey)
            "to_system": "project"
        }
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data

            coords = data.get("coordinates", {})
            from_unit = data.get("from_unit", "mm")
            to_unit = data.get("to_unit", "mm")
            from_system = data.get("from_system", "internal")
            to_system = data.get("to_system", "internal")

            x = coords.get("x", 0)
            y = coords.get("y", 0)
            z = coords.get("z", 0)

            # Unit conversion factors to internal feet
            to_feet = {
                "mm": 1 / 304.8,
                "m": 1 / 0.3048,
                "ft": 1,
                "in": 1 / 12
            }

            from_feet = {
                "mm": 304.8,
                "m": 0.3048,
                "ft": 1,
                "in": 12
            }

            # Convert to internal feet first
            x_ft = x * to_feet.get(from_unit, 1)
            y_ft = y * to_feet.get(from_unit, 1)
            z_ft = z * to_feet.get(from_unit, 1)

            point = DB.XYZ(x_ft, y_ft, z_ft)

            # Coordinate system transformations
            if from_system != to_system:
                try:
                    # Get project base point and survey point info
                    project_location = doc.ActiveProjectLocation
                    position = project_location.GetProjectPosition(DB.XYZ.Zero)

                    if from_system == "internal" and to_system == "shared":
                        # Transform from internal to shared
                        transform = project_location.GetTotalTransform()
                        point = transform.OfPoint(point)
                    elif from_system == "shared" and to_system == "internal":
                        # Transform from shared to internal
                        transform = project_location.GetTotalTransform().Inverse
                        point = transform.OfPoint(point)
                    elif from_system == "internal" and to_system == "project":
                        # Add project base point offset
                        point = DB.XYZ(
                            point.X + position.EastWest,
                            point.Y + position.NorthSouth,
                            point.Z + position.Elevation
                        )
                    elif from_system == "project" and to_system == "internal":
                        # Subtract project base point offset
                        point = DB.XYZ(
                            point.X - position.EastWest,
                            point.Y - position.NorthSouth,
                            point.Z - position.Elevation
                        )
                except Exception as e:
                    logger.warning("Coordinate transform failed: {}".format(str(e)))

            # Convert to output units
            result_x = point.X * from_feet.get(to_unit, 1)
            result_y = point.Y * from_feet.get(to_unit, 1)
            result_z = point.Z * from_feet.get(to_unit, 1)

            return routes.make_response(data={
                "status": "success",
                "input": {
                    "coordinates": coords,
                    "unit": from_unit,
                    "system": from_system
                },
                "output": {
                    "coordinates": {
                        "x": round(result_x, 6),
                        "y": round(result_y, 6),
                        "z": round(result_z, 6)
                    },
                    "unit": to_unit,
                    "system": to_system
                }
            })

        except Exception as e:
            logger.error("Coordinate conversion failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/place_kozijnen/", methods=["POST"])
    def place_kozijnen(doc, request):
        """
        Place WorkPlaneBased window families with correct orientation.
        
        Expected data:
        {
            "placements": [
                {
                    "symbol_id": 6714302,
                    "x_mm": 1000, "y_mm": 2000, "z_mm": 0,
                    "normal_x": 0.5, "normal_y": -0.5,
                    "mark": "K02-02"
                }
            ]
        }
        """
        import math
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            placements = data.get("placements", [])
            
            if not placements:
                return routes.make_response(data={"error": "placements list required"}, status=400)
            
            results = {"placed": [], "failed": []}
            
            with Transaction(doc, "MCP Place Kozijnen") as t:
                t.Start()
                
                for i, p in enumerate(placements):
                    try:
                        symbol_id = p.get("symbol_id")
                        symbol = doc.GetElement(DB.ElementId(int(symbol_id)))
                        
                        if not symbol:
                            results["failed"].append({"index": i, "error": "Symbol not found"})
                            continue
                        
                        if not symbol.IsActive:
                            symbol.Activate()
                            doc.Regenerate()
                        
                        # Convert mm to feet
                        x_ft = p.get("x_mm", 0) / 304.8
                        y_ft = p.get("y_mm", 0) / 304.8
                        z_ft = p.get("z_mm", 0) / 304.8
                        nx = p.get("normal_x", 1.0)
                        ny = p.get("normal_y", 0.0)
                        
                        point = DB.XYZ(x_ft, y_ft, z_ft)
                        normal = DB.XYZ(nx, ny, 0)
                        
                        # Create vertical sketch plane with EXPLICIT basis vectors
                        # This fixes the issue where normals along Y-axis cause horizontal orientation
                        # XVec: horizontal, perpendicular to normal (in XY plane)
                        xvec = DB.XYZ(-ny, nx, 0).Normalize()
                        # YVec: always vertical (up)
                        yvec = DB.XYZ(0, 0, 1)
                        
                        plane = DB.Plane.CreateByOriginAndBasis(point, xvec, yvec)
                        sketch_plane = DB.SketchPlane.Create(doc, plane)
                        
                        # Place WorkPlaneBased family
                        instance = doc.Create.NewFamilyInstance(
                            point, symbol, sketch_plane,
                            DB.Structure.StructuralType.NonStructural
                        )
                        
                        # Check facing en flip indien nodig
                        current_facing = instance.FacingOrientation
                        dot = current_facing.X * nx + current_facing.Y * ny
                        if dot < 0:
                            instance.flipFacing()
                        
                        # Set Mark parameter
                        mark = p.get("mark", "")
                        if mark:
                            mark_param = instance.LookupParameter("Mark")
                            if mark_param:
                                mark_param.Set(str(mark))
                        
                        results["placed"].append({
                            "index": i,
                            "instance_id": instance.Id.IntegerValue,
                            "mark": mark
                        })
                        
                    except Exception as e:
                        results["failed"].append({"index": i, "error": str(e)})
                
                t.Commit()
            
            return routes.make_response(data={
                "status": "success",
                "placed_count": len(results["placed"]),
                "failed_count": len(results["failed"]),
                "results": results
            })
            
        except Exception as e:
            logger.error("Place kozijnen failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/create_model_lines/", methods=["POST"])
    def create_model_lines(doc, request):
        """
        Create model lines (crosses/markers) at specified positions.
        
        Expected data:
        {
            "positions": [
                {"x": 1000, "y": 2000, "z": 0, "normal_x": 0.5, "normal_y": 0.5},
                {"x": 3000, "y": 4000, "z": 100}
            ],
            "line_length": 500  (optional, default 500mm)
        }
        Coordinates in mm. If normal provided, creates directional cross.
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            
            positions = data.get("positions", [])
            line_length = data.get("line_length", 500)
            
            if not positions:
                return routes.make_response(data={"error": "positions list required"}, status=400)
            
            half_len_ft = (line_length / 2.0) / 304.8
            results = {"created": [], "failed": []}
            
            with Transaction(doc, "MCP Create Model Lines") as t:
                t.Start()
                
                for i, pos in enumerate(positions):
                    try:
                        x_ft = pos.get("x", 0) / 304.8
                        y_ft = pos.get("y", 0) / 304.8
                        z_ft = pos.get("z", 0) / 304.8
                        nx = pos.get("normal_x", 1.0)
                        ny = pos.get("normal_y", 0.0)
                        
                        center = DB.XYZ(x_ft, y_ft, z_ft)
                        normal = DB.XYZ(nx, ny, 0).Normalize()
                        
                        # Perpendicular direction
                        up = DB.XYZ(0, 0, 1)
                        perp = normal.CrossProduct(up).Normalize()
                        
                        # Line 1: along normal
                        p1_start = DB.XYZ(center.X - normal.X * half_len_ft,
                                          center.Y - normal.Y * half_len_ft, center.Z)
                        p1_end = DB.XYZ(center.X + normal.X * half_len_ft,
                                        center.Y + normal.Y * half_len_ft, center.Z)
                        
                        # Line 2: perpendicular
                        p2_start = DB.XYZ(center.X - perp.X * half_len_ft,
                                          center.Y - perp.Y * half_len_ft, center.Z)
                        p2_end = DB.XYZ(center.X + perp.X * half_len_ft,
                                        center.Y + perp.Y * half_len_ft, center.Z)
                        
                        line1 = DB.Line.CreateBound(p1_start, p1_end)
                        line2 = DB.Line.CreateBound(p2_start, p2_end)
                        
                        # Create sketch plane
                        plane = DB.Plane.CreateByNormalAndOrigin(up, center)
                        sketch_plane = DB.SketchPlane.Create(doc, plane)
                        
                        # Create model curves
                        mc1 = doc.Create.NewModelCurve(line1, sketch_plane)
                        mc2 = doc.Create.NewModelCurve(line2, sketch_plane)
                        
                        results["created"].append({
                            "position_index": i,
                            "line_ids": [mc1.Id.IntegerValue, mc2.Id.IntegerValue]
                        })
                        
                    except Exception as e:
                        results["failed"].append({
                            "position_index": i,
                            "error": str(e)
                        })
                
                t.Commit()
            
            return routes.make_response(data={
                "status": "success",
                "created_count": len(results["created"]),
                "failed_count": len(results["failed"]),
                "results": results
            })
            
        except Exception as e:
            logger.error("Create model lines failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/fix_liggende_kozijnen/", methods=["POST"])
    def fix_liggende_kozijnen(doc, request):
        """
        Fix kozijnen that are lying flat (HandOrientation.Z != 0).
        Deletes and replaces with correct orientation.
        """
        import math
        doc = revit.doc
        try:
            # Find all windows with wrong orientation
            windows = DB.FilteredElementCollector(doc).OfCategory(
                DB.BuiltInCategory.OST_Windows
            ).WhereElementIsNotElementType().ToElements()
            
            liggende = []
            for w in windows:
                ho = w.HandOrientation
                if abs(ho.Z) > 0.5:
                    mark_param = w.LookupParameter("Mark")
                    mark = mark_param.AsString() if mark_param else ""
                    loc = w.Location.Point
                    fo = w.FacingOrientation
                    liggende.append({
                        "id": w.Id.IntegerValue,
                        "mark": mark,
                        "type_id": w.GetTypeId().IntegerValue,
                        "x_ft": loc.X,
                        "y_ft": loc.Y,
                        "z_ft": loc.Z,
                        "nx": fo.X,
                        "ny": fo.Y
                    })
            
            if not liggende:
                return routes.make_response(data={
                    "status": "success",
                    "message": "No lying windows found",
                    "fixed_count": 0
                })
            
            results = {"fixed": [], "failed": []}
            
            with Transaction(doc, "Fix Liggende Kozijnen") as t:
                t.Start()
                
                for data in liggende:
                    try:
                        # Delete old
                        doc.Delete(DB.ElementId(data["id"]))
                        
                        # Get symbol
                        symbol = doc.GetElement(DB.ElementId(data["type_id"]))
                        if not symbol.IsActive:
                            symbol.Activate()
                            doc.Regenerate()
                        
                        point = DB.XYZ(data["x_ft"], data["y_ft"], data["z_ft"])
                        nx, ny = data["nx"], data["ny"]
                        
                        # CORRECT plane with explicit basis
                        xvec = DB.XYZ(-ny, nx, 0).Normalize()
                        yvec = DB.XYZ(0, 0, 1)
                        plane = DB.Plane.CreateByOriginAndBasis(point, xvec, yvec)
                        sketch_plane = DB.SketchPlane.Create(doc, plane)
                        
                        # Place family
                        instance = doc.Create.NewFamilyInstance(
                            point, symbol, sketch_plane,
                            DB.Structure.StructuralType.NonStructural
                        )
                        
                        # Rotate if ny > 0
                        if ny > 0:
                            axis = DB.Line.CreateBound(point, DB.XYZ(point.X, point.Y, point.Z + 1))
                            DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, math.pi)
                        
                        # Set Mark
                        if data["mark"]:
                            mark_param = instance.LookupParameter("Mark")
                            if mark_param:
                                mark_param.Set(data["mark"])
                        
                        results["fixed"].append({
                            "old_id": data["id"],
                            "new_id": instance.Id.IntegerValue,
                            "mark": data["mark"]
                        })
                        
                    except Exception as e:
                        results["failed"].append({
                            "id": data["id"],
                            "mark": data["mark"],
                            "error": str(e)
                        })
                
                t.Commit()
            
            return routes.make_response(data={
                "status": "success",
                "fixed_count": len(results["fixed"]),
                "failed_count": len(results["failed"]),
                "results": results
            })
            
        except Exception as e:
            logger.error("Fix liggende kozijnen failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    logger.info("Modification routes registered successfully")
