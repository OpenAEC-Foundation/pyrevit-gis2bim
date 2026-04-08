# -*- coding: UTF-8 -*-
"""
Selection & Inspection Module for Revit MCP
Handles active selection and element inspection functionality
"""

from pyrevit import routes, revit, DB
import logging

from utils import get_element_name

logger = logging.getLogger(__name__)


def safe_str(value):
    """Safely convert value to ASCII-safe string for JSON serialization."""
    if value is None:
        return None
    try:
        return str(value).encode('ascii', 'replace').decode('ascii')
    except:
        return str(type(value))


def register_selection_routes(api):
    """Register all selection-related routes with the API"""

    @api.route("/active_selection/", methods=["GET"])
    def get_active_selection():
        """Get details about currently selected elements in Revit."""
        doc = revit.doc
        uidoc = revit.uidoc
        try:
            if not doc or not uidoc:
                return routes.make_response(
                    data={"error": "No active Revit document"}, status=503
                )

            logger.info("Getting active selection info")

            selection = uidoc.Selection
            selected_ids = selection.GetElementIds()
            
            if selected_ids.Count == 0:
                return routes.make_response(
                    data={
                        "status": "success",
                        "count": 0,
                        "message": "No elements selected",
                        "by_category": {},
                        "ids": [],
                        "first_element": None
                    }
                )

            elements_by_category = {}
            all_ids = []
            first_element_info = None
            
            for elem_id in selected_ids:
                try:
                    elem = doc.GetElement(elem_id)
                    if not elem:
                        continue
                    
                    all_ids.append(elem_id.IntegerValue)
                    
                    category_name = "Unknown"
                    if elem.Category:
                        category_name = safe_str(elem.Category.Name)

                    if category_name not in elements_by_category:
                        elements_by_category[category_name] = 0
                    elements_by_category[category_name] += 1
                    
                    if first_element_info is None:
                        first_element_info = _get_element_details(doc, elem)
                        
                except Exception as e:
                    logger.warning("Could not process element {}: {}".format(
                        elem_id.IntegerValue, str(e)
                    ))
                    continue

            result = {
                "status": "success",
                "count": len(all_ids),
                "by_category": elements_by_category,
                "ids": all_ids,
                "first_element": first_element_info
            }

            return routes.make_response(data=result)

        except Exception as e:
            logger.error("Get active selection failed: {}".format(str(e)))
            return routes.make_response(
                data={"error": "Failed to get active selection: {}".format(str(e))},
                status=500,
            )

    @api.route("/inspect_selected/<int:index>", methods=["GET"])
    def inspect_selected_element(index):
        """Get detailed information about a specific selected element by index."""
        doc = revit.doc
        uidoc = revit.uidoc
        try:
            if not doc or not uidoc:
                return routes.make_response(
                    data={"error": "No active Revit document"}, status=503
                )

            selection = uidoc.Selection
            selected_ids = list(selection.GetElementIds())
            
            if not selected_ids:
                return routes.make_response(
                    data={"error": "No elements selected"}, status=404
                )
            
            if index < 0 or index >= len(selected_ids):
                return routes.make_response(
                    data={
                        "error": "Index {} out of range. Selection has {} elements.".format(
                            index, len(selected_ids)
                        )
                    },
                    status=400
                )

            elem_id = selected_ids[index]
            elem = doc.GetElement(elem_id)
            
            if not elem:
                return routes.make_response(
                    data={"error": "Element not found"}, status=404
                )

            element_info = _get_element_details(doc, elem, include_all_parameters=True)
            
            return routes.make_response(
                data={
                    "status": "success",
                    "index": index,
                    "total_selected": len(selected_ids),
                    "element": element_info
                }
            )

        except Exception as e:
            logger.error("Inspect selected element failed: {}".format(str(e)))
            return routes.make_response(
                data={"error": "Failed to inspect element: {}".format(str(e))},
                status=500,
            )

    @api.route("/inspect_element/<int:element_id>", methods=["GET"])
    def inspect_element_by_id(element_id):
        """Get comprehensive information about ANY element by ID."""
        doc = revit.doc
        try:
            if not doc:
                return routes.make_response(
                    data={"error": "No active Revit document"}, status=503
                )

            logger.info("Inspecting element ID: {}".format(element_id))

            elem = doc.GetElement(DB.ElementId(element_id))
            
            if not elem:
                return routes.make_response(
                    data={"error": "Element {} not found".format(element_id)},
                    status=404
                )

            element_info = _get_element_details(doc, elem, include_all_parameters=True, include_geometry=True)
            
            return routes.make_response(
                data={
                    "status": "success",
                    "element": element_info
                }
            )

        except Exception as e:
            logger.error("Inspect element failed: {}".format(str(e)))
            return routes.make_response(
                data={"error": "Failed to inspect element: {}".format(str(e))},
                status=500,
            )

    @api.route("/link_status/", methods=["GET"])
    def get_link_status():
        """Get comprehensive information about all linked models."""
        doc = revit.doc
        try:
            if not doc:
                return routes.make_response(
                    data={"error": "No active Revit document"}, status=503
                )

            logger.info("Getting link status")

            links_info = {}
            
            # Get all RevitLinkInstances
            collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
            
            for link_instance in collector:
                try:
                    link_doc = link_instance.GetLinkDocument()
                    link_name = safe_str(get_element_name(link_instance))

                    link_data = {
                        "instance_id": link_instance.Id.IntegerValue,
                        "loaded": link_doc is not None,
                        "pinned": link_instance.Pinned if hasattr(link_instance, "Pinned") else False,
                    }
                    
                    # Get transform info
                    try:
                        transform = link_instance.GetTotalTransform()
                        link_data["transform"] = {
                            "origin_feet": {
                                "x": transform.Origin.X,
                                "y": transform.Origin.Y,
                                "z": transform.Origin.Z
                            },
                            "origin_mm": {
                                "x": round(transform.Origin.X * 304.8, 1),
                                "y": round(transform.Origin.Y * 304.8, 1),
                                "z": round(transform.Origin.Z * 304.8, 1)
                            },
                            "is_identity": transform.IsIdentity
                        }
                    except Exception:
                        link_data["transform"] = None
                    
                    if link_doc:
                        link_data["document_title"] = safe_str(link_doc.Title)
                        link_data["path"] = safe_str(link_doc.PathName)

                        # Count elements by category
                        category_counts = {}
                        try:
                            all_elements = DB.FilteredElementCollector(link_doc).WhereElementIsNotElementType().ToElements()
                            for elem in all_elements:
                                if elem.Category:
                                    cat_name = safe_str(elem.Category.Name)
                                    if cat_name not in category_counts:
                                        category_counts[cat_name] = 0
                                    category_counts[cat_name] += 1
                            link_data["element_counts"] = category_counts
                            link_data["total_elements"] = sum(category_counts.values())
                        except Exception:
                            link_data["element_counts"] = {}
                            link_data["total_elements"] = 0
                    
                    links_info[link_name] = link_data
                    
                except Exception as e:
                    logger.warning("Could not process link: {}".format(str(e)))
                    continue

            return routes.make_response(
                data={
                    "status": "success",
                    "link_count": len(links_info),
                    "links": links_info
                }
            )

        except Exception as e:
            logger.error("Get link status failed: {}".format(str(e)))
            return routes.make_response(
                data={"error": "Failed to get link status: {}".format(str(e))},
                status=500,
            )

    @api.route("/quick_count/", methods=["POST"])
    def quick_count(request):
        """Fast element count with filters."""
        doc = revit.doc
        uidoc = revit.uidoc
        try:
            if not doc:
                return routes.make_response(
                    data={"error": "No active Revit document"}, status=503
                )

            import json
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            
            category_name = data.get("category")
            if not category_name:
                return routes.make_response(
                    data={"error": "category is required"}, status=400
                )
            
            type_contains = data.get("type_contains")
            type_excludes = data.get("type_excludes", [])
            level_name = data.get("level")
            in_view = data.get("in_view", False)
            
            logger.info("Quick count for category: {}".format(category_name))

            # Find category
            target_category = None
            for cat in doc.Settings.Categories:
                if cat.Name == category_name:
                    target_category = cat
                    break
            
            if not target_category:
                return routes.make_response(
                    data={"error": "Category '{}' not found".format(category_name)},
                    status=404
                )

            # Build collector
            if in_view and uidoc and uidoc.ActiveView:
                collector = DB.FilteredElementCollector(doc, uidoc.ActiveView.Id)
            else:
                collector = DB.FilteredElementCollector(doc)
            
            collector = collector.OfCategoryId(target_category.Id).WhereElementIsNotElementType()
            elements = collector.ToElements()

            # Apply filters
            filtered_elements = []
            filters_applied = []
            
            for elem in elements:
                try:
                    # Get type name
                    type_name = ""
                    type_id = elem.GetTypeId()
                    if type_id and type_id != DB.ElementId.InvalidElementId:
                        type_elem = doc.GetElement(type_id)
                        if type_elem:
                            type_name = safe_str(get_element_name(type_elem))
                    
                    # Type contains filter
                    if type_contains and type_contains not in type_name:
                        continue
                    
                    # Type excludes filter
                    skip = False
                    if isinstance(type_excludes, list):
                        for exclude in type_excludes:
                            if exclude in type_name:
                                skip = True
                                break
                    if skip:
                        continue
                    
                    # Level filter
                    if level_name:
                        elem_level = None
                        level_param = elem.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
                        if level_param:
                            level_id = level_param.AsElementId()
                            if level_id and level_id != DB.ElementId.InvalidElementId:
                                level_elem = doc.GetElement(level_id)
                                if level_elem:
                                    elem_level = safe_str(get_element_name(level_elem))
                        if elem_level != level_name:
                            continue
                    
                    filtered_elements.append(elem.Id.IntegerValue)
                    
                except Exception:
                    continue

            if type_contains:
                filters_applied.append("type_contains: {}".format(type_contains))
            if type_excludes:
                filters_applied.append("type_excludes: {}".format(type_excludes))
            if level_name:
                filters_applied.append("level: {}".format(level_name))
            if in_view:
                filters_applied.append("in_view: True")

            return routes.make_response(
                data={
                    "status": "success",
                    "count": len(filtered_elements),
                    "category": category_name,
                    "filters_applied": filters_applied,
                    "element_ids": filtered_elements[:100]  # Limit to first 100 IDs
                }
            )

        except Exception as e:
            logger.error("Quick count failed: {}".format(str(e)))
            return routes.make_response(
                data={"error": "Failed to count elements: {}".format(str(e))},
                status=500,
            )

    logger.info("Selection routes registered successfully")


def _get_element_details(doc, elem, include_all_parameters=False, include_geometry=False):
    """Get detailed information about an element."""
    try:
        element_info = {
            "id": elem.Id.IntegerValue,
            "name": safe_str(get_element_name(elem)),
            "element_class": safe_str(elem.GetType().Name),
        }

        # Category
        if elem.Category:
            element_info["category"] = safe_str(elem.Category.Name)
            element_info["category_id"] = elem.Category.Id.IntegerValue
        else:
            element_info["category"] = "Unknown"
            element_info["category_id"] = None

        # Type name
        try:
            type_id = elem.GetTypeId()
            if type_id and type_id != DB.ElementId.InvalidElementId:
                type_elem = doc.GetElement(type_id)
                if type_elem:
                    element_info["type_name"] = safe_str(get_element_name(type_elem))
                    element_info["type_id"] = type_id.IntegerValue
                else:
                    element_info["type_name"] = None
                    element_info["type_id"] = None
            else:
                element_info["type_name"] = None
                element_info["type_id"] = None
        except Exception:
            element_info["type_name"] = None
            element_info["type_id"] = None
        
        # Location
        element_info["location"] = _get_location_info(elem)
        
        # Bounding box
        element_info["bbox"] = _get_bbox_info(elem)
        
        # Level
        element_info["level"] = _get_level_info(doc, elem)
        
        # Host element (for hosted elements like windows, doors)
        element_info["host"] = _get_host_info(doc, elem)
        
        # Parameters
        if include_all_parameters:
            element_info["parameters"] = _get_all_parameters(elem)
        else:
            element_info["key_parameters"] = _get_key_parameters(elem)
        
        # Geometry summary
        if include_geometry:
            element_info["geometry"] = _get_geometry_summary(elem)
        
        return element_info
        
    except Exception as e:
        logger.error("Failed to get element details: {}".format(str(e)))
        return {"id": elem.Id.IntegerValue if elem else None, "error": str(e)}


def _get_location_info(elem):
    """Extract location information from element."""
    try:
        location = elem.Location
        if location is None:
            return {"type": "none"}
        
        if hasattr(location, "Point"):
            pt = location.Point
            result = {
                "type": "point",
                "x_feet": pt.X,
                "y_feet": pt.Y,
                "z_feet": pt.Z,
                "x_mm": round(pt.X * 304.8, 1),
                "y_mm": round(pt.Y * 304.8, 1),
                "z_mm": round(pt.Z * 304.8, 1)
            }
            # Try to get rotation
            try:
                if hasattr(location, "Rotation"):
                    import math
                    result["rotation_rad"] = location.Rotation
                    result["rotation_deg"] = round(math.degrees(location.Rotation), 2)
            except Exception:
                pass
            return result
            
        elif hasattr(location, "Curve"):
            curve = location.Curve
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            return {
                "type": "curve",
                "start_mm": {
                    "x": round(start.X * 304.8, 1),
                    "y": round(start.Y * 304.8, 1),
                    "z": round(start.Z * 304.8, 1)
                },
                "end_mm": {
                    "x": round(end.X * 304.8, 1),
                    "y": round(end.Y * 304.8, 1),
                    "z": round(end.Z * 304.8, 1)
                },
                "length_mm": round(curve.Length * 304.8, 1)
            }
        else:
            return {"type": "other", "class": location.GetType().Name}
    except Exception as e:
        return {"type": "error", "message": str(e)}


def _get_bbox_info(elem):
    """Extract bounding box information from element."""
    try:
        bbox = elem.get_BoundingBox(None)
        if bbox is None:
            return None
        
        min_pt = bbox.Min
        max_pt = bbox.Max
        
        return {
            "min_mm": {
                "x": round(min_pt.X * 304.8, 1),
                "y": round(min_pt.Y * 304.8, 1),
                "z": round(min_pt.Z * 304.8, 1)
            },
            "max_mm": {
                "x": round(max_pt.X * 304.8, 1),
                "y": round(max_pt.Y * 304.8, 1),
                "z": round(max_pt.Z * 304.8, 1)
            },
            "width_mm": round((max_pt.X - min_pt.X) * 304.8, 1),
            "depth_mm": round((max_pt.Y - min_pt.Y) * 304.8, 1),
            "height_mm": round((max_pt.Z - min_pt.Z) * 304.8, 1),
            "center_mm": {
                "x": round((min_pt.X + max_pt.X) / 2 * 304.8, 1),
                "y": round((min_pt.Y + max_pt.Y) / 2 * 304.8, 1),
                "z": round((min_pt.Z + max_pt.Z) / 2 * 304.8, 1)
            }
        }
    except Exception as e:
        return {"error": str(e)}


def _get_level_info(doc, elem):
    """Extract level information from element."""
    try:
        # Try FAMILY_LEVEL_PARAM first
        level_param = elem.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
        if level_param:
            level_id = level_param.AsElementId()
            if level_id and level_id != DB.ElementId.InvalidElementId:
                level_elem = doc.GetElement(level_id)
                if level_elem:
                    return {
                        "name": safe_str(get_element_name(level_elem)),
                        "id": level_id.IntegerValue,
                        "elevation_mm": round(level_elem.Elevation * 304.8, 1)
                    }

        # Try LEVEL_PARAM
        level_param = elem.get_Parameter(DB.BuiltInParameter.LEVEL_PARAM)
        if level_param:
            level_id = level_param.AsElementId()
            if level_id and level_id != DB.ElementId.InvalidElementId:
                level_elem = doc.GetElement(level_id)
                if level_elem:
                    return {
                        "name": safe_str(get_element_name(level_elem)),
                        "id": level_id.IntegerValue,
                        "elevation_mm": round(level_elem.Elevation * 304.8, 1)
                    }

        return None
    except Exception:
        return None


def _get_host_info(doc, elem):
    """Get host element info for hosted elements."""
    try:
        if hasattr(elem, "Host") and elem.Host:
            host = elem.Host
            return {
                "id": host.Id.IntegerValue,
                "name": safe_str(get_element_name(host)),
                "category": safe_str(host.Category.Name) if host.Category else "Unknown"
            }
        return None
    except Exception:
        return None


def _get_key_parameters(elem):
    """Get commonly useful parameters only."""
    key_params = {}
    try:
        # Mark
        mark_param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        if mark_param and mark_param.HasValue:
            key_params["Mark"] = safe_str(mark_param.AsString()) or ""

        # Comments
        comments_param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if comments_param and comments_param.HasValue:
            key_params["Comments"] = safe_str(comments_param.AsString()) or ""

        # Width/Height for windows/doors
        width_param = elem.get_Parameter(DB.BuiltInParameter.WINDOW_WIDTH)
        if width_param and width_param.HasValue:
            key_params["Width_mm"] = round(width_param.AsDouble() * 304.8, 1)

        height_param = elem.get_Parameter(DB.BuiltInParameter.WINDOW_HEIGHT)
        if height_param and height_param.HasValue:
            key_params["Height_mm"] = round(height_param.AsDouble() * 304.8, 1)

    except Exception:
        pass
    return key_params


def _get_all_parameters(elem):
    """Get all non-empty parameters from element."""
    all_params = {}
    try:
        for param in elem.Parameters:
            try:
                if not param.HasValue:
                    continue

                param_name = safe_str(param.Definition.Name)
                storage_type = param.StorageType

                if storage_type == DB.StorageType.String:
                    value = param.AsString()
                    if value:
                        all_params[param_name] = safe_str(value)
                elif storage_type == DB.StorageType.Double:
                    value = param.AsDouble()
                    try:
                        display_value = param.AsValueString()
                        all_params[param_name] = safe_str(display_value) if display_value else str(value)
                    except:
                        all_params[param_name] = value
                elif storage_type == DB.StorageType.Integer:
                    all_params[param_name] = param.AsInteger()
                elif storage_type == DB.StorageType.ElementId:
                    elem_id = param.AsElementId()
                    if elem_id and elem_id != DB.ElementId.InvalidElementId:
                        all_params[param_name] = elem_id.IntegerValue
            except Exception:
                continue
    except Exception:
        pass
    return all_params


def _get_geometry_summary(elem):
    """Get summary of element geometry."""
    try:
        options = DB.Options()
        options.ComputeReferences = False
        options.DetailLevel = DB.ViewDetailLevel.Medium
        
        geom = elem.get_Geometry(options)
        if not geom:
            return {"has_geometry": False}
        
        solid_count = 0
        face_count = 0
        curve_count = 0
        
        for geom_obj in geom:
            if isinstance(geom_obj, DB.Solid):
                solid_count += 1
                face_count += geom_obj.Faces.Size
            elif isinstance(geom_obj, DB.Curve):
                curve_count += 1
            elif isinstance(geom_obj, DB.GeometryInstance):
                # Nested geometry
                instance_geom = geom_obj.GetInstanceGeometry()
                if instance_geom:
                    for inst_obj in instance_geom:
                        if isinstance(inst_obj, DB.Solid):
                            solid_count += 1
                            face_count += inst_obj.Faces.Size
                        elif isinstance(inst_obj, DB.Curve):
                            curve_count += 1
        
        return {
            "has_geometry": True,
            "solid_count": solid_count,
            "face_count": face_count,
            "curve_count": curve_count
        }
    except Exception as e:
        return {"has_geometry": False, "error": str(e)}
