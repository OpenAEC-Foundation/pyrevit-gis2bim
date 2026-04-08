# -*- coding: UTF-8 -*-
"""
IFC Query Module for Revit MCP
Query elements in linked IFC models by category, parameters, or IFC properties.
"""

from pyrevit import routes, revit, DB
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


def get_param_value(param):
    """Extract parameter value safely."""
    if not param or not param.HasValue:
        return None
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        elif param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == DB.StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == DB.StorageType.ElementId:
            return param.AsElementId().IntegerValue
    except:
        pass
    return None


def register_ifc_query_routes(api):
    """Register IFC query routes with the API."""

    @api.route("/query_ifc_elements/", methods=["POST"])
    def query_ifc_elements(doc, request):
        """
        Query elements in linked IFC models.

        Filters:
        - link_name: Name of the linked model (partial match)
        - category: Revit category name (e.g., "Windows", "Doors", "Walls")
        - ifc_class: IFC class name (e.g., "IfcWall", "IfcDoor")
        - parameter_name: Parameter to filter on
        - parameter_value: Value to match (partial match for strings)
        - max_results: Maximum number of results (default 100)
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data

            link_name_filter = data.get("link_name", "")
            category_filter = data.get("category", "")
            ifc_class_filter = data.get("ifc_class", "")
            param_name = data.get("parameter_name", "")
            param_value = data.get("parameter_value", "")
            max_results = data.get("max_results", 100)

            logger.info("Querying IFC elements - link: {}, category: {}, ifc_class: {}".format(
                link_name_filter, category_filter, ifc_class_filter))

            results = []
            links_searched = []

            # Get all linked models
            collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)

            for link_instance in collector:
                try:
                    link_doc = link_instance.GetLinkDocument()
                    if not link_doc:
                        continue

                    link_title = safe_str(link_doc.Title)

                    # Filter by link name if specified
                    if link_name_filter and link_name_filter.lower() not in link_title.lower():
                        continue

                    links_searched.append(link_title)

                    # Build element collector for linked doc
                    elem_collector = DB.FilteredElementCollector(link_doc).WhereElementIsNotElementType()

                    # Filter by category if specified
                    if category_filter:
                        target_category = None
                        for cat in link_doc.Settings.Categories:
                            if cat.Name.lower() == category_filter.lower():
                                target_category = cat
                                break
                        if target_category:
                            elem_collector = elem_collector.OfCategoryId(target_category.Id)
                        else:
                            # Category not found in this link, skip
                            continue

                    elements = elem_collector.ToElements()

                    for elem in elements:
                        if len(results) >= max_results:
                            break

                        try:
                            # Filter by IFC class if specified
                            if ifc_class_filter:
                                ifc_type_param = elem.LookupParameter("Export to IFC As")
                                if not ifc_type_param:
                                    ifc_type_param = elem.LookupParameter("IfcExportAs")
                                ifc_type = get_param_value(ifc_type_param) if ifc_type_param else ""

                                # Also check category-based IFC mapping
                                if not ifc_type and elem.Category:
                                    cat_name = elem.Category.Name
                                    # Map common Revit categories to IFC classes
                                    ifc_map = {
                                        "Walls": "IfcWall",
                                        "Doors": "IfcDoor",
                                        "Windows": "IfcWindow",
                                        "Floors": "IfcSlab",
                                        "Roofs": "IfcRoof",
                                        "Columns": "IfcColumn",
                                        "Stairs": "IfcStair",
                                        "Railings": "IfcRailing",
                                        "Furniture": "IfcFurnishingElement",
                                    }
                                    ifc_type = ifc_map.get(cat_name, "")

                                if ifc_class_filter.lower() not in str(ifc_type).lower():
                                    continue

                            # Filter by parameter name/value if specified
                            if param_name:
                                param = elem.LookupParameter(param_name)
                                if not param:
                                    # Try type parameter
                                    type_id = elem.GetTypeId()
                                    if type_id and type_id.IntegerValue > 0:
                                        elem_type = link_doc.GetElement(type_id)
                                        if elem_type:
                                            param = elem_type.LookupParameter(param_name)

                                if not param:
                                    continue

                                value = get_param_value(param)
                                if param_value:
                                    # Check if value matches (partial for strings)
                                    if isinstance(value, str):
                                        if param_value.lower() not in value.lower():
                                            continue
                                    else:
                                        if str(param_value) != str(value):
                                            continue

                            # Element passed all filters, add to results
                            elem_info = {
                                "link_name": link_title,
                                "link_instance_id": link_instance.Id.IntegerValue,
                                "element_id": elem.Id.IntegerValue,
                                "name": safe_str(get_element_name(elem)),
                                "category": safe_str(elem.Category.Name) if elem.Category else "Unknown",
                            }

                            # Add type info
                            try:
                                type_id = elem.GetTypeId()
                                if type_id and type_id.IntegerValue > 0:
                                    elem_type = link_doc.GetElement(type_id)
                                    if elem_type:
                                        elem_info["type_name"] = safe_str(get_element_name(elem_type))
                            except:
                                pass

                            # Add some common IFC parameters
                            ifc_params = {}
                            for pname in ["IfcGUID", "IfcName", "IfcDescription", "IfcTag",
                                         "IFC Predefined Type", "Export to IFC As"]:
                                p = elem.LookupParameter(pname)
                                if p:
                                    val = get_param_value(p)
                                    if val:
                                        ifc_params[pname] = safe_str(str(val))
                            if ifc_params:
                                elem_info["ifc_properties"] = ifc_params

                            # Add matched parameter if filtering by parameter
                            if param_name and param:
                                elem_info["matched_parameter"] = {
                                    "name": param_name,
                                    "value": safe_str(str(get_param_value(param)))
                                }

                            results.append(elem_info)

                        except Exception as e:
                            logger.warning("Could not process element: {}".format(str(e)))
                            continue

                    if len(results) >= max_results:
                        break

                except Exception as e:
                    logger.warning("Could not process link: {}".format(str(e)))
                    continue

            return routes.make_response(data={
                "status": "success",
                "count": len(results),
                "max_results": max_results,
                "links_searched": links_searched,
                "filters": {
                    "link_name": link_name_filter or None,
                    "category": category_filter or None,
                    "ifc_class": ifc_class_filter or None,
                    "parameter_name": param_name or None,
                    "parameter_value": param_value or None
                },
                "elements": results
            })

        except Exception as e:
            logger.error("Query IFC elements failed: {}".format(str(e)))
            return routes.make_response(
                data={"error": "Query failed: {}".format(str(e))},
                status=500
            )

    @api.route("/get_ifc_element_properties/", methods=["POST"])
    def get_ifc_element_properties(doc, request):
        """
        Get all properties from an element in a linked IFC model.

        Required:
        - link_instance_id: ID of the RevitLinkInstance
        - element_id: ID of the element within the linked model
        """
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data

            link_instance_id = data.get("link_instance_id")
            element_id = data.get("element_id")

            if not link_instance_id or not element_id:
                return routes.make_response(
                    data={"error": "link_instance_id and element_id are required"},
                    status=400
                )

            # Get the link instance
            link_instance = doc.GetElement(DB.ElementId(int(link_instance_id)))
            if not link_instance:
                return routes.make_response(
                    data={"error": "Link instance not found"},
                    status=404
                )

            link_doc = link_instance.GetLinkDocument()
            if not link_doc:
                return routes.make_response(
                    data={"error": "Linked document not loaded"},
                    status=404
                )

            # Get the element from the linked doc
            elem = link_doc.GetElement(DB.ElementId(int(element_id)))
            if not elem:
                return routes.make_response(
                    data={"error": "Element not found in linked model"},
                    status=404
                )

            # Collect all parameters
            instance_params = []
            type_params = []

            for param in elem.Parameters:
                try:
                    if not param.HasValue:
                        continue

                    param_info = {
                        "name": safe_str(param.Definition.Name),
                        "value": safe_str(str(get_param_value(param))),
                        "storage_type": str(param.StorageType),
                        "is_read_only": param.IsReadOnly
                    }

                    # Try to get display value
                    try:
                        display = param.AsValueString()
                        if display:
                            param_info["display_value"] = safe_str(display)
                    except:
                        pass

                    instance_params.append(param_info)
                except:
                    continue

            # Get type parameters
            type_id = elem.GetTypeId()
            if type_id and type_id.IntegerValue > 0:
                elem_type = link_doc.GetElement(type_id)
                if elem_type:
                    for param in elem_type.Parameters:
                        try:
                            if not param.HasValue:
                                continue

                            param_info = {
                                "name": safe_str(param.Definition.Name),
                                "value": safe_str(str(get_param_value(param))),
                                "storage_type": str(param.StorageType),
                                "is_read_only": param.IsReadOnly
                            }

                            try:
                                display = param.AsValueString()
                                if display:
                                    param_info["display_value"] = safe_str(display)
                            except:
                                pass

                            type_params.append(param_info)
                        except:
                            continue

            return routes.make_response(data={
                "status": "success",
                "link_name": safe_str(link_doc.Title),
                "element_id": element_id,
                "name": safe_str(get_element_name(elem)),
                "category": safe_str(elem.Category.Name) if elem.Category else "Unknown",
                "instance_parameters": instance_params,
                "type_parameters": type_params,
                "total_parameters": len(instance_params) + len(type_params)
            })

        except Exception as e:
            logger.error("Get IFC element properties failed: {}".format(str(e)))
            return routes.make_response(
                data={"error": "Failed to get properties: {}".format(str(e))},
                status=500
            )

    logger.info("IFC query routes registered successfully")
