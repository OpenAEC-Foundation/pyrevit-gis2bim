# -*- coding: UTF-8 -*-
"""
Parameter routes for Revit MCP
"""

from pyrevit import routes
import logging

logger = logging.getLogger(__name__)


def safe_str(value):
    """Safely convert value to ASCII-safe string for JSON serialization."""
    if value is None:
        return None
    try:
        return str(value).encode('ascii', 'replace').decode('ascii')
    except:
        return str(type(value))


def register_parameter_routes(api):
    """Register parameter routes"""

    @api.route("/param_test/", methods=["GET"])
    def param_test():
        return routes.make_response(data={"status": "parameter_routes_working"})

    @api.route("/get_parameter/", methods=["POST"])
    def get_parameter(request):
        """Get a single parameter value from an element."""
        from pyrevit import revit, DB
        import json
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            element_id = data.get("element_id")
            param_name = data.get("parameter_name")

            if not element_id or not param_name:
                return routes.make_response(data={"error": "element_id and parameter_name required"}, status=400)

            element = doc.GetElement(DB.ElementId(int(element_id)))
            if not element:
                return routes.make_response(data={"error": "Element not found"}, status=404)

            param = element.LookupParameter(param_name)
            if not param:
                type_id = element.GetTypeId()
                if type_id and type_id.IntegerValue > 0:
                    elem_type = doc.GetElement(type_id)
                    if elem_type:
                        param = elem_type.LookupParameter(param_name)

            if not param:
                return routes.make_response(data={"error": "Parameter not found"}, status=404)

            value = None
            display = None
            storage = str(param.StorageType)

            if param.HasValue:
                if param.StorageType == DB.StorageType.String:
                    value = param.AsString()
                    display = value
                elif param.StorageType == DB.StorageType.Integer:
                    value = param.AsInteger()
                    display = param.AsValueString() or str(value)
                elif param.StorageType == DB.StorageType.Double:
                    value = param.AsDouble()
                    display = param.AsValueString() or str(value)
                elif param.StorageType == DB.StorageType.ElementId:
                    value = param.AsElementId().IntegerValue
                    display = str(value)

            return routes.make_response(data={
                "status": "success",
                "element_id": element_id,
                "parameter_name": param.Definition.Name,
                "value": value,
                "display_value": display,
                "storage_type": storage,
                "is_read_only": param.IsReadOnly
            })
        except Exception as e:
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/set_parameter/", methods=["POST"])
    def set_parameter(doc, request):
        """Set a parameter value on an element."""
        from pyrevit import revit, DB
        from Autodesk.Revit.DB import Transaction
        import json
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            element_id = data.get("element_id")
            param_name = data.get("parameter_name")
            new_value = data.get("value")

            if element_id is None or not param_name:
                return routes.make_response(data={"error": "element_id, parameter_name required"}, status=400)

            element = doc.GetElement(DB.ElementId(int(element_id)))
            if not element:
                return routes.make_response(data={"error": "Element not found"}, status=404)

            # Find parameter (instance first, then type)
            param = element.LookupParameter(param_name)
            is_type_param = False
            if not param:
                type_id = element.GetTypeId()
                if type_id and type_id.IntegerValue > 0:
                    elem_type = doc.GetElement(type_id)
                    if elem_type:
                        param = elem_type.LookupParameter(param_name)
                        is_type_param = True

            if not param:
                return routes.make_response(data={"error": "Parameter '{}' not found".format(param_name)}, status=404)

            if param.IsReadOnly:
                return routes.make_response(data={"error": "Parameter '{}' is read-only".format(param_name)}, status=400)

            # Get old value for response
            old_value = None
            if param.HasValue:
                if param.StorageType == DB.StorageType.String:
                    old_value = param.AsString()
                elif param.StorageType == DB.StorageType.Integer:
                    old_value = param.AsInteger()
                elif param.StorageType == DB.StorageType.Double:
                    old_value = param.AsDouble()
                elif param.StorageType == DB.StorageType.ElementId:
                    old_value = param.AsElementId().IntegerValue

            # Set value in transaction
            success = False
            with Transaction(doc, "MCP Set Parameter") as t:
                t.Start()
                if param.StorageType == DB.StorageType.String:
                    success = param.Set(str(new_value) if new_value is not None else "")
                elif param.StorageType == DB.StorageType.Integer:
                    # Handle Yes/No parameters
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
                t.Commit()

            if success:
                return routes.make_response(data={
                    "status": "success",
                    "element_id": element_id,
                    "parameter_name": safe_str(param.Definition.Name),
                    "old_value": old_value,
                    "new_value": new_value,
                    "is_type_parameter": is_type_param
                })
            else:
                return routes.make_response(data={"error": "Failed to set parameter value"}, status=500)

        except Exception as e:
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/get_all_parameters/", methods=["POST"])
    def get_all_parameters(request):
        """Get all parameters from an element (instance + type)."""
        from pyrevit import revit, DB
        import json
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            element_id = data.get("element_id")
            include_empty = data.get("include_empty", False)
            include_readonly = data.get("include_readonly", True)

            if not element_id:
                return routes.make_response(data={"error": "element_id required"}, status=400)

            element = doc.GetElement(DB.ElementId(int(element_id)))
            if not element:
                return routes.make_response(data={"error": "Element not found"}, status=404)

            def extract_params(elem, is_type=False):
                params_list = []
                for param in elem.Parameters:
                    try:
                        if not include_readonly and param.IsReadOnly:
                            continue

                        has_value = param.HasValue
                        if not include_empty and not has_value:
                            continue

                        param_info = {
                            "name": safe_str(param.Definition.Name),
                            "storage_type": str(param.StorageType),
                            "is_read_only": param.IsReadOnly,
                            "is_type_parameter": is_type,
                            "has_value": has_value
                        }

                        if has_value:
                            if param.StorageType == DB.StorageType.String:
                                param_info["value"] = safe_str(param.AsString())
                                param_info["display_value"] = safe_str(param.AsString())
                            elif param.StorageType == DB.StorageType.Integer:
                                param_info["value"] = param.AsInteger()
                                param_info["display_value"] = safe_str(param.AsValueString()) or str(param.AsInteger())
                            elif param.StorageType == DB.StorageType.Double:
                                param_info["value"] = param.AsDouble()
                                param_info["display_value"] = safe_str(param.AsValueString()) or str(param.AsDouble())
                            elif param.StorageType == DB.StorageType.ElementId:
                                param_info["value"] = param.AsElementId().IntegerValue
                                param_info["display_value"] = str(param.AsElementId().IntegerValue)
                        else:
                            param_info["value"] = None
                            param_info["display_value"] = None

                        params_list.append(param_info)
                    except Exception:
                        continue
                return params_list

            instance_params = extract_params(element, is_type=False)

            type_params = []
            type_id = element.GetTypeId()
            if type_id and type_id.IntegerValue > 0:
                elem_type = doc.GetElement(type_id)
                if elem_type:
                    type_params = extract_params(elem_type, is_type=True)

            return routes.make_response(data={
                "status": "success",
                "element_id": element_id,
                "instance_parameters": instance_params,
                "type_parameters": type_params,
                "total_count": len(instance_params) + len(type_params)
            })

        except Exception as e:
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/set_parameters_bulk/", methods=["POST"])
    def set_parameters_bulk(doc, request):
        """Set multiple parameters on a single element in one transaction."""
        from pyrevit import revit, DB
        from Autodesk.Revit.DB import Transaction
        import json
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            element_id = data.get("element_id")
            parameters = data.get("parameters", {})

            if not element_id:
                return routes.make_response(data={"error": "element_id required"}, status=400)

            if not parameters:
                return routes.make_response(data={"error": "parameters dict required"}, status=400)

            element = doc.GetElement(DB.ElementId(int(element_id)))
            if not element:
                return routes.make_response(data={"error": "Element not found"}, status=404)

            # Get type element for type parameters
            elem_type = None
            type_id = element.GetTypeId()
            if type_id and type_id.IntegerValue > 0:
                elem_type = doc.GetElement(type_id)

            results = {"success": [], "failed": []}

            with Transaction(doc, "MCP Bulk Set Parameters") as t:
                t.Start()
                for param_name, new_value in parameters.items():
                    try:
                        # Find parameter
                        param = element.LookupParameter(param_name)
                        if not param and elem_type:
                            param = elem_type.LookupParameter(param_name)

                        if not param:
                            results["failed"].append({"parameter": param_name, "error": "Not found"})
                            continue

                        if param.IsReadOnly:
                            results["failed"].append({"parameter": param_name, "error": "Read-only"})
                            continue

                        # Set value based on storage type
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
                            results["success"].append({"parameter": param_name, "value": new_value})
                        else:
                            results["failed"].append({"parameter": param_name, "error": "Set failed"})

                    except Exception as e:
                        results["failed"].append({"parameter": param_name, "error": str(e)})
                t.Commit()

            return routes.make_response(data={
                "status": "success",
                "element_id": element_id,
                "results": results,
                "success_count": len(results["success"]),
                "failed_count": len(results["failed"])
            })

        except Exception as e:
            return routes.make_response(data={"error": str(e)}, status=500)

    @api.route("/set_parameters_multi/", methods=["POST"])
    def set_parameters_multi(doc, request):
        """Set same parameters on multiple elements in one transaction."""
        from pyrevit import revit, DB
        from Autodesk.Revit.DB import Transaction
        import json
        doc = revit.doc
        try:
            data = json.loads(request.data) if isinstance(request.data, str) else request.data
            element_ids = data.get("element_ids", [])
            parameters = data.get("parameters", {})

            if not element_ids:
                return routes.make_response(data={"error": "element_ids list required"}, status=400)

            if not parameters:
                return routes.make_response(data={"error": "parameters dict required"}, status=400)

            results = {"elements": [], "summary": {"total": len(element_ids), "success": 0, "partial": 0, "failed": 0}}

            with Transaction(doc, "MCP Multi-Element Set Parameters") as t:
                t.Start()
                for elem_id in element_ids:
                    elem_result = {"element_id": elem_id, "success": [], "failed": []}

                    element = doc.GetElement(DB.ElementId(int(elem_id)))
                    if not element:
                        elem_result["failed"].append({"error": "Element not found"})
                        results["elements"].append(elem_result)
                        results["summary"]["failed"] += 1
                        continue

                    # Get type element
                    elem_type = None
                    type_id = element.GetTypeId()
                    if type_id and type_id.IntegerValue > 0:
                        elem_type = doc.GetElement(type_id)

                    for param_name, new_value in parameters.items():
                        try:
                            param = element.LookupParameter(param_name)
                            if not param and elem_type:
                                param = elem_type.LookupParameter(param_name)

                            if not param:
                                elem_result["failed"].append({"parameter": param_name, "error": "Not found"})
                                continue

                            if param.IsReadOnly:
                                elem_result["failed"].append({"parameter": param_name, "error": "Read-only"})
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
                                elem_result["success"].append(param_name)
                            else:
                                elem_result["failed"].append({"parameter": param_name, "error": "Set failed"})

                        except Exception as e:
                            elem_result["failed"].append({"parameter": param_name, "error": str(e)})

                    results["elements"].append(elem_result)

                    # Update summary
                    if not elem_result["failed"]:
                        results["summary"]["success"] += 1
                    elif elem_result["success"]:
                        results["summary"]["partial"] += 1
                    else:
                        results["summary"]["failed"] += 1
                t.Commit()

            return routes.make_response(data={
                "status": "success",
                "results": results
            })

        except Exception as e:
            return routes.make_response(data={"error": str(e)}, status=500)

    logger.info("Parameter routes registered")
