# -*- coding: UTF-8 -*-
"""
Revit MCP Extension Startup
Registers all MCP routes and initializes the API
"""

from pyrevit import routes
import logging
import sys
import os

logger = logging.getLogger(__name__)

# Add revit_mcp folder to sys.path so "from utils import ..." works
_this_dir = os.path.dirname(__file__)
_mcp_dir = os.path.join(_this_dir, "revit_mcp")
if _mcp_dir not in sys.path:
    sys.path.insert(0, _mcp_dir)

# Initialize the main API
api = routes.API("revit_mcp")


def register_routes():
    """Register all MCP route modules"""
    try:
        from revit_mcp.status import register_status_routes
        register_status_routes(api)

        from revit_mcp.model_info import register_model_info_routes
        register_model_info_routes(api)

        from revit_mcp.views import register_views_routes
        register_views_routes(api)

        from revit_mcp.placement import register_placement_routes
        register_placement_routes(api)

        from revit_mcp.colors import register_color_routes
        register_color_routes(api)

        from revit_mcp.code_execution import register_code_execution_routes
        register_code_execution_routes(api)

        from revit_mcp.selection import register_selection_routes
        register_selection_routes(api)

        from revit_mcp.parameters import register_parameter_routes
        register_parameter_routes(api)

        from revit_mcp.ifc_query import register_ifc_query_routes
        register_ifc_query_routes(api)

        from revit_mcp.modification import register_modification_routes
        register_modification_routes(api)

        logger.info("All MCP routes registered successfully")

    except Exception as e:
        logger.error("Failed to register MCP routes: %s", str(e))
        raise


# Register all routes when the extension loads
register_routes()
