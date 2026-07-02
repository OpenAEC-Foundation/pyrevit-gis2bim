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
    """Register all MCP route modules independently.

    A failure in one module no longer blocks the others, and full
    tracebacks are written to a diagnostics file (GITHUB copy marker).
    """
    import traceback

    diag_path = os.path.join(
        os.environ.get("TEMP", _this_dir), "revit_mcp_startup_GITHUB.log"
    )
    results = ["LOADED COPY: " + _this_dir]

    def _try(name, func):
        try:
            func()
            results.append("ok    " + name)
        except Exception:
            results.append("FAIL  " + name + "\n" + traceback.format_exc())

    def _status():
        from revit_mcp.status import register_status_routes
        register_status_routes(api)

    def _model_info():
        from revit_mcp.model_info import register_model_info_routes
        register_model_info_routes(api)

    def _views():
        from revit_mcp.views import register_views_routes
        register_views_routes(api)

    def _placement():
        from revit_mcp.placement import register_placement_routes
        register_placement_routes(api)

    def _colors():
        from revit_mcp.colors import register_color_routes
        register_color_routes(api)

    def _code_execution():
        from revit_mcp.code_execution import register_code_execution_routes
        register_code_execution_routes(api)

    def _selection():
        from revit_mcp.selection import register_selection_routes
        register_selection_routes(api)

    def _parameters():
        from revit_mcp.parameters import register_parameter_routes
        register_parameter_routes(api)

    def _ifc_query():
        from revit_mcp.ifc_query import register_ifc_query_routes
        register_ifc_query_routes(api)

    def _modification():
        from revit_mcp.modification import register_modification_routes
        register_modification_routes(api)

    _try("status", _status)
    _try("model_info", _model_info)
    _try("views", _views)
    _try("placement", _placement)
    _try("colors", _colors)
    _try("code_execution", _code_execution)
    _try("selection", _selection)
    _try("parameters", _parameters)
    _try("ifc_query", _ifc_query)
    _try("modification", _modification)

    # --- Diagnostics: engine + readback of the actual route store ---
    try:
        results.append("ENGINE sys.version: " + str(sys.version))
    except Exception:
        results.append("ENGINE sys.version: <unavailable>")
    try:
        results.append("sys.executable: " + str(getattr(sys, "executable", "?")))
    except Exception:
        pass
    try:
        from pyrevit.routes.server import router as _router
        _stored = _router.get_routes("revit_mcp")
        results.append("STORE revit_mcp route count: " + str(len(_stored)))
        for _rt in list(_stored.keys()):
            results.append("   route: %s %s" % (_rt.method, _rt.pattern))
    except Exception:
        import traceback as _tb
        results.append("STORE readback FAILED:\n" + _tb.format_exc())
    try:
        from pyrevit.coreutils import envvars as _ev
        _root = _ev.get_pyrevit_env_vars()
        results.append("ENVVAR root dict id: " + str(id(_root)))
        results.append("ROUTES_SERVER present: " + str(bool(_ev.get_pyrevit_env_var(_ev.ROUTES_SERVER))))
    except Exception:
        import traceback as _tb2
        results.append("ENVVAR introspect FAILED:\n" + _tb2.format_exc())

    report = "\n".join(results)
    try:
        with open(diag_path, "w") as f:
            f.write(report)
    except Exception:
        pass
    logger.info("MCP route registration report:\n%s", report)


# Register all routes when the extension loads
register_routes()
