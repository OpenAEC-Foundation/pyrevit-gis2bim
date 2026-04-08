# -*- coding: UTF-8 -*-
"""
Code Execution Module for Revit MCP
Handles direct execution of IronPython code in Revit context.
"""
from pyrevit import routes, revit, DB
import json
import logging
import sys
import traceback
from StringIO import StringIO

# Standard logger setup
logger = logging.getLogger(__name__)


def register_code_execution_routes(api):
    """Register code execution routes with the API."""

    @api.route("/execute_code/", methods=["POST"])
    def execute_code(request):
        """Execute IronPython code in Revit context."""
        doc = revit.doc
        try:
            # Parse the request data
            data = (
                json.loads(request.data)
                if isinstance(request.data, str)
                else request.data
            )
            code_to_execute = data.get("code", "")
            description = data.get("description", "Code execution")
            # New option: wrap in transaction
            use_transaction = data.get("use_transaction", False)

            if not code_to_execute:
                return routes.make_response(
                    data={"error": "No code provided"}, status=400
                )

            logger.info("Executing code: {}".format(description))

            # Capture stdout
            old_stdout = sys.stdout
            captured_output = StringIO()
            sys.stdout = captured_output

            try:
                # Create namespace with common Revit objects
                namespace = {
                    "doc": doc,
                    "DB": DB,
                    "revit": revit,
                    "__builtins__": __builtins__,
                    "print": lambda *args: captured_output.write(
                        " ".join(str(arg) for arg in args) + "\n"
                    ),
                }

                # Execute with or without transaction wrapper
                if use_transaction:
                    # Start transaction in route handler context (where it works!)
                    t = DB.Transaction(doc, "MCP: {}".format(description))
                    t.Start()
                    try:
                        exec(code_to_execute, namespace)
                        t.Commit()
                    except Exception as exec_err:
                        if t.HasStarted() and not t.HasEnded():
                            t.RollBack()
                        raise exec_err
                else:
                    # Read-only execution
                    exec(code_to_execute, namespace)

                # Restore stdout
                sys.stdout = old_stdout
                output = captured_output.getvalue()
                captured_output.close()

                return routes.make_response(
                    data={
                        "status": "success",
                        "description": description,
                        "output": output if output else "Code executed successfully (no output)",
                        "transaction_used": use_transaction,
                    }
                )

            except Exception as exec_error:
                sys.stdout = old_stdout
                error_traceback = traceback.format_exc()
                logger.error("Code execution failed: {}".format(str(exec_error)))

                return routes.make_response(
                    data={
                        "status": "error",
                        "error": str(exec_error),
                        "traceback": error_traceback,
                        "code_attempted": code_to_execute,
                    },
                    status=500,
                )

        except Exception as e:
            logger.error("Execute code request failed: {}".format(str(e)))
            return routes.make_response(data={"error": str(e)}, status=500)

    logger.info("Code execution routes registered successfully.")
