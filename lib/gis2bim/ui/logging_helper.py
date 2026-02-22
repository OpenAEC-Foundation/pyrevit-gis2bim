# -*- coding: utf-8 -*-
"""
Logging Helper - GIS2BIM
========================

Gedeelde logging functionaliteit voor GIS2BIM pyRevit tools.
Logt naar zowel de pyRevit output window als een logbestand.

Gebruik:
    from gis2bim.ui.logging_helper import create_tool_logger

    log, LOG_FILE = create_tool_logger("WMS")
    log("Modules geladen")
    # Output: "WMS: Modules geladen" (naar console + logbestand)
"""

import os


def _get_log_directory():
    """Bepaal de centrale log directory met fallback."""
    appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
    log_paths = [
        os.path.join(appdata, 'GIS2BIM', 'logs'),
    ]
    for path in log_paths:
        try:
            if not os.path.exists(path):
                parent = os.path.dirname(path)
                if os.path.exists(parent):
                    os.makedirs(path)
            if os.path.isdir(path):
                test_file = os.path.join(path, '.write_test')
                with open(test_file, 'w') as f:
                    f.write('test')
                os.remove(test_file)
                return path
        except Exception:
            continue
    # Fallback naar temp
    fallback = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), 'GIS2BIM_logs')
    if not os.path.exists(fallback):
        os.makedirs(fallback)
    return fallback


def create_tool_logger(tool_name, script_file=None):
    """Maak een logger functie voor een tool.

    Args:
        tool_name: Naam van de tool (prefix voor log berichten)
        script_file: __file__ van het aanroepende script (niet meer
                     gebruikt voor pad bepaling, behouden voor
                     compatibiliteit).

    Returns:
        Tuple van (log_functie, log_bestandspad)
    """
    log_dir = _get_log_directory()
    log_file = os.path.join(log_dir, "{0}_debug.log".format(tool_name))

    def log(msg):
        """Log naar pyRevit output window EN logbestand."""
        text = "{0}: {1}".format(tool_name, msg)
        print(text)
        try:
            log_dir = os.path.dirname(log_file)
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            with open(log_file, "a") as f:
                f.write(text + "\n")
        except Exception:
            pass

    return log, log_file


def clear_log_file(log_file):
    """Leeg een logbestand (aanroepen bij start van de tool).

    Args:
        log_file: Pad naar het logbestand
    """
    try:
        log_dir = os.path.dirname(log_file)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        with open(log_file, "w") as f:
            f.write("")
    except Exception:
        pass
