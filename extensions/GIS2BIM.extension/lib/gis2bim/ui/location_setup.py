# -*- coding: utf-8 -*-
"""
Location Setup - GIS2BIM
========================

Gedeelde locatie-setup logica voor GIS2BIM pyRevit tools.
Haalt RD coordinaten op uit project parameters of Survey Point,
valideert ze, en stelt de UI in.

Vereist dat het Window object de volgende attributen heeft:
    - pnl_location (StackPanel)
    - pnl_no_location (StackPanel)
    - txt_location_x (TextBlock)
    - txt_location_y (TextBlock)
    - btn_execute (Button)

Gebruik:
    from gis2bim.ui.location_setup import setup_project_location

    class MijnWindow(Window):
        def __init__(self, doc):
            ...
            self.location_rd = setup_project_location(self, doc, log)
"""

import traceback
from System.Windows import Visibility

from gis2bim.revit.location import get_project_location_rd, get_rd_from_project_params


def is_valid_rd(rd_x, rd_y):
    """Controleer of RD coordinaten binnen Nederland vallen.

    Args:
        rd_x: RD X coordinaat
        rd_y: RD Y coordinaat

    Returns:
        True als de coordinaten geldig zijn
    """
    return 0 < rd_x < 300000 and 289000 < rd_y < 629000


def setup_project_location(window, doc, log=None):
    """Haal projectlocatie op en stel de UI in.

    Probeert eerst GIS2BIM_RD_X/Y project parameters,
    daarna Survey Point als fallback.
    Valideert RD coordinaten en corrigeert X/Y swap indien nodig.

    Vereist UI elementen op window:
        pnl_location, pnl_no_location, txt_location_x,
        txt_location_y, btn_execute

    Args:
        window: Het Window object met UI elementen
        doc: Revit Document
        log: Optionele log functie (callable)

    Returns:
        Tuple (rd_x, rd_y) bij succes, of None bij geen locatie
    """
    if log is None:
        log = lambda msg: None

    try:
        rd_params = get_rd_from_project_params(doc)

        if rd_params:
            rd_x = rd_params["rd_x"]
            rd_y = rd_params["rd_y"]
            log("Locatie uit project parameters: RD {0}, {1}".format(rd_x, rd_y))
        else:
            location = get_project_location_rd(doc)
            if location and (location["rd_x"] != 0 or location["rd_y"] != 0):
                rd_x = location["rd_x"]
                rd_y = location["rd_y"]
                log("Locatie uit Survey Point: RD {0}, {1}".format(rd_x, rd_y))
            else:
                _show_no_location(window)
                return None

        # Valideer RD coordinaten (Nederland)
        if is_valid_rd(rd_x, rd_y):
            _show_location(window, rd_x, rd_y)
            log("Locatie OK: RD X={0:.0f}, Y={1:.0f}".format(rd_x, rd_y))
            return (rd_x, rd_y)
        else:
            log("WAARSCHUWING: RD coords buiten bereik: {0}, {1}".format(rd_x, rd_y))
            # Probeer omgedraaid (veelvoorkomend bij Survey Point)
            if is_valid_rd(rd_y, rd_x):
                _show_location(window, rd_y, rd_x)
                log("Locatie gecorrigeerd (X/Y swap)")
                return (rd_y, rd_x)
            else:
                _show_no_location(window)
                return None

    except Exception as e:
        log("Error getting location: {0}".format(e))
        log(traceback.format_exc())
        _show_no_location(window)
        return None


def _show_location(window, rd_x, rd_y):
    """Toon locatie in de UI."""
    window.pnl_location.Visibility = Visibility.Visible
    window.pnl_no_location.Visibility = Visibility.Collapsed
    window.txt_location_x.Text = "{0:,.0f} m".format(rd_x)
    window.txt_location_y.Text = "{0:,.0f} m".format(rd_y)
    window.btn_execute.IsEnabled = True


def _show_no_location(window):
    """Toon geen-locatie waarschuwing in de UI."""
    window.pnl_location.Visibility = Visibility.Collapsed
    window.pnl_no_location.Visibility = Visibility.Visible
    window.btn_execute.IsEnabled = False
