# -*- coding: utf-8 -*-
"""
NAP Peilmerken - GIS2BIM
=========================

Haal NAP peilmerken op van Rijkswaterstaat en plaats ze
als kruismarkers met hoogte-labels op een Revit plan view.
"""

__title__ = "NAP"
__author__ = "OpenAEC Foundation"
__doc__ = "Plaats NAP hoogtepunten (peilmerken) van Rijkswaterstaat op een plan view"

# CLR references voor WPF
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
clr.AddReference('System.Xml')

import System
from System.Windows import Window, Visibility

# pyRevit
from pyrevit import revit, DB, script, forms

# Standaard library
import sys
import os
import traceback

# Voeg lib folder toe aan path
extension_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(extension_path, "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Gedeelde modules
from gis2bim.ui.logging_helper import create_tool_logger, clear_log_file
from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements
from gis2bim.ui.location_setup import setup_project_location
from gis2bim.ui.view_setup import populate_view_dropdown, get_selected_view
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

log, LOG_FILE = create_tool_logger("NAP", __file__)

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.nap import NAPClient
    from gis2bim.revit.geometry import (
        rd_to_revit_xyz,
        create_text_notes,
        meters_to_feet,
    )
    from gis2bim.coordinates import create_bbox_rd
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


class NAPPeilmerkenWindow(Window):
    """WPF Window voor NAP Peilmerken laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location', 'cmb_bbox_size',
            'chk_puntnummer', 'chk_bereikbaar',
            'cmb_kruis_grootte', 'cmb_view',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        populate_view_dropdown(self.cmb_view, doc, log=log)
        self._bind_events()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_execute(self, sender, args):
        log("Execute gestart")

        if self.location_rd is None:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        try:
            show_progress(self, "Voorbereiden...")
            self.btn_execute.IsEnabled = False

            bbox_size = self._get_bbox_size()
            view = get_selected_view(self.cmb_view, self.doc)
            alleen_bereikbaar = self.chk_bereikbaar.IsChecked
            toon_puntnummer = self.chk_puntnummer.IsChecked
            kruis_grootte = self._get_kruis_grootte()

            if view is None:
                self.txt_status.Text = "Selecteer een view"
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            center_x, center_y = self.location_rd
            bbox = create_bbox_rd(center_x, center_y, bbox_size)
            log("Bbox: {0}".format(bbox))

            # NAP data ophalen
            show_progress(self, "NAP peilmerken ophalen van Rijkswaterstaat...")
            update_ui()

            client = NAPClient()
            peilmerken = client.get_peilmerken(bbox, alleen_bereikbaar)
            log("Peilmerken opgehaald: {0}".format(len(peilmerken)))

            if not peilmerken:
                self.txt_status.Text = "Geen peilmerken gevonden in zoekgebied"
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            # Tekenen in Revit
            show_progress(self, "Peilmerken tekenen in Revit ({0} stuks)...".format(
                len(peilmerken)))
            update_ui()

            stats = self._draw_peilmerken(
                peilmerken, view, kruis_grootte, toon_puntnummer)

            hide_progress(self)
            self.DialogResult = True
            self.Close()

            self._show_result(stats, bbox_size, len(peilmerken))

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            hide_progress(self)
            self.btn_execute.IsEnabled = True

    def _get_bbox_size(self):
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 1000

    def _get_kruis_grootte(self):
        item = self.cmb_kruis_grootte.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 5

    def _draw_peilmerken(self, peilmerken, view, kruis_grootte, toon_puntnummer):
        """Teken kruismarkers en labels voor peilmerken in Revit.

        Args:
            peilmerken: Lijst van NAPPeilmerk objecten
            view: Revit View voor plaatsing
            kruis_grootte: Grootte van het kruis in meters
            toon_puntnummer: Of het puntnummer bij het label moet

        Returns:
            Dict met statistieken
        """
        stats = {
            "kruisen": 0,
            "labels": 0,
            "errors": 0
        }

        origin_x, origin_y = self.location_rd
        half_size = meters_to_feet(kruis_grootte / 2.0)

        # Bouw annotatie-lijst voor text notes
        annotations = []
        for pm in peilmerken:
            label = pm.hoogte_label
            if toon_puntnummer:
                label = "{0}\n{1}".format(pm.puntnummer, label)
            annotations.append({
                "text": label,
                "x": pm.x_rd,
                "y": pm.y_rd,
                "rotation": 0
            })

        t = DB.Transaction(self.doc, "GIS2BIM - NAP Peilmerken")
        t.Start()

        try:
            # Teken kruismarkers (detail lines op de view)
            for pm in peilmerken:
                try:
                    center = rd_to_revit_xyz(
                        pm.x_rd, pm.y_rd, origin_x, origin_y)

                    # Horizontale lijn
                    pt_h1 = DB.XYZ(center.X - half_size, center.Y, 0)
                    pt_h2 = DB.XYZ(center.X + half_size, center.Y, 0)
                    line_h = DB.Line.CreateBound(pt_h1, pt_h2)
                    self.doc.Create.NewDetailCurve(view, line_h)

                    # Verticale lijn
                    pt_v1 = DB.XYZ(center.X, center.Y - half_size, 0)
                    pt_v2 = DB.XYZ(center.X, center.Y + half_size, 0)
                    line_v = DB.Line.CreateBound(pt_v1, pt_v2)
                    self.doc.Create.NewDetailCurve(view, line_v)

                    stats["kruisen"] += 1

                except Exception as e:
                    stats["errors"] += 1
                    if stats["errors"] <= 3:
                        log("Fout bij kruis voor {0}: {1}".format(
                            pm.puntnummer, e))

            # Teken labels (text notes)
            label_ids = create_text_notes(
                self.doc, view, annotations,
                (origin_x, origin_y))
            stats["labels"] = len(label_ids)

            t.Commit()
            log("Transaction commit OK")

        except Exception as e:
            t.RollBack()
            log("Transaction rollback: {0}".format(e))
            raise

        log("Stats: {0}".format(stats))
        return stats

    def _show_result(self, stats, bbox_size, totaal):
        msg_lines = [
            "NAP peilmerken succesvol geplaatst!",
            "",
            "Zoekgebied: {0} x {0} m".format(bbox_size),
            "Peilmerken gevonden: {0}".format(totaal),
            "",
            "Aangemaakt:",
            "  Kruismarkers: {0}".format(stats["kruisen"]),
            "  Hoogte-labels: {0}".format(stats["labels"]),
        ]

        if stats["errors"] > 0:
            msg_lines.append("  Fouten: {0}".format(stats["errors"]))

        forms.alert("\n".join(msg_lines), title="GIS2BIM - NAP Peilmerken")


def main():
    clear_log_file(LOG_FILE)
    log("=== GIS2BIM NAP Peilmerken Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = NAPPeilmerkenWindow(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM NAP Peilmerken Tool beeindigd ===")


if __name__ == "__main__":
    main()
