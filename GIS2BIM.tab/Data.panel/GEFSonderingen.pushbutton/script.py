# -*- coding: utf-8 -*-
"""
GEF Sonderingen & Boringen - GIS2BIM
======================================

Haal geotechnische data (CPT sonderingen en boringen) op van de
BRO (Basisregistratie Ondergrond) en visualiseer als 3D cilinders
met gekleurde grondlagen in Revit.
"""

__title__ = "GEF"
__author__ = "OpenAEC Foundation"
__doc__ = "Laad CPT sonderingen en boringen uit de BRO als 3D cilinders"

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
extension_path = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(__file__))))
lib_path = os.path.join(extension_path, "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Gedeelde modules
from gis2bim.ui.logging_helper import create_tool_logger, clear_log_file
from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements
from gis2bim.ui.location_setup import setup_project_location
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

log, LOG_FILE = create_tool_logger("GEFSonderingen", __file__)

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.bro import (
        BROClient, CPT_KLEUR, get_grondsoort_kleur,
        get_qc_kleur, BRO_VIEWER_URL,
    )
    from gis2bim.revit.geometry import (
        rd_to_revit_xyz,
        create_cylinder_solid,
        create_directshape,
        set_element_color,
        get_or_create_material,
        set_element_material,
        ensure_bro_parameters,
        set_element_parameter,
        METER_TO_FEET,
    )
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


class GEFSonderingenWindow(Window):
    """WPF Window voor GEF Sonderingen & Boringen laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.result_count = 0

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x',
            'txt_location_y', 'pnl_no_location',
            'cmb_straal', 'cmb_diameter',
            'chk_cpt', 'chk_bhr',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
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

            straal = self._get_straal()
            diameter = self._get_diameter()
            laden_cpt = self.chk_cpt.IsChecked
            laden_bhr = self.chk_bhr.IsChecked

            if not laden_cpt and not laden_bhr:
                self.txt_status.Text = "Selecteer CPT en/of Boringen"
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            center_x, center_y = self.location_rd
            log("Locatie: {0}, {1}, straal: {2}m".format(
                center_x, center_y, straal))

            client = BROClient()
            stats = {"cpt": 0, "bhr": 0, "lagen": 0, "errors": 0}

            # CPT laden (search → detail per CPT voor meetdata)
            if laden_cpt:
                show_progress(self, "CPT sonderingen zoeken in BRO...")
                update_ui()

                cpt_data_list = client.search_cpt(
                    center_x, center_y, straal)
                log("CPT gevonden: {0}".format(len(cpt_data_list)))

                if cpt_data_list:
                    # Detail ophalen per CPT (BUITEN transactie)
                    show_progress(self,
                        "{0} CPT meetdata ophalen...".format(
                            len(cpt_data_list)))
                    update_ui()

                    detailed_list = []
                    for i, basic_cpt in enumerate(cpt_data_list):
                        log("CPT detail ophalen: {0}".format(
                            basic_cpt.bro_id))
                        try:
                            detail = client.get_cpt(basic_cpt.bro_id)
                        except Exception as e:
                            log("CPT detail FOUT {0}: {1}".format(
                                basic_cpt.bro_id, e))
                            detail = None

                        if detail and detail.metingen:
                            log("CPT detail OK: {0} metingen".format(
                                len(detail.metingen)))
                            # Gebruik locatie uit search als detail
                            # geen locatie heeft
                            if detail.rd_x == 0.0 and detail.rd_y == 0.0:
                                detail.rd_x = basic_cpt.rd_x
                                detail.rd_y = basic_cpt.rd_y
                            if detail.nap_hoogte == 0.0:
                                detail.nap_hoogte = basic_cpt.nap_hoogte
                            detailed_list.append(detail)
                        else:
                            log("CPT detail fallback: {0} (detail={1}, metingen={2})".format(
                                basic_cpt.bro_id,
                                detail is not None,
                                len(detail.metingen) if detail else "N/A"))
                            # Fallback: basic data zonder metingen
                            detailed_list.append(basic_cpt)

                        if (i + 1) % 5 == 0:
                            show_progress(self,
                                "CPT meetdata ophalen... {0}/{1}".format(
                                    i + 1, len(cpt_data_list)))
                            update_ui()

                    show_progress(self,
                        "{0} CPT sonderingen tekenen...".format(
                            len(detailed_list)))
                    update_ui()
                    stats["cpt"] = self._draw_cpt(
                        detailed_list, diameter)

            # BHR laden
            if laden_bhr:
                show_progress(self, "Boringen zoeken in BRO...")
                update_ui()

                bhr_ids = client.search_bhr(center_x, center_y, straal)
                log("BHR gevonden: {0}".format(len(bhr_ids)))

                if bhr_ids:
                    show_progress(self, "{0} boringen ophalen...".format(
                        len(bhr_ids)))
                    update_ui()

                    bhr_data_list = []
                    for i, bro_id in enumerate(bhr_ids):
                        data = client.get_bhr(bro_id)
                        if data and data.grondlagen:
                            bhr_data_list.append(data)
                        if (i + 1) % 5 == 0:
                            show_progress(self,
                                "Boringen ophalen... {0}/{1}".format(
                                    i + 1, len(bhr_ids)))
                            update_ui()

                    if bhr_data_list:
                        show_progress(self,
                            "{0} boringen tekenen...".format(
                                len(bhr_data_list)))
                        update_ui()
                        bhr_count, laag_count = self._draw_bhr(
                            bhr_data_list, diameter)
                        stats["bhr"] = bhr_count
                        stats["lagen"] = laag_count

            self.result_count = stats["cpt"] + stats["bhr"]
            hide_progress(self)
            self.DialogResult = True
            self.Close()

            self._show_result(stats, straal)

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            hide_progress(self)
            self.btn_execute.IsEnabled = True

    def _get_straal(self):
        item = self.cmb_straal.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 250

    def _get_diameter(self):
        item = self.cmb_diameter.SelectedItem
        if item and hasattr(item, 'Tag'):
            return float(item.Tag)
        return 1.0

    def _draw_cpt(self, cpt_list, diameter):
        """Teken CPT sonderingen als cilinders in Revit.

        Met meetdata: gesegmenteerde cilinders gekleurd op qc-waarde.
        Zonder meetdata (fallback): één blauwe cilinder.

        Args:
            cpt_list: Lijst van CPTData objecten (met of zonder metingen)
            diameter: Cilinder diameter in meters

        Returns:
            Aantal getekende sonderingen
        """
        origin_x, origin_y = self.location_rd
        radius_m = diameter / 2.0
        count = 0
        view = self.doc.ActiveView

        t = DB.Transaction(self.doc, "GIS2BIM - CPT Sonderingen")
        t.Start()

        try:
            # Parameters en materiaal voorbereiden
            ensure_bro_parameters(self.doc)
            cpt_mat_id = get_or_create_material(
                self.doc, "BRO - CPT Sondering", CPT_KLEUR)

            for data in cpt_list:
                try:
                    # Positie berekenen
                    pos = rd_to_revit_xyz(
                        data.rd_x, data.rd_y,
                        origin_x, origin_y, z=0.0
                    )

                    bro_url = BRO_VIEWER_URL.format(
                        bro_id=data.bro_id)

                    if data.metingen:
                        # Gesegmenteerd tekenen per meting
                        for meting in data.metingen:
                            z_bottom = data.nap_hoogte - meting.onderkant
                            height = meting.dikte

                            if height <= 0:
                                continue

                            solid = create_cylinder_solid(
                                pos.X, pos.Y,
                                z_bottom * METER_TO_FEET,
                                radius_m * METER_TO_FEET,
                                height * METER_TO_FEET
                            )

                            ds = create_directshape(
                                self.doc, [solid],
                                name="BRO CPT - {0} ({1})".format(
                                    data.bro_id,
                                    meting.classificatie)
                            )

                            # Materiaal per qc classificatie
                            kleur = get_qc_kleur(meting.qc)
                            mat_naam = "BRO - qc {0}".format(
                                meting.classificatie)
                            mat_id = get_or_create_material(
                                self.doc, mat_naam, kleur)

                            if not set_element_material(
                                    self.doc, ds, mat_id):
                                set_element_color(
                                    self.doc, view, ds.Id, kleur)

                            # Parameters
                            set_element_parameter(
                                ds, "CPT_diepte",
                                data.einddiepte * METER_TO_FEET)
                            set_element_parameter(
                                ds, "CPT_qc", meting.qc)
                            set_element_parameter(
                                ds, "CPT_classificatie",
                                meting.classificatie)
                            set_element_parameter(
                                ds, "BRO_url", bro_url)

                        count += 1

                    else:
                        # Fallback: één blauwe cilinder
                        z_bottom = data.nap_hoogte - data.einddiepte
                        height = data.einddiepte

                        if height <= 0:
                            continue

                        solid = create_cylinder_solid(
                            pos.X, pos.Y,
                            z_bottom * METER_TO_FEET,
                            radius_m * METER_TO_FEET,
                            height * METER_TO_FEET
                        )

                        ds = create_directshape(
                            self.doc, [solid],
                            name="BRO CPT - {0}".format(data.bro_id)
                        )

                        if not set_element_material(
                                self.doc, ds, cpt_mat_id):
                            set_element_color(
                                self.doc, view, ds.Id, CPT_KLEUR)

                        set_element_parameter(
                            ds, "CPT_diepte",
                            data.einddiepte * METER_TO_FEET)
                        set_element_parameter(
                            ds, "BRO_url", bro_url)

                        count += 1

                except Exception as e:
                    log("CPT draw error ({0}): {1}".format(
                        data.bro_id, e))

            t.Commit()
            log("CPT transaction commit: {0} elementen".format(count))

        except Exception as e:
            t.RollBack()
            log("CPT transaction rollback: {0}\n{1}".format(
                e, traceback.format_exc()))
            raise

        return count

    def _draw_bhr(self, bhr_list, diameter):
        """Teken boringen als gestapelde gekleurde cilinders.

        Args:
            bhr_list: Lijst van BHRData objecten
            diameter: Cilinder diameter in meters

        Returns:
            Tuple (aantal_boringen, aantal_lagen)
        """
        origin_x, origin_y = self.location_rd
        radius_m = diameter / 2.0
        bhr_count = 0
        laag_count = 0
        view = self.doc.ActiveView

        t = DB.Transaction(self.doc, "GIS2BIM - Boringen")
        t.Start()

        try:
            # Parameters voorbereiden
            ensure_bro_parameters(self.doc)

            for data in bhr_list:
                try:
                    # Positie berekenen
                    pos = rd_to_revit_xyz(
                        data.rd_x, data.rd_y,
                        origin_x, origin_y, z=0.0
                    )

                    log("BHR {0}: RD({1:.0f},{2:.0f}) -> pos({3:.1f},{4:.1f}), NAP={5:.2f}, {6} lagen".format(
                        data.bro_id, data.rd_x, data.rd_y,
                        pos.X, pos.Y, data.nap_hoogte,
                        len(data.grondlagen)))

                    bro_url = BRO_VIEWER_URL.format(
                        bro_id=data.bro_id)

                    boring_ok = False
                    for laag in data.grondlagen:
                        # Z posities (NAP referentie)
                        z_bottom = data.nap_hoogte - laag.onderkant
                        height = laag.dikte

                        if height <= 0:
                            continue

                        log("  laag: {0:.1f}-{1:.1f}m, z_bottom={2:.2f}m, h={3:.2f}m, {4} ({5})".format(
                            laag.bovenkant, laag.onderkant,
                            z_bottom, height,
                            laag.grondsoort, laag.beschrijving))

                        # Cilinder per grondlaag
                        solid = create_cylinder_solid(
                            pos.X, pos.Y,
                            z_bottom * METER_TO_FEET,
                            radius_m * METER_TO_FEET,
                            height * METER_TO_FEET
                        )

                        ds = create_directshape(
                            self.doc, [solid],
                            name="BRO BHR - {0} ({1})".format(
                                data.bro_id, laag.grondsoort)
                        )

                        # Materiaal per grondsoort
                        kleur = get_grondsoort_kleur(laag.grondsoort)
                        mat_naam = "BRO - {0}".format(
                            laag.grondsoort.capitalize())
                        mat_id = get_or_create_material(
                            self.doc, mat_naam, kleur)

                        if not set_element_material(
                                self.doc, ds, mat_id):
                            set_element_color(
                                self.doc, view, ds.Id, kleur)

                        # Parameters: diepte, materiaal en BRO url
                        set_element_parameter(
                            ds, "boring_diepte",
                            data.einddiepte * METER_TO_FEET)
                        set_element_parameter(
                            ds, "boring_materiaal",
                            laag.grondsoort)
                        set_element_parameter(
                            ds, "BRO_url", bro_url)

                        laag_count += 1
                        boring_ok = True

                    if boring_ok:
                        bhr_count += 1

                except Exception as e:
                    log("BHR draw error ({0}): {1}\n{2}".format(
                        data.bro_id, e, traceback.format_exc()))

            t.Commit()
            log("BHR transaction commit: {0} boringen, {1} lagen".format(
                bhr_count, laag_count))

        except Exception as e:
            t.RollBack()
            log("BHR transaction rollback: {0}\n{1}".format(
                e, traceback.format_exc()))
            raise

        return (bhr_count, laag_count)

    def _show_result(self, stats, straal):
        msg_lines = [
            "Geotechnische data succesvol geladen!",
            "",
            "Zoekstraal: {0} m".format(straal),
        ]

        if stats["cpt"] > 0:
            msg_lines.append("CPT sonderingen: {0}".format(stats["cpt"]))

        if stats["bhr"] > 0:
            msg_lines.append("Boringen: {0} ({1} grondlagen)".format(
                stats["bhr"], stats["lagen"]))

        if stats["cpt"] == 0 and stats["bhr"] == 0:
            msg_lines.append("Geen data gevonden in zoekgebied")

        forms.alert("\n".join(msg_lines),
                    title="GIS2BIM - GEF Sonderingen")


def main():
    clear_log_file(LOG_FILE)
    log("=== GIS2BIM GEF Sonderingen Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = GEFSonderingenWindow(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM GEF Sonderingen Tool beeindigd ===")


if __name__ == "__main__":
    main()
