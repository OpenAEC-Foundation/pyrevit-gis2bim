# -*- coding: utf-8 -*-
"""
KLIC Data Laden - GIS2BIM
==========================

Importeer KLIC kabels en leidingen data (IMKL GML formaat) in Revit.
Ondersteunt zowel map als .zip leveringen.
"""

__title__ = "KLIC"
__author__ = "OpenAEC Foundation"
__doc__ = "Importeer KLIC kabels en leidingen data in Revit"

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
from gis2bim.ui.logging_helper import create_tool_logger
from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements
from gis2bim.ui.location_setup import setup_project_location
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui
from gis2bim.revit.geometry import (
    create_model_lines, create_text_notes,
    create_or_get_line_style, KLIC_COLORS,
)

log, LOG_FILE = create_tool_logger("KLIC", __file__)

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.parsers.klic import (
        parse_klic_delivery, KLICDelivery, KLICError,
        FEATURE_TYPE_MAP,
    )
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# Map van netwerk type naar NL display naam
NETWORK_DISPLAY = {
    "electricity": "Elektriciteitskabel",
    "telecom": "Telecommunicatiekabel",
    "gas": "OlieGasChemicalienPijpleiding",
    "water": "Waterleiding",
    "sewer": "Rioolleiding",
    "duct": "Kabelbed / Mantelbuis",
    "other": "Overig",
}


class KLICWindow(Window):
    """WPF Window voor KLIC data laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.delivery = None
        self.result_count = 0

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle',
            'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location',
            'txt_klic_path', 'btn_browse',
            'pnl_klic_info', 'txt_klic_number', 'txt_klic_date',
            'txt_klic_features', 'txt_klic_annotations',
            'pnl_features',
            'chk_electricity', 'txt_electricity',
            'chk_telecom', 'txt_telecom',
            'chk_gas', 'txt_gas',
            'chk_sewer', 'txt_sewer',
            'chk_water', 'txt_water',
            'chk_duct', 'txt_duct',
            'chk_annotations', 'txt_annotations_chk',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        self._bind_events()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.btn_browse.Click += self._on_browse

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_browse(self, sender, args):
        """Selecteer KLIC levering map."""
        self.txt_status.Text = ""

        # Gebruik forms.pick_folder voor map selectie
        folder = forms.pick_folder(title="Selecteer KLIC levering map")
        if not folder:
            return

        self.txt_klic_path.Text = folder
        self._parse_delivery(folder)

    def _parse_delivery(self, path):
        """Parse KLIC levering en toon preview."""
        show_progress(self, "KLIC levering parsen...")
        update_ui()

        try:
            self.delivery = parse_klic_delivery(path)
            log("KLIC levering geparsed: {0}".format(self.delivery))
            log("  Features: {0}, Annotaties: {1}".format(
                len(self.delivery.features), len(self.delivery.annotations)))

            self._show_preview()
            hide_progress(self)

        except KLICError as e:
            log("KLIC parse fout: {0}".format(e))
            hide_progress(self)
            self.delivery = None
            self.btn_execute.IsEnabled = False
            self.pnl_klic_info.Visibility = Visibility.Collapsed
            self.pnl_features.Visibility = Visibility.Collapsed
            self.txt_status.Text = "Fout: {0}".format(str(e))

        except Exception as e:
            log("Onverwachte fout: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.delivery = None
            self.btn_execute.IsEnabled = False
            self.pnl_klic_info.Visibility = Visibility.Collapsed
            self.pnl_features.Visibility = Visibility.Collapsed
            self.txt_status.Text = "Fout bij parsen: {0}".format(str(e))

    def _show_preview(self):
        """Toon preview van geparsede KLIC data."""
        d = self.delivery
        if d is None:
            return

        # Info panel
        self.pnl_klic_info.Visibility = Visibility.Visible
        self.txt_klic_number.Text = d.klic_number or "Onbekend"
        self.txt_klic_date.Text = d.delivery_date or "Onbekend"
        self.txt_klic_features.Text = "{0} kabels/leidingen".format(len(d.features))
        self.txt_klic_annotations.Text = str(len(d.annotations))

        # Feature type checkboxes met aantallen
        summary = d.feature_summary()
        counts_per_network = {}
        for f in d.features:
            nt = f.network_type
            counts_per_network[nt] = counts_per_network.get(nt, 0) + 1

        self._update_checkbox("electricity", counts_per_network,
                              self.chk_electricity, self.txt_electricity)
        self._update_checkbox("telecom", counts_per_network,
                              self.chk_telecom, self.txt_telecom)
        self._update_checkbox("gas", counts_per_network,
                              self.chk_gas, self.txt_gas)
        self._update_checkbox("sewer", counts_per_network,
                              self.chk_sewer, self.txt_sewer)
        self._update_checkbox("water", counts_per_network,
                              self.chk_water, self.txt_water)
        self._update_checkbox("duct", counts_per_network,
                              self.chk_duct, self.txt_duct)

        # Annotaties
        ann_count = len(d.annotations)
        self.txt_annotations_chk.Text = "Annotaties ({0})".format(ann_count)
        self.chk_annotations.IsEnabled = ann_count > 0

        self.pnl_features.Visibility = Visibility.Visible

        # Enable execute als er locatie en features zijn
        if self.location_rd and len(d.features) > 0:
            self.btn_execute.IsEnabled = True

    def _update_checkbox(self, network_type, counts, checkbox, textblock):
        """Update checkbox label met aantal features."""
        count = counts.get(network_type, 0)
        display_name = NETWORK_DISPLAY.get(network_type, network_type)
        textblock.Text = "{0} ({1})".format(display_name, count)
        checkbox.IsEnabled = count > 0
        if count == 0:
            checkbox.IsChecked = False

    def _on_execute(self, sender, args):
        """Start KLIC data importeren in Revit."""
        if not self.location_rd:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        if not self.delivery:
            self.txt_status.Text = "Geen KLIC data geladen"
            return

        show_progress(self, "KLIC data importeren...")
        self.btn_execute.IsEnabled = False

        try:
            self._import_klic_data()
            self.DialogResult = True
            self.Close()
        except Exception as e:
            log("Error importing KLIC: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    def _get_selected_network_types(self):
        """Geeft set van geselecteerde netwerk types."""
        type_map = {
            "electricity": self.chk_electricity,
            "telecom": self.chk_telecom,
            "gas": self.chk_gas,
            "sewer": self.chk_sewer,
            "water": self.chk_water,
            "duct": self.chk_duct,
        }

        selected = set()
        for network_type, checkbox in type_map.items():
            if checkbox.IsChecked:
                selected.add(network_type)

        # "other" altijd meenemen als er andere types geselecteerd zijn
        if selected:
            selected.add("other")

        return selected

    def _import_klic_data(self):
        """Importeer KLIC features als Model Lines in Revit."""
        from Autodesk.Revit.DB import Group

        origin_rd = self.location_rd
        selected_types = self._get_selected_network_types()
        import_annotations = self.chk_annotations.IsChecked

        log("Importeren - types: {0}, annotaties: {1}".format(
            selected_types, import_annotations))

        # Filter features op geselecteerde types
        features_by_type = {}
        for f in self.delivery.features:
            if f.network_type in selected_types:
                if f.network_type not in features_by_type:
                    features_by_type[f.network_type] = []
                features_by_type[f.network_type].append(f)

        total_lines = 0
        total_annotations = 0

        with revit.Transaction("GIS2BIM - KLIC Laden"):
            # Maak lijnstijlen en teken features per netwerk type
            for network_type, features in features_by_type.items():
                display_name = NETWORK_DISPLAY.get(network_type, network_type)
                show_progress(self, "Laden: {0} ({1})...".format(
                    display_name, len(features)))
                update_ui()

                # Maak of haal lijnstijl
                style_name = "KLIC_{0}".format(display_name.replace(" / ", "_"))
                rgb = KLIC_COLORS.get(network_type, (200, 200, 200))
                line_style = create_or_get_line_style(
                    self.doc, style_name, rgb, line_weight=3)

                if line_style:
                    log("Lijnstijl: {0} ({1})".format(style_name, rgb))
                else:
                    log("WAARSCHUWING: Kon lijnstijl '{0}' niet aanmaken".format(style_name))

                # Verzamel polylines
                polylines = []
                for f in features:
                    if f.geometry_type == "line" and f.geometry:
                        polylines.append(f.geometry)
                    elif f.geometry_type == "polygon" and f.geometry:
                        polylines.append(f.geometry)

                if polylines:
                    ids = create_model_lines(
                        self.doc, polylines, origin_rd,
                        line_style=line_style)
                    line_count = len(ids)
                    total_lines += line_count
                    log("{0}: {1} lijnsegmenten".format(display_name, line_count))

                    # Groepeer elementen
                    if ids:
                        self._group_elements(ids, "KLIC_{0}".format(display_name))

            # Annotaties (labels als TextNotes)
            if import_annotations and self.delivery.annotations:
                show_progress(self, "Laden: Annotaties ({0})...".format(
                    len(self.delivery.annotations)))
                update_ui()

                ann_count = self._import_annotations(origin_rd)
                total_annotations = ann_count
                log("Annotaties: {0} text notes".format(ann_count))

        self.result_count = total_lines + total_annotations
        log("Totaal: {0} lijnsegmenten, {1} annotaties".format(
            total_lines, total_annotations))

    def _import_annotations(self, origin_rd):
        """Importeer annotaties als TextNotes in de actieve view."""
        view = self.doc.ActiveView

        # Filter annotaties met labels (punt-annotaties)
        annotations = []
        for ann in self.delivery.annotations:
            if ann.label and ann.label_position:
                annotations.append({
                    "text": ann.label,
                    "x": ann.label_position[0],
                    "y": ann.label_position[1],
                    "rotation": ann.label_rotation,
                })

        if not annotations:
            return 0

        ids = create_text_notes(self.doc, view, annotations, origin_rd)
        return len(ids)

    def _group_elements(self, element_ids, group_name):
        """Groepeer elementen in een Revit Group."""
        if not element_ids:
            return

        try:
            from System.Collections.Generic import List as GenericList

            id_list = GenericList[DB.ElementId]()
            for eid in element_ids:
                id_list.Add(eid)

            group = self.doc.Create.NewGroup(id_list)
            try:
                # Hernoem de groep
                group_type = self.doc.GetElement(group.GetTypeId())
                if group_type:
                    DB.Element.Name.__set__(group_type, group_name)
            except Exception:
                # Naam bestaat mogelijk al, geen probleem
                pass

        except Exception as e:
            log("Kon elementen niet groeperen als '{0}': {1}".format(
                group_name, e))


def main():
    doc = revit.doc

    if not GIS2BIM_LOADED:
        forms.alert(
            "GIS2BIM modules konden niet worden geladen.\n\n"
            "Fout: {0}".format(IMPORT_ERROR),
            title="GIS2BIM Error",
            warn_icon=True
        )
        return

    window = KLICWindow(doc)
    result = window.ShowDialog()

    if result and window.result_count > 0:
        forms.alert(
            "{0} KLIC elementen aangemaakt.".format(window.result_count),
            title="GIS2BIM KLIC",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
