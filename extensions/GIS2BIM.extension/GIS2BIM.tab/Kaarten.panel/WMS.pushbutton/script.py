# -*- coding: utf-8 -*-
"""
WMS Kaarten Laden - GIS2BIM
============================

Laad WMS kaartbeelden (luchtfoto, bestemmingsplan, geluid, etc.)
en plaats deze als rasterafbeeldingen in Revit views.
"""

__title__ = "WMS\nKaarten"
__author__ = "OpenAEC Foundation"
__doc__ = "Laad WMS kaartbeelden (luchtfoto, bestemmingsplan, geluid, etc.)"

# CLR references voor WPF
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
clr.AddReference('System.Xml')

import System
from System.Windows import Window, Visibility, Thickness
from System.Windows.Controls import (
    CheckBox, TextBlock, StackPanel, Expander
)
from System.Windows.Media import SolidColorBrush
from System.Windows.Media import Color as WpfColor

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
from bm_logger import get_logger
from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements
from gis2bim.ui.location_setup import setup_project_location
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

log = get_logger("WMS")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.wms import WMSClient, WMS_LAYERS, get_layers_by_category
    from gis2bim.coordinates import create_bbox_rd
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# Conversieconstante
METER_TO_FEET = 1.0 / 0.3048


class WMSWindow(Window):
    """WPF Window voor WMS kaarten laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.layer_checkboxes = {}
        self.result_count = 0

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location', 'cmb_bbox_size',
            'chk_select_all', 'pnl_layers', 'chk_create_views',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        self._populate_layers()
        self._bind_events()

    def _populate_layers(self):
        """Vul het layers panel dynamisch met categorie-expanders en checkboxes."""
        violet = SolidColorBrush(WpfColor.FromRgb(53, 14, 53))
        text_primary = SolidColorBrush(WpfColor.FromRgb(50, 50, 50))

        default_checked = {"luchtfoto_actueel", "enkelbestemming"}

        for category, layers in get_layers_by_category():
            expander = Expander()
            expander.IsExpanded = True
            expander.Margin = Thickness(0, 0, 0, 4)

            header = TextBlock()
            header.Text = category
            header.FontSize = 11
            header.FontWeight = System.Windows.FontWeights.SemiBold
            header.Foreground = violet
            expander.Header = header

            content_panel = StackPanel()
            content_panel.Margin = Thickness(16, 4, 0, 4)

            for layer in layers:
                chk = CheckBox()
                chk.Content = layer["name"]
                chk.Tag = layer["key"]
                chk.IsChecked = layer["key"] in default_checked
                chk.Margin = Thickness(0, 0, 0, 4)
                chk.FontSize = 13
                chk.Foreground = text_primary
                content_panel.Children.Add(chk)
                self.layer_checkboxes[layer["key"]] = chk

            expander.Content = content_panel
            self.pnl_layers.Children.Add(expander)

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.chk_select_all.Checked += self._on_select_all
        self.chk_select_all.Unchecked += self._on_deselect_all

    def _on_select_all(self, sender, args):
        for chk in self.layer_checkboxes.values():
            chk.IsChecked = True

    def _on_deselect_all(self, sender, args):
        for chk in self.layer_checkboxes.values():
            chk.IsChecked = False

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_execute(self, sender, args):
        """Voer het laden van WMS kaarten uit."""
        log("Execute gestart")

        if self.location_rd is None:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        try:
            selected = self._get_selected_layers()
            if not selected:
                self.txt_status.Text = "Selecteer minimaal 1 kaartlaag"
                return

            show_progress(self, "Voorbereiden...")
            self.btn_execute.IsEnabled = False

            bbox_size = self._get_bbox_size()
            create_views = self.chk_create_views.IsChecked

            center_x, center_y = self.location_rd
            bbox = create_bbox_rd(center_x, center_y, bbox_size)
            log("Bbox: {0}, grootte: {1}m".format(bbox, bbox_size))

            client = WMSClient()
            loaded_count = 0
            errors = []

            t = DB.Transaction(self.doc, "GIS2BIM - WMS Kaarten Laden")
            t.Start()

            try:
                for layer_key in selected:
                    layer = WMS_LAYERS[layer_key]
                    layer_name = layer["name"]
                    view_name = layer["view_name"]

                    show_progress(self, "Laden: {0}...".format(layer_name))

                    view = self._find_view(view_name)
                    if view is None and create_views:
                        view = self._create_floor_plan(view_name)
                        if view:
                            log("View aangemaakt: {0}".format(view_name))
                    elif view is None:
                        errors.append("{0}: view '{1}' niet gevonden".format(
                            layer_name, view_name))
                        continue

                    show_progress(self, "Downloaden: {0}...".format(layer_name))

                    try:
                        image_path = client.download_image(layer, bbox)
                        log("Download OK: {0} -> {1}".format(layer_name, image_path))
                    except Exception as e:
                        errors.append("{0}: download fout - {1}".format(
                            layer_name, str(e)))
                        log("Download fout {0}: {1}".format(layer_name, e))
                        continue

                    self._remove_existing_images(view, view_name)

                    show_progress(self, "Plaatsen: {0}...".format(layer_name))

                    try:
                        self._place_image_in_view(view, image_path, bbox_size)
                        loaded_count += 1
                        log("Geplaatst: {0} in view {1}".format(
                            layer_name, view_name))
                    except Exception as e:
                        errors.append("{0}: plaatsingsfout - {1}".format(
                            layer_name, str(e)))
                        log("Plaatsingsfout {0}: {1}".format(layer_name, e))
                        log(traceback.format_exc())

                    try:
                        if os.path.exists(image_path):
                            os.remove(image_path)
                    except Exception:
                        pass

                t.Commit()
                log("Transactie committed")

            except Exception as e:
                t.RollBack()
                log("Transactie rollback: {0}".format(e))
                log(traceback.format_exc())
                raise

            hide_progress(self)
            self.result_count = loaded_count
            self.DialogResult = True
            self.Close()

            self._show_result(loaded_count, errors)

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            hide_progress(self)
            self.btn_execute.IsEnabled = True

    def _get_selected_layers(self):
        selected = []
        for key, chk in self.layer_checkboxes.items():
            if chk.IsChecked:
                selected.append(key)
        return selected

    def _get_bbox_size(self):
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 500

    def _find_view(self, view_name):
        collector = DB.FilteredElementCollector(self.doc)
        views = collector.OfClass(DB.View).ToElements()

        for view in views:
            if view.IsTemplate:
                continue
            try:
                if view.Name == view_name:
                    return view
            except Exception:
                pass
        return None

    def _create_floor_plan(self, view_name):
        try:
            collector = DB.FilteredElementCollector(self.doc)
            vfts = collector.OfClass(DB.ViewFamilyType).ToElements()

            floor_plan_type = None
            for vft in vfts:
                if vft.ViewFamily == DB.ViewFamily.FloorPlan:
                    floor_plan_type = vft
                    break

            if floor_plan_type is None:
                log("Geen FloorPlan ViewFamilyType gevonden")
                return None

            level_collector = DB.FilteredElementCollector(self.doc)
            levels = level_collector.OfClass(DB.Level).ToElements()

            if not levels:
                log("Geen Levels gevonden")
                return None

            level = levels[0]

            new_view = DB.ViewPlan.Create(
                self.doc, floor_plan_type.Id, level.Id
            )
            new_view.Name = view_name

            log("FloorPlan aangemaakt: {0}".format(view_name))
            return new_view

        except Exception as e:
            log("Fout bij aanmaken view: {0}".format(e))
            log(traceback.format_exc())
            return None

    def _remove_existing_images(self, view, view_name):
        try:
            collector = DB.FilteredElementCollector(self.doc, view.Id)
            images = collector.OfCategory(
                DB.BuiltInCategory.OST_RasterImages
            ).WhereElementIsNotElementType().ToElements()

            for img in images:
                try:
                    self.doc.Delete(img.Id)
                    log("Bestaand beeld verwijderd uit {0}".format(view_name))
                except Exception:
                    pass
        except Exception:
            pass

    def _place_image_in_view(self, view, image_path, bbox_size_m):
        options = DB.ImageTypeOptions(image_path, False, DB.ImageTypeSource.Import)
        image_type = DB.ImageType.Create(self.doc, options)

        placement = DB.ImagePlacementOptions()
        placement.PlacementPoint = DB.BoxPlacement.Center
        placement.Location = DB.XYZ(0, 0, 0)

        image_instance = DB.ImageInstance.Create(
            self.doc, view, image_type.Id, placement
        )

        desired_width_ft = bbox_size_m * METER_TO_FEET

        current_width = image_instance.Width
        if current_width > 0:
            scale_factor = desired_width_ft / current_width
            image_instance.Width = desired_width_ft
            log("Image geschaald: {0:.1f}ft (was {1:.1f}ft, factor {2:.3f})".format(
                desired_width_ft, current_width, scale_factor))

    def _show_result(self, loaded_count, errors):
        msg_lines = [
            "WMS kaarten geladen!",
            "",
            "Kaartlagen geplaatst: {0}".format(loaded_count),
        ]

        if errors:
            msg_lines.append("")
            msg_lines.append("Waarschuwingen:")
            for err in errors:
                msg_lines.append("  - {0}".format(err))

        forms.alert("\n".join(msg_lines), title="GIS2BIM - WMS Kaarten")


def main():
    log("=== GIS2BIM WMS Kaarten Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = WMSWindow(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM WMS Kaarten Tool beeindigd ===")


if __name__ == "__main__":
    main()
