# -*- coding: utf-8 -*-
"""
WFS Data Laden - GIS2BIM
========================

Laad geo-data van WFS services (PDOK Kadaster, BAG) en
teken deze in Revit als Model Lines en Text Notes.
"""

__title__ = "WFS\nLaden"
__author__ = "OpenAEC Foundation"
__doc__ = "Laad perceelgrenzen, straatnamen en huisnummers van PDOK"

# CLR references voor WPF
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
clr.AddReference('System.Xml')

import System
from System.Windows import Window, Visibility
from System.Windows.Controls import ComboBoxItem

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
from gis2bim.ui.view_setup import populate_view_dropdown, get_selected_view
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

log = get_logger("WFS")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.wfs import WFSClient, WFSLayer
    from gis2bim.api.wfs_layers import (
        KADASTER_PERCELEN,
        KADASTER_STRAATNAMEN,
        BAG_HUISNUMMERS,
        BAG_PAND,
        get_layer
    )
    from gis2bim.revit.geometry import (
        create_model_lines_from_features,
        create_text_notes_from_features,
        create_filled_regions_from_features,
        rd_to_revit_xyz
    )
    from gis2bim.coordinates import create_bbox_rd
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


class WFSWindow(Window):
    """WPF Window voor WFS data laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location', 'cmb_bbox_size',
            'chk_percelen', 'chk_perceelnummers', 'chk_straatnamen',
            'chk_huisnummers', 'chk_panden',
            'cmb_view', 'cmb_line_style', 'cmb_filled_type', 'cmb_text_type',
            'chk_group',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        populate_view_dropdown(self.cmb_view, doc, log=log)
        self._setup_styles()
        self._bind_events()

    def _setup_styles(self):
        """Vul lijnstijl, filled region type en tekst type dropdowns."""
        try:
            cat = self.doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
            line_styles = []
            for subcat in cat.SubCategories:
                try:
                    style = subcat.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                    if style:
                        line_styles.append((subcat.Name, style))
                except Exception:
                    pass

            line_styles.sort(key=lambda x: x[0])
            default_line_idx = 0

            for i, (name, style) in enumerate(line_styles):
                item = ComboBoxItem()
                item.Content = name
                item.Tag = style
                self.cmb_line_style.Items.Add(item)
                if name == "<Thin Lines>":
                    default_line_idx = i

            if self.cmb_line_style.Items.Count > 0:
                self.cmb_line_style.SelectedIndex = default_line_idx

        except Exception as e:
            log("Error lijnstijlen: {0}".format(e))

        try:
            collector = DB.FilteredElementCollector(self.doc)
            filled_types = collector.OfClass(DB.FilledRegionType).ToElements()

            filled_list = []
            for frt in filled_types:
                try:
                    name = DB.Element.Name.__get__(frt)
                    filled_list.append((name, frt))
                except Exception:
                    filled_list.append(("ID={0}".format(frt.Id.IntegerValue), frt))

            filled_list.sort(key=lambda x: x[0])

            for i, (name, frt) in enumerate(filled_list):
                item = ComboBoxItem()
                item.Content = name
                item.Tag = frt
                self.cmb_filled_type.Items.Add(item)

            if self.cmb_filled_type.Items.Count > 0:
                self.cmb_filled_type.SelectedIndex = 0

        except Exception as e:
            log("Error filled region types: {0}".format(e))

        try:
            collector = DB.FilteredElementCollector(self.doc)
            text_types = collector.OfClass(DB.TextNoteType).ToElements()

            text_list = []
            for tt in text_types:
                try:
                    name = DB.Element.Name.__get__(tt)
                    text_list.append((name, tt))
                except Exception:
                    text_list.append(("ID={0}".format(tt.Id.IntegerValue), tt))

            text_list.sort(key=lambda x: x[0])

            for i, (name, tt) in enumerate(text_list):
                item = ComboBoxItem()
                item.Content = name
                item.Tag = tt
                self.cmb_text_type.Items.Add(item)

            if self.cmb_text_type.Items.Count > 0:
                self.cmb_text_type.SelectedIndex = 0

        except Exception as e:
            log("Error tekst types: {0}".format(e))

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
            selected_layers = self._get_selected_layers()
            view = get_selected_view(self.cmb_view, self.doc)

            if not selected_layers:
                self.txt_status.Text = "Selecteer minimaal 1 layer"
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            if view is None:
                self.txt_status.Text = "Selecteer een view voor labels"
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            center_x, center_y = self.location_rd
            bbox = create_bbox_rd(center_x, center_y, bbox_size)
            log("Bbox: {0}".format(bbox))

            show_progress(self, "WFS data ophalen...")

            wfs_results = self._fetch_wfs_data(selected_layers, bbox)

            show_progress(self, "Geometry aanmaken in Revit...")

            stats = self._draw_in_revit(wfs_results, view)

            hide_progress(self)
            self.DialogResult = True
            self.Close()

            self._show_result(stats, bbox_size)

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            hide_progress(self)
            self.btn_execute.IsEnabled = True

    def _get_bbox_size(self):
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 100

    def _get_selected_layers(self):
        layers = []
        if self.chk_percelen.IsChecked:
            layers.append(("percelen", KADASTER_PERCELEN))
        if self.chk_perceelnummers.IsChecked:
            layers.append(("perceelnummers", KADASTER_PERCELEN))
        if self.chk_straatnamen.IsChecked:
            layers.append(("straatnamen", KADASTER_STRAATNAMEN))
        if self.chk_huisnummers.IsChecked:
            layers.append(("huisnummers", BAG_HUISNUMMERS))
        if self.chk_panden.IsChecked:
            layers.append(("panden", BAG_PAND))
        return layers

    def _get_selected_line_style(self):
        item = self.cmb_line_style.SelectedItem
        if item and hasattr(item, 'Tag'):
            return item.Tag
        return None

    def _get_selected_filled_type(self):
        item = self.cmb_filled_type.SelectedItem
        if item and hasattr(item, 'Tag'):
            return item.Tag
        return None

    def _get_selected_text_type(self):
        item = self.cmb_text_type.SelectedItem
        if item and hasattr(item, 'Tag'):
            return item.Tag
        return None

    def _fetch_wfs_data(self, selected_layers, bbox):
        client = WFSClient()
        results = {}
        fetched_layers = set()

        for layer_name, layer_config in selected_layers:
            layer_key = layer_config.layer_name

            if layer_key not in fetched_layers:
                log("Ophalen: {0}".format(layer_config.name))
                features = client.get_features(layer_config, bbox)
                results[layer_key] = features
                fetched_layers.add(layer_key)
                log("  {0} features opgehaald".format(len(features)))

        return results

    def _draw_in_revit(self, wfs_results, view):
        from System.Collections.Generic import List as GenericList

        stats = {
            "lines": 0,
            "texts": 0,
            "regions": 0,
            "features": 0
        }

        model_ids = []
        detail_ids = []

        origin_rd = self.location_rd
        line_style = self._get_selected_line_style()
        filled_type = self._get_selected_filled_type()
        text_type = self._get_selected_text_type()

        log("=== DRAW START ===")
        log("Origin RD: {0}".format(origin_rd))
        log("View: {0} (type={1})".format(view.Name, view.ViewType))

        t = DB.Transaction(self.doc, "GIS2BIM - WFS Data Laden")
        t.Start()

        try:
            if self.chk_percelen.IsChecked:
                kadaster_key = KADASTER_PERCELEN.layer_name
                if kadaster_key in wfs_results:
                    features = wfs_results[kadaster_key]
                    polygon_features = [f for f in features
                                        if f.geometry_type in ["polygon", "multipolygon"]]
                    log("Percelen: {0} polygon features".format(len(polygon_features)))
                    try:
                        ids, feats = create_model_lines_from_features(
                            self.doc, polygon_features, origin_rd,
                            line_style=line_style
                        )
                        stats["lines"] += len(ids)
                        stats["features"] += feats
                        model_ids.extend(ids)
                        log("Perceelgrenzen: {0} lijnen".format(len(ids)))
                    except Exception as e:
                        log("ERROR perceelgrenzen: {0}".format(e))
                        log(traceback.format_exc())

            if self.chk_perceelnummers.IsChecked:
                kadaster_key = KADASTER_PERCELEN.layer_name
                if kadaster_key in wfs_results:
                    features = wfs_results[kadaster_key]
                    label_features = [f for f in features if f.label_position is not None]
                    log("Perceelnummers: {0} met label_position".format(len(label_features)))
                    try:
                        ids, feats = create_text_notes_from_features(
                            self.doc, view, label_features, origin_rd,
                            text_type=text_type
                        )
                        stats["texts"] += len(ids)
                        detail_ids.extend(ids)
                        log("Perceelnummers: {0} labels".format(len(ids)))
                    except Exception as e:
                        log("ERROR perceelnummers: {0}".format(e))
                        log(traceback.format_exc())

            if self.chk_straatnamen.IsChecked:
                straat_key = KADASTER_STRAATNAMEN.layer_name
                if straat_key in wfs_results:
                    features = wfs_results[straat_key]
                    for f in features:
                        if not f.label and "tekst" in f.properties:
                            f.label = f.properties["tekst"]
                        if f.geometry_type == "point" and f.geometry:
                            f.label_position = f.geometry
                    log("Straatnamen: {0} features".format(len(features)))
                    try:
                        ids, feats = create_text_notes_from_features(
                            self.doc, view, features, origin_rd,
                            text_type=text_type
                        )
                        stats["texts"] += len(ids)
                        detail_ids.extend(ids)
                        log("Straatnamen: {0} labels".format(len(ids)))
                    except Exception as e:
                        log("ERROR straatnamen: {0}".format(e))
                        log(traceback.format_exc())

            if self.chk_huisnummers.IsChecked:
                bag_key = BAG_HUISNUMMERS.layer_name
                if bag_key in wfs_results:
                    features = wfs_results[bag_key]
                    for f in features:
                        huisnr = f.properties.get("huisnummer", "")
                        huisletter = f.properties.get("huisletter", "") or ""
                        toevoeging = f.properties.get("toevoeging", "") or ""
                        label = str(huisnr)
                        if huisletter:
                            label += huisletter
                        if toevoeging:
                            label += "-" + str(toevoeging)
                        f.label = label
                        if f.geometry_type == "point" and f.geometry:
                            f.label_position = f.geometry
                    log("Huisnummers: {0} features".format(len(features)))
                    try:
                        ids, feats = create_text_notes_from_features(
                            self.doc, view, features, origin_rd,
                            text_type=text_type
                        )
                        stats["texts"] += len(ids)
                        detail_ids.extend(ids)
                        log("Huisnummers: {0} labels".format(len(ids)))
                    except Exception as e:
                        log("ERROR huisnummers: {0}".format(e))
                        log(traceback.format_exc())

            if self.chk_panden.IsChecked:
                pand_key = BAG_PAND.layer_name
                if pand_key in wfs_results:
                    features = wfs_results[pand_key]
                    polygon_features = [f for f in features
                                        if f.geometry_type in ["polygon", "multipolygon"]]
                    log("Panden: {0} polygon features".format(len(polygon_features)))
                    try:
                        ids, feats = create_filled_regions_from_features(
                            self.doc, view, polygon_features, origin_rd,
                            filled_type=filled_type
                        )
                        stats["regions"] += len(ids)
                        stats["features"] += feats
                        detail_ids.extend(ids)
                        log("Panden: {0} filled regions".format(len(ids)))
                    except Exception as e:
                        log("ERROR panden: {0}".format(e))
                        log(traceback.format_exc())

            if self.chk_group.IsChecked:
                self._group_elements(model_ids, detail_ids)

            t.Commit()
            log("=== DRAW COMMIT OK ===")

        except Exception as e:
            t.RollBack()
            log("=== DRAW ROLLBACK: {0} ===".format(e))
            raise e

        log("Stats: {0}".format(stats))
        return stats

    def _group_elements(self, model_ids, detail_ids):
        from System.Collections.Generic import List as GenericList

        if model_ids:
            try:
                id_list = GenericList[DB.ElementId]()
                for eid in model_ids:
                    id_list.Add(eid)
                group = self.doc.Create.NewGroup(id_list)
                log("3D groep aangemaakt: {0} elementen".format(len(model_ids)))
            except Exception as e:
                log("ERROR 3D groep: {0}".format(e))

        if detail_ids:
            try:
                id_list = GenericList[DB.ElementId]()
                for eid in detail_ids:
                    id_list.Add(eid)
                group = self.doc.Create.NewGroup(id_list)
                log("2D groep aangemaakt: {0} elementen".format(len(detail_ids)))
            except Exception as e:
                log("ERROR 2D groep: {0}".format(e))

    def _show_result(self, stats, bbox_size):
        msg_lines = [
            "WFS data succesvol geladen!",
            "",
            "Zoekgebied: {0} x {0} m".format(bbox_size),
            "",
            "Aangemaakt:",
            "  Model Lines: {0}".format(stats["lines"]),
            "  Filled Regions: {0}".format(stats["regions"]),
            "  Text Notes: {0}".format(stats["texts"]),
        ]

        forms.alert("\n".join(msg_lines), title="GIS2BIM - WFS Resultaat")


def main():
    log("=== GIS2BIM WFS Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = WFSWindow(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM WFS Tool beeindigd ===")


if __name__ == "__main__":
    main()
