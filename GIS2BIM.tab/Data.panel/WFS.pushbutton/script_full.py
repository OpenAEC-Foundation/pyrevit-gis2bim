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
from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows import Window, Visibility
from System.Windows.Markup import XamlReader
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

# Logger setup
output = script.get_output()


def log(msg):
    """Log naar pyRevit output window."""
    print("WFS: {0}".format(msg))


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
    from gis2bim.revit.location import get_project_location_rd
    from gis2bim.revit.geometry import (
        create_model_lines_from_features,
        create_text_notes_from_features,
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
        self._load_xaml(xaml_path)

        self._setup_location()
        self._setup_views()
        self._bind_events()

    def _load_xaml(self, xaml_path):
        with open(xaml_path, 'r') as f:
            xaml_content = f.read()

        reader = StringReader(xaml_content)
        xml_reader = SysXmlReader.Create(reader)
        loaded = XamlReader.Load(xml_reader)

        self.Title = loaded.Title
        self.Width = loaded.Width
        self.SizeToContent = loaded.SizeToContent
        self.WindowStartupLocation = loaded.WindowStartupLocation
        self.ResizeMode = loaded.ResizeMode
        self.Background = loaded.Background
        self.Content = loaded.Content

        self._bind_elements(loaded)

    def _bind_elements(self, root):
        elements = [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location', 'cmb_bbox_size',
            'chk_percelen', 'chk_perceelnummers', 'chk_straatnamen',
            'chk_huisnummers', 'chk_panden',
            'cmb_view', 'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ]

        for name in elements:
            element = root.FindName(name)
            if element:
                setattr(self, name, element)

    def _setup_location(self):
        """Haal projectlocatie op en toon in UI."""
        try:
            location = get_project_location_rd(self.doc)

            if location and (location["rd_x"] != 0 or location["rd_y"] != 0):
                self.location_rd = (location["rd_x"], location["rd_y"])
                self.pnl_location.Visibility = Visibility.Visible
                self.pnl_no_location.Visibility = Visibility.Collapsed
                self.txt_location_x.Text = "{0:,.0f} m".format(location["rd_x"])
                self.txt_location_y.Text = "{0:,.0f} m".format(location["rd_y"])
                self.btn_execute.IsEnabled = True
                log("Locatie gevonden: {0}, {1}".format(location["rd_x"], location["rd_y"]))
            else:
                self.pnl_location.Visibility = Visibility.Collapsed
                self.pnl_no_location.Visibility = Visibility.Visible
                self.btn_execute.IsEnabled = False
                log("Geen locatie ingesteld")

        except Exception as e:
            log("Error getting location: {0}".format(e))
            self.pnl_location.Visibility = Visibility.Collapsed
            self.pnl_no_location.Visibility = Visibility.Visible
            self.btn_execute.IsEnabled = False

    def _setup_views(self):
        """Vul view dropdown met beschikbare views."""
        try:
            # Haal alle floor plans en drafting views op
            collector = DB.FilteredElementCollector(self.doc)
            views = collector.OfClass(DB.View).ToElements()

            # Filter op bruikbare views
            usable_views = []
            for view in views:
                if view.IsTemplate:
                    continue
                if view.ViewType in [
                    DB.ViewType.FloorPlan,
                    DB.ViewType.CeilingPlan,
                    DB.ViewType.AreaPlan,
                    DB.ViewType.DraftingView
                ]:
                    usable_views.append(view)

            # Sorteer op naam
            usable_views.sort(key=lambda v: v.Name)

            # Voeg toe aan dropdown
            active_view = self.doc.ActiveView
            selected_index = 0

            for i, view in enumerate(usable_views):
                item = ComboBoxItem()
                item.Content = view.Name
                item.Tag = view.Id
                self.cmb_view.Items.Add(item)

                if view.Id == active_view.Id:
                    selected_index = i

            if self.cmb_view.Items.Count > 0:
                self.cmb_view.SelectedIndex = selected_index

        except Exception as e:
            log("Error setting up views: {0}".format(e))

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_execute(self, sender, args):
        """Voer WFS data ophalen en tekenen uit."""
        log("Execute gestart")

        if self.location_rd is None:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        try:
            # Toon progress
            self.pnl_progress.Visibility = Visibility.Visible
            self.txt_status.Text = ""
            self.btn_execute.IsEnabled = False

            # Haal parameters op
            bbox_size = self._get_bbox_size()
            selected_layers = self._get_selected_layers()
            view = self._get_selected_view()

            if not selected_layers:
                self.txt_status.Text = "Selecteer minimaal 1 layer"
                self.pnl_progress.Visibility = Visibility.Collapsed
                self.btn_execute.IsEnabled = True
                return

            if view is None:
                self.txt_status.Text = "Selecteer een view voor labels"
                self.pnl_progress.Visibility = Visibility.Collapsed
                self.btn_execute.IsEnabled = True
                return

            # Bereken bbox
            center_x, center_y = self.location_rd
            bbox = create_bbox_rd(center_x, center_y, bbox_size)
            log("Bbox: {0}".format(bbox))

            # Haal WFS data op
            self.txt_progress.Text = "WFS data ophalen..."
            self._update_ui()

            wfs_results = self._fetch_wfs_data(selected_layers, bbox)

            # Teken in Revit
            self.txt_progress.Text = "Geometry aanmaken in Revit..."
            self._update_ui()

            stats = self._draw_in_revit(wfs_results, view)

            # Klaar
            self.pnl_progress.Visibility = Visibility.Collapsed
            self.DialogResult = True
            self.Close()

            # Toon resultaat
            self._show_result(stats, bbox_size)

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.pnl_progress.Visibility = Visibility.Collapsed
            self.btn_execute.IsEnabled = True

    def _get_bbox_size(self):
        """Haal geselecteerde bbox grootte op."""
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 100

    def _get_selected_layers(self):
        """Haal geselecteerde layers op."""
        layers = []

        if self.chk_percelen.IsChecked:
            layers.append(("percelen", KADASTER_PERCELEN))
        if self.chk_perceelnummers.IsChecked:
            # Perceelnummers komen uit dezelfde layer als percelen
            layers.append(("perceelnummers", KADASTER_PERCELEN))
        if self.chk_straatnamen.IsChecked:
            layers.append(("straatnamen", KADASTER_STRAATNAMEN))
        if self.chk_huisnummers.IsChecked:
            layers.append(("huisnummers", BAG_HUISNUMMERS))
        if self.chk_panden.IsChecked:
            layers.append(("panden", BAG_PAND))

        return layers

    def _get_selected_view(self):
        """Haal geselecteerde view op."""
        item = self.cmb_view.SelectedItem
        if item and hasattr(item, 'Tag'):
            view_id = item.Tag
            return self.doc.GetElement(view_id)
        return None

    def _update_ui(self):
        """Force UI update."""
        try:
            from System.Windows.Threading import Dispatcher, DispatcherPriority
            Dispatcher.CurrentDispatcher.Invoke(
                System.Action(lambda: None),
                DispatcherPriority.Render
            )
        except Exception:
            pass

    def _fetch_wfs_data(self, selected_layers, bbox):
        """Haal WFS data op voor geselecteerde layers."""
        client = WFSClient()
        results = {}

        # Unieke layers (percelen en perceelnummers delen dezelfde data)
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
        """Teken WFS data in Revit."""
        stats = {
            "lines": 0,
            "texts": 0,
            "features": 0
        }

        origin_rd = self.location_rd

        # Start transactie
        t = DB.Transaction(self.doc, "GIS2BIM - WFS Data Laden")
        t.Start()

        try:
            # Perceelgrenzen (polygonen als model lines)
            if self.chk_percelen.IsChecked:
                kadaster_key = KADASTER_PERCELEN.layer_name
                if kadaster_key in wfs_results:
                    features = wfs_results[kadaster_key]
                    polygon_features = [f for f in features if f.geometry_type in ["polygon", "multipolygon"]]
                    lines, feats = create_model_lines_from_features(
                        self.doc, polygon_features, origin_rd
                    )
                    stats["lines"] += lines
                    stats["features"] += feats
                    log("Perceelgrenzen: {0} lijnen van {1} percelen".format(lines, feats))

            # Perceelnummers (text notes)
            if self.chk_perceelnummers.IsChecked:
                kadaster_key = KADASTER_PERCELEN.layer_name
                if kadaster_key in wfs_results:
                    features = wfs_results[kadaster_key]
                    # Filter features met label positie
                    label_features = [f for f in features if f.label_position is not None]
                    texts, feats = create_text_notes_from_features(
                        self.doc, view, label_features, origin_rd
                    )
                    stats["texts"] += texts
                    log("Perceelnummers: {0} labels".format(texts))

            # Straatnamen (text notes)
            if self.chk_straatnamen.IsChecked:
                straat_key = KADASTER_STRAATNAMEN.layer_name
                if straat_key in wfs_results:
                    features = wfs_results[straat_key]
                    # Voeg labels toe aan features
                    for f in features:
                        if not f.label and "tekst" in f.properties:
                            f.label = f.properties["tekst"]
                        if f.geometry_type == "point" and f.geometry:
                            f.label_position = f.geometry

                    texts, feats = create_text_notes_from_features(
                        self.doc, view, features, origin_rd
                    )
                    stats["texts"] += texts
                    log("Straatnamen: {0} labels".format(texts))

            # Huisnummers (text notes via bag:verblijfsobject)
            if self.chk_huisnummers.IsChecked:
                bag_key = BAG_HUISNUMMERS.layer_name
                if bag_key in wfs_results:
                    features = wfs_results[bag_key]
                    # Bouw volledige huisnummer labels (huisnummer + huisletter + toevoeging)
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

                    texts, feats = create_text_notes_from_features(
                        self.doc, view, features, origin_rd
                    )
                    stats["texts"] += texts
                    log("Huisnummers: {0} labels".format(texts))

            # Panden (polygonen als model lines)
            if self.chk_panden.IsChecked:
                pand_key = BAG_PAND.layer_name
                if pand_key in wfs_results:
                    features = wfs_results[pand_key]
                    polygon_features = [f for f in features if f.geometry_type in ["polygon", "multipolygon"]]
                    lines, feats = create_model_lines_from_features(
                        self.doc, polygon_features, origin_rd
                    )
                    stats["lines"] += lines
                    stats["features"] += feats
                    log("Panden: {0} lijnen van {1} panden".format(lines, feats))

            t.Commit()

        except Exception as e:
            t.RollBack()
            raise e

        return stats

    def _show_result(self, stats, bbox_size):
        """Toon resultaat dialoog."""
        msg_lines = [
            "WFS data succesvol geladen!",
            "",
            "Zoekgebied: {0} x {0} m".format(bbox_size),
            "",
            "Aangemaakt:",
            "  Model Lines: {0}".format(stats["lines"]),
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
