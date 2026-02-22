# -*- coding: utf-8 -*-
"""
BGT 3D Elementen - GIS2BIM
============================

Plaats bomen en lantaarnpalen als Revit family instances op basis van
BGT punt-data via de PDOK OGC API Features.

Collecties:
- vegetatieobject: bomen (punt-geometrie)
- paal: lantaarnpalen (punt-geometrie, filter op plus-type == "lichtmast")

Families:
- Bomen: loofboom (generic)
- Lantaarns: lichtmast (generic)
"""

__title__ = "BGT\n3D"
__author__ = "OpenAEC Foundation"
__doc__ = "BGT 3D elementen plaatsen (bomen en lantaarnpalen) op basis van PDOK OGC API"

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
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

log, LOG_FILE = create_tool_logger("BGT3D", __file__)

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.ogc_api import OGCAPIClient
    from gis2bim.api.bgt_layers import BGT_API_URL, BGT_VEGETATIEOBJECT, BGT_PAAL
    from gis2bim.revit.geometry import rd_to_revit_xyz
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# Standaard family namen (worden voorgeselecteerd als ze bestaan)
DEFAULT_FAMILY_BOMEN = "loofboom"
DEFAULT_FAMILY_LANTAARNS = "lichtmast"


def _get_all_family_symbols(doc):
    """Verzamel alle FamilySymbols gegroepeerd per Family naam.

    Returns:
        Dict van {family_name: eerste FamilySymbol}
    """
    families = {}
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
    for fs in collector:
        name = fs.Family.Name
        if name not in families:
            families[name] = fs
    return families


def _centroid_of_ring(ring):
    """Bereken centroid van een polygon ring (lijst van (x,y) tuples)."""
    if not ring:
        return None
    n = len(ring)
    cx = sum(p[0] for p in ring) / n
    cy = sum(p[1] for p in ring) / n
    return (cx, cy)


def _extract_points_from_features(features):
    """Extraheer (x, y) punten uit OGCAPIFeature objecten.

    Ondersteunt point, multipoint en polygon geometrie (centroid).

    Returns:
        Lijst van (x, y) tuples
    """
    points = []
    for feature in features:
        if feature.geometry_type == "point":
            # geometry is een tuple (x, y)
            points.append(feature.geometry)
        elif feature.geometry_type == "multipoint":
            # geometry is een lijst van (x, y) tuples
            for pt in feature.geometry:
                points.append(pt)
        elif feature.geometry_type in ("polygon", "multipolygon"):
            # Sommige vegetatieobjecten hebben vlakgeometrie;
            # gebruik centroid van de outer ring als plaatsingspunt
            if feature.geometry_type == "polygon":
                # geometry = [[outer_ring], [hole1], ...]
                if feature.geometry and feature.geometry[0]:
                    centroid = _centroid_of_ring(feature.geometry[0])
                    if centroid:
                        points.append(centroid)
            else:
                # multipolygon: geometry = [[[outer1], ...], [[outer2], ...]]
                for polygon_rings in feature.geometry:
                    if polygon_rings and polygon_rings[0]:
                        centroid = _centroid_of_ring(polygon_rings[0])
                        if centroid:
                            points.append(centroid)
    return points


class BGT3DWindow(Window):
    """WPF Window voor BGT 3D elementen plaatsing."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self._family_map = {}

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location', 'cmb_search_radius',
            'chk_bomen', 'chk_lantaarns',
            'cmb_family_bomen', 'cmb_family_lantaarns', 'txt_family_warning',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        self._populate_family_dropdowns()
        self._bind_events()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute

    def _populate_family_dropdowns(self):
        """Vul family dropdowns met alle families uit het document."""
        from System.Windows.Controls import ComboBoxItem

        all_families = _get_all_family_symbols(self.doc)
        family_names = sorted(all_families.keys())
        log("Families in document: {0}".format(len(family_names)))

        # Bewaar mapping voor later opvragen
        self._family_map = all_families

        # Vul beide dropdowns
        for combo, default_name in [
            (self.cmb_family_bomen, DEFAULT_FAMILY_BOMEN),
            (self.cmb_family_lantaarns, DEFAULT_FAMILY_LANTAARNS),
        ]:
            combo.Items.Clear()
            default_index = -1

            for i, name in enumerate(family_names):
                item = ComboBoxItem()
                item.Content = name
                item.Tag = name
                combo.Items.Add(item)

                if name == default_name:
                    default_index = i

            if default_index >= 0:
                combo.SelectedIndex = default_index
                log("Family '{0}' voorgeselecteerd".format(default_name))
            elif combo.Items.Count > 0:
                combo.SelectedIndex = 0

        # Waarschuwing als standaard families niet gevonden
        warnings = []
        if DEFAULT_FAMILY_BOMEN not in all_families:
            warnings.append("'{0}' niet in project".format(DEFAULT_FAMILY_BOMEN))
        if DEFAULT_FAMILY_LANTAARNS not in all_families:
            warnings.append("'{0}' niet in project".format(DEFAULT_FAMILY_LANTAARNS))

        if warnings:
            self.txt_family_warning.Text = "Standaard families niet gevonden: {0}".format(
                ", ".join(warnings))
        else:
            self.txt_family_warning.Text = ""

        if not family_names:
            self.txt_family_warning.Text = "Geen families geladen in het project"

    def _get_selected_family_symbol(self, combo):
        """Haal de geselecteerde FamilySymbol op uit een ComboBox.

        Returns:
            FamilySymbol of None
        """
        item = combo.SelectedItem
        if item and hasattr(item, 'Tag'):
            family_name = item.Tag
            return self._family_map.get(family_name)
        return None

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_execute(self, sender, args):
        log("Execute gestart")

        if self.location_rd is None:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        plaats_bomen = self.chk_bomen.IsChecked
        plaats_lantaarns = self.chk_lantaarns.IsChecked

        if not plaats_bomen and not plaats_lantaarns:
            self.txt_status.Text = "Selecteer minimaal een element type"
            return

        # Haal geselecteerde families op
        family_bomen = None
        family_lantaarns = None

        if plaats_bomen:
            family_bomen = self._get_selected_family_symbol(self.cmb_family_bomen)
            if family_bomen is None:
                self.txt_status.Text = "Selecteer een family voor bomen"
                return

        if plaats_lantaarns:
            family_lantaarns = self._get_selected_family_symbol(self.cmb_family_lantaarns)
            if family_lantaarns is None:
                self.txt_status.Text = "Selecteer een family voor lantaarnpalen"
                return

        try:
            show_progress(self, "Voorbereiden...")
            self.btn_execute.IsEnabled = False

            search_radius = self._get_search_radius()
            center_x, center_y = self.location_rd
            bbox = (
                center_x - search_radius,
                center_y - search_radius,
                center_x + search_radius,
                center_y + search_radius
            )

            log("Center: {0}, {1} - Zoekstraal: {2}m".format(
                center_x, center_y, search_radius))
            log("BBox: {0}".format(bbox))

            client = OGCAPIClient(BGT_API_URL)
            bomen_count = 0
            lantaarns_count = 0

            # Bomen ophalen en plaatsen
            if plaats_bomen:
                show_progress(self, "Bomen ophalen van PDOK...")
                update_ui()

                bomen_features = client.get_features(
                    BGT_VEGETATIEOBJECT.collection_id, bbox)
                bomen_points = _extract_points_from_features(bomen_features)
                log("Bomen features: {0}, punten: {1}".format(
                    len(bomen_features), len(bomen_points)))

                if bomen_points:
                    show_progress(self, "Bomen plaatsen ({0} stuks)...".format(
                        len(bomen_points)))
                    update_ui()
                    bomen_count = self._place_instances(
                        bomen_points, family_bomen,
                        center_x, center_y, "Bomen")

            # Lantaarns ophalen en plaatsen
            if plaats_lantaarns:
                show_progress(self, "Lantaarnpalen ophalen van PDOK...")
                update_ui()

                paal_features = client.get_features(
                    BGT_PAAL.collection_id, bbox)

                # Filter op lichtmast
                lantaarn_features = [
                    f for f in paal_features
                    if f.properties.get("plus_type") == "lichtmast"
                ]
                lantaarn_points = _extract_points_from_features(lantaarn_features)
                log("Paal features: {0}, lichtmasten: {1}, punten: {2}".format(
                    len(paal_features), len(lantaarn_features), len(lantaarn_points)))

                if lantaarn_points:
                    show_progress(self, "Lantaarnpalen plaatsen ({0} stuks)...".format(
                        len(lantaarn_points)))
                    update_ui()
                    lantaarns_count = self._place_instances(
                        lantaarn_points, family_lantaarns,
                        center_x, center_y, "Lantaarnpalen")

            hide_progress(self)
            self.DialogResult = True
            self.Close()

            self._show_result(search_radius, bomen_count, lantaarns_count)

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            hide_progress(self)
            self.btn_execute.IsEnabled = True

    def _get_search_radius(self):
        item = self.cmb_search_radius.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 100

    def _place_instances(self, points, family_symbol, origin_x, origin_y, label):
        """Plaats family instances op de gegeven punten.

        Args:
            points: Lijst van (rd_x, rd_y) tuples
            family_symbol: FamilySymbol om te plaatsen
            origin_x, origin_y: RD origin van het project
            label: Naam voor logging

        Returns:
            Aantal geplaatste instances
        """
        t = DB.Transaction(self.doc, "GIS2BIM - BGT 3D {0}".format(label))
        t.Start()
        try:
            # Activeer family symbol indien nodig
            if not family_symbol.IsActive:
                family_symbol.Activate()
                self.doc.Regenerate()

            placed = 0
            errors = 0

            for pt in points:
                try:
                    xyz = rd_to_revit_xyz(pt[0], pt[1], origin_x, origin_y)
                    self.doc.Create.NewFamilyInstance(
                        xyz, family_symbol,
                        DB.Structure.StructuralType.NonStructural
                    )
                    placed += 1
                except Exception as e:
                    errors += 1
                    if errors <= 3:
                        log("Fout bij plaatsen {0}: {1}".format(label, e))

            t.Commit()
            log("{0}: {1} geplaatst, {2} fouten".format(label, placed, errors))
            return placed

        except Exception as e:
            t.RollBack()
            log("Fout bij {0} transactie: {1}".format(label, e))
            log(traceback.format_exc())
            return 0

    def _show_result(self, search_radius, bomen_count, lantaarns_count):
        """Toon resultaat dialog."""
        msg_lines = [
            "BGT 3D elementen geplaatst!",
            "",
            "Zoekgebied: {0} m straal".format(search_radius),
        ]

        if bomen_count > 0:
            msg_lines.append("Bomen geplaatst: {0}".format(bomen_count))
        if lantaarns_count > 0:
            msg_lines.append("Lantaarnpalen geplaatst: {0}".format(lantaarns_count))

        total = bomen_count + lantaarns_count
        if total == 0:
            msg_lines.append("")
            msg_lines.append("Geen elementen gevonden binnen zoekgebied.")

        forms.alert("\n".join(msg_lines), title="GIS2BIM - BGT 3D")


def main():
    clear_log_file(LOG_FILE)
    log("=== GIS2BIM BGT 3D Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = BGT3DWindow(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM BGT 3D Tool beeindigd ===")


if __name__ == "__main__":
    main()
