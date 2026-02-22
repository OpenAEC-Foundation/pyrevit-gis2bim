# -*- coding: utf-8 -*-
"""
BGT Data Laden - GIS2BIM
========================

Laad BGT (Basisregistratie Grootschalige Topografie) data van PDOK
en teken deze in Revit als Filled Regions en Detail Lines.
"""

__title__ = "BGT"
__author__ = "3BM Bouwkunde"
__doc__ = "Laad BGT topografische data (wegen, water, terrein, gebouwen)"

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
from gis2bim.ui.view_setup import populate_view_dropdown, get_selected_view
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui
from gis2bim.revit.geometry import rd_to_revit_xyz

log, LOG_FILE = create_tool_logger("BGT", __file__)

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.ogc_api import OGCAPIClient
    from gis2bim.api.bgt_layers import (
        BGT_API_URL,
        BGT_LAYERS,
        get_bgt_layer
    )
    from gis2bim.coordinates import create_bbox_rd
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


class BGTWindow(Window):
    """WPF Window voor BGT data laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.result_count = 0

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location', 'cmb_bbox_size',
            'chk_wegdeel', 'chk_ondersteunendwegdeel', 'chk_overbruggingsdeel',
            'chk_begroeidterreindeel', 'chk_onbegroeidterreindeel',
            'chk_waterdeel', 'chk_ondersteunendwaterdeel',
            'chk_pand', 'chk_overigbouwwerk',
            'chk_scheiding_lijn',
            'cmb_view',
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
        if not self.location_rd:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        show_progress(self, "Data ophalen van PDOK...")
        self.btn_execute.IsEnabled = False

        try:
            self._load_bgt_data()
            self.DialogResult = True
            self.Close()
        except Exception as e:
            log("Error loading BGT: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    def _get_selected_layers(self):
        layer_map = {
            "wegdeel": self.chk_wegdeel,
            "ondersteunendwegdeel": self.chk_ondersteunendwegdeel,
            "overbruggingsdeel": self.chk_overbruggingsdeel,
            "begroeidterreindeel": self.chk_begroeidterreindeel,
            "onbegroeidterreindeel": self.chk_onbegroeidterreindeel,
            "waterdeel": self.chk_waterdeel,
            "ondersteunendwaterdeel": self.chk_ondersteunendwaterdeel,
            "pand": self.chk_pand,
            "overigbouwwerk": self.chk_overigbouwwerk,
            "scheiding_lijn": self.chk_scheiding_lijn,
        }

        selected = []
        for layer_id, checkbox in layer_map.items():
            if checkbox.IsChecked:
                selected.append(layer_id)

        return selected

    def _get_bbox_size(self):
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 100

    def _load_bgt_data(self):
        rd_x, rd_y = self.location_rd
        bbox_size = self._get_bbox_size()
        half_size = bbox_size / 2.0

        bbox = (
            rd_x - half_size,
            rd_y - half_size,
            rd_x + half_size,
            rd_y + half_size
        )

        log("BBOX: {0}".format(bbox))

        selected_layers = self._get_selected_layers()
        log("Geselecteerde layers: {0}".format(selected_layers))

        view = get_selected_view(self.cmb_view, self.doc)
        if not view:
            raise ValueError("Geen view geselecteerd")

        log("View: {0}".format(view.Name))

        client = OGCAPIClient(BGT_API_URL)

        all_features = {}
        for layer_id in selected_layers:
            layer = get_bgt_layer(layer_id)
            if not layer:
                continue

            show_progress(self, "Laden: {0}...".format(layer.name))

            features = client.get_features(
                layer.collection_id,
                bbox,
                limit=1000,
                max_features=5000
            )

            log("{0}: {1} features".format(layer_id, len(features)))
            all_features[layer_id] = {
                "layer": layer,
                "features": features
            }

        show_progress(self, "Tekenen in Revit...")
        self._draw_features(view, all_features)

    def _draw_features(self, view, all_features):
        from System.Collections.Generic import List as GenericList

        origin_rd = self.location_rd
        total_regions = 0
        total_lines = 0

        boundary_style = self._get_line_style("bgt_bounderyline")

        with revit.Transaction("GIS2BIM BGT Laden"):
            for layer_id, data in all_features.items():
                layer = data["layer"]
                features = data["features"]

                if not features:
                    continue

                if layer.geometry_type in ("polygon", "multipolygon"):
                    region_count = self._draw_filled_regions(
                        view, features, origin_rd, layer, boundary_style
                    )
                    total_regions += region_count

                elif layer.geometry_type in ("line", "multiline"):
                    count = self._draw_detail_lines(
                        view, features, origin_rd, layer
                    )
                    total_lines += count

        log("Resultaat: {0} filled regions, {1} detail lines".format(
            total_regions, total_lines))
        self.result_count = total_regions + total_lines

    def _draw_filled_regions(self, view, features, origin_rd, layer,
                             boundary_style=None):
        from System.Collections.Generic import List as GenericList

        origin_x, origin_y = origin_rd
        count = 0

        filled_type = self._get_filled_region_type(layer.filled_region_type)
        if not filled_type:
            filled_type = self._get_default_filled_region_type()

        if not filled_type:
            log("Geen FilledRegionType beschikbaar voor {0}".format(layer.name))
            return 0

        holes_found = 0
        holes_created = 0

        for feature in features:
            polygon_ring_sets = self._get_polygon_rings_from_feature(feature)

            for rings in polygon_ring_sets:
                if not rings:
                    continue

                has_holes = len(rings) > 1
                if has_holes:
                    holes_found += 1

                try:
                    loops = GenericList[DB.CurveLoop]()

                    for ring_index, ring in enumerate(rings):
                        if len(ring) < 3:
                            continue

                        is_ccw = self._is_ring_ccw(ring)
                        ordered_ring = ring if is_ccw else list(reversed(ring))

                        curve_loop = self._create_curve_loop(ordered_ring, origin_x, origin_y)
                        if curve_loop is None:
                            continue

                        loops.Add(curve_loop)

                    if loops.Count == 0:
                        continue

                    filled_region = DB.FilledRegion.Create(
                        self.doc, filled_type.Id, view.Id, loops
                    )

                    if boundary_style is not None:
                        DB.FilledRegion.SetLineStyleId(filled_region, boundary_style.Id)

                    count += 1
                    if has_holes and loops.Count > 1:
                        holes_created += 1

                except Exception as e:
                    if count < 3:
                        log("Fout bij filled region ({0} loops): {1}".format(
                            loops.Count if loops else 0, e))
                    if has_holes and len(rings) > 0:
                        try:
                            outer_ring = rings[0]
                            is_ccw = self._is_ring_ccw(outer_ring)
                            ordered = outer_ring if is_ccw else list(reversed(outer_ring))
                            cl = self._create_curve_loop(ordered, origin_x, origin_y)
                            if cl:
                                single = GenericList[DB.CurveLoop]()
                                single.Add(cl)
                                fr = DB.FilledRegion.Create(
                                    self.doc, filled_type.Id, view.Id, single
                                )
                                if boundary_style is not None:
                                    DB.FilledRegion.SetLineStyleId(fr, boundary_style.Id)
                                count += 1
                        except Exception:
                            pass

        if holes_found > 0:
            log("{0}: {1} polygonen met holes gevonden, {2} met holes aangemaakt".format(
                layer.name, holes_found, holes_created))

        log("{0}: {1} filled regions".format(layer.name, count))
        return count

    def _draw_detail_lines(self, view, features, origin_rd, layer):
        origin_x, origin_y = origin_rd
        count = 0

        for feature in features:
            lines = self._get_lines_from_feature(feature)

            for line_coords in lines:
                if len(line_coords) < 2:
                    continue

                for i in range(len(line_coords) - 1):
                    pt1 = line_coords[i]
                    pt2 = line_coords[i + 1]

                    xyz1 = rd_to_revit_xyz(pt1[0], pt1[1], origin_x, origin_y)
                    xyz2 = rd_to_revit_xyz(pt2[0], pt2[1], origin_x, origin_y)

                    dist = xyz1.DistanceTo(xyz2)
                    if dist < 0.003:
                        continue

                    try:
                        line = DB.Line.CreateBound(xyz1, xyz2)
                        detail_line = self.doc.Create.NewDetailCurve(view, line)
                        count += 1
                    except Exception as e:
                        if count < 3:
                            log("Fout bij detail line: {0}".format(e))

        log("{0}: {1} detail lines".format(layer.name, count))
        return count

    def _get_polygon_rings_from_feature(self, feature):
        geom = feature.geometry
        geom_type = feature.geometry_type

        if geom_type == "polygon" and geom:
            return [geom]
        elif geom_type == "multipolygon" and geom:
            return geom
        return []

    def _get_lines_from_feature(self, feature):
        geom = feature.geometry
        geom_type = feature.geometry_type

        if geom_type == "line" and geom:
            return [geom]
        elif geom_type == "multiline" and geom:
            return geom
        return []

    def _is_ring_ccw(self, ring):
        area = 0.0
        n = len(ring)
        for i in range(n):
            x1, y1 = ring[i][0], ring[i][1]
            x2, y2 = ring[(i + 1) % n][0], ring[(i + 1) % n][1]
            area += (x2 - x1) * (y2 + y1)
        return area < 0

    def _create_curve_loop(self, polygon, origin_x, origin_y):
        raw_points = []
        for pt in polygon:
            xyz = rd_to_revit_xyz(pt[0], pt[1], origin_x, origin_y)
            raw_points.append(xyz)

        points = [raw_points[0]]
        for j in range(1, len(raw_points)):
            if raw_points[j].DistanceTo(points[-1]) >= 0.003:
                points.append(raw_points[j])

        if len(points) > 1 and points[0].DistanceTo(points[-1]) < 0.003:
            points = points[:-1]

        if len(points) < 3:
            return None

        curve_loop = DB.CurveLoop()
        for i in range(len(points)):
            p1 = points[i]
            p2 = points[(i + 1) % len(points)]
            line = DB.Line.CreateBound(p1, p2)
            curve_loop.Append(line)

        return curve_loop

    def _get_filled_region_type(self, type_name):
        if not type_name:
            return None
        try:
            collector = DB.FilteredElementCollector(self.doc)
            types = collector.OfClass(DB.FilledRegionType).ToElements()
            for frt in types:
                try:
                    name = DB.Element.Name.__get__(frt)
                    if name == type_name:
                        return frt
                except Exception:
                    pass
        except Exception:
            pass
        return None

    def _get_default_filled_region_type(self):
        try:
            collector = DB.FilteredElementCollector(self.doc)
            types = collector.OfClass(DB.FilledRegionType).ToElements()
            if types:
                return types[0]
        except Exception:
            pass
        return None

    def _get_line_style(self, style_name):
        if not style_name:
            return None
        try:
            from Autodesk.Revit.DB import GraphicsStyle

            collector = DB.FilteredElementCollector(self.doc)
            styles = collector.OfClass(GraphicsStyle).ToElements()
            for style in styles:
                try:
                    if style.Name == style_name:
                        return style
                except Exception:
                    pass
        except Exception as e:
            log("Fout bij ophalen lijnstijl '{0}': {1}".format(style_name, e))
        return None


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

    window = BGTWindow(doc)
    result = window.ShowDialog()

    if result and window.result_count > 0:
        forms.alert(
            "{0} BGT elementen aangemaakt.".format(window.result_count),
            title="GIS2BIM BGT",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
