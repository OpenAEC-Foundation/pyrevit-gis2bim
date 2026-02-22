# -*- coding: utf-8 -*-
"""
OSM Data Laden - GIS2BIM
========================

Laad OpenStreetMap data via de Overpass API
en teken deze in Revit als Filled Regions.

Coordinaten flow:
1. Projectlocatie (RD) -> WGS84 bbox voor Overpass query
2. Overpass retourneert WGS84 (lat/lon) coordinaten
3. WGS84 -> RD -> Revit lokaal (feet)

Filled region type naamconventie (conform Dynamo):
    osm_{feature}_{subtype}
    Bijv: osm_building_residential, osm_landuse_forest

Fallback: osm_building -> default type
"""

__title__ = "OSM"
__author__ = "3BM Bouwkunde"
__doc__ = "Laad OpenStreetMap data (gebouwen, landgebruik, natuur, water)"

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
from gis2bim.coordinates import rd_to_wgs84, wgs84_to_rd

log, LOG_FILE = create_tool_logger("OSM", __file__)

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.overpass import OverpassClient
    from gis2bim.api.osm_layers import (
        OSM_LAYERS,
        get_osm_layer
    )
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


class OSMWindow(Window):
    """WPF Window voor OSM data laden."""

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
            'chk_building', 'chk_landuse', 'chk_natural',
            'chk_water', 'chk_amenity', 'chk_leisure',
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

        show_progress(self, "Data ophalen van OpenStreetMap...")
        self.btn_execute.IsEnabled = False

        try:
            self._load_osm_data()
            self.DialogResult = True
            self.Close()
        except Exception as e:
            log("Error loading OSM: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    def _get_selected_layers(self):
        layer_map = {
            "building": self.chk_building,
            "landuse": self.chk_landuse,
            "natural": self.chk_natural,
            "water": self.chk_water,
            "amenity": self.chk_amenity,
            "leisure": self.chk_leisure,
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
        return 1000

    def _rd_bbox_to_wgs84(self, rd_bbox):
        """Converteer RD bounding box naar WGS84 (south, west, north, east)."""
        xmin, ymin, xmax, ymax = rd_bbox

        # Converteer SW en NE hoeken
        lat_sw, lon_sw = rd_to_wgs84(xmin, ymin)
        lat_ne, lon_ne = rd_to_wgs84(xmax, ymax)

        # south, west, north, east
        return (lat_sw, lon_sw, lat_ne, lon_ne)

    def _load_osm_data(self):
        rd_x, rd_y = self.location_rd
        bbox_size = self._get_bbox_size()
        half_size = bbox_size / 2.0

        # RD bounding box
        rd_bbox = (
            rd_x - half_size,
            rd_y - half_size,
            rd_x + half_size,
            rd_y + half_size
        )

        # Converteer naar WGS84 voor Overpass API
        wgs84_bbox = self._rd_bbox_to_wgs84(rd_bbox)

        log("RD BBOX: {0}".format(rd_bbox))
        log("WGS84 BBOX: {0}".format(wgs84_bbox))

        selected_layers = self._get_selected_layers()
        log("Geselecteerde layers: {0}".format(selected_layers))

        view = get_selected_view(self.cmb_view, self.doc)
        if not view:
            raise ValueError("Geen view geselecteerd")

        log("View: {0}".format(view.Name))

        # Verzamel alle layer objecten
        layer_objects = []
        for layer_id in selected_layers:
            layer = get_osm_layer(layer_id)
            if layer:
                layer_objects.append(layer)

        if not layer_objects:
            raise ValueError("Geen layers geselecteerd")

        # Combineer alle query tags in EEN Overpass query
        all_query_tags = []
        for layer in layer_objects:
            for tag in layer.query_tags:
                if tag not in all_query_tags:
                    all_query_tags.append(tag)

        log("Gecombineerde query: {0} tags".format(len(all_query_tags)))

        show_progress(self, "OSM data downloaden...")
        update_ui()

        client = OverpassClient(timeout=180, log_func=log)
        all_raw_features = client.get_features(
            wgs84_bbox,
            all_query_tags,
            as_polygon=True
        )

        log("Totaal features ontvangen: {0}".format(len(all_raw_features)))

        # Categoriseer features per layer op basis van tags
        features_by_layer = {}
        for layer in layer_objects:
            features_by_layer[layer.layer_id] = {
                "layer": layer,
                "features": []
            }

        for feature in all_raw_features:
            layer_id = self._categorize_feature(feature, selected_layers)
            if layer_id and layer_id in features_by_layer:
                features_by_layer[layer_id]["features"].append(feature)

        for layer_id, data in features_by_layer.items():
            log("{0}: {1} features".format(layer_id, len(data["features"])))

        show_progress(self, "Tekenen in Revit...")
        update_ui()
        self._draw_features(view, features_by_layer)

    def _categorize_feature(self, feature, selected_layers):
        """Wijs een feature toe aan de juiste layer op basis van OSM tags.

        Prioriteit: water > building > landuse > natural > amenity > leisure
        (water eerst omdat natural=water naar water layer moet, niet natural)
        """
        tags = feature.tags
        if not tags:
            return None

        # Water check (voor natural, want natural=water hoort bij water)
        if "water" in selected_layers:
            if "water" in tags:
                return "water"
            if tags.get("natural") == "water":
                return "water"
            waterway = tags.get("waterway", "")
            if waterway in ("riverbank", "dock", "canal"):
                return "water"

        # Gebouwen
        if "building" in selected_layers and "building" in tags:
            return "building"

        # Landgebruik
        if "landuse" in selected_layers and "landuse" in tags:
            return "landuse"

        # Natuur (excl. water, al afgehandeld)
        if "natural" in selected_layers and "natural" in tags:
            return "natural"

        # Voorzieningen
        if "amenity" in selected_layers and "amenity" in tags:
            return "amenity"

        # Recreatie
        if "leisure" in selected_layers and "leisure" in tags:
            return "leisure"

        return None

    def _draw_features(self, view, all_features):
        origin_rd = self.location_rd
        total_regions = 0

        boundary_style = self._get_line_style("OSM_boundery")

        with revit.Transaction("GIS2BIM OSM Laden"):
            for layer_id, data in all_features.items():
                layer = data["layer"]
                features = data["features"]

                if not features:
                    continue

                region_count = self._draw_filled_regions(
                    view, features, origin_rd, layer, boundary_style
                )
                total_regions += region_count

        log("Resultaat: {0} filled regions".format(total_regions))
        self.result_count = total_regions

    def _wgs84_to_revit_xyz(self, lat, lon, origin_x, origin_y):
        """Converteer WGS84 lat/lon naar Revit XYZ via RD."""
        rd_x, rd_y = wgs84_to_rd(lat, lon)
        return rd_to_revit_xyz(rd_x, rd_y, origin_x, origin_y)

    def _resolve_filled_region_type(self, layer, feature):
        """Zoek filled region type per subtype, met fallback.

        Volgorde: osm_building_residential -> osm_building -> default
        Conform Dynamo naamconventie: osm_{feature}_{subtype}
        """
        # Probeer subtype-specifiek type (bijv. osm_building_residential)
        if layer.tag_key and feature.tags:
            subtype = feature.tags.get(layer.tag_key, "")
            if subtype and subtype != "yes":
                specific_name = "{0}_{1}".format(
                    layer.filled_region_type, subtype
                )
                frt = self._get_filled_region_type(specific_name)
                if frt:
                    return frt

        # Fallback naar generiek type (bijv. osm_building)
        frt = self._get_filled_region_type(layer.filled_region_type)
        if frt:
            return frt

        # Laatste fallback: default type
        return self._get_default_filled_region_type()

    def _draw_filled_regions(self, view, features, origin_rd, layer,
                             boundary_style=None):
        from System.Collections.Generic import List as GenericList

        origin_x, origin_y = origin_rd
        count = 0

        # Cache voor filled region types (voorkom herhaalde lookups)
        frt_cache = {}
        default_frt = self._get_default_filled_region_type()

        if not default_frt:
            log("Geen FilledRegionType beschikbaar voor {0}".format(layer.name))
            return 0

        holes_found = 0
        holes_created = 0

        for feature in features:
            polygon_ring_sets = self._get_polygon_rings_from_feature(feature)

            # Bepaal filled region type per feature (met subtype)
            subtype_key = ""
            if layer.tag_key and feature.tags:
                subtype_key = feature.tags.get(layer.tag_key, "")

            # Cache lookup
            cache_key = "{0}_{1}".format(layer.layer_id, subtype_key)
            if cache_key not in frt_cache:
                frt_cache[cache_key] = self._resolve_filled_region_type(
                    layer, feature
                )
            filled_type = frt_cache[cache_key] or default_frt

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

                        curve_loop = self._create_curve_loop_wgs84(
                            ordered_ring, origin_x, origin_y
                        )
                        if curve_loop is None:
                            continue

                        loops.Add(curve_loop)

                    if loops.Count == 0:
                        continue

                    filled_region = DB.FilledRegion.Create(
                        self.doc, filled_type.Id, view.Id, loops
                    )

                    if boundary_style is not None:
                        DB.FilledRegion.SetLineStyleId(
                            filled_region, boundary_style.Id
                        )

                    count += 1
                    if has_holes and loops.Count > 1:
                        holes_created += 1

                except Exception as e:
                    if count < 3:
                        log("Fout bij filled region ({0} loops): {1}".format(
                            loops.Count if loops else 0, e))
                    # Fallback: probeer alleen outer ring
                    if has_holes and len(rings) > 0:
                        try:
                            outer_ring = rings[0]
                            is_ccw = self._is_ring_ccw(outer_ring)
                            ordered = outer_ring if is_ccw else list(
                                reversed(outer_ring))
                            cl = self._create_curve_loop_wgs84(
                                ordered, origin_x, origin_y
                            )
                            if cl:
                                single = GenericList[DB.CurveLoop]()
                                single.Add(cl)
                                fr = DB.FilledRegion.Create(
                                    self.doc, filled_type.Id, view.Id, single
                                )
                                if boundary_style is not None:
                                    DB.FilledRegion.SetLineStyleId(
                                        fr, boundary_style.Id
                                    )
                                count += 1
                        except Exception:
                            pass

        if holes_found > 0:
            log("{0}: {1} polygonen met holes, {2} met holes aangemaakt".format(
                layer.name, holes_found, holes_created))

        log("{0}: {1} filled regions".format(layer.name, count))
        return count

    def _get_polygon_rings_from_feature(self, feature):
        """Haal polygon ringen op uit een feature.

        Returns: lijst van ring-sets, elk ring-set is [outer, hole1, hole2, ...]
        """
        geom = feature.geometry
        geom_type = feature.geometry_type

        if geom_type == "polygon" and geom:
            # geom = [ring1, ring2, ...] waar ring1=outer, rest=holes
            return [geom]
        elif geom_type == "multipolygon" and geom:
            return geom
        return []

    def _is_ring_ccw(self, ring):
        """Bepaal of een ring counter-clockwise is (shoelace formule).

        Ring bevat (lat, lon) tuples. We gebruiken lon als X en lat als Y
        voor de orientatie berekening.
        """
        area = 0.0
        n = len(ring)
        for i in range(n):
            # Gebruik lon als x, lat als y
            x1, y1 = ring[i][1], ring[i][0]
            x2, y2 = ring[(i + 1) % n][1], ring[(i + 1) % n][0]
            area += (x2 - x1) * (y2 + y1)
        return area < 0

    def _create_curve_loop_wgs84(self, polygon, origin_x, origin_y):
        """Maak een CurveLoop van WGS84 coordinaten."""
        # Converteer WGS84 naar Revit XYZ
        raw_points = []
        for pt in polygon:
            xyz = self._wgs84_to_revit_xyz(pt[0], pt[1], origin_x, origin_y)
            raw_points.append(xyz)

        # Filter opeenvolgende duplicaten
        points = [raw_points[0]]
        for j in range(1, len(raw_points)):
            if raw_points[j].DistanceTo(points[-1]) >= 0.003:
                points.append(raw_points[j])

        # Verwijder sluitpunt als het al dicht bij eerste punt ligt
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
            log("Fout bij ophalen lijnstijl '{0}': {1}".format(
                style_name, e))
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

    window = OSMWindow(doc)
    result = window.ShowDialog()

    if result and window.result_count > 0:
        forms.alert(
            "{0} OSM elementen aangemaakt.".format(window.result_count),
            title="GIS2BIM OSM",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
