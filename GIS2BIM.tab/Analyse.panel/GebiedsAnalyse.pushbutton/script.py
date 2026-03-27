# -*- coding: utf-8 -*-
"""
GebiedsAnalyse - GIS2BIM
=========================

Analyseer de voorzieningenwaarde van de projectlocatie op basis van
OpenStreetMap data en visualiseer het resultaat als heatmap in Revit.

Workflow:
    1. Locatie ophalen uit project parameters
    2. Instellingen UI (categorieen, gewichten, grid)
    3. OSM data ophalen via Overpass API
    4. Grid scores berekenen
    5. Heatmap renderen via SpatialFieldManager
    6. Samenvatting tonen
"""

__title__ = "Gebiedsanalyse"
__author__ = "3BM Bouwkunde"
__doc__ = "Voorzieningenwaarde heatmap op basis van OSM data"

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
import json
import time
import traceback
import copy

# Voeg lib folder toe aan path
extension_path = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(__file__))))
lib_path = os.path.join(extension_path, "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Logger
from bm_logger import get_logger
log = get_logger("GebiedsAnalyse")

# GIS2BIM UI helpers
from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements
from gis2bim.ui.location_setup import setup_project_location
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.coordinates import rd_to_wgs84, wgs84_to_rd
    from gis2bim.bbox import BoundingBox
    from gis2bim.api.overpass import OverpassClient
    from gis2bim.analysis.categories import (
        CATEGORIES, PRESETS, RING_PROFILES,
        get_max_ring, get_all_osm_tags, apply_preset, apply_ring_profile
    )
    from gis2bim.analysis.grid import generate_grid, calculate_scores, smooth_scores
    from gis2bim.analysis.heatmap import create_and_apply_heatmap
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# ─── Cache helpers ───────────────────────────────────────────────

CACHE_DIR = os.path.join(
    os.environ.get("APPDATA", ""),
    "3BM_Bouwkunde", "osm_cache"
)
CACHE_MAX_AGE = 7 * 24 * 3600  # 7 dagen in seconden


def _get_cache_path(lat, lon, radius, category):
    """Genereer cache bestandspad."""
    key = "{0:.4f}_{1:.4f}_{2}_{3}.json".format(lat, lon, radius, category)
    return os.path.join(CACHE_DIR, key)


def _read_cache(path):
    """Lees gecachte POIs als het bestand niet verlopen is."""
    try:
        if not os.path.exists(path):
            return None
        age = time.time() - os.path.getmtime(path)
        if age > CACHE_MAX_AGE:
            return None
        with open(path, "r") as f:
            return json.load(f)
    except Exception:
        return None


def _write_cache(path, data):
    """Schrijf POIs naar cache."""
    try:
        cache_dir = os.path.dirname(path)
        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)
        with open(path, "w") as f:
            json.dump(data, f)
    except Exception:
        pass


# ─── Window ──────────────────────────────────────────────────────

class GebiedsAnalyseWindow(Window):
    """WPF Window voor gebiedsanalyse instellingen."""

    # Mapping categorie-id -> (checkbox, rings combobox, weight textbox)
    CAT_UI_MAP = {
        "school":     ("chk_school",     "cmb_rings_school",     "txt_weight_school"),
        "bus":        ("chk_bus",        "cmb_rings_bus",        "txt_weight_bus"),
        "trein":      ("chk_trein",      "cmb_rings_trein",      "txt_weight_trein"),
        "park":       ("chk_park",       "cmb_rings_park",       "txt_weight_park"),
        "supermarkt": ("chk_supermarkt", "cmb_rings_supermarkt", "txt_weight_supermarkt"),
        "ziekenhuis": ("chk_ziekenhuis", "cmb_rings_ziekenhuis", "txt_weight_ziekenhuis"),
        "huisarts":   ("chk_huisarts",   "cmb_rings_huisarts",   "txt_weight_huisarts"),
    }

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.result = None

        xaml_path = os.path.join(os.path.dirname(__file__), "UI.xaml")
        root = load_xaml_window(self, xaml_path)

        # Bind alle UI elementen
        ui_elements = [
            "txt_subtitle",
            "pnl_location", "txt_location_x", "txt_location_y",
            "pnl_no_location",
            "cmb_preset",
            "chk_school", "cmb_rings_school", "txt_weight_school",
            "chk_bus", "cmb_rings_bus", "txt_weight_bus",
            "chk_trein", "cmb_rings_trein", "txt_weight_trein",
            "chk_park", "cmb_rings_park", "txt_weight_park",
            "chk_supermarkt", "cmb_rings_supermarkt", "txt_weight_supermarkt",
            "chk_ziekenhuis", "cmb_rings_ziekenhuis", "txt_weight_ziekenhuis",
            "chk_huisarts", "cmb_rings_huisarts", "txt_weight_huisarts",
            "cmb_grid_size", "cmb_resolution", "chk_smoothing",
            "pnl_progress", "txt_progress", "progress_bar",
            "txt_status",
            "btn_cancel", "btn_execute",
        ]
        bind_ui_elements(self, root, ui_elements)

        # Locatie ophalen
        self.location_rd = setup_project_location(self, doc, log)

        # Events
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.cmb_preset.SelectionChanged += self._on_preset_changed

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_preset_changed(self, sender, args):
        """Pas categorie checkboxes, bereik en gewichten aan bij preset wijziging."""
        item = self.cmb_preset.SelectedItem
        if not item or not hasattr(item, "Tag"):
            return

        preset_name = str(item.Tag)
        if preset_name == "custom":
            return

        preset = PRESETS.get(preset_name)
        if not preset:
            return

        for cat_id, (chk_name, rings_name, weight_name) in self.CAT_UI_MAP.items():
            overrides = preset.get(cat_id, {})
            chk = getattr(self, chk_name, None)
            cmb = getattr(self, rings_name, None)
            txt = getattr(self, weight_name, None)
            if chk is not None:
                chk.IsChecked = overrides.get("enabled", True)
            if cmb is not None:
                ring_profile = overrides.get("ring_profile")
                if ring_profile is not None:
                    self._select_combobox_by_tag(cmb, str(ring_profile))
            if txt is not None:
                txt.Text = str(overrides.get("max_score", 0.5))

    @staticmethod
    def _select_combobox_by_tag(combobox, tag_value):
        """Selecteer een ComboBox item op basis van Tag waarde."""
        for i in range(combobox.Items.Count):
            item = combobox.Items[i]
            if hasattr(item, "Tag") and str(item.Tag) == tag_value:
                combobox.SelectedIndex = i
                return

    def _get_categories(self):
        """Lees categorie instellingen uit de UI, inclusief ringprofiel."""
        cats = copy.deepcopy(CATEGORIES)
        for cat_id, (chk_name, rings_name, weight_name) in self.CAT_UI_MAP.items():
            chk = getattr(self, chk_name, None)
            cmb = getattr(self, rings_name, None)
            txt = getattr(self, weight_name, None)
            if cat_id in cats:
                if chk is not None:
                    cats[cat_id]["enabled"] = bool(chk.IsChecked)
                if cmb is not None:
                    ring_item = cmb.SelectedItem
                    if ring_item and hasattr(ring_item, "Tag"):
                        try:
                            profile_key = int(ring_item.Tag)
                            cats[cat_id]["ring_profile"] = profile_key
                            apply_ring_profile(cats[cat_id])
                        except (ValueError, TypeError):
                            pass
                if txt is not None:
                    try:
                        cats[cat_id]["max_score"] = float(txt.Text)
                    except (ValueError, TypeError):
                        pass
        return cats

    def _get_grid_size(self):
        """Lees grid grootte uit combobox."""
        item = self.cmb_grid_size.SelectedItem
        if item and hasattr(item, "Tag"):
            return int(item.Tag)
        return 2000

    def _get_resolution(self):
        """Lees grid resolutie uit combobox."""
        item = self.cmb_resolution.SelectedItem
        if item and hasattr(item, "Tag"):
            return int(item.Tag)
        return 50

    def _on_execute(self, sender, args):
        """Start de analyse."""
        if not self.location_rd:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        # Check of minstens 1 categorie aan staat
        categories = self._get_categories()
        enabled = [c for c in categories.values() if c.get("enabled")]
        if not enabled:
            self.txt_status.Text = "Selecteer minstens 1 categorie"
            return

        show_progress(self, "Analyse starten...")
        self.btn_execute.IsEnabled = False

        try:
            self.result = self._run_analysis(categories)
            self.DialogResult = True
            self.Close()
        except Exception as e:
            log("Analyse fout: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    def _run_analysis(self, categories):
        """Voer de volledige analyse uit."""
        rd_x, rd_y = self.location_rd
        grid_size = self._get_grid_size()
        resolution = self._get_resolution()
        use_smoothing = bool(self.chk_smoothing.IsChecked)

        # Converteer center naar WGS84
        center_lat, center_lon = rd_to_wgs84(rd_x, rd_y)
        log("Center: RD ({0:.0f}, {1:.0f}) -> WGS84 ({2:.6f}, {3:.6f})".format(
            rd_x, rd_y, center_lat, center_lon))

        # ─── Stap 1: OSM data ophalen ───
        show_progress(self, "OSM data ophalen...")
        update_ui()

        max_ring = get_max_ring(categories)
        fetch_size = grid_size + 2 * max_ring
        bbox_rd = BoundingBox.from_center(rd_x, rd_y, fetch_size, fetch_size)

        sw_lat, sw_lon = rd_to_wgs84(bbox_rd.xmin, bbox_rd.ymin)
        ne_lat, ne_lon = rd_to_wgs84(bbox_rd.xmax, bbox_rd.ymax)
        bbox_wgs84 = (sw_lat, sw_lon, ne_lat, ne_lon)

        client = OverpassClient(timeout=90, log_func=log)
        pois_per_category = {}

        enabled_cats = {k: v for k, v in categories.items() if v.get("enabled")}
        total_cats = len(enabled_cats)

        for idx, (cat_id, cat) in enumerate(enabled_cats.items()):
            show_progress(self, "Ophalen: {0} ({1}/{2})...".format(
                cat["label"], idx + 1, total_cats))
            update_ui()

            # Check cache
            cache_path = _get_cache_path(
                center_lat, center_lon, fetch_size, cat_id)
            cached = _read_cache(cache_path)

            if cached is not None and len(cached) > 0:
                pois_per_category[cat_id] = cached
                log("{0}: {1} POIs (cache)".format(cat_id, len(cached)))
            else:
                # Rate limiting: 2 seconden tussen requests (voorkomt 429)
                if idx > 0:
                    time.sleep(2.0)

                pois = client.get_pois(bbox_wgs84, cat["osm_tags"])
                pois_per_category[cat_id] = pois
                log("{0}: {1} POIs (API)".format(cat_id, len(pois)))

                # Alleen cachen als er resultaten zijn
                if pois:
                    _write_cache(cache_path, pois)

        # ─── Stap 2: Grid genereren en scores berekenen ───
        show_progress(self, "Grid scores berekenen...")
        update_ui()

        grid_points, (n_rows, n_cols) = generate_grid(
            center_lat, center_lon, grid_size, resolution
        )
        log("Grid: {0} x {1} = {2} punten".format(n_rows, n_cols, len(grid_points)))

        scores = calculate_scores(grid_points, pois_per_category, categories)

        if use_smoothing:
            scores = smooth_scores(scores, n_rows, n_cols)
            log("Smoothing toegepast")

        # ─── Stap 3: Heatmap visualiseren ───
        show_progress(self, "Heatmap renderen in Revit...")
        update_ui()

        result = create_and_apply_heatmap(
            self.doc,
            grid_points, scores,
            grid_size, n_rows, n_cols,
            origin_rd=(rd_x, rd_y),
            center_rd=(rd_x, rd_y),
            log=log
        )

        if result is None:
            raise RuntimeError("Heatmap kon niet worden aangemaakt")

        # Score op projectlocatie berekenen
        center_score = self._get_center_score(
            grid_points, scores, center_lat, center_lon)
        result["center_score"] = center_score
        result["grid_size"] = grid_size
        result["resolution"] = resolution
        result["categories_used"] = len(enabled_cats)

        # POI totalen toevoegen
        total_pois = sum(len(p) for p in pois_per_category.values())
        result["total_pois"] = total_pois

        return result

    def _get_center_score(self, grid_points, scores, center_lat, center_lon):
        """Vind de score van het punt dichtst bij het centrum."""
        from gis2bim.analysis.grid import haversine

        min_dist = None
        center_score = 0.0

        for i, pt in enumerate(grid_points):
            dist = haversine(pt["lat"], pt["lon"], center_lat, center_lon)
            if min_dist is None or dist < min_dist:
                min_dist = dist
                center_score = scores[i]

        return center_score


# ─── Samenvatting ────────────────────────────────────────────────

def _format_summary(result):
    """Formatteer de analyse resultaten voor weergave."""
    lines = [
        "Gebiedsanalyse voltooid",
        "",
        "Grid: {0} x {0} m, resolutie {1} m".format(
            result.get("grid_size", "?"),
            result.get("resolution", "?")),
        "Categorieen: {0}".format(result.get("categories_used", "?")),
        "POIs gevonden: {0}".format(result.get("total_pois", "?")),
        "Gridpunten: {0}".format(result.get("point_count", "?")),
        "",
        "Score op projectlocatie: {0:.2f}".format(
            result.get("center_score", 0)),
        "Gemiddelde score: {0:.2f}".format(result.get("avg_score", 0)),
        "Maximum score: {0:.2f}".format(result.get("max_score", 0)),
        "Minimum score: {0:.2f}".format(result.get("min_score", 0)),
    ]
    return "\n".join(lines)


def _write_exchange_json(result, rd_x, rd_y):
    """Schrijf resultaten naar uitwisselings-JSON voor andere tools."""
    try:
        exchange_dir = os.path.join(os.environ.get("TEMP", ""), "3bm_exchange")
        if not os.path.exists(exchange_dir):
            os.makedirs(exchange_dir)

        data = {
            "tool": "GebiedsAnalyse",
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "location_rd": {"x": rd_x, "y": rd_y},
            "grid_size": result.get("grid_size"),
            "resolution": result.get("resolution"),
            "score_center": result.get("center_score", 0),
            "score_avg": result.get("avg_score", 0),
            "score_max": result.get("max_score", 0),
            "score_min": result.get("min_score", 0),
            "total_pois": result.get("total_pois", 0),
            "categories_used": result.get("categories_used", 0),
        }

        path = os.path.join(exchange_dir, "gebiedsanalyse_result.json")
        with open(path, "w") as f:
            json.dump(data, f, indent=2)

        log("Exchange JSON geschreven: {0}".format(path))
    except Exception as e:
        log("Fout bij exchange JSON: {0}".format(e))


# ─── Main ────────────────────────────────────────────────────────

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

    window = GebiedsAnalyseWindow(doc)
    dialog_result = window.ShowDialog()

    if dialog_result and window.result:
        result = window.result
        summary = _format_summary(result)

        # Exchange JSON schrijven
        if window.location_rd:
            _write_exchange_json(result, window.location_rd[0], window.location_rd[1])

        forms.alert(
            summary,
            title="GIS2BIM - Gebiedsanalyse",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
