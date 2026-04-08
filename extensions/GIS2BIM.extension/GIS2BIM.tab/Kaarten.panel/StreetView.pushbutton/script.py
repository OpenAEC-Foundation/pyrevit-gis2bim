# -*- coding: utf-8 -*-
"""
Street View 360° - GIS2BIM
============================

Download 8 Google Street View images (elke 45°) en plaats
deze op een Revit sheet in een 4x2 grid.
"""

__title__ = "Street View\n360\u00B0"
__author__ = "OpenAEC Foundation"
__doc__ = "Download Google Street View 360\u00B0 panorama en plaats op een sheet"

# CLR references voor WPF
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
clr.AddReference('System.Xml')

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
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui
from gis2bim.revit.sheets import (
    get_sheet_bounds,
    calculate_grid_position,
    place_image_on_sheet,
    place_label_on_sheet,
    find_a3_titleblock,
    find_any_titleblock,
    populate_sheets_dropdown,
)

log = get_logger("StreetView360")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.streetview import StreetViewClient
    from gis2bim.coordinates import rd_to_wgs84
    from gis2bim.config import get_api_key, set_api_key
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# 8 kompasrichtingen, elke 45 graden
HEADINGS = [
    {"heading": 0,   "label": "N (0\u00B0)",     "short": "N"},
    {"heading": 45,  "label": "NO (45\u00B0)",    "short": "NO"},
    {"heading": 90,  "label": "O (90\u00B0)",     "short": "O"},
    {"heading": 135, "label": "ZO (135\u00B0)",   "short": "ZO"},
    {"heading": 180, "label": "Z (180\u00B0)",    "short": "Z"},
    {"heading": 225, "label": "ZW (225\u00B0)",   "short": "ZW"},
    {"heading": 270, "label": "W (270\u00B0)",    "short": "W"},
    {"heading": 315, "label": "NW (315\u00B0)",   "short": "NW"},
]

# Grid layout constanten (4 kolommen x 2 rijen)
GRID_COLS = 4
GRID_ROWS = 2
GAP_H_MM = 12.0
GAP_V_MM = 15.0
LABEL_OFFSET_MM = 3.0
TOP_MARGIN_MM = 25.0
MM_TO_FEET = 1.0 / 304.8


class StreetView360Window(Window):
    """WPF Window voor Street View 360°."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle',
            'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location',
            'txt_api_key', 'btn_save_key',
            'cmb_fov', 'cmb_pitch', 'cmb_image_size',
            'txt_area_width', 'txt_area_height',
            'rdo_new_sheet', 'rdo_existing_sheet',
            'pnl_existing_sheet', 'cmb_existing_sheet',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status',
            'btn_cancel', 'btn_execute',
        ])

        # Locatie ophalen
        self.location_rd = setup_project_location(self, doc, log)

        # API key laden
        saved_key = get_api_key()
        if saved_key:
            self.txt_api_key.Text = saved_key
            log("API key geladen uit config")

        # Sheets dropdown vullen
        populate_sheets_dropdown(self.cmb_existing_sheet, doc,
                                 default_sheet_number="092", log=log)

        # Events binden
        self._bind_events()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.btn_save_key.Click += self._on_save_key
        self.rdo_new_sheet.Checked += self._on_sheet_option_changed
        self.rdo_existing_sheet.Checked += self._on_sheet_option_changed

    def _on_sheet_option_changed(self, sender, args):
        """Toggle zichtbaarheid van bestaande sheet selectie."""
        if self.rdo_existing_sheet.IsChecked:
            self.pnl_existing_sheet.Visibility = Visibility.Visible
        else:
            self.pnl_existing_sheet.Visibility = Visibility.Collapsed

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_save_key(self, sender, args):
        """Sla API key op."""
        key = self.txt_api_key.Text.strip()
        if key:
            set_api_key(key)
            self.txt_status.Text = "API key opgeslagen"
            log("API key opgeslagen")
        else:
            self.txt_status.Text = "Voer een API key in"

    def _on_execute(self, sender, args):
        """Voer de Street View 360° download en plaatsing uit."""
        log("Execute gestart")

        # Validaties
        if self.location_rd is None:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        api_key = self.txt_api_key.Text.strip()
        if not api_key:
            self.txt_status.Text = "Voer een Google API key in"
            return

        try:
            # Toon progress
            show_progress(self, "Voorbereiden...")
            self.btn_execute.IsEnabled = False

            # RD naar WGS84
            center_x, center_y = self.location_rd
            lat, lon = rd_to_wgs84(center_x, center_y)
            log("Locatie: RD ({0}, {1}) -> WGS84 ({2:.6f}, {3:.6f})".format(
                center_x, center_y, lat, lon))

            # Camera parameters ophalen
            fov = self._get_combo_value(self.cmb_fov, 75)
            pitch = self._get_combo_value(self.cmb_pitch, 0)
            img_pixels = self._get_combo_value(self.cmb_image_size, 1024)

            log("Camera: FOV={0}, pitch={1}, pixels={2}".format(fov, pitch, img_pixels))

            # API key opslaan als dat nog niet was gedaan
            set_api_key(api_key)

            # Street View Client
            client = StreetViewClient(api_key)
            errors = []
            downloaded_images = []

            # Download 8 images (elke 45°)
            for slot_idx, heading_info in enumerate(HEADINGS):
                heading = heading_info["heading"]
                label = heading_info["label"]
                show_progress(self, "Downloaden: {0}...".format(label))

                try:
                    image_path = client.download_image(
                        lat, lon, heading,
                        fov=fov, pitch=pitch,
                        width=img_pixels, height=img_pixels
                    )
                    downloaded_images.append((slot_idx, heading_info, image_path))
                    log("Download OK: {0} -> {1}".format(label, image_path))
                except Exception as e:
                    errors.append("{0}: download fout - {1}".format(label, str(e)))
                    log("Download fout {0}: {1}".format(label, e))
                    log(traceback.format_exc())

            if not downloaded_images:
                self.txt_status.Text = "Geen images gedownload. Controleer API key en verbinding."
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            # Start Revit transactie
            show_progress(self, "Sheet en images plaatsen...")

            t = DB.Transaction(self.doc, "GIS2BIM - Street View 360")
            t.Start()

            try:
                # Sheet ophalen of aanmaken
                sheet = self._get_or_create_sheet()
                if sheet is None:
                    t.RollBack()
                    self.txt_status.Text = "Kon sheet niet aanmaken/vinden"
                    hide_progress(self)
                    self.btn_execute.IsEnabled = True
                    return

                log("Sheet: {0} - {1}".format(sheet.SheetNumber, sheet.Name))

                # Bereken grid layout (4x2)
                layout = self._build_centered_layout(sheet)

                # Images plaatsen in grid
                placed_count = 0
                for slot_idx, heading_info, image_path in downloaded_images:
                    label = heading_info["label"]
                    show_progress(self, "Plaatsen: {0}...".format(label))

                    try:
                        cx, cy = calculate_grid_position(slot_idx, layout)
                        place_image_on_sheet(
                            self.doc, sheet, image_path,
                            cx, cy, layout["img_size"], log
                        )
                        place_label_on_sheet(
                            self.doc, sheet, label,
                            cx, cy, layout["img_size"],
                            label_offset_mm=LABEL_OFFSET_MM, log=log
                        )
                        placed_count += 1
                        log("Geplaatst: {0} op positie {1}".format(label, slot_idx + 1))
                    except Exception as e:
                        errors.append("{0}: plaatsingsfout - {1}".format(label, str(e)))
                        log("Plaatsingsfout {0}: {1}".format(label, e))
                        log(traceback.format_exc())

                    # Opruimen temp bestand
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

            # Klaar
            hide_progress(self)
            self.DialogResult = True
            self.Close()

            # Toon resultaat
            self._show_result(placed_count, errors)

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            hide_progress(self)
            self.btn_execute.IsEnabled = True

    def _get_combo_value(self, combo, default):
        """Haal numerieke waarde op uit ComboBox Tag."""
        item = combo.SelectedItem
        if item and hasattr(item, 'Tag'):
            try:
                return int(item.Tag)
            except (ValueError, TypeError):
                pass
        return default

    def _get_area_size(self):
        """Haal beschikbare ruimte op uit invulvelden (mm).

        Returns:
            Tuple (breedte_mm, hoogte_mm)
        """
        try:
            w = float(self.txt_area_width.Text.strip())
        except (ValueError, AttributeError):
            w = 400.0
        try:
            h = float(self.txt_area_height.Text.strip())
        except (ValueError, AttributeError):
            h = 210.0
        return (w, h)

    def _build_centered_layout(self, sheet):
        """Bereken 4x2 grid layout gecentreerd op de sheet.

        Returns:
            Layout dict compatible met calculate_grid_position()
        """
        # Haal sheet bounds op via titelblok
        sheet_bounds = get_sheet_bounds(self.doc, sheet, log)
        sheet_cx = (sheet_bounds[0] + sheet_bounds[2]) / 2.0
        sheet_top = sheet_bounds[3]  # y_max = bovenkant sheet
        top_margin_ft = TOP_MARGIN_MM * MM_TO_FEET

        # Gebruiker-opgegeven beschikbare ruimte
        area_w_mm, area_h_mm = self._get_area_size()
        area_w_ft = area_w_mm * MM_TO_FEET
        area_h_ft = area_h_mm * MM_TO_FEET

        log("Beschikbare ruimte: {0:.0f} x {1:.0f} mm, top marge: {2:.0f} mm".format(
            area_w_mm, area_h_mm, TOP_MARGIN_MM))
        log("Sheet bovenkant: {0:.4f} ft, centrum X: {1:.4f} ft".format(
            sheet_top, sheet_cx))

        gap_h_ft = GAP_H_MM * MM_TO_FEET
        gap_v_ft = GAP_V_MM * MM_TO_FEET

        # Maximale image grootte (vierkant) die past in het 4x2 grid
        max_img_w = (area_w_ft - (GRID_COLS - 1) * gap_h_ft) / GRID_COLS
        max_img_h = (area_h_ft - (GRID_ROWS - 1) * gap_v_ft) / GRID_ROWS
        img_size = min(max_img_w, max_img_h)

        # Content grootte
        content_w = GRID_COLS * img_size + (GRID_COLS - 1) * gap_h_ft
        content_h = GRID_ROWS * img_size + (GRID_ROWS - 1) * gap_v_ft

        # Horizontaal: gecentreerd op sheet centrum
        x_start = sheet_cx - content_w / 2.0

        # Verticaal: beschikbare ruimte vanaf bovenkant sheet (met marge),
        # content gecentreerd binnen die ruimte
        area_top = sheet_top - top_margin_ft
        area_center_y = area_top - area_h_ft / 2.0
        y_start = area_center_y + content_h / 2.0

        img_mm = img_size / MM_TO_FEET
        log("Grid layout: image={0:.0f}mm, content={1:.0f}x{2:.0f}mm".format(
            img_mm, content_w / MM_TO_FEET, content_h / MM_TO_FEET))

        return {
            "x_start": x_start,
            "y_start": y_start,
            "img_size": img_size,
            "cols": GRID_COLS,
            "rows": GRID_ROWS,
            "gap_h": gap_h_ft,
            "gap_v": gap_v_ft,
        }

    def _get_or_create_sheet(self):
        """Maak nieuwe sheet of haal bestaande op."""
        if self.rdo_existing_sheet.IsChecked:
            item = self.cmb_existing_sheet.SelectedItem
            if item and hasattr(item, 'Tag'):
                sheet_id = item.Tag
                return self.doc.GetElement(sheet_id)
            return None
        else:
            return self._create_new_sheet()

    def _create_new_sheet(self):
        """Maak een nieuwe A3 sheet aan."""
        try:
            titleblock_id = find_a3_titleblock(self.doc, log)
            if titleblock_id is None:
                titleblock_id = find_any_titleblock(self.doc, log)
            if titleblock_id is None:
                log("Geen titelblok gevonden")
                return None

            sheet = DB.ViewSheet.Create(self.doc, titleblock_id)
            sheet.Name = "Street View 360"

            sheet_number = self._generate_sheet_number()
            try:
                sheet.SheetNumber = sheet_number
            except Exception:
                pass

            log("Sheet aangemaakt: {0} - {1}".format(sheet.SheetNumber, sheet.Name))
            return sheet

        except Exception as e:
            log("Fout bij aanmaken sheet: {0}".format(e))
            log(traceback.format_exc())
            return None

    def _generate_sheet_number(self):
        """Genereer een uniek sheet nummer."""
        collector = DB.FilteredElementCollector(self.doc)
        sheets = collector.OfClass(DB.ViewSheet).ToElements()
        existing_numbers = set()
        for s in sheets:
            existing_numbers.add(s.SheetNumber)

        for i in range(1, 100):
            number = "SV-{0:02d}".format(i)
            if number not in existing_numbers:
                return number

        return "SV-99"

    def _show_result(self, placed_count, errors):
        """Toon resultaat dialoog."""
        msg_lines = [
            "Street View 360\u00B0 gereed!",
            "",
            "Foto's geplaatst: {0} van {1}".format(placed_count, len(HEADINGS)),
        ]

        if errors:
            msg_lines.append("")
            msg_lines.append("Waarschuwingen:")
            for err in errors:
                msg_lines.append("  - {0}".format(err))

        forms.alert("\n".join(msg_lines), title="GIS2BIM - Street View 360\u00B0")


def main():
    log("=== GIS2BIM Street View 360 Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = StreetView360Window(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM Street View 360 Tool beeindigd ===")


if __name__ == "__main__":
    main()
