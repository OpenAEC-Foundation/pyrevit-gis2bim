# -*- coding: utf-8 -*-
"""
Kaart Tijdreis - GIS2BIM
=========================

Download historische kaarten van verschillende jaren via de
ArcGIS Historische Tijdreis service en plaats deze op een
Revit sheet in een 4x2 grid. Automatische chronologische
sortering en duplicaat-detectie.
"""

__title__ = "Kaart\nTijdreis"
__author__ = "OpenAEC Foundation"
__doc__ = "Download historische kaarten van verschillende jaren en plaats op een sheet"

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
    calculate_grid_layout,
    calculate_grid_position,
    place_image_on_sheet,
    place_label_on_sheet,
    find_a3_titleblock,
    find_any_titleblock,
    populate_sheets_dropdown,
)

log = get_logger("KaartTijdreis")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.wmts_tiles import ArcGISTileClient, TIJDREIS_YEARS
    from gis2bim.coordinates import create_bbox_rd
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# Beschikbare kaartlagen
KAART_LAYERS = []
for _y in TIJDREIS_YEARS if GIS2BIM_LOADED else []:
    _label = _y.replace("_", "-")
    KAART_LAYERS.append({"key": _y, "name": _label})
KAART_LAYERS.append({"key": "geen", "name": "- Geen -"})

# Default selecties per slot (8 stuks, breed gespreid)
DEFAULT_SLOTS = [
    "1850",
    "1900",
    "1938",
    "1958",
    "1975",
    "1990",
    "2005",
    "2019",
]

# Grid layout constanten
GRID_COLS = 4
GRID_ROWS = 2
GAP_H_MM = 12.0
GAP_V_MM = 15.0
LABEL_OFFSET_MM = 3.0
TOP_MARGIN_MM = 25.0
MM_TO_FEET = 1.0 / 304.8


class KaartTijdreisWindow(Window):
    """WPF Window voor Kaart Tijdreis."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.slot_combos = []

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle',
            'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location',
            'cmb_bbox_size',
            'cmb_slot_1', 'cmb_slot_2', 'cmb_slot_3', 'cmb_slot_4',
            'cmb_slot_5', 'cmb_slot_6', 'cmb_slot_7', 'cmb_slot_8',
            'txt_area_width', 'txt_area_height',
            'rdo_new_sheet', 'rdo_existing_sheet',
            'pnl_existing_sheet', 'cmb_existing_sheet',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status',
            'btn_cancel', 'btn_execute',
        ])

        self.slot_combos = [
            self.cmb_slot_1, self.cmb_slot_2, self.cmb_slot_3, self.cmb_slot_4,
            self.cmb_slot_5, self.cmb_slot_6, self.cmb_slot_7, self.cmb_slot_8,
        ]

        self.location_rd = setup_project_location(self, doc, log)
        self._populate_slot_combos()
        populate_sheets_dropdown(self.cmb_existing_sheet, doc,
                                 default_sheet_number="090", log=log)
        self._bind_events()

    def _populate_slot_combos(self):
        """Vul de 8 slot ComboBoxes met beschikbare kaartjaren."""
        for slot_idx, combo in enumerate(self.slot_combos):
            default_key = DEFAULT_SLOTS[slot_idx] if slot_idx < len(DEFAULT_SLOTS) else "geen"

            for layer in KAART_LAYERS:
                item = ComboBoxItem()
                item.Content = layer["name"]
                item.Tag = layer["key"]
                combo.Items.Add(item)

                if layer["key"] == default_key:
                    combo.SelectedItem = item

            if combo.SelectedItem is None and combo.Items.Count > 0:
                combo.SelectedIndex = 0

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
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

    @staticmethod
    def _sort_key(year_key):
        """Sorteersleutel voor chronologische volgorde.

        Handelt ranges af zoals '1823_1829' -> sorteert op eerste jaar.
        """
        try:
            return int(year_key.split("_")[0])
        except (ValueError, IndexError):
            return 9999

    def _on_execute(self, sender, args):
        """Voer de kaart tijdreis uit."""
        log("Execute gestart")

        if self.location_rd is None:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        try:
            # Haal geselecteerde jaren op
            selected_slots = self._get_selected_slots()
            active_years = [s for s in selected_slots if s is not None]

            if not active_years:
                self.txt_status.Text = "Selecteer minimaal 1 kaartjaar"
                return

            # Verwijder dubbele selecties
            seen = set()
            unique_years = []
            for y in active_years:
                if y not in seen:
                    seen.add(y)
                    unique_years.append(y)

            # Sorteer chronologisch
            unique_years.sort(key=self._sort_key)
            log("Chronologisch gesorteerd: {0}".format(unique_years))

            # Herindex als (slot_idx, year_key) voor grid plaatsing
            active_slots = list(enumerate(unique_years))

            # Toon progress
            show_progress(self, "Voorbereiden...")
            self.btn_execute.IsEnabled = False

            # Haal parameters op
            bbox_size = self._get_bbox_size()
            center_x, center_y = self.location_rd
            bbox = create_bbox_rd(center_x, center_y, bbox_size)
            log("Bbox: {0}, grootte: {1}m".format(bbox, bbox_size))

            # ArcGIS Tile Client
            client = ArcGISTileClient()
            errors = []
            downloaded_images = []

            # Duplicaat-detectie: check sample tiles
            show_progress(self, "Controleren op duplicaat-kaarten...")
            update_ui()
            hash_map = {}  # hash -> eerste year_key
            duplicates = []
            for _, year_key in active_slots:
                h = client.get_sample_hash(year_key, bbox)
                if h and h in hash_map:
                    duplicates.append((year_key, hash_map[h]))
                    log("Duplicaat: {0} = {1}".format(year_key, hash_map[h]))
                elif h:
                    hash_map[h] = year_key

            if duplicates:
                dup_msgs = []
                for dup_year, orig_year in duplicates:
                    dup_msgs.append("{0} = {1}".format(
                        dup_year.replace("_", "-"), orig_year.replace("_", "-")))
                msg = (
                    "Sommige jaren tonen dezelfde kaart op deze locatie:\n\n"
                    + "\n".join(dup_msgs)
                    + "\n\nToch doorgaan?"
                )
                if not forms.alert(msg, title="Duplicaten gedetecteerd",
                                   yes=True, no=True):
                    hide_progress(self)
                    self.btn_execute.IsEnabled = True
                    return

            # Download alle kaarten
            for slot_idx, year_key in active_slots:
                show_progress(self, "Downloaden: {0}...".format(year_key))

                try:
                    image_path = client.download_image(year_key, bbox)
                    downloaded_images.append((slot_idx, year_key, image_path))
                    log("Download OK: {0} -> {1}".format(year_key, image_path))
                except Exception as e:
                    errors.append("{0}: download fout - {1}".format(year_key, str(e)))
                    log("Download fout {0}: {1}".format(year_key, e))
                    log(traceback.format_exc())

            if not downloaded_images:
                self.txt_status.Text = "Geen kaarten gedownload. Controleer de verbinding."
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            # Start Revit transactie
            show_progress(self, "Sheet en kaarten plaatsen...")

            t = DB.Transaction(self.doc, "GIS2BIM - Kaart Tijdreis")
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

                # Bereken grid layout vanuit centrum van de sheet
                layout = self._build_centered_layout(sheet)

                # Kaarten plaatsen in grid
                placed_count = 0
                for slot_idx, year_key, image_path in downloaded_images:
                    label = year_key.replace("_", "-")
                    show_progress(self, "Plaatsen: {0}...".format(label))

                    try:
                        center_x, center_y = calculate_grid_position(slot_idx, layout)
                        place_image_on_sheet(
                            self.doc, sheet, image_path,
                            center_x, center_y, layout["img_size"], log
                        )
                        place_label_on_sheet(
                            self.doc, sheet, label,
                            center_x, center_y, layout["img_size"],
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

    def _get_selected_slots(self):
        """Haal geselecteerde jaar keys op per slot. None = 'geen'."""
        result = []
        for combo in self.slot_combos:
            item = combo.SelectedItem
            if item and hasattr(item, 'Tag'):
                key = item.Tag
                if key == "geen":
                    result.append(None)
                else:
                    result.append(key)
            else:
                result.append(None)
        return result

    def _get_bbox_size(self):
        """Haal geselecteerde bbox grootte op."""
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 500

    def _get_area_size(self):
        """Haal beschikbare ruimte op uit de invulvelden (mm).

        Returns:
            Tuple (breedte_mm, hoogte_mm)
        """
        try:
            w = float(self.txt_area_width.Text.strip())
        except (ValueError, AttributeError):
            w = 380.0
        try:
            h = float(self.txt_area_height.Text.strip())
        except (ValueError, AttributeError):
            h = 210.0
        return (w, h)

    def _build_centered_layout(self, sheet):
        """Bereken grid layout gecentreerd op de sheet.

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

        # Maximale image grootte (vierkant) die past in het grid
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
            sheet.Name = "Kaart tijdreis"

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
            number = "TT-{0:02d}".format(i)
            if number not in existing_numbers:
                return number

        return "TT-99"

    def _show_result(self, placed_count, errors):
        """Toon resultaat dialoog."""
        msg_lines = [
            "Kaart tijdreis gereed!",
            "",
            "Kaarten geplaatst: {0}".format(placed_count),
        ]

        if errors:
            msg_lines.append("")
            msg_lines.append("Waarschuwingen:")
            for err in errors:
                msg_lines.append("  - {0}".format(err))

        forms.alert("\n".join(msg_lines), title="GIS2BIM - Kaart Tijdreis")


def main():
    log("=== GIS2BIM Kaart Tijdreis Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = KaartTijdreisWindow(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM Kaart Tijdreis Tool beeindigd ===")


if __name__ == "__main__":
    main()
