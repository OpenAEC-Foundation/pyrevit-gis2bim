# -*- coding: utf-8 -*-
"""
Locatie Button - GIS2BIM (WPF versie)
=====================================

Stel projectlocatie in via PDOK Locatieserver.
"""

__title__ = "Locatie\nInstellen"
__author__ = "3BM Bouwkunde"
__doc__ = "Stel projectlocatie in via PDOK (adres naar RD coördinaten)"

# CLR references voor WPF
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
clr.AddReference('System.Xml')

from System.Windows import Window, Visibility

# pyRevit
from pyrevit import revit, DB, script, forms

# Standaard library
import sys
import os
import traceback
import re

# Voeg lib folder toe aan path
extension_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(extension_path, "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Gedeelde modules
from gis2bim.ui.logging_helper import create_tool_logger
from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements

log, LOG_FILE = create_tool_logger("Locatie", __file__)

# GIS2BIM modules
try:
    from gis2bim.api.pdok import PDOKLocatie
    from gis2bim.revit.location import (
        set_project_location_rd,
        set_site_location_wgs84,
        get_project_location_rd,
        get_site_location,
        set_project_info_from_location
    )
    from gis2bim.coordinates import wgs84_to_rd
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    GIS2BIM_LOADED = False
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))


class LocatieWindow(Window):
    """WPF Window voor Locatie instellen"""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.result_data = None

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_current', 'txt_current_x', 'txt_current_y',
            'rb_revit', 'rb_address', 'rb_postcode', 'txt_revit_location',
            'pnl_address', 'txt_city', 'txt_street', 'txt_number_addr',
            'pnl_postcode', 'txt_postcode', 'txt_number_post',
            'pnl_result', 'txt_result_address', 'txt_result_x', 'txt_result_y',
            'txt_result_prov', 'txt_result_wind', 'txt_result_wgs',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self._setup_current_location()
        self._setup_revit_site()
        self._bind_events()

    def _setup_current_location(self):
        try:
            current = get_project_location_rd(self.doc)
            if current and (current["rd_x"] != 0 or current["rd_y"] != 0):
                self.pnl_current.Visibility = Visibility.Visible
                self.txt_current_x.Text = "{0:,.0f} m".format(current["rd_x"])
                self.txt_current_y.Text = "{0:,.0f} m".format(current["rd_y"])
        except Exception as e:
            log("Error getting current location: {0}".format(e))

    def _setup_revit_site(self):
        try:
            site = get_site_location(self.doc)
            has_site = site and site.get("place_name") and site["place_name"] != "Default"

            if has_site:
                self.rb_revit.IsEnabled = True
                self.txt_revit_location.Text = "({0})".format(site["place_name"])
                self._site_data = site
            else:
                self.rb_revit.IsEnabled = False
                self.txt_revit_location.Text = "(niet ingesteld)"
                self._site_data = None
        except Exception as e:
            log("Error getting site location: {0}".format(e))
            self.rb_revit.IsEnabled = False
            self._site_data = None

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute

        self.rb_revit.Checked += self._on_method_changed
        self.rb_address.Checked += self._on_method_changed
        self.rb_postcode.Checked += self._on_method_changed

        self.txt_city.LostFocus += self._on_input_changed
        self.txt_street.LostFocus += self._on_input_changed
        self.txt_number_addr.LostFocus += self._on_input_changed
        self.txt_postcode.LostFocus += self._on_input_changed
        self.txt_number_post.LostFocus += self._on_input_changed

    def _on_method_changed(self, sender, args):
        if self.rb_address.IsChecked:
            self.pnl_address.Visibility = Visibility.Visible
            self.pnl_postcode.Visibility = Visibility.Collapsed
        elif self.rb_postcode.IsChecked:
            self.pnl_address.Visibility = Visibility.Collapsed
            self.pnl_postcode.Visibility = Visibility.Visible
        else:
            self.pnl_address.Visibility = Visibility.Collapsed
            self.pnl_postcode.Visibility = Visibility.Collapsed
            if self._site_data:
                self._show_revit_site_result()

        if not self.rb_revit.IsChecked:
            self.pnl_result.Visibility = Visibility.Collapsed
            self.result_data = None

    def _on_input_changed(self, sender, args):
        self._search_location()

    def _search_location(self):
        self.txt_status.Text = ""
        self.pnl_result.Visibility = Visibility.Collapsed
        self.result_data = None

        try:
            loc = PDOKLocatie()
            result = None

            if self.rb_address.IsChecked:
                city = self.txt_city.Text.strip()
                street = self.txt_street.Text.strip()
                number = self.txt_number_addr.Text.strip()

                if not city or not street or not number:
                    return

                log("Zoeken: {0} {1} {2}".format(city, street, number))
                result = loc.search_address(city, street, number)

            elif self.rb_postcode.IsChecked:
                postcode = self.txt_postcode.Text.strip().replace(" ", "")
                number = self.txt_number_post.Text.strip()

                if not postcode or not number:
                    return

                log("Zoeken: {0} {1}".format(postcode, number))
                result = loc.search_postcode(postcode, number)

            if result:
                log("Gevonden: {0} ({1}, {2})".format(result.gemeente, result.rd_x, result.rd_y))
                self._show_pdok_result(result)
            else:
                self.txt_status.Text = "Geen locatie gevonden"

        except Exception as e:
            log("Search error: {0}".format(e))
            self.txt_status.Text = "Zoekfout: {0}".format(str(e))

    def _show_pdok_result(self, result):
        self.result_data = result
        self.pnl_result.Visibility = Visibility.Visible

        self.txt_result_address.Text = "{0} {1}".format(result.postcode, result.gemeente)
        self.txt_result_x.Text = "{0:,.2f} m".format(result.rd_x)
        self.txt_result_y.Text = "{0:,.2f} m".format(result.rd_y)
        self.txt_result_prov.Text = result.provincie
        self.txt_result_wind.Text = str(result.windgebied)
        self.txt_result_wgs.Text = "{0:.6f}, {1:.6f}".format(result.lat, result.lon)

    def _show_revit_site_result(self):
        if not self._site_data:
            return

        rd_x, rd_y = wgs84_to_rd(self._site_data["latitude"], self._site_data["longitude"])

        self.pnl_result.Visibility = Visibility.Visible
        self.txt_result_address.Text = self._site_data["place_name"]
        self.txt_result_x.Text = "{0:,.2f} m".format(rd_x)
        self.txt_result_y.Text = "{0:,.2f} m".format(rd_y)
        self.txt_result_prov.Text = "-"
        self.txt_result_wind.Text = "-"
        self.txt_result_wgs.Text = "{0:.6f}, {1:.6f}".format(
            self._site_data["latitude"], self._site_data["longitude"]
        )

        self.result_data = {
            "method": "revit",
            "rd_x": rd_x,
            "rd_y": rd_y,
            "lat": self._site_data["latitude"],
            "lon": self._site_data["longitude"],
            "place_name": self._site_data["place_name"]
        }

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_execute(self, sender, args):
        log("Execute geklikt")
        try:
            if self.rb_revit.IsChecked:
                if not self._site_data:
                    self.txt_status.Text = "Geen Revit Site Location beschikbaar"
                    return
                self._execute_revit_site()
            else:
                if not self.result_data:
                    self._search_location()
                    if not self.result_data:
                        self.txt_status.Text = "Zoek eerst een geldige locatie"
                        return
                self._execute_pdok_result()
        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))

    def _execute_revit_site(self):
        log("Execute Revit Site...")
        place_name = self._site_data.get("place_name", "")
        log("Place name: {0}".format(place_name))

        rd_x, rd_y = wgs84_to_rd(self._site_data["latitude"], self._site_data["longitude"])

        result = set_project_location_rd(self.doc, rd_x=rd_x, rd_y=rd_y)
        log("Survey Point result: {0}".format(result))

        if result["success"]:
            postcode = ""
            housenumber = ""
            street = ""

            postcode_match = re.search(r'(\d{4})\s?([A-Za-z]{2})', place_name)
            if postcode_match:
                postcode = postcode_match.group(1) + postcode_match.group(2)
                log("Postcode gevonden: {0}".format(postcode))

            number_match = re.search(r'\s(\d+[A-Za-z]*)\s*[,\s]+\d{4}', place_name)
            if number_match:
                housenumber = number_match.group(1)
                log("Huisnummer gevonden: {0}".format(housenumber))

            if housenumber:
                street_match = re.match(r'^([^,]+?)\s+' + re.escape(housenumber), place_name)
                if street_match:
                    street = street_match.group(1).strip()
                    log("Straat gevonden: {0}".format(street))

            location_data = None
            if postcode and housenumber:
                log("PDOK lookup via postcode: {0} {1}".format(postcode, housenumber))
                try:
                    loc = PDOKLocatie()
                    pdok_result = loc.search_postcode(postcode, housenumber)
                    if pdok_result:
                        log("PDOK gevonden: {0}".format(pdok_result.gemeente))
                        location_data = pdok_result
                    else:
                        log("PDOK: geen resultaat voor postcode")
                except Exception as e:
                    log("PDOK postcode error: {0}".format(e))

            if not location_data:
                log("Fallback: reverse geocoding via RD coordinaten...")
                try:
                    loc = PDOKLocatie()
                    pdok_result = loc.search_rd_coordinates(rd_x, rd_y)
                    if pdok_result:
                        log("PDOK reverse gevonden: {0}".format(pdok_result.gemeente))
                        location_data = pdok_result
                except Exception as e:
                    log("PDOK reverse error: {0}".format(e))

            if not location_data:
                log("Gebruik minimale data uit place_name")
                location_data = {
                    "rd_x": rd_x,
                    "rd_y": rd_y,
                    "gemeente": place_name.split(",")[-2].strip() if "," in place_name else place_name,
                    "postcode": postcode,
                    "provincie": "",
                    "windgebied": "",
                    "kadaster_gemeente": "",
                    "kadaster_sectie": "",
                    "kadaster_perceel": "",
                }

            log("Setting Project Info parameters (Revit Site)...")
            log("Street: {0}, Housenumber: {1}".format(street, housenumber))
            try:
                param_result = set_project_info_from_location(
                    self.doc,
                    location_data,
                    street=street,
                    housenumber=housenumber
                )
                log("Parameter result: {0}".format(param_result))
            except Exception as e:
                log("Parameter error: {0}\n{1}".format(e, traceback.format_exc()))
                param_result = {"filled": {}, "not_found": [], "errors": [str(e)], "created": []}

            msg_lines = [
                "Survey Point ingesteld!",
                "",
                "RD X: {0:,.2f} m".format(rd_x),
                "RD Y: {0:,.2f} m".format(rd_y),
                ""
            ]

            filled = param_result.get("filled", {})
            created = param_result.get("created", [])

            if created:
                msg_lines.append("Nieuw aangemaakte parameters: {0}".format(len(created)))
            if filled:
                msg_lines.append("Ingevulde parameters: {0}".format(len(filled)))

            self.DialogResult = True
            self.Close()
            forms.alert("\n".join(msg_lines), title="GIS2BIM - Succes")
        else:
            self.txt_status.Text = "Fout: {0}".format(result["message"])

    def _execute_pdok_result(self):
        result = self.result_data
        log("Execute PDOK result: {0}".format(result.gemeente if hasattr(result, 'gemeente') else result))

        if self.rb_address.IsChecked:
            street = self.txt_street.Text.strip()
            housenumber = self.txt_number_addr.Text.strip()
        elif self.rb_postcode.IsChecked:
            street = ""
            housenumber = self.txt_number_post.Text.strip()
        else:
            street = ""
            housenumber = ""

        log("Straat: {0}, Huisnummer: {1}".format(street, housenumber))

        log("Setting Survey Point...")
        set_result = set_project_location_rd(
            self.doc,
            rd_x=result.rd_x,
            rd_y=result.rd_y,
            elevation=0.0,
            angle=0.0
        )
        log("Survey Point result: {0}".format(set_result))

        if set_result["success"]:
            log("Setting Site Location...")
            set_site_location_wgs84(
                self.doc,
                latitude=result.lat,
                longitude=result.lon,
                place_name="{0}, {1}".format(result.gemeente, result.provincie)
            )

            log("Setting Project Info parameters...")
            try:
                param_result = set_project_info_from_location(
                    self.doc,
                    result,
                    street=street,
                    housenumber=housenumber
                )
                log("Parameter result: {0}".format(param_result))
            except Exception as e:
                log("Parameter error: {0}\n{1}".format(e, traceback.format_exc()))
                param_result = {"filled": {}, "not_found": [], "errors": [str(e)], "created": []}

            msg_lines = [
                "Locatie ingesteld!",
                "",
                "Survey Point:",
                "  RD X: {0:,.2f} m".format(result.rd_x),
                "  RD Y: {0:,.2f} m".format(result.rd_y),
                "  Windgebied: {0}".format(result.windgebied),
                ""
            ]

            created = param_result.get("created", [])
            filled = param_result.get("filled", {})
            not_found = param_result.get("not_found", [])
            errors = param_result.get("errors", [])

            if created:
                msg_lines.append("Nieuw aangemaakte parameters:")
                msg_lines.append("  " + ", ".join(created))
                msg_lines.append("")

            if filled:
                msg_lines.append("Ingevulde parameters ({0}):".format(len(filled)))
                for k, v in filled.items():
                    v_str = str(v)
                    if len(v_str) > 30:
                        v_str = v_str[:27] + "..."
                    msg_lines.append("  {0}: {1}".format(k, v_str))
            else:
                msg_lines.append("Geen parameters ingevuld.")

            if not_found:
                msg_lines.append("")
                msg_lines.append("Parameters niet aangemaakt:")
                msg_lines.append("  " + ", ".join(not_found))

            if errors:
                msg_lines.append("")
                msg_lines.append("Fouten:")
                for err in errors:
                    msg_lines.append("  " + err)

            self.DialogResult = True
            self.Close()

            forms.alert("\n".join(msg_lines), title="GIS2BIM - Resultaat")
        else:
            self.txt_status.Text = "Fout: {0}".format(set_result["message"])


def main():
    log("=== GIS2BIM Locatie Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = LocatieWindow(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM Locatie Tool beeindigd ===")


if __name__ == "__main__":
    main()
