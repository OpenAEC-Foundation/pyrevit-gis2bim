# -*- coding: utf-8 -*-
"""
Natura 2000 - GIS2BIM
======================

Haal Natura 2000 gebieden op via WFS, bereken de minimale afstand
tot het dichtstbijzijnde beschermde natuurgebied, sla resultaten
op in project parameters en plaats optioneel:
- Luchtfoto als achtergrond (WMS, meetbaar)
- Filled regions van Natura 2000 polygonen (meetbaar in Revit)
- Text labels met gebiedsnamen per Natura 2000 gebied

Alles op schaal 1:1000 vanwege Revit's 30.000 ft image width limiet.
"""

__title__ = "Natura\n2000"
__author__ = "OpenAEC Foundation"
__doc__ = "Natura 2000 kaart en afstandsberekening tot beschermde natuurgebieden"

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
from bm_logger import get_logger
from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements
from gis2bim.ui.location_setup import setup_project_location
from gis2bim.ui.view_setup import populate_view_dropdown, get_selected_view
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

log = get_logger("Natura2000")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.natura2000 import Natura2000Client
    from gis2bim.api.wms import WMSClient, WMS_LAYERS
    from gis2bim.coordinates import create_bbox_rd
    from gis2bim.revit.location import _create_project_parameter
    from gis2bim.revit.geometry import (
        get_filled_region_type,
        _get_default_text_type,
    )
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# Conversieconstanten
METER_TO_FEET = 1.0 / 0.3048
IMAGE_SCALE = 1000  # Plaatsing op schaal 1:1000 (Revit limiet 30.000 ft)

# Project parameter namen
PARAM_OVERZICHT = "GIS2BIM_natura2000_overzicht"
PARAM_AFSTAND = "GIS2BIM_natura2000_afstand"


def _rd_to_scaled_xyz(rd_x, rd_y, origin_x, origin_y):
    """RD coordinaten naar Revit XYZ op schaal 1:1000."""
    local_x = (rd_x - origin_x) / IMAGE_SCALE * METER_TO_FEET
    local_y = (rd_y - origin_y) / IMAGE_SCALE * METER_TO_FEET
    return DB.XYZ(local_x, local_y, 0)


def _is_ccw(ring):
    """Check of een polygon ring counter-clockwise is (shoelace formule)."""
    area = 0.0
    n = len(ring)
    for i in range(n):
        j = (i + 1) % n
        area += ring[i][0] * ring[j][1]
        area -= ring[j][0] * ring[i][1]
    return area > 0


class Natura2000Window(Window):
    """WPF Window voor Natura 2000 kaart en afstandsberekening."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location', 'cmb_bbox_size',
            'chk_luchtfoto', 'chk_filled_regions', 'cmb_view',
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
        log("Execute gestart")

        if self.location_rd is None:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        try:
            show_progress(self, "Voorbereiden...")
            self.btn_execute.IsEnabled = False

            search_radius = self._get_search_radius()
            plaats_luchtfoto = self.chk_luchtfoto.IsChecked
            plaats_filled = self.chk_filled_regions.IsChecked
            view = get_selected_view(self.cmb_view, self.doc)

            needs_view = plaats_luchtfoto or plaats_filled
            if needs_view and view is None:
                self.txt_status.Text = "Selecteer een view voor kaartplaatsing"
                hide_progress(self)
                self.btn_execute.IsEnabled = True
                return

            center_x, center_y = self.location_rd
            log("Center: {0}, {1} - Zoekstraal: {2}m".format(
                center_x, center_y, search_radius))

            # WFS query: Natura 2000 gebieden ophalen
            show_progress(self, "Natura 2000 gebieden ophalen van PDOK...")
            update_ui()

            client = Natura2000Client()
            result = client.get_natura2000_info(
                (center_x, center_y), search_radius)

            log("Gebieden gevonden: {0}".format(len(result.areas)))
            log("Gebiedsnamen: {0}".format(result.area_names))
            log("Min afstand: {0}".format(result.min_distance))
            log("Dichtstbij: {0}".format(result.nearest_name))

            # Project parameters opslaan
            show_progress(self, "Project parameters opslaan...")
            update_ui()

            self._save_project_parameters(result)

            bbox_size = search_radius * 2

            # Luchtfoto als achtergrond plaatsen
            if plaats_luchtfoto and view is not None:
                show_progress(self, "Luchtfoto downloaden en plaatsen...")
                update_ui()
                self._place_luchtfoto(view, center_x, center_y, bbox_size)

            # Filled regions + text labels tekenen
            filled_getekend = False
            if plaats_filled and view is not None and result.areas:
                show_progress(self, "Natura 2000 gebieden tekenen...")
                update_ui()
                self._draw_filled_regions(view, result, center_x, center_y)

                show_progress(self, "Gebiedsnamen plaatsen...")
                update_ui()
                self._draw_area_labels(view, result, center_x, center_y)
                filled_getekend = True

            hide_progress(self)
            self.DialogResult = True
            self.Close()

            self._show_result(result, search_radius,
                              plaats_luchtfoto, filled_getekend)

        except Exception as e:
            log("Execute error: {0}\n{1}".format(e, traceback.format_exc()))
            self.txt_status.Text = "Fout: {0}".format(str(e))
            hide_progress(self)
            self.btn_execute.IsEnabled = True

    def _get_search_radius(self):
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 10000

    def _save_project_parameters(self, result):
        """Sla Natura 2000 resultaten op in project parameters."""
        # Gebiedsnamen overzicht
        overzicht = "; ".join(result.area_names) if result.area_names else "Geen gebieden gevonden"

        # Afstand als geheel getal in meters
        if result.min_distance is not None:
            afstand_str = str(int(round(result.min_distance)))
        else:
            afstand_str = "N/A"

        param_values = {
            PARAM_OVERZICHT: overzicht,
            PARAM_AFSTAND: afstand_str,
        }

        project_info = self.doc.ProjectInformation

        # Fase 1: Parameters aanmaken indien nodig
        params_to_create = []
        for param_name in param_values:
            param = project_info.LookupParameter(param_name)
            if param is None:
                params_to_create.append(param_name)

        if params_to_create:
            log("Parameters aanmaken: {0}".format(params_to_create))
            t_create = DB.Transaction(self.doc, "GIS2BIM - Natura 2000 Parameters Aanmaken")
            t_create.Start()
            try:
                for param_name in params_to_create:
                    success = _create_project_parameter(self.doc, param_name)
                    log("Parameter {0} aangemaakt: {1}".format(param_name, success))
                t_create.Commit()
                # Refresh project info reference
                project_info = self.doc.ProjectInformation
            except Exception as e:
                t_create.RollBack()
                log("Fout bij aanmaken parameters: {0}".format(e))

        # Fase 2: Parameters vullen
        t_fill = DB.Transaction(self.doc, "GIS2BIM - Natura 2000 Parameters Vullen")
        t_fill.Start()
        try:
            for param_name, value in param_values.items():
                param = project_info.LookupParameter(param_name)
                if param is not None and not param.IsReadOnly:
                    param.Set(str(value))
                    log("Parameter {0} = {1}".format(param_name, value))
                else:
                    log("Parameter {0} niet gevonden of read-only".format(param_name))
            t_fill.Commit()
        except Exception as e:
            t_fill.RollBack()
            log("Fout bij vullen parameters: {0}".format(e))

    def _place_luchtfoto(self, view, center_x, center_y, bbox_size):
        """Download en plaats luchtfoto als achtergrond op schaal 1:1000."""
        layer = WMS_LAYERS.get("luchtfoto_actueel")
        if layer is None:
            log("WMS layer 'luchtfoto_actueel' niet gevonden")
            return

        bbox = create_bbox_rd(center_x, center_y, bbox_size)
        log("Luchtfoto bbox: {0}".format(bbox))

        wms_client = WMSClient()
        image_path = wms_client.download_image(layer, bbox)
        log("Luchtfoto gedownload: {0}".format(image_path))

        t = DB.Transaction(self.doc, "GIS2BIM - Luchtfoto Plaatsen")
        t.Start()
        try:
            # Verwijder bestaande rasterbeelden in de view
            self._remove_existing_images(view)

            # Plaats nieuw beeld
            options = DB.ImageTypeOptions(image_path, False, DB.ImageTypeSource.Import)
            image_type = DB.ImageType.Create(self.doc, options)

            placement = DB.ImagePlacementOptions()
            placement.PlacementPoint = DB.BoxPlacement.Center
            placement.Location = DB.XYZ(0, 0, 0)

            image_instance = DB.ImageInstance.Create(
                self.doc, view, image_type.Id, placement
            )

            # Schaal 1:1000 — Revit limiet is 30.000 ft (~9.1 km)
            # Bij 1:1000 wordt 40km bbox -> 40m in Revit -> ~131 ft
            desired_width_ft = (bbox_size / IMAGE_SCALE) * METER_TO_FEET
            current_width = image_instance.Width
            if current_width > 0:
                image_instance.Width = desired_width_ft
                log("Luchtfoto geschaald naar {0:.1f}ft (bbox {1}m, schaal 1:{2})".format(
                    desired_width_ft, bbox_size, IMAGE_SCALE))

            t.Commit()
            log("Luchtfoto geplaatst in view")

        except Exception as e:
            t.RollBack()
            log("Fout bij plaatsen luchtfoto: {0}".format(e))
            log(traceback.format_exc())

        # Opruimen temp bestand
        try:
            if os.path.exists(image_path):
                os.remove(image_path)
        except Exception:
            pass

    def _draw_filled_regions(self, view, result, origin_x, origin_y):
        """Teken filled regions voor Natura 2000 polygonen op schaal 1:1000."""
        from System.Collections.Generic import List as GenericList

        # Zoek filled region type "Natura2000"
        fr_type = get_filled_region_type(self.doc, "Natura2000")
        if fr_type is None:
            log("FilledRegionType 'Natura2000' niet gevonden")
            return

        log("FilledRegionType: {0}".format(fr_type.Id))

        t = DB.Transaction(self.doc, "GIS2BIM - Natura 2000 Filled Regions")
        t.Start()
        try:
            created = 0
            skipped = 0
            errors = 0

            for area in result.areas:
                for ring in area.polygons:
                    if len(ring) < 3:
                        skipped += 1
                        continue

                    try:
                        # Ensure CCW winding voor outer boundary
                        pts = list(ring)
                        if not _is_ccw(pts):
                            pts = pts[::-1]

                        # Converteer naar geschaalde Revit XYZ
                        raw_points = []
                        for pt in pts:
                            xyz = _rd_to_scaled_xyz(
                                pt[0], pt[1], origin_x, origin_y)
                            raw_points.append(xyz)

                        # Filter opeenvolgende duplicaten (< 0.003 ft ~ 1mm)
                        points = [raw_points[0]]
                        for j in range(1, len(raw_points)):
                            if raw_points[j].DistanceTo(points[-1]) >= 0.003:
                                points.append(raw_points[j])

                        # Verwijder sluitpunt als dicht bij eerste punt
                        if len(points) > 1 and \
                                points[0].DistanceTo(points[-1]) < 0.003:
                            points = points[:-1]

                        if len(points) < 3:
                            skipped += 1
                            continue

                        # Bouw gesloten CurveLoop
                        curve_loop = DB.CurveLoop()
                        for i in range(len(points)):
                            p1 = points[i]
                            p2 = points[(i + 1) % len(points)]
                            line = DB.Line.CreateBound(p1, p2)
                            curve_loop.Append(line)

                        loops = GenericList[DB.CurveLoop]()
                        loops.Add(curve_loop)

                        DB.FilledRegion.Create(
                            self.doc, fr_type.Id, view.Id, loops
                        )
                        created += 1

                    except Exception as e:
                        errors += 1
                        if errors <= 5:
                            log("Fout bij filled region: {0}".format(e))

            t.Commit()
            log("Filled regions: {0} aangemaakt, {1} overgeslagen, {2} errors".format(
                created, skipped, errors))

        except Exception as e:
            t.RollBack()
            log("Fout bij filled regions transactie: {0}".format(e))
            log(traceback.format_exc())

    def _draw_area_labels(self, view, result, origin_x, origin_y):
        """Teken text labels met gebiedsnamen op schaal 1:1000."""
        text_type = _get_default_text_type(self.doc)
        if text_type is None:
            log("Geen TextNoteType beschikbaar")
            return

        # Groepeer polygonpunten per unieke gebiedsnaam voor centroid
        name_points = {}
        for area in result.areas:
            if area.name not in name_points:
                name_points[area.name] = []
            for ring in area.polygons:
                for pt in ring:
                    name_points[area.name].append(pt)

        t = DB.Transaction(self.doc, "GIS2BIM - Natura 2000 Labels")
        t.Start()
        try:
            created = 0
            for name, points in name_points.items():
                if not points:
                    continue

                # Bereken centroid (gemiddelde x,y van alle polygonpunten)
                avg_x = sum(p[0] for p in points) / len(points)
                avg_y = sum(p[1] for p in points) / len(points)

                xyz = _rd_to_scaled_xyz(avg_x, avg_y, origin_x, origin_y)

                try:
                    DB.TextNote.Create(
                        self.doc, view.Id, xyz, name, text_type.Id
                    )
                    created += 1
                except Exception as e:
                    log("Fout bij label '{0}': {1}".format(name, e))

            t.Commit()
            log("Labels: {0} geplaatst".format(created))

        except Exception as e:
            t.RollBack()
            log("Fout bij labels transactie: {0}".format(e))
            log(traceback.format_exc())

    def _remove_existing_images(self, view):
        """Verwijder bestaande rasterbeelden uit een view."""
        try:
            collector = DB.FilteredElementCollector(self.doc, view.Id)
            images = collector.OfCategory(
                DB.BuiltInCategory.OST_RasterImages
            ).WhereElementIsNotElementType().ToElements()

            for img in images:
                try:
                    self.doc.Delete(img.Id)
                    log("Bestaand beeld verwijderd")
                except Exception:
                    pass
        except Exception:
            pass

    def _show_result(self, result, search_radius,
                     luchtfoto_geplaatst, filled_getekend):
        """Toon resultaat dialog."""
        msg_lines = [
            "Natura 2000 analyse voltooid!",
            "",
            "Zoekgebied: {0} km straal".format(search_radius / 1000),
            "Gebieden gevonden: {0}".format(len(result.area_names)),
        ]

        if result.area_names:
            msg_lines.append("")
            msg_lines.append("Gevonden gebieden:")
            for name in result.area_names:
                msg_lines.append("  - {0}".format(name))

        msg_lines.append("")
        if result.min_distance is not None:
            if result.min_distance == 0:
                msg_lines.append("Afstand: 0 m (project ligt BINNEN {0})".format(
                    result.nearest_name))
            else:
                msg_lines.append("Minimale afstand: {0} m".format(
                    int(round(result.min_distance))))
                msg_lines.append("Dichtstbijzijnd: {0}".format(
                    result.nearest_name))
        else:
            msg_lines.append("Geen Natura 2000 gebieden gevonden binnen zoekgebied")

        msg_lines.append("")
        msg_lines.append("Project parameters bijgewerkt:")
        msg_lines.append("  - {0}".format(PARAM_OVERZICHT))
        msg_lines.append("  - {0}".format(PARAM_AFSTAND))

        if luchtfoto_geplaatst:
            msg_lines.append("")
            msg_lines.append("Luchtfoto geplaatst in view (schaal 1:1000)")

        if filled_getekend:
            msg_lines.append("")
            msg_lines.append("Filled regions + labels getekend in view")

        forms.alert("\n".join(msg_lines), title="GIS2BIM - Natura 2000")


def main():
    log("=== GIS2BIM Natura 2000 Tool gestart ===")

    if not GIS2BIM_LOADED:
        forms.alert(
            "Kon GIS2BIM modules niet laden:\n\n{0}".format(IMPORT_ERROR),
            title="GIS2BIM - Fout"
        )
        return

    doc = revit.doc
    log("Document: {0}".format(doc.Title))

    try:
        window = Natura2000Window(doc)
        window.ShowDialog()
    except Exception as e:
        log("Window error: {0}\n{1}".format(e, traceback.format_exc()))
        forms.alert(
            "Fout bij laden window:\n\n{0}".format(str(e)),
            title="GIS2BIM - Fout"
        )

    log("=== GIS2BIM Natura 2000 Tool beeindigd ===")


if __name__ == "__main__":
    main()
