# -*- coding: utf-8 -*-
"""
AHN Data Laden - GIS2BIM
========================

Laad AHN (Actueel Hoogtebestand Nederland) hoogtedata van PDOK
en maak een TopographySurface in Revit.

Twee methoden:
- WCS (GeoTIFF): Snel, klein bestand, geen externe tools
- LAZ (Puntenwolk): Hogere dichtheid, vereist LAStools
"""

__title__ = "AHN\nLaden"
__author__ = "OpenAEC Foundation"
__doc__ = "Laad AHN hoogtedata (maaiveld/oppervlakte) als Topography"

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
import tempfile
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
# view_setup niet meer nodig - TopographySurface is altijd 3D
from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui
from gis2bim.revit.geometry import rd_to_revit_xyz, create_textured_material, set_element_material

log = get_logger("AHN")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.ahn import AHNClient, AHNError
    from gis2bim.api.wms import WMSClient, WMS_LAYERS
    from gis2bim.parsers.geotiff import GeoTiffReader, GeoTiffError
    from gis2bim.parsers.las import LASReader, LASError
    from gis2bim.coordinates import create_bbox_rd
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


METER_TO_FEET = 1.0 / 0.3048


class AHNWindow(Window):
    """WPF Window voor AHN data laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.result_count = 0
        self._lastools_status = None

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location',
            'cmb_bbox_size',
            'rdo_dtm', 'rdo_dsm',
            'rdo_wcs', 'rdo_laz',
            'pnl_lastools', 'txt_lastools',
            'lbl_keep_nth', 'pnl_keep_nth', 'txt_keep_nth',
            'cmb_resolution', 'cmb_texture',
            'pnl_estimate', 'txt_estimate',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        self._bind_events()
        self._init_method_ui()
        self._set_default_keep_nth()
        self._update_estimate()

    def _init_method_ui(self):
        """Initialiseer UI staat voor de geselecteerde methode."""
        is_laz = hasattr(self, 'rdo_laz') and self.rdo_laz.IsChecked
        if is_laz:
            self.pnl_lastools.Visibility = Visibility.Visible
            self._check_lastools()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.cmb_bbox_size.SelectionChanged += self._on_settings_changed
        self.cmb_resolution.SelectionChanged += self._on_settings_changed
        self.rdo_wcs.Checked += self._on_method_changed
        self.rdo_laz.Checked += self._on_method_changed
        self.txt_keep_nth.TextChanged += self._on_settings_changed

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_settings_changed(self, sender, args):
        # Herbereken default bij bbox/resolutie wijziging
        # (maar niet als de gebruiker zelf een waarde heeft getypt)
        if sender != self.txt_keep_nth:
            self._set_default_keep_nth()
        self._update_estimate()

    def _on_method_changed(self, sender, args):
        is_laz = hasattr(self, 'rdo_laz') and self.rdo_laz.IsChecked
        laz_vis = Visibility.Visible if is_laz else Visibility.Collapsed
        self.pnl_lastools.Visibility = laz_vis
        self.lbl_keep_nth.Visibility = Visibility.Visible
        self.pnl_keep_nth.Visibility = Visibility.Visible
        if is_laz:
            self._check_lastools()
        self._set_default_keep_nth()
        self._update_estimate()

    def _check_lastools(self):
        try:
            client = AHNClient()
            log("LAStools zoekpaden: {0}".format(client.LASTOOLS_SEARCH_PATHS))
            status = client.get_lastools_status()
            self._lastools_status = status
            log("LAStools status: {0}".format(status["message"]))

            if status["available"]:
                self.txt_lastools.Text = status["message"]
                self.txt_lastools.Foreground = System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(69, 182, 168))
            else:
                self.txt_lastools.Text = (
                    "Geen LAStools gevonden.\n"
                    "Installeer LAStools in C:\\LAStools\\bin\\ "
                    "of voeg toe aan PATH."
                )
                self.txt_lastools.Foreground = System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(219, 76, 64))
        except Exception as e:
            self.txt_lastools.Text = "Fout: {0}".format(e)

    def _set_default_keep_nth(self):
        """Bereken een default keep_nth zodat ~5000 punten overblijven.

        WCS: keep_nth wordt toegepast NA de grid (pixels).
             effective = pixels^2 / keep_nth
        LAZ: keep_nth wordt toegepast VOOR de grid thinning.
             effective = raw_points / keep_nth (mits < grid^2)
        """
        try:
            bbox_size = self._get_bbox_size()
            resolution = self._get_resolution()
            is_laz = self._get_method() == "laz"
            target = 5000

            if is_laz:
                # LAZ: ~8 punten/m2, keep_nth vóór grid thinning
                raw_points = int(bbox_size * bbox_size * 8)
                keep_nth = max(2, raw_points // target)
            else:
                # WCS: pixels = (bbox_size / resolution)^2
                pixels = int(round(bbox_size / resolution))
                total = pixels * pixels
                keep_nth = max(2, total // target)

            self.txt_keep_nth.Text = str(keep_nth)
        except Exception:
            pass

    def _update_estimate(self):
        try:
            bbox_size = self._get_bbox_size()
            resolution = self._get_resolution()
            is_laz = self._get_method() == "laz"

            if is_laz:
                raw_points = int(bbox_size * bbox_size * 8)
                keep_nth = self._get_keep_every_nth()
                if keep_nth:
                    after_nth = raw_points // keep_nth
                else:
                    after_nth = raw_points
                pixels = int(round(bbox_size / resolution))
                grid_points = pixels * pixels
                final = min(after_nth, grid_points)
                parts = "~{0:,} ruwe punten".format(raw_points)
                if keep_nth:
                    parts += " -> elke {0}e = ~{1:,}".format(keep_nth, after_nth)
                parts += " -> grid {0}m = ~{1:,}".format(resolution, final)
                self.txt_estimate.Text = "Geschat: {0}".format(parts)
            else:
                pixels = int(round(bbox_size / resolution))
                total = pixels * pixels
                keep_nth = self._get_keep_every_nth()
                if keep_nth:
                    final = total // keep_nth
                    self.txt_estimate.Text = (
                        "Geschat: {0:,} pixels -> elke {1}e = ~{2:,} punten".format(
                            total, keep_nth, final)
                    )
                else:
                    self.txt_estimate.Text = (
                        "Geschat: ~{0:,} punten ({1}x{1} pixels)".format(
                            total, pixels)
                    )
        except Exception:
            pass

    def _on_execute(self, sender, args):
        if not self.location_rd:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        method = self._get_method()
        if method == "laz":
            if not self._lastools_status or not self._lastools_status["available"]:
                self.txt_status.Text = "LAStools niet gevonden. Installeer LAStools eerst."
                return

        show_progress(self, "AHN data laden...")
        self.btn_execute.IsEnabled = False

        try:
            if method == "laz":
                self._load_ahn_laz()
            else:
                self._load_ahn_wcs()
            self.DialogResult = True
            self.Close()
        except (AHNError, GeoTiffError, LASError) as e:
            log("AHN error: {0}".format(e))
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True
        except Exception as e:
            log("Error loading AHN: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    # =========================================================================
    # UI getters
    # =========================================================================

    def _get_bbox_size(self):
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 100

    def _get_coverage_type(self):
        if hasattr(self, 'rdo_dsm') and self.rdo_dsm.IsChecked:
            return "dsm"
        return "dtm"

    def _get_method(self):
        if hasattr(self, 'rdo_laz') and self.rdo_laz.IsChecked:
            return "laz"
        return "wcs"

    def _get_resolution(self):
        item = self.cmb_resolution.SelectedItem
        if item and hasattr(item, 'Tag'):
            return float(item.Tag)
        return 0.5

    def _get_keep_every_nth(self):
        if hasattr(self, 'txt_keep_nth'):
            text = self.txt_keep_nth.Text.strip()
            if text:
                try:
                    val = int(text)
                    if val >= 2:
                        return val
                except ValueError:
                    pass
        return None

    def _get_texture_type(self):
        """Haal geselecteerde textuur type op.

        Returns:
            "geen", "luchtfoto", of "bgt"
        """
        item = self.cmb_texture.SelectedItem
        if item and hasattr(item, 'Tag'):
            return str(item.Tag)
        return "geen"

    def _get_bbox(self):
        rd_x, rd_y = self.location_rd
        bbox_size = self._get_bbox_size()
        half_size = bbox_size / 2.0
        return (
            rd_x - half_size,
            rd_y - half_size,
            rd_x + half_size,
            rd_y + half_size
        )

    # =========================================================================
    # WCS (GeoTIFF) workflow
    # =========================================================================

    def _load_ahn_wcs(self):
        bbox = self._get_bbox()
        coverage = self._get_coverage_type()
        resolution = self._get_resolution()
        keep_nth = self._get_keep_every_nth()

        log("WCS: BBOX={0}, Coverage={1}, Resolutie={2}m, keep_nth={3}".format(
            bbox, coverage, resolution, keep_nth))

        show_progress(self, "Downloaden van PDOK ({0}, {1}m)...".format(
            coverage.upper(), resolution))

        client = AHNClient(timeout=60)
        tiff_path = client.download_geotiff(bbox, coverage, resolution)
        log("GeoTIFF gedownload: {0}".format(tiff_path))

        show_progress(self, "GeoTIFF verwerken...")

        reader = GeoTiffReader()
        grid = reader.read(tiff_path)
        log("Grid: {0}x{1} pixels".format(grid["width"], grid["height"]))

        show_progress(self, "Punten converteren...")

        xyz_points = reader.to_xyz_points(grid)
        log("{0} punten na filtering nodata".format(len(xyz_points)))

        if keep_nth and len(xyz_points) > 3:
            xyz_points = xyz_points[::keep_nth]
            log("{0} punten na keep_every_nth={1}".format(
                len(xyz_points), keep_nth))

        if len(xyz_points) < 3:
            raise AHNError(
                "Te weinig geldige punten ({0}). "
                "Mogelijk geen AHN data beschikbaar.".format(len(xyz_points)))

        show_progress(self, "Topography aanmaken ({0} punten)...".format(
            len(xyz_points)))

        self._create_topography(xyz_points, coverage)

    # =========================================================================
    # LAZ (Puntenwolk) workflow
    # =========================================================================

    def _load_ahn_laz(self):
        bbox = self._get_bbox()
        coverage = self._get_coverage_type()
        resolution = self._get_resolution()
        keep_nth = self._get_keep_every_nth()

        log("LAZ: BBOX={0}, Coverage={1}, Resolutie={2}m, keep_nth={3}".format(
            bbox, coverage, resolution, keep_nth))

        client = AHNClient(timeout=300)

        tools = client.find_lastools()
        found_tools = [n for n, p in tools.items() if p]
        log("LAStools gevonden: {0}".format(found_tools))

        def progress_cb(msg):
            show_progress(self, msg)

        laz_paths = client.download_laz_tiles(bbox, progress_callback=progress_cb)
        log("LAZ tiles gedownload: {0}".format(len(laz_paths)))

        all_points = []
        las_reader = LASReader()

        keep_classification = [2, 9] if coverage == "dtm" else None
        keep_highest = (coverage == "dsm")

        for i, laz_path in enumerate(laz_paths):
            show_progress(self, "Verwerken tile {0}/{1} met LAStools...".format(
                i + 1, len(laz_paths)))

            tile_name = os.path.splitext(os.path.basename(laz_path))[0]
            output_base = os.path.join(tempfile.gettempdir(),
                                       "ahn_processed_{0}".format(tile_name))

            result_path, fmt = client.process_laz(
                laz_path, output_base, bbox=bbox, tools=tools,
                keep_every_nth=keep_nth,
                keep_classification=keep_classification
            )
            log("Tile verwerkt: {0} ({1})".format(result_path, fmt))

            show_progress(self, "Punten lezen uit {0}...".format(fmt.upper()))

            if fmt == "xyz":
                points = las_reader.read_xyz_text(
                    result_path,
                    bbox=bbox,
                    thin_grid=resolution,
                    keep_highest=keep_highest,
                )
            else:
                points = las_reader.read(
                    result_path,
                    bbox=bbox,
                    thin_grid=resolution,
                    classification=keep_classification,
                    keep_highest=keep_highest,
                )

            log("Tile {0}: {1} punten".format(i + 1, len(points)))
            all_points.extend(points)

            try:
                os.remove(result_path)
            except OSError:
                pass

        log("Totaal: {0} punten uit {1} tiles".format(
            len(all_points), len(laz_paths)))

        if len(all_points) < 3:
            raise AHNError(
                "Te weinig geldige punten ({0}). "
                "Mogelijk geen AHN data beschikbaar of LAStools "
                "kon het bestand niet verwerken.".format(len(all_points)))

        show_progress(self, "Topography aanmaken ({0} punten)...".format(
            len(all_points)))

        self._create_topography(all_points, coverage)

    # =========================================================================
    # Revit Topography aanmaak
    # =========================================================================

    def _create_topography(self, xyz_points, coverage):
        from System.Collections.Generic import List as GenericList

        origin_x, origin_y = self.location_rd

        revit_points = GenericList[DB.XYZ]()
        for rd_x, rd_y, z_m in xyz_points:
            xyz = rd_to_revit_xyz(rd_x, rd_y, origin_x, origin_y, z_m)
            revit_points.Add(xyz)

        log("Revit punten: {0}".format(revit_points.Count))

        # Download textuur afbeelding (BUITEN transactie)
        texture_type = self._get_texture_type()
        log("Textuur type: {0}".format(texture_type))
        image_path = None
        if texture_type != "geen":
            image_path = self._download_texture_image(texture_type)
            if image_path:
                log("Textuur afbeelding: {0} ({1} bytes)".format(
                    image_path,
                    os.path.getsize(image_path) if os.path.exists(image_path) else 0))
            else:
                log("WAARSCHUWING: Textuur download mislukt")

        with revit.Transaction("GIS2BIM AHN {0}".format(coverage.upper())):
            topo = self._create_topography_surface(revit_points)

            if topo:
                self.result_count = revit_points.Count
                log("Topography aangemaakt: {0} punten".format(revit_points.Count))

                # Textuur materiaal toepassen
                if image_path:
                    self._apply_texture(topo, image_path, texture_type)
            else:
                raise AHNError("Kon geen topography aanmaken")

    def _create_topography_surface(self, revit_points):
        try:
            from Autodesk.Revit.DB.Architecture import TopographySurface
            topo = TopographySurface.Create(self.doc, revit_points)
            log("TopographySurface aangemaakt")
            return topo
        except Exception as e:
            log("TopographySurface.Create fout: {0}".format(e))
            return None

    # =========================================================================
    # Textuur download en toepassing
    # =========================================================================

    def _download_texture_image(self, texture_type):
        """Download WMS afbeelding voor textuur.

        Args:
            texture_type: "luchtfoto" of "bgt"

        Returns:
            Pad naar gedownloade afbeelding, of None bij fout
        """
        bbox = self._get_bbox()

        if texture_type == "luchtfoto":
            layer = WMS_LAYERS.get("luchtfoto_actueel")
            label = "luchtfoto"
        else:
            return None

        if not layer:
            return None

        show_progress(self, "{0} downloaden...".format(label))
        update_ui()

        try:
            client = WMSClient(timeout=60)

            # Sla op in GIS2BIM appdata map (persistent)
            appdata = os.path.join(
                os.environ.get("APPDATA", tempfile.gettempdir()),
                "GIS2BIM", "textures")
            if not os.path.exists(appdata):
                os.makedirs(appdata)

            bbox_str = "{0}_{1}_{2}_{3}".format(
                int(bbox[0]), int(bbox[1]), int(bbox[2]), int(bbox[3]))
            ext = ".jpg" if "jpeg" in layer.get("format", "") else ".png"
            filename = "ahn_{0}_{1}{2}".format(texture_type, bbox_str, ext)
            output_path = os.path.join(appdata, filename)

            # Cache: hergebruik bestaande afbeelding
            if os.path.exists(output_path) and os.path.getsize(output_path) > 1000:
                log("Textuur uit cache: {0}".format(output_path))
                return output_path

            image_path = client.download_image(layer, bbox, output_path)
            log("Textuur gedownload: {0}".format(image_path))
            return image_path

        except Exception as e:
            log("Textuur download fout: {0}".format(e))
            log(traceback.format_exc())
            # Niet fataal: ga door zonder textuur
            return None

    def _apply_texture(self, topo, image_path, texture_type):
        """Pas textuur materiaal toe op TopographySurface.

        Probeert achtereenvolgens:
        1. doc.Paint() op faces
        2. TopographySurface materiaal parameter
        3. TopographySurface type materiaal

        Args:
            topo: TopographySurface element
            image_path: Pad naar textuur afbeelding
            texture_type: "luchtfoto" of "bgt"
        """
        bbox_size = self._get_bbox_size()
        label = "Luchtfoto" if texture_type == "luchtfoto" else "BGT"
        mat_name = "AHN - {0}".format(label)

        show_progress(self, "Textuur materiaal toepassen...")
        update_ui()

        log("Textuur toepassen: afbeelding={0}, "
            "grootte={1}m".format(image_path, bbox_size))

        try:
            mat_id = create_textured_material(
                self.doc, mat_name, image_path,
                float(bbox_size), float(bbox_size))
            log("Materiaal aangemaakt: {0} (ID={1})".format(
                mat_name, mat_id.IntegerValue))

            # Methode 1: Paint op faces
            log("Poging 1: Paint op faces...")
            success = set_element_material(self.doc, topo, mat_id)
            if success:
                log("Textuur via Paint ingesteld")
                return

            # Methode 2: TopographySurface materiaal parameter
            log("Poging 2: Materiaal parameter...")
            if self._set_topo_material_param(topo, mat_id):
                log("Textuur via parameter ingesteld")
                return

            # Methode 3: TopographySurface type materiaal
            log("Poging 3: TopographySurface type materiaal...")
            if self._set_topo_surface_type_material(topo, mat_id):
                log("Textuur via TopographySurface type ingesteld")
                return

            log("Geen methode werkte voor materiaal toewijzing. "
                "Materiaal '{0}' is wel aangemaakt - handmatig "
                "toewijzen via Revit.".format(mat_name))

        except Exception as e:
            log("Textuur toepassen fout: {0}".format(e))
            log(traceback.format_exc())

    def _set_topo_material_param(self, topo, mat_id):
        """Probeer materiaal in te stellen via parameter.

        Probeert verschillende benaderingen:
        1. MATERIAL_ID_PARAM BuiltInParameter
        2. LookupParameter op naam
        3. Type materiaal parameter

        Returns:
            True als succesvol
        """
        # Poging 1: BuiltInParameter
        try:
            from Autodesk.Revit.DB import BuiltInParameter
            param = topo.get_Parameter(
                BuiltInParameter.MATERIAL_ID_PARAM)
            if param and not param.IsReadOnly:
                param.Set(mat_id)
                log("Materiaal ingesteld via MATERIAL_ID_PARAM")
                return True
            else:
                log("MATERIAL_ID_PARAM: read-only of niet gevonden")
        except Exception as e:
            log("MATERIAL_ID_PARAM fout: {0}".format(e))

        # Poging 2: LookupParameter op naam
        try:
            for param_name in ["Material", "Materiaal",
                               "Phasierung - Material",
                               "Surface Material"]:
                param = topo.LookupParameter(param_name)
                if param and not param.IsReadOnly:
                    param.Set(mat_id)
                    log("Materiaal ingesteld via LookupParameter "
                        "'{0}'".format(param_name))
                    return True
        except Exception as e:
            log("LookupParameter fout: {0}".format(e))

        # Poging 3: Type parameter
        try:
            topo_type = self.doc.GetElement(topo.GetTypeId())
            if topo_type:
                from Autodesk.Revit.DB import BuiltInParameter
                param = topo_type.get_Parameter(
                    BuiltInParameter.MATERIAL_ID_PARAM)
                if param and not param.IsReadOnly:
                    param.Set(mat_id)
                    log("Materiaal ingesteld via type MATERIAL_ID_PARAM")
                    return True
        except Exception as e:
            log("Type parameter fout: {0}".format(e))

        # Log alle beschikbare parameters voor debugging
        log("Geen schrijfbare materiaal parameter gevonden")
        try:
            for p in topo.Parameters:
                if p and hasattr(p, 'Definition'):
                    is_ro = "RO" if p.IsReadOnly else "RW"
                    log("  param: {0} [{1}] {2}".format(
                        p.Definition.Name, p.StorageType, is_ro))
        except Exception:
            pass
        return False

    def _set_topo_surface_type_material(self, topo, mat_id):
        """Stel materiaal in op TopographySurface type.

        Dupliceert het type om andere surfaces niet te beinvloeden,
        en stelt het materiaal in via alle beschikbare parameters.

        Returns:
            True als succesvol
        """
        try:
            from Autodesk.Revit.DB import BuiltInParameter

            topo_type = self.doc.GetElement(topo.GetTypeId())
            if not topo_type:
                log("Geen type gevonden voor TopographySurface")
                return False

            # Dupliceer type
            type_name = "AHN - Textuur"
            try:
                new_type = topo_type.Duplicate(type_name)
                log("TopographySurface type gedupliceerd: "
                    "{0}".format(type_name))
            except Exception:
                new_type = topo_type
                log("Gebruik bestaand type (duplicaat bestond al)")

            # Probeer materiaal op alle parameters
            set_ok = False
            for p in new_type.Parameters:
                if p and not p.IsReadOnly:
                    try:
                        if (p.StorageType.ToString() == "ElementId"
                                and "aterial" in p.Definition.Name):
                            p.Set(mat_id)
                            log("Type materiaal param ingesteld: "
                                "{0}".format(p.Definition.Name))
                            set_ok = True
                    except Exception:
                        pass

            # Probeer specifieke BuiltInParameters op het type
            if not set_ok:
                for bip in [BuiltInParameter.MATERIAL_ID_PARAM,
                            BuiltInParameter.PHY_MATERIAL_PARAM_FINISH_FACE_1,
                            BuiltInParameter.STRUCTURAL_MATERIAL_PARAM]:
                    try:
                        param = new_type.get_Parameter(bip)
                        if param and not param.IsReadOnly:
                            param.Set(mat_id)
                            log("Type materiaal via BIP {0}".format(
                                bip.ToString()))
                            set_ok = True
                            break
                    except Exception:
                        pass

            # Wijs nieuwe type toe
            if new_type.Id != topo.GetTypeId():
                topo.ChangeTypeId(new_type.Id)
                log("Type gewijzigd naar: {0}".format(type_name))

            return set_ok

        except Exception as e:
            log("TopographySurface type materiaal fout: {0}".format(e))
            log(traceback.format_exc())
            return False


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

    window = AHNWindow(doc)
    result = window.ShowDialog()

    if result and window.result_count > 0:
        forms.alert(
            "AHN Topography aangemaakt met {0:,} punten.".format(
                window.result_count),
            title="GIS2BIM AHN",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
