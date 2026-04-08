# -*- coding: utf-8 -*-
"""
CityJSON 3D Data Laden - GIS2BIM
==================================

Laad 3D gebouwen en omgevingsdata (wegen, water, terrein) vanuit CityJSON.

Data bronnen:
- PDOK 3D Basisvoorziening (automatisch downloaden)
- Lokaal CityJSON bestand (.json / .city.json)

Workflow:
1. Download of selecteer CityJSON bestand
2. Parse CityJSON objecten (gebouwen, wegen, water, terrein)
3. Maak DirectShape elementen in Revit met kleuren per type
4. Stel parameters in voor gebouwen (bouwjaar, status, ID)
"""

__title__ = "CityJSON\n3D"
__author__ = "OpenAEC Foundation"
__doc__ = "Laad 3D CityJSON gebouwen en omgevingsdata"

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
from gis2bim.revit.geometry import (
    rd_to_revit_xyz, get_or_create_material, set_element_material,
    set_element_parameter,
)

log = get_logger("CityJSON")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.parsers.cityjson import (
        CityJSONParser, CityJSONObject, CityJSONError,
        CITYJSON_KLEUREN, MATERIAL_CATEGORIE,
    )
    from gis2bim.api.pdok3d import PDOK3DClient, PDOK3DError
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# =============================================================================
# Materiaal mapping
# =============================================================================

MATERIAL_MAP = {
    "Gebouw":  ("CityJSON - Gebouw",  (200, 180, 160)),
    "Weg":     ("CityJSON - Weg",     (100, 100, 100)),
    "Water":   ("CityJSON - Water",   (100, 150, 220)),
    "Terrein": ("CityJSON - Terrein", (160, 140, 120)),
    "Groen":   ("CityJSON - Groen",   (80, 160, 60)),
    "Overig":  ("CityJSON - Overig",  (200, 200, 200)),
}


# =============================================================================
# CityJSON Parameters
# =============================================================================

_cityjson_params_created = False


def ensure_cityjson_parameters(doc):
    """Zorg dat CityJSON parameters bestaan op GenericModel categorie.

    Maakt instance parameters aan voor gebouw-attributen:
    - gebouw_id (text): BAG/kadaster identificatie
    - bouwjaar (number): Oorspronkelijk bouwjaar
    - gebouw_status (text): Pand status

    Args:
        doc: Revit document

    Returns:
        True als succesvol
    """
    global _cityjson_params_created
    if _cityjson_params_created:
        return True

    from Autodesk.Revit.DB import (
        ExternalDefinitionCreationOptions, BuiltInCategory
    )

    app = doc.Application

    param_defs = [
        ("gebouw_id", "text"),
        ("bouwjaar", "number"),
        ("gebouw_status", "text"),
    ]

    # Check welke al bestaan
    existing = set()
    bm = doc.ParameterBindings
    it = bm.ForwardIterator()
    while it.MoveNext():
        existing.add(it.Key.Name)

    needed = [(n, t) for n, t in param_defs if n not in existing]
    if not needed:
        _cityjson_params_created = True
        return True

    # Bewaar origineel shared parameter file
    original_spf = ""
    try:
        original_spf = app.SharedParametersFilename
    except Exception:
        pass

    try:
        temp_dir = os.environ.get(
            "TEMP", os.path.join(
                os.path.expanduser("~"), "AppData", "Local", "Temp"))
        temp_path = os.path.join(temp_dir, "GIS2BIM_SharedParams.txt")

        if not os.path.exists(temp_path):
            with open(temp_path, 'w') as f:
                pass

        app.SharedParametersFilename = temp_path
        def_file = app.OpenSharedParameterFile()

        # Zoek of maak groep
        group = None
        for g in def_file.Groups:
            if g.Name == "GIS2BIM":
                group = g
                break
        if group is None:
            group = def_file.Groups.Create("GIS2BIM")

        # Categorie set: GenericModel
        cat_set = app.Create.NewCategorySet()
        gm_cat = doc.Settings.Categories.get_Item(
            BuiltInCategory.OST_GenericModel)
        cat_set.Insert(gm_cat)

        for param_name, type_key in needed:
            # Zoek bestaande definitie in groep
            definition = None
            for d in group.Definitions:
                if d.Name == param_name:
                    definition = d
                    break

            if definition is None:
                opts = _create_param_options(param_name, type_key)
                definition = group.Definitions.Create(opts)

            # Bind als instance parameter
            binding = app.Create.NewInstanceBinding(cat_set)
            _bind_parameter(doc, bm, definition, binding)

        doc.Regenerate()
        _cityjson_params_created = True
        return True

    except Exception as e:
        log("ensure_cityjson_parameters fout: {0}".format(e))
        return False
    finally:
        try:
            if original_spf:
                app.SharedParametersFilename = original_spf
        except Exception:
            pass


def _create_param_options(name, type_key):
    """Maak ExternalDefinitionCreationOptions (version-safe)."""
    from Autodesk.Revit.DB import ExternalDefinitionCreationOptions

    try:
        from Autodesk.Revit.DB import SpecTypeId
        if type_key == "number":
            return ExternalDefinitionCreationOptions(
                name, SpecTypeId.Number)
        else:
            return ExternalDefinitionCreationOptions(
                name, SpecTypeId.String.Text)
    except (ImportError, AttributeError):
        pass

    from Autodesk.Revit.DB import ParameterType as RevitPT
    if type_key == "number":
        return ExternalDefinitionCreationOptions(name, RevitPT.Number)
    else:
        return ExternalDefinitionCreationOptions(name, RevitPT.Text)


def _bind_parameter(doc, binding_map, definition, binding):
    """Bind parameter aan document (version-safe)."""
    try:
        from Autodesk.Revit.DB import GroupTypeId
        binding_map.Insert(definition, binding, GroupTypeId.Data)
        return
    except (ImportError, AttributeError):
        pass

    try:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        binding_map.Insert(
            definition, binding, BuiltInParameterGroup.PG_DATA)
        return
    except Exception:
        pass

    binding_map.Insert(definition, binding)


# =============================================================================
# Main Window
# =============================================================================

class CityJSONWindow(Window):
    """WPF Window voor CityJSON 3D data laden."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.result_count = 0

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location',
            'rdo_source_pdok', 'rdo_source_local',
            'pnl_pdok_options', 'cmb_bbox_size', 'cmb_collection',
            'pnl_local_options', 'txt_filepath', 'btn_browse',
            'rdo_lod22', 'rdo_lod13', 'rdo_lod12',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        self._bind_events()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.btn_browse.Click += self._on_browse
        self.rdo_source_pdok.Checked += self._on_source_changed
        self.rdo_source_local.Checked += self._on_source_changed

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_source_changed(self, sender, args):
        """Wissel tussen PDOK en lokaal bestand opties."""
        if self.rdo_source_pdok.IsChecked:
            self.pnl_pdok_options.Visibility = Visibility.Visible
            self.pnl_local_options.Visibility = Visibility.Collapsed
        else:
            self.pnl_pdok_options.Visibility = Visibility.Collapsed
            self.pnl_local_options.Visibility = Visibility.Visible

    def _on_browse(self, sender, args):
        """Open bestandskiezer voor lokaal CityJSON bestand."""
        from Microsoft.Win32 import OpenFileDialog

        dlg = OpenFileDialog()
        dlg.Title = "Selecteer CityJSON bestand"
        dlg.Filter = (
            "CityJSON bestanden (*.json;*.city.json)|*.json;*.city.json|"
            "Alle bestanden (*.*)|*.*"
        )

        if dlg.ShowDialog():
            self.txt_filepath.Text = dlg.FileName

    def _on_execute(self, sender, args):
        """Start het laden van CityJSON data."""
        # Validatie
        is_pdok = self.rdo_source_pdok.IsChecked
        if is_pdok and not self.location_rd:
            self.txt_status.Text = "Geen locatie beschikbaar voor PDOK download"
            return

        if not is_pdok:
            filepath = self.txt_filepath.Text.strip()
            if not filepath or not os.path.exists(filepath):
                self.txt_status.Text = "Selecteer een geldig CityJSON bestand"
                return

        show_progress(self, "CityJSON data laden...")
        self.btn_execute.IsEnabled = False

        try:
            self._load_cityjson()
            self.DialogResult = True
            self.Close()
        except (CityJSONError, PDOK3DError) as e:
            log("CityJSON error: {0}".format(e))
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True
        except Exception as e:
            log("Error loading CityJSON: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    # =========================================================================
    # UI getters
    # =========================================================================

    def _get_lod(self):
        """Haal geselecteerde LoD waarde op."""
        if hasattr(self, 'rdo_lod13') and self.rdo_lod13.IsChecked:
            return "1.3"
        if hasattr(self, 'rdo_lod12') and self.rdo_lod12.IsChecked:
            return "1.2"
        return "2.2"

    def _get_lod_label(self):
        lod = self._get_lod()
        labels = {
            "1.2": "LoD 1.2",
            "1.3": "LoD 1.3",
            "2.2": "LoD 2.2",
        }
        return labels.get(lod, lod)

    def _get_bbox_size(self):
        """Haal geselecteerde zoekgebied grootte op.

        Returns:
            Grootte in meters, of 0 voor hele tile (geen crop)
        """
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 200

    def _get_bbox(self):
        """Bereken bounding box rond projectlocatie.

        Returns:
            (xmin, ymin, xmax, ymax) in RD, of None voor hele tile
        """
        bbox_size = self._get_bbox_size()
        if bbox_size == 0 or not self.location_rd:
            return None

        rd_x, rd_y = self.location_rd
        half = bbox_size / 2.0
        return (
            rd_x - half,
            rd_y - half,
            rd_x + half,
            rd_y + half,
        )

    def _get_collection(self):
        """Haal geselecteerde PDOK collectie op."""
        item = self.cmb_collection.SelectedItem
        if item and hasattr(item, 'Tag'):
            return str(item.Tag)
        return "gebouwen_terreinen"

    # =========================================================================
    # CityJSON workflow
    # =========================================================================

    def _load_cityjson(self):
        """Hoofdworkflow: download/laad + parse + maak DirectShapes."""
        target_lod = self._get_lod()
        lod_label = self._get_lod_label()
        crop_bbox = self._get_bbox()

        # Stap 1: Verkrijg CityJSON bestand (BUITEN transactie)
        if self.rdo_source_pdok.IsChecked:
            filepath = self._download_from_pdok()
        else:
            filepath = self.txt_filepath.Text.strip()

        # Stap 2: Parse CityJSON (met bbox crop)
        if crop_bbox:
            bbox_size = self._get_bbox_size()
            show_progress(self, "CityJSON parsen (crop {0}x{0}m)...".format(
                bbox_size))
        else:
            show_progress(self, "CityJSON parsen (hele tile)...")
        update_ui()

        parser = CityJSONParser()
        objects = parser.parse_file(
            filepath, target_lod=target_lod, bbox=crop_bbox)

        if not objects:
            if crop_bbox:
                raise CityJSONError(
                    "Geen objecten gevonden binnen zoekgebied "
                    "({0}x{0}m). Probeer een groter gebied.".format(
                        self._get_bbox_size()))
            raise CityJSONError("Geen objecten gevonden in CityJSON bestand")

        # Log statistieken
        counts = parser.count_by_type(objects)
        count_str = ", ".join(
            "{0}: {1}".format(t, c) for t, c in sorted(counts.items()))
        if crop_bbox:
            log("{0} objecten na crop {1}x{1}m ({2})".format(
                len(objects), self._get_bbox_size(), count_str))
        else:
            log("{0} objecten gevonden ({1})".format(
                len(objects), count_str))

        # Stap 3: Maak DirectShapes (BINNEN transactie)
        show_progress(self, "DirectShapes aanmaken...")
        update_ui()

        total_shapes = self._create_directshapes(objects, lod_label)

        self.result_count = total_shapes
        log("Totaal: {0} DirectShape(s) aangemaakt".format(total_shapes))

    def _download_from_pdok(self):
        """Download CityJSON van PDOK.

        Returns:
            Pad naar gedownload CityJSON bestand
        """
        rd_x, rd_y = self.location_rd
        collection = self._get_collection()

        show_progress(self, "PDOK tile zoeken...")
        update_ui()

        client = PDOK3DClient(timeout=120)

        def progress_cb(msg):
            show_progress(self, msg)

        tile_url = client.get_tile_url(rd_x, rd_y, collection=collection)
        log("Tile URL: {0}".format(tile_url))

        filepath = client.download_cityjson(
            tile_url, progress_callback=progress_cb)
        log("CityJSON bestand: {0}".format(filepath))

        return filepath

    # =========================================================================
    # DirectShape aanmaak
    # =========================================================================

    def _create_directshapes(self, objects, lod_label):
        """Maak DirectShape elementen voor alle CityJSON objecten.

        Args:
            objects: Lijst van CityJSONObject
            lod_label: LoD label string

        Returns:
            Aantal aangemaakte DirectShapes
        """
        from Autodesk.Revit.DB import (
            TessellatedShapeBuilder, TessellatedShapeBuilderTarget,
            TessellatedShapeBuilderFallback, TessellatedFace,
            DirectShape, ElementId, BuiltInCategory
        )
        from System.Collections.Generic import List as GenericList

        origin_x, origin_y = self.location_rd or (0, 0)
        total_shapes = 0

        with revit.Transaction("GIS2BIM - CityJSON 3D"):
            # Parameters voor gebouwen aanmaken
            has_buildings = any(obj.is_building() for obj in objects)
            if has_buildings:
                ensure_cityjson_parameters(self.doc)

            # Materiaal cache per categorie
            material_cache = {}

            for i, obj in enumerate(objects):
                if (i + 1) % 50 == 0 or i == 0:
                    show_progress(self, "DirectShape {0}/{1}...".format(
                        i + 1, len(objects)))
                    update_ui()

                ds = self._build_directshape(
                    obj, origin_x, origin_y, lod_label,
                    material_cache)

                if ds:
                    total_shapes += 1

                    # Gebouw parameters instellen
                    if obj.is_building():
                        self._set_building_params(ds, obj)

        return total_shapes

    def _build_directshape(self, obj, origin_x, origin_y, lod_label,
                           material_cache):
        """Bouw een DirectShape element van een CityJSONObject.

        Args:
            obj: CityJSONObject
            origin_x: RD X origin
            origin_y: RD Y origin
            lod_label: LoD label
            material_cache: Dict voor materiaal caching

        Returns:
            DirectShape element of None
        """
        from Autodesk.Revit.DB import (
            TessellatedShapeBuilder, TessellatedShapeBuilderTarget,
            TessellatedShapeBuilderFallback, TessellatedFace,
            DirectShape, ElementId, BuiltInCategory
        )
        from System.Collections.Generic import List as GenericList

        vertices = obj.vertices
        faces = obj.faces

        if not vertices or not faces:
            return None

        # Pre-convert vertices naar Revit XYZ (alleen gebruikte indices)
        xyz_cache = {}
        for face in faces:
            for idx in face:
                if idx not in xyz_cache and 0 <= idx < len(vertices):
                    vx, vy, vz = vertices[idx]
                    xyz_cache[idx] = rd_to_revit_xyz(
                        vx, vy, origin_x, origin_y, vz)

        # Bouw tessellated shape
        builder = TessellatedShapeBuilder()
        builder.Target = TessellatedShapeBuilderTarget.Mesh
        builder.Fallback = TessellatedShapeBuilderFallback.Salvage
        builder.OpenConnectedFaceSet(False)

        face_count = 0
        skipped = 0

        for face_indices in faces:
            try:
                if len(face_indices) == 3:
                    # Directe triangle
                    pts = []
                    valid = True
                    for idx in face_indices:
                        if idx in xyz_cache:
                            pts.append(xyz_cache[idx])
                        else:
                            valid = False
                            break

                    if valid and len(pts) == 3:
                        if not self._is_degenerate_triangle(
                                pts[0], pts[1], pts[2]):
                            face_points = GenericList[DB.XYZ]()
                            for p in pts:
                                face_points.Add(p)
                            tf = TessellatedFace(
                                face_points,
                                ElementId.InvalidElementId)
                            builder.AddFace(tf)
                            face_count += 1
                        else:
                            skipped += 1
                    else:
                        skipped += 1

                elif len(face_indices) > 3:
                    # Fan-triangulatie (al gedaan in parser, maar safety check)
                    p0_idx = face_indices[0]
                    if p0_idx not in xyz_cache:
                        skipped += 1
                        continue

                    for j in range(1, len(face_indices) - 1):
                        p1_idx = face_indices[j]
                        p2_idx = face_indices[j + 1]

                        if p1_idx in xyz_cache and p2_idx in xyz_cache:
                            p0 = xyz_cache[p0_idx]
                            p1 = xyz_cache[p1_idx]
                            p2 = xyz_cache[p2_idx]

                            if not self._is_degenerate_triangle(p0, p1, p2):
                                face_points = GenericList[DB.XYZ]()
                                face_points.Add(p0)
                                face_points.Add(p1)
                                face_points.Add(p2)
                                tf = TessellatedFace(
                                    face_points,
                                    ElementId.InvalidElementId)
                                builder.AddFace(tf)
                                face_count += 1
                            else:
                                skipped += 1
                        else:
                            skipped += 1
                else:
                    skipped += 1

            except Exception:
                skipped += 1

        builder.CloseConnectedFaceSet()

        if face_count == 0:
            return None

        builder.Build()
        result = builder.GetBuildResult()

        # Maak DirectShape
        ds = DirectShape.CreateElement(
            self.doc,
            ElementId(BuiltInCategory.OST_GenericModel)
        )
        ds.SetShape(result.GetGeometricalObjects())

        # Naam
        if obj.is_building():
            ds.Name = "CityJSON {0} {1}".format(obj.obj_id, lod_label)
        else:
            ds.Name = "CityJSON {0} {1}".format(obj.obj_type, lod_label)

        # Materiaal toewijzen
        cat = MATERIAL_CATEGORIE.get(obj.obj_type, "Overig")
        if cat not in material_cache:
            mat_name, rgb = MATERIAL_MAP.get(
                cat, ("CityJSON - Overig", (200, 200, 200)))
            material_cache[cat] = get_or_create_material(
                self.doc, mat_name, rgb)

        mat_id = material_cache[cat]
        set_element_material(self.doc, ds, mat_id)

        if skipped > 0 and skipped > face_count:
            log("{0}: {1} faces, {2} overgeslagen".format(
                obj.obj_id, face_count, skipped))

        return ds

    def _set_building_params(self, ds, obj):
        """Stel gebouw parameters in op DirectShape.

        Args:
            ds: DirectShape element
            obj: CityJSONObject met building attributen
        """
        attrs = obj.attributes

        # gebouw_id: identificatie uit CityJSON
        obj_id = obj.obj_id
        # Soms bevat de ID een prefix zoals "NL.IMBAG.Pand."
        if obj_id:
            set_element_parameter(ds, "gebouw_id", str(obj_id))

        # bouwjaar: oorspronkelijk bouwjaar
        bouwjaar = attrs.get("oorspronkelijkBouwjaar",
                   attrs.get("bouwjaar",
                   attrs.get("yearOfConstruction", "")))
        if bouwjaar:
            try:
                set_element_parameter(ds, "bouwjaar", float(bouwjaar))
            except (ValueError, TypeError):
                pass

        # gebouw_status: pand status
        status = attrs.get("status",
                 attrs.get("pandstatus", ""))
        if status:
            set_element_parameter(ds, "gebouw_status", str(status))

    def _is_degenerate_triangle(self, p0, p1, p2):
        """Check of een driehoek gedegenereerd is (oppervlakte ~0).

        Args:
            p0, p1, p2: XYZ punten

        Returns:
            True als de driehoek gedegenereerd is
        """
        ax = p1.X - p0.X
        ay = p1.Y - p0.Y
        az = p1.Z - p0.Z
        bx = p2.X - p0.X
        by = p2.Y - p0.Y
        bz = p2.Z - p0.Z

        cx = ay * bz - az * by
        cy = az * bx - ax * bz
        cz = ax * by - ay * bx

        length_sq = cx * cx + cy * cy + cz * cz
        return length_sq < 0.0001


# =============================================================================
# Main
# =============================================================================

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

    window = CityJSONWindow(doc)
    result = window.ShowDialog()

    if result and window.result_count > 0:
        forms.alert(
            "CityJSON: {0} DirectShape(s) aangemaakt.".format(
                window.result_count),
            title="GIS2BIM CityJSON 3D",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
