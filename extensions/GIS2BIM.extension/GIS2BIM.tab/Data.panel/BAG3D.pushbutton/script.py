# -*- coding: utf-8 -*-
"""
BAG3D Data Laden - GIS2BIM
==========================

Laad 3D gebouwmodellen van 3DBAG en maak DirectShape elementen in Revit.

Workflow:
1. WFS query voor tiles in zoekgebied
2. Download OBJ ZIP bestanden
3. Parse OBJ meshes
4. Maak DirectShape (GenericModel) per tile
"""

__title__ = "BAG3D\nLaden"
__author__ = "OpenAEC Foundation"
__doc__ = "Laad 3D gebouwmodellen (3DBAG) als DirectShape"

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
from gis2bim.revit.geometry import rd_to_revit_xyz, set_element_parameter

log = get_logger("BAG3D")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.api.bag3d import BAG3DClient, BAG3DError
    from gis2bim.parsers.obj import OBJReader, OBJError
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


_bag3d_params_created = False

# Maximale bbox grootte (m) bij per-gebouw modus
PER_BUILDING_MAX_BBOX = 200


def ensure_bag3d_parameters(doc):
    """Zorg dat BAG3D parameters bestaan op GenericModel categorie.

    Maakt instance parameters aan:
    - BAG_ID (text): BAG pand identificatie
    - BAG_bouwjaar (number): Oorspronkelijk bouwjaar
    - BAG_leeftijd (number): Leeftijd in jaren (huidig jaar - bouwjaar)

    Args:
        doc: Revit document

    Returns:
        True als succesvol
    """
    global _bag3d_params_created
    if _bag3d_params_created:
        return True

    from Autodesk.Revit.DB import (
        ExternalDefinitionCreationOptions, BuiltInCategory
    )

    app = doc.Application

    param_defs = [
        ("BAG_ID", "text"),
        ("BAG_bouwjaar", "number"),
        ("BAG_leeftijd", "number"),
    ]

    # Check welke al bestaan
    existing = set()
    bm = doc.ParameterBindings
    it = bm.ForwardIterator()
    while it.MoveNext():
        existing.add(it.Key.Name)

    needed = [(n, t) for n, t in param_defs if n not in existing]
    if not needed:
        _bag3d_params_created = True
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
        _bag3d_params_created = True
        return True

    except Exception as e:
        log("ensure_bag3d_parameters fout: {0}".format(e))
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


def extract_bag_id(obj_name):
    """Haal de numerieke BAG pand-ID uit een OBJ objectnaam.

    3DBAG objectnamen zijn bv. "NL.IMBAG.Pand.0363100012345678".
    BAG WFS retourneert identificatie als "0363100012345678".

    Args:
        obj_name: OBJ object naam string

    Returns:
        Numerieke pand-ID string, of de originele naam als fallback
    """
    if not obj_name:
        return ""
    # Strip prefix: alles na de laatste punt als het op cijfers eindigt
    parts = obj_name.split(".")
    if len(parts) > 1 and parts[-1].isdigit():
        return parts[-1]
    return obj_name


class BAG3DWindow(Window):
    """WPF Window voor BAG3D data laden."""

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
            'cmb_bbox_size',
            'rdo_lod22', 'rdo_lod13', 'rdo_lod12',
            'chk_per_building',
            'txt_info',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        self._bind_events()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.chk_per_building.Checked += self._on_per_building_changed
        self.chk_per_building.Unchecked += self._on_per_building_changed

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_execute(self, sender, args):
        if not self.location_rd:
            self.txt_status.Text = "Geen locatie beschikbaar"
            return

        show_progress(self, "3DBAG data laden...")
        self.btn_execute.IsEnabled = False

        try:
            self._load_bag3d()
            self.DialogResult = True
            self.Close()
        except (BAG3DError, OBJError) as e:
            log("BAG3D error: {0}".format(e))
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True
        except Exception as e:
            log("Error loading BAG3D: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    def _on_per_building_changed(self, sender, args):
        """Update UI wanneer per-gebouw checkbox verandert."""
        checked = self.chk_per_building.IsChecked
        max_size = PER_BUILDING_MAX_BBOX

        for item in self.cmb_bbox_size.Items:
            tag = int(item.Tag)
            item.IsEnabled = not checked or tag <= max_size

        # Als huidige selectie te groot is, selecteer max toegestane
        if checked:
            current = self._get_bbox_size()
            if current > max_size:
                for item in self.cmb_bbox_size.Items:
                    if int(item.Tag) == max_size:
                        self.cmb_bbox_size.SelectedItem = item
                        break

    # =========================================================================
    # UI getters
    # =========================================================================

    def _get_bbox_size(self):
        item = self.cmb_bbox_size.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 200

    def _get_lod(self):
        if hasattr(self, 'rdo_lod13') and self.rdo_lod13.IsChecked:
            return "lod13"
        if hasattr(self, 'rdo_lod12') and self.rdo_lod12.IsChecked:
            return "lod12"
        return "lod22"

    def _get_lod_label(self):
        lod = self._get_lod()
        labels = {
            "lod12": "LoD 1.2",
            "lod13": "LoD 1.3",
            "lod22": "LoD 2.2",
        }
        return labels.get(lod, lod)

    def _is_per_building(self):
        return (hasattr(self, 'chk_per_building')
                and self.chk_per_building.IsChecked)

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
    # BAG3D workflow
    # =========================================================================

    def _load_bag3d(self):
        bbox = self._get_bbox()
        lod = self._get_lod()
        lod_label = self._get_lod_label()
        per_building = self._is_per_building()

        log("BAG3D: BBOX={0}, LoD={1}, per_building={2}".format(
            bbox, lod, per_building))

        show_progress(self, "Zoeken naar 3DBAG tiles...")

        client = BAG3DClient(timeout=60)
        tiles = client.get_tiles(bbox)
        log("{0} tiles gevonden".format(len(tiles)))

        # ── Fase 1: Downloads (parallel) ──
        show_progress(self, "Downloaden {0} tiles...".format(len(tiles)))
        update_ui()
        obj_paths = client.download_tiles_parallel(tiles, lod=lod)
        log("{0} OBJ bestanden gedownload".format(len(obj_paths)))

        reader = OBJReader()
        origin_x, origin_y = self.location_rd
        total_shapes = 0

        if per_building:
            # ── Fase 2: Pre-processing (buiten transactie) ──
            show_progress(self, "Parameters aanmaken...")
            ensure_bag3d_parameters(self.doc)

            show_progress(self, "Bouwjaren ophalen uit BAG...")
            update_ui()
            bouwjaren = client.get_bouwjaren(bbox)
            log("{0} bouwjaren opgehaald uit BAG".format(len(bouwjaren)))

            # Parse alle OBJ + converteer vertices vooraf
            show_progress(self, "OBJ parsen en vertices converteren...")
            update_ui()
            all_buildings = []

            for i, obj_path in enumerate(obj_paths):
                meshes = reader.read(obj_path)
                for mesh in meshes:
                    name = mesh.get("name", "unnamed")
                    vertices = mesh["vertices"]
                    faces = mesh["faces"]
                    if not vertices or not faces:
                        continue
                    # Pre-convert vertices naar Revit XYZ
                    xyz_cache = []
                    for vx, vy, vz in vertices:
                        xyz_cache.append(
                            rd_to_revit_xyz(vx, vy, origin_x, origin_y, vz))
                    all_buildings.append({
                        "name": name,
                        "xyz_cache": xyz_cache,
                        "faces": faces,
                    })

            log("{0} gebouwen geparsed uit {1} tiles".format(
                len(all_buildings), len(obj_paths)))

            # ── Fase 3: Revit DirectShapes (één transactie) ──
            show_progress(self, "DirectShapes aanmaken: {0} gebouwen...".format(
                len(all_buildings)))
            update_ui()

            with revit.Transaction("GIS2BIM BAG3D {0}".format(lod_label)):
                for j, bldg in enumerate(all_buildings):
                    if j % 50 == 0:
                        show_progress(
                            self,
                            "DirectShape {0}/{1}...".format(
                                j + 1, len(all_buildings)))
                        update_ui()

                    ds = self._build_directshape_from_cache(
                        bldg["xyz_cache"], bldg["faces"],
                        name="BAG3D {0}".format(bldg["name"]))
                    if ds:
                        total_shapes += 1
                        self._set_bag_parameters(
                            ds, bldg["name"], bouwjaren)

            log("{0} DirectShapes aangemaakt".format(total_shapes))
        else:
            for i, obj_path in enumerate(obj_paths):
                show_progress(self, "Verwerken mesh {0}/{1}...".format(
                    i + 1, len(obj_paths)))

                mesh = reader.read_as_single_mesh(obj_path)
                log("Mesh {0}: {1} vertices, {2} faces".format(
                    i + 1, len(mesh["vertices"]), len(mesh["faces"])))

                show_progress(self, "DirectShape aanmaken {0}/{1} ({2} faces)...".format(
                    i + 1, len(obj_paths), len(mesh["faces"])))

                ds = self._create_directshape(mesh, origin_x, origin_y, lod_label)
                if ds:
                    total_shapes += 1

        self.result_count = total_shapes
        log("Totaal: {0} DirectShape(s) aangemaakt".format(total_shapes))

    # =========================================================================
    # DirectShape aanmaak
    # =========================================================================

    def _create_directshape(self, mesh, origin_x, origin_y, lod_label, name=None):
        """Bouw DirectShape met eigen transactie (voor per-tile modus)."""
        with revit.Transaction("GIS2BIM BAG3D {0}".format(lod_label)):
            return self._build_directshape(
                mesh, origin_x, origin_y, lod_label, name)

    def _build_directshape(self, mesh, origin_x, origin_y, lod_label, name=None):
        """Bouw DirectShape element (caller beheert transactie).

        Args:
            mesh: Dict met "vertices" en "faces"
            origin_x, origin_y: RD origin
            lod_label: LoD label string
            name: Optionele element naam

        Returns:
            DirectShape element, of None bij lege/ongeldige mesh
        """
        vertices = mesh["vertices"]
        faces = mesh["faces"]

        if not vertices or not faces:
            log("Lege mesh, overgeslagen")
            return None

        xyz_cache = []
        for vx, vy, vz in vertices:
            xyz_cache.append(rd_to_revit_xyz(vx, vy, origin_x, origin_y, vz))

        ds = self._build_directshape_from_cache(xyz_cache, faces, name=name)
        if ds is None:
            log("DirectShape overgeslagen (geen geldige faces)")
        return ds

    def _build_directshape_from_cache(self, xyz_cache, faces, name=None):
        """Bouw DirectShape uit pre-geconverteerde XYZ vertices.

        Geoptimaliseerd voor batch gebruik: geen logging per gebouw,
        imports cached, minimale overhead.

        Args:
            xyz_cache: Lijst van Revit XYZ punten
            faces: Lijst van face index-lijsten
            name: Optionele element naam

        Returns:
            DirectShape element, of None bij lege/ongeldige mesh
        """
        from Autodesk.Revit.DB import (
            TessellatedShapeBuilder, TessellatedShapeBuilderTarget,
            TessellatedShapeBuilderFallback, TessellatedFace,
            DirectShape, ElementId, BuiltInCategory
        )
        from System.Collections.Generic import List as GenericList

        if not xyz_cache or not faces:
            return None

        n_verts = len(xyz_cache)
        builder = TessellatedShapeBuilder()
        builder.Target = TessellatedShapeBuilderTarget.Mesh
        builder.Fallback = TessellatedShapeBuilderFallback.Salvage
        builder.OpenConnectedFaceSet(False)

        face_count = 0
        invalid_id = ElementId.InvalidElementId

        for face_indices in faces:
            n_idx = len(face_indices)
            try:
                if n_idx == 3:
                    i0, i1, i2 = face_indices
                    if (0 <= i0 < n_verts and 0 <= i1 < n_verts
                            and 0 <= i2 < n_verts):
                        p0 = xyz_cache[i0]
                        p1 = xyz_cache[i1]
                        p2 = xyz_cache[i2]
                        if not self._is_degenerate_triangle(p0, p1, p2):
                            fp = GenericList[DB.XYZ]()
                            fp.Add(p0)
                            fp.Add(p1)
                            fp.Add(p2)
                            builder.AddFace(TessellatedFace(fp, invalid_id))
                            face_count += 1

                elif n_idx > 3:
                    p0_idx = face_indices[0]
                    if 0 <= p0_idx < n_verts:
                        p0 = xyz_cache[p0_idx]
                        for k in range(1, n_idx - 1):
                            i1 = face_indices[k]
                            i2 = face_indices[k + 1]
                            if (0 <= i1 < n_verts and 0 <= i2 < n_verts):
                                p1 = xyz_cache[i1]
                                p2 = xyz_cache[i2]
                                if not self._is_degenerate_triangle(p0, p1, p2):
                                    fp = GenericList[DB.XYZ]()
                                    fp.Add(p0)
                                    fp.Add(p1)
                                    fp.Add(p2)
                                    builder.AddFace(
                                        TessellatedFace(fp, invalid_id))
                                    face_count += 1
            except Exception:
                pass

        builder.CloseConnectedFaceSet()

        if face_count == 0:
            return None

        builder.Build()
        result = builder.GetBuildResult()

        ds = DirectShape.CreateElement(
            self.doc,
            ElementId(BuiltInCategory.OST_GenericModel)
        )
        ds.SetShape(result.GetGeometricalObjects())
        if name:
            ds.Name = name

        return ds

    def _set_bag_parameters(self, ds, obj_name, bouwjaren):
        """Stel BAG_ID, BAG_bouwjaar en BAG_leeftijd in op een DirectShape.

        Args:
            ds: DirectShape element
            obj_name: OBJ object naam (bv. "NL.IMBAG.Pand.0363100012345678")
            bouwjaren: Dict mapping pand_id -> bouwjaar (int)
        """
        import time

        bag_id = extract_bag_id(obj_name)
        set_element_parameter(ds, "BAG_ID", str(bag_id))

        bouwjaar = bouwjaren.get(bag_id)
        if bouwjaar:
            set_element_parameter(ds, "BAG_bouwjaar", float(bouwjaar))
            huidig_jaar = int(time.strftime("%Y"))
            leeftijd = huidig_jaar - int(bouwjaar)
            set_element_parameter(ds, "BAG_leeftijd", float(leeftijd))

    def _is_degenerate_triangle(self, p0, p1, p2):
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

    window = BAG3DWindow(doc)
    result = window.ShowDialog()

    if result and window.result_count > 0:
        forms.alert(
            "BAG3D: {0} DirectShape(s) aangemaakt.".format(
                window.result_count),
            title="GIS2BIM BAG3D",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
