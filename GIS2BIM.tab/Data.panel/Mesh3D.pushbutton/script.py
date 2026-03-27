# -*- coding: utf-8 -*-
"""
3D Mesh Importeren - GIS2BIM
==============================

Importeer 3D mesh bestanden als DirectShape elementen in Revit.
Twee bronnen:
- OBJ bestand: vanuit Kavel10, Blender, Google Earth export, etc.
- Google 3D Tiles: fotogrammetrische meshes via Google Map Tiles API

Workflow:
1. Kies bron (bestand of Google 3D)
2. Kies opties (coordinaten, detail, etc.)
3. Parse mesh data
4. Maak DirectShape (GenericModel) elementen
"""

__title__ = "3D Mesh\nImport"
__author__ = "OpenAEC Foundation"
__doc__ = "Importeer OBJ / Google 3D mesh als DirectShape"

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
    METER_TO_FEET,
)

log = get_logger("Mesh3D")

# GIS2BIM modules
GIS2BIM_LOADED = False
IMPORT_ERROR = ""

try:
    from gis2bim.parsers.obj import OBJReader, OBJError
    from gis2bim.parsers.mtl import MTLReader, MTLError
    from gis2bim.parsers.glb import GLBReader, GLBError
    from gis2bim.api.google3d import (
        Google3DClient, Google3DError,
        ecef_to_wgs84,
    )
    from gis2bim.coordinates import rd_to_wgs84, wgs84_to_rd
    from gis2bim.config import get_api_key, set_api_key
    GIS2BIM_LOADED = True
    log("Modules geladen")
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())


# Standaard materiaalkleur (lichtgrijs)
DEFAULT_COLOR = (180, 180, 180)
GOOGLE3D_COLOR = (200, 195, 185)

# Config key voor Google API
GOOGLE_API_KEY_NAME = "google_maps_api_key"


class Mesh3DWindow(Window):
    """WPF Window voor 3D mesh import."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.location_rd = None
        self.result_count = 0
        self._preview_info = None

        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        root = load_xaml_window(self, xaml_path)

        bind_ui_elements(self, root, [
            'txt_subtitle', 'pnl_location', 'txt_location_x', 'txt_location_y',
            'pnl_no_location',
            'rdo_source_file', 'rdo_source_google',
            'pnl_file_options', 'pnl_google_options',
            'txt_filepath', 'btn_browse', 'pnl_file_info', 'txt_file_info',
            'rdo_coords_rd', 'rdo_coords_local',
            'cmb_unit',
            'chk_per_object', 'chk_use_materials',
            'txt_api_key', 'cmb_google_radius',
            'rdo_detail_low', 'rdo_detail_medium', 'rdo_detail_high',
            'txt_info',
            'pnl_progress', 'txt_progress', 'progress_bar',
            'txt_status', 'btn_cancel', 'btn_execute'
        ])

        self.location_rd = setup_project_location(self, doc, log)
        self._load_saved_api_key()
        self._bind_events()

    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
        self.btn_browse.Click += self._on_browse
        self.rdo_source_file.Checked += self._on_source_changed
        self.rdo_source_google.Checked += self._on_source_changed

    def _load_saved_api_key(self):
        """Laad opgeslagen Google API key."""
        try:
            key = get_api_key(GOOGLE_API_KEY_NAME)
            if key:
                self.txt_api_key.Text = key
        except Exception:
            pass

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_source_changed(self, sender, args):
        """Wissel tussen OBJ bestand en Google 3D opties."""
        if self.rdo_source_file.IsChecked:
            self.pnl_file_options.Visibility = Visibility.Visible
            self.pnl_google_options.Visibility = Visibility.Collapsed
        else:
            self.pnl_file_options.Visibility = Visibility.Collapsed
            self.pnl_google_options.Visibility = Visibility.Visible

    def _on_browse(self, sender, args):
        """Open bestandskiezer voor OBJ bestand."""
        from Microsoft.Win32 import OpenFileDialog

        dlg = OpenFileDialog()
        dlg.Title = "Selecteer mesh bestand"
        dlg.Filter = (
            "Mesh bestanden (*.obj;*.glb)|*.obj;*.glb|"
            "Wavefront OBJ (*.obj)|*.obj|"
            "Binary glTF (*.glb)|*.glb|"
            "Alle bestanden (*.*)|*.*"
        )

        if dlg.ShowDialog():
            self.txt_filepath.Text = dlg.FileName
            self._preview_file(dlg.FileName)

    def _preview_file(self, filepath):
        """Toon preview info van het geselecteerde bestand."""
        try:
            file_size = os.path.getsize(filepath)
            if file_size > 1024 * 1024:
                size_str = "{0:.1f} MB".format(file_size / (1024.0 * 1024.0))
            else:
                size_str = "{0:.0f} KB".format(file_size / 1024.0)

            ext = os.path.splitext(filepath)[1].lower()

            if ext == ".obj":
                vertex_count = 0
                face_count = 0
                has_mtl = False

                with open(filepath, "r") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("v "):
                            vertex_count += 1
                        elif line.startswith("f "):
                            face_count += 1
                        elif line.startswith("mtllib "):
                            has_mtl = True

                info_parts = [
                    "{0} vertices, {1} faces".format(
                        self._format_number(vertex_count),
                        self._format_number(face_count)),
                    "Grootte: {0}".format(size_str),
                ]
                if has_mtl:
                    info_parts.append("MTL materialen: ja")

            elif ext == ".glb":
                info_parts = [
                    "GLB (binary glTF 2.0)",
                    "Grootte: {0}".format(size_str),
                    "Coordinaten: waarschijnlijk ECEF (wordt automatisch geconverteerd)",
                ]
            else:
                info_parts = ["Grootte: {0}".format(size_str)]

            self.txt_file_info.Text = "\n".join(info_parts)
            self.pnl_file_info.Visibility = Visibility.Visible

        except Exception as e:
            self.pnl_file_info.Visibility = Visibility.Collapsed
            log("Preview fout: {0}".format(e))

    def _format_number(self, n):
        """Formatteer getal met duizendtallen separator."""
        s = str(n)
        parts = []
        while s:
            parts.append(s[-3:])
            s = s[:-3]
        return ".".join(reversed(parts))

    def _on_execute(self, sender, args):
        """Start de mesh import."""
        is_google = self.rdo_source_google.IsChecked

        if is_google:
            # Google 3D validatie
            if not self.location_rd:
                self.txt_status.Text = (
                    "Geen locatie beschikbaar. Stel eerst een locatie in.")
                return
            api_key = self.txt_api_key.Text.strip()
            if not api_key:
                self.txt_status.Text = "Vul een Google API key in"
                return
        else:
            # OBJ validatie
            filepath = self.txt_filepath.Text.strip()
            if not filepath or not os.path.exists(filepath):
                self.txt_status.Text = "Selecteer een geldig bestand"
                return

            is_rd = self.rdo_coords_rd.IsChecked
            if is_rd and not self.location_rd:
                self.txt_status.Text = (
                    "Geen locatie beschikbaar voor RD-modus. "
                    "Gebruik 'Lokaal' of stel eerst een locatie in.")
                return

        show_progress(self, "Mesh importeren...")
        self.btn_execute.IsEnabled = False

        try:
            if is_google:
                self._import_google3d()
            else:
                self._import_file()
            self.DialogResult = True
            self.Close()
        except (OBJError, MTLError, GLBError, Google3DError) as e:
            log("Mesh error: {0}".format(e))
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True
        except Exception as e:
            log("Error importing mesh: {0}".format(e))
            log(traceback.format_exc())
            hide_progress(self)
            self.txt_status.Text = "Fout: {0}".format(str(e))
            self.btn_execute.IsEnabled = True

    # =========================================================================
    # UI getters
    # =========================================================================

    def _get_unit_scale(self):
        """Haal eenheid schaalfactor op (naar meters)."""
        item = self.cmb_unit.SelectedItem
        if item and hasattr(item, 'Tag'):
            return float(item.Tag)
        return 1.0

    def _is_rd_mode(self):
        return self.rdo_coords_rd.IsChecked

    def _is_per_object(self):
        return self.chk_per_object.IsChecked

    def _use_materials(self):
        return self.chk_use_materials.IsChecked

    def _get_google_radius(self):
        item = self.cmb_google_radius.SelectedItem
        if item and hasattr(item, 'Tag'):
            return int(item.Tag)
        return 100

    def _get_google_detail(self):
        """Haal geometric error threshold op."""
        if self.rdo_detail_low.IsChecked:
            return 50.0
        if self.rdo_detail_high.IsChecked:
            return 5.0
        return 20.0  # medium

    # =========================================================================
    # OBJ / GLB file import workflow
    # =========================================================================

    def _import_file(self):
        """Import OBJ of GLB bestand."""
        filepath = self.txt_filepath.Text.strip()
        ext = os.path.splitext(filepath)[1].lower()

        show_progress(self, "Bestand parsen...")
        update_ui()

        if ext == ".glb":
            self._import_glb_file(filepath)
        else:
            self._import_obj_file(filepath)

    def _import_obj_file(self, filepath):
        """Import OBJ bestand als DirectShape(s)."""
        per_object = self._is_per_object()
        reader = OBJReader()

        if per_object:
            meshes = reader.read(filepath)
        else:
            meshes = [reader.read_as_single_mesh(filepath)]

        log("{0} mesh(es) geparsed".format(len(meshes)))

        # Laad MTL materialen
        mtl_materials = {}
        if self._use_materials():
            mtl_materials = self._load_mtl_materials(filepath, meshes)

        # Maak DirectShapes
        mesh_name = os.path.splitext(os.path.basename(filepath))[0]
        total = self._create_directshapes_obj(meshes, mtl_materials, mesh_name)
        self.result_count = total

    def _import_glb_file(self, filepath):
        """Import lokaal GLB bestand als DirectShape(s)."""
        reader = GLBReader()
        meshes = reader.read(filepath)
        log("{0} mesh(es) uit GLB geparsed".format(len(meshes)))

        # Detecteer of vertices ECEF zijn (grote absolute waarden)
        is_ecef = self._detect_ecef(meshes)
        mesh_name = os.path.splitext(os.path.basename(filepath))[0]

        if is_ecef:
            log("ECEF coordinaten gedetecteerd, conversie naar RD")
            total = self._create_directshapes_ecef(meshes, mesh_name)
        else:
            total = self._create_directshapes_obj(meshes, {}, mesh_name)

        self.result_count = total

    def _detect_ecef(self, meshes):
        """Detecteer of mesh vertices in ECEF coordinaten zijn.

        ECEF waarden zijn typisch 6.3M+ (aardstraal), terwijl
        RD waarden 0-300K zijn en lokale waarden klein.
        """
        if not meshes or not meshes[0].get("vertices"):
            return False
        vx, vy, vz = meshes[0]["vertices"][0]
        magnitude = (vx ** 2 + vy ** 2 + vz ** 2) ** 0.5
        return magnitude > 1000000  # > 1000km = ECEF

    # =========================================================================
    # Google 3D Tiles import workflow
    # =========================================================================

    def _import_google3d(self):
        """Download en importeer Google 3D Tiles."""
        api_key = self.txt_api_key.Text.strip()

        # Sla API key op voor volgende keer
        try:
            set_api_key(api_key, GOOGLE_API_KEY_NAME)
        except Exception:
            pass

        rd_x, rd_y = self.location_rd
        lat, lon = rd_to_wgs84(rd_x, rd_y)
        radius = self._get_google_radius()
        max_error = self._get_google_detail()

        log("Google 3D: lat={0:.6f}, lon={1:.6f}, radius={2}m, "
            "max_error={3}".format(lat, lon, radius, max_error))

        def progress_cb(msg):
            show_progress(self, msg)
            update_ui()

        # Stap 1: Download tiles
        client = Google3DClient(api_key, timeout=60)

        glb_paths = client.get_tiles_for_location(
            lat, lon,
            radius_m=radius,
            max_geometric_error=max_error,
            max_tiles=80,
            progress_callback=progress_cb,
        )

        log("{0} GLB tiles gedownload".format(len(glb_paths)))

        # Stap 2: Parse alle GLB tiles
        show_progress(self, "GLB tiles parsen...")
        update_ui()

        reader = GLBReader()
        all_meshes = []

        for i, glb_path in enumerate(glb_paths):
            try:
                meshes = reader.read(glb_path)
                all_meshes.extend(meshes)
            except GLBError as e:
                log("GLB parse fout tile {0}: {1}".format(i, e))
            finally:
                # Verwijder temp bestand
                try:
                    os.remove(glb_path)
                except Exception:
                    pass

        if not all_meshes:
            raise Google3DError("Geen geldige mesh data in gedownloade tiles")

        log("{0} meshes totaal uit {1} tiles".format(
            len(all_meshes), len(glb_paths)))

        # Stap 3: Maak DirectShapes (ECEF → RD conversie)
        total = self._create_directshapes_ecef(all_meshes, "Google3D")
        self.result_count = total
        log("Totaal: {0} DirectShape(s) aangemaakt".format(total))

    # =========================================================================
    # DirectShape aanmaak - OBJ modus (RD of lokaal)
    # =========================================================================

    def _create_directshapes_obj(self, meshes, mtl_materials, mesh_name):
        """Maak DirectShapes van OBJ meshes.

        Args:
            meshes: Lijst van mesh dicts
            mtl_materials: Dict van MTL materialen
            mesh_name: Basis naam

        Returns:
            Aantal aangemaakte DirectShapes
        """
        total_shapes = 0

        with revit.Transaction("GIS2BIM - 3D Mesh Import"):
            revit_material_cache = {}

            for i, mesh in enumerate(meshes):
                if len(meshes) > 1 and ((i + 1) % 25 == 0 or i == 0):
                    show_progress(self, "DirectShape {0}/{1}...".format(
                        i + 1, len(meshes)))
                    update_ui()

                xyz_cache = self._convert_vertices_obj(mesh["vertices"])
                ds = self._build_directshape(mesh["faces"], xyz_cache)

                if ds:
                    total_shapes += 1
                    per_object = self._is_per_object()
                    if per_object:
                        obj_name = mesh.get("name", "unnamed")
                        ds.Name = "Mesh {0} - {1}".format(mesh_name, obj_name)
                    else:
                        ds.Name = "Mesh {0}".format(mesh_name)

                    mat_id = self._resolve_material(
                        mesh, mtl_materials, revit_material_cache, mesh_name)
                    if mat_id:
                        set_element_material(self.doc, ds, mat_id)

        return total_shapes

    def _convert_vertices_obj(self, vertices):
        """Converteer OBJ vertices naar Revit XYZ (RD of lokaal modus)."""
        from Autodesk.Revit.DB import XYZ

        unit_scale = self._get_unit_scale()
        is_rd = self._is_rd_mode()

        if is_rd and self.location_rd:
            origin_x, origin_y = self.location_rd
            xyz_list = []
            for vx, vy, vz in vertices:
                xyz = rd_to_revit_xyz(
                    vx * unit_scale, vy * unit_scale,
                    origin_x, origin_y, vz * unit_scale)
                xyz_list.append(xyz)
            return xyz_list
        else:
            scale = unit_scale * METER_TO_FEET
            return [XYZ(vx * scale, vy * scale, vz * scale)
                    for vx, vy, vz in vertices]

    # =========================================================================
    # DirectShape aanmaak - ECEF modus (Google 3D / GLB)
    # =========================================================================

    def _create_directshapes_ecef(self, meshes, mesh_name):
        """Maak DirectShapes van ECEF meshes (Google 3D Tiles).

        Converteert ECEF → WGS84 → RD → Revit coordinaten.

        Args:
            meshes: Lijst van mesh dicts met ECEF vertices
            mesh_name: Basis naam

        Returns:
            Aantal aangemaakte DirectShapes
        """
        if not self.location_rd:
            raise Google3DError("Geen projectlocatie voor coordinaat-conversie")

        origin_x, origin_y = self.location_rd
        total_shapes = 0

        with revit.Transaction("GIS2BIM - 3D Mesh Import"):
            mat_id = get_or_create_material(
                self.doc, "Mesh - {0}".format(mesh_name), GOOGLE3D_COLOR)

            for i, mesh in enumerate(meshes):
                if (i + 1) % 10 == 0 or i == 0:
                    show_progress(self, "DirectShape {0}/{1}...".format(
                        i + 1, len(meshes)))
                    update_ui()

                xyz_cache = self._convert_vertices_ecef(
                    mesh["vertices"], origin_x, origin_y)
                ds = self._build_directshape(mesh["faces"], xyz_cache)

                if ds:
                    total_shapes += 1
                    ds.Name = "{0} tile {1}".format(mesh_name, i + 1)
                    set_element_material(self.doc, ds, mat_id)

        return total_shapes

    def _convert_vertices_ecef(self, vertices, origin_x, origin_y):
        """Converteer ECEF vertices naar Revit XYZ via WGS84 → RD.

        Args:
            vertices: Lijst van (x, y, z) ECEF tuples
            origin_x, origin_y: RD origin van project

        Returns:
            Lijst van Revit XYZ objecten
        """
        xyz_list = []
        for ex, ey, ez in vertices:
            lat, lon, h = ecef_to_wgs84(ex, ey, ez)
            rd_x, rd_y = wgs84_to_rd(lat, lon)
            xyz = rd_to_revit_xyz(rd_x, rd_y, origin_x, origin_y, h)
            xyz_list.append(xyz)
        return xyz_list

    # =========================================================================
    # Gedeelde DirectShape builder
    # =========================================================================

    def _build_directshape(self, faces, xyz_cache):
        """Bouw DirectShape element van faces + pre-converted vertices.

        Args:
            faces: Lijst van face index tuples
            xyz_cache: Lijst van Revit XYZ objecten (per vertex index)

        Returns:
            DirectShape element, of None
        """
        from Autodesk.Revit.DB import (
            TessellatedShapeBuilder, TessellatedShapeBuilderTarget,
            TessellatedShapeBuilderFallback, TessellatedFace,
            DirectShape, ElementId, BuiltInCategory,
        )
        from System.Collections.Generic import List as GenericList

        if not faces or not xyz_cache:
            return None

        builder = TessellatedShapeBuilder()
        builder.Target = TessellatedShapeBuilderTarget.Mesh
        builder.Fallback = TessellatedShapeBuilderFallback.Salvage
        builder.OpenConnectedFaceSet(False)

        face_count = 0
        skipped = 0
        cache_len = len(xyz_cache)

        for face_indices in faces:
            try:
                if len(face_indices) == 3:
                    face_points = GenericList[DB.XYZ]()
                    valid = True
                    for idx in face_indices:
                        if 0 <= idx < cache_len:
                            face_points.Add(xyz_cache[idx])
                        else:
                            valid = False
                            break

                    if valid and face_points.Count == 3:
                        if not self._is_degenerate_triangle(
                                face_points[0], face_points[1],
                                face_points[2]):
                            tf = TessellatedFace(
                                face_points, ElementId.InvalidElementId)
                            builder.AddFace(tf)
                            face_count += 1
                        else:
                            skipped += 1
                    else:
                        skipped += 1

                elif len(face_indices) > 3:
                    p0_idx = face_indices[0]
                    if p0_idx < 0 or p0_idx >= cache_len:
                        skipped += 1
                        continue

                    for j in range(1, len(face_indices) - 1):
                        p1_idx = face_indices[j]
                        p2_idx = face_indices[j + 1]

                        if (0 <= p1_idx < cache_len and
                                0 <= p2_idx < cache_len):
                            face_points = GenericList[DB.XYZ]()
                            face_points.Add(xyz_cache[p0_idx])
                            face_points.Add(xyz_cache[p1_idx])
                            face_points.Add(xyz_cache[p2_idx])

                            if not self._is_degenerate_triangle(
                                    face_points[0], face_points[1],
                                    face_points[2]):
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

        if skipped > 0 and skipped > face_count * 0.1:
            log("{0} faces, {1} overgeslagen".format(face_count, skipped))

        ds = DirectShape.CreateElement(
            self.doc,
            ElementId(BuiltInCategory.OST_GenericModel)
        )
        ds.SetShape(result.GetGeometricalObjects())
        return ds

    def _is_degenerate_triangle(self, p0, p1, p2):
        """Check of een driehoek gedegenereerd is (oppervlakte ~0)."""
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

    # =========================================================================
    # MTL materiaal support
    # =========================================================================

    def _load_mtl_materials(self, obj_filepath, meshes):
        """Laad MTL materialen als die bestaan."""
        mtllib = None
        for mesh in meshes:
            mtllib = mesh.get("mtllib")
            if mtllib:
                break

        if not mtllib:
            return {}

        obj_dir = os.path.dirname(os.path.abspath(obj_filepath))
        mtl_path = os.path.join(obj_dir, mtllib)

        if not os.path.exists(mtl_path):
            log("MTL niet gevonden: {0}".format(mtl_path))
            return {}

        try:
            mtl_reader = MTLReader()
            materials = mtl_reader.read(mtl_path)
            log("{0} materialen uit MTL".format(len(materials)))
            return materials
        except MTLError as e:
            log("MTL fout: {0}".format(e))
            return {}

    def _resolve_material(self, mesh, mtl_materials, cache, mesh_name):
        """Bepaal en maak Revit materiaal voor een mesh."""
        mat_name = mesh.get("material")

        if mat_name and mat_name in mtl_materials:
            if mat_name in cache:
                return cache[mat_name]

            mtl_mat = mtl_materials[mat_name]
            mtl_reader = MTLReader()
            rgb = mtl_reader.get_rgb_255(mtl_mat)
            revit_name = "Mesh - {0}".format(mat_name)
            mat_id = get_or_create_material(self.doc, revit_name, rgb)
            cache[mat_name] = mat_id
            return mat_id

        fallback_key = "__default__"
        if fallback_key in cache:
            return cache[fallback_key]

        revit_name = "Mesh - {0}".format(mesh_name)
        mat_id = get_or_create_material(self.doc, revit_name, DEFAULT_COLOR)
        cache[fallback_key] = mat_id
        return mat_id


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

    window = Mesh3DWindow(doc)
    result = window.ShowDialog()

    if result and window.result_count > 0:
        forms.alert(
            "3D Mesh: {0} DirectShape(s) aangemaakt.".format(
                window.result_count),
            title="GIS2BIM 3D Mesh",
            warn_icon=False
        )


if __name__ == "__main__":
    main()
