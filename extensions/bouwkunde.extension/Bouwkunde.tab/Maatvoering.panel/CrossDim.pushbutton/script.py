# -*- coding: utf-8 -*-
"""CrossDim - Room Maatvoering

Plaats kruisende maatlijnen in rooms met één klik.
Klik in een room om horizontale en verticale maatlijnen te plaatsen.

Auteur: 3BM Bouwkunde
Versie: 1.5.1 - Wandfilter alleen op 'tegel' in typenaam (gipsplaat e.d. weer meegenomen)
"""

import clr
import sys
import os
import math

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from bm_logger import get_logger
log = get_logger("CrossDim")
log.info("CrossDim v1.5.1")

from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl

try:
    clr.AddReference('RevitAPI')
    clr.AddReference('RevitAPIUI')
except:
    pass

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI.Selection import ObjectSnapTypes
from Autodesk.Revit.Exceptions import OperationCanceledException

from System.Windows.Forms import DialogResult, ComboBox, ComboBoxStyle
from System.Drawing import Point, Size

from pyrevit import revit, forms

# GEEN doc/uidoc/active_view = revit.doc hier! 
# Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
uidoc = None
active_view = None



# =============================================================================
# HELPERS
# =============================================================================
def get_dimension_types():
    """Haal alle lineaire dimension types op."""
    collector = FilteredElementCollector(doc).OfClass(DimensionType)
    types = []
    for dt in collector:
        try:
            if dt.StyleType == DimensionStyleType.Linear:
                name = dt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if name:
                    types.append((dt.Id, name.AsString()))
        except:
            pass
    types.sort(key=lambda x: x[1])
    return types


def find_default_dim_type_index(dim_types):
    """Zoek index van default type: eerst 'verkoop' in de naam, dan '1.8'."""
    for i, (dt_id, dt_name) in enumerate(dim_types):
        if "verkoop" in dt_name.lower():
            return i
    for i, (dt_id, dt_name) in enumerate(dim_types):
        if "1.8" in dt_name:
            return i
    return -1  # Geen default gevonden


def get_link_instances():
    """Verzamel geladen RevitLinkInstances: (instance, link_doc, transform)."""
    links = []
    for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        try:
            link_doc = link.GetLinkDocument()
            if link_doc is None:
                continue  # link niet geladen
            if link.IsHidden(active_view):
                continue
            links.append((link, link_doc, link.GetTotalTransform()))
        except:
            pass
    return links


COLUMN_CATEGORIES = [
    BuiltInCategory.OST_Columns,            # bouwkundige kolommen
    BuiltInCategory.OST_StructuralColumns,  # constructieve kolommen
]

def is_afwerkingswand(wall):
    """Detecteer tegelwanden: 'tegel' in de typenaam.

    Bewust alleen op 'tegel': dunne voorzetwanden (gipsplaat e.d.)
    moeten wél gemaatvoerd worden.
    """
    try:
        wt = wall.Document.GetElement(wall.GetTypeId())
        if wt:
            name_param = wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if name_param:
                name = (name_param.AsString() or "").lower()
                if "tegel" in name:
                    return True
    except:
        pass
    return False


def collect_dim_elements(view, links, ignore_afwerking=True):
    """Verzamel wanden en kolommen uit host-document en gelinkte modellen.

    Retourneert lijst van (element, transform, link_instance, kind).
    kind is 'wall' of 'column'; transform en link_instance zijn None
    voor host-elementen. Met ignore_afwerking=True worden afwerkings-
    wanden (tegelwerk/stucwerk) overgeslagen, zodat de maatlijn de
    wand erachter pakt.
    """
    entries = []
    skipped = 0
    for wall in FilteredElementCollector(doc, view.Id).OfClass(Wall):
        if ignore_afwerking and is_afwerkingswand(wall):
            skipped += 1
            continue
        entries.append((wall, None, None, 'wall'))
    wall_count = len(entries)

    for cat in COLUMN_CATEGORIES:
        try:
            for col in FilteredElementCollector(doc, view.Id)\
                    .OfCategory(cat)\
                    .WhereElementIsNotElementType():
                entries.append((col, None, None, 'column'))
        except Exception as ex:
            log.debug("Host columns failed: {}".format(ex))
    log.info("Host: {} walls, {} columns".format(wall_count, len(entries) - wall_count))

    for link, link_doc, transform in links:
        try:
            n_wall = 0
            n_col = 0
            for wall in FilteredElementCollector(link_doc)\
                    .OfClass(Wall)\
                    .WhereElementIsNotElementType():
                if ignore_afwerking and is_afwerkingswand(wall):
                    skipped += 1
                    continue
                entries.append((wall, transform, link, 'wall'))
                n_wall += 1
            for cat in COLUMN_CATEGORIES:
                for col in FilteredElementCollector(link_doc)\
                        .OfCategory(cat)\
                        .WhereElementIsNotElementType():
                    entries.append((col, transform, link, 'column'))
                    n_col += 1
            log.info("Link '{}': {} walls, {} columns".format(
                link_doc.Title, n_wall, n_col))
        except Exception as ex:
            log.debug("Link elements failed: {}".format(ex))

    if skipped:
        log.info("Afwerkingswanden genegeerd: {}".format(skipped))

    return entries


ROOM_TEST_HOOGTE_FT = 2.0  # ~600mm boven vloer/level


def point_in_room(room, point):
    """IsPointInRoom met Z-fallbacks.

    PickPoint klikt op level-hoogte; een room-bounding afwerkvloer
    (tegelvloer in badkamer/toilet) tilt de onderkant van het room-
    volume daar net bovenuit, waardoor IsPointInRoom op klik-Z faalt.
    Daarom hertesten op klik-Z + 600mm en op room-level + 600mm.
    """
    try:
        if room.IsPointInRoom(point):
            return True
        if room.IsPointInRoom(XYZ(point.X, point.Y,
                                  point.Z + ROOM_TEST_HOOGTE_FT)):
            return True
        level = room.Level
        if level:
            adjusted = XYZ(point.X, point.Y,
                           level.Elevation + ROOM_TEST_HOOGTE_FT)
            if room.IsPointInRoom(adjusted):
                return True
    except:
        pass
    return False


def get_room_at_point(point, view, links):
    """Vind room op een punt: eerst host-document, dan gelinkte modellen.

    Retourneert (room, transform). transform is None voor host-rooms,
    anders de link-transform (link-coords -> host-coords).
    """
    rooms = FilteredElementCollector(doc, view.Id)\
        .OfCategory(BuiltInCategory.OST_Rooms)\
        .WhereElementIsNotElementType()\
        .ToElements()

    for room in rooms:
        if point_in_room(room, point):
            return room, None

    for link, link_doc, transform in links:
        try:
            link_point = transform.Inverse.OfPoint(point)
            link_rooms = FilteredElementCollector(link_doc)\
                .OfCategory(BuiltInCategory.OST_Rooms)\
                .WhereElementIsNotElementType()\
                .ToElements()
            for room in link_rooms:
                if point_in_room(room, link_point):
                    return room, transform
        except Exception as ex:
            log.debug("Link rooms failed: {}".format(ex))

    return None, None


def get_room_bounding_box(room, transform=None):
    """Haal bounding box van room, in host-coordinaten."""
    if transform is None:
        bbox = room.get_BoundingBox(active_view)
        if bbox:
            return bbox.Min, bbox.Max
        return None, None

    bbox = room.get_BoundingBox(None)
    if not bbox:
        return None, None
    p1 = transform.OfPoint(bbox.Min)
    p2 = transform.OfPoint(bbox.Max)
    min_pt = XYZ(min(p1.X, p2.X), min(p1.Y, p2.Y), min(p1.Z, p2.Z))
    max_pt = XYZ(max(p1.X, p2.X), max(p1.Y, p2.Y), max(p1.Z, p2.Z))
    return min_pt, max_pt


def get_solids(geom):
    """Haal solids uit een GeometryElement, incl. GeometryInstance (families)."""
    solids = []
    for obj in geom:
        if isinstance(obj, Solid) and obj.Volume > 0:
            solids.append(obj)
        elif isinstance(obj, GeometryInstance):
            try:
                for inst_obj in obj.GetInstanceGeometry():
                    if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                        solids.append(inst_obj)
            except:
                pass
    return solids


def column_ray_hit(column, transform, view, start_point, direction, room_min, room_max):
    """Ray vs kolom-faces: dichtstbijzijnde vertikale face die naar het
    startpunt wijst. Retourneert (face, dist) of (None, None).

    Kolommen hebben geen LocationCurve, dus de ray wordt direct met de
    geometrie-faces gesneden (ray-plane + containment-check via Project).
    """
    options = Options()
    options.ComputeReferences = True
    if transform is None:
        options.View = view

    geom = column.get_Geometry(options)
    if not geom:
        return None, None

    inv = transform.Inverse if transform is not None else None
    margin = 2.0  # ~600mm marge, gelijk aan wand-check
    best_face = None
    best_t = None

    for solid in get_solids(geom):
        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue

            fn = face.FaceNormal
            origin = face.Origin
            if transform is not None:
                fn = transform.OfVector(fn)
                origin = transform.OfPoint(origin)

            if abs(fn.Z) > 0.1:
                continue

            # Face moet naar startpunt wijzen (tegengesteld aan direction)
            face_dot = fn.X * (-direction.X) + fn.Y * (-direction.Y)
            if face_dot < 0.5:
                continue

            # Ray-plane snijpunt
            denom = direction.X * fn.X + direction.Y * fn.Y
            if abs(denom) < 0.0001:
                continue
            t = ((origin.X - start_point.X) * fn.X
                 + (origin.Y - start_point.Y) * fn.Y) / denom
            if t <= 0:
                continue

            hit_x = start_point.X + t * direction.X
            hit_y = start_point.Y + t * direction.Y
            if hit_x < room_min.X - margin or hit_x > room_max.X + margin:
                continue
            if hit_y < room_min.Y - margin or hit_y > room_max.Y + margin:
                continue

            # Containment: ligt het snijpunt echt op deze face (niet alleen
            # op het oneindige vlak)? Alleen horizontaal checken, zodat een
            # Z-offset van klikpunt t.o.v. kolomhoogte niet uitmaakt.
            check_pt = XYZ(hit_x, hit_y, start_point.Z)
            if inv is not None:
                check_pt = inv.OfPoint(check_pt)
            try:
                proj = face.Project(check_pt)
            except:
                proj = None
            if not proj:
                continue
            dx = proj.XYZPoint.X - check_pt.X
            dy = proj.XYZPoint.Y - check_pt.Y
            if math.sqrt(dx * dx + dy * dy) > 0.05:
                continue

            if best_t is None or t < best_t:
                best_t = t
                best_face = face

    return best_face, best_t


def find_wall_faces_in_direction(start_point, direction, wall_entries, view, room_min, room_max):
    """
    Zoek face references van wanden/kolommen in een richting, binnen room
    bounds. Zoekt het dichtstbijzijnde element in de gegeven richting.
    wall_entries: lijst van (element, transform, link_instance, kind) —
    transform/link zijn None voor host-elementen.
    """
    best_ref = None
    best_dist = 1000.0  # Max zoekafstand

    for wall, transform, link, kind in wall_entries:
        if kind == 'column':
            face, t = column_ray_hit(wall, transform, view,
                                     start_point, direction,
                                     room_min, room_max)
            if face is not None and t < best_dist:
                ref = face.Reference
                if ref:
                    if link is not None:
                        ref = ref.CreateLinkReference(link)
                    best_ref = ref
                    best_dist = t
            continue

        wall_loc = wall.Location
        if not isinstance(wall_loc, LocationCurve):
            continue

        wall_curve = wall_loc.Curve
        if transform is not None:
            wall_curve = wall_curve.CreateTransformed(transform)
        ws = wall_curve.GetEndPoint(0)
        we = wall_curve.GetEndPoint(1)
        
        # Wand richting en normaal
        wall_dir = (we - ws).Normalize()
        wall_normal = XYZ(-wall_dir.Y, wall_dir.X, 0)
        
        # Wand moet loodrecht op zoekrichting staan
        dot = abs(direction.X * wall_normal.X + direction.Y * wall_normal.Y)
        if dot < 0.9:
            continue
        
        # Ray-line intersectie: zoek waar ray de wand-lijn kruist
        # Ray: P = start_point + t * direction
        # Lijn: Q = ws + s * (we - ws)
        
        denom = direction.X * (we.Y - ws.Y) - direction.Y * (we.X - ws.X)
        if abs(denom) < 0.0001:
            continue
        
        t = ((ws.X - start_point.X) * (we.Y - ws.Y) - (ws.Y - start_point.Y) * (we.X - ws.X)) / denom
        s = ((ws.X - start_point.X) * direction.Y - (ws.Y - start_point.Y) * direction.X) / denom
        
        # t moet positief zijn (in de richting), s moet tussen 0 en 1 (op de wand)
        if t <= 0 or s < 0 or s > 1:
            continue
        
        # Kruispunt
        hit_x = start_point.X + t * direction.X
        hit_y = start_point.Y + t * direction.Y
        
        # Check of binnen room bounds (met marge)
        margin = 2.0  # ~600mm marge
        if hit_x < room_min.X - margin or hit_x > room_max.X + margin:
            continue
        if hit_y < room_min.Y - margin or hit_y > room_max.Y + margin:
            continue
        
        dist = t  # Afstand is de ray parameter
        
        if dist >= best_dist:
            continue
        
        # Haal face reference
        options = Options()
        options.ComputeReferences = True
        if transform is None:
            options.View = view
        # Voor link-wanden geen View zetten: Options.View moet uit hetzelfde
        # document komen als het element — default detail level volstaat

        geom = wall.get_Geometry(options)
        if not geom:
            continue

        for geom_obj in geom:
            if not isinstance(geom_obj, Solid) or geom_obj.Volume <= 0:
                continue

            for face in geom_obj.Faces:
                if not isinstance(face, PlanarFace):
                    continue

                fn = face.FaceNormal
                if transform is not None:
                    fn = transform.OfVector(fn)
                if abs(fn.Z) > 0.1:
                    continue

                # Face moet naar startpunt wijzen (tegengesteld aan direction)
                face_dot = fn.X * (-direction.X) + fn.Y * (-direction.Y)
                if face_dot < 0.5:
                    continue

                ref = face.Reference
                if ref:
                    if link is not None:
                        ref = ref.CreateLinkReference(link)
                    best_ref = ref
                    best_dist = dist
                    break

    return best_ref, best_dist


def create_room_dimensions(click_point, room, walls, view, dim_type_id, room_transform=None):
    """
    Maak horizontale en verticale maatlijnen die de room breed/hoog beslaan.
    De maatlijnen lopen door het klikpunt.
    room_transform: link-transform als de room uit een gelinkt model komt.
    """
    log.info("Creating dimensions at ({:.2f}, {:.2f})".format(click_point.X, click_point.Y))

    min_pt, max_pt = get_room_bounding_box(room, room_transform)
    if not min_pt or not max_pt:
        return None, None, "Kon room bounds niet bepalen"
    
    log.debug("Room bounds: ({:.2f},{:.2f}) to ({:.2f},{:.2f})".format(
        min_pt.X, min_pt.Y, max_pt.X, max_pt.Y
    ))
    
    dim_h = None
    dim_v = None
    
    # Richtingen
    dir_left = XYZ(-1, 0, 0)
    dir_right = XYZ(1, 0, 0)
    dir_down = XYZ(0, -1, 0)
    dir_up = XYZ(0, 1, 0)
    
    # HORIZONTALE MAATLIJN (volledige breedte)
    # Start van midden Y van room, zoek links en rechts
    h_start = XYZ(click_point.X, click_point.Y, click_point.Z)
    
    ref_left, dist_left = find_wall_faces_in_direction(h_start, dir_left, walls, view, min_pt, max_pt)
    ref_right, dist_right = find_wall_faces_in_direction(h_start, dir_right, walls, view, min_pt, max_pt)
    
    log.debug("H: left={:.2f}, right={:.2f}".format(dist_left, dist_right))
    
    if ref_left and ref_right:
        ref_array_h = ReferenceArray()
        ref_array_h.Append(ref_left)
        ref_array_h.Append(ref_right)
        
        # Lijn van links naar rechts door klikpunt Y
        line_h = Line.CreateBound(
            XYZ(click_point.X - dist_left - 1, click_point.Y, click_point.Z),
            XYZ(click_point.X + dist_right + 1, click_point.Y, click_point.Z)
        )
        
        try:
            if dim_type_id:
                dim_type = doc.GetElement(dim_type_id)
                dim_h = doc.Create.NewDimension(view, line_h, ref_array_h, dim_type)
            else:
                dim_h = doc.Create.NewDimension(view, line_h, ref_array_h)
            
            if dim_h:
                log.info("H dim: {}".format(dim_h.Id.IntegerValue))
        except Exception as ex:
            log.debug("H failed: {}".format(ex))
    
    # VERTICALE MAATLIJN (volledige hoogte)
    v_start = XYZ(click_point.X, click_point.Y, click_point.Z)
    
    ref_down, dist_down = find_wall_faces_in_direction(v_start, dir_down, walls, view, min_pt, max_pt)
    ref_up, dist_up = find_wall_faces_in_direction(v_start, dir_up, walls, view, min_pt, max_pt)
    
    log.debug("V: down={:.2f}, up={:.2f}".format(dist_down, dist_up))
    
    if ref_down and ref_up:
        ref_array_v = ReferenceArray()
        ref_array_v.Append(ref_down)
        ref_array_v.Append(ref_up)
        
        # Lijn van onder naar boven door klikpunt X
        line_v = Line.CreateBound(
            XYZ(click_point.X, click_point.Y - dist_down - 1, click_point.Z),
            XYZ(click_point.X, click_point.Y + dist_up + 1, click_point.Z)
        )
        
        try:
            if dim_type_id:
                dim_type = doc.GetElement(dim_type_id)
                dim_v = doc.Create.NewDimension(view, line_v, ref_array_v, dim_type)
            else:
                dim_v = doc.Create.NewDimension(view, line_v, ref_array_v)
            
            if dim_v:
                log.info("V dim: {}".format(dim_v.Id.IntegerValue))
        except Exception as ex:
            log.debug("V failed: {}".format(ex))
    
    if not dim_h and not dim_v:
        return None, None, "Kon geen wanden vinden voor maatlijnen"
    
    return dim_h, dim_v, None


# =============================================================================
# UI
# =============================================================================
class CrossDimForm(BaseForm):
    def __init__(self):
        super(CrossDimForm, self).__init__(
            title="CrossDim",
            width=450,
            height=500,
            show_header=True,
            show_footer=True
        )
        self.set_subtitle("Room maatvoering")
        self.options = None
        self.dim_types = get_dimension_types()
        self._setup_ui()
    
    def _setup_ui(self):
        m = DPIScaler.scale(20)
        y = DPIScaler.scale(10)
        
        # Dimension Type
        lbl_type = UIFactory.create_label("Maatlijn type", font_size=12, bold=True, color=Huisstijl.VIOLET)
        lbl_type.Location = Point(m, y)
        self.pnl_content.Controls.Add(lbl_type)
        y += DPIScaler.scale(30)
        
        self.cmb_type = ComboBox()
        self.cmb_type.Location = Point(m, y)
        self.cmb_type.Size = DPIScaler.scale_size(380, 28)
        self.cmb_type.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_type.DropDownHeight = DPIScaler.scale(200)
        
        for dt_id, dt_name in self.dim_types:
            self.cmb_type.Items.Add(dt_name)
        
        # Standaard: eerste met "1.8" in naam, anders geen selectie
        default_idx = find_default_dim_type_index(self.dim_types)
        if default_idx >= 0:
            self.cmb_type.SelectedIndex = default_idx
        
        self.pnl_content.Controls.Add(self.cmb_type)
        y += DPIScaler.scale(55)

        # Tegelwanden negeren
        self.chk_afwerking = UIFactory.create_checkbox(
            "Negeer tegelwanden ('tegel' in typenaam)", checked=True
        )
        self.chk_afwerking.Location = Point(m, y)
        self.pnl_content.Controls.Add(self.chk_afwerking)
        y += DPIScaler.scale(40)

        # Info
        info = UIFactory.create_label(
            "Werkwijze:\n"
            "1. Klik in een room\n"
            "2. Horizontale en verticale maatlijnen worden geplaatst\n"
            "   over de volledige breedte en hoogte van de room\n"
            "3. Ga door naar volgende room\n"
            "4. Druk ESC om te stoppen\n"
            "\n"
            "Tegelwanden ('tegel' in typenaam) worden genegeerd:\n"
            "de maatlijn pakt de wand erachter.",
            font_size=9, italic=True, color=Huisstijl.TEXT_SECONDARY
        )
        info.Location = Point(m, y)
        info.MaximumSize = Size(DPIScaler.scale(380), 0)
        info.AutoSize = True
        self.pnl_content.Controls.Add(info)
        
        self.add_footer_button("Start", 'primary', self._on_run, 120)
    
    def _on_run(self, sender, args):
        dim_type_id = None
        if self.cmb_type.SelectedIndex >= 0 and self.cmb_type.SelectedIndex < len(self.dim_types):
            dim_type_id = self.dim_types[self.cmb_type.SelectedIndex][0]
        
        self.options = {
            'dim_type_id': dim_type_id,
            'ignore_afwerking': self.chk_afwerking.Checked
        }
        self.DialogResult = DialogResult.OK
        self.Close()


# =============================================================================
# MAIN
# =============================================================================
def main():
    global doc, uidoc, active_view
    
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    uidoc = revit.uidoc
    
    if not doc:
        forms.alert("Open eerst een Revit project.", title="CrossDim")
        return
    
    active_view = doc.ActiveView
    
    log.log_revit_info()
    log.section("Main")
    
    if not isinstance(active_view, ViewPlan):
        forms.alert("Alleen in plattegronden.", title="CrossDim")
        return
    
    form = CrossDimForm()
    if form.ShowDialog() != DialogResult.OK:
        return
    
    options = form.options
    log.log_options(options)
    
    links = get_link_instances()
    log.info("Linked models: {}".format(len(links)))

    walls = collect_dim_elements(active_view, links,
                                 options.get('ignore_afwerking', True))
    log.info("Elementen totaal (host + links, wanden + kolommen): {}".format(len(walls)))

    rooms = list(FilteredElementCollector(doc, active_view.Id)\
        .OfCategory(BuiltInCategory.OST_Rooms)\
        .WhereElementIsNotElementType().ToElements())
    log.info("Rooms in view: {}".format(len(rooms)))

    if not rooms:
        has_link_rooms = False
        for link, link_doc, transform in links:
            try:
                count = FilteredElementCollector(link_doc)\
                    .OfCategory(BuiltInCategory.OST_Rooms)\
                    .WhereElementIsNotElementType()\
                    .GetElementCount()
                if count > 0:
                    has_link_rooms = True
                    break
            except:
                pass
        if not has_link_rooms:
            forms.alert("Geen rooms in view of in gelinkte modellen.", title="CrossDim")
            return
    
    dim_count = 0
    room_count = 0
    miss_count = 0

    while True:
        try:
            click_point = uidoc.Selection.PickPoint(
                ObjectSnapTypes.None,
                "Klik in een room (ESC om te stoppen)"
            )
            
            room, room_transform = get_room_at_point(click_point, active_view, links)

            if not room:
                miss_count += 1
                log.info("Geen room op klikpunt ({:.2f}, {:.2f})".format(
                    click_point.X, click_point.Y))
                continue

            room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or "Unnamed"
            log.info("Room: {}{}".format(room_name, " (link)" if room_transform else ""))

            with revit.Transaction("CrossDim - {}".format(room_name)):
                dim_h, dim_v, err = create_room_dimensions(
                    click_point, room, walls, active_view,
                    options.get('dim_type_id'), room_transform
                )
                
                if not err:
                    room_count += 1
                    if dim_h:
                        dim_count += 1
                    if dim_v:
                        dim_count += 1

            # Direct tonen zodat de gebruiker het resultaat ziet
            # voordat de volgende klik gevraagd wordt
            try:
                uidoc.RefreshActiveView()
            except:
                pass

        except OperationCanceledException:
            break
        except Exception as ex:
            if "cancel" in str(ex).lower() or "aborted" in str(ex).lower():
                break
            log.exception("Error")
            break
    
    if dim_count > 0 or miss_count > 0:
        msg = "{} maatlijn(en) in {} room(s) geplaatst!".format(dim_count, room_count)
        if miss_count:
            msg += "\n\n{} klik(s) zonder room genegeerd.".format(miss_count)
        forms.alert(msg, title="CrossDim")

    log.finalize(True, "{} dimensions in {} rooms, {} misses".format(
        dim_count, room_count, miss_count))


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log.exception("Error")
        log.finalize(False)
        raise
