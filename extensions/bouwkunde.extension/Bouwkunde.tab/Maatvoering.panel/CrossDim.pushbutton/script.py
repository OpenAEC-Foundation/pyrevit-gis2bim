# -*- coding: utf-8 -*-
"""CrossDim - Room Maatvoering

Plaats kruisende maatlijnen in rooms met één klik.
Klik in een room om horizontale en verticale maatlijnen te plaatsen.

Auteur: 3BM Bouwkunde
Versie: 1.2.0 - Ondersteuning voor gelinkte modellen (wanden + rooms uit links)
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
log.info("CrossDim v1.2.0")

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
    """Zoek index van eerste type met '1.8' in de naam."""
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


def collect_walls(view, links):
    """Verzamel wanden uit host-document en gelinkte modellen.

    Retourneert lijst van (wall, transform, link_instance).
    transform en link_instance zijn None voor host-wanden.
    """
    entries = []
    for wall in FilteredElementCollector(doc, view.Id).OfClass(Wall):
        entries.append((wall, None, None))
    log.info("Host walls: {}".format(len(entries)))

    for link, link_doc, transform in links:
        try:
            link_walls = list(FilteredElementCollector(link_doc)
                              .OfClass(Wall)
                              .WhereElementIsNotElementType())
            log.info("Link '{}': {} walls".format(link_doc.Title, len(link_walls)))
            for wall in link_walls:
                entries.append((wall, transform, link))
        except Exception as ex:
            log.debug("Link walls failed: {}".format(ex))

    return entries


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
        if room.IsPointInRoom(point):
            return room, None

    for link, link_doc, transform in links:
        try:
            link_point = transform.Inverse.OfPoint(point)
            link_rooms = FilteredElementCollector(link_doc)\
                .OfCategory(BuiltInCategory.OST_Rooms)\
                .WhereElementIsNotElementType()\
                .ToElements()
            for room in link_rooms:
                try:
                    if room.IsPointInRoom(link_point):
                        return room, transform
                    # Klikpunt-Z kan buiten de room-hoogte vallen (verticale
                    # link-offset) — hertest op level-hoogte van de room
                    level = room.Level
                    if level:
                        adjusted = XYZ(link_point.X, link_point.Y,
                                       level.Elevation + 1.0)
                        if room.IsPointInRoom(adjusted):
                            return room, transform
                except:
                    pass
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


def find_wall_faces_in_direction(start_point, direction, wall_entries, view, room_min, room_max):
    """
    Zoek wand face references in een richting, binnen room bounds.
    Zoekt de dichtstbijzijnde wand in de gegeven richting.
    wall_entries: lijst van (wall, transform, link_instance) — transform/link
    zijn None voor host-wanden, gevuld voor wanden uit gelinkte modellen.
    """
    best_ref = None
    best_dist = 1000.0  # Max zoekafstand

    for wall, transform, link in wall_entries:
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
        
        # Info
        info = UIFactory.create_label(
            "Werkwijze:\n"
            "1. Klik in een room\n"
            "2. Horizontale en verticale maatlijnen worden geplaatst\n"
            "   over de volledige breedte en hoogte van de room\n"
            "3. Ga door naar volgende room\n"
            "4. Druk ESC om te stoppen",
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
            'dim_type_id': dim_type_id
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

    walls = collect_walls(active_view, links)
    log.info("Walls totaal (host + links): {}".format(len(walls)))

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
    
    while True:
        try:
            click_point = uidoc.Selection.PickPoint(
                ObjectSnapTypes.None,
                "Klik in een room (ESC om te stoppen)"
            )
            
            room, room_transform = get_room_at_point(click_point, active_view, links)

            if not room:
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
        
        except OperationCanceledException:
            break
        except Exception as ex:
            if "cancel" in str(ex).lower() or "aborted" in str(ex).lower():
                break
            log.exception("Error")
            break
    
    if dim_count > 0:
        forms.alert("{} maatlijn(en) in {} room(s) geplaatst!".format(dim_count, room_count), title="CrossDim")
    
    log.finalize(True, "{} dimensions in {} rooms".format(dim_count, room_count))


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log.exception("Error")
        log.finalize(False)
        raise
