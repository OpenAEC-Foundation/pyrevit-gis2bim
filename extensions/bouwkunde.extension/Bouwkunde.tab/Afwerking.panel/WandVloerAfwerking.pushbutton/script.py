# -*- coding: utf-8 -*-
"""
Wand & Vloer Afwerking Tool
===========================
Automatisch wand-, vloer- en plafondafwerking modelleren per ruimte.
Klik op ruimtes om ze toe te voegen, ESC om te stoppen.
"""
__title__ = "Wand Vloer\nAfwerking"
__author__ = "3BM Bouwkunde"

# ==============================================================================
# IMPORTS
# ==============================================================================
import sys
import os

# Voeg lib toe aan path voor ui_template
lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl
from bm_logger import get_logger

log = get_logger("WandVloerAfwerking")

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    ComboBox, TextBox, CheckBox, RadioButton, DialogResult,
    ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    WallType, FloorType, Floor, Wall, Ceiling, CeilingType, Element, Level,
    Transaction, ElementId, XYZ, Line, CurveLoop, CurveArray,
    SpatialElementBoundaryOptions, SpatialElementBoundaryLocation,
    WallLocationLine
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

# GEEN doc/uidoc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
uidoc = None
output = script.get_output()

# Conversie
MM_TO_FEET = 1.0 / 304.8
FEET_TO_MM = 304.8


# ==============================================================================
# SELECTION FILTER
# ==============================================================================
class RoomSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if isinstance(elem, Room):
                return elem.Area > 0
        except:
            pass
        return False
    
    def AllowReference(self, reference, position):
        return False


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def get_element_name(elem):
    """Veilig element naam ophalen"""
    try:
        name_param = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if name_param and name_param.AsString():
            return name_param.AsString()
        name_param2 = elem.LookupParameter("Type Name")
        if name_param2 and name_param2.AsString():
            return name_param2.AsString()
    except:
        pass
    return "Onbekend"


def get_assembly_code(elem):
    """Haal Assembly Code (SfB) op van element type"""
    try:
        param = elem.get_Parameter(BuiltInParameter.UNIFORMAT_CODE)
        if param and param.AsString():
            return param.AsString()
        param2 = elem.LookupParameter("Assembly Code")
        if param2 and param2.AsString():
            return param2.AsString()
        param3 = elem.LookupParameter("NL-SfB")
        if param3 and param3.AsString():
            return param3.AsString()
    except:
        pass
    return ""


def check_sfb_code(assembly_code, sfb_number):
    """Check of assembly code het SfB nummer bevat (bijv. '42' in '2E(42.12)')"""
    if not assembly_code or not sfb_number:
        return False
    return sfb_number in assembly_code


def get_wall_types(sfb_filter="42"):
    """Haal wall types, gefilterd op SfB code (bijv. 42)"""
    collector = FilteredElementCollector(doc).OfClass(WallType)
    types_filtered = []
    types_all = []
    
    for wt in collector:
        name = get_element_name(wt)
        if name == "Onbekend":
            continue
        types_all.append((wt.Id, name))
        if sfb_filter:
            assembly_code = get_assembly_code(wt)
            if assembly_code and check_sfb_code(assembly_code, sfb_filter):
                types_filtered.append((wt.Id, name))
    
    if types_filtered:
        output.print_md("*SfB {} wandtypes gevonden: {}*".format(sfb_filter, len(types_filtered)))
        return sorted(types_filtered, key=lambda x: x[1])
    output.print_md("*Geen SfB {} wandtypes, {} totaal geladen*".format(sfb_filter, len(types_all)))
    return sorted(types_all, key=lambda x: x[1])


def get_floor_types(sfb_filter="43"):
    """Haal floor types, gefilterd op SfB code (bijv. 43)"""
    collector = FilteredElementCollector(doc).OfClass(FloorType)
    types_filtered = []
    types_all = []
    
    for ft in collector:
        name = get_element_name(ft)
        if name == "Onbekend":
            continue
        types_all.append((ft.Id, name))
        if sfb_filter:
            assembly_code = get_assembly_code(ft)
            if assembly_code and check_sfb_code(assembly_code, sfb_filter):
                types_filtered.append((ft.Id, name))
    
    if types_filtered:
        output.print_md("*SfB {} vloertypes gevonden: {}*".format(sfb_filter, len(types_filtered)))
        return sorted(types_filtered, key=lambda x: x[1])
    output.print_md("*Geen SfB {} vloertypes, {} totaal geladen*".format(sfb_filter, len(types_all)))
    return sorted(types_all, key=lambda x: x[1])


def get_ceiling_types():
    """Haal alle ceiling types (geen filter)"""
    collector = FilteredElementCollector(doc).OfClass(CeilingType)
    types_all = []
    for ct in collector:
        name = get_element_name(ct)
        if name == "Onbekend":
            continue
        types_all.append((ct.Id, name))
    return sorted(types_all, key=lambda x: x[1])


def get_type_thickness(type_id):
    """Haal dikte van type in feet (voor floor/ceiling)"""
    try:
        elem_type = doc.GetElement(type_id)
        cs = elem_type.GetCompoundStructure()
        if cs:
            return cs.GetWidth()
    except:
        pass
    return 10 * MM_TO_FEET


def get_room_base_offset(room):
    """Haal room base offset tov level"""
    try:
        param = room.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET)
        if param:
            return param.AsDouble()
    except:
        pass
    return 0.0


def get_room_boundary_segments(room):
    """Haal room boundary segments"""
    options = SpatialElementBoundaryOptions()
    options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
    segments = room.GetBoundarySegments(options)
    if segments and len(segments) > 0:
        return segments[0]
    return None


def get_room_height(room):
    """Bepaal room hoogte (unbounded height)"""
    try:
        param = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT)
        if param:
            return param.AsDouble()
    except:
        pass
    return 2.8 / 0.3048


def get_room_upper_level_elevation(room):
    """Haal elevatie van bovenliggend level"""
    try:
        upper_limit = room.UpperLimit
        if upper_limit:
            limit_offset = room.LimitOffset
            return upper_limit.Elevation + limit_offset
    except:
        pass
    level = doc.GetElement(room.LevelId)
    return level.Elevation + get_room_height(room)


def create_floor_finish(room, floor_type_id):
    """Maak vloerafwerking voor room"""
    segments = get_room_boundary_segments(room)
    if not segments:
        return None, "Geen boundary gevonden"
    
    try:
        curve_loop = CurveLoop()
        for seg in segments:
            curve_loop.Append(seg.GetCurve())
        
        level_id = room.LevelId
        floor_thickness = get_type_thickness(floor_type_id)
        floor = Floor.Create(doc, [curve_loop], floor_type_id, level_id)
        doc.Regenerate()
        
        offset_param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
        if offset_param and not offset_param.IsReadOnly:
            offset_param.Set(floor_thickness)
        
        comments = floor.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if comments:
            comments.Set("3BM_Afwerking_Vloer_{}".format(room.Number))
        
        return floor, None
    except Exception as e:
        return None, str(e)


def create_ceiling_finish(room, ceiling_type_id):
    """Maak plafondafwerking voor room"""
    segments = get_room_boundary_segments(room)
    if not segments:
        return None, "Geen boundary gevonden"
    
    try:
        curve_loop = CurveLoop()
        for seg in segments:
            curve_loop.Append(seg.GetCurve())
        
        level_id = room.LevelId
        level = doc.GetElement(level_id)
        upper_elevation = get_room_upper_level_elevation(room)
        ceiling_offset = upper_elevation - level.Elevation
        
        ceiling = Ceiling.Create(doc, [curve_loop], ceiling_type_id, level_id)
        
        height_param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
        if height_param:
            height_param.Set(ceiling_offset)
        
        comments = ceiling.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if comments:
            comments.Set("3BM_Afwerking_Plafond_{}".format(room.Number))
        
        return ceiling, None
    except Exception as e:
        return None, str(e)


def get_wall_type_thickness(wall_type_id):
    """Haal dikte van wall type in feet"""
    try:
        wall_type = doc.GetElement(wall_type_id)
        return wall_type.Width
    except:
        pass
    return 10 * MM_TO_FEET


def offset_curve_to_room_center(curve, offset_distance, room):
    """Offset een curve naar het centrum van de room"""
    room_bb = room.get_BoundingBox(None)
    if not room_bb:
        return curve
    
    room_center = XYZ(
        (room_bb.Min.X + room_bb.Max.X) / 2,
        (room_bb.Min.Y + room_bb.Max.Y) / 2,
        (room_bb.Min.Z + room_bb.Max.Z) / 2
    )
    
    curve_mid = curve.Evaluate(0.5, True)
    direction = XYZ(room_center.X - curve_mid.X, room_center.Y - curve_mid.Y, 0)
    length = (direction.X**2 + direction.Y**2)**0.5
    
    if length < 0.001:
        return curve
    
    direction = XYZ(direction.X / length, direction.Y / length, 0)
    offset_vec = XYZ(direction.X * offset_distance, direction.Y * offset_distance, 0)
    
    if hasattr(curve, 'GetEndPoint'):
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        new_start = XYZ(start.X + offset_vec.X, start.Y + offset_vec.Y, start.Z)
        new_end = XYZ(end.X + offset_vec.X, end.Y + offset_vec.Y, end.Z)
        return Line.CreateBound(new_start, new_end)
    
    return curve


def create_wall_finish(room, wall_type_id, height_mm, tot_plafond, floor_thickness, ceiling_thickness, has_ceiling):
    """Maak wandafwerking rondom room met Finish Face Exterior op room boundary"""
    segments = get_room_boundary_segments(room)
    if not segments:
        return [], "Geen boundary gevonden"
    
    level_id = room.LevelId
    level = doc.GetElement(level_id)
    
    if tot_plafond:
        upper_elevation = get_room_upper_level_elevation(room)
        room_base = level.Elevation + get_room_base_offset(room)
        if has_ceiling:
            wall_height = (upper_elevation - ceiling_thickness) - room_base - floor_thickness
        else:
            wall_height = upper_elevation - room_base - floor_thickness
    else:
        wall_height = height_mm * MM_TO_FEET
    
    base_offset = get_room_base_offset(room) + floor_thickness
    wall_thickness = get_wall_type_thickness(wall_type_id)
    
    created_walls = []
    errors = []
    
    for seg in segments:
        curve = seg.GetCurve()
        try:
            offset_curve = offset_curve_to_room_center(curve, wall_thickness / 2.0, room)
            
            wall = Wall.Create(doc, offset_curve, wall_type_id, level_id,
                              wall_height, base_offset, False, False)
            
            location_param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
            if location_param:
                location_param.Set(2)
            
            comments = wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if comments:
                comments.Set("3BM_Afwerking_Wand_{}".format(room.Number))
            
            created_walls.append(wall)
        except Exception as e:
            errors.append(str(e))
    
    return created_walls, errors


def pick_rooms():
    """Laat gebruiker ruimtes aanklikken tot ESC"""
    rooms = []
    room_filter = RoomSelectionFilter()
    
    while True:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                room_filter,
                "Klik op ruimte (ESC = klaar, {} geselecteerd)".format(len(rooms))
            )
            room = doc.GetElement(ref.ElementId)
            if room and room.Id not in [r.Id for r in rooms]:
                rooms.append(room)
        except OperationCanceledException:
            break
        except:
            break
    
    return rooms


# ==============================================================================
# UI FORM - MET UI_TEMPLATE
# ==============================================================================
class AfwerkingForm(BaseForm):
    """Configuratie formulier met 3BM huisstijl"""
    
    def __init__(self, wall_types, floor_types, ceiling_types):
        self.wall_types = wall_types
        self.floor_types = floor_types
        self.ceiling_types = ceiling_types
        self.result = None
        
        super(AfwerkingForm, self).__init__(
            "Wand, Vloer & Plafond Afwerking",
            width=520,
            height=480
        )
        self.set_subtitle("Selecteer afwerkingstypes")
        self._setup_ui()
    
    def _setup_ui(self):
        y = 10
        
        # === VLOER ===
        self.chk_floor = UIFactory.create_checkbox("Vloerafwerking", checked=True)
        self.chk_floor.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.chk_floor.Location = DPIScaler.scale_point(10, y)
        self.chk_floor.CheckedChanged += self._on_floor_changed
        self.pnl_content.Controls.Add(self.chk_floor)
        y += 28
        
        self.cmb_floor = UIFactory.create_combobox(440)
        self.cmb_floor.Location = DPIScaler.scale_point(30, y)
        for _, name in self.floor_types:
            self.cmb_floor.Items.Add(name)
        if self.cmb_floor.Items.Count > 0:
            self.cmb_floor.SelectedIndex = 0
        self.pnl_content.Controls.Add(self.cmb_floor)
        y += 38
        
        # === WAND ===
        self.chk_wall = UIFactory.create_checkbox("Wandafwerking", checked=True)
        self.chk_wall.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.chk_wall.Location = DPIScaler.scale_point(10, y)
        self.chk_wall.CheckedChanged += self._on_wall_changed
        self.pnl_content.Controls.Add(self.chk_wall)
        y += 28
        
        self.cmb_wall = UIFactory.create_combobox(440)
        self.cmb_wall.Location = DPIScaler.scale_point(30, y)
        for _, name in self.wall_types:
            self.cmb_wall.Items.Add(name)
        if self.cmb_wall.Items.Count > 0:
            self.cmb_wall.SelectedIndex = 0
        self.pnl_content.Controls.Add(self.cmb_wall)
        y += 32
        
        # Hoogte opties
        lbl_height = UIFactory.create_label("Hoogte:")
        lbl_height.Location = DPIScaler.scale_point(30, y)
        self.pnl_content.Controls.Add(lbl_height)
        y += 24
        
        self.rb_plafond = RadioButton()
        self.rb_plafond.Text = "Tot plafond (of bovenliggend level)"
        self.rb_plafond.Font = Font("Segoe UI", 10)
        self.rb_plafond.Location = DPIScaler.scale_point(30, y)
        self.rb_plafond.AutoSize = True
        self.rb_plafond.Checked = True
        self.rb_plafond.CheckedChanged += self._on_height_changed
        self.pnl_content.Controls.Add(self.rb_plafond)
        y += 24
        
        self.rb_custom = RadioButton()
        self.rb_custom.Text = "Vaste hoogte:"
        self.rb_custom.Font = Font("Segoe UI", 10)
        self.rb_custom.Location = DPIScaler.scale_point(30, y)
        self.rb_custom.AutoSize = True
        self.pnl_content.Controls.Add(self.rb_custom)
        
        self.txt_height = UIFactory.create_textbox(60)
        self.txt_height.Text = "1200"
        self.txt_height.Location = DPIScaler.scale_point(145, y - 2)
        self.txt_height.Enabled = False
        self.pnl_content.Controls.Add(self.txt_height)
        
        lbl_mm = UIFactory.create_label("mm")
        lbl_mm.Location = DPIScaler.scale_point(210, y)
        self.pnl_content.Controls.Add(lbl_mm)
        y += 38
        
        # === PLAFOND ===
        self.chk_ceiling = UIFactory.create_checkbox("Plafondafwerking", checked=False)
        self.chk_ceiling.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.chk_ceiling.Location = DPIScaler.scale_point(10, y)
        self.chk_ceiling.CheckedChanged += self._on_ceiling_changed
        self.pnl_content.Controls.Add(self.chk_ceiling)
        y += 28
        
        self.cmb_ceiling = UIFactory.create_combobox(440)
        self.cmb_ceiling.Location = DPIScaler.scale_point(30, y)
        self.cmb_ceiling.Enabled = False
        for _, name in self.ceiling_types:
            self.cmb_ceiling.Items.Add(name)
        if self.cmb_ceiling.Items.Count > 0:
            self.cmb_ceiling.SelectedIndex = 0
        self.pnl_content.Controls.Add(self.cmb_ceiling)
        y += 40
        
        # Info label
        lbl_info = UIFactory.create_label(
            "Tip: Bij plafond wordt wandhoogte tot onderzijde plafond",
            color=Huisstijl.TEXT_SECONDARY
        )
        lbl_info.Location = DPIScaler.scale_point(10, y)
        self.pnl_content.Controls.Add(lbl_info)
        
        # Footer buttons - vervang standaard sluiten knop
        self.btn_close.Text = "Annuleren"
        self.add_footer_button("Selecteer Ruimtes", 'primary', self._on_select_click, 160)
    
    def _on_floor_changed(self, sender, args):
        self.cmb_floor.Enabled = self.chk_floor.Checked
    
    def _on_wall_changed(self, sender, args):
        self.cmb_wall.Enabled = self.chk_wall.Checked
        self.rb_plafond.Enabled = self.chk_wall.Checked
        self.rb_custom.Enabled = self.chk_wall.Checked
        self.txt_height.Enabled = self.chk_wall.Checked and self.rb_custom.Checked
    
    def _on_ceiling_changed(self, sender, args):
        self.cmb_ceiling.Enabled = self.chk_ceiling.Checked
    
    def _on_height_changed(self, sender, args):
        self.txt_height.Enabled = self.rb_custom.Checked
    
    def _on_select_click(self, sender, args):
        if not self.chk_floor.Checked and not self.chk_wall.Checked and not self.chk_ceiling.Checked:
            self.show_warning("Selecteer minimaal één afwerkingstype.")
            return
        
        if self.chk_wall.Checked and self.rb_custom.Checked:
            try:
                height = int(self.txt_height.Text)
                if height < 100 or height > 5000:
                    self.show_warning("Hoogte moet tussen 100 en 5000 mm zijn.")
                    return
            except:
                self.show_warning("Ongeldige hoogte.")
                return
        
        floor_idx = self.cmb_floor.SelectedIndex
        wall_idx = self.cmb_wall.SelectedIndex
        ceiling_idx = self.cmb_ceiling.SelectedIndex
        
        self.result = {
            'do_floor': self.chk_floor.Checked,
            'floor_type_id': self.floor_types[floor_idx][0] if floor_idx >= 0 else None,
            'floor_type_name': self.floor_types[floor_idx][1] if floor_idx >= 0 else "",
            'do_wall': self.chk_wall.Checked,
            'wall_type_id': self.wall_types[wall_idx][0] if wall_idx >= 0 else None,
            'wall_type_name': self.wall_types[wall_idx][1] if wall_idx >= 0 else "",
            'tot_plafond': self.rb_plafond.Checked,
            'height_mm': int(self.txt_height.Text) if self.rb_custom.Checked else 0,
            'do_ceiling': self.chk_ceiling.Checked,
            'ceiling_type_id': self.ceiling_types[ceiling_idx][0] if ceiling_idx >= 0 and self.ceiling_types else None,
            'ceiling_type_name': self.ceiling_types[ceiling_idx][1] if ceiling_idx >= 0 and self.ceiling_types else "",
        }
        
        self.DialogResult = DialogResult.OK
        self.Close()


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    global doc, uidoc
    
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    uidoc = revit.uidoc
    
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Wand Vloer Afwerking")
        return
    
    output.print_md("# Wand, Vloer & Plafond Afwerking")
    output.print_md("*3BM Bouwkunde*")
    
    wall_types = get_wall_types("42")
    floor_types = get_floor_types("43")
    ceiling_types = get_ceiling_types()
    
    if not wall_types and not floor_types:
        forms.alert("Geen wand- of vloertypes gevonden.", exitscript=True)
        return
    
    form = AfwerkingForm(wall_types, floor_types, ceiling_types)
    result = form.ShowDialog()
    
    if result != DialogResult.OK or not form.result:
        return
    
    config = form.result
    output.print_md("")
    output.print_md("## Configuratie")
    if config['do_floor']:
        output.print_md("- **Vloer:** {}".format(config['floor_type_name']))
    if config['do_wall']:
        output.print_md("- **Wand:** {}".format(config['wall_type_name']))
        output.print_md("- **Hoogte:** {}".format(
            "Tot plafond/level" if config['tot_plafond'] else "{} mm".format(config['height_mm'])
        ))
    if config['do_ceiling']:
        output.print_md("- **Plafond:** {}".format(config['ceiling_type_name']))
    output.print_md("")
    
    output.print_md("## Selecteer ruimtes")
    output.print_md("*Klik op ruimtes, druk ESC om te stoppen*")
    
    rooms = pick_rooms()
    
    if not rooms:
        forms.alert("Geen ruimtes geselecteerd.")
        return
    
    output.print_md("")
    output.print_md("**{} ruimte(s) geselecteerd:**".format(len(rooms)))
    for r in rooms:
        rname = r.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or ""
        output.print_md("- {} - {}".format(r.Number, rname))
    
    if not forms.alert("Afwerking toepassen op {} ruimte(s)?".format(len(rooms)), 
                       yes=True, no=True):
        return
    
    output.print_md("")
    output.print_md("## Uitvoeren")
    
    floor_thickness = 0.0
    ceiling_thickness = 0.0
    
    if config['do_floor']:
        floor_thickness = get_type_thickness(config['floor_type_id'])
        output.print_md("*Vloerdikte: {:.0f} mm*".format(floor_thickness * FEET_TO_MM))
    
    if config['do_ceiling']:
        ceiling_thickness = get_type_thickness(config['ceiling_type_id'])
        output.print_md("*Plafonddikte: {:.0f} mm*".format(ceiling_thickness * FEET_TO_MM))
    
    created_floors = 0
    created_walls = 0
    created_ceilings = 0
    all_errors = []
    
    with revit.Transaction("Wand/Vloer/Plafond Afwerking"):
        for room in rooms:
            rname = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or ""
            rnumber = room.Number or "?"
            
            output.print_md("")
            output.print_md("### {} - {}".format(rnumber, rname))
            
            if config['do_floor']:
                floor, err = create_floor_finish(room, config['floor_type_id'])
                if floor:
                    created_floors += 1
                    output.print_md("✓ Vloer aangemaakt")
                else:
                    output.print_md("✗ Vloer mislukt: {}".format(err))
                    all_errors.append("Vloer {}: {}".format(rnumber, err))
            
            if config['do_ceiling']:
                ceiling, err = create_ceiling_finish(room, config['ceiling_type_id'])
                if ceiling:
                    created_ceilings += 1
                    output.print_md("✓ Plafond aangemaakt")
                else:
                    output.print_md("✗ Plafond mislukt: {}".format(err))
                    all_errors.append("Plafond {}: {}".format(rnumber, err))
            
            if config['do_wall']:
                walls, errors = create_wall_finish(
                    room, config['wall_type_id'], config['height_mm'],
                    config['tot_plafond'], floor_thickness,
                    ceiling_thickness, config['do_ceiling']
                )
                if walls:
                    created_walls += len(walls)
                    output.print_md("✓ {} wandsegmenten aangemaakt".format(len(walls)))
                else:
                    output.print_md("✗ Wanden mislukt")
                if errors:
                    for e in errors:
                        all_errors.append("Wand {}: {}".format(rnumber, e))
    
    output.print_md("")
    output.print_md("## Resultaat")
    if config['do_floor']:
        output.print_md("- **Vloeren:** {}".format(created_floors))
    if config['do_wall']:
        output.print_md("- **Wandsegmenten:** {}".format(created_walls))
    if config['do_ceiling']:
        output.print_md("- **Plafonds:** {}".format(created_ceilings))
    
    if all_errors:
        output.print_md("")
        output.print_md("### Fouten")
        for e in all_errors[:10]:
            output.print_md("- {}".format(e))
    
    forms.alert(
        "Afwerking voltooid!\n\n"
        "Vloeren: {}\n"
        "Wandsegmenten: {}\n"
        "Plafonds: {}".format(created_floors, created_walls, created_ceilings),
        title="Gereed"
    )


if __name__ == "__main__":
    main()
