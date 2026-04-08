# -*- coding: utf-8 -*-
"""Elementen Nummeren - Automatisch nummeren van elementen

Nummert elementen zoals je een boek leest: links naar rechts, boven naar onder.
Plaatst tekstannotatie bij elk element.
"""

__title__ = "Elementen\nNummeren"
__author__ = "3BM Bouwkunde"
__doc__ = "Nummer elementen automatisch (L>R, boven>onder) en plaats tekst"

import sys
import os

sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from bm_logger import get_logger

log = get_logger("PalenNummeren")

from pyrevit import revit, DB, forms, script

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, Button, ComboBox, ComboBoxStyle, TextBox,
    Panel, MessageBox, MessageBoxButtons, MessageBoxIcon,
    FormStartPosition, FormBorderStyle, FlatStyle
)
from System.Drawing import Point, Size, Color, Font, FontStyle, ContentAlignment

# ==============================================================================
# 3BM HUISSTIJL KLEUREN
# ==============================================================================
COLOR_VIOLET = Color.FromArgb(53, 14, 53)       # #350E35 - Magic Violet
COLOR_TEAL = Color.FromArgb(69, 182, 168)       # #44B6A8 - Verdigris
COLOR_YELLOW = Color.FromArgb(239, 189, 117)    # #EFBD75 - Friendly Yellow
COLOR_MAGENTA = Color.FromArgb(160, 28, 72)     # #A01C48 - Warm Magenta
COLOR_PEACH = Color.FromArgb(219, 76, 64)       # #DB4C40 - Flaming Peach

# Tolerantie voor Y-coordinaat groepering (in feet, ~150mm)
Y_TOLERANCE = 0.5

# Categorieen
CATEGORY_OPTIONS = [
    ("Structural Columns (Palen/Kolommen)", DB.BuiltInCategory.OST_StructuralColumns),
    ("Structural Foundations (Funderingen)", DB.BuiltInCategory.OST_StructuralFoundation),
    ("Structural Framing (Liggers)", DB.BuiltInCategory.OST_StructuralFraming),
    ("Columns (Architectonische kolommen)", DB.BuiltInCategory.OST_Columns),
    ("Floors (Vloeren)", DB.BuiltInCategory.OST_Floors),
    ("Walls (Wanden)", DB.BuiltInCategory.OST_Walls),
    ("Doors (Deuren)", DB.BuiltInCategory.OST_Doors),
    ("Windows (Ramen)", DB.BuiltInCategory.OST_Windows),
    ("Generic Models", DB.BuiltInCategory.OST_GenericModel),
    ("Furniture (Meubilair)", DB.BuiltInCategory.OST_Furniture),
    ("Plumbing Fixtures (Sanitair)", DB.BuiltInCategory.OST_PlumbingFixtures),
    ("Lighting Fixtures (Verlichting)", DB.BuiltInCategory.OST_LightingFixtures),
    ("Mechanical Equipment", DB.BuiltInCategory.OST_MechanicalEquipment),
    ("Electrical Equipment", DB.BuiltInCategory.OST_ElectricalEquipment),
    ("Parking (Parkeerplaatsen)", DB.BuiltInCategory.OST_Parking),
]

SCOPE_OPTIONS = [
    "Zichtbaar in huidige view",
    "Geselecteerde elementen",
    "Alle elementen in model",
]


def get_elements_from_selection(category):
    selection = revit.get_selection()
    elements = []
    cat_id = int(category)
    for el in selection:
        if el.Category and el.Category.Id.IntegerValue == cat_id:
            elements.append(el)
    return elements


def get_elements_visible_in_view(view, category):
    collector = DB.FilteredElementCollector(revit.doc, view.Id)\
        .OfCategory(category)\
        .WhereElementIsNotElementType()
    return list(collector)


def get_all_elements(category):
    collector = DB.FilteredElementCollector(revit.doc)\
        .OfCategory(category)\
        .WhereElementIsNotElementType()
    return list(collector)


def get_location(element):
    loc = element.Location
    if isinstance(loc, DB.LocationPoint):
        pt = loc.Point
        return (pt.X, pt.Y, pt.Z)
    elif isinstance(loc, DB.LocationCurve):
        curve = loc.Curve
        pt = curve.Evaluate(0.5, True)
        return (pt.X, pt.Y, pt.Z)
    try:
        bb = element.get_BoundingBox(None)
        if bb:
            center = (bb.Min + bb.Max) / 2
            return (center.X, center.Y, center.Z)
    except:
        pass
    return (0, 0, 0)


def sort_elements_reading_order(elements):
    elements_met_loc = []
    for el in elements:
        x, y, z = get_location(el)
        elements_met_loc.append((el, x, y))
    
    elements_met_loc.sort(key=lambda p: -p[2])
    
    rows = []
    current_row = []
    current_y = None
    
    for el, x, y in elements_met_loc:
        if current_y is None or abs(y - current_y) <= Y_TOLERANCE:
            current_row.append((el, x, y))
            if current_y is None:
                current_y = y
        else:
            if current_row:
                rows.append(current_row)
            current_row = [(el, x, y)]
            current_y = y
    
    if current_row:
        rows.append(current_row)
    
    sorted_elements = []
    for row in rows:
        row.sort(key=lambda p: p[1])
        for el, x, y in row:
            sorted_elements.append(el)
    
    return sorted_elements


def create_text_note(view, location, text, offset_x=1.0, offset_y=1.0):
    text_point = DB.XYZ(location.X + offset_x, location.Y + offset_y, 0)
    text_type_id = revit.doc.GetDefaultElementTypeId(DB.ElementTypeGroup.TextNoteType)
    options = DB.TextNoteOptions()
    options.TypeId = text_type_id
    options.HorizontalAlignment = DB.HorizontalTextAlignment.Left
    text_note = DB.TextNote.Create(revit.doc, view.Id, text_point, text, options)
    return text_note


class NummerenForm(Form):
    def __init__(self):
        self._setup_form()
    
    def _setup_form(self):
        self.Text = "Elementen Nummeren - 3BM Bouwkunde"
        self.Size = Size(480, 360)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.BackColor = Color.White
        
        # Header panel met 3BM kleuren
        self.pnl_header = Panel()
        self.pnl_header.Location = Point(0, 0)
        self.pnl_header.Size = Size(480, 55)
        self.pnl_header.BackColor = COLOR_VIOLET
        self.Controls.Add(self.pnl_header)
        
        self.lbl_title = Label()
        self.lbl_title.Text = "Elementen Nummeren"
        self.lbl_title.Font = Font("Segoe UI", 16, FontStyle.Bold)
        self.lbl_title.ForeColor = Color.White
        self.lbl_title.Location = Point(20, 12)
        self.lbl_title.Size = Size(400, 35)
        self.pnl_header.Controls.Add(self.lbl_title)
        
        # Accent streep
        self.pnl_accent = Panel()
        self.pnl_accent.Location = Point(0, 55)
        self.pnl_accent.Size = Size(480, 5)
        self.pnl_accent.BackColor = COLOR_TEAL
        self.Controls.Add(self.pnl_accent)
        
        y = 75
        
        # Categorie
        self.lbl_cat = Label()
        self.lbl_cat.Text = "Categorie:"
        self.lbl_cat.Location = Point(25, y)
        self.lbl_cat.Size = Size(120, 20)
        self.lbl_cat.Font = Font("Segoe UI", 9)
        self.Controls.Add(self.lbl_cat)
        
        self.cmb_category = ComboBox()
        self.cmb_category.Location = Point(25, y + 22)
        self.cmb_category.Size = Size(415, 28)
        self.cmb_category.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_category.Font = Font("Segoe UI", 9)
        for name, cat in CATEGORY_OPTIONS:
            self.cmb_category.Items.Add(name)
        self.cmb_category.SelectedIndex = 0
        self.Controls.Add(self.cmb_category)
        
        y += 55
        
        # Scope
        self.lbl_scope = Label()
        self.lbl_scope.Text = "Bereik:"
        self.lbl_scope.Location = Point(25, y)
        self.lbl_scope.Size = Size(120, 20)
        self.lbl_scope.Font = Font("Segoe UI", 9)
        self.Controls.Add(self.lbl_scope)
        
        self.cmb_scope = ComboBox()
        self.cmb_scope.Location = Point(25, y + 22)
        self.cmb_scope.Size = Size(415, 28)
        self.cmb_scope.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_scope.Font = Font("Segoe UI", 9)
        for scope in SCOPE_OPTIONS:
            self.cmb_scope.Items.Add(scope)
        self.cmb_scope.SelectedIndex = 0
        self.Controls.Add(self.cmb_scope)
        
        y += 55
        
        # Prefix en Startnummer op zelfde rij
        self.lbl_prefix = Label()
        self.lbl_prefix.Text = "Prefix:"
        self.lbl_prefix.Location = Point(25, y)
        self.lbl_prefix.Size = Size(120, 20)
        self.lbl_prefix.Font = Font("Segoe UI", 9)
        self.Controls.Add(self.lbl_prefix)
        
        self.lbl_start = Label()
        self.lbl_start.Text = "Startnummer:"
        self.lbl_start.Location = Point(245, y)
        self.lbl_start.Size = Size(120, 20)
        self.lbl_start.Font = Font("Segoe UI", 9)
        self.Controls.Add(self.lbl_start)
        
        self.txt_prefix = TextBox()
        self.txt_prefix.Location = Point(25, y + 22)
        self.txt_prefix.Size = Size(195, 28)
        self.txt_prefix.Font = Font("Segoe UI", 10)
        self.txt_prefix.Text = "P"
        self.Controls.Add(self.txt_prefix)
        
        self.txt_start = TextBox()
        self.txt_start.Location = Point(245, y + 22)
        self.txt_start.Size = Size(195, 28)
        self.txt_start.Font = Font("Segoe UI", 10)
        self.txt_start.Text = "1"
        self.Controls.Add(self.txt_start)
        
        y += 70
        
        # Buttons
        self.btn_cancel = Button()
        self.btn_cancel.Text = "Annuleren"
        self.btn_cancel.Location = Point(25, y)
        self.btn_cancel.Size = Size(120, 38)
        self.btn_cancel.Font = Font("Segoe UI", 9)
        self.btn_cancel.FlatStyle = FlatStyle.Flat
        self.btn_cancel.BackColor = Color.White
        self.btn_cancel.ForeColor = COLOR_VIOLET
        self.btn_cancel.Click += self._cancel_click
        self.Controls.Add(self.btn_cancel)
        
        self.btn_run = Button()
        self.btn_run.Text = "Nummeren"
        self.btn_run.Location = Point(300, y)
        self.btn_run.Size = Size(140, 38)
        self.btn_run.Font = Font("Segoe UI", 11, FontStyle.Bold)
        self.btn_run.FlatStyle = FlatStyle.Flat
        self.btn_run.BackColor = COLOR_TEAL
        self.btn_run.ForeColor = Color.White
        self.btn_run.Click += self._run_click
        self.Controls.Add(self.btn_run)
        
        # Result
        self.result = None
    
    def _cancel_click(self, sender, args):
        self.result = None
        self.Close()
    
    def _run_click(self, sender, args):
        try:
            start_num = int(self.txt_start.Text)
        except:
            MessageBox.Show("Ongeldig startnummer.", "Fout", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        
        self.result = {
            'category_index': self.cmb_category.SelectedIndex,
            'scope_index': self.cmb_scope.SelectedIndex,
            'prefix': self.txt_prefix.Text,
            'start_num': start_num
        }
        self.Close()


def main():
    view = revit.doc.ActiveView
    if not isinstance(view, (DB.ViewPlan, DB.ViewSection, DB.ViewDrafting)):
        forms.alert("Open een plattegrond, doorsnede of tekenblad.", title="Elementen Nummeren")
        return
    
    # Toon form
    form = NummerenForm()
    form.ShowDialog()
    
    if not form.result:
        return
    
    # Haal settings op
    category_name, category = CATEGORY_OPTIONS[form.result['category_index']]
    scope = SCOPE_OPTIONS[form.result['scope_index']]
    prefix = form.result['prefix']
    start_num = form.result['start_num']
    
    # Haal elementen op
    if scope == "Geselecteerde elementen":
        elements = get_elements_from_selection(category)
        if not elements:
            forms.alert("Selecteer eerst elementen van de gekozen categorie.", title="Elementen Nummeren")
            return
    elif scope == "Zichtbaar in huidige view":
        elements = get_elements_visible_in_view(view, category)
        if not elements:
            forms.alert("Geen elementen van deze categorie zichtbaar in deze view.", title="Elementen Nummeren")
            return
    else:
        elements = get_all_elements(category)
        if not elements:
            forms.alert("Geen elementen van deze categorie gevonden in het model.", title="Elementen Nummeren")
            return
    
    # Sorteer
    sorted_elements = sort_elements_reading_order(elements)
    
    # Plaats tekst
    with revit.Transaction("Elementen Nummeren"):
        for i, el in enumerate(sorted_elements):
            nummer = start_num + i
            label = "{}{}".format(prefix, str(nummer).zfill(3))
            x, y, z = get_location(el)
            loc_point = DB.XYZ(x, y, 0)
            try:
                create_text_note(view, loc_point, label)
            except Exception as e:
                print("Fout bij element {}: {}".format(nummer, str(e)))
    
    forms.alert(
        "{} elementen genummerd:\n{} t/m {}".format(
            len(sorted_elements),
            "{}{}".format(prefix, str(start_num).zfill(3)),
            "{}{}".format(prefix, str(start_num + len(sorted_elements) - 1).zfill(3))
        ),
        title="Elementen Nummeren"
    )


if __name__ == '__main__':
    main()
