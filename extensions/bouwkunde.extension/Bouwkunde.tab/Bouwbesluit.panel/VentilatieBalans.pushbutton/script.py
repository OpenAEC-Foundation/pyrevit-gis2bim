# -*- coding: utf-8 -*-
"""Ventilatie Balans Calculator - Native pyRevit Tool

Berekent ventilatie-eisen volgens Bouwbesluit/BBL op basis van Revit ruimtes.
Gebruikt ui_template.py helpers voor consistente styling en DPI scaling.
"""

__title__ = "Ventilatie\nBalans"
__author__ = "3BM Bouwkunde"

import os
import sys
import json
import codecs

# Voeg lib folder toe aan path
lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

from pyrevit import revit, DB, forms, script

# UI helpers via template (geen BaseForm - eigen Form layout)
from ui_template import DPIScaler, Huisstijl, UIFactory
from bm_logger import get_logger

log = get_logger("VentilatieBalans")

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, Button, DataGridView, DataGridViewTextBoxColumn,
    DataGridViewComboBoxColumn, ComboBox, ComboBoxStyle, TabControl, 
    TabPage, GroupBox, Panel, MessageBox, MessageBoxButtons, MessageBoxIcon, 
    FormStartPosition, SaveFileDialog, DialogResult, BorderStyle, 
    DataGridViewAutoSizeColumnsMode, DataGridViewSelectionMode, 
    FlatStyle, NumericUpDown, ListBox, AnchorStyles, AutoScaleMode
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# ==============================================================================
# CONSTANTEN
# ==============================================================================
ZONE_PARAMETER = 'ruimtezone'
FUNCTIE_PARAMETER = 'gebruiksfunctie'
PERSONEN_PARAMETERS = ['bezetting', 'PvE_aantal personen', 'Aantal personen', 'Occupancy']

TYPE_TOEVOER = 'toevoer'
TYPE_AFVOER = 'afvoer'
TYPE_GEEN = 'geen'

NORMEN_BBL = {
    'bijeenkomstfunctie': {'dm3_per_m2': 0.9, 'minimum': 7, 'type': TYPE_TOEVOER},
    'kantoorfunctie': {'dm3_per_m2': 0.9, 'minimum': 7, 'type': TYPE_TOEVOER},
    'onderwijsfunctie': {'dm3_per_m2': 0.9, 'minimum': 7, 'type': TYPE_TOEVOER},
    'sportfunctie': {'dm3_per_m2': 0.9, 'minimum': 7, 'type': TYPE_TOEVOER},
    'winkelfunctie': {'dm3_per_m2': 0.9, 'minimum': 7, 'type': TYPE_TOEVOER},
    'woonfunctie': {'dm3_per_m2': 0.7, 'minimum': 7, 'type': TYPE_TOEVOER},
    'gezondheidszorgfunctie': {'dm3_per_m2': 0.9, 'minimum': 7, 'type': TYPE_TOEVOER},
    'logiesfunctie': {'dm3_per_m2': 0.7, 'minimum': 7, 'type': TYPE_TOEVOER},
    'verblijfsgebied': {'dm3_per_m2': 0.9, 'minimum': 7, 'type': TYPE_TOEVOER},
    'verblijfsruimte': {'dm3_per_m2': 0.7, 'minimum': 7, 'type': TYPE_TOEVOER},
    'overige gebruiksfunctie': {'dm3_per_m2': 0.7, 'minimum': 7, 'type': TYPE_AFVOER},
    'toiletruimte': {'dm3_per_m2': 0, 'minimum': 7, 'type': TYPE_AFVOER},
    'badruimte': {'dm3_per_m2': 0, 'minimum': 14, 'type': TYPE_AFVOER},
    'keuken': {'dm3_per_m2': 0, 'minimum': 21, 'type': TYPE_AFVOER},
    'wasruimte': {'dm3_per_m2': 0, 'minimum': 14, 'type': TYPE_AFVOER},
    'technische ruimte': {'dm3_per_m2': 0, 'minimum': 2, 'type': TYPE_AFVOER},
    'meterruimte': {'dm3_per_m2': 0, 'minimum': 2, 'type': TYPE_AFVOER},
    'bergruimte': {'dm3_per_m2': 0, 'minimum': 0, 'type': TYPE_GEEN},
    'berging': {'dm3_per_m2': 0, 'minimum': 0, 'type': TYPE_GEEN},
    'verkeersruimte': {'dm3_per_m2': 0, 'minimum': 0, 'type': TYPE_GEEN},
    'gang': {'dm3_per_m2': 0, 'minimum': 0, 'type': TYPE_GEEN},
    'hal': {'dm3_per_m2': 0, 'minimum': 0, 'type': TYPE_GEEN},
    'trappenhuis': {'dm3_per_m2': 0, 'minimum': 0, 'type': TYPE_GEEN},
}
DEFAULT_NORM = {'dm3_per_m2': 0.7, 'minimum': 7, 'type': TYPE_AFVOER}


# ==============================================================================
# SETTINGS PERSISTENCE
# ==============================================================================
def get_project_settings_path():
    doc_path = revit.doc.PathName
    if doc_path:
        project_dir = os.path.dirname(doc_path)
        project_name = os.path.splitext(os.path.basename(doc_path))[0]
        return os.path.join(project_dir, "{}_ventilatie_settings.json".format(project_name))
    return None


def load_project_settings():
    settings_path = get_project_settings_path()
    if settings_path and os.path.exists(settings_path):
        try:
            with codecs.open(settings_path, 'r', 'utf-8') as f:
                return json.load(f)
        except:
            pass
    return {}


def save_project_settings(settings):
    settings_path = get_project_settings_path()
    if settings_path:
        try:
            with codecs.open(settings_path, 'w', 'utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
            return True
        except:
            pass
    return False


def load_units_database():
    script_dir = os.path.dirname(__file__)
    json_path = os.path.join(script_dir, 'ventilatie_units.json')
    if os.path.exists(json_path):
        try:
            with codecs.open(json_path, 'r', 'utf-8') as f:
                return json.load(f)
        except:
            pass
    return {'wtw_units': [], 'mv_units': []}


# ==============================================================================
# DATA CLASSES
# ==============================================================================
class AirTerminalData(object):
    def __init__(self, element):
        self.element = element
        self.element_id = element.Id.IntegerValue
        self.mark = ''
        self.family_name = ''
        self.system_classification = ''
        self.is_toevoer = False
        self.is_afvoer = False
        self.location = None
        self.ruimte_id = None
        self._load_data()
    
    def _load_data(self):
        mark_param = self.element.LookupParameter('Mark')
        if mark_param and mark_param.HasValue:
            self.mark = mark_param.AsString() or ''
        
        fam_param = self.element.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        if fam_param and fam_param.HasValue:
            self.family_name = fam_param.AsValueString() or ''
        
        sys_param = self.element.LookupParameter('System Classification')
        if sys_param and sys_param.HasValue:
            self.system_classification = sys_param.AsValueString() or ''
        
        fam_lower = self.family_name.lower()
        sys_lower = self.system_classification.lower()
        
        if 'toevoer' in fam_lower or 'supply' in fam_lower or 'supply' in sys_lower:
            self.is_toevoer = True
        elif 'afvoer' in fam_lower or 'exhaust' in fam_lower or 'exhaust' in sys_lower or 'return' in sys_lower:
            self.is_afvoer = True
        
        loc = self.element.Location
        if loc and hasattr(loc, 'Point'):
            pt = loc.Point
            self.location = (pt.X, pt.Y, pt.Z)


class RuimteData(object):
    def __init__(self, element, dm3_per_persoon=0):
        self.element = element
        self.element_id = element.Id.IntegerValue
        self.naam = self._get_param_value('Name') or 'Naamloos'
        self.nummer = self._get_param_value('Number') or '-'
        self.niveau = self._get_level_name()
        self.oppervlakte = self._get_area()
        self.aantal_personen = self._get_personen()
        self.zone = self._get_param_value(ZONE_PARAMETER)
        self.zone_missing = self.zone is None or self.zone.strip() == ''
        if self.zone_missing:
            self.zone = '(geen zone)'
        self.gebruiksfunctie = self._get_param_value(FUNCTIE_PARAMETER) or ''
        self.gebruiksfunctie_missing = self.gebruiksfunctie.strip() == ''
        self.norm = self._get_norm()
        self.dm3_per_persoon = dm3_per_persoon
        self.ventilatie_type = self.norm['type']
        self.ventilatie_type_override = None
        self.ventilatie_eis = self._bereken_ventilatie_eis()
        self.afvoer_correctie = 0.0
        self.air_terminals_toevoer = []
        self.air_terminals_afvoer = []
        self.bbox = self._get_bounding_box()
    
    def _get_bounding_box(self):
        bbox = self.element.get_BoundingBox(None)
        if bbox:
            return {'min_x': bbox.Min.X, 'max_x': bbox.Max.X, 'min_y': bbox.Min.Y, 'max_y': bbox.Max.Y}
        return None
    
    def punt_in_ruimte(self, x, y):
        if not self.bbox:
            return False
        return (self.bbox['min_x'] <= x <= self.bbox['max_x'] and self.bbox['min_y'] <= y <= self.bbox['max_y'])
    
    @property
    def aantal_toevoer_punten(self):
        return len(self.air_terminals_toevoer)
    
    @property
    def aantal_afvoer_punten(self):
        return len(self.air_terminals_afvoer)
    
    @property
    def mv_status(self):
        eff_type = self.get_effective_type()
        if eff_type == TYPE_TOEVOER:
            return 'OK' if self.aantal_toevoer_punten > 0 else 'ONTBREEKT'
        elif eff_type == TYPE_AFVOER:
            return 'OK' if self.aantal_afvoer_punten > 0 else 'ONTBREEKT'
        else:
            if self.aantal_toevoer_punten > 0 or self.aantal_afvoer_punten > 0:
                return 'OVERBODIG'
            return '-'
    
    def set_type_override(self, new_type):
        self.ventilatie_type_override = new_type if new_type in [TYPE_TOEVOER, TYPE_AFVOER, TYPE_GEEN] else None
    
    def get_effective_type(self):
        return self.ventilatie_type_override if self.ventilatie_type_override else self.ventilatie_type
    
    def update_dm3_per_persoon(self, value):
        self.dm3_per_persoon = value
        self.ventilatie_eis = self._bereken_ventilatie_eis()
    
    def _get_param_value(self, param_name):
        param = self.element.LookupParameter(param_name)
        if param and param.HasValue:
            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
        return None
    
    def _get_norm(self):
        functie = self.gebruiksfunctie.lower().strip()
        if functie in NORMEN_BBL:
            return NORMEN_BBL[functie].copy()
        for key, norm in NORMEN_BBL.items():
            if key in functie or functie in key:
                return norm.copy()
        return DEFAULT_NORM.copy()
    
    def _get_level_name(self):
        level_id = self.element.LevelId
        if level_id and level_id != DB.ElementId.InvalidElementId:
            level = revit.doc.GetElement(level_id)
            if level:
                return level.Name
        return 'Onbekend'
    
    def _get_area(self):
        area_param = self.element.get_Parameter(DB.BuiltInParameter.ROOM_AREA)
        if area_param and area_param.HasValue:
            return area_param.AsDouble() * 0.092903
        return 0.0
    
    def _get_personen(self):
        for param_name in PERSONEN_PARAMETERS:
            val = self._get_param_value(param_name)
            if val is not None:
                try:
                    return int(float(val))
                except:
                    pass
        return 0
    
    def _bereken_ventilatie_eis(self):
        eis = 0.0
        if self.norm['dm3_per_m2'] > 0:
            eis = max(eis, self.oppervlakte * self.norm['dm3_per_m2'])
        if self.dm3_per_persoon > 0 and self.aantal_personen > 0:
            eis = max(eis, self.aantal_personen * self.dm3_per_persoon)
        eis = max(eis, self.norm['minimum'])
        return round(eis, 1)
    
    @property
    def toevoer(self):
        return self.ventilatie_eis if self.get_effective_type() == TYPE_TOEVOER else 0.0
    
    @property
    def afvoer(self):
        return self.ventilatie_eis if self.get_effective_type() == TYPE_AFVOER else 0.0
    
    @property
    def afvoer_totaal(self):
        return self.afvoer + self.afvoer_correctie


class ZoneUnitToewijzing(object):
    def __init__(self, zone_naam):
        self.zone_naam = zone_naam
        self.units = []
        self.gekoppeld_aan = None
    
    def voeg_unit_toe(self, unit_data, aantal=1):
        for item in self.units:
            if item['unit']['id'] == unit_data['id']:
                item['aantal'] += aantal
                return
        self.units.append({'unit': unit_data, 'aantal': aantal})
    
    def get_totaal_capaciteit_m3h(self):
        return sum(u['unit']['capaciteit_m3h'] * u['aantal'] for u in self.units)
    
    def get_totaal_capaciteit_dm3s(self):
        return self.get_totaal_capaciteit_m3h() / 3.6


# ==============================================================================
# MAIN FORM
# ==============================================================================
class VentilatieBalansForm(Form):
    """Hoofdformulier voor Ventilatie Balans - met DPI scaling."""
    
    def __init__(self):
        self.ruimtes = []
        self.filtered_ruimtes = []
        self.ruimtes_zonder_zone = []
        self.ruimtes_zonder_functie = []
        self.air_terminals = []
        self.dm3_per_persoon = 4.0
        self._updating_linked_fields = False
        self._updating_koppeling = False
        self.units_db = load_units_database()
        self.zone_toewijzingen = {}
        self.project_settings = load_project_settings()
        self.settings_loaded = bool(self.project_settings)
        self._setup_form()
        self._load_ruimtes()
        self._load_air_terminals()
        self._koppel_terminals_aan_ruimtes()
        self._setup_filters()
        self._init_zone_toewijzingen()
        self._load_settings_into_form()
        self._bereken_overdruk_verdeling()
        self._update_display()
        self._check_missing_data()
        self._show_settings_status()
    
    def _setup_form(self):
        """Setup form met DPI scaling."""
        self.Text = "Ventilatie Balans - 3BM Bouwkunde"
        self.Size = DPIScaler.scale_size(1300, 750)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.White
        self.AutoScaleMode = AutoScaleMode.Dpi
        self.FormClosing += self._form_closing
        
        # Header
        self.pnl_header = Panel()
        self.pnl_header.Location = Point(0, 0)
        self.pnl_header.Size = Size(self.ClientSize.Width, DPIScaler.scale(55))
        self.pnl_header.BackColor = Huisstijl.VIOLET
        self.pnl_header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_header)
        
        self.lbl_header = Label()
        self.lbl_header.Text = "Ventilatie Balans Berekening"
        self.lbl_header.Font = Font("Segoe UI", 16, FontStyle.Bold)
        self.lbl_header.ForeColor = Color.White
        self.lbl_header.Location = DPIScaler.scale_point(20, 12)
        self.lbl_header.AutoSize = True
        self.pnl_header.Controls.Add(self.lbl_header)
        
        self.lbl_status = Label()
        self.lbl_status.Location = DPIScaler.scale_point(450, 18)
        self.lbl_status.Size = DPIScaler.scale_size(800, 20)
        self.lbl_status.ForeColor = Huisstijl.TEAL
        self.lbl_status.Font = Font("Segoe UI", 9)
        self.pnl_header.Controls.Add(self.lbl_status)
        
        # Accent
        accent_top = DPIScaler.scale(55)
        self.pnl_accent = Panel()
        self.pnl_accent.Location = Point(0, accent_top)
        self.pnl_accent.Size = Size(self.ClientSize.Width, DPIScaler.scale(5))
        self.pnl_accent.BackColor = Huisstijl.TEAL
        self.pnl_accent.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_accent)
        
        # Settings panel
        settings_top = DPIScaler.scale(60)
        self.pnl_settings = Panel()
        self.pnl_settings.Location = Point(DPIScaler.scale(10), settings_top)
        self.pnl_settings.Size = Size(self.ClientSize.Width - DPIScaler.scale(20), DPIScaler.scale(40))
        self.pnl_settings.BackColor = Huisstijl.LIGHT_GRAY
        self.pnl_settings.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_settings)
        
        x = DPIScaler.scale(10)
        y = DPIScaler.scale(10)
        
        lbl_p = UIFactory.create_label("Ventilatie per persoon:")
        lbl_p.Location = Point(x, y)
        self.pnl_settings.Controls.Add(lbl_p)
        x += DPIScaler.scale(135)
        
        self.nud_dm3s = NumericUpDown()
        self.nud_dm3s.Location = Point(x, y - DPIScaler.scale(3))
        self.nud_dm3s.Size = DPIScaler.scale_size(60, 25)
        self.nud_dm3s.Font = Font("Segoe UI", 9)
        self.nud_dm3s.DecimalPlaces = 1
        self.nud_dm3s.Minimum = 0
        self.nud_dm3s.Maximum = 100
        self.nud_dm3s.Value = 4
        self.nud_dm3s.Increment = 0.5
        self.nud_dm3s.ValueChanged += self._dm3s_changed
        self.pnl_settings.Controls.Add(self.nud_dm3s)
        x += DPIScaler.scale(65)
        
        lbl_d = UIFactory.create_label("dm³/s  =")
        lbl_d.Location = Point(x, y)
        self.pnl_settings.Controls.Add(lbl_d)
        x += DPIScaler.scale(55)
        
        self.nud_m3h = NumericUpDown()
        self.nud_m3h.Location = Point(x, y - DPIScaler.scale(3))
        self.nud_m3h.Size = DPIScaler.scale_size(60, 25)
        self.nud_m3h.Font = Font("Segoe UI", 9)
        self.nud_m3h.DecimalPlaces = 1
        self.nud_m3h.Minimum = 0
        self.nud_m3h.Maximum = 360
        self.nud_m3h.Value = 14.4
        self.nud_m3h.Increment = 1
        self.nud_m3h.ValueChanged += self._m3h_changed
        self.pnl_settings.Controls.Add(self.nud_m3h)
        x += DPIScaler.scale(65)
        
        lbl_m = UIFactory.create_label("m³/h pp")
        lbl_m.Location = Point(x, y)
        self.pnl_settings.Controls.Add(lbl_m)
        x += DPIScaler.scale(70)
        
        # Button: Vul Air Flow in
        self.btn_fill_airflow = UIFactory.create_button("Vul Air Flow in", 110, 28, 'danger')
        self.btn_fill_airflow.Location = Point(x, y - DPIScaler.scale(5))
        self.btn_fill_airflow.Click += self._fill_airflow_click
        self.pnl_settings.Controls.Add(self.btn_fill_airflow)
        x += DPIScaler.scale(120)
        
        # Button: Opslaan
        self.btn_save = UIFactory.create_button("Opslaan", 80, 28, 'primary')
        self.btn_save.Location = Point(x, y - DPIScaler.scale(5))
        self.btn_save.Click += self._save_settings_click
        self.pnl_settings.Controls.Add(self.btn_save)
        x += DPIScaler.scale(90)
        
        self.lbl_info = UIFactory.create_label("", italic=True, color=Huisstijl.TEXT_SECONDARY)
        self.lbl_info.AutoSize = False
        self.lbl_info.Size = DPIScaler.scale_size(400, 20)
        self.lbl_info.Location = Point(x, y)
        self.pnl_settings.Controls.Add(self.lbl_info)
        
        # Tabs
        tabs_top = DPIScaler.scale(105)
        tabs_height = self.ClientSize.Height - tabs_top - DPIScaler.scale(55)
        
        self.tabs = TabControl()
        self.tabs.Location = Point(DPIScaler.scale(10), tabs_top)
        self.tabs.Size = Size(self.ClientSize.Width - DPIScaler.scale(20), tabs_height)
        self.tabs.Font = Font("Segoe UI", 9)
        self.tabs.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        
        self.tab_ruimtes = TabPage("Ruimtes")
        self.tab_ruimtes.BackColor = Color.White
        self._setup_ruimtes_tab()
        self.tabs.TabPages.Add(self.tab_ruimtes)
        
        self.tab_balans = TabPage("Balans per Zone")
        self.tab_balans.BackColor = Color.White
        self._setup_balans_tab()
        self.tabs.TabPages.Add(self.tab_balans)
        
        self.tab_units = TabPage("WTW/MV Units")
        self.tab_units.BackColor = Color.White
        self._setup_units_tab()
        self.tabs.TabPages.Add(self.tab_units)
        
        self.tab_samenvatting = TabPage("Samenvatting")
        self.tab_samenvatting.BackColor = Color.White
        self._setup_samenvatting_tab()
        self.tabs.TabPages.Add(self.tab_samenvatting)
        
        self.Controls.Add(self.tabs)
        
        # Bottom buttons
        btn_y = self.ClientSize.Height - DPIScaler.scale(45)
        
        self.btn_export_csv = UIFactory.create_button("Export CSV", 100, 38, 'secondary')
        self.btn_export_csv.Location = Point(self.ClientSize.Width - DPIScaler.scale(330), btn_y)
        self.btn_export_csv.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_export_csv.Click += self._export_csv_click
        self.Controls.Add(self.btn_export_csv)
        
        self.btn_export_sheet = UIFactory.create_button("Naar Sheet", 100, 38, 'warning')
        self.btn_export_sheet.Location = Point(self.ClientSize.Width - DPIScaler.scale(220), btn_y)
        self.btn_export_sheet.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_export_sheet.Click += self._export_sheet_click
        self.Controls.Add(self.btn_export_sheet)
        
        self.btn_close = UIFactory.create_button("Sluiten", 100, 38, 'primary')
        self.btn_close.Location = Point(self.ClientSize.Width - DPIScaler.scale(110), btn_y)
        self.btn_close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_close.Click += self._close_click
        self.Controls.Add(self.btn_close)
    
    def _show_settings_status(self):
        if self.settings_loaded:
            n_overrides = len(self.project_settings.get('type_overrides', {}))
            self.lbl_info.Text = "Settings geladen ({} type overrides)".format(n_overrides)
            self.lbl_info.ForeColor = Huisstijl.TEAL
        else:
            self.lbl_info.Text = "Geen opgeslagen settings gevonden"
            self.lbl_info.ForeColor = Huisstijl.TEXT_SECONDARY
    
    def _load_settings_into_form(self):
        if 'dm3_per_persoon' in self.project_settings:
            self._updating_linked_fields = True
            try:
                dm3s = self.project_settings['dm3_per_persoon']
                self.dm3_per_persoon = dm3s
                self.nud_dm3s.Value = min(dm3s, 100)
                self.nud_m3h.Value = min(round(dm3s * 3.6, 1), 360)
                for r in self.ruimtes:
                    r.update_dm3_per_persoon(dm3s)
            finally:
                self._updating_linked_fields = False
        
        if 'type_overrides' in self.project_settings:
            overrides = self.project_settings['type_overrides']
            for r in self.ruimtes:
                key = str(r.element_id)
                if key in overrides:
                    r.set_type_override(overrides[key])
    
    def _save_current_settings(self):
        settings = {
            'dm3_per_persoon': self.dm3_per_persoon,
            'type_overrides': {}
        }
        for r in self.ruimtes:
            if r.ventilatie_type_override:
                settings['type_overrides'][str(r.element_id)] = r.ventilatie_type_override
        return save_project_settings(settings)
    
    def _save_settings_click(self, s, e):
        if self._save_current_settings():
            n_overrides = sum(1 for r in self.ruimtes if r.ventilatie_type_override)
            self.lbl_info.Text = "Settings opgeslagen! ({} type overrides)".format(n_overrides)
            self.lbl_info.ForeColor = Huisstijl.TEAL
            MessageBox.Show("Settings opgeslagen naar:\n{}".format(get_project_settings_path()), 
                          "Opgeslagen", MessageBoxButtons.OK, MessageBoxIcon.Information)
        else:
            self.lbl_info.Text = "Opslaan mislukt!"
            self.lbl_info.ForeColor = Huisstijl.PEACH
            MessageBox.Show("Kon settings niet opslaan.\nControleer of het project is opgeslagen.", 
                          "Fout", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    
    def _form_closing(self, s, e):
        self._save_current_settings()
    
    def _load_air_terminals(self):
        self.air_terminals = []
        collector = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType()
        for elem in collector:
            at = AirTerminalData(elem)
            if at.location and (at.is_toevoer or at.is_afvoer):
                self.air_terminals.append(at)
        n_toev = sum(1 for at in self.air_terminals if at.is_toevoer)
        n_afv = sum(1 for at in self.air_terminals if at.is_afvoer)
        self.lbl_status.Text = "{} ruimtes | {} Air Terminals ({} TV, {} AV)".format(len(self.ruimtes), len(self.air_terminals), n_toev, n_afv)
    
    def _koppel_terminals_aan_ruimtes(self):
        for at in self.air_terminals:
            if not at.location:
                continue
            x, y, z = at.location
            for r in self.ruimtes:
                if r.punt_in_ruimte(x, y):
                    at.ruimte_id = r.element_id
                    if at.is_toevoer:
                        r.air_terminals_toevoer.append(at)
                    elif at.is_afvoer:
                        r.air_terminals_afvoer.append(at)
                    break
    
    def _init_zone_toewijzingen(self):
        for zone in set(r.zone for r in self.ruimtes):
            if zone not in self.zone_toewijzingen:
                self.zone_toewijzingen[zone] = ZoneUnitToewijzing(zone)
    
    def _get_hoofdzone(self, zone_naam):
        if zone_naam not in self.zone_toewijzingen:
            return zone_naam
        tw = self.zone_toewijzingen[zone_naam]
        if tw.gekoppeld_aan and tw.gekoppeld_aan in self.zone_toewijzingen:
            return tw.gekoppeld_aan
        return zone_naam
    
    def _get_gekoppelde_zones(self, hoofdzone):
        zones = [hoofdzone]
        for zn, tw in self.zone_toewijzingen.items():
            if tw.gekoppeld_aan == hoofdzone and zn != hoofdzone:
                zones.append(zn)
        return zones
    
    def _get_gecombineerde_eis(self, hoofdzone):
        zones_data = self._get_zones_data()
        gekoppelde = self._get_gekoppelde_zones(hoofdzone)
        totaal_eis = 0
        for zn in gekoppelde:
            if zn in zones_data:
                zd = zones_data[zn]
                totaal_eis += max(zd['toevoer'], zd['afvoer_totaal'])
        return totaal_eis
    
    def _bereken_overdruk_verdeling(self):
        for r in self.ruimtes:
            r.afvoer_correctie = 0.0
        zones = {}
        for r in self.ruimtes:
            if r.zone not in zones:
                zones[r.zone] = []
            zones[r.zone].append(r)
        for zone_naam, zone_ruimtes in zones.items():
            totaal_toevoer = sum(r.toevoer for r in zone_ruimtes)
            totaal_afvoer = sum(r.afvoer for r in zone_ruimtes)
            overdruk = totaal_toevoer - totaal_afvoer
            if overdruk > 0:
                afvoer_ruimtes = [r for r in zone_ruimtes if r.get_effective_type() == TYPE_AFVOER]
                if afvoer_ruimtes:
                    totaal_opp = sum(r.oppervlakte for r in afvoer_ruimtes)
                    if totaal_opp > 0:
                        for r in afvoer_ruimtes:
                            r.afvoer_correctie = round(overdruk * r.oppervlakte / totaal_opp, 1)
    
    def _setup_ruimtes_tab(self):
        y = DPIScaler.scale(12)
        x = DPIScaler.scale(10)
        
        lbl_z = UIFactory.create_label("Zone:")
        lbl_z.Location = Point(x, y + DPIScaler.scale(3))
        self.tab_ruimtes.Controls.Add(lbl_z)
        x += DPIScaler.scale(40)
        
        self.cmb_zone = UIFactory.create_combobox(180)
        self.cmb_zone.Location = Point(x, y)
        self.cmb_zone.SelectedIndexChanged += self._filter_changed
        self.tab_ruimtes.Controls.Add(self.cmb_zone)
        x += DPIScaler.scale(200)
        
        lbl_n = UIFactory.create_label("Niveau:")
        lbl_n.Location = Point(x, y + DPIScaler.scale(3))
        self.tab_ruimtes.Controls.Add(lbl_n)
        x += DPIScaler.scale(50)
        
        self.cmb_niveau = UIFactory.create_combobox(180)
        self.cmb_niveau.Location = Point(x, y)
        self.cmb_niveau.SelectedIndexChanged += self._filter_changed
        self.tab_ruimtes.Controls.Add(self.cmb_niveau)
        x += DPIScaler.scale(200)
        
        self.lbl_warning = UIFactory.create_label("", bold=True, color=Huisstijl.PEACH)
        self.lbl_warning.AutoSize = False
        self.lbl_warning.Size = DPIScaler.scale_size(500, 20)
        self.lbl_warning.Location = Point(x, y + DPIScaler.scale(3))
        self.lbl_warning.Visible = False
        self.tab_ruimtes.Controls.Add(self.lbl_warning)
        
        # Grid
        grid_top = DPIScaler.scale(45)
        self.grid = DataGridView()
        self.grid.Location = Point(DPIScaler.scale(10), grid_top)
        self.grid.Size = Size(self.tabs.Width - DPIScaler.scale(35), self.tabs.Height - grid_top - DPIScaler.scale(35))
        self.grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid.AllowUserToAddRows = False
        self.grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.grid.BackgroundColor = Color.White
        self.grid.BorderStyle = BorderStyle.None
        self.grid.RowHeadersVisible = False
        self.grid.EnableHeadersVisualStyles = False
        self.grid.ColumnHeadersDefaultCellStyle.BackColor = Huisstijl.VIOLET
        self.grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        self.grid.ColumnHeadersDefaultCellStyle.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.grid.ColumnHeadersHeight = DPIScaler.scale(30)
        self.grid.RowTemplate.Height = DPIScaler.scale(26)
        self.grid.AlternatingRowsDefaultCellStyle.BackColor = Huisstijl.LIGHT_GRAY
        
        for name, header, width in [("zone", "Zone", 90), ("naam", "Naam", 120), ("niveau", "Niveau", 80), ("functie", "Functie", 110), ("opp", "m²", 45), ("pers", "Pers.", 40), ("eis", "Eis", 50)]:
            col = DataGridViewTextBoxColumn()
            col.Name = name
            col.HeaderText = header
            col.Width = DPIScaler.scale(width)
            col.ReadOnly = True
            self.grid.Columns.Add(col)
        
        type_col = DataGridViewComboBoxColumn()
        type_col.Name = "type"
        type_col.HeaderText = "Type"
        type_col.Width = DPIScaler.scale(70)
        type_col.Items.Add("toevoer")
        type_col.Items.Add("afvoer")
        type_col.Items.Add("geen")
        type_col.FlatStyle = FlatStyle.Flat
        self.grid.Columns.Add(type_col)
        
        for name, header, width in [("toev", "Toev.", 50), ("afv", "Afv.", 50), ("corr", "Corr.", 50), ("tot", "Tot.", 50), ("at_t", "TV", 35), ("at_a", "AV", 35), ("at_st", "Status", 70)]:
            col = DataGridViewTextBoxColumn()
            col.Name = name
            col.HeaderText = header
            col.Width = DPIScaler.scale(width)
            col.ReadOnly = True
            self.grid.Columns.Add(col)
        
        self.grid.CellValueChanged += self._grid_changed
        self.grid.CurrentCellDirtyStateChanged += self._grid_dirty
        self.tab_ruimtes.Controls.Add(self.grid)
    
    def _grid_dirty(self, s, e):
        if self.grid.IsCurrentCellDirty:
            self.grid.CommitEdit(1)
    
    def _grid_changed(self, s, e):
        if e.ColumnIndex == self.grid.Columns["type"].Index:
            idx = e.RowIndex
            if 0 <= idx < len(self.filtered_ruimtes):
                new_type = self.grid.Rows[idx].Cells["type"].Value
                r = self.filtered_ruimtes[idx]
                r.set_type_override(None if new_type == r.ventilatie_type else str(new_type))
                self._bereken_overdruk_verdeling()
                self._update_display()
    
    def _setup_balans_tab(self):
        self.pnl_balans = Panel()
        self.pnl_balans.Location = Point(DPIScaler.scale(10), DPIScaler.scale(10))
        self.pnl_balans.Size = Size(self.tabs.Width - DPIScaler.scale(35), self.tabs.Height - DPIScaler.scale(45))
        self.pnl_balans.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.pnl_balans.AutoScroll = True
        self.pnl_balans.BackColor = Color.White
        self.tab_balans.Controls.Add(self.pnl_balans)
    
    def _setup_samenvatting_tab(self):
        self.lbl_samenvatting = Label()
        self.lbl_samenvatting.Location = Point(DPIScaler.scale(10), DPIScaler.scale(10))
        self.lbl_samenvatting.Size = Size(self.tabs.Width - DPIScaler.scale(35), self.tabs.Height - DPIScaler.scale(45))
        self.lbl_samenvatting.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lbl_samenvatting.Font = Font("Consolas", 10)
        self.tab_samenvatting.Controls.Add(self.lbl_samenvatting)
    
    def _setup_units_tab(self):
        # Linker panel
        pnl_left = Panel()
        pnl_left.Location = Point(DPIScaler.scale(10), DPIScaler.scale(10))
        pnl_left.Size = DPIScaler.scale_size(400, 480)
        pnl_left.BorderStyle = BorderStyle.FixedSingle
        self.tab_units.Controls.Add(pnl_left)
        
        y = DPIScaler.scale(10)
        
        lbl_z = UIFactory.create_label("Zone:", bold=True)
        lbl_z.Location = Point(DPIScaler.scale(10), y)
        pnl_left.Controls.Add(lbl_z)
        
        self.cmb_unit_zone = UIFactory.create_combobox(320)
        self.cmb_unit_zone.Location = Point(DPIScaler.scale(65), y - DPIScaler.scale(3))
        self.cmb_unit_zone.SelectedIndexChanged += self._unit_zone_changed
        pnl_left.Controls.Add(self.cmb_unit_zone)
        y += DPIScaler.scale(30)
        
        self.lbl_zone_info = UIFactory.create_label("", color=Huisstijl.TEXT_SECONDARY)
        self.lbl_zone_info.AutoSize = False
        self.lbl_zone_info.Size = DPIScaler.scale_size(375, 20)
        self.lbl_zone_info.Location = Point(DPIScaler.scale(10), y)
        pnl_left.Controls.Add(self.lbl_zone_info)
        y += DPIScaler.scale(28)
        
        lbl_koppel = UIFactory.create_label("Koppel aan zone:", bold=True, color=Huisstijl.MAGENTA)
        lbl_koppel.Location = Point(DPIScaler.scale(10), y)
        pnl_left.Controls.Add(lbl_koppel)
        
        self.cmb_koppel_zone = UIFactory.create_combobox(270)
        self.cmb_koppel_zone.Location = Point(DPIScaler.scale(115), y - DPIScaler.scale(3))
        self.cmb_koppel_zone.SelectedIndexChanged += self._koppel_zone_changed
        pnl_left.Controls.Add(self.cmb_koppel_zone)
        y += DPIScaler.scale(28)
        
        self.lbl_koppel_info = UIFactory.create_label("", italic=True, color=Huisstijl.MAGENTA)
        self.lbl_koppel_info.AutoSize = False
        self.lbl_koppel_info.Size = DPIScaler.scale_size(375, 20)
        self.lbl_koppel_info.Location = Point(DPIScaler.scale(10), y)
        pnl_left.Controls.Add(self.lbl_koppel_info)
        y += DPIScaler.scale(30)
        
        lbl_t = UIFactory.create_label("Toegewezen units:", bold=True)
        lbl_t.Location = Point(DPIScaler.scale(10), y)
        pnl_left.Controls.Add(lbl_t)
        y += DPIScaler.scale(25)
        
        self.lst_toegewezen = ListBox()
        self.lst_toegewezen.Location = Point(DPIScaler.scale(10), y)
        self.lst_toegewezen.Size = DPIScaler.scale_size(375, 250)
        self.lst_toegewezen.Font = Font("Segoe UI", 9)
        pnl_left.Controls.Add(self.lst_toegewezen)
        y += DPIScaler.scale(260)
        
        self.btn_verwijder = UIFactory.create_button("Verwijder unit", 150, 30, 'danger')
        self.btn_verwijder.Location = Point(DPIScaler.scale(10), y)
        self.btn_verwijder.Click += self._verwijder_unit
        pnl_left.Controls.Add(self.btn_verwijder)
        y += DPIScaler.scale(40)
        
        self.lbl_cap_status = UIFactory.create_label("", bold=True)
        self.lbl_cap_status.AutoSize = False
        self.lbl_cap_status.Size = DPIScaler.scale_size(375, 35)
        self.lbl_cap_status.Location = Point(DPIScaler.scale(10), y)
        pnl_left.Controls.Add(self.lbl_cap_status)
        
        # Rechter panel
        pnl_right = Panel()
        pnl_right.Location = Point(DPIScaler.scale(420), DPIScaler.scale(10))
        pnl_right.Size = DPIScaler.scale_size(830, 480)
        pnl_right.BorderStyle = BorderStyle.FixedSingle
        pnl_right.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.tab_units.Controls.Add(pnl_right)
        
        y = DPIScaler.scale(10)
        
        lbl_type = UIFactory.create_label("Type:", bold=True)
        lbl_type.Location = Point(DPIScaler.scale(10), y + DPIScaler.scale(3))
        pnl_right.Controls.Add(lbl_type)
        
        self.cmb_unit_type = UIFactory.create_combobox(120, ["WTW units", "MV units"])
        self.cmb_unit_type.Location = Point(DPIScaler.scale(55), y)
        self.cmb_unit_type.SelectedIndexChanged += self._unit_type_changed
        pnl_right.Controls.Add(self.cmb_unit_type)
        
        lbl_fab = UIFactory.create_label("Fabrikant:")
        lbl_fab.Location = Point(DPIScaler.scale(200), y + DPIScaler.scale(3))
        pnl_right.Controls.Add(lbl_fab)
        
        self.cmb_fabrikant = UIFactory.create_combobox(150)
        self.cmb_fabrikant.Location = Point(DPIScaler.scale(270), y)
        self.cmb_fabrikant.SelectedIndexChanged += self._fab_changed
        pnl_right.Controls.Add(self.cmb_fabrikant)
        y += DPIScaler.scale(35)
        
        self.grid_units = DataGridView()
        self.grid_units.Location = Point(DPIScaler.scale(10), y)
        self.grid_units.Size = DPIScaler.scale_size(805, 350)
        self.grid_units.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.grid_units.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid_units.AllowUserToAddRows = False
        self.grid_units.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.grid_units.BackgroundColor = Color.White
        self.grid_units.BorderStyle = BorderStyle.None
        self.grid_units.ColumnHeadersDefaultCellStyle.BackColor = Huisstijl.VIOLET
        self.grid_units.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        self.grid_units.ColumnHeadersDefaultCellStyle.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.grid_units.EnableHeadersVisualStyles = False
        self.grid_units.RowHeadersVisible = False
        self.grid_units.ReadOnly = True
        self.grid_units.MultiSelect = False
        self.grid_units.ColumnHeadersHeight = DPIScaler.scale(30)
        self.grid_units.RowTemplate.Height = DPIScaler.scale(26)
        
        for name, header, width in [("fab", "Fabrikant", 90), ("model", "Model", 120), ("cap", "Capaciteit", 100), ("rend", "Rend.", 60), ("geluid", "Geluid", 60), ("toep", "Toepassing", 90), ("opm", "Opmerkingen", 180)]:
            col = DataGridViewTextBoxColumn()
            col.Name = name
            col.HeaderText = header
            col.Width = DPIScaler.scale(width)
            self.grid_units.Columns.Add(col)
        pnl_right.Controls.Add(self.grid_units)
        y += DPIScaler.scale(360)
        
        lbl_a = UIFactory.create_label("Aantal:")
        lbl_a.Location = Point(DPIScaler.scale(10), y + DPIScaler.scale(5))
        pnl_right.Controls.Add(lbl_a)
        
        self.nud_aantal = NumericUpDown()
        self.nud_aantal.Location = Point(DPIScaler.scale(65), y)
        self.nud_aantal.Size = DPIScaler.scale_size(50, 25)
        self.nud_aantal.Font = Font("Segoe UI", 9)
        self.nud_aantal.Minimum = 1
        self.nud_aantal.Maximum = 10
        self.nud_aantal.Value = 1
        pnl_right.Controls.Add(self.nud_aantal)
        
        self.btn_voeg_toe = UIFactory.create_button("Voeg toe aan zone", 150, 30, 'primary')
        self.btn_voeg_toe.Location = Point(DPIScaler.scale(130), y)
        self.btn_voeg_toe.Click += self._voeg_unit_toe
        pnl_right.Controls.Add(self.btn_voeg_toe)
        
        self._populate_fabrikanten()
        self._populate_units_grid()
    
    def _populate_fabrikanten(self):
        self.cmb_fabrikant.Items.Clear()
        self.cmb_fabrikant.Items.Add("Alle")
        unit_type = "wtw_units" if self.cmb_unit_type.SelectedIndex == 0 else "mv_units"
        for fab in sorted(set(u.get('fabrikant', '') for u in self.units_db.get(unit_type, []))):
            if fab:
                self.cmb_fabrikant.Items.Add(fab)
        self.cmb_fabrikant.SelectedIndex = 0
    
    def _populate_units_grid(self):
        self.grid_units.Rows.Clear()
        unit_type = "wtw_units" if self.cmb_unit_type.SelectedIndex == 0 else "mv_units"
        units = self.units_db.get(unit_type, [])
        fab_filter = str(self.cmb_fabrikant.SelectedItem) if self.cmb_fabrikant.SelectedIndex > 0 else None
        for u in units:
            if fab_filter and u.get('fabrikant') != fab_filter:
                continue
            self.grid_units.Rows.Add(u.get('fabrikant', ''), u.get('model', ''), "{} m³/h".format(u.get('capaciteit_m3h', 0)), "{}%".format(u.get('rendement_pct', '-')) if 'rendement_pct' in u else '-', "{} dB".format(u.get('geluid_dba', '-')) if 'geluid_dba' in u else '-', u.get('toepassing', ''), u.get('opmerkingen', ''))
    
    def _unit_type_changed(self, s, e):
        self._populate_fabrikanten()
        self._populate_units_grid()
    
    def _fab_changed(self, s, e):
        self._populate_units_grid()
    
    def _unit_zone_changed(self, s, e):
        self._update_koppel_dropdown()
        self._update_zone_units_display()
    
    def _update_koppel_dropdown(self):
        if self._updating_koppeling or self.cmb_unit_zone.SelectedIndex < 0:
            return
        self._updating_koppeling = True
        try:
            zone = str(self.cmb_unit_zone.SelectedItem)
            self.cmb_koppel_zone.Items.Clear()
            self.cmb_koppel_zone.Items.Add("(geen)")
            for zn in sorted(self.zone_toewijzingen.keys()):
                if zn != zone:
                    self.cmb_koppel_zone.Items.Add(zn)
            if zone in self.zone_toewijzingen:
                tw = self.zone_toewijzingen[zone]
                if tw.gekoppeld_aan:
                    for i in range(self.cmb_koppel_zone.Items.Count):
                        if str(self.cmb_koppel_zone.Items[i]) == tw.gekoppeld_aan:
                            self.cmb_koppel_zone.SelectedIndex = i
                            break
                else:
                    self.cmb_koppel_zone.SelectedIndex = 0
            else:
                self.cmb_koppel_zone.SelectedIndex = 0
        finally:
            self._updating_koppeling = False
    
    def _koppel_zone_changed(self, s, e):
        if self._updating_koppeling or self.cmb_unit_zone.SelectedIndex < 0:
            return
        zone = str(self.cmb_unit_zone.SelectedItem)
        if zone not in self.zone_toewijzingen:
            return
        tw = self.zone_toewijzingen[zone]
        tw.gekoppeld_aan = None if self.cmb_koppel_zone.SelectedIndex == 0 else str(self.cmb_koppel_zone.SelectedItem)
        self._update_zone_units_display()
        self._update_balans()
        self._update_samenvatting()
    
    def _update_zone_units_display(self):
        if self.cmb_unit_zone.SelectedIndex < 0:
            return
        zone = str(self.cmb_unit_zone.SelectedItem)
        zones_data = self._get_zones_data()
        if zone in zones_data:
            zd = zones_data[zone]
            self.lbl_zone_info.Text = "Eis: {:.1f} dm³/s toevoer | {:.1f} dm³/s afvoer".format(zd['toevoer'], zd['afvoer_totaal'])
        tw = self.zone_toewijzingen.get(zone)
        is_gekoppeld = tw and tw.gekoppeld_aan
        hoofdzone = self._get_hoofdzone(zone)
        self.lst_toegewezen.Items.Clear()
        if is_gekoppeld:
            self.lbl_koppel_info.Text = "→ Gebruikt units van '{}'".format(tw.gekoppeld_aan)
            self.btn_voeg_toe.Enabled = False
            self.btn_verwijder.Enabled = False
            if hoofdzone in self.zone_toewijzingen:
                for item in self.zone_toewijzingen[hoofdzone].units:
                    u = item['unit']
                    self.lst_toegewezen.Items.Add("(via {}) {}x {} {}".format(hoofdzone, item['aantal'], u.get('fabrikant', ''), u.get('model', '')))
        else:
            self.lbl_koppel_info.Text = ""
            self.btn_voeg_toe.Enabled = True
            self.btn_verwijder.Enabled = True
            if tw:
                for item in tw.units:
                    u = item['unit']
                    self.lst_toegewezen.Items.Add("{}x {} {} - {} m³/h".format(item['aantal'], u.get('fabrikant', ''), u.get('model', ''), u.get('capaciteit_m3h', 0) * item['aantal']))
        if hoofdzone in self.zone_toewijzingen:
            hoofdtw = self.zone_toewijzingen[hoofdzone]
            totaal_cap = hoofdtw.get_totaal_capaciteit_dm3s()
            gecomb_eis = self._get_gecombineerde_eis(hoofdzone)
            if totaal_cap >= gecomb_eis:
                self.lbl_cap_status.Text = "OK: {:.1f} / {:.1f} dm³/s".format(totaal_cap, gecomb_eis)
                self.lbl_cap_status.ForeColor = Huisstijl.TEAL
            elif totaal_cap > 0:
                self.lbl_cap_status.Text = "TEKORT: {:.1f} dm³/s nodig".format(gecomb_eis - totaal_cap)
                self.lbl_cap_status.ForeColor = Huisstijl.PEACH
            else:
                self.lbl_cap_status.Text = "Geen units toegewezen"
                self.lbl_cap_status.ForeColor = Huisstijl.TEXT_SECONDARY
    
    def _get_zones_data(self):
        zones = {}
        for r in self.ruimtes:
            if r.zone not in zones:
                zones[r.zone] = {'toevoer': 0, 'afvoer': 0, 'afvoer_totaal': 0}
            zones[r.zone]['toevoer'] += r.toevoer
            zones[r.zone]['afvoer'] += r.afvoer
            zones[r.zone]['afvoer_totaal'] += r.afvoer_totaal
        return zones
    
    def _voeg_unit_toe(self, s, e):
        if self.cmb_unit_zone.SelectedIndex < 0 or self.grid_units.SelectedRows.Count == 0:
            return
        zone = str(self.cmb_unit_zone.SelectedItem)
        if zone in self.zone_toewijzingen and self.zone_toewijzingen[zone].gekoppeld_aan:
            return
        row_idx = self.grid_units.SelectedRows[0].Index
        aantal = int(self.nud_aantal.Value)
        unit_type = "wtw_units" if self.cmb_unit_type.SelectedIndex == 0 else "mv_units"
        units = self.units_db.get(unit_type, [])
        fab_filter = str(self.cmb_fabrikant.SelectedItem) if self.cmb_fabrikant.SelectedIndex > 0 else None
        filtered = [u for u in units if not fab_filter or u.get('fabrikant') == fab_filter]
        if row_idx < len(filtered):
            if zone not in self.zone_toewijzingen:
                self.zone_toewijzingen[zone] = ZoneUnitToewijzing(zone)
            self.zone_toewijzingen[zone].voeg_unit_toe(filtered[row_idx], aantal)
            self._update_zone_units_display()
            self._update_balans()
            self._update_samenvatting()
    
    def _verwijder_unit(self, s, e):
        if self.cmb_unit_zone.SelectedIndex < 0 or self.lst_toegewezen.SelectedIndex < 0:
            return
        zone = str(self.cmb_unit_zone.SelectedItem)
        if zone in self.zone_toewijzingen and self.zone_toewijzingen[zone].gekoppeld_aan:
            return
        idx = self.lst_toegewezen.SelectedIndex
        if zone in self.zone_toewijzingen and idx < len(self.zone_toewijzingen[zone].units):
            del self.zone_toewijzingen[zone].units[idx]
            self._update_zone_units_display()
            self._update_balans()
            self._update_samenvatting()
    
    def _dm3s_changed(self, s, e):
        if self._updating_linked_fields:
            return
        self._updating_linked_fields = True
        try:
            dm3s = float(self.nud_dm3s.Value)
            self.nud_m3h.Value = min(round(dm3s * 3.6, 1), 360)
            self.dm3_per_persoon = dm3s
            for r in self.ruimtes:
                r.update_dm3_per_persoon(dm3s)
            self._bereken_overdruk_verdeling()
            self._update_display()
        finally:
            self._updating_linked_fields = False
    
    def _m3h_changed(self, s, e):
        if self._updating_linked_fields:
            return
        self._updating_linked_fields = True
        try:
            m3h = float(self.nud_m3h.Value)
            dm3s = round(m3h / 3.6, 1)
            self.nud_dm3s.Value = min(dm3s, 100)
            self.dm3_per_persoon = dm3s
            for r in self.ruimtes:
                r.update_dm3_per_persoon(dm3s)
            self._bereken_overdruk_verdeling()
            self._update_display()
        finally:
            self._updating_linked_fields = False
    
    def _load_ruimtes(self):
        collector = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
        for room in collector:
            if room.Area > 0:
                r = RuimteData(room, self.dm3_per_persoon)
                self.ruimtes.append(r)
                if r.zone_missing:
                    self.ruimtes_zonder_zone.append(r)
                if r.gebruiksfunctie_missing:
                    self.ruimtes_zonder_functie.append(r)
        self.ruimtes.sort(key=lambda r: (r.zone, r.naam))
        self.filtered_ruimtes = list(self.ruimtes)
        self.lbl_status.Text = "{} ruimtes geladen".format(len(self.ruimtes))
    
    def _check_missing_data(self):
        w = []
        if self.ruimtes_zonder_zone:
            w.append("{} zonder zone".format(len(self.ruimtes_zonder_zone)))
        if self.ruimtes_zonder_functie:
            w.append("{} zonder functie".format(len(self.ruimtes_zonder_functie)))
        n_ontbreekt = sum(1 for r in self.ruimtes if r.mv_status == 'ONTBREEKT')
        if n_ontbreekt > 0:
            w.append("{} zonder Air Terminal".format(n_ontbreekt))
        if w:
            self.lbl_warning.Text = "! " + ", ".join(w)
            self.lbl_warning.Visible = True
    
    def _setup_filters(self):
        zones = sorted(set(r.zone for r in self.ruimtes))
        self.cmb_zone.Items.Add("Alle zones")
        for z in zones:
            self.cmb_zone.Items.Add(z)
        self.cmb_zone.SelectedIndex = 0
        self.cmb_unit_zone.Items.Clear()
        for z in zones:
            self.cmb_unit_zone.Items.Add(z)
        if zones:
            self.cmb_unit_zone.SelectedIndex = 0
        niveaus = sorted(set(r.niveau for r in self.ruimtes))
        self.cmb_niveau.Items.Add("Alle niveaus")
        for n in niveaus:
            self.cmb_niveau.Items.Add(n)
        self.cmb_niveau.SelectedIndex = 0
    
    def _filter_changed(self, s, e):
        self._apply_filters()
        self._update_display()
    
    def _apply_filters(self):
        self.filtered_ruimtes = list(self.ruimtes)
        z = self.cmb_zone.SelectedItem
        if z and str(z) != "Alle zones":
            self.filtered_ruimtes = [r for r in self.filtered_ruimtes if r.zone == str(z)]
        n = self.cmb_niveau.SelectedItem
        if n and str(n) != "Alle niveaus":
            self.filtered_ruimtes = [r for r in self.filtered_ruimtes if r.niveau == str(n)]
        self.filtered_ruimtes.sort(key=lambda r: (r.zone, r.naam))
    
    def _update_display(self):
        self.grid.Rows.Clear()
        for r in self.filtered_ruimtes:
            at_t = str(r.aantal_toevoer_punten) if r.aantal_toevoer_punten > 0 else "-"
            at_a = str(r.aantal_afvoer_punten) if r.aantal_afvoer_punten > 0 else "-"
            idx = self.grid.Rows.Add(r.zone, r.naam, r.niveau, r.gebruiksfunctie or "(geen)", "{:.1f}".format(r.oppervlakte), r.aantal_personen if r.aantal_personen > 0 else "-", "{:.1f}".format(r.ventilatie_eis), r.get_effective_type(), "{:.1f}".format(r.toevoer) if r.toevoer > 0 else "-", "{:.1f}".format(r.afvoer) if r.afvoer > 0 else "-", "+{:.1f}".format(r.afvoer_correctie) if r.afvoer_correctie > 0 else "-", "{:.1f}".format(r.afvoer_totaal) if r.afvoer_totaal > 0 else ("{:.1f}".format(r.toevoer) if r.toevoer > 0 else "-"), at_t, at_a, r.mv_status)
            if r.zone_missing or r.gebruiksfunctie_missing:
                self.grid.Rows[idx].DefaultCellStyle.BackColor = Color.FromArgb(255, 230, 230)
            if r.ventilatie_type_override:
                self.grid.Rows[idx].Cells["type"].Style.BackColor = Color.FromArgb(230, 255, 230)
            if r.afvoer_correctie > 0:
                self.grid.Rows[idx].Cells["corr"].Style.BackColor = Color.FromArgb(230, 245, 255)
                self.grid.Rows[idx].Cells["tot"].Style.BackColor = Color.FromArgb(230, 245, 255)
            if r.mv_status == 'ONTBREEKT':
                self.grid.Rows[idx].Cells["at_st"].Style.BackColor = Huisstijl.PEACH
                self.grid.Rows[idx].Cells["at_st"].Style.ForeColor = Color.White
            elif r.mv_status == 'OK':
                self.grid.Rows[idx].Cells["at_st"].Style.BackColor = Huisstijl.TEAL
                self.grid.Rows[idx].Cells["at_st"].Style.ForeColor = Color.White
            elif r.mv_status == 'OVERBODIG':
                self.grid.Rows[idx].Cells["at_st"].Style.BackColor = Huisstijl.YELLOW
        self._update_balans()
        self._update_samenvatting()
        self._update_zone_units_display()
    
    def _update_balans(self):
        self.pnl_balans.Controls.Clear()
        zones = {}
        for r in self.ruimtes:
            if r.zone not in zones:
                zones[r.zone] = {'toevoer': 0, 'afvoer': 0, 'afvoer_totaal': 0, 'missing': False, 'at_t': 0, 'at_a': 0}
            zones[r.zone]['toevoer'] += r.toevoer
            zones[r.zone]['afvoer'] += r.afvoer
            zones[r.zone]['afvoer_totaal'] += r.afvoer_totaal
            zones[r.zone]['at_t'] += r.aantal_toevoer_punten
            zones[r.zone]['at_a'] += r.aantal_afvoer_punten
            if r.zone_missing or r.gebruiksfunctie_missing:
                zones[r.zone]['missing'] = True
        y = DPIScaler.scale(5)
        for zn, zd in sorted(zones.items()):
            gb = GroupBox()
            gb.Text = "  {}  ".format(zn)
            gb.Location = Point(DPIScaler.scale(5), y)
            gb.Size = DPIScaler.scale_size(1200, 100)
            gb.Font = Font("Segoe UI", 10, FontStyle.Bold)
            gb.ForeColor = Huisstijl.PEACH if zd['missing'] else Huisstijl.VIOLET
            
            lbl1 = UIFactory.create_label("Eis: {:.1f} dm³/s toevoer | {:.1f} dm³/s afvoer | Air Terminals: {} TV, {} AV".format(zd['toevoer'], zd['afvoer'], zd['at_t'], zd['at_a']))
            lbl1.Location = Point(DPIScaler.scale(15), DPIScaler.scale(25))
            gb.Controls.Add(lbl1)
            
            balans = zd['toevoer'] - zd['afvoer']
            if balans > 0:
                lbl2 = UIFactory.create_label("Ontwerp: {:.1f} dm³/s afvoer [+{:.1f} verdeeld]".format(zd['afvoer_totaal'], balans), bold=True, color=Huisstijl.TEAL)
                lbl2.Location = Point(DPIScaler.scale(15), DPIScaler.scale(45))
                gb.Controls.Add(lbl2)
            
            balans_gecorr = zd['toevoer'] - zd['afvoer_totaal']
            if abs(balans_gecorr) < 1:
                status, color = "In balans", Huisstijl.TEAL
            elif balans_gecorr > 0:
                status, color = "Overdruk +{:.1f}".format(balans_gecorr), Huisstijl.YELLOW
            else:
                status, color = "Onderdruk {:.1f}".format(balans_gecorr), Huisstijl.PEACH
            lbl_s = UIFactory.create_label(status, bold=True, color=color)
            lbl_s.Location = Point(DPIScaler.scale(15), DPIScaler.scale(70))
            gb.Controls.Add(lbl_s)
            self.pnl_balans.Controls.Add(gb)
            y += DPIScaler.scale(110)
    
    def _update_samenvatting(self):
        tot_t = sum(r.toevoer for r in self.ruimtes)
        tot_a = sum(r.afvoer for r in self.ruimtes)
        tot_at = sum(r.afvoer_totaal for r in self.ruimtes)
        tot_at_t = sum(r.aantal_toevoer_punten for r in self.ruimtes)
        tot_at_a = sum(r.aantal_afvoer_punten for r in self.ruimtes)
        txt = "SAMENVATTING\n============\n\nRuimtes: {} | Ventilatie/persoon: {:.1f} dm³/s\nAir Terminals: {} TV, {} AV\n\n".format(len(self.ruimtes), self.dm3_per_persoon, tot_at_t, tot_at_a)
        zones = {}
        for r in self.ruimtes:
            if r.zone not in zones:
                zones[r.zone] = {'t': 0, 'a': 0, 'at': 0, 'at_t': 0, 'at_a': 0}
            zones[r.zone]['t'] += r.toevoer
            zones[r.zone]['a'] += r.afvoer
            zones[r.zone]['at'] += r.afvoer_totaal
            zones[r.zone]['at_t'] += r.aantal_toevoer_punten
            zones[r.zone]['at_a'] += r.aantal_afvoer_punten
        for zn, zd in sorted(zones.items()):
            txt += "{}\n  Eis: {:.1f} TV | {:.1f} AV | Terminals: {} TV, {} AV\n".format(zn.upper(), zd['t'], zd['a'], zd['at_t'], zd['at_a'])
        txt += "\nTOTALEN\n  Toevoer: {:.1f} dm³/s\n  Afvoer: {:.1f} dm³/s\n  Ontwerp: {:.1f} dm³/s\n".format(tot_t, tot_a, tot_at)
        self.lbl_samenvatting.Text = txt
    
    def _fill_airflow_click(self, s, e):
        try:
            updated = 0
            tv_counter = 0
            av_counter = 0
            with revit.Transaction("Vul Air Flow in"):
                for r in sorted(self.ruimtes, key=lambda x: (x.zone, x.naam)):
                    if r.air_terminals_toevoer and r.toevoer > 0:
                        flow_per_terminal = r.toevoer / len(r.air_terminals_toevoer)
                        for at in r.air_terminals_toevoer:
                            tv_counter += 1
                            param = at.element.LookupParameter('Air Flow')
                            if param and not param.IsReadOnly:
                                param.Set(flow_per_terminal / 28.3168)
                            comment_param = at.element.LookupParameter('Comments')
                            if comment_param and not comment_param.IsReadOnly:
                                comment_param.Set("{:02d}_TV_{:.1f}l/s".format(tv_counter, flow_per_terminal))
                            updated += 1
                    if r.air_terminals_afvoer and r.afvoer_totaal > 0:
                        flow_per_terminal = r.afvoer_totaal / len(r.air_terminals_afvoer)
                        for at in r.air_terminals_afvoer:
                            av_counter += 1
                            param = at.element.LookupParameter('Air Flow')
                            if param and not param.IsReadOnly:
                                param.Set(flow_per_terminal / 28.3168)
                            comment_param = at.element.LookupParameter('Comments')
                            if comment_param and not comment_param.IsReadOnly:
                                comment_param.Set("{:02d}_AV_{:.1f}l/s".format(av_counter, flow_per_terminal))
                            updated += 1
            MessageBox.Show("{} Air Terminals bijgewerkt!\n\n{} Toevoer (TV)\n{} Afvoer (AV)".format(updated, tv_counter, av_counter), "Air Flow", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as ex:
            MessageBox.Show(str(ex), "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
    
    def _export_csv_click(self, s, e):
        try:
            dlg = SaveFileDialog()
            dlg.Filter = "CSV (*.csv)|*.csv"
            dlg.FileName = "VentilatieBalans.csv"
            if dlg.ShowDialog() == DialogResult.OK:
                with open(dlg.FileName, 'w') as f:
                    f.write("Zone;Naam;Niveau;Functie;m2;Pers;Type;Eis;Toevoer;Afvoer;Correctie;Totaal;TV;AV;Status\n")
                    for r in self.ruimtes:
                        f.write("{};{};{};{};{:.1f};{};{};{:.1f};{:.1f};{:.1f};{:.1f};{:.1f};{};{};{}\n".format(r.zone, r.naam, r.niveau, r.gebruiksfunctie or "", r.oppervlakte, r.aantal_personen, r.get_effective_type(), r.ventilatie_eis, r.toevoer, r.afvoer, r.afvoer_correctie, r.afvoer_totaal, r.aantal_toevoer_punten, r.aantal_afvoer_punten, r.mv_status))
                MessageBox.Show("Export succesvol!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as ex:
            MessageBox.Show(str(ex), "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
    
    def _export_sheet_click(self, s, e):
        try:
            sheet_nr = forms.ask_for_string(default='OV-700c', prompt='Sheet nummer:', title='Export')
            if not sheet_nr:
                return
            sheets = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSheet).ToElements()
            target_sheet = None
            for sh in sheets:
                if sh.SheetNumber == sheet_nr:
                    target_sheet = sh
                    break
            if not target_sheet:
                MessageBox.Show("Sheet niet gevonden", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            
            zones_data = self._get_zones_data()
            
            lines = []
            lines.append("RUIMTELIJST VENTILATIE")
            lines.append("=" * 60)
            lines.append("")
            lines.append("{:<14}{:<16}{:<14}{:>6} {:>7} {:>6} {:>6}".format("Zone", "Naam", "Functie", "m²", "Type", "TV", "AV"))
            lines.append("-" * 70)
            
            for r in sorted(self.ruimtes, key=lambda x: (x.zone, x.naam)):
                zone = r.zone[:13] if len(r.zone) > 13 else r.zone
                naam = r.naam[:15] if len(r.naam) > 15 else r.naam
                functie = (r.gebruiksfunctie or "-")[:13]
                tv = "{:.1f}".format(r.toevoer) if r.toevoer > 0 else "-"
                av = "{:.1f}".format(r.afvoer_totaal) if r.afvoer_totaal > 0 else "-"
                lines.append("{:<14}{:<16}{:<14}{:>5.1f} {:>7} {:>6} {:>6}".format(zone, naam, functie, r.oppervlakte, r.get_effective_type()[:3], tv, av))
            
            txt1 = "\n".join(lines)
            
            lines2 = []
            lines2.append("BALANS PER ZONE")
            lines2.append("=" * 40)
            lines2.append("")
            
            for zn, zd in sorted(zones_data.items()):
                lines2.append(zn.upper())
                lines2.append("  Toevoer:  {:>6.1f} dm³/s".format(zd['toevoer']))
                lines2.append("  Afvoer:   {:>6.1f} dm³/s".format(zd['afvoer_totaal']))
                balans = zd['toevoer'] - zd['afvoer_totaal']
                if abs(balans) < 1:
                    lines2.append("  Status:   IN BALANS")
                elif balans > 0:
                    lines2.append("  Status:   OVERDRUK +{:.1f}".format(balans))
                else:
                    lines2.append("  Status:   ONDERDRUK {:.1f}".format(balans))
                lines2.append("")
            
            txt2 = "\n".join(lines2)
            
            tot_tv = sum(zd['toevoer'] for zd in zones_data.values())
            tot_av = sum(zd['afvoer_totaal'] for zd in zones_data.values())
            
            lines3 = []
            lines3.append("TOTALEN")
            lines3.append("=" * 30)
            lines3.append("")
            lines3.append("Ventilatie/persoon: {:.1f} dm³/s".format(self.dm3_per_persoon))
            lines3.append("")
            lines3.append("Totaal toevoer:  {:>6.1f} dm³/s".format(tot_tv))
            lines3.append("Totaal afvoer:   {:>6.1f} dm³/s".format(tot_av))
            
            txt3 = "\n".join(lines3)
            
            with revit.Transaction("Ventilatie Export"):
                text_type_id = DB.FilteredElementCollector(revit.doc).OfClass(DB.TextNoteType).FirstElementId()
                if text_type_id and text_type_id != DB.ElementId.InvalidElementId:
                    DB.TextNote.Create(revit.doc, target_sheet.Id, DB.XYZ(0.02, 0.85, 0), txt1, text_type_id)
                    DB.TextNote.Create(revit.doc, target_sheet.Id, DB.XYZ(0.02, 0.25, 0), txt2, text_type_id)
                    DB.TextNote.Create(revit.doc, target_sheet.Id, DB.XYZ(0.55, 0.25, 0), txt3, text_type_id)
            
            MessageBox.Show("Geplaatst op sheet {}!".format(sheet_nr), "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as ex:
            MessageBox.Show(str(ex), "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
    
    def _close_click(self, s, e):
        self.Close()


# ==============================================================================
# ENTRY POINT
# ==============================================================================
if __name__ == '__main__':
    try:
        form = VentilatieBalansForm()
        form.ShowDialog()
    except Exception as e:
        forms.alert("Fout: {}".format(str(e)), title="Ventilatie Balans")
