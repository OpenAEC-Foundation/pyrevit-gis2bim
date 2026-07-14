# -*- coding: utf-8 -*-
"""Ventilatievoud Bouwaanvraag - Native pyRevit Tool

Berekent per ruimte de ventilatie-eis volgens het Bouwbesluit/Bbl
(nieuwbouw woonfunctie) en de benodigde lengte Ducoton 10 'ZR'
glasroosters voor de toevoer. Plaatst het resultaat als tekst-tabel
op de actieve view (bouwaanvraag).
"""

__title__ = "Ventilatie\nTabel"
__author__ = "3BM Bouwkunde"

import os
import sys
import json
import math
import codecs
from datetime import date

# Voeg lib folder toe aan path
lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

from pyrevit import revit, DB, forms

from ui_template import DPIScaler, Huisstijl, UIFactory
from bm_logger import get_logger

log = get_logger("Ventilatievoud")

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, Panel, DataGridView, DataGridViewTextBoxColumn,
    DataGridViewComboBoxColumn, MessageBox, MessageBoxButtons,
    MessageBoxIcon, FormStartPosition, BorderStyle, FlatStyle,
    DataGridViewAutoSizeColumnsMode, DataGridViewSelectionMode,
    NumericUpDown, AnchorStyles, AutoScaleMode, CheckBox
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# ==============================================================================
# CONSTANTEN
# ==============================================================================
ROOSTER_NAAM = "Ducoton 10 'ZR'"
ROOSTER_CAPACITEIT = 10.2   # dm3/s per m1 bij 1 Pa (zonder DucoFilter)
ROOSTER_AFRONDING_MM = 50   # roosterlengte afronden op veelvoud
ROOSTER_MAX_MM = 2500       # max. lengte per rooster (garantiegrens Duco)

NORM_VERBLIJF_DM3_M2 = 0.9  # dm3/s per m2 (Bbl art. 4.122, verblijfsgebied)
NORM_VERBLIJF_MIN = 7.0     # dm3/s minimum per verblijfsruimte

DEFAULT_TEXT_TYPE = '3BM_2mm'

FT2_NAAR_M2 = 0.092903
FT3_NAAR_M3 = 0.0283168

EXCHANGE_MAP = os.path.join(os.environ.get('TEMP', ''), '3bm_exchange')
EXCHANGE_BESTAND = 'ventilatie_bouwaanvraag.json'

# Ruimtetypen: regel bepaalt berekening.
# toevoer = 0,9 dm3/s per m2 met minimum (verblijfsruimte)
# afvoer  = vaste eis in dm3/s (mechanische afvoer)
RUIMTE_TYPES = [
    ('verblijfsruimte', {'toevoer': True, 'afvoer': 0.0}),
    ('woonkeuken',      {'toevoer': True, 'afvoer': 21.0}),
    ('keuken',          {'toevoer': False, 'afvoer': 21.0}),
    ('badruimte',       {'toevoer': False, 'afvoer': 14.0}),
    ('toiletruimte',    {'toevoer': False, 'afvoer': 7.0}),
    ('wasruimte',       {'toevoer': False, 'afvoer': 14.0}),
    ('geen eis',        {'toevoer': False, 'afvoer': 0.0}),
]
RUIMTE_TYPE_MAP = dict(RUIMTE_TYPES)

GEEN_EIS_KEYWORDS = [
    'hal', 'gang', 'entree', 'overloop', 'trap', 'berging', 'garage',
    'zolder', 'vide', 'meterkast', 'techniek', 'installatie', 'schacht',
    'kast', 'buiten', 'terras', 'balkon', 'carport', 'vliering'
]


def detecteer_type(naam):
    """Bepaal ruimtetype op basis van de ruimtenaam."""
    n = naam.lower().strip()
    if 'woonkeuken' in n or ('woon' in n and 'keuken' in n):
        return 'woonkeuken'
    if 'keuken' in n:
        return 'keuken'
    if 'bad' in n or 'douche' in n:
        return 'badruimte'
    if 'toilet' in n or n == 'wc' or n.startswith('wc ') or n.endswith(' wc'):
        return 'toiletruimte'
    if 'washok' in n or 'wasruimte' in n or 'wasmachine' in n:
        return 'wasruimte'
    for kw in GEEN_EIS_KEYWORDS:
        if kw in n:
            return 'geen eis'
    return 'verblijfsruimte'


def nl_getal(waarde, decimalen=1):
    """Formatteer getal met komma als decimaalteken (NL tekeningen)."""
    return ('{:.' + str(decimalen) + 'f}').format(waarde).replace('.', ',')


# ==============================================================================
# DATA
# ==============================================================================
class RuimteRij(object):
    """Een ruimte met berekende ventilatie-eis en roosterlengte."""

    def __init__(self, room):
        self.element = room
        self.element_id = room.Id.IntegerValue
        self.naam = self._param_str(room, DB.BuiltInParameter.ROOM_NAME) or 'Naamloos'
        self.nummer = self._param_str(room, DB.BuiltInParameter.ROOM_NUMBER) or '-'
        self.niveau = self._level_naam(room)
        self.opp = self._area(room)
        self.volume = self._volume(room)
        self.type_key = detecteer_type(self.naam)
        self.toevoer = 0.0
        self.afvoer = 0.0
        self.rooster_mm = 0
        self.rooster_aantal = 0
        self.rooster_tekst = '-'
        self.voud = 0.0

    def _param_str(self, room, bip):
        p = room.get_Parameter(bip)
        if p and p.HasValue:
            return p.AsString()
        return None

    def _level_naam(self, room):
        level_id = room.LevelId
        if level_id and level_id != DB.ElementId.InvalidElementId:
            level = revit.doc.GetElement(level_id)
            if level:
                return level.Name
        return 'Onbekend'

    def _area(self, room):
        p = room.get_Parameter(DB.BuiltInParameter.ROOM_AREA)
        if p and p.HasValue:
            return p.AsDouble() * FT2_NAAR_M2
        return 0.0

    def _volume(self, room):
        p = room.get_Parameter(DB.BuiltInParameter.ROOM_VOLUME)
        if p and p.HasValue:
            return p.AsDouble() * FT3_NAAR_M3
        return 0.0

    def bereken(self, factor, minimum, capaciteit, afronding_mm):
        """Herbereken eis, roosterlengte en ventilatievoud."""
        cfg = RUIMTE_TYPE_MAP.get(self.type_key, RUIMTE_TYPE_MAP['geen eis'])
        self.toevoer = 0.0
        self.afvoer = float(cfg['afvoer'])
        self.rooster_mm = 0
        self.rooster_aantal = 0
        self.rooster_tekst = '-'
        self.voud = 0.0

        if cfg['toevoer'] and self.opp > 0:
            self.toevoer = round(max(self.opp * factor, minimum), 1)
            if capaciteit > 0:
                ruw_mm = self.toevoer / capaciteit * 1000.0
                aantal = int(math.ceil(ruw_mm / ROOSTER_MAX_MM))
                per_stuk = ruw_mm / aantal
                per_stuk = int(math.ceil(per_stuk / afronding_mm) * afronding_mm)
                self.rooster_mm = per_stuk * aantal
                self.rooster_aantal = aantal
                if aantal > 1:
                    self.rooster_tekst = '{}x {} mm'.format(aantal, per_stuk)
                else:
                    self.rooster_tekst = '{} mm'.format(per_stuk)

        qv = max(self.toevoer, self.afvoer)
        if qv > 0 and self.volume > 0:
            self.voud = round(qv * 3.6 / self.volume, 1)

    @property
    def heeft_eis(self):
        return self.toevoer > 0 or self.afvoer > 0


def laad_ruimtes():
    """Verzamel alle geplaatste ruimtes met oppervlak uit het model."""
    ruimtes = []
    collector = DB.FilteredElementCollector(revit.doc)\
        .OfCategory(DB.BuiltInCategory.OST_Rooms)\
        .WhereElementIsNotElementType()
    for room in collector:
        if room.Area > 0:
            ruimtes.append(RuimteRij(room))
    ruimtes.sort(key=lambda r: (r.niveau, r.nummer, r.naam))
    return ruimtes


def laad_teksttypes():
    """Alle TextNoteTypes in het document, gesorteerd op naam."""
    types = []
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.TextNoteType)
    for tt in collector:
        naam = DB.Element.Name.__get__(tt)
        types.append((naam, tt.Id))
    types.sort(key=lambda t: t[0].lower())
    return types


# ==============================================================================
# TABEL-TEKST
# ==============================================================================
def bouw_tabel_tekst(ruimtes, factor, minimum, capaciteit, toon_zonder_eis):
    """Bouw de tekst-tabel voor op de tekening (vaste kolombreedtes)."""
    heeft_voud = any(r.voud > 0 for r in ruimtes)

    kop = "{:<5}{:<20}{:<12}{:>7}  {:>9}  {:<13}{:>9}".format(
        "Nr", "Ruimte", "Verdieping", "Opp.", "Toevoer", "Rooster", "Afvoer")
    kop2 = "{:<5}{:<20}{:<12}{:>7}  {:>9}  {:<13}{:>9}".format(
        "", "", "", "[m2]", "[dm3/s]", ROOSTER_NAAM.split(' ')[0], "[dm3/s]")
    if heeft_voud:
        kop += "  {:>7}".format("Voud")
        kop2 += "  {:>7}".format("[1/h]")
    breedte = max(len(kop), len(kop2))

    regels = []
    regels.append(u"VENTILATIEBEREKENING BOUWAANVRAAG")
    regels.append(u"Luchtverversing conform Bbl afd. 4.3 (nieuwbouw, woonfunctie)")
    regels.append(u"Project: {}   Datum: {}".format(
        revit.doc.Title, date.today().strftime('%d-%m-%Y')))
    regels.append(u"")
    regels.append(kop)
    regels.append(kop2)
    regels.append(u"-" * breedte)

    tot_toevoer = 0.0
    tot_afvoer = 0.0
    for r in ruimtes:
        if not r.heeft_eis and not toon_zonder_eis:
            continue
        tot_toevoer += r.toevoer
        tot_afvoer += r.afvoer
        regel = u"{:<5}{:<20}{:<12}{:>7}  {:>9}  {:<13}{:>9}".format(
            r.nummer[:4],
            r.naam[:19],
            r.niveau[:11],
            nl_getal(r.opp),
            nl_getal(r.toevoer) if r.toevoer > 0 else u"-",
            r.rooster_tekst,
            nl_getal(r.afvoer) if r.afvoer > 0 else u"-")
        if heeft_voud:
            regel += u"  {:>7}".format(nl_getal(r.voud) if r.voud > 0 else u"-")
        regels.append(regel)

    regels.append(u"-" * breedte)
    regels.append(u"Totaal toevoer: {} dm3/s   Totaal afvoer: {} dm3/s".format(
        nl_getal(tot_toevoer), nl_getal(tot_afvoer)))
    regels.append(u"")
    regels.append(u"Uitgangspunten:")
    regels.append(u"- Toevoer via {} zelfregelende glasroosters: {} dm3/s per m1 bij 1 Pa".format(
        ROOSTER_NAAM, nl_getal(capaciteit)))
    regels.append(u"- Verblijfsruimten: {} dm3/s per m2, minimaal {} dm3/s per ruimte".format(
        nl_getal(factor), nl_getal(minimum, 0)))
    regels.append(u"- Afvoer: keuken 21 / badruimte 14 / toiletruimte 7 dm3/s (mechanisch)")
    regels.append(u"- Roosterlengten (werkend) afgerond op {} mm, max. {} mm per rooster".format(
        ROOSTER_AFRONDING_MM, ROOSTER_MAX_MM))
    return u"\n".join(regels)


def schrijf_exchange_json(ruimtes, factor, minimum, capaciteit):
    """Schrijf resultaat naar %TEMP%\\3bm_exchange voor andere tools."""
    try:
        if not os.path.isdir(EXCHANGE_MAP):
            os.makedirs(EXCHANGE_MAP)
        data = {
            'project': revit.doc.Title,
            'datum': date.today().isoformat(),
            'rooster': ROOSTER_NAAM,
            'capaciteit_dm3s_per_m': capaciteit,
            'norm_dm3s_per_m2': factor,
            'norm_minimum_dm3s': minimum,
            'ruimtes': [{
                'nummer': r.nummer,
                'naam': r.naam,
                'niveau': r.niveau,
                'oppervlakte_m2': round(r.opp, 2),
                'type': r.type_key,
                'toevoer_dm3s': r.toevoer,
                'afvoer_dm3s': r.afvoer,
                'rooster_mm': r.rooster_mm,
                'rooster_aantal': r.rooster_aantal,
                'ventilatievoud_per_h': r.voud,
            } for r in ruimtes],
        }
        pad = os.path.join(EXCHANGE_MAP, EXCHANGE_BESTAND)
        with codecs.open(pad, 'w', 'utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as ex:
        log.warning("Exchange JSON niet geschreven: {}".format(ex))


# ==============================================================================
# FORM
# ==============================================================================
class VentilatievoudForm(Form):
    """Huisstijl-form: ruimtelijst met type-override en instellingen."""

    def __init__(self, ruimtes):
        self.ruimtes = ruimtes
        self.teksttypes = laad_teksttypes()
        self.plaats_data = None  # gezet bij klik op 'Plaats tabel'
        self._updating = False
        self._setup_form()
        self._herbereken()
        self._vul_grid()

    # -- instellingen uit de UI --------------------------------------------
    @property
    def factor(self):
        return float(self.nud_factor.Value)

    @property
    def minimum(self):
        return float(self.nud_minimum.Value)

    @property
    def capaciteit(self):
        return float(self.nud_capaciteit.Value)

    def _setup_form(self):
        self.Text = "Ventilatievoud Bouwaanvraag - 3BM Bouwkunde"
        self.Size = DPIScaler.scale_size(1050, 700)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.White
        self.AutoScaleMode = AutoScaleMode.Dpi

        # Header
        self.pnl_header = Panel()
        self.pnl_header.Location = Point(0, 0)
        self.pnl_header.Size = Size(self.ClientSize.Width, DPIScaler.scale(55))
        self.pnl_header.BackColor = Huisstijl.VIOLET
        self.pnl_header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_header)

        lbl_header = Label()
        lbl_header.Text = "Ventilatievoud Bouwaanvraag"
        lbl_header.Font = Font("Segoe UI", 16, FontStyle.Bold)
        lbl_header.ForeColor = Color.White
        lbl_header.Location = DPIScaler.scale_point(20, 12)
        lbl_header.AutoSize = True
        self.pnl_header.Controls.Add(lbl_header)

        self.lbl_status = Label()
        self.lbl_status.Location = DPIScaler.scale_point(420, 18)
        self.lbl_status.Size = DPIScaler.scale_size(580, 20)
        self.lbl_status.ForeColor = Huisstijl.TEAL
        self.lbl_status.Font = Font("Segoe UI", 9)
        self.pnl_header.Controls.Add(self.lbl_status)

        # Accent
        pnl_accent = Panel()
        pnl_accent.Location = Point(0, DPIScaler.scale(55))
        pnl_accent.Size = Size(self.ClientSize.Width, DPIScaler.scale(5))
        pnl_accent.BackColor = Huisstijl.TEAL
        pnl_accent.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(pnl_accent)

        # Instellingen-balk
        pnl_settings = Panel()
        pnl_settings.Location = Point(DPIScaler.scale(10), DPIScaler.scale(65))
        pnl_settings.Size = Size(self.ClientSize.Width - DPIScaler.scale(20), DPIScaler.scale(40))
        pnl_settings.BackColor = Huisstijl.LIGHT_GRAY
        pnl_settings.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(pnl_settings)

        x = DPIScaler.scale(10)
        y = DPIScaler.scale(10)

        lbl_c = UIFactory.create_label("Roostercapaciteit:")
        lbl_c.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_c)
        x += DPIScaler.scale(115)

        self.nud_capaciteit = self._maak_nud(x, y, ROOSTER_CAPACITEIT, 1, 50, 0.1, 1)
        pnl_settings.Controls.Add(self.nud_capaciteit)
        x += DPIScaler.scale(65)

        lbl_c2 = UIFactory.create_label("dm3/s per m1")
        lbl_c2.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_c2)
        x += DPIScaler.scale(95)

        lbl_f = UIFactory.create_label("Verblijfsruimte:")
        lbl_f.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_f)
        x += DPIScaler.scale(95)

        self.nud_factor = self._maak_nud(x, y, NORM_VERBLIJF_DM3_M2, 0.1, 5, 0.1, 1)
        pnl_settings.Controls.Add(self.nud_factor)
        x += DPIScaler.scale(65)

        lbl_f2 = UIFactory.create_label("dm3/s per m2, min.")
        lbl_f2.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_f2)
        x += DPIScaler.scale(120)

        self.nud_minimum = self._maak_nud(x, y, NORM_VERBLIJF_MIN, 0, 50, 1, 0)
        pnl_settings.Controls.Add(self.nud_minimum)
        x += DPIScaler.scale(65)

        lbl_m2 = UIFactory.create_label("dm3/s")
        lbl_m2.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_m2)
        x += DPIScaler.scale(60)

        lbl_t = UIFactory.create_label("Teksttype:")
        lbl_t.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_t)
        x += DPIScaler.scale(70)

        self.cmb_teksttype = UIFactory.create_combobox(160, [t[0] for t in self.teksttypes])
        self.cmb_teksttype.Location = Point(x, y - DPIScaler.scale(3))
        for i, (naam, _) in enumerate(self.teksttypes):
            if naam == DEFAULT_TEXT_TYPE:
                self.cmb_teksttype.SelectedIndex = i
                break
        pnl_settings.Controls.Add(self.cmb_teksttype)
        x += DPIScaler.scale(175)

        self.chk_zonder_eis = CheckBox()
        self.chk_zonder_eis.Text = "Ruimtes zonder eis tonen"
        self.chk_zonder_eis.AutoSize = True
        self.chk_zonder_eis.Font = Font("Segoe UI", 9)
        self.chk_zonder_eis.Checked = True
        self.chk_zonder_eis.Location = Point(x, y - DPIScaler.scale(2))
        pnl_settings.Controls.Add(self.chk_zonder_eis)

        # Grid
        grid_top = DPIScaler.scale(115)
        self.grid = DataGridView()
        self.grid.Location = Point(DPIScaler.scale(10), grid_top)
        self.grid.Size = Size(self.ClientSize.Width - DPIScaler.scale(20),
                              self.ClientSize.Height - grid_top - DPIScaler.scale(60))
        self.grid.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom |
                            AnchorStyles.Left | AnchorStyles.Right)
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

        for name, header, width, readonly in [
                ("nummer", "Nr", 45, True),
                ("naam", "Ruimte", 160, True),
                ("niveau", "Verdieping", 100, True),
                ("opp", "m²", 55, True)]:
            col = DataGridViewTextBoxColumn()
            col.Name = name
            col.HeaderText = header
            col.Width = DPIScaler.scale(width)
            col.ReadOnly = readonly
            self.grid.Columns.Add(col)

        type_col = DataGridViewComboBoxColumn()
        type_col.Name = "type"
        type_col.HeaderText = "Type"
        type_col.Width = DPIScaler.scale(120)
        for key, _ in RUIMTE_TYPES:
            type_col.Items.Add(key)
        type_col.FlatStyle = FlatStyle.Flat
        self.grid.Columns.Add(type_col)

        for name, header, width in [
                ("toevoer", "Toevoer [dm³/s]", 95),
                ("rooster", "Ducoton 10", 100),
                ("afvoer", "Afvoer [dm³/s]", 95),
                ("voud", "Voud [1/h]", 75)]:
            col = DataGridViewTextBoxColumn()
            col.Name = name
            col.HeaderText = header
            col.Width = DPIScaler.scale(width)
            col.ReadOnly = True
            self.grid.Columns.Add(col)

        self.grid.CellValueChanged += self._grid_changed
        self.grid.CurrentCellDirtyStateChanged += self._grid_dirty
        self.Controls.Add(self.grid)

        # Footer knoppen
        btn_y = self.ClientSize.Height - DPIScaler.scale(48)

        self.btn_plaats = UIFactory.create_button("Plaats tabel op view", 170, 38, 'primary')
        self.btn_plaats.Location = Point(self.ClientSize.Width - DPIScaler.scale(300), btn_y)
        self.btn_plaats.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_plaats.Click += self._plaats_click
        self.Controls.Add(self.btn_plaats)

        self.btn_close = UIFactory.create_button("Sluiten", 100, 38, 'secondary')
        self.btn_close.Location = Point(self.ClientSize.Width - DPIScaler.scale(120), btn_y)
        self.btn_close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_close.Click += self._close_click
        self.Controls.Add(self.btn_close)

        # Instellingen -> herberekenen
        self.nud_capaciteit.ValueChanged += self._instelling_changed
        self.nud_factor.ValueChanged += self._instelling_changed
        self.nud_minimum.ValueChanged += self._instelling_changed

    def _maak_nud(self, x, y, waarde, minimum, maximum, stap, decimalen):
        nud = NumericUpDown()
        nud.Location = Point(x, y - DPIScaler.scale(3))
        nud.Size = DPIScaler.scale_size(58, 25)
        nud.Font = Font("Segoe UI", 9)
        nud.DecimalPlaces = decimalen
        nud.Minimum = minimum
        nud.Maximum = maximum
        nud.Increment = stap
        nud.Value = waarde
        return nud

    # -- berekening & weergave ---------------------------------------------
    def _herbereken(self):
        for r in self.ruimtes:
            r.bereken(self.factor, self.minimum, self.capaciteit, ROOSTER_AFRONDING_MM)
        tot_t = sum(r.toevoer for r in self.ruimtes)
        tot_a = sum(r.afvoer for r in self.ruimtes)
        self.lbl_status.Text = "{} ruimtes | toevoer {} dm3/s | afvoer {} dm3/s".format(
            len(self.ruimtes), nl_getal(tot_t), nl_getal(tot_a))

    def _vul_grid(self):
        self._updating = True
        try:
            self.grid.Rows.Clear()
            for r in self.ruimtes:
                idx = self.grid.Rows.Add(
                    r.nummer, r.naam, r.niveau, nl_getal(r.opp), r.type_key,
                    nl_getal(r.toevoer) if r.toevoer > 0 else "-",
                    r.rooster_tekst,
                    nl_getal(r.afvoer) if r.afvoer > 0 else "-",
                    nl_getal(r.voud) if r.voud > 0 else "-")
                if r.type_key != detecteer_type(r.naam):
                    self.grid.Rows[idx].Cells["type"].Style.BackColor = Color.FromArgb(230, 255, 230)
        finally:
            self._updating = False

    def _instelling_changed(self, s, e):
        if self._updating:
            return
        self._herbereken()
        self._vul_grid()

    def _grid_dirty(self, s, e):
        if self.grid.IsCurrentCellDirty:
            self.grid.CommitEdit(1)

    def _grid_changed(self, s, e):
        if self._updating or e.RowIndex < 0:
            return
        if e.ColumnIndex == self.grid.Columns["type"].Index:
            idx = e.RowIndex
            if 0 <= idx < len(self.ruimtes):
                self.ruimtes[idx].type_key = str(self.grid.Rows[idx].Cells["type"].Value)
                self._herbereken()
                self._vul_grid()

    # -- acties --------------------------------------------------------------
    def _plaats_click(self, s, e):
        if not self.teksttypes:
            MessageBox.Show("Geen teksttypes in dit document gevonden.",
                            "Fout", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        idx = self.cmb_teksttype.SelectedIndex
        if idx < 0:
            idx = 0
        tekst = bouw_tabel_tekst(self.ruimtes, self.factor, self.minimum,
                                 self.capaciteit, self.chk_zonder_eis.Checked)
        self.plaats_data = {
            'tekst': tekst,
            'teksttype_id': self.teksttypes[idx][1],
            'teksttype_naam': self.teksttypes[idx][0],
        }
        self.Close()

    def _close_click(self, s, e):
        self.Close()


# ==============================================================================
# PLAATSING
# ==============================================================================
def _view_centrum(view):
    """Middelpunt van de crop box van een view, in modelcoordinaten."""
    try:
        box = view.CropBox
        midden = (box.Min + box.Max) * 0.5
        return box.Transform.OfPoint(midden)
    except Exception:
        return DB.XYZ(0, 0, 0)


def plaats_tabel(plaats_data, ruimtes, factor, minimum, capaciteit):
    """Vraag klikpunt op de actieve view en plaats de TextNote."""
    view = revit.doc.ActiveView
    if isinstance(view, DB.View3D):
        forms.alert("De actieve view is een 3D-view. Open een plattegrond, "
                    "drafting view of sheet en start de tool opnieuw.",
                    title="Ventilatievoud")
        return

    from Autodesk.Revit.Exceptions import OperationCanceledException
    try:
        punt = revit.uidoc.Selection.PickPoint("Klik het plaatsingspunt voor de ventilatietabel")
    except OperationCanceledException:
        return  # geannuleerd door gebruiker
    except Exception:
        # PickPoint niet mogelijk (bv. geen werkvlak) -> plaats in view-centrum
        punt = _view_centrum(view)

    with revit.Transaction("Ventilatietabel plaatsen"):
        DB.TextNote.Create(revit.doc, view.Id, punt,
                           plaats_data['tekst'], plaats_data['teksttype_id'])

    schrijf_exchange_json(ruimtes, factor, minimum, capaciteit)
    log.info("Ventilatietabel geplaatst op view '{}' (teksttype {})".format(
        view.Name, plaats_data['teksttype_naam']))
    forms.alert("Ventilatietabel geplaatst op view '{}'.\n\n"
                "Teksttype: {}\nTip: gebruik een monospace-teksttype als de "
                "kolommen niet uitlijnen.".format(view.Name, plaats_data['teksttype_naam']),
                title="Ventilatievoud")


# ==============================================================================
# ENTRY POINT
# ==============================================================================
if __name__ == '__main__':
    try:
        alle_ruimtes = laad_ruimtes()
        if not alle_ruimtes:
            forms.alert("Geen geplaatste ruimtes (Rooms) gevonden in het model.\n"
                        "Plaats eerst Rooms met een oppervlak.",
                        title="Ventilatievoud", exitscript=True)
        form = VentilatievoudForm(alle_ruimtes)
        form.ShowDialog()
        if form.plaats_data:
            plaats_tabel(form.plaats_data, form.ruimtes,
                         form.factor, form.minimum, form.capaciteit)
    except Exception as e:
        log.error("Ventilatievoud fout: {}".format(e))
        forms.alert("Fout: {}".format(str(e)), title="Ventilatievoud")
