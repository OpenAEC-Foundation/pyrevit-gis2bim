# -*- coding: utf-8 -*-
"""Bbl-Toets Bouwaanvraag - Native pyRevit Tool

Toetst per ruimte aan het Bbl (nieuwbouw woonfunctie):
- Ventilatie (par. 4.3.5): eis per ruimte + benodigde lengte
  Ducoton 10 'ZR' glasroosters voor de toevoer
- Daglicht (par. 4.3.7): equivalente daglichtoppervlakte
  (10% vloeroppervlak, min. 0,5 m2 per verblijfsruimte)

Plaatst het resultaat als tekst-tabel op de actieve view (bouwaanvraag).
"""

__title__ = "Bbl\nToets"
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

log = get_logger("BblToets")

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, Panel, DataGridView, DataGridViewTextBoxColumn,
    DataGridViewComboBoxColumn, MessageBox, MessageBoxButtons,
    MessageBoxIcon, FormStartPosition, BorderStyle, FlatStyle,
    DataGridViewAutoSizeColumnsMode, DataGridViewSelectionMode,
    NumericUpDown, AnchorStyles, AutoScaleMode, CheckBox,
    TabControl, TabPage
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# ==============================================================================
# CONSTANTEN — VENTILATIE
# ==============================================================================
ROOSTER_NAAM = "Ducoton 10 'ZR'"
ROOSTER_CAPACITEIT = 10.2   # dm3/s per m1 bij 1 Pa (zonder DucoFilter)
ROOSTER_AFRONDING_MM = 50   # roosterlengte afronden op veelvoud
ROOSTER_MAX_MM = 2500       # max. lengte per rooster (garantiegrens Duco)

NORM_VERBLIJF_DM3_M2 = 0.9  # dm3/s per m2 (Bbl par. 4.3.5, verblijfsgebied)
NORM_VERBLIJF_MIN = 7.0     # dm3/s minimum per verblijfsruimte

# ==============================================================================
# CONSTANTEN — DAGLICHT
# ==============================================================================
DAGLICHT_PCT = 0.10         # equivalente daglichtopp. >= 10% vloeroppervlak
DAGLICHT_MIN_M2 = 0.5       # minimum per verblijfsruimte (Bbl par. 4.3.7)
GLASFACTOR = 0.80           # glasdeel van bruto kozijnoppervlak (default)
BELEMMERINGSFACTOR = 1.0    # Cb x Cu cf. NEN 2057 (default: vrije ligging)

RAAM_BREEDTE_PARAMS = ('kozijn_breedte', 'Breedte', 'Width')
RAAM_HOOGTE_PARAMS = ('kozijn_hoogte', 'Hoogte', 'Height')

DEFAULT_TEXT_TYPE = '3BM_2mm'

FT_NAAR_M = 0.3048
FT2_NAAR_M2 = 0.092903
FT3_NAAR_M3 = 0.0283168

EXCHANGE_MAP = os.path.join(os.environ.get('TEMP', ''), '3bm_exchange')
EXCHANGE_BESTAND = 'bbl_toets_bouwaanvraag.json'

# Ruimtetypen: regels per aspect.
# toevoer  = 0,9 dm3/s per m2 met minimum (verblijfsruimte)
# afvoer   = vaste eis in dm3/s (mechanische afvoer)
# daglicht = equivalente daglichtoppervlakte-eis van toepassing
RUIMTE_TYPES = [
    ('verblijfsruimte', {'toevoer': True, 'afvoer': 0.0, 'daglicht': True}),
    ('woonkeuken',      {'toevoer': True, 'afvoer': 21.0, 'daglicht': True}),
    ('keuken',          {'toevoer': False, 'afvoer': 21.0, 'daglicht': True}),
    ('badruimte',       {'toevoer': False, 'afvoer': 14.0, 'daglicht': False}),
    ('toiletruimte',    {'toevoer': False, 'afvoer': 7.0, 'daglicht': False}),
    ('wasruimte',       {'toevoer': False, 'afvoer': 14.0, 'daglicht': False}),
    ('geen eis',        {'toevoer': False, 'afvoer': 0.0, 'daglicht': False}),
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


def parse_getal(tekst):
    """Parse gebruikersinvoer met komma of punt naar float, None bij falen."""
    try:
        schoon = str(tekst).strip().replace('*', '').replace(',', '.')
        if not schoon:
            return None
        return float(schoon)
    except (ValueError, TypeError):
        return None


# ==============================================================================
# DATA
# ==============================================================================
class RuimteRij(object):
    """Een ruimte met berekende ventilatie- en daglicht-toets."""

    def __init__(self, room):
        self.element = room
        self.element_id = room.Id.IntegerValue
        self.naam = self._param_str(room, DB.BuiltInParameter.ROOM_NAME) or 'Naamloos'
        self.nummer = self._param_str(room, DB.BuiltInParameter.ROOM_NUMBER) or '-'
        self.niveau = self._level_naam(room)
        self.opp = self._area(room)
        self.volume = self._volume(room)
        self.type_key = detecteer_type(self.naam)
        # ventilatie
        self.toevoer = 0.0
        self.afvoer = 0.0
        self.rooster_mm = 0
        self.rooster_aantal = 0
        self.rooster_tekst = '-'
        self.voud = 0.0
        # daglicht
        self.raam_aantal = 0
        self.raam_opp = 0.0        # bruto kozijnoppervlak in de ruimte [m2]
        self.ae_vereist = 0.0
        self.ae_aanwezig = 0.0
        self.ae_override = None    # handmatige correctie [m2]
        self.daglicht_toets = '-'

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

    @property
    def config(self):
        return RUIMTE_TYPE_MAP.get(self.type_key, RUIMTE_TYPE_MAP['geen eis'])

    def bereken_ventilatie(self, factor, minimum, capaciteit, afronding_mm):
        """Herbereken ventilatie-eis, roosterlengte en ventilatievoud."""
        cfg = self.config
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

    def bereken_daglicht(self, glasfactor, belemmering):
        """Herbereken daglicht-eis en toets (equivalente daglichtopp.)."""
        cfg = self.config
        self.ae_vereist = 0.0
        self.ae_aanwezig = 0.0
        self.daglicht_toets = '-'
        if not cfg['daglicht'] or self.opp <= 0:
            return
        self.ae_vereist = round(max(self.opp * DAGLICHT_PCT, DAGLICHT_MIN_M2), 2)
        if self.ae_override is not None:
            self.ae_aanwezig = round(self.ae_override, 2)
        else:
            self.ae_aanwezig = round(self.raam_opp * glasfactor * belemmering, 2)
        if self.ae_aanwezig >= self.ae_vereist - 0.005:
            self.daglicht_toets = 'OK'
        else:
            self.daglicht_toets = 'TEKORT {} m2'.format(
                nl_getal(self.ae_vereist - self.ae_aanwezig, 2))

    @property
    def heeft_eis(self):
        return self.toevoer > 0 or self.afvoer > 0 or self.ae_vereist > 0


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


def _lees_lengte_ft(elem, namen, bip):
    """Lees een lengte-parameter (feet) van element: named params, dan BIP."""
    if elem is None:
        return 0.0
    for naam in namen:
        try:
            p = elem.LookupParameter(naam)
            if p and p.HasValue and p.StorageType == DB.StorageType.Double:
                v = p.AsDouble()
                if v > 0:
                    return v
        except Exception:
            pass
    if bip is not None:
        try:
            p = elem.get_Parameter(bip)
            if p and p.HasValue:
                v = p.AsDouble()
                if v > 0:
                    return v
        except Exception:
            pass
    return 0.0


def _raam_oppervlak_m2(inst):
    """Bruto kozijnoppervlak (b x h) van een window-instance in m2."""
    symbool = None
    try:
        symbool = inst.Symbol
    except Exception:
        pass
    b = _lees_lengte_ft(inst, RAAM_BREEDTE_PARAMS, None)
    if b <= 0:
        b = _lees_lengte_ft(symbool, RAAM_BREEDTE_PARAMS,
                            DB.BuiltInParameter.FAMILY_WIDTH_PARAM)
    h = _lees_lengte_ft(inst, RAAM_HOOGTE_PARAMS, None)
    if h <= 0:
        h = _lees_lengte_ft(symbool, RAAM_HOOGTE_PARAMS,
                            DB.BuiltInParameter.FAMILY_HEIGHT_PARAM)
    return (b * FT_NAAR_M) * (h * FT_NAAR_M)


def koppel_ramen(ruimtes):
    """Koppel windows aan ruimtes via FromRoom/ToRoom (laatste fase)."""
    per_ruimte = {}
    collector = DB.FilteredElementCollector(revit.doc)\
        .OfCategory(DB.BuiltInCategory.OST_Windows)\
        .WhereElementIsNotElementType()
    for inst in collector:
        room = None
        try:
            room = inst.ToRoom
        except Exception:
            pass
        if room is None:
            try:
                room = inst.FromRoom
            except Exception:
                pass
        if room is None:
            continue
        rid = room.Id.IntegerValue
        per_ruimte.setdefault(rid, []).append(_raam_oppervlak_m2(inst))
    for r in ruimtes:
        opps = per_ruimte.get(r.element_id, [])
        r.raam_aantal = len(opps)
        r.raam_opp = sum(opps)


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
def bouw_tabel_tekst(ruimtes, instellingen, toon_zonder_eis):
    """Bouw de gecombineerde Bbl-toets-tabel (vaste kolombreedtes)."""
    heeft_voud = any(r.voud > 0 for r in ruimtes)
    zichtbaar = [r for r in ruimtes if r.heeft_eis or toon_zonder_eis]

    regels = []
    regels.append(u"BBL-TOETS BOUWAANVRAAG")
    regels.append(u"Project: {}   Datum: {}".format(
        revit.doc.Title, date.today().strftime('%d-%m-%Y')))
    regels.append(u"")

    # -- 1. Ventilatie -------------------------------------------------------
    kop = u"{:<5}{:<20}{:<12}{:>7}  {:>9}  {:<13}{:>9}".format(
        "Nr", "Ruimte", "Verdieping", "Opp.", "Toevoer", "Rooster", "Afvoer")
    kop2 = u"{:<5}{:<20}{:<12}{:>7}  {:>9}  {:<13}{:>9}".format(
        "", "", "", "[m2]", "[dm3/s]", "Ducoton", "[dm3/s]")
    if heeft_voud:
        kop += u"  {:>7}".format("Voud")
        kop2 += u"  {:>7}".format("[1/h]")
    breedte = max(len(kop), len(kop2))

    regels.append(u"1. VENTILATIE (Bbl par. 4.3.5 - luchtverversing, nieuwbouw)")
    regels.append(kop)
    regels.append(kop2)
    regels.append(u"-" * breedte)

    tot_toevoer = 0.0
    tot_afvoer = 0.0
    for r in zichtbaar:
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

    # -- 2. Daglicht ---------------------------------------------------------
    kop_d = u"{:<5}{:<20}{:<12}{:>7}  {:>6}  {:>9}  {:>10}  {:<14}".format(
        "Nr", "Ruimte", "Verdieping", "Opp.", "Ramen", "Vereist", "Aanwezig", "Toets")
    kop_d2 = u"{:<5}{:<20}{:<12}{:>7}  {:>6}  {:>9}  {:>10}".format(
        "", "", "", "[m2]", "", "Ae [m2]", "Ae [m2]")
    breedte_d = len(kop_d)

    regels.append(u"2. DAGLICHT (Bbl par. 4.3.7 - equivalente daglichtoppervlakte)")
    regels.append(kop_d)
    regels.append(kop_d2)
    regels.append(u"-" * breedte_d)

    for r in zichtbaar:
        if r.ae_vereist <= 0 and not toon_zonder_eis:
            continue
        aanwezig = u"-"
        if r.ae_vereist > 0:
            aanwezig = nl_getal(r.ae_aanwezig, 2)
            if r.ae_override is not None:
                aanwezig += u"*"
        regels.append(u"{:<5}{:<20}{:<12}{:>7}  {:>6}  {:>9}  {:>10}  {:<14}".format(
            r.nummer[:4],
            r.naam[:19],
            r.niveau[:11],
            nl_getal(r.opp),
            str(r.raam_aantal) if r.raam_aantal > 0 else u"-",
            nl_getal(r.ae_vereist, 2) if r.ae_vereist > 0 else u"-",
            aanwezig,
            r.daglicht_toets))

    regels.append(u"-" * breedte_d)
    regels.append(u"")

    # -- Uitgangspunten ------------------------------------------------------
    regels.append(u"Uitgangspunten:")
    regels.append(u"- Toevoer via {} zelfregelende glasroosters: {} dm3/s per m1 bij 1 Pa".format(
        ROOSTER_NAAM, nl_getal(instellingen['capaciteit'])))
    regels.append(u"- Verblijfsruimten: {} dm3/s per m2, minimaal {} dm3/s per ruimte".format(
        nl_getal(instellingen['factor']), nl_getal(instellingen['minimum'], 0)))
    regels.append(u"- Afvoer: keuken 21 / badruimte 14 / toiletruimte 7 dm3/s (mechanisch)")
    regels.append(u"- Roosterlengten (werkend) afgerond op {} mm, max. {} mm per rooster".format(
        ROOSTER_AFRONDING_MM, ROOSTER_MAX_MM))
    regels.append(u"- Daglicht cf. NEN 2057: vereist Ae = 10% vloeroppervlak, "
                  u"min. {} m2 per verblijfsruimte".format(nl_getal(DAGLICHT_MIN_M2)))
    regels.append(u"- Aanwezig Ae = kozijnoppervlak x glasfactor {} x belemmeringsfactor "
                  u"CbxCu {}  (* = handmatig)".format(
        nl_getal(instellingen['glasfactor'], 2), nl_getal(instellingen['belemmering'], 2)))
    regels.append(u"- Spuiventilatie (Bbl par. 4.3.5) buiten beschouwing")
    return u"\n".join(regels)


def schrijf_exchange_json(ruimtes, instellingen):
    """Schrijf resultaat naar %TEMP%\\3bm_exchange voor andere tools."""
    try:
        if not os.path.isdir(EXCHANGE_MAP):
            os.makedirs(EXCHANGE_MAP)
        data = {
            'project': revit.doc.Title,
            'datum': date.today().isoformat(),
            'rooster': ROOSTER_NAAM,
            'instellingen': instellingen,
            'daglicht_pct': DAGLICHT_PCT,
            'daglicht_min_m2': DAGLICHT_MIN_M2,
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
                'raam_aantal': r.raam_aantal,
                'raam_opp_m2': round(r.raam_opp, 2),
                'ae_vereist_m2': r.ae_vereist,
                'ae_aanwezig_m2': r.ae_aanwezig,
                'ae_handmatig': r.ae_override is not None,
                'daglicht_toets': r.daglicht_toets,
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
class BblToetsForm(Form):
    """Huisstijl-form: tabs Ventilatie + Daglicht met type-override."""

    def __init__(self, ruimtes):
        self.ruimtes = ruimtes
        self.teksttypes = laad_teksttypes()
        self.plaats_data = None  # gezet bij klik op 'Plaats tabel'
        self._updating = False
        self._setup_form()
        self._herbereken()
        self._vul_grids()

    # -- instellingen uit de UI ----------------------------------------------
    @property
    def instellingen(self):
        return {
            'capaciteit': float(self.nud_capaciteit.Value),
            'factor': float(self.nud_factor.Value),
            'minimum': float(self.nud_minimum.Value),
            'glasfactor': float(self.nud_glasfactor.Value),
            'belemmering': float(self.nud_belemmering.Value),
        }

    def _setup_form(self):
        self.Text = "Bbl-Toets Bouwaanvraag - 3BM Bouwkunde"
        self.Size = DPIScaler.scale_size(1080, 730)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.White
        self.AutoScaleMode = AutoScaleMode.Dpi

        # Header
        pnl_header = Panel()
        pnl_header.Location = Point(0, 0)
        pnl_header.Size = Size(self.ClientSize.Width, DPIScaler.scale(55))
        pnl_header.BackColor = Huisstijl.VIOLET
        pnl_header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(pnl_header)

        lbl_header = Label()
        lbl_header.Text = "Bbl-Toets Bouwaanvraag"
        lbl_header.Font = Font("Segoe UI", 16, FontStyle.Bold)
        lbl_header.ForeColor = Color.White
        lbl_header.Location = DPIScaler.scale_point(20, 12)
        lbl_header.AutoSize = True
        pnl_header.Controls.Add(lbl_header)

        self.lbl_status = Label()
        self.lbl_status.Location = DPIScaler.scale_point(360, 18)
        self.lbl_status.Size = DPIScaler.scale_size(680, 20)
        self.lbl_status.ForeColor = Huisstijl.TEAL
        self.lbl_status.Font = Font("Segoe UI", 9)
        pnl_header.Controls.Add(self.lbl_status)

        # Accent
        pnl_accent = Panel()
        pnl_accent.Location = Point(0, DPIScaler.scale(55))
        pnl_accent.Size = Size(self.ClientSize.Width, DPIScaler.scale(5))
        pnl_accent.BackColor = Huisstijl.TEAL
        pnl_accent.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(pnl_accent)

        # Instellingen (2 rijen)
        pnl_settings = Panel()
        pnl_settings.Location = Point(DPIScaler.scale(10), DPIScaler.scale(65))
        pnl_settings.Size = Size(self.ClientSize.Width - DPIScaler.scale(20), DPIScaler.scale(70))
        pnl_settings.BackColor = Huisstijl.LIGHT_GRAY
        pnl_settings.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(pnl_settings)

        # Rij 1: ventilatie + teksttype
        x = DPIScaler.scale(10)
        y = DPIScaler.scale(8)

        lbl_v = UIFactory.create_label("Ventilatie:", bold=True)
        lbl_v.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_v)
        x += DPIScaler.scale(75)

        lbl_c = UIFactory.create_label("roostercap.")
        lbl_c.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_c)
        x += DPIScaler.scale(75)

        self.nud_capaciteit = self._maak_nud(x, y, ROOSTER_CAPACITEIT, 1, 50, 0.1, 1)
        pnl_settings.Controls.Add(self.nud_capaciteit)
        x += DPIScaler.scale(65)

        lbl_c2 = UIFactory.create_label("dm3/s per m1  |  verblijfsruimte")
        lbl_c2.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_c2)
        x += DPIScaler.scale(190)

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

        self.cmb_teksttype = UIFactory.create_combobox(170, [t[0] for t in self.teksttypes])
        self.cmb_teksttype.Location = Point(x, y - DPIScaler.scale(3))
        for i, (naam, _) in enumerate(self.teksttypes):
            if naam == DEFAULT_TEXT_TYPE:
                self.cmb_teksttype.SelectedIndex = i
                break
        pnl_settings.Controls.Add(self.cmb_teksttype)

        # Rij 2: daglicht + checkbox
        x = DPIScaler.scale(10)
        y = DPIScaler.scale(38)

        lbl_d = UIFactory.create_label("Daglicht:", bold=True)
        lbl_d.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_d)
        x += DPIScaler.scale(75)

        lbl_g = UIFactory.create_label("glasfactor")
        lbl_g.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_g)
        x += DPIScaler.scale(75)

        self.nud_glasfactor = self._maak_nud(x, y, GLASFACTOR, 0.1, 1, 0.05, 2)
        pnl_settings.Controls.Add(self.nud_glasfactor)
        x += DPIScaler.scale(65)

        lbl_g2 = UIFactory.create_label("x belemmering CbxCu")
        lbl_g2.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_g2)
        x += DPIScaler.scale(135)

        self.nud_belemmering = self._maak_nud(x, y, BELEMMERINGSFACTOR, 0.1, 1, 0.05, 2)
        pnl_settings.Controls.Add(self.nud_belemmering)
        x += DPIScaler.scale(75)

        lbl_g3 = UIFactory.create_label(
            "vereist: 10% vloeropp., min. 0,5 m2 per verblijfsruimte (vast)",
            italic=True, color=Huisstijl.TEXT_SECONDARY)
        lbl_g3.Location = Point(x, y)
        pnl_settings.Controls.Add(lbl_g3)
        x += DPIScaler.scale(360)

        self.chk_zonder_eis = CheckBox()
        self.chk_zonder_eis.Text = "Ruimtes zonder eis tonen"
        self.chk_zonder_eis.AutoSize = True
        self.chk_zonder_eis.Font = Font("Segoe UI", 9)
        self.chk_zonder_eis.Checked = True
        self.chk_zonder_eis.Location = Point(x, y - DPIScaler.scale(2))
        pnl_settings.Controls.Add(self.chk_zonder_eis)

        # Tabs met grids
        tabs_top = DPIScaler.scale(145)
        self.tabs = TabControl()
        self.tabs.Location = Point(DPIScaler.scale(10), tabs_top)
        self.tabs.Size = Size(self.ClientSize.Width - DPIScaler.scale(20),
                              self.ClientSize.Height - tabs_top - DPIScaler.scale(60))
        self.tabs.Font = Font("Segoe UI", 9)
        self.tabs.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom |
                            AnchorStyles.Left | AnchorStyles.Right)

        tab_vent = TabPage("Ventilatie")
        tab_vent.BackColor = Color.White
        self.grid_vent = self._maak_grid()
        self._kolommen_ventilatie(self.grid_vent)
        self.grid_vent.CellValueChanged += self._vent_grid_changed
        self.grid_vent.CurrentCellDirtyStateChanged += self._grid_dirty
        tab_vent.Controls.Add(self.grid_vent)
        self.tabs.TabPages.Add(tab_vent)

        tab_dag = TabPage("Daglicht")
        tab_dag.BackColor = Color.White
        self.grid_dag = self._maak_grid()
        self._kolommen_daglicht(self.grid_dag)
        self.grid_dag.CellValueChanged += self._dag_grid_changed
        self.grid_dag.CurrentCellDirtyStateChanged += self._grid_dirty
        tab_dag.Controls.Add(self.grid_dag)
        self.tabs.TabPages.Add(tab_dag)

        self.Controls.Add(self.tabs)

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
        for nud in [self.nud_capaciteit, self.nud_factor, self.nud_minimum,
                    self.nud_glasfactor, self.nud_belemmering]:
            nud.ValueChanged += self._instelling_changed

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

    def _maak_grid(self):
        grid = DataGridView()
        grid.Location = Point(DPIScaler.scale(5), DPIScaler.scale(5))
        grid.Size = Size(self.tabs.Width - DPIScaler.scale(20),
                         self.tabs.Height - DPIScaler.scale(40))
        grid.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom |
                       AnchorStyles.Left | AnchorStyles.Right)
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        grid.AllowUserToAddRows = False
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.BackgroundColor = Color.White
        grid.BorderStyle = BorderStyle.None
        grid.RowHeadersVisible = False
        grid.EnableHeadersVisualStyles = False
        grid.ColumnHeadersDefaultCellStyle.BackColor = Huisstijl.VIOLET
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        grid.ColumnHeadersDefaultCellStyle.Font = Font("Segoe UI", 9, FontStyle.Bold)
        grid.ColumnHeadersHeight = DPIScaler.scale(30)
        grid.RowTemplate.Height = DPIScaler.scale(26)
        grid.AlternatingRowsDefaultCellStyle.BackColor = Huisstijl.LIGHT_GRAY
        return grid

    def _tekst_kolom(self, grid, name, header, width, readonly=True):
        col = DataGridViewTextBoxColumn()
        col.Name = name
        col.HeaderText = header
        col.Width = DPIScaler.scale(width)
        col.ReadOnly = readonly
        grid.Columns.Add(col)

    def _kolommen_ventilatie(self, grid):
        self._tekst_kolom(grid, "nummer", "Nr", 45)
        self._tekst_kolom(grid, "naam", "Ruimte", 150)
        self._tekst_kolom(grid, "niveau", "Verdieping", 95)
        self._tekst_kolom(grid, "opp", "m²", 55)

        type_col = DataGridViewComboBoxColumn()
        type_col.Name = "type"
        type_col.HeaderText = "Type"
        type_col.Width = DPIScaler.scale(120)
        for key, _ in RUIMTE_TYPES:
            type_col.Items.Add(key)
        type_col.FlatStyle = FlatStyle.Flat
        grid.Columns.Add(type_col)

        self._tekst_kolom(grid, "toevoer", "Toevoer [dm³/s]", 95)
        self._tekst_kolom(grid, "rooster", "Ducoton 10", 100)
        self._tekst_kolom(grid, "afvoer", "Afvoer [dm³/s]", 95)
        self._tekst_kolom(grid, "voud", "Voud [1/h]", 75)

    def _kolommen_daglicht(self, grid):
        self._tekst_kolom(grid, "nummer", "Nr", 45)
        self._tekst_kolom(grid, "naam", "Ruimte", 150)
        self._tekst_kolom(grid, "niveau", "Verdieping", 95)
        self._tekst_kolom(grid, "opp", "m²", 55)
        self._tekst_kolom(grid, "type", "Type", 110)
        self._tekst_kolom(grid, "vereist", "Vereist Ae [m²]", 95)
        self._tekst_kolom(grid, "ramen", "Ramen", 55)
        self._tekst_kolom(grid, "raamopp", "Kozijnopp. [m²]", 95)
        self._tekst_kolom(grid, "aanwezig", "Aanwezig Ae [m²]", 100, readonly=False)
        self._tekst_kolom(grid, "toets", "Toets", 110)

    # -- berekening & weergave ------------------------------------------------
    def _herbereken(self):
        s = self.instellingen
        for r in self.ruimtes:
            r.bereken_ventilatie(s['factor'], s['minimum'], s['capaciteit'],
                                 ROOSTER_AFRONDING_MM)
            r.bereken_daglicht(s['glasfactor'], s['belemmering'])
        tot_t = sum(r.toevoer for r in self.ruimtes)
        tot_a = sum(r.afvoer for r in self.ruimtes)
        n_tekort = sum(1 for r in self.ruimtes if r.daglicht_toets.startswith('TEKORT'))
        self.lbl_status.Text = ("{} ruimtes | toevoer {} dm3/s | afvoer {} dm3/s | "
                                "daglicht: {} tekort").format(
            len(self.ruimtes), nl_getal(tot_t), nl_getal(tot_a), n_tekort)

    def _vul_grids(self):
        self._updating = True
        try:
            self.grid_vent.Rows.Clear()
            self.grid_dag.Rows.Clear()
            for r in self.ruimtes:
                idx = self.grid_vent.Rows.Add(
                    r.nummer, r.naam, r.niveau, nl_getal(r.opp), r.type_key,
                    nl_getal(r.toevoer) if r.toevoer > 0 else "-",
                    r.rooster_tekst,
                    nl_getal(r.afvoer) if r.afvoer > 0 else "-",
                    nl_getal(r.voud) if r.voud > 0 else "-")
                if r.type_key != detecteer_type(r.naam):
                    self.grid_vent.Rows[idx].Cells["type"].Style.BackColor = \
                        Color.FromArgb(230, 255, 230)

                aanwezig = "-"
                if r.ae_vereist > 0:
                    aanwezig = nl_getal(r.ae_aanwezig, 2)
                    if r.ae_override is not None:
                        aanwezig += "*"
                idx_d = self.grid_dag.Rows.Add(
                    r.nummer, r.naam, r.niveau, nl_getal(r.opp), r.type_key,
                    nl_getal(r.ae_vereist, 2) if r.ae_vereist > 0 else "-",
                    str(r.raam_aantal) if r.raam_aantal > 0 else "-",
                    nl_getal(r.raam_opp, 2) if r.raam_opp > 0 else "-",
                    aanwezig,
                    r.daglicht_toets)
                cel = self.grid_dag.Rows[idx_d].Cells["toets"]
                if r.daglicht_toets == 'OK':
                    cel.Style.BackColor = Huisstijl.TEAL
                    cel.Style.ForeColor = Color.White
                elif r.daglicht_toets.startswith('TEKORT'):
                    cel.Style.BackColor = Huisstijl.PEACH
                    cel.Style.ForeColor = Color.White
                if r.ae_override is not None:
                    self.grid_dag.Rows[idx_d].Cells["aanwezig"].Style.BackColor = \
                        Color.FromArgb(230, 255, 230)
        finally:
            self._updating = False

    def _instelling_changed(self, s, e):
        if self._updating:
            return
        self._herbereken()
        self._vul_grids()

    def _grid_dirty(self, s, e):
        grid = s
        if grid.IsCurrentCellDirty:
            grid.CommitEdit(1)

    def _vent_grid_changed(self, s, e):
        if self._updating or e.RowIndex < 0:
            return
        if e.ColumnIndex == self.grid_vent.Columns["type"].Index:
            idx = e.RowIndex
            if 0 <= idx < len(self.ruimtes):
                self.ruimtes[idx].type_key = \
                    str(self.grid_vent.Rows[idx].Cells["type"].Value)
                self._herbereken()
                self._vul_grids()

    def _dag_grid_changed(self, s, e):
        if self._updating or e.RowIndex < 0:
            return
        if e.ColumnIndex == self.grid_dag.Columns["aanwezig"].Index:
            idx = e.RowIndex
            if 0 <= idx < len(self.ruimtes):
                waarde = self.grid_dag.Rows[idx].Cells["aanwezig"].Value
                self.ruimtes[idx].ae_override = parse_getal(waarde)
                self._herbereken()
                self._vul_grids()

    # -- acties ----------------------------------------------------------------
    def _plaats_click(self, s, e):
        if not self.teksttypes:
            MessageBox.Show("Geen teksttypes in dit document gevonden.",
                            "Fout", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        idx = self.cmb_teksttype.SelectedIndex
        if idx < 0:
            idx = 0
        tekst = bouw_tabel_tekst(self.ruimtes, self.instellingen,
                                 self.chk_zonder_eis.Checked)
        self.plaats_data = {
            'tekst': tekst,
            'teksttype_id': self.teksttypes[idx][1],
            'teksttype_naam': self.teksttypes[idx][0],
            'instellingen': self.instellingen,
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


def plaats_tabel(plaats_data, ruimtes):
    """Vraag klikpunt op de actieve view en plaats de TextNote."""
    view = revit.doc.ActiveView
    if isinstance(view, DB.View3D):
        forms.alert("De actieve view is een 3D-view. Open een plattegrond, "
                    "drafting view of sheet en start de tool opnieuw.",
                    title="Bbl-Toets")
        return

    from Autodesk.Revit.Exceptions import OperationCanceledException
    try:
        punt = revit.uidoc.Selection.PickPoint("Klik het plaatsingspunt voor de Bbl-toets-tabel")
    except OperationCanceledException:
        return  # geannuleerd door gebruiker
    except Exception:
        # PickPoint niet mogelijk (bv. geen werkvlak) -> plaats in view-centrum
        punt = _view_centrum(view)

    with revit.Transaction("Bbl-toets-tabel plaatsen"):
        DB.TextNote.Create(revit.doc, view.Id, punt,
                           plaats_data['tekst'], plaats_data['teksttype_id'])

    schrijf_exchange_json(ruimtes, plaats_data['instellingen'])
    log.info("Bbl-toets-tabel geplaatst op view '{}' (teksttype {})".format(
        view.Name, plaats_data['teksttype_naam']))
    forms.alert("Bbl-toets-tabel geplaatst op view '{}'.\n\n"
                "Teksttype: {}\nTip: gebruik een monospace-teksttype als de "
                "kolommen niet uitlijnen.".format(view.Name, plaats_data['teksttype_naam']),
                title="Bbl-Toets")


# ==============================================================================
# ENTRY POINT
# ==============================================================================
if __name__ == '__main__':
    try:
        alle_ruimtes = laad_ruimtes()
        if not alle_ruimtes:
            forms.alert("Geen geplaatste ruimtes (Rooms) gevonden in het model.\n"
                        "Plaats eerst Rooms met een oppervlak.",
                        title="Bbl-Toets", exitscript=True)
        koppel_ramen(alle_ruimtes)
        form = BblToetsForm(alle_ruimtes)
        form.ShowDialog()
        if form.plaats_data:
            plaats_tabel(form.plaats_data, form.ruimtes)
    except Exception as e:
        log.error("BblToets fout: {}".format(e))
        forms.alert("Fout: {}".format(str(e)), title="Bbl-Toets")
