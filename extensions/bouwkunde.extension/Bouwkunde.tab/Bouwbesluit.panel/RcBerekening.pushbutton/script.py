# -*- coding: utf-8 -*-
"""Rc-waarde Berekening - Thermische weerstand van wanden incl. curtain walls"""

__title__ = "Rc-waarde\nBerekening"
__author__ = "3BM Bouwkunde"
__doc__ = "Bereken Rc en U-waarde van geselecteerde wanden"

from pyrevit import revit, DB, forms, script

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System.Net')

from System.Windows.Forms import Panel, BorderStyle
from System.Drawing import Point, Size, Color, Font, FontStyle, Pen, SolidBrush, StringFormat, RectangleF
from System.Drawing.Drawing2D import DashStyle, SmoothingMode
from System.Net import WebClient
from System.Text import Encoding
from System.IO import MemoryStream
from System.Drawing import Bitmap, Rectangle
from System.Drawing.Imaging import ImageFormat
from System.Collections.Generic import List as NetList
import math
import os
import json
import sys
import base64

# UI Template imports
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl, LayoutHelper
from bm_logger import get_logger

log = get_logger("RcBerekening")

# ==============================================================================
# CONSTANTEN
# ==============================================================================
SPOUW_RD = 0.17
RSI_DEFAULT = 0.13
RSE_DEFAULT = 0.04
FT_TO_M = 0.3048
FT_TO_MM = 304.8
MIN_SPOUW_GAP = 5.0
MEMBRANE_DEFAULT_MM = 0.2

T_BINNEN = 20.0
T_BUITEN = -10.0
RH_BINNEN = 50.0
RH_BUITEN = 90.0

# Uniforme breedte voor alle panels
PANEL_WIDTH = 940

# Nederlands klimaat - De Bilt referentiejaar (maandgemiddelden)
NL_KLIMAAT = {
    'maanden': ['Jan', 'Feb', 'Mrt', 'Apr', 'Mei', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dec'],
    'T_buiten': [3.1, 3.3, 6.2, 9.2, 13.1, 15.6, 17.9, 17.5, 14.5, 10.7, 6.7, 3.7],  # °C
    'RH_buiten': [87, 84, 81, 75, 75, 76, 77, 79, 83, 85, 89, 89],  # %
    'T_binnen': [20, 20, 20, 20, 20, 22, 23, 23, 22, 20, 20, 20],  # °C (iets hoger in zomer)
    'RH_binnen': [50, 50, 50, 50, 50, 55, 55, 55, 50, 50, 50, 50],  # %
}


# ==============================================================================
# MATERIAAL DATABASE
# ==============================================================================
def load_materiaal_database():
    """Laad materialen uit externe JSON database"""
    lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
    json_path = os.path.join(lib_path, 'materialen_database.json')
    
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r') as f:
                data = json.load(f)
                return data.get('materialen', [])
        except:
            pass
    return []

# Lazy loading - database wordt pas geladen bij gebruik
MATERIAAL_DB = None

def get_materiaal_db():
    """Lazy loader voor materiaal database"""
    global MATERIAAL_DB
    if MATERIAAL_DB is None:
        MATERIAAL_DB = load_materiaal_database()
    return MATERIAAL_DB


def get_db_tuples():
    """Retourneer materialen als tuples voor dropdown: (cat, naam, lambda, rd, mu)"""
    db = get_materiaal_db()
    if db:
        return [(m.get('categorie',''), m.get('naam',''), m.get('lambda'), 
                 m.get('rd_vast'), m.get('mu', 1)) for m in db]
    return [("Spouw", "Spouw Rd=0.17", None, 0.17, 1)]


def match_material_by_name(name):
    """Zoek materiaal op basis van keywords"""
    db = get_materiaal_db()
    if not name or not db:
        return None
    name_lower = name.lower()
    best, best_score = None, 0
    for m in db:
        score = sum(len(kw) for kw in m.get('keywords', []) if kw.lower() in name_lower)
        if score > best_score:
            best, best_score = m, score
    return best if best_score >= 3 else None


# ==============================================================================
# REVIT HELPERS
# ==============================================================================
def psat(T):
    """Verzadigingsdruk waterdamp (Magnus formule)"""
    if T >= 0:
        return 611.2 * math.exp((17.62 * T) / (243.12 + T))
    return 611.2 * math.exp((22.46 * T) / (272.62 + T))


def get_thermal_conductivity(material):
    if not material:
        return None
    thermal_asset_id = material.ThermalAssetId
    if not thermal_asset_id or thermal_asset_id == DB.ElementId.InvalidElementId:
        return None
    thermal_asset = revit.doc.GetElement(thermal_asset_id)
    if not thermal_asset:
        return None
    try:
        tc_param = thermal_asset.get_Parameter(DB.BuiltInParameter.PHY_MATERIAL_PARAM_THERMAL_CONDUCTIVITY)
        if tc_param and tc_param.HasValue:
            raw = tc_param.AsDouble()
            try:
                return DB.UnitUtils.ConvertFromInternalUnits(raw, DB.UnitTypeId.WattsPerMeterKelvin)
            except:
                return raw
    except:
        pass
    return None


def get_mu_from_material(material):
    """Lees Mu parameter uit Revit materiaal (Number - dimensieloos)"""
    if not material:
        return None
    try:
        mu_param = material.LookupParameter("Mu")
        if mu_param and mu_param.HasValue:
            if mu_param.StorageType == DB.StorageType.Double:
                # Number parameter: direct uitlezen, geen conversie
                return int(round(mu_param.AsDouble()))
            elif mu_param.StorageType == DB.StorageType.Integer:
                return mu_param.AsInteger()
    except:
        pass
    return None


def get_wall_layers(wall):
    """Haal lagen op uit Revit wand"""
    layers = []
    wall_type = wall.WallType
    compound = wall_type.GetCompoundStructure()
    if not compound:
        return layers
    
    wall_type_name = wall_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    
    for layer in compound.GetLayers():
        width_m = layer.Width * FT_TO_M
        width_mm = width_m * 1000
        mat_id = layer.MaterialId
        mat_name, lambda_val, mu_val = "Onbekend", None, None
        is_membrane = width_mm < 0.1
        
        if mat_id and mat_id != DB.ElementId.InvalidElementId:
            mat = revit.doc.GetElement(mat_id)
            if mat:
                mat_name = mat.Name
                lambda_val = get_thermal_conductivity(mat)
                mu_val = get_mu_from_material(mat)
                
                # Auto-match als waarden ontbreken
                if lambda_val is None or mu_val is None:
                    db_match = match_material_by_name(mat_name)
                    if db_match:
                        lambda_val = lambda_val or db_match.get('lambda')
                        mu_val = mu_val or db_match.get('mu')
        
        if is_membrane:
            width_mm = MEMBRANE_DEFAULT_MM
        
        r_val = (width_m / lambda_val) if lambda_val and lambda_val > 0 and not is_membrane else None
        
        layers.append({
            'wall_type': wall_type_name, 'material': mat_name, 'width_mm': width_mm,
            'lambda': lambda_val, 'r_value': r_val, 'mu': mu_val,
            'is_air_gap': False, 'is_membrane': is_membrane
        })
    return layers


def get_curtain_panel_layers(panel):
    """Haal lagen op uit een curtain wall panel"""
    layers = []
    
    # Haal panel type op
    panel_type_id = panel.GetTypeId()
    if not panel_type_id or panel_type_id == DB.ElementId.InvalidElementId:
        return layers
    
    panel_type = revit.doc.GetElement(panel_type_id)
    if not panel_type:
        return layers
    
    type_name = panel_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    type_name = type_name.AsString() if type_name else "Panel"
    
    # Probeer compound structure (sommige panel families hebben dit)
    try:
        compound = panel_type.GetCompoundStructure()
        if compound:
            for layer in compound.GetLayers():
                width_m = layer.Width * FT_TO_M
                width_mm = width_m * 1000
                mat_id = layer.MaterialId
                mat_name, lambda_val, mu_val = "Onbekend", None, None
                is_membrane = width_mm < 0.1
                
                if mat_id and mat_id != DB.ElementId.InvalidElementId:
                    mat = revit.doc.GetElement(mat_id)
                    if mat:
                        mat_name = mat.Name
                        lambda_val = get_thermal_conductivity(mat)
                        mu_val = get_mu_from_material(mat)
                        
                        if lambda_val is None or mu_val is None:
                            db_match = match_material_by_name(mat_name)
                            if db_match:
                                lambda_val = lambda_val or db_match.get('lambda')
                                mu_val = mu_val or db_match.get('mu')
                
                if is_membrane:
                    width_mm = MEMBRANE_DEFAULT_MM
                
                r_val = (width_m / lambda_val) if lambda_val and lambda_val > 0 and not is_membrane else None
                
                layers.append({
                    'wall_type': type_name, 'material': mat_name, 'width_mm': width_mm,
                    'lambda': lambda_val, 'r_value': r_val, 'mu': mu_val,
                    'is_air_gap': False, 'is_membrane': is_membrane
                })
            return layers
    except:
        pass
    
    # Fallback: lees dikte en materiaal direct van panel
    try:
        # Dikte parameter
        thickness_param = panel_type.LookupParameter("Thickness") or panel_type.LookupParameter("Dikte")
        if not thickness_param:
            thickness_param = panel_type.get_Parameter(DB.BuiltInParameter.CURTAIN_WALL_SYSPANEL_THICKNESS)
        
        width_mm = 100  # default
        if thickness_param and thickness_param.HasValue:
            if thickness_param.StorageType == DB.StorageType.Double:
                width_mm = thickness_param.AsDouble() * FT_TO_MM
            else:
                try:
                    width_mm = float(thickness_param.AsValueString().replace('mm', '').replace(',', '.').strip())
                except:
                    pass
        
        # Materiaal
        mat_id = panel_type.get_Parameter(DB.BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if not mat_id:
            mat_id = panel_type.LookupParameter("Material") or panel_type.LookupParameter("Materiaal")
        
        mat_name, lambda_val, mu_val = type_name, None, None
        
        if mat_id and mat_id.HasValue:
            mat_elem_id = mat_id.AsElementId()
            if mat_elem_id and mat_elem_id != DB.ElementId.InvalidElementId:
                mat = revit.doc.GetElement(mat_elem_id)
                if mat:
                    mat_name = mat.Name
                    lambda_val = get_thermal_conductivity(mat)
                    mu_val = get_mu_from_material(mat)
        
        # Auto-match op type naam als geen materiaal gevonden
        if lambda_val is None or mu_val is None:
            db_match = match_material_by_name(type_name)
            if db_match:
                lambda_val = lambda_val or db_match.get('lambda')
                mu_val = mu_val or db_match.get('mu')
        
        width_m = width_mm / 1000.0
        r_val = (width_m / lambda_val) if lambda_val and lambda_val > 0 else None
        
        layers.append({
            'wall_type': "CW: " + type_name, 'material': mat_name, 'width_mm': width_mm,
            'lambda': lambda_val, 'r_value': r_val, 'mu': mu_val,
            'is_air_gap': False, 'is_membrane': False
        })
    except:
        pass
    
    return layers


def get_curtain_wall_panels(curtain_wall):
    """Haal alle unieke panels uit een curtain wall"""
    panels = []
    try:
        # Haal curtain grid
        cw_grid = curtain_wall.CurtainGrid
        if not cw_grid:
            return panels
        
        # Haal panel IDs
        panel_ids = cw_grid.GetPanelIds()
        
        seen_types = set()
        for panel_id in panel_ids:
            panel = revit.doc.GetElement(panel_id)
            if not panel:
                continue
            
            # Skip mullions en andere elementen
            if not isinstance(panel, DB.Panel) and not isinstance(panel, DB.FamilyInstance):
                continue
            
            # Check of het een panel is (niet een deur/raam in de CW)
            cat = panel.Category
            if cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_CurtainWallPanels):
                type_id = panel.GetTypeId()
                if type_id not in seen_types:
                    seen_types.add(type_id)
                    panels.append(panel)
    except:
        pass
    
    return panels


# ==============================================================================
# DATA KLASSE
# ==============================================================================
class LaagData:
    def __init__(self, wall_type, material, width_mm, lambda_val, r_value, mu=None, is_air_gap=False, is_membrane=False):
        self.wall_type = wall_type
        self.material = material
        self.width_mm = width_mm
        self.lambda_val = lambda_val
        self.r_value = r_value
        self.mu = mu
        self.is_air_gap = is_air_gap
        self.is_membrane = is_membrane
    
    @property
    def sd(self):
        return self.mu * (self.width_mm / 1000.0) if self.mu and self.width_mm else None


# ==============================================================================
# VOCHTBALANS BEREKENING
# ==============================================================================
def calculate_monthly_moisture(lagen, rsi, rse):
    """
    Bereken maandelijkse vochtaccumulatie met vereenvoudigde Glaser methode.
    Returns: tuple (monthly_change, cumulative) - beide lijsten van 12 waarden
    """
    if not lagen:
        return [0] * 12, [0] * 12
    
    r_vals = [l.r_value or 0 for l in lagen]
    sd_vals = [l.sd or 0 for l in lagen]
    r_tot = rsi + sum(r_vals) + rse
    sd_tot = sum(sd_vals)
    
    if r_tot == 0 or sd_tot == 0:
        return [0] * 12, [0] * 12
    
    monthly_change = []
    
    for month in range(12):
        t_in = NL_KLIMAAT['T_binnen'][month]
        t_out = NL_KLIMAAT['T_buiten'][month]
        rh_in = NL_KLIMAAT['RH_binnen'][month]
        rh_out = NL_KLIMAAT['RH_buiten'][month]
        
        # Temperatuurverloop berekenen
        dt = t_in - t_out
        temps = [t_in]
        r_cum = rsi
        temps.append(t_in - (r_cum / r_tot) * dt)
        for r in r_vals:
            r_cum += r
            temps.append(t_in - (r_cum / r_tot) * dt)
        
        # Dampdrukverloop
        p_in = psat(t_in) * (rh_in / 100.0)
        p_out = psat(t_out) * (rh_out / 100.0)
        
        pvap = [p_in, p_in]
        sd_cum = 0
        for sd in sd_vals:
            sd_cum += sd
            pvap.append(p_in - (sd_cum / sd_tot) * (p_in - p_out))
        
        psat_vals = [psat(t) for t in temps]
        
        # Bereken maximale overschrijding OF droogpotentieel
        max_excess = None
        min_margin = None
        
        for i in range(1, len(pvap)):
            if i < len(psat_vals):
                excess = pvap[i] - psat_vals[i]
                if max_excess is None or excess > max_excess:
                    max_excess = excess
                if min_margin is None or excess < min_margin:
                    min_margin = excess
        
        if max_excess is None:
            max_excess = 0
        if min_margin is None:
            min_margin = 0
        
        # Condensatie of droging berekenen
        if max_excess > 0:
            # Er is condensatie: vocht accumuleert
            change = max_excess * 0.1
        else:
            # Geen condensatie: droogpotentieel (min_margin is negatief)
            change = min_margin * 0.03
        
        monthly_change.append(change)
    
    # Cumulatief berekenen - ZONDER max(0) zodat droging zichtbaar is
    cumulative = []
    total = 0
    for mc in monthly_change:
        total += mc
        # Minimaal 0: kan niet meer drogen dan er aan vocht is
        if total < 0:
            total = 0
        cumulative.append(total)
    
    return monthly_change, cumulative


# ==============================================================================
# PDF RAPPORT GENERATIE
# ==============================================================================
REPORT_API_BASE = os.environ.get("BM_REPORT_API", "https://report.3bm.co.nl")
REPORT_API_KEY = os.environ.get("BM_REPORT_API_KEY", "")

# API key cache (blijft geldig zolang pyRevit draait)
_api_key = None


def _get_api_key():
    """Haal de API key op voor de Report API.

    Gebruikt BM_REPORT_API_KEY environment variable.
    Als die niet gezet is, vraagt een prompt om de key.

    Returns:
        API key string

    Raises:
        Exception bij geannuleerde invoer
    """
    global _api_key
    if _api_key:
        return _api_key

    key = REPORT_API_KEY

    if not key:
        key = forms.ask_for_string(
            prompt="API key voor report.3bm.co.nl (begint met 3bm_k_):",
            title="API Key"
        )
        if not key:
            raise Exception("API key invoer geannuleerd")

    _api_key = key
    return _api_key


def capture_3d_view(view_name="3D IFC export"):
    """Exporteer een 3D view als base64 PNG voor gebruik in rapport.

    Args:
        view_name: naam van de 3D view (zoekt op exacte match, daarna contains)

    Returns:
        dict {"data": "<base64>", "media_type": "image/png"} of None
    """
    try:
        views = DB.FilteredElementCollector(revit.doc) \
            .OfClass(DB.View3D) \
            .ToElements()

        view_3d = None
        # Zoek op exacte naam
        for v in views:
            if not v.IsTemplate and v.Name == view_name:
                view_3d = v
                break
        # Fallback: zoek op gedeeltelijke naam
        if not view_3d:
            for v in views:
                if not v.IsTemplate and view_name.lower() in v.Name.lower():
                    view_3d = v
                    break
        # Fallback: eerste beschikbare 3D view
        if not view_3d:
            for v in views:
                if not v.IsTemplate:
                    view_3d = v
                    break
        if not view_3d:
            return None

        # Export naar temp bestand
        import tempfile
        temp_dir = tempfile.gettempdir()
        temp_base = "rc_3d_export"
        temp_path = os.path.join(temp_dir, temp_base)

        # Verwijder eventueel oud bestand
        for f in os.listdir(temp_dir):
            if f.startswith(temp_base) and f.endswith('.png'):
                try:
                    os.remove(os.path.join(temp_dir, f))
                except:
                    pass

        opts = DB.ImageExportOptions()
        opts.ExportRange = DB.ExportRange.SetOfViews
        view_ids = NetList[DB.ElementId]()
        view_ids.Add(view_3d.Id)
        opts.SetViewsAndSheets(view_ids)
        opts.FilePath = temp_path
        opts.HLRandWFViewsFileType = DB.ImageFileType.PNG
        opts.ShadowViewsFileType = DB.ImageFileType.PNG
        opts.ImageResolution = DB.ImageResolution.DPI_150
        opts.ZoomType = DB.ZoomFitType.FitToPage
        opts.PixelSize = 1200

        revit.doc.ExportImage(opts)

        # Zoek het geexporteerde bestand (Revit voegt soms suffix toe)
        actual_path = None
        for f in os.listdir(temp_dir):
            if f.startswith(temp_base) and f.endswith('.png'):
                actual_path = os.path.join(temp_dir, f)
                break

        if not actual_path or not os.path.exists(actual_path):
            return None

        with open(actual_path, 'rb') as f:
            data = base64.b64encode(f.read()).decode('ascii')

        try:
            os.remove(actual_path)
        except:
            pass

        return {"data": data, "media_type": "image/png"}
    except:
        return None


def capture_panel_image(panel):
    """Capture een WinForms Panel als base64 PNG.

    Args:
        panel: WinForms Panel/Control met OnPaint

    Returns:
        dict {"data": "<base64>", "media_type": "image/png"} of None
    """
    try:
        bmp = Bitmap(panel.Width, panel.Height)
        panel.DrawToBitmap(bmp, Rectangle(0, 0, panel.Width, panel.Height))

        stream = MemoryStream()
        bmp.Save(stream, ImageFormat.Png)
        data = base64.b64encode(bytes(bytearray(stream.ToArray()))).decode('ascii')
        stream.Dispose()
        bmp.Dispose()

        return {"data": data, "media_type": "image/png"}
    except:
        return None


def generate_pdf_report(lagen, rsi, rse, ti, te, rhi, rhe, rc_total, u_value,
                        monthly_change, cumulative, output_path,
                        cover_image=None, diagram_image=None):
    """Genereer PDF rapport via 3BM Report Generator API.

    Args:
        lagen: lijst van LaagData objecten
        rsi, rse: oppervlakteweerstanden
        ti, te: temperaturen binnen/buiten
        rhi, rhe: relatieve vochtigheid binnen/buiten
        rc_total: berekende Rc-waarde
        u_value: berekende U-waarde
        monthly_change: maandelijkse vochtverandering (12 waarden)
        cumulative: cumulatieve vochtbalans (12 waarden)
        output_path: pad voor PDF bestand

    Returns:
        output_path bij succes

    Raises:
        Exception bij API fout
    """
    # Project info uit Revit
    proj = revit.doc.ProjectInformation
    proj_naam = proj.Name if proj.Name else "Onbekend project"
    proj_nummer = proj.Number if proj.Number else ""
    proj_adres = proj.Address if proj.Address else ""

    # Wandtype uit eerste laag
    wand_type = lagen[0].wall_type if lagen else "Onbekend"

    # Sectie 1: Uitgangspunten
    sectie_uitgangspunten = {
        "number": "1",
        "title": "Uitgangspunten",
        "level": 1,
        "content": [
            {
                "type": "paragraph",
                "text": "Klimaatcondities conform NEN-EN ISO 6946 / NEN 2778."
            },
            {
                "type": "table",
                "title": "Klimaatcondities",
                "headers": ["Parameter", "Waarde", "Eenheid"],
                "rows": [
                    ["Rsi (binnen)", "{:.2f}".format(float(rsi)), u"m\u00b2K/W"],
                    ["Rse (buiten)", "{:.2f}".format(float(rse)), u"m\u00b2K/W"],
                    ["Ti (temperatuur binnen)", "{:.1f}".format(float(ti)), u"\u00b0C"],
                    ["Te (temperatuur buiten)", "{:.1f}".format(float(te)), u"\u00b0C"],
                    ["RVi (rel. vochtigheid binnen)", "{:.0f}".format(float(rhi)), "%"],
                    ["RVe (rel. vochtigheid buiten)", "{:.0f}".format(float(rhe)), "%"],
                ]
            }
        ]
    }

    # Sectie 2: Wandopbouw
    laag_rows = []
    for i, l in enumerate(lagen):
        lam_str = "(spouw)" if l.is_air_gap else ("{:.4f}".format(float(l.lambda_val)) if l.lambda_val else "-")
        r_str = "{:.3f}".format(float(l.r_value)) if l.r_value else "-"
        mu_str = str(l.mu) if l.mu else "-"
        sd_str = "{:.2f}".format(float(l.sd)) if l.sd else "-"
        mat_name = l.material if len(l.material) <= 40 else l.material[:37] + "..."
        laag_rows.append([
            str(i + 1), mat_name, "{:.1f}".format(float(l.width_mm)),
            lam_str, r_str, mu_str, sd_str
        ])

    wandopbouw_content = [
        {
            "type": "table",
            "title": "Lagenopbouw (binnen naar buiten)",
            "headers": ["#", "Materiaal", "Dikte [mm]",
                        u"\u03bb [W/mK]", u"R [m\u00b2K/W]",
                        u"\u03bc [-]", "Sd [m]"],
            "rows": laag_rows,
            "style": "striped",
            "column_widths": [8, 48, 18, 22, 24, 14, 18]
        }
    ]

    if diagram_image:
        wandopbouw_content.append({"type": "spacer", "height_mm": 5})
        wandopbouw_content.append({
            "type": "image",
            "src": diagram_image,
            "caption": "Temperatuur- en dampspanningsverloop door de constructie",
            "width_mm": 160,
            "alignment": "center"
        })

    sectie_wandopbouw = {
        "number": "2",
        "title": "Wandopbouw",
        "level": 1,
        "content": wandopbouw_content
    }

    # Sectie 3: Resultaat
    dikte_totaal = sum(l.width_mm for l in lagen) or 0.0
    sd_totaal = sum(l.sd for l in lagen if l.sd) or 0.0
    missing = [l for l in lagen if not l.lambda_val and not l.r_value and not l.is_membrane]

    rc_label = "Rc" if not missing else "Rc (indicatief)"
    rc_eis = 3.5
    voldoet = rc_total >= rc_eis and not missing

    sectie_resultaat = {
        "number": "3",
        "title": "Resultaat",
        "level": 1,
        "content": [
            {
                "type": "calculation",
                "title": "Totale dikte",
                "formula": "d_tot = som(d_i)",
                "result": "{:.0f}".format(float(dikte_totaal)),
                "unit": "mm"
            },
            {
                "type": "calculation",
                "title": "Totale dampdiffusieweerstand",
                "formula": "Sd_tot = som(mu_i * d_i)",
                "result": "{:.2f}".format(float(sd_totaal)),
                "unit": "m"
            },
            {
                "type": "calculation",
                "title": rc_label,
                "formula": "Rc = Rsi + som(d_i / lambda_i) + Rse",
                "substitution": "Rc = {:.2f} + som(R_i) + {:.2f}".format(float(rsi), float(rse)),
                "result": "{:.2f}".format(float(rc_total)),
                "unit": u"m\u00b2K/W",
                "reference": "NEN-EN ISO 6946"
            },
            {
                "type": "calculation",
                "title": "U-waarde",
                "formula": "U = 1 / Rc",
                "substitution": "U = 1 / {:.2f}".format(float(rc_total)) if rc_total > 0 else "U = -",
                "result": "{:.3f}".format(float(u_value)) if u_value else "-",
                "unit": u"W/m\u00b2K"
            },
            {
                "type": "check",
                "description": "Toets Bouwbesluit thermische isolatie",
                "required_value": u"Rc \u2265 {:.1f} m\u00b2K/W".format(float(rc_eis)),
                "calculated_value": "Rc = {:.2f}".format(float(rc_total)),
                "unity_check": round(rc_eis / rc_total, 2) if rc_total > 0 else 99,
                "limit": 1.0,
                "result": "VOLDOET" if voldoet else "VOLDOET NIET",
                "reference": "Bouwbesluit 2012, art. 5.3"
            }
        ]
    }

    if missing:
        sectie_resultaat["content"].insert(0, {
            "type": "paragraph",
            "text": "Let op: {} laag/lagen hebben geen lambda-waarde. "
                    "De Rc-waarde is indicatief.".format(len(missing))
        })

    # Sectie 4: Vochtbalans
    maanden = NL_KLIMAAT['maanden']
    vocht_rows = []
    for i in range(12):
        t_out = NL_KLIMAAT['T_buiten'][i]
        rh_out = NL_KLIMAAT['RH_buiten'][i]
        t_in = NL_KLIMAAT['T_binnen'][i]
        rh_in = NL_KLIMAAT['RH_binnen'][i]
        vocht_rows.append([
            maanden[i],
            "{:.1f}".format(float(t_in)), "{:.0f}".format(float(rh_in)),
            "{:.1f}".format(float(t_out)), "{:.0f}".format(float(rh_out)),
            "{:.1f}".format(float(monthly_change[i])),
            "{:.1f}".format(float(cumulative[i]))
        ])

    end_val = float(cumulative[-1]) if cumulative else 0.0
    risico_tekst = ""
    if end_val > 500:
        risico_tekst = "WAARSCHUWING: Hoog condensatierisico! Cumulatief vocht na 12 maanden: {:.0f} g/m2.".format(end_val)
    elif end_val > 0:
        risico_tekst = "Er treedt beperkte condensatie op ({:.0f} g/m2 na 12 maanden). De constructie droogt niet volledig.".format(end_val)
    else:
        risico_tekst = "De constructie droogt volledig uit. Geen condensatierisico."

    sectie_vochtbalans = {
        "number": "4",
        "title": "Vochtbalans (Glaser)",
        "level": 1,
        "content": [
            {
                "type": "paragraph",
                "text": "Vereenvoudigde Glaser analyse op basis van Nederlands klimaat (De Bilt referentiejaar)."
            },
            {
                "type": "table",
                "title": u"Maandelijkse vochtbalans (g/m\u00b2)",
                "headers": [
                    "Maand",
                    u"Ti [\u00b0C]", "RVi [%]",
                    u"Te [\u00b0C]", "RVe [%]",
                    u"Maand [g/m\u00b2]", u"Cum. [g/m\u00b2]"
                ],
                "rows": vocht_rows,
                "style": "striped",
                "column_widths": [18, 18, 18, 18, 18, 28, 28]
            },
            {
                "type": "paragraph",
                "text": risico_tekst
            }
        ]
    }

    # Rapport JSON samenstellen
    cover_data = {
        "enabled": True,
        "subtitle": u"Rc-waarde berekening \u2014 thermische weerstand en vochtbalans",
        "extra_fields": {}
    }
    if cover_image:
        cover_data["image"] = cover_image

    rapport_data = {
        "project": wand_type,
        "project_number": proj_nummer,
        "report_type": "Rc-waarde berekening",
        "author": "3BM Bouwkunde",
        "status": "CONCEPT",
        "brand": "3bm_cooperatie",
        "cover": cover_data,
        "colofon": {"enabled": False},
        "toc": {"enabled": False},
        "backcover": {"enabled": False},
        "sections": [
            sectie_uitgangspunten,
            sectie_wandopbouw,
            sectie_resultaat,
            sectie_vochtbalans
        ]
    }

    if proj_naam:
        rapport_data["cover"]["extra_fields"]["Project"] = proj_naam
    if proj_adres:
        rapport_data["cover"]["extra_fields"]["Adres"] = proj_adres

    # POST naar API via System.Net.WebClient (met API key)
    json_str = json.dumps(rapport_data, ensure_ascii=False, indent=2)
    json_bytes = Encoding.UTF8.GetBytes(json_str)

    # Debug: schrijf JSON naar bestand naast de PDF
    debug_path = output_path.replace('.pdf', '_debug.json')
    with open(debug_path, 'wb') as f:
        f.write(json_str.encode('utf-8'))

    # API key ophalen
    api_key = _get_api_key()

    generate_url = "{}/api/generate/v2".format(REPORT_API_BASE)
    client = WebClient()
    client.Headers.Add("Content-Type", "application/json; charset=utf-8")
    client.Headers.Add("X-API-Key", api_key)
    response_bytes = client.UploadData(generate_url, "POST", json_bytes)

    # Controleer content type
    content_type = client.ResponseHeaders.Get("Content-Type") or ""
    if "application/pdf" not in content_type:
        error_text = Encoding.UTF8.GetString(response_bytes)
        # Bij 401 (ongeldige key): key resetten en opnieuw proberen
        if "401" in error_text or "Niet ingelogd" in error_text or "verlopen" in error_text:
            global _api_key
            _api_key = None
            raise Exception("Ongeldige API key — probeer opnieuw")
        raise Exception("API fout: {}".format(error_text))

    # PDF opslaan
    with open(output_path, 'wb') as f:
        f.write(bytes(bytearray(response_bytes)))

    return output_path


# ==============================================================================
# WANDOPBOUW PANEL (Custom Paint)
# ==============================================================================
class WandOpbouwPanel(Panel):
    def __init__(self):
        self.lagen = []
        self.rsi, self.rse = RSI_DEFAULT, RSE_DEFAULT
        self.t_binnen, self.t_buiten = T_BINNEN, T_BUITEN
        self.rh_binnen, self.rh_buiten = RH_BINNEN, RH_BUITEN
        self.BackColor = Color.White
        self.BorderStyle = BorderStyle.None
    
    def set_data(self, lagen, rsi, rse, ti, te, rhi, rhe):
        self.lagen = lagen
        self.rsi, self.rse = rsi, rse
        self.t_binnen, self.t_buiten = ti, te
        self.rh_binnen, self.rh_buiten = rhi, rhe
        self.Invalidate()
    
    def _calc_glaser(self):
        if not self.lagen:
            return [], [], []
        r_vals = [l.r_value or 0 for l in self.lagen]
        sd_vals = [l.sd or 0 for l in self.lagen]
        r_tot = self.rsi + sum(r_vals) + self.rse
        if r_tot == 0:
            return [], [], []
        
        dt = self.t_binnen - self.t_buiten
        temps = [self.t_binnen]
        r_cum = self.rsi
        temps.append(self.t_binnen - (r_cum / r_tot) * dt)
        for r in r_vals:
            r_cum += r
            temps.append(self.t_binnen - (r_cum / r_tot) * dt)
        
        psat_vals = [psat(t) for t in temps]
        p_in = psat(self.t_binnen) * (self.rh_binnen / 100.0)
        p_out = psat(self.t_buiten) * (self.rh_buiten / 100.0)
        
        sd_tot = sum(sd_vals)
        if sd_tot == 0:
            pvap_vals = [p_in] + [p_in - ((i+1)/(len(self.lagen)+1))*(p_in-p_out) for i in range(len(self.lagen)+1)]
        else:
            pvap_vals = [p_in, p_in]
            sd_cum = 0
            for sd in sd_vals:
                sd_cum += sd
                pvap_vals.append(p_in - (sd_cum / sd_tot) * (p_in - p_out))
        
        return temps, psat_vals, pvap_vals
    
    def OnPaint(self, e):
        g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        if not self.lagen:
            return
        
        w_draw, m_top, m_bot = 600, 25, 25
        h_draw = self.Height - m_top - m_bot
        x_start = (self.Width - w_draw) / 2
        
        # Labels BINNEN/BUITEN
        font_lbl = Font("Segoe UI", 10, FontStyle.Bold)
        brush_txt = SolidBrush(Huisstijl.VIOLET)
        sf = StringFormat()
        sf.Alignment = sf.Alignment.Center
        
        state = g.Save()
        g.TranslateTransform(float(x_start - 15), float(self.Height / 2))
        g.RotateTransform(90)
        g.DrawString("BINNEN", font_lbl, brush_txt, 0, -10, sf)
        g.Restore(state)
        
        state = g.Save()
        g.TranslateTransform(float(x_start + w_draw + 15), float(self.Height / 2))
        g.RotateTransform(90)
        g.DrawString("BUITEN", font_lbl, brush_txt, 0, -10, sf)
        g.Restore(state)
        
        # Lagen tekenen
        tot_d = sum(l.width_mm for l in self.lagen)
        if tot_d == 0:
            return
        
        positions = []
        cur_x = x_start
        for l in self.lagen:
            w = max(25, int((l.width_mm / tot_d) * w_draw))
            positions.append((int(cur_x), w))
            cur_x += w
        
        pen_border = Pen(Color.FromArgb(100, 100, 100), 1)
        pen_membrane = Pen(Color.FromArgb(200, 50, 50), 3)
        pen_membrane.DashStyle = DashStyle.Dash
        
        for i, laag in enumerate(self.lagen):
            lx, lw = positions[i]
            nr = i + 1
            
            if laag.is_membrane:
                line_x = lx + lw / 2
                g.DrawLine(pen_membrane, int(line_x), m_top, int(line_x), m_top + h_draw)
                font_nr = Font("Segoe UI", 9, FontStyle.Bold)
                g.DrawString(str(nr), font_nr, SolidBrush(Huisstijl.MAGENTA), int(line_x) - 5, m_top + 5)
            else:
                color = Huisstijl.get_material_color(laag.material)
                g.FillRectangle(SolidBrush(color), lx, m_top, lw, h_draw)
                g.DrawRectangle(pen_border, lx, m_top, lw, h_draw)
                
                # Nummer in cirkel
                cx, cy, cs = lx + lw/2 - 12, m_top + h_draw/2 - 12, 24
                g.FillEllipse(SolidBrush(Color.White), int(cx), int(cy), cs, cs)
                g.DrawEllipse(pen_border, int(cx), int(cy), cs, cs)
                
                font_nr = Font("Segoe UI", 12, FontStyle.Bold)
                sf_c = StringFormat()
                sf_c.Alignment = sf_c.Alignment.Center
                sf_c.LineAlignment = sf_c.LineAlignment.Center
                g.DrawString(str(nr), font_nr, SolidBrush(Huisstijl.TEXT_PRIMARY), 
                            RectangleF(float(cx), float(cy), float(cs), float(cs)), sf_c)
                
                # Dikte onderaan
                if lw >= 35:
                    font_d = Font("Segoe UI", 7)
                    sf_d = StringFormat()
                    sf_d.Alignment = sf_d.Alignment.Center
                    g.DrawString("{:.0f}".format(laag.width_mm), font_d, SolidBrush(Huisstijl.DARK_GRAY),
                                RectangleF(float(lx), float(m_top + h_draw - 18), float(lw), float(16)), sf_d)
        
        # Glaser grafieken
        temps, psat_v, pvap_v = self._calc_glaser()
        if len(temps) < 2:
            return
        
        bounds = [int(x_start)] + [lx + lw for lx, lw in positions]
        
        t_min, t_max = min(temps), max(temps)
        t_rng = t_max - t_min if t_max != t_min else 1
        def t2y(t):
            return int(m_top + h_draw - ((t - t_min) / t_rng) * h_draw * 0.8 - h_draw * 0.1)
        
        all_p = psat_v + pvap_v
        p_min, p_max = min(all_p), max(all_p)
        p_rng = p_max - p_min if p_max != p_min else 1
        def p2y(p):
            return int(m_top + h_draw - ((p - p_min) / p_rng) * h_draw * 0.8 - h_draw * 0.1)
        
        # Temperatuur lijn
        pen_t = Pen(Huisstijl.GRAPH_TEMP, 3)
        pts_t = [Point(bounds[i], t2y(temps[i+1])) for i in range(len(bounds)) if i+1 < len(temps)]
        for i in range(len(pts_t) - 1):
            g.DrawLine(pen_t, pts_t[i], pts_t[i + 1])
        
        # Psat lijn
        pen_ps = Pen(Huisstijl.GRAPH_PSAT, 2)
        pts_ps = [Point(bounds[i], p2y(psat_v[i+1])) for i in range(len(bounds)) if i+1 < len(psat_v)]
        for i in range(len(pts_ps) - 1):
            g.DrawLine(pen_ps, pts_ps[i], pts_ps[i + 1])
        
        # Pvap lijn (gestreept)
        pen_pv = Pen(Huisstijl.GRAPH_PVAP, 2)
        pen_pv.DashStyle = DashStyle.Dash
        pts_pv = [Point(bounds[i], p2y(pvap_v[i+1])) for i in range(len(bounds)) if i+1 < len(pvap_v)]
        for i in range(len(pts_pv) - 1):
            g.DrawLine(pen_pv, pts_pv[i], pts_pv[i + 1])
        
        # Condensatie indicatoren
        for i in range(min(len(pts_ps), len(pts_pv))):
            if pts_pv[i].Y < pts_ps[i].Y - 3:
                zone_x = bounds[i]
                y_top, y_bot = min(pts_ps[i].Y, pts_pv[i].Y) - 5, max(pts_ps[i].Y, pts_pv[i].Y) + 5
                h_line = y_bot - y_top
                
                pen_ln = Pen(Color.FromArgb(180, 60, 140, 220), 2)
                pen_ln.DashStyle = DashStyle.Dot
                g.DrawLine(pen_ln, zone_x, int(y_top), zone_x, int(y_bot))
                
                n_drops = min(6, max(3, int(h_line / 20)))
                spacing = h_line / (n_drops + 1)
                brush_w = SolidBrush(Color.FromArgb(200, 80, 160, 240))
                pen_w = Pen(Color.FromArgb(255, 40, 120, 200), 1)
                
                for d in range(n_drops):
                    ox = -6 if d % 2 == 0 else 6
                    dy = y_top + spacing * (d + 1)
                    sz = int(5 + (1 - abs(d - n_drops/2.0) / (n_drops/2.0)) * 4)
                    dx, dyi = zone_x + ox, int(dy)
                    g.FillEllipse(brush_w, dx - sz/2, dyi - sz, sz, sz)
                    g.DrawEllipse(pen_w, dx - sz/2, dyi - sz, sz, sz)
                    ps = int(sz * 0.4)
                    g.DrawLine(pen_w, dx, dyi + ps, dx - ps, dyi)
                    g.DrawLine(pen_w, dx, dyi + ps, dx + ps, dyi)
                break  # Alleen eerste zone
        
        # Legenda
        lx, ly = int(x_start + w_draw - 180), m_top + 5
        font_leg = Font("Segoe UI", 7)
        g.DrawLine(pen_t, lx, ly + 6, lx + 25, ly + 6)
        g.DrawString(u"Temperatuur (\u00b0C)", font_leg, brush_txt, lx + 30, ly)
        g.DrawLine(pen_ps, lx, ly + 22, lx + 25, ly + 22)
        g.DrawString("Verzadigingsdruk", font_leg, brush_txt, lx + 30, ly + 16)
        g.DrawLine(pen_pv, lx, ly + 38, lx + 25, ly + 38)
        g.DrawString("Dampdruk", font_leg, brush_txt, lx + 30, ly + 32)


# ==============================================================================
# VOCHTBALANS PANEL (Staafdiagram met pos/neg en klimaatdata)
# ==============================================================================
class VochtBalansPanel(Panel):
    def __init__(self):
        self.monthly_change = [0] * 12
        self.cumulative = [0] * 12
        self.BackColor = Color.White
        self.BorderStyle = BorderStyle.None
    
    def set_data(self, monthly_change, cumulative):
        self.monthly_change = monthly_change
        self.cumulative = cumulative
        self.Invalidate()
    
    def OnPaint(self, e):
        g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        
        m_left, m_right, m_top, m_bot = 50, 20, 20, 55  # Meer ruimte onderaan voor klimaatdata
        w = self.Width - m_left - m_right
        h = self.Height - m_top - m_bot
        
        # Titel
        font_title = Font("Segoe UI", 9, FontStyle.Bold)
        brush_txt = SolidBrush(Huisstijl.VIOLET)
        g.DrawString(u"Vochtbalans (g/m\u00b2) - cumulatief", font_title, brush_txt, m_left, 3)
        
        # Bepaal schaal (moet zowel pos als neg kunnen tonen)
        max_val = max(self.cumulative) if self.cumulative else 0
        min_val = min(self.monthly_change) if self.monthly_change else 0
        
        # Zorg voor minimale range
        if max_val < 10:
            max_val = 10
        if min_val > -10:
            min_val = -10
        
        # Y bereik
        y_range = max_val - min_val
        if y_range == 0:
            y_range = 20
        
        # Positie van y=0 lijn
        zero_y = m_top + int((max_val / y_range) * h)
        
        # Assen
        pen_axis = Pen(Huisstijl.DARK_GRAY, 1)
        g.DrawLine(pen_axis, m_left, m_top, m_left, m_top + h)
        g.DrawLine(pen_axis, m_left, zero_y, m_left + w, zero_y)  # X-as op y=0
        
        # Y-as labels
        font_small = Font("Segoe UI", 7)
        
        # Positieve labels
        if max_val > 0:
            for i in range(1, 5):
                y_val = float(max_val) * i / 4.0
                y_pos = m_top + int(((max_val - y_val) / y_range) * h)
                if y_pos >= m_top:
                    g.DrawString("{0:.0f}".format(y_val), font_small, brush_txt, 5, y_pos - 5)
        
        # Negatieve labels
        if min_val < 0:
            for i in range(1, 3):
                y_val = float(min_val) * i / 2.0
                y_pos = m_top + int(((max_val - y_val) / y_range) * h)
                if y_pos <= m_top + h:
                    g.DrawString("{0:.0f}".format(y_val), font_small, brush_txt, 5, y_pos - 5)
        
        # 0-label
        g.DrawString("0", font_small, brush_txt, 5, zero_y - 5)
        
        # Gridlijnen
        pen_grid = Pen(Color.FromArgb(50, 150, 150, 150), 1)
        pen_grid.DashStyle = DashStyle.Dot
        
        # Staven en lijn tekenen
        bar_w = int(w / 14)
        bar_gap = int((w - 12 * bar_w) / 13)
        
        font_tiny = Font("Segoe UI", 6)
        brush_gray = SolidBrush(Huisstijl.TEXT_SECONDARY)
        sf_center = StringFormat()
        sf_center.Alignment = sf_center.Alignment.Center
        
        # Eerst de maandelijkse verandering staven (achtergrond)
        for i, change in enumerate(self.monthly_change):
            x = m_left + bar_gap + i * (bar_w + bar_gap)
            
            if change > 0:
                # Condensatie: blauwe staaf omhoog
                bar_h = int((change / y_range) * h)
                y = zero_y - bar_h
                color = Color.FromArgb(100, 80, 140, 220)
                g.FillRectangle(SolidBrush(color), x, y, bar_w, bar_h)
            elif change < 0:
                # Droging: groene staaf omlaag
                bar_h = int((abs(change) / y_range) * h)
                y = zero_y
                color = Color.FromArgb(100, 80, 180, 120)
                g.FillRectangle(SolidBrush(color), x, y, bar_w, bar_h)
        
        # Dan de cumulatieve lijn
        pen_cum = Pen(Huisstijl.VIOLET, 2)
        points = []
        for i, cum in enumerate(self.cumulative):
            x = m_left + bar_gap + i * (bar_w + bar_gap) + bar_w / 2
            y = m_top + int(((max_val - cum) / y_range) * h)
            points.append(Point(int(x), int(y)))
        
        for i in range(len(points) - 1):
            g.DrawLine(pen_cum, points[i], points[i + 1])
        
        # Punten op de lijn
        for pt in points:
            g.FillEllipse(SolidBrush(Huisstijl.VIOLET), pt.X - 3, pt.Y - 3, 6, 6)
        
        # Maandlabels + klimaatdata
        for i in range(12):
            x = m_left + bar_gap + i * (bar_w + bar_gap)
            
            # Maandnaam
            g.DrawString(NL_KLIMAAT['maanden'][i], font_small, brush_txt, 
                        RectangleF(float(x - 5), float(m_top + h + 3), float(bar_w + 10), float(12)), sf_center)
            
            # Temperatuur buiten
            t_out = float(NL_KLIMAAT['T_buiten'][i])
            g.DrawString(u"{0:.0f}\u00b0C".format(t_out), font_tiny, brush_gray,
                        RectangleF(float(x - 5), float(m_top + h + 15), float(bar_w + 10), float(10)), sf_center)
            
            # RV buiten
            rh_out = int(NL_KLIMAAT['RH_buiten'][i])
            g.DrawString("{0}%".format(rh_out), font_tiny, brush_gray,
                        RectangleF(float(x - 5), float(m_top + h + 25), float(bar_w + 10), float(10)), sf_center)
        
        # Legenda
        leg_x = m_left + w - 200
        leg_y = 3
        
        # Cumulatief lijn
        g.DrawLine(pen_cum, leg_x, leg_y + 6, leg_x + 20, leg_y + 6)
        g.FillEllipse(SolidBrush(Huisstijl.VIOLET), leg_x + 7, leg_y + 3, 6, 6)
        g.DrawString("Cumulatief", font_small, brush_txt, leg_x + 25, leg_y)
        
        # Condensatie
        g.FillRectangle(SolidBrush(Color.FromArgb(150, 80, 140, 220)), leg_x + 90, leg_y + 2, 12, 10)
        g.DrawString("Cond.", font_small, brush_txt, leg_x + 105, leg_y)
        
        # Droging  
        g.FillRectangle(SolidBrush(Color.FromArgb(150, 80, 180, 120)), leg_x + 145, leg_y + 2, 12, 10)
        g.DrawString("Droog", font_small, brush_txt, leg_x + 160, leg_y)
        
        # Status tekst
        end_val = self.cumulative[-1] if self.cumulative else 0
        if end_val > 500:
            font_warn = Font("Segoe UI", 8, FontStyle.Bold)
            g.DrawString(u"\u26a0 Risico!", font_warn, SolidBrush(Huisstijl.PEACH), m_left, 3)


# ==============================================================================
# MATERIAAL KIEZER DIALOG
# ==============================================================================
class MateriaalKiezerForm(BaseForm):
    def __init__(self, huidige_naam=""):
        super(MateriaalKiezerForm, self).__init__("Materiaal kiezen", 620, 280)
        self.result = None
        self.db_tuples = get_db_tuples()  # Cache de database tuples
        self._setup_ui(huidige_naam)
    
    def _setup_ui(self, huidige_naam):
        y = 10
        kort = huidige_naam[:45] + "..." if len(huidige_naam) > 45 else huidige_naam
        lbl = UIFactory.create_label("Huidig: {}".format(kort or "-"), color=Huisstijl.TEXT_SECONDARY)
        lbl.Location = DPIScaler.scale_point(10, y)
        self.pnl_content.Controls.Add(lbl)
        y += 35
        
        items = []
        for cat, naam, lam, rd, mu in self.db_tuples:
            if rd is not None and rd > 0:
                items.append(u"{} - {} [Rd={:.2f}, \u03bc={}]".format(cat, naam, rd, mu))
            elif lam and lam > 0:
                items.append(u"{} - {} [\u03bb={:.3f}, \u03bc={}]".format(cat, naam, lam, mu))
            else:
                items.append(u"{} - {} [\u03bc={}]".format(cat, naam, mu))
        
        self.cmb = UIFactory.create_combobox(560, items)
        self.cmb.Location = DPIScaler.scale_point(10, y)
        self.cmb.SelectedIndex = 0  # Selecteer eerste item standaard
        self.pnl_content.Controls.Add(self.cmb)
        
        # Info label
        y += 40
        self.lbl_info = UIFactory.create_label("", color=Huisstijl.TEXT_SECONDARY, italic=True)
        self.lbl_info.Location = DPIScaler.scale_point(10, y)
        self.lbl_info.AutoSize = True
        self.pnl_content.Controls.Add(self.lbl_info)
        
        # Update info bij selectie wijziging
        self.cmb.SelectedIndexChanged += self._on_selection_changed
        self._on_selection_changed(None, None)  # Initial update
        
        self.add_footer_button("OK", 'primary', self._on_ok)
        self.add_footer_button("Annuleren", 'secondary', self._on_cancel)
    
    def _on_selection_changed(self, sender, args):
        idx = self.cmb.SelectedIndex
        if 0 <= idx < len(self.db_tuples):
            cat, naam, lam, rd, mu = self.db_tuples[idx]
            if rd is not None and rd > 0:
                self.lbl_info.Text = u"Spouw/luchtlaag met vaste Rd={:.2f} m\u00b2K/W".format(rd)
            elif lam and lam > 0:
                self.lbl_info.Text = u"\u03bb={:.4f} W/(m\u00b7K), \u03bc={} (dampdiffusieweerstand)".format(lam, mu)
            else:
                self.lbl_info.Text = u"Folie/membraan: \u03bc={} (alleen dampdiffusie, geen R-waarde)".format(mu)
    
    def _on_ok(self, sender, args):
        idx = self.cmb.SelectedIndex
        if 0 <= idx < len(self.db_tuples):
            self.result = self.db_tuples[idx]  # (cat, naam, lambda, rd, mu)
        self.Close()
    
    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()


# ==============================================================================
# HOOFD FORMULIER
# ==============================================================================
class RcBerekeningForm(BaseForm):
    def __init__(self, walls, n_normal, n_curtain):
        super(RcBerekeningForm, self).__init__("Rc-waarde", 1000, 1000)  # Compacter formaat
        self.walls = walls
        self.n_normal = n_normal
        self.n_curtain = n_curtain
        self.lagen = []
        self._load_layers()
        self._setup_ui()
        self._refresh()
    
    def _load_layers(self):
        """Laad lagen van alle wanden - GEEN automatische spouw detectie"""
        for wall in self.walls:
            # Check of het een curtain wall is
            if isinstance(wall, DB.Wall):
                try:
                    if wall.WallType.Kind == DB.WallKind.Curtain:
                        panels = get_curtain_wall_panels(wall)
                        for panel in panels:
                            layers = get_curtain_panel_layers(panel)
                            for l in layers:
                                self.lagen.append(LaagData(
                                    l['wall_type'], l['material'], l['width_mm'], l['lambda'],
                                    l['r_value'], l['mu'], l['is_air_gap'], l['is_membrane']
                                ))
                    else:
                        # Normale wand
                        layers = get_wall_layers(wall)
                        for l in layers:
                            self.lagen.append(LaagData(
                                l['wall_type'], l['material'], l['width_mm'], l['lambda'],
                                l['r_value'], l['mu'], l['is_air_gap'], l['is_membrane']
                            ))
                except:
                    layers = get_wall_layers(wall)
                    for l in layers:
                        self.lagen.append(LaagData(
                            l['wall_type'], l['material'], l['width_mm'], l['lambda'],
                            l['r_value'], l['mu'], l['is_air_gap'], l['is_membrane']
                        ))
    
    def _setup_ui(self):
        info = ""
        if self.n_normal > 0:
            info += "{} wand".format(self.n_normal)
        if self.n_curtain > 0:
            info += " + {} curtain wall".format(self.n_curtain) if info else "{} curtain wall".format(self.n_curtain)
        info += " | {} materialen".format(len(get_materiaal_db()))
        self.set_subtitle(info)
        
        y = 5
        
        # Preview panel (Glaser diagram)
        self.pnl_preview = Panel()
        self.pnl_preview.Location = DPIScaler.scale_point(10, y)
        self.pnl_preview.Size = DPIScaler.scale_size(PANEL_WIDTH, 180)
        self.pnl_preview.BorderStyle = BorderStyle.FixedSingle
        self.pnl_content.Controls.Add(self.pnl_preview)
        
        self.pnl_opbouw = WandOpbouwPanel()
        self.pnl_opbouw.Location = Point(0, 0)
        self.pnl_opbouw.Size = Size(self.pnl_preview.Width - 2, self.pnl_preview.Height - 2)
        self.pnl_preview.Controls.Add(self.pnl_opbouw)
        y += 190
        
        # Vochtbalans panel (groter voor klimaatdata)
        self.pnl_moisture_border = Panel()
        self.pnl_moisture_border.Location = DPIScaler.scale_point(10, y)
        self.pnl_moisture_border.Size = DPIScaler.scale_size(PANEL_WIDTH, 130)  # Hoger voor klimaatdata
        self.pnl_moisture_border.BorderStyle = BorderStyle.FixedSingle
        self.pnl_content.Controls.Add(self.pnl_moisture_border)
        
        self.pnl_moisture = VochtBalansPanel()
        self.pnl_moisture.Location = Point(0, 0)
        self.pnl_moisture.Size = Size(self.pnl_moisture_border.Width - 2, self.pnl_moisture_border.Height - 2)
        self.pnl_moisture_border.Controls.Add(self.pnl_moisture)
        y += 140
        
        # DataGridView - zelfde breedte
        cols = [("#", "#", 35), ("wandtype", "Wandtype", 160), ("materiaal", "Materiaal", 220),
                ("dikte", "Dikte", 70), ("lambda", u"\u03bb", 85), ("r", "R", 85),
                ("mu", u"\u03bc", 75), ("sd", "Sd", 75)]
        self.grid = UIFactory.create_datagridview(cols, PANEL_WIDTH, 200, allow_edit=True)
        self.grid.Location = DPIScaler.scale_point(10, y)
        self.grid.Columns["#"].ReadOnly = True
        self.grid.CellDoubleClick += self._on_cell_dblclick
        self.grid.CellEndEdit += self._on_cell_edit
        self.pnl_content.Controls.Add(self.grid)
        y += 210
        
        # Knoppen rij
        btns = [
            ("Laag +", 'primary', self._add_layer),
            ("Laag -", 'secondary', self._remove_layer),
            (u"\u2191", 'icon', self._move_up),
            (u"\u2193", 'icon', self._move_down),
            (u"\u21c4", 'warning', self._flip),
        ]
        x = 10
        for text, style, handler in btns:
            w = 45 if style == 'icon' else 90
            btn = UIFactory.create_button(text, w, 38, style)
            btn.Location = DPIScaler.scale_point(x, y)
            btn.Click += handler
            self.pnl_content.Controls.Add(btn)
            x += w + 10
        
        lbl_hint = UIFactory.create_label(u"Dubbelklik \u03bb of \u03bc voor materiaalkeuze", italic=True, color=Huisstijl.TEXT_SECONDARY)
        lbl_hint.Location = DPIScaler.scale_point(x + 20, y + 10)
        self.pnl_content.Controls.Add(lbl_hint)
        y += 50
        
        # Climate inputs
        climate_data = [
            ("Rsi:", RSI_DEFAULT, 50), ("Rse:", RSE_DEFAULT, 50),
            ("Ti:", T_BINNEN, 40), (u"\u00b0C", None, 25),
            ("Te:", T_BUITEN, 40), (u"\u00b0C", None, 25),
            ("RVi:", RH_BINNEN, 40), ("%", None, 20),
            ("RVe:", RH_BUITEN, 40), ("%", None, 20),
        ]
        self.txt_inputs = {}
        x = 10
        for label, default, w in climate_data:
            if default is not None:
                lbl = UIFactory.create_label(label)
                lbl.Location = DPIScaler.scale_point(x, y + 5)
                self.pnl_content.Controls.Add(lbl)
                x += lbl.Width + 5
                
                txt = UIFactory.create_textbox(w)
                txt.Text = str(default)
                txt.Location = DPIScaler.scale_point(x, y)
                txt.TextChanged += self._on_values_changed
                self.pnl_content.Controls.Add(txt)
                self.txt_inputs[label.replace(":", "")] = txt
                x += w + 15
            else:
                lbl = UIFactory.create_label(label)
                lbl.Location = DPIScaler.scale_point(x - 10, y + 5)
                self.pnl_content.Controls.Add(lbl)
                x += 15
        y += 45
        
        # Resultaat box - zelfde breedte
        gb = UIFactory.create_groupbox("Resultaat", PANEL_WIDTH, 120)
        gb.Location = DPIScaler.scale_point(10, y)
        self.pnl_content.Controls.Add(gb)
        
        self.lbl_dikte = UIFactory.create_label("Dikte: - mm")
        self.lbl_dikte.Location = DPIScaler.scale_point(20, 30)
        gb.Controls.Add(self.lbl_dikte)
        
        self.lbl_sd = UIFactory.create_label("Sd: - m")
        self.lbl_sd.Location = DPIScaler.scale_point(200, 30)
        gb.Controls.Add(self.lbl_sd)
        
        self.lbl_status = UIFactory.create_label("", color=Huisstijl.PEACH, italic=True)
        self.lbl_status.Location = DPIScaler.scale_point(400, 30)
        gb.Controls.Add(self.lbl_status)
        
        self.lbl_rc = UIFactory.create_label("Rc: - m²K/W", font_size=20, bold=True, color=Huisstijl.TEAL)
        self.lbl_rc.Location = DPIScaler.scale_point(20, 65)
        gb.Controls.Add(self.lbl_rc)
        
        self.lbl_u = UIFactory.create_label("U: - W/m²K", font_size=20, bold=True, color=Huisstijl.TEAL)
        self.lbl_u.Location = DPIScaler.scale_point(400, 65)
        gb.Controls.Add(self.lbl_u)
        
        # Footer buttons
        self.add_footer_button("CSV", 'secondary', self._export_csv, 90)
        self.add_footer_button("PDF", 'primary', self._export_pdf, 90)
    
    def _refresh(self):
        self._update_grid()
        self._update_preview()
        self._update_moisture()
        self._calc_totals()
    
    def _update_grid(self):
        self.grid.Rows.Clear()
        for i, l in enumerate(self.lagen):
            nr = str(i + 1)
            lam = "(spouw)" if l.is_air_gap else ("{:.4f}".format(l.lambda_val) if l.lambda_val else "")
            r = "{:.3f}".format(l.r_value) if l.r_value else "-"
            mu = str(l.mu) if l.mu else "-"
            sd = "{:.2f}".format(l.sd) if l.sd else "-"
            self.grid.Rows.Add(nr, l.wall_type or "(hand)", l.material, "{:.1f}".format(l.width_mm), lam, r, mu, sd)
    
    def _update_preview(self):
        try:
            rsi = float(self.txt_inputs.get("Rsi", type('', (), {'Text': str(RSI_DEFAULT)})()).Text.replace(',', '.'))
            rse = float(self.txt_inputs.get("Rse", type('', (), {'Text': str(RSE_DEFAULT)})()).Text.replace(',', '.'))
            ti = float(self.txt_inputs.get("Ti", type('', (), {'Text': str(T_BINNEN)})()).Text.replace(',', '.'))
            te = float(self.txt_inputs.get("Te", type('', (), {'Text': str(T_BUITEN)})()).Text.replace(',', '.'))
            rhi = float(self.txt_inputs.get("RVi", type('', (), {'Text': str(RH_BINNEN)})()).Text.replace(',', '.'))
            rhe = float(self.txt_inputs.get("RVe", type('', (), {'Text': str(RH_BUITEN)})()).Text.replace(',', '.'))
        except:
            rsi, rse, ti, te, rhi, rhe = RSI_DEFAULT, RSE_DEFAULT, T_BINNEN, T_BUITEN, RH_BINNEN, RH_BUITEN
        self.pnl_opbouw.set_data(self.lagen, rsi, rse, ti, te, rhi, rhe)
    
    def _update_moisture(self):
        try:
            rsi = float(self.txt_inputs["Rsi"].Text.replace(',', '.'))
            rse = float(self.txt_inputs["Rse"].Text.replace(',', '.'))
        except:
            rsi, rse = RSI_DEFAULT, RSE_DEFAULT
        
        monthly_change, cumulative = calculate_monthly_moisture(self.lagen, rsi, rse)
        self.pnl_moisture.set_data(monthly_change, cumulative)
    
    def _calc_totals(self):
        try:
            rsi = float(self.txt_inputs["Rsi"].Text.replace(',', '.'))
            rse = float(self.txt_inputs["Rse"].Text.replace(',', '.'))
        except:
            rsi, rse = RSI_DEFAULT, RSE_DEFAULT
        
        dikte = sum(l.width_mm for l in self.lagen)
        sd_tot = sum(l.sd for l in self.lagen if l.sd)
        r_tot = rsi + sum(l.r_value for l in self.lagen if l.r_value) + rse
        
        self.lbl_dikte.Text = "Dikte: {:.0f} mm".format(dikte)
        self.lbl_sd.Text = "Sd: {:.2f} m".format(sd_tot)
        
        missing = [l for l in self.lagen if not l.lambda_val and not l.r_value and not l.is_membrane]
        if missing:
            self.lbl_status.Text = "{} laag/lagen zonder lambda".format(len(missing))
            self.lbl_rc.Text = u"Rc: ~{:.2f} m\u00b2K/W".format(r_tot)
            self.lbl_rc.ForeColor = Huisstijl.YELLOW
        else:
            self.lbl_status.Text = ""
            self.lbl_rc.Text = u"Rc: {:.2f} m\u00b2K/W".format(r_tot)
            self.lbl_rc.ForeColor = Huisstijl.TEAL
        
        self.lbl_u.Text = u"U: {:.3f} W/m\u00b2K".format(1/r_tot) if r_tot > 0 else "U: - W/m²K"
    
    def _on_cell_dblclick(self, sender, args):
        if args.ColumnIndex in [4, 6]:  # lambda of mu
            idx = args.RowIndex
            if 0 <= idx < len(self.lagen):
                dlg = MateriaalKiezerForm(self.lagen[idx].material)
                dlg.ShowDialog()
                if dlg.result:
                    cat, naam, lam, rd, mu = dlg.result
                    
                    # Debug: toon wat er geselecteerd is
                    # print("Geselecteerd: {} | lam={} | rd={} | mu={}".format(naam, lam, rd, mu))
                    
                    # Update alle waarden
                    self.lagen[idx].material = naam
                    self.lagen[idx].mu = mu
                    w_m = self.lagen[idx].width_mm / 1000.0
                    
                    # Bepaal r_value en is_air_gap
                    if rd is not None and rd > 0:
                        # Spouw met vaste Rd-waarde (> 0)
                        self.lagen[idx].lambda_val = None
                        self.lagen[idx].r_value = rd
                        self.lagen[idx].is_air_gap = True
                    elif lam is not None and lam > 0:
                        # Normaal materiaal met lambda
                        self.lagen[idx].lambda_val = lam
                        self.lagen[idx].r_value = w_m / lam if w_m > 0 else 0
                        self.lagen[idx].is_air_gap = False
                    else:
                        # Folie of materiaal zonder lambda (rd=0 of rd=None, lam=None)
                        self.lagen[idx].lambda_val = None
                        self.lagen[idx].r_value = 0
                        self.lagen[idx].is_air_gap = False
                    
                    self._refresh()
                    
                    # Toon bevestiging
                    # self.show_info("Materiaal gewijzigd naar:\n{}\nλ={}, Rd={}, μ={}".format(naam, lam, rd, mu))
    
    def _on_cell_edit(self, sender, args):
        idx, col = args.RowIndex, args.ColumnIndex
        if idx < 0 or idx >= len(self.lagen):
            return
        try:
            val = str(self.grid.Rows[idx].Cells[col].Value).strip().replace(',', '.')
            if col == 2:  # materiaal
                self.lagen[idx].material = val
            elif col == 3:  # dikte
                self.lagen[idx].width_mm = float(val)
                if self.lagen[idx].lambda_val and self.lagen[idx].lambda_val > 0:
                    self.lagen[idx].r_value = (float(val) / 1000.0) / self.lagen[idx].lambda_val
            elif col == 4:  # lambda
                if val.lower() in ['spouw', '(spouw)', '']:
                    self.lagen[idx].is_air_gap = True
                    self.lagen[idx].lambda_val = None
                    self.lagen[idx].r_value = SPOUW_RD
                else:
                    self.lagen[idx].lambda_val = float(val)
                    self.lagen[idx].is_air_gap = False
                    if self.lagen[idx].lambda_val > 0:
                        self.lagen[idx].r_value = (self.lagen[idx].width_mm / 1000.0) / self.lagen[idx].lambda_val
            elif col == 6:  # mu
                self.lagen[idx].mu = int(float(val)) if val and val != "-" else None
            self._refresh()
        except:
            pass
    
    def _on_values_changed(self, sender, args):
        self._update_preview()
        self._update_moisture()
        self._calc_totals()
    
    def _add_layer(self, sender, args):
        idx = self.grid.SelectedRows[0].Index + 1 if self.grid.SelectedRows.Count > 0 else len(self.lagen)
        self.lagen.insert(idx, LaagData("(hand)", "PE-folie", 0.2, None, None, 50000, False, False))
        self._refresh()
        if idx < self.grid.Rows.Count:
            self.grid.ClearSelection()
            self.grid.Rows[idx].Selected = True
    
    def _remove_layer(self, sender, args):
        if self.grid.SelectedRows.Count > 0:
            idx = self.grid.SelectedRows[0].Index
            if idx < len(self.lagen):
                del self.lagen[idx]
                self._refresh()
    
    def _move_up(self, sender, args):
        if self.grid.SelectedRows.Count > 0:
            idx = self.grid.SelectedRows[0].Index
            if idx > 0:
                self.lagen[idx], self.lagen[idx-1] = self.lagen[idx-1], self.lagen[idx]
                self._refresh()
                self.grid.ClearSelection()
                self.grid.Rows[idx-1].Selected = True
    
    def _move_down(self, sender, args):
        if self.grid.SelectedRows.Count > 0:
            idx = self.grid.SelectedRows[0].Index
            if idx < len(self.lagen) - 1:
                self.lagen[idx], self.lagen[idx+1] = self.lagen[idx+1], self.lagen[idx]
                self._refresh()
                self.grid.ClearSelection()
                self.grid.Rows[idx+1].Selected = True
    
    def _flip(self, sender, args):
        self.lagen.reverse()
        self._refresh()
    
    def _export_csv(self, sender, args):
        path = self.save_file_dialog("CSV (*.csv)|*.csv", "Rc_berekening.csv")
        if path:
            try:
                with open(path, 'w') as f:
                    f.write("#;Wandtype;Materiaal;Dikte mm;Lambda;R-waarde;Mu;Sd (m)\n")
                    for i, l in enumerate(self.lagen):
                        lam = "(spouw)" if l.is_air_gap else ("{:.4f}".format(l.lambda_val) if l.lambda_val else "")
                        f.write("{};{};{};{:.1f};{};{};{};{}\n".format(
                            i+1, l.wall_type or "(hand)", l.material, l.width_mm, lam,
                            "{:.3f}".format(l.r_value) if l.r_value else "",
                            l.mu or "", "{:.2f}".format(l.sd) if l.sd else ""
                        ))
                    f.write("\nRc;{}\nU;{}\n".format(self.lbl_rc.Text, self.lbl_u.Text))
                    
                    # Vochtbalans toevoegen
                    try:
                        rsi = float(self.txt_inputs["Rsi"].Text.replace(',', '.'))
                        rse = float(self.txt_inputs["Rse"].Text.replace(',', '.'))
                    except:
                        rsi, rse = RSI_DEFAULT, RSE_DEFAULT
                    monthly_change, cumulative = calculate_monthly_moisture(self.lagen, rsi, rse)
                    f.write("\nVochtbalans (g/m2)\n")
                    f.write(";".join(NL_KLIMAAT['maanden']) + "\n")
                    f.write("Maandelijks;" + ";".join("{:.1f}".format(m) for m in monthly_change) + "\n")
                    f.write("Cumulatief;" + ";".join("{:.1f}".format(m) for m in cumulative) + "\n")
                    
                self.show_info("Export klaar!\n\n" + path)
            except Exception as e:
                self.show_error("Fout: " + str(e))

    def _export_pdf(self, sender, args):
        path = self.save_file_dialog("PDF (*.pdf)|*.pdf", "Rc_berekening.pdf")
        if not path:
            return
        try:
            # Lees huidige inputwaarden
            try:
                rsi = float(self.txt_inputs["Rsi"].Text.replace(',', '.'))
                rse = float(self.txt_inputs["Rse"].Text.replace(',', '.'))
                ti = float(self.txt_inputs["Ti"].Text.replace(',', '.'))
                te = float(self.txt_inputs["Te"].Text.replace(',', '.'))
                rhi = float(self.txt_inputs["RVi"].Text.replace(',', '.'))
                rhe = float(self.txt_inputs["RVe"].Text.replace(',', '.'))
            except:
                rsi, rse = RSI_DEFAULT, RSE_DEFAULT
                ti, te = T_BINNEN, T_BUITEN
                rhi, rhe = RH_BINNEN, RH_BUITEN

            # Bereken totalen
            rc_total = rsi + sum(l.r_value for l in self.lagen if l.r_value) + rse
            u_value = (1.0 / rc_total) if rc_total > 0 else 0

            # Vochtbalans
            monthly_change, cumulative = calculate_monthly_moisture(self.lagen, rsi, rse)

            # Capture afbeeldingen
            cover_image = capture_3d_view()
            diagram_image = capture_panel_image(self.pnl_opbouw)

            # Genereer PDF
            generate_pdf_report(
                self.lagen, rsi, rse, ti, te, rhi, rhe,
                rc_total, u_value,
                monthly_change, cumulative,
                path,
                cover_image=cover_image,
                diagram_image=diagram_image
            )
            self.show_info("PDF rapport opgeslagen!\n\n" + path)
        except Exception as e:
            self.show_error("PDF export mislukt:\n" + str(e))


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    selection = revit.get_selection()
    walls = []
    n_normal = 0
    n_curtain = 0
    
    for el in selection:
        if isinstance(el, DB.Wall):
            walls.append(el)
            try:
                if el.WallType.Kind == DB.WallKind.Curtain:
                    n_curtain += 1
                else:
                    n_normal += 1
            except:
                n_normal += 1
        elif isinstance(el, DB.Panel):
            # Direct geselecteerd curtain wall panel - niet ondersteund
            pass
    
    if not walls:
        forms.alert("Selecteer eerst wanden (normale of curtain walls).", title="Rc-waarde")
        return
    
    form = RcBerekeningForm(walls, n_normal, n_curtain)
    form.ShowDialog()


if __name__ == '__main__':
    main()
