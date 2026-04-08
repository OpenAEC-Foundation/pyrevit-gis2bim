# -*- coding: utf-8 -*-
"""
Hellingbaan Generator - NEN 2443 conforme hellingbanen
Genereert concept vloeren voor parkeergarage hellingbanen
"""
__title__ = "Helling"
__author__ = "3BM Bouwkunde"
__doc__ = "Genereer NEN 2443 conforme hellingbanen met concept vloeren"

# Imports
import math
import clr

from System.Collections.Generic import List
from System.Windows import Visibility, FontWeights
from System.Windows.Shapes import Line as WpfLine, Rectangle as WpfRect
from System.Windows.Controls import Canvas as WpfCanvas, TextBlock as WpfTextBlock
from System.Windows.Media import SolidColorBrush, DoubleCollection

from Autodesk.Revit.DB import (
    FilteredElementCollector, Floor, FloorType, Level,
    Transaction, XYZ, CurveLoop, Line, ElementId,
    BuiltInCategory, BuiltInParameter,
    SketchPlane, Plane, ViewPlan,
    DirectShape, GeometryCreationUtilities, SolidOptions,
    SolidUtils, Transform
)
from Autodesk.Revit.UI.Selection import ObjectSnapTypes
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms, script

# UI imports
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from wpf_template import WPFWindow, Huisstijl
from bm_logger import get_logger

# Setup - wordt later geinitialiseerd
doc = None
output = None
log = get_logger("HellingbaanGenerator")


# ==============================================================================
# NEN 2443 CONSTANTEN
# ==============================================================================
WIELBASIS = 2770  # mm
OVERGANG_VOET = WIELBASIS  # 2770 mm
OVERGANG_TOP = WIELBASIS / 2  # 1385 mm
OVERGANG_TOTAAL = OVERGANG_VOET + OVERGANG_TOP  # 4155 mm

# Garage types met grenzen volgens NEN 2443
GARAGE_TYPES = [
    {'id': 'openbaar', 'label': 'Openbaar', 'max': 14, 'min': 14, 'breedte_min': 3000},
    {'id': 'openbaar_dhumy', 'label': "Openbaar (d'Humy)", 'max': 15, 'min': 14, 'breedte_min': 3000},
    {'id': 'niet_openbaar', 'label': 'Niet-openbaar', 'max': 20, 'min': 14, 'breedte_min': 2750},
    {'id': 'stalling', 'label': 'Stalling', 'max': 24, 'min': 14, 'breedte_min': 2750},
]

LEN_MIN = 10000  # mm
LEN_MAX = 40000  # mm


# ==============================================================================
# BEREKENING
# ==============================================================================
def bereken_hellingbaan(hoogte, garage_type, met_overgang=True, helling_override=None):
    """
    Bereken hellingbaan volgens NEN 2443.
    """
    max_helling = garage_type['max']
    min_helling = garage_type['min']

    # Hoogtegrenzen voor de drie zones
    hoogte_min = LEN_MIN * max_helling / 100.0
    hoogte_max = LEN_MAX * min_helling / 100.0

    # Zonder overgangshelling: simpele berekening
    if not met_overgang:
        helling = float(helling_override) if helling_override else float(max_helling)
        lengte = hoogte / helling * 100.0

        return {
            'helling': float(helling),
            'helling_berekend': float(max_helling),
            'is_override': helling_override is not None,
            'lengte': float(lengte),
            'zone': 'simpel',
            'overgang_onder_lengte': 0.0,
            'overgang_boven_lengte': 0.0,
            'hoofd_lengte': float(lengte),
            'overgang_helling': 0.0,
            'overgang_onder_hoogte': 0.0,
            'overgang_boven_hoogte': 0.0,
            'hoofd_hoogte': float(hoogte),
        }

    # Bepaal zone en bereken helling
    if max_helling == min_helling:
        zone = 'vast'
        helling_berekend = float(max_helling)
    elif hoogte <= hoogte_min:
        zone = 'kort'
        helling_berekend = float(max_helling)
    elif hoogte >= hoogte_max:
        zone = 'lang'
        helling_berekend = float(min_helling)
    else:
        zone = 'midden'
        factor = (LEN_MAX - LEN_MIN) / float(max_helling - min_helling)
        a = -factor
        b = LEN_MIN + factor * max_helling - OVERGANG_TOTAAL / 2.0
        c = -hoogte * 100.0
        discriminant = b * b - 4 * a * c

        if discriminant >= 0:
            helling_berekend = (-b - math.sqrt(discriminant)) / (2 * a)
            helling_berekend = max(min_helling, min(max_helling, helling_berekend))
        else:
            helling_berekend = float(min_helling)

    # Gebruik override of berekende waarde
    helling = helling_override if helling_override else helling_berekend

    # Bereken segmenten
    overgang_helling = helling / 2.0
    overgang_onder_lengte = OVERGANG_VOET
    overgang_boven_lengte = OVERGANG_TOP

    overgang_onder_hoogte = overgang_onder_lengte * overgang_helling / 100.0
    overgang_boven_hoogte = overgang_boven_lengte * overgang_helling / 100.0
    hoofd_hoogte = hoogte - overgang_onder_hoogte - overgang_boven_hoogte

    hoofd_lengte = hoofd_hoogte / helling * 100.0
    lengte = overgang_onder_lengte + hoofd_lengte + overgang_boven_lengte

    return {
        'helling': float(helling),
        'helling_berekend': float(helling_berekend),
        'is_override': helling_override is not None,
        'lengte': float(lengte),
        'zone': zone,
        'overgang_onder_lengte': float(overgang_onder_lengte),
        'overgang_boven_lengte': float(overgang_boven_lengte),
        'hoofd_lengte': float(hoofd_lengte),
        'overgang_helling': float(overgang_helling),
        'overgang_onder_hoogte': float(overgang_onder_hoogte),
        'overgang_boven_hoogte': float(overgang_boven_hoogte),
        'hoofd_hoogte': float(hoofd_hoogte),
    }


# ==============================================================================
# MAIN FORM
# ==============================================================================
class HellingbaanWindow(WPFWindow):
    """Hoofdvenster voor Hellingbaan Generator - WPF versie"""

    def __init__(self, floor_types, levels):
        xaml_file = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        super(HellingbaanWindow, self).__init__(
            xaml_file, "Hellingbaan Generator", width=500, height=None
        )

        self.floor_types = floor_types
        self.levels = levels
        self.helling_override = None
        self.breedte_override = None
        self.berekening = None
        self.hoogte = 3600
        self.met_overgang = True
        self.plaats_gevraagd = False

        self._populate_combos()
        self._bind_events()
        self._update_berekening()

    def _populate_combos(self):
        """Vul comboboxen"""
        for gt in GARAGE_TYPES:
            self.cmb_garage.Items.Add("{} (max {}%)".format(gt['label'], gt['max']))
        self.cmb_garage.SelectedIndex = 3

        self.txt_helling.Text = "24.0"

        for ft in self.floor_types:
            param = ft.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM)
            name = param.AsString() if param else ft.Name
            self.cmb_vloer.Items.Add(name)
        if self.cmb_vloer.Items.Count > 0:
            self.cmb_vloer.SelectedIndex = 0

    def _bind_events(self):
        """Bind UI events"""
        self.cmb_garage.SelectionChanged += self._on_garage_changed
        self.txt_helling.TextChanged += self._on_helling_changed
        self.txt_hoogte.TextChanged += self._on_input_changed
        self.txt_breedte.TextChanged += self._on_breedte_changed
        self.chk_overgang.Checked += self._on_input_changed
        self.chk_overgang.Unchecked += self._on_input_changed
        self.btn_reset_helling.Click += self._on_reset_helling
        self.btn_reset_breedte.Click += self._on_reset_breedte
        self.btn_plaats.Click += self._on_plaats_click
        self.preview_canvas.SizeChanged += self._on_canvas_resized

    # ---- Canvas preview drawing ----

    def _add_line(self, x1, y1, x2, y2, brush, thickness=1, dash=None):
        """Voeg lijn toe aan preview canvas"""
        line = WpfLine()
        line.X1 = float(x1)
        line.Y1 = float(y1)
        line.X2 = float(x2)
        line.Y2 = float(y2)
        line.Stroke = brush
        line.StrokeThickness = thickness
        if dash:
            dc = DoubleCollection()
            for d in dash:
                dc.Add(float(d))
            line.StrokeDashArray = dc
        self.preview_canvas.Children.Add(line)

    def _add_text(self, text, x, y, brush, size=9, bold=False):
        """Voeg tekst toe aan preview canvas"""
        tb = WpfTextBlock()
        tb.Text = text
        tb.FontSize = size
        tb.Foreground = brush
        if bold:
            tb.FontWeight = FontWeights.Bold
        WpfCanvas.SetLeft(tb, float(x))
        WpfCanvas.SetTop(tb, float(y))
        self.preview_canvas.Children.Add(tb)

    def _add_rect(self, x, y, w, h, brush):
        """Voeg rechthoek toe aan preview canvas"""
        rect = WpfRect()
        rect.Width = float(w)
        rect.Height = float(h)
        rect.Fill = brush
        WpfCanvas.SetLeft(rect, float(x))
        WpfCanvas.SetTop(rect, float(y))
        self.preview_canvas.Children.Add(rect)

    def _update_preview(self):
        """Update preview canvas met hellingbaan visualisatie"""
        self.preview_canvas.Children.Clear()

        b = self.berekening
        if not b:
            return

        canvas_width = self.preview_canvas.ActualWidth if self.preview_canvas.ActualWidth > 0 else 440
        canvas_height = self.preview_canvas.ActualHeight if self.preview_canvas.ActualHeight > 0 else 150

        padding = 40
        draw_width = canvas_width - 2 * padding
        draw_height = canvas_height - 2 * padding - 20

        # Schaal berekenen
        schaal_x = draw_width / max(b['lengte'], 1)
        schaal_y = draw_height / max(self.hoogte, 1)
        schaal = min(schaal_x, schaal_y) * 0.85

        start_x = padding
        start_y = canvas_height - padding

        # Punten berekenen
        p1 = (start_x, start_y)
        p2 = (start_x + b['overgang_onder_lengte'] * schaal,
              start_y - b['overgang_onder_hoogte'] * schaal)
        p3 = (p2[0] + b['hoofd_lengte'] * schaal,
              p2[1] - b['hoofd_hoogte'] * schaal)
        p4 = (p3[0] + b['overgang_boven_lengte'] * schaal,
              p3[1] - b['overgang_boven_hoogte'] * schaal)

        # Kleuren
        teal_brush = Huisstijl.get_brush(Huisstijl.TEAL_HEX)
        yellow_brush = Huisstijl.get_brush(Huisstijl.YELLOW_HEX)
        violet_brush = Huisstijl.get_brush(Huisstijl.VIOLET_HEX)
        gray_brush = Huisstijl.get_brush(Huisstijl.MEDIUM_GRAY_HEX)
        sec_brush = Huisstijl.get_brush(Huisstijl.TEXT_SECONDARY_HEX)

        # Vloerlijnen (stippel)
        self._add_line(padding - 10, start_y, canvas_width - padding + 10, start_y,
                       gray_brush, 1, [4, 2])
        self._add_line(padding - 10, p4[1], canvas_width - padding + 10, p4[1],
                       gray_brush, 1, [4, 2])

        # Hellingbaan segmenten
        if self.met_overgang and b['overgang_onder_lengte'] > 0:
            self._add_line(p1[0], p1[1], p2[0], p2[1], yellow_brush, 4)
            self._add_line(p2[0], p2[1], p3[0], p3[1], teal_brush, 4)
            self._add_line(p3[0], p3[1], p4[0], p4[1], yellow_brush, 4)
        else:
            self._add_line(p1[0], p1[1], p4[0], p4[1], teal_brush, 4)

        # Maatlijnen - lengte
        y_maat = start_y + 15
        self._add_line(p1[0], y_maat, p4[0], y_maat, violet_brush, 1)
        self._add_line(p1[0], y_maat - 5, p1[0], y_maat + 5, violet_brush, 1)
        self._add_line(p4[0], y_maat - 5, p4[0], y_maat + 5, violet_brush, 1)

        lengte_tekst = "{:.2f} m".format(b['lengte'] / 1000.0)
        text_x = (p1[0] + p4[0]) / 2 - 20
        self._add_text(lengte_tekst, text_x, y_maat + 2, violet_brush, 9)

        # Maatlijnen - hoogte
        x_maat = canvas_width - padding + 15
        self._add_line(x_maat, start_y, x_maat, p4[1], violet_brush, 1)
        self._add_line(x_maat - 5, start_y, x_maat + 5, start_y, violet_brush, 1)
        self._add_line(x_maat - 5, p4[1], x_maat + 5, p4[1], violet_brush, 1)

        hoogte_tekst = "{:.2f} m".format(self.hoogte / 1000.0)
        text_y = (start_y + p4[1]) / 2
        self._add_text(hoogte_tekst, x_maat + 5, text_y, violet_brush, 9)

        # Titel
        self._add_text("Hellingbaan Preview", padding, 8, violet_brush, 11, bold=True)

        # Legenda
        legend_x = canvas_width - 160
        if self.met_overgang:
            self._add_rect(legend_x, 10, 12, 4, yellow_brush)
            self._add_text("Overgang ({:.1f}%)".format(b['overgang_helling']),
                           legend_x + 16, 6, sec_brush, 8)
        self._add_rect(legend_x, 22, 12, 4, teal_brush)
        self._add_text("Helling ({:.1f}%)".format(b['helling']),
                       legend_x + 16, 18, sec_brush, 8)

    # ---- Berekening en UI updates ----

    def _get_garage_type(self):
        """Haal geselecteerde garage type op"""
        idx = self.cmb_garage.SelectedIndex
        if 0 <= idx < len(GARAGE_TYPES):
            return GARAGE_TYPES[idx]
        return GARAGE_TYPES[3]

    def _get_hoogte(self):
        """Haal hoogte op uit textbox"""
        try:
            return float(self.txt_hoogte.Text.replace(',', '.'))
        except (ValueError, TypeError):
            return 3600

    def _get_breedte(self):
        """Haal breedte op"""
        garage = self._get_garage_type()
        if self.breedte_override is not None:
            return self.breedte_override
        return garage['breedte_min']

    def _update_berekening(self):
        """Update berekening en UI"""
        self.hoogte = self._get_hoogte()
        garage = self._get_garage_type()
        self.met_overgang = self.chk_overgang.IsChecked == True
        breedte = self._get_breedte()

        b = bereken_hellingbaan(self.hoogte, garage, self.met_overgang, self.helling_override)
        self.berekening = b
        self.berekening['breedte'] = breedte

        # Update helling textbox (zonder event triggeren)
        if not b['is_override']:
            self.txt_helling.TextChanged -= self._on_helling_changed
            self.txt_helling.Text = "{:.1f}".format(b['helling'])
            self.txt_helling.TextChanged += self._on_helling_changed
            self.btn_reset_helling.Visibility = Visibility.Collapsed
        else:
            self.btn_reset_helling.Visibility = Visibility.Visible

        # Update breedte textbox
        if self.breedte_override is None:
            self.txt_breedte.TextChanged -= self._on_breedte_changed
            self.txt_breedte.Text = str(int(garage['breedte_min']))
            self.txt_breedte.TextChanged += self._on_breedte_changed
            self.btn_reset_breedte.Visibility = Visibility.Collapsed
        else:
            self.btn_reset_breedte.Visibility = Visibility.Visible

        # Update min breedte label
        self.lbl_min_breedte.Text = "(min: {})".format(garage['breedte_min'])

        # Update preview
        self._update_preview()

        # Update zone label
        if b['is_override']:
            zone_text = "Handmatig"
        elif not self.met_overgang:
            zone_text = "Zonder overgang"
        elif b['zone'] == 'vast':
            zone_text = "Vast ({}%)".format(garage['max'])
        elif b['zone'] == 'kort':
            zone_text = "Kort ({}%)".format(garage['max'])
        elif b['zone'] == 'lang':
            zone_text = "Lang ({}%)".format(garage['min'])
        else:
            zone_text = "Geoptimaliseerd"
        self.lbl_zone.Text = zone_text

        # Update resultaat tekst
        lines = []
        lines.append("Totale lengte:       {:.2f} m".format(b['lengte'] / 1000.0))
        lines.append("Breedte:             {:.2f} m".format(breedte / 1000.0))
        lines.append("Gebruikte helling:   {:.1f}%".format(b['helling']))
        if b['is_override']:
            lines.append("Berekende helling:   {:.1f}%".format(b['helling_berekend']))
        if self.met_overgang:
            lines.append("Overgangshelling:    {:.1f}%".format(b['overgang_helling']))
            lines.append("Overgang onder:      {:.2f} m".format(b['overgang_onder_lengte'] / 1000.0))
            lines.append("Hoofdhelling:        {:.2f} m".format(b['hoofd_lengte'] / 1000.0))
            lines.append("Overgang boven:      {:.2f} m".format(b['overgang_boven_lengte'] / 1000.0))

        self.lbl_resultaat.Text = "\n".join(lines)

    # ---- Event handlers ----

    def _on_garage_changed(self, sender, args):
        self.helling_override = None
        self.breedte_override = None
        self._update_berekening()

    def _on_helling_changed(self, sender, args):
        try:
            val = float(self.txt_helling.Text.replace(',', '.'))
            if 0 < val <= 30:
                self.helling_override = val
                self._update_berekening()
        except (ValueError, TypeError):
            pass

    def _on_breedte_changed(self, sender, args):
        try:
            val = float(self.txt_breedte.Text.replace(',', '.'))
            if val > 0:
                self.breedte_override = val
                self._update_berekening()
        except (ValueError, TypeError):
            pass

    def _on_reset_helling(self, sender, args):
        self.helling_override = None
        self._update_berekening()

    def _on_reset_breedte(self, sender, args):
        self.breedte_override = None
        self._update_berekening()

    def _on_input_changed(self, sender, args):
        self._update_berekening()

    def _on_canvas_resized(self, sender, args):
        if self.berekening:
            self._update_preview()

    def _on_plaats_click(self, sender, args):
        if not self.berekening:
            self.show_warning("Geen geldige berekening")
            return

        if self.cmb_vloer.SelectedIndex < 0:
            self.show_warning("Selecteer een vloertype")
            return

        log.info("Plaats in Revit geklikt - resultaat: {}".format(self.berekening))

        self.plaats_gevraagd = True
        self.close_ok()


# ==============================================================================
# VLOER PLAATSING MET DIRECTSHAPE
# ==============================================================================
def mm_to_feet(mm):
    """Converteer mm naar Revit internal units (feet)"""
    return mm / 304.8


def create_ramp_solid(x_start, y_start, z_start, length, width, height_delta, thickness=0.3):
    """
    Maak een solide voor een hellende vloer via extrusie.
    """
    t = thickness  # vloerdikte

    p0 = XYZ(0, 0, 0)
    p1 = XYZ(length, 0, height_delta)
    p2 = XYZ(length, 0, height_delta + t)
    p3 = XYZ(0, 0, t)

    profile = CurveLoop()
    profile.Append(Line.CreateBound(p0, p1))
    profile.Append(Line.CreateBound(p1, p2))
    profile.Append(Line.CreateBound(p2, p3))
    profile.Append(Line.CreateBound(p3, p0))

    profiles = List[CurveLoop]()
    profiles.Add(profile)

    extrusion_dir = XYZ(0, 1, 0)

    solid = GeometryCreationUtilities.CreateExtrusionGeometry(
        profiles, extrusion_dir, width
    )

    return solid


def plaats_hellingbaan(document, uidoc, result, floor_type, level):
    """
    Plaats hellingbaan als DirectShape elementen in Revit.
    """
    log.info("Start plaats_hellingbaan (DirectShape methode)")

    try:
        pt = uidoc.Selection.PickPoint(
            ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections | ObjectSnapTypes.Midpoints,
            "Klik startpunt hellingbaan (linksonder)"
        )
        log.info("Punt geselecteerd: X={:.2f}, Y={:.2f}, Z={:.2f}".format(pt.X, pt.Y, pt.Z))
    except OperationCanceledException:
        log.info("Selectie geannuleerd door gebruiker")
        return None
    except Exception as e:
        log.error("Fout bij punt selectie: {}".format(e), exc_info=True)
        return None

    # Converteer maten naar feet
    breedte = mm_to_feet(result['breedte'])
    dikte = mm_to_feet(300)  # 300mm vloerdikte

    overgang_onder_len = mm_to_feet(result['overgang_onder_lengte'])
    hoofd_len = mm_to_feet(result['hoofd_lengte'])
    overgang_boven_len = mm_to_feet(result['overgang_boven_lengte'])

    overgang_onder_hoogte = mm_to_feet(result['overgang_onder_hoogte'])
    hoofd_hoogte = mm_to_feet(result['hoofd_hoogte'])
    overgang_boven_hoogte = mm_to_feet(result['overgang_boven_hoogte'])

    log.info("Breedte: {:.2f} ft, Dikte: {:.2f} ft".format(breedte, dikte))
    log.info("Hoogtes (ft): onder={:.3f}, hoofd={:.3f}, boven={:.3f}".format(
        overgang_onder_hoogte, hoofd_hoogte, overgang_boven_hoogte))

    x0, y0, z0 = pt.X, pt.Y, pt.Z

    created_shapes = []
    category_id = ElementId(BuiltInCategory.OST_Floors)

    with Transaction(document, "Plaats Hellingbaan (DirectShape)") as t:
        t.Start()
        log.info("Transaction gestart")

        if result['overgang_onder_lengte'] > 0:
            # === MET OVERGANGSHELLING: 3 segmenten ===

            # 1. Overgang onder (voetboog)
            log.info("Segment 1: Overgang onder - len={:.2f}m, hoogte={:.2f}m".format(
                overgang_onder_len * 0.3048, overgang_onder_hoogte * 0.3048))

            solid1 = create_ramp_solid(0, 0, 0, overgang_onder_len, breedte,
                                       overgang_onder_hoogte, dikte)
            transform1 = Transform.CreateTranslation(XYZ(x0, y0, z0))
            solid1_moved = SolidUtils.CreateTransformed(solid1, transform1)

            ds1 = DirectShape.CreateElement(document, category_id)
            ds1.SetShape([solid1_moved])
            ds1.SetName("Hellingbaan_Overgang_Onder")
            created_shapes.append(ds1)
            log.info("DirectShape 1 aangemaakt: {}".format(ds1.Id.IntegerValue))

            # 2. Hoofdhelling
            x1 = x0 + overgang_onder_len
            z1 = z0 + overgang_onder_hoogte

            log.info("Segment 2: Hoofdhelling - len={:.2f}m, hoogte={:.2f}m".format(
                hoofd_len * 0.3048, hoofd_hoogte * 0.3048))

            solid2 = create_ramp_solid(0, 0, 0, hoofd_len, breedte,
                                       hoofd_hoogte, dikte)
            transform2 = Transform.CreateTranslation(XYZ(x1, y0, z1))
            solid2_moved = SolidUtils.CreateTransformed(solid2, transform2)

            ds2 = DirectShape.CreateElement(document, category_id)
            ds2.SetShape([solid2_moved])
            ds2.SetName("Hellingbaan_Hoofdhelling")
            created_shapes.append(ds2)
            log.info("DirectShape 2 aangemaakt: {}".format(ds2.Id.IntegerValue))

            # 3. Overgang boven (topboog)
            x2 = x1 + hoofd_len
            z2 = z1 + hoofd_hoogte

            log.info("Segment 3: Overgang boven - len={:.2f}m, hoogte={:.2f}m".format(
                overgang_boven_len * 0.3048, overgang_boven_hoogte * 0.3048))

            solid3 = create_ramp_solid(0, 0, 0, overgang_boven_len, breedte,
                                       overgang_boven_hoogte, dikte)
            transform3 = Transform.CreateTranslation(XYZ(x2, y0, z2))
            solid3_moved = SolidUtils.CreateTransformed(solid3, transform3)

            ds3 = DirectShape.CreateElement(document, category_id)
            ds3.SetShape([solid3_moved])
            ds3.SetName("Hellingbaan_Overgang_Boven")
            created_shapes.append(ds3)
            log.info("DirectShape 3 aangemaakt: {}".format(ds3.Id.IntegerValue))

        else:
            # === ZONDER OVERGANGSHELLING: 1 segment ===
            lengte = mm_to_feet(result['lengte'])
            hoogte_totaal = mm_to_feet(result['hoofd_hoogte'])

            log.info("Enkel segment - len={:.2f}m, hoogte={:.2f}m".format(
                lengte * 0.3048, hoogte_totaal * 0.3048))

            solid = create_ramp_solid(0, 0, 0, lengte, breedte,
                                      hoogte_totaal, dikte)
            transform = Transform.CreateTranslation(XYZ(x0, y0, z0))
            solid_moved = SolidUtils.CreateTransformed(solid, transform)

            ds = DirectShape.CreateElement(document, category_id)
            ds.SetShape([solid_moved])
            ds.SetName("Hellingbaan")
            created_shapes.append(ds)
            log.info("DirectShape aangemaakt: {}".format(ds.Id.IntegerValue))

        t.Commit()
        log.info("Transaction voltooid - {} DirectShapes aangemaakt".format(len(created_shapes)))

    return created_shapes


# ==============================================================================
# REVIT DATA HELPERS
# ==============================================================================
def get_floor_types(document):
    collector = FilteredElementCollector(document).OfClass(FloorType)
    return list(collector.ToElements())


def get_levels(document):
    collector = FilteredElementCollector(document).OfClass(Level)
    return sorted(list(collector.ToElements()), key=lambda x: x.Elevation)


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    global doc, output
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()

    log.info("Hellingbaan Generator gestart")

    floor_types = get_floor_types(doc)
    levels = get_levels(doc)

    if not floor_types:
        forms.alert("Geen vloertypes gevonden in het model.", exitscript=True)

    if not levels:
        forms.alert("Geen levels gevonden in het model.", exitscript=True)

    window = HellingbaanWindow(floor_types, levels)
    window.ShowDialog()

    # Als gebruiker op "Plaats in Revit" heeft geklikt
    if window.plaats_gevraagd and window.berekening:
        log.info("Plaatsing gestart met resultaat: {}".format(window.berekening))

        # Haal geselecteerde floor type
        floor_type = floor_types[window.cmb_vloer.SelectedIndex]
        floor_type_name = window.cmb_vloer.SelectedItem
        log.info("Floor type: {}".format(floor_type_name))

        # Gebruik actieve view level of eerste level
        active_view = doc.ActiveView
        if isinstance(active_view, ViewPlan) and active_view.GenLevel:
            level = active_view.GenLevel
        else:
            level = levels[0]
        try:
            level_name = level.Name
        except Exception:
            level_name = str(level.Id.IntegerValue)
        log.info("Level: {}".format(level_name))

        # Plaats de hellingbaan
        try:
            created = plaats_hellingbaan(doc, uidoc, window.berekening, floor_type, level)

            if created:
                log.info("Hellingbaan geplaatst: {} vloeren".format(len(created)))
                forms.alert(
                    "Hellingbaan geplaatst!\n\n"
                    "{} vloer(en) aangemaakt.\n"
                    "Lengte: {:.2f} m\n"
                    "Breedte: {:.2f} m\n"
                    "Helling: {:.1f}%".format(
                        len(created),
                        window.berekening['lengte'] / 1000.0,
                        window.berekening['breedte'] / 1000.0,
                        window.berekening['helling']
                    ),
                    title="Hellingbaan Generator"
                )
        except Exception as e:
            log.error("Fout bij plaatsing: {}".format(e), exc_info=True)
            forms.alert("Fout bij plaatsing:\n{}".format(e), title="Fout")


if __name__ == "__main__":
    main()
