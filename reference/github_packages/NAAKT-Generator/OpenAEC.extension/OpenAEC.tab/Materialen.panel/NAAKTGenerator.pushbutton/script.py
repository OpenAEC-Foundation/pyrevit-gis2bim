# -*- coding: utf-8 -*-
"""
NAA.K.T. Materiaal Generator - WPF versie
Genereer eenduidige materiaalbenamingen met pattern & kleur settings.

Format: naam_kenmerk_toepassing_[eigen-invulling]
"""

__title__ = "NAA.K.T."
__author__ = "3BM Bouwkunde"
__doc__ = "Genereer materiaalbenamingen volgens NAA.K.T. standaard"

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color as WpfColor
from System.Windows.Controls import ComboBoxItem

import os
import sys
import json

# Lib path - supports both full extension and standalone installation
SCRIPT_DIR = os.path.dirname(__file__)

def find_lib_dir():
    """Find lib directory - works for both full extension and standalone"""
    # Option 1: Standalone (lib is sibling to pushbutton folder)
    standalone_lib = os.path.join(os.path.dirname(SCRIPT_DIR), 'lib')
    if os.path.exists(standalone_lib):
        return standalone_lib
    
    # Option 2: Full extension (lib is in extension root)
    extension_lib = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), 
        'lib'
    )
    if os.path.exists(extension_lib):
        return extension_lib
    
    # Fallback
    return standalone_lib

LIB_DIR = find_lib_dir()
sys.path.insert(0, LIB_DIR)

from bm_logger import get_logger

# pyRevit imports
from pyrevit import revit, forms, script

# Revit API
from Autodesk.Revit.DB import (
    FilteredElementCollector, Material, Transaction, ElementId,
    FillPatternElement, Color as RevitColor
)

log = get_logger("NAAKTGenerator")

# Data paths
DATA_DIR = os.path.join(LIB_DIR, 'naakt_data')
NAMEN_FILE = os.path.join(DATA_DIR, 'naakt_namen.json')
KENMERKEN_FILE = os.path.join(DATA_DIR, 'naakt_kenmerken.json')
TOEPASSINGEN_FILE = os.path.join(DATA_DIR, 'naakt_toepassingen.json')


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def hex_to_wpf_color(hex_color):
    """Convert hex string to WPF Color"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return WpfColor.FromRgb(r, g, b)

def hex_to_revit_color(hex_color):
    """Convert hex string to Revit Color"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RevitColor(r, g, b)

def revit_color_to_hex(color):
    """Convert Revit Color to hex string"""
    return "#{:02X}{:02X}{:02X}".format(color.Red, color.Green, color.Blue)


# ==============================================================================
# MATERIAAL KEYWORDS
# ==============================================================================
MATERIAAL_KEYWORDS = {
    'beton': ['beton', 'concrete', 'cementgebonden'],
    'hout': ['hout', 'wood', 'timber', 'eiken', 'grenen', 'multiplex', 'mdf', 'osb'],
    'isolatie': ['isolatie', 'insulation', 'eps', 'xps', 'pir', 'pur', 'minerale wol', 'glaswol', 'rotswol'],
    'metaal': ['metaal', 'metal', 'staal', 'steel', 'aluminium', 'alu', 'zink', 'koper'],
    'steen': ['steen', 'stone', 'baksteen', 'brick', 'kalkzandsteen', 'natuursteen'],
    'gips': ['gips', 'gypsum', 'gyproc', 'fermacell'],
    'glas': ['glas', 'glass', 'beglazing'],
    'folie': ['folie', 'dampremmend', 'dampopen', 'membraan', 'barrier'],
    'lucht': ['lucht', 'air', 'spouw', 'cavity'],
    'kunststof': ['kunststof', 'plastic', 'pvc', 'pe', 'hdpe', 'ldpe'],
    'bitumen': ['bitumen', 'dakbedekking', 'roofing'],
    'verf': ['verf', 'paint', 'coating', 'latex'],
    'tegels': ['tegel', 'tile', 'keramisch', 'ceramic'],
    'mortel': ['mortel', 'mortar', 'voeg', 'specie'],
}


# ==============================================================================
# DATA MANAGEMENT
# ==============================================================================
class NAAKTData:
    """Beheer NAA.K.T. data uit JSON bestanden"""
    
    def __init__(self):
        self.namen = []
        self.kenmerken_per_naam = {}
        self.toepassingen_per_naam = {}
        self.load_data()
    
    def load_data(self):
        """Laad alle JSON data"""
        try:
            if os.path.exists(NAMEN_FILE):
                with open(NAMEN_FILE, 'r') as f:
                    content = f.read()
                    if content.startswith('\xef\xbb\xbf'):
                        content = content[3:]
                    data = json.loads(content)
                    self.namen = data.get('namen', [])
            
            if os.path.exists(KENMERKEN_FILE):
                with open(KENMERKEN_FILE, 'r') as f:
                    content = f.read()
                    if content.startswith('\xef\xbb\xbf'):
                        content = content[3:]
                    data = json.loads(content)
                    self.kenmerken_per_naam = data.get('kenmerken_per_naam', {})
            
            if os.path.exists(TOEPASSINGEN_FILE):
                with open(TOEPASSINGEN_FILE, 'r') as f:
                    content = f.read()
                    if content.startswith('\xef\xbb\xbf'):
                        content = content[3:]
                    data = json.loads(content)
                    self.toepassingen_per_naam = data.get('toepassingen_per_naam', {})
                    
        except Exception as e:
            log.error("Fout bij laden data: {}".format(e))
    
    def _write_json(self, filepath, data):
        """Schrijf JSON bestand"""
        with open(filepath, 'w') as f:
            json_str = json.dumps(data, indent=2, ensure_ascii=False)
            if hasattr(json_str, 'encode'):
                f.write(json_str.encode('utf-8'))
            else:
                f.write(json_str)
    
    def save_namen(self):
        data = {
            "beschrijving": "Geldige NAA.K.T. NAMEN",
            "namen": sorted(self.namen)
        }
        self._write_json(NAMEN_FILE, data)
    
    def save_kenmerken(self):
        sorted_kenmerken = {naam: sorted(self.kenmerken_per_naam[naam]) 
                           for naam in sorted(self.kenmerken_per_naam.keys())}
        data = {
            "beschrijving": "Geldige NAA.K.T. KENMERKEN per NAAM",
            "kenmerken_per_naam": sorted_kenmerken
        }
        self._write_json(KENMERKEN_FILE, data)
    
    def save_toepassingen(self):
        sorted_toepassingen = {naam: sorted(self.toepassingen_per_naam[naam]) 
                              for naam in sorted(self.toepassingen_per_naam.keys())}
        data = {
            "beschrijving": "Geldige NAA.K.T. TOEPASSINGEN per NAAM",
            "toepassingen_per_naam": sorted_toepassingen
        }
        self._write_json(TOEPASSINGEN_FILE, data)
    
    def add_naam(self, naam):
        naam = naam.lower().strip()
        if naam and naam not in self.namen:
            self.namen.append(naam)
            self.kenmerken_per_naam[naam] = ['generiek', 'ntb']
            self.toepassingen_per_naam[naam] = ['generiek', 'ntb']
            self.save_namen()
            self.save_kenmerken()
            self.save_toepassingen()
            return True
        return False
    
    def add_kenmerk(self, naam, kenmerk):
        kenmerk = kenmerk.lower().strip()
        if naam in self.kenmerken_per_naam and kenmerk:
            if kenmerk not in self.kenmerken_per_naam[naam]:
                self.kenmerken_per_naam[naam].append(kenmerk)
                self.save_kenmerken()
                return True
        return False
    
    def add_toepassing(self, naam, toepassing):
        toepassing = toepassing.lower().strip()
        if naam in self.toepassingen_per_naam and toepassing:
            if toepassing not in self.toepassingen_per_naam[naam]:
                self.toepassingen_per_naam[naam].append(toepassing)
                self.save_toepassingen()
                return True
        return False
    
    def get_kenmerken(self, naam):
        return sorted(self.kenmerken_per_naam.get(naam, []))
    
    def get_toepassingen(self, naam):
        return sorted(self.toepassingen_per_naam.get(naam, []))


# ==============================================================================
# PATTERN HELPER
# ==============================================================================
class PatternHelper:
    """Helper voor Revit fill patterns"""
    
    @staticmethod
    def get_all_patterns():
        """Haal alle fill patterns op uit het model"""
        collector = FilteredElementCollector(revit.doc).OfClass(FillPatternElement)
        patterns = {}
        for pattern in collector:
            try:
                fp = pattern.GetFillPattern()
                if fp:
                    patterns[pattern.Name] = pattern.Id
            except:
                pass
        return patterns
    
    @staticmethod
    def get_solid_pattern_id():
        """Haal het Solid fill pattern ID op"""
        collector = FilteredElementCollector(revit.doc).OfClass(FillPatternElement)
        for pattern in collector:
            try:
                fp = pattern.GetFillPattern()
                if fp and fp.IsSolidFill:
                    return pattern.Id
            except:
                pass
        return ElementId.InvalidElementId
    
    @staticmethod
    def find_pattern_by_name(name):
        """Zoek pattern op naam (case-insensitive)"""
        collector = FilteredElementCollector(revit.doc).OfClass(FillPatternElement)
        name_lower = name.lower()
        for pattern in collector:
            if pattern.Name.lower() == name_lower:
                return pattern.Id
            # Partial match
            if name_lower in pattern.Name.lower():
                return pattern.Id
        return None


# ==============================================================================
# MATERIAAL HELPER
# ==============================================================================
class MateriaalHelper:
    """Helper voor Revit materiaal operaties"""
    
    @staticmethod
    def get_all_materials():
        collector = FilteredElementCollector(revit.doc).OfClass(Material)
        return list(collector)
    
    @staticmethod
    def find_closest_material(naakt_naam):
        """Zoek materiaal dat best matcht met NAA.K.T. naam"""
        materials = MateriaalHelper.get_all_materials()
        if not materials:
            return None
        
        keywords = MATERIAAL_KEYWORDS.get(naakt_naam.lower(), [naakt_naam.lower()])
        best_match = None
        best_score = 0
        
        for mat in materials:
            mat_name = mat.Name.lower()
            score = 0
            for keyword in keywords:
                if keyword in mat_name:
                    if mat_name == keyword:
                        score += 10
                    elif mat_name.startswith(keyword):
                        score += 5
                    else:
                        score += 2
            
            if score > best_score:
                best_score = score
                best_match = mat
        
        if not best_match and materials:
            best_match = materials[0]
        
        return best_match
    
    @staticmethod
    def material_exists(name):
        materials = FilteredElementCollector(revit.doc).OfClass(Material)
        for mat in materials:
            if mat.Name == name:
                return True
        return False
    
    @staticmethod
    def duplicate_material(source_material, new_name):
        """Dupliceer materiaal met assets"""
        if not source_material:
            return None
        
        try:
            doc = revit.doc
            existing = FilteredElementCollector(doc).OfClass(Material)
            for mat in existing:
                if mat.Name == new_name:
                    return None
            
            # Material.Duplicate() geeft direct het nieuwe Material object terug
            new_mat = source_material.Duplicate(new_name)
            
            if not new_mat:
                return None
            
            # Dupliceer Appearance Asset
            appearance_asset_id = new_mat.AppearanceAssetId
            if appearance_asset_id and appearance_asset_id != ElementId.InvalidElementId:
                appearance_asset = doc.GetElement(appearance_asset_id)
                if appearance_asset:
                    try:
                        # AppearanceAssetElement.Duplicate() geeft ook direct element terug
                        new_asset = appearance_asset.Duplicate(new_name)
                        if new_asset:
                            new_mat.AppearanceAssetId = new_asset.Id
                    except Exception as ex:
                        log.warning("Asset duplicate failed: {}".format(ex))
            
            return new_mat
            
        except Exception as e:
            log.error("Duplicate error: {}".format(e))
            return None


# ==============================================================================
# MAIN WINDOW
# ==============================================================================
class NAAKTGeneratorWindow(Window):
    """NAA.K.T. Materiaal Generator - WPF"""
    
    def __init__(self):
        Window.__init__(self)
        self.data = NAAKTData()
        self.patterns = PatternHelper.get_all_patterns()
        
        # Color settings (hex)
        self.surface_fg_color = "#323232"
        self.surface_bg_color = "#C8C8C8"
        self.cut_fg_color = "#323232"
        self.cut_bg_color = "#B4B4B4"
        
        self._load_xaml()
        self._setup_controls()
        self._bind_events()
        self._fill_data()
    
    def _load_xaml(self):
        """Laad XAML layout"""
        xaml_path = os.path.join(SCRIPT_DIR, 'UI.xaml')
        
        try:
            with open(xaml_path, 'r') as f:
                xaml_content = f.read()
            
            reader = StringReader(xaml_content)
            loaded = XamlReader.Load(SysXmlReader.Create(reader))
            
            # Copy window properties
            self.Title = loaded.Title
            self.Width = loaded.Width
            self.Height = loaded.Height
            self.WindowStartupLocation = loaded.WindowStartupLocation
            self.ResizeMode = loaded.ResizeMode
            self.Background = loaded.Background
            self.Content = loaded.Content
            
            # Bind all named elements
            self._bind_elements(loaded)
            
            log.info("XAML loaded successfully")
            
        except Exception as e:
            log.error("XAML load error: {}".format(e), exc_info=True)
            raise
    
    def _bind_elements(self, loaded):
        """Bind all named XAML elements"""
        element_names = [
            'txt_subtitle', 'txt_preview',
            'cmb_naam', 'cmb_kenmerk', 'cmb_toepassing', 'txt_eigen',
            'btn_add_naam', 'btn_add_kenmerk', 'btn_add_toepassing',
            'cmb_surface_fg', 'cmb_cut_fg',
            'btn_surface_fg_color', 'btn_surface_bg_color',
            'btn_cut_fg_color', 'btn_cut_bg_color',
            'rect_surface_fg', 'rect_surface_bg', 'rect_cut_fg', 'rect_cut_bg',
            'txt_surface_fg_color', 'txt_surface_bg_color',
            'txt_cut_fg_color', 'txt_cut_bg_color',
            'preview_surface', 'preview_cut',
            'txt_surface_preview', 'txt_cut_preview',
            'btn_copy', 'btn_create', 'tab_main'
        ]
        
        for name in element_names:
            element = loaded.FindName(name)
            if element:
                setattr(self, name, element)
            else:
                log.warning("Element not found: {}".format(name))
    
    def _setup_controls(self):
        """Setup initial control values"""
        # Fill pattern comboboxes
        pattern_names = sorted(self.patterns.keys())
        
        for cmb in [self.cmb_surface_fg, self.cmb_cut_fg]:
            if cmb:
                cmb.Items.Clear()
                for name in pattern_names:
                    cmb.Items.Add(name)
                # Select Solid fill if available
                for i, name in enumerate(pattern_names):
                    if 'solid' in name.lower():
                        cmb.SelectedIndex = i
                        break
                if cmb.SelectedIndex < 0 and pattern_names:
                    cmb.SelectedIndex = 0
    
    def _bind_events(self):
        """Bind event handlers"""
        # NAA.K.T. events
        if self.cmb_naam:
            self.cmb_naam.SelectionChanged += self._on_naam_changed
        if self.cmb_kenmerk:
            self.cmb_kenmerk.SelectionChanged += self._on_update_preview
        if self.cmb_toepassing:
            self.cmb_toepassing.SelectionChanged += self._on_update_preview
        if self.txt_eigen:
            self.txt_eigen.TextChanged += self._on_update_preview
        
        # Add buttons
        if self.btn_add_naam:
            self.btn_add_naam.Click += self._on_add_naam
        if self.btn_add_kenmerk:
            self.btn_add_kenmerk.Click += self._on_add_kenmerk
        if self.btn_add_toepassing:
            self.btn_add_toepassing.Click += self._on_add_toepassing
        
        # Color buttons
        if self.btn_surface_fg_color:
            self.btn_surface_fg_color.Click += lambda s, e: self._pick_color('surface_fg')
        if self.btn_surface_bg_color:
            self.btn_surface_bg_color.Click += lambda s, e: self._pick_color('surface_bg')
        if self.btn_cut_fg_color:
            self.btn_cut_fg_color.Click += lambda s, e: self._pick_color('cut_fg')
        if self.btn_cut_bg_color:
            self.btn_cut_bg_color.Click += lambda s, e: self._pick_color('cut_bg')
        
        # Pattern comboboxes
        if self.cmb_surface_fg:
            self.cmb_surface_fg.SelectionChanged += self._on_surface_pattern_changed
        if self.cmb_cut_fg:
            self.cmb_cut_fg.SelectionChanged += self._on_cut_pattern_changed
        
        # Footer buttons
        if self.btn_copy:
            self.btn_copy.Click += self._on_copy
        if self.btn_create:
            self.btn_create.Click += self._on_create
    
    def _fill_data(self):
        """Fill initial data"""
        # Fill NAAM combobox
        if self.cmb_naam:
            self.cmb_naam.Items.Clear()
            for naam in sorted(self.data.namen):
                self.cmb_naam.Items.Add(naam)
            if self.cmb_naam.Items.Count > 0:
                self.cmb_naam.SelectedIndex = 0
    
    def _on_naam_changed(self, sender, args):
        """NAAM selection changed"""
        if not self.cmb_naam or self.cmb_naam.SelectedIndex < 0:
            return
        
        naam = str(self.cmb_naam.SelectedItem)
        
        # Fill KENMERK
        if self.cmb_kenmerk:
            self.cmb_kenmerk.Items.Clear()
            for kenmerk in self.data.get_kenmerken(naam):
                self.cmb_kenmerk.Items.Add(kenmerk)
            # Select 'generiek' if available
            for i in range(self.cmb_kenmerk.Items.Count):
                if str(self.cmb_kenmerk.Items[i]) == 'generiek':
                    self.cmb_kenmerk.SelectedIndex = i
                    break
            else:
                if self.cmb_kenmerk.Items.Count > 0:
                    self.cmb_kenmerk.SelectedIndex = 0
        
        # Fill TOEPASSING
        if self.cmb_toepassing:
            self.cmb_toepassing.Items.Clear()
            for toepassing in self.data.get_toepassingen(naam):
                self.cmb_toepassing.Items.Add(toepassing)
            for i in range(self.cmb_toepassing.Items.Count):
                if str(self.cmb_toepassing.Items[i]) == 'generiek':
                    self.cmb_toepassing.SelectedIndex = i
                    break
            else:
                if self.cmb_toepassing.Items.Count > 0:
                    self.cmb_toepassing.SelectedIndex = 0
        
        self._update_preview()
    
    def _on_update_preview(self, sender, args):
        self._update_preview()
    
    def _update_preview(self):
        """Update preview name"""
        parts = []
        
        if self.cmb_naam and self.cmb_naam.SelectedIndex >= 0:
            parts.append(str(self.cmb_naam.SelectedItem))
        if self.cmb_kenmerk and self.cmb_kenmerk.SelectedIndex >= 0:
            parts.append(str(self.cmb_kenmerk.SelectedItem))
        if self.cmb_toepassing and self.cmb_toepassing.SelectedIndex >= 0:
            parts.append(str(self.cmb_toepassing.SelectedItem))
        
        if self.txt_eigen:
            eigen = self.txt_eigen.Text.strip().lower().replace(' ', '_')
            if eigen:
                parts.append(eigen)
        
        if self.txt_preview:
            if len(parts) >= 3:
                self.txt_preview.Text = '_'.join(parts)
                self.txt_preview.Foreground = SolidColorBrush(hex_to_wpf_color("#350E35"))
            else:
                self.txt_preview.Text = '(selecteer naam, kenmerk en toepassing)'
                self.txt_preview.Foreground = SolidColorBrush(hex_to_wpf_color("#808080"))
    
    def _get_generated_name(self):
        """Get generated name if valid"""
        if self.txt_preview:
            text = self.txt_preview.Text
            if text and not text.startswith('('):
                return text
        return None
    
    def _pick_color(self, target):
        """Open color picker dialog"""
        try:
            # Use Windows Forms ColorDialog
            clr.AddReference('System.Windows.Forms')
            from System.Windows.Forms import ColorDialog, DialogResult
            from System.Drawing import Color as DrawingColor
            
            dialog = ColorDialog()
            dialog.FullOpen = True
            
            # Set current color
            current_hex = getattr(self, target + '_color', '#808080')
            current_hex = current_hex.lstrip('#')
            r = int(current_hex[0:2], 16)
            g = int(current_hex[2:4], 16)
            b = int(current_hex[4:6], 16)
            dialog.Color = DrawingColor.FromArgb(r, g, b)
            
            if dialog.ShowDialog() == DialogResult.OK:
                new_color = dialog.Color
                new_hex = "#{:02X}{:02X}{:02X}".format(new_color.R, new_color.G, new_color.B)
                
                # Update stored color
                setattr(self, target + '_color', new_hex)
                
                # Update UI
                self._update_color_ui(target, new_hex)
                
        except Exception as e:
            log.error("Color picker error: {}".format(e))
    
    def _update_color_ui(self, target, hex_color):
        """Update color button UI"""
        wpf_color = hex_to_wpf_color(hex_color)
        brush = SolidColorBrush(wpf_color)
        
        # Determine text color (light or dark)
        hex_clean = hex_color.lstrip('#')
        r = int(hex_clean[0:2], 16)
        g = int(hex_clean[2:4], 16)
        b = int(hex_clean[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        text_color = "#FFFFFF" if luminance < 0.5 else "#323232"
        
        if target == 'surface_fg':
            if self.rect_surface_fg:
                self.rect_surface_fg.Background = brush
            if self.txt_surface_fg_color:
                self.txt_surface_fg_color.Text = hex_color.upper()
            if self.btn_surface_fg_color:
                self.btn_surface_fg_color.Background = brush
                
        elif target == 'surface_bg':
            if self.rect_surface_bg:
                self.rect_surface_bg.Background = brush
            if self.txt_surface_bg_color:
                self.txt_surface_bg_color.Text = hex_color.upper()
            if self.btn_surface_bg_color:
                self.btn_surface_bg_color.Background = brush
            if self.preview_surface:
                self.preview_surface.Background = brush
                
        elif target == 'cut_fg':
            if self.rect_cut_fg:
                self.rect_cut_fg.Background = brush
            if self.txt_cut_fg_color:
                self.txt_cut_fg_color.Text = hex_color.upper()
            if self.btn_cut_fg_color:
                self.btn_cut_fg_color.Background = brush
                
        elif target == 'cut_bg':
            if self.rect_cut_bg:
                self.rect_cut_bg.Background = brush
            if self.txt_cut_bg_color:
                self.txt_cut_bg_color.Text = hex_color.upper()
            if self.btn_cut_bg_color:
                self.btn_cut_bg_color.Background = brush
            if self.preview_cut:
                self.preview_cut.Background = brush
    
    def _on_surface_pattern_changed(self, sender, args):
        """Surface pattern selection changed"""
        if self.cmb_surface_fg and self.cmb_surface_fg.SelectedIndex >= 0:
            pattern_name = str(self.cmb_surface_fg.SelectedItem)
            if self.txt_surface_preview:
                self.txt_surface_preview.Text = pattern_name
    
    def _on_cut_pattern_changed(self, sender, args):
        """Cut pattern selection changed"""
        if self.cmb_cut_fg and self.cmb_cut_fg.SelectedIndex >= 0:
            pattern_name = str(self.cmb_cut_fg.SelectedItem)
            if self.txt_cut_preview:
                self.txt_cut_preview.Text = pattern_name
    
    def _on_add_naam(self, sender, args):
        """Add new NAAM"""
        result = forms.ask_for_string(
            prompt="Voer nieuwe NAAM (hoofdgroep) in:",
            title="Nieuwe NAAM toevoegen"
        )
        if result:
            if self.data.add_naam(result):
                self._fill_data()
                # Select new naam
                for i in range(self.cmb_naam.Items.Count):
                    if str(self.cmb_naam.Items[i]) == result.lower():
                        self.cmb_naam.SelectedIndex = i
                        break
                MessageBox.Show("NAAM '{}' toegevoegd!".format(result.lower()), 
                               "Succes", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                MessageBox.Show("NAAM bestaat al of is ongeldig.", 
                               "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
    
    def _on_add_kenmerk(self, sender, args):
        """Add new KENMERK"""
        if not self.cmb_naam or self.cmb_naam.SelectedIndex < 0:
            MessageBox.Show("Selecteer eerst een NAAM.", 
                           "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        naam = str(self.cmb_naam.SelectedItem)
        result = forms.ask_for_string(
            prompt="Voer nieuw KENMERK in voor '{}':".format(naam),
            title="Nieuw KENMERK toevoegen"
        )
        if result:
            if self.data.add_kenmerk(naam, result):
                self._on_naam_changed(None, None)
                for i in range(self.cmb_kenmerk.Items.Count):
                    if str(self.cmb_kenmerk.Items[i]) == result.lower():
                        self.cmb_kenmerk.SelectedIndex = i
                        break
                MessageBox.Show("KENMERK '{}' toegevoegd!".format(result.lower()), 
                               "Succes", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                MessageBox.Show("KENMERK bestaat al of is ongeldig.", 
                               "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
    
    def _on_add_toepassing(self, sender, args):
        """Add new TOEPASSING"""
        if not self.cmb_naam or self.cmb_naam.SelectedIndex < 0:
            MessageBox.Show("Selecteer eerst een NAAM.", 
                           "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        naam = str(self.cmb_naam.SelectedItem)
        result = forms.ask_for_string(
            prompt="Voer nieuwe TOEPASSING in voor '{}':".format(naam),
            title="Nieuwe TOEPASSING toevoegen"
        )
        if result:
            if self.data.add_toepassing(naam, result):
                self._on_naam_changed(None, None)
                for i in range(self.cmb_toepassing.Items.Count):
                    if str(self.cmb_toepassing.Items[i]) == result.lower():
                        self.cmb_toepassing.SelectedIndex = i
                        break
                MessageBox.Show("TOEPASSING '{}' toegevoegd!".format(result.lower()), 
                               "Succes", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                MessageBox.Show("TOEPASSING bestaat al of is ongeldig.", 
                               "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
    
    def _on_copy(self, sender, args):
        """Copy to clipboard"""
        name = self._get_generated_name()
        if name:
            try:
                clr.AddReference('System.Windows.Forms')
                from System.Windows.Forms import Clipboard
                Clipboard.SetText(name)
                MessageBox.Show("Gekopieerd naar clipboard:\n\n{}".format(name), 
                               "Gekopieerd!", MessageBoxButton.OK, MessageBoxImage.Information)
            except Exception as e:
                log.error("Clipboard error: {}".format(e))
        else:
            MessageBox.Show("Geen geldige materiaalnaam om te kopieren.", 
                           "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
    
    def _on_create(self, sender, args):
        """Create material in Revit"""
        new_name = self._get_generated_name()
        if not new_name:
            MessageBox.Show("Geen geldige materiaalnaam. Selecteer naam, kenmerk en toepassing.", 
                           "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Check if exists
        if MateriaalHelper.material_exists(new_name):
            MessageBox.Show("Materiaal '{}' bestaat al in het model.".format(new_name), 
                           "Waarschuwing", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Find source material
        naakt_naam = str(self.cmb_naam.SelectedItem) if self.cmb_naam.SelectedIndex >= 0 else None
        source_mat = MateriaalHelper.find_closest_material(naakt_naam)
        
        if not source_mat:
            MessageBox.Show("Geen bronmateriaal gevonden in het model.", 
                           "Fout", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        # Confirm
        result = MessageBox.Show(
            "Materiaal '{}' aanmaken?\n\nBron: '{}'\n\nHet materiaal wordt aangemaakt met de geselecteerde patterns en kleuren.".format(
                new_name, source_mat.Name
            ),
            "Materiaal aanmaken",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        )
        
        if result != MessageBoxResult.Yes:
            return
        
        # Create material
        try:
            with Transaction(revit.doc, "NAA.K.T. Materiaal aanmaken") as t:
                t.Start()
                
                new_mat = MateriaalHelper.duplicate_material(source_mat, new_name)
                
                if new_mat:
                    # Apply patterns and colors
                    self._apply_patterns(new_mat)
                    
                    t.Commit()
                    
                    MessageBox.Show(
                        "Materiaal aangemaakt!\n\nNaam: {}\nBron: {}\n\nHet materiaal is nu beschikbaar in de Material Browser.".format(
                            new_name, source_mat.Name
                        ),
                        "Succes",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    )
                else:
                    t.RollBack()
                    MessageBox.Show("Kon materiaal niet aanmaken.", 
                                   "Fout", MessageBoxButton.OK, MessageBoxImage.Error)
                    
        except Exception as e:
            log.error("Create material error: {}".format(e), exc_info=True)
            MessageBox.Show("Fout bij aanmaken materiaal:\n\n{}".format(str(e)), 
                           "Fout", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _apply_patterns(self, material):
        """Apply pattern and color settings to material"""
        try:
            # Surface patterns
            if self.cmb_surface_fg and self.cmb_surface_fg.SelectedIndex >= 0:
                pattern_name = str(self.cmb_surface_fg.SelectedItem)
                pattern_id = self.patterns.get(pattern_name)
                if pattern_id:
                    material.SurfaceForegroundPatternId = pattern_id
            
            # Surface colors
            material.SurfaceForegroundPatternColor = hex_to_revit_color(self.surface_fg_color)
            material.SurfaceBackgroundPatternColor = hex_to_revit_color(self.surface_bg_color)
            
            # Set solid fill for background
            solid_id = PatternHelper.get_solid_pattern_id()
            if solid_id != ElementId.InvalidElementId:
                material.SurfaceBackgroundPatternId = solid_id
            
            # Cut patterns
            if self.cmb_cut_fg and self.cmb_cut_fg.SelectedIndex >= 0:
                pattern_name = str(self.cmb_cut_fg.SelectedItem)
                pattern_id = self.patterns.get(pattern_name)
                if pattern_id:
                    material.CutForegroundPatternId = pattern_id
            
            # Cut colors
            material.CutForegroundPatternColor = hex_to_revit_color(self.cut_fg_color)
            material.CutBackgroundPatternColor = hex_to_revit_color(self.cut_bg_color)
            
            # Set solid fill for cut background
            if solid_id != ElementId.InvalidElementId:
                material.CutBackgroundPatternId = solid_id
            
            log.info("Patterns applied successfully")
            
        except Exception as e:
            log.warning("Could not apply all patterns: {}".format(e))


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    log.info("=== NAA.K.T. Generator Started (WPF) ===")
    
    try:
        window = NAAKTGeneratorWindow()
        window.ShowDialog()
    except Exception as e:
        log.error("Error: {}".format(e), exc_info=True)
        forms.alert("Fout bij laden NAA.K.T. Generator:\n\n{}".format(e), title="Error")
    
    log.info("=== NAA.K.T. Generator Closed ===")


if __name__ == "__main__":
    main()
