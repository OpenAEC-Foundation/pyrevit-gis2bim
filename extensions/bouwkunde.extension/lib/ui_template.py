# -*- coding: utf-8 -*-
"""
3BM Bouwkunde - UI Template Module
==================================
Schaalbare Windows Forms UI voor pyRevit tools.
Werkt automatisch op HD (1080p) tot 4K displays.

Auteur: 3BM Bouwkunde
Versie: 1.3 - Fixed header text positioning
"""

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, Button, Panel, GroupBox, TextBox, ComboBox,
    DataGridView, DataGridViewTextBoxColumn, DataGridViewSelectionMode,
    FormStartPosition, FormBorderStyle, FlatStyle, BorderStyle,
    DockStyle, AnchorStyles, AutoScaleMode, Padding, ScrollBars,
    ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon,
    SaveFileDialog, OpenFileDialog, DialogResult, CheckBox
)
from System.Drawing import (
    Point, Size, Color, Font, FontStyle, Graphics,
    ContentAlignment, SystemFonts
)
from System import IntPtr


# ==============================================================================
# 3BM HUISSTIJL
# ==============================================================================
class Huisstijl:
    """3BM Bouwkunde huisstijl kleuren"""
    
    VIOLET = Color.FromArgb(53, 14, 53)
    TEAL = Color.FromArgb(69, 182, 168)
    YELLOW = Color.FromArgb(239, 189, 117)
    MAGENTA = Color.FromArgb(160, 28, 72)
    PEACH = Color.FromArgb(219, 76, 64)
    
    WHITE = Color.White
    LIGHT_GRAY = Color.FromArgb(245, 245, 245)
    MEDIUM_GRAY = Color.FromArgb(180, 180, 180)
    DARK_GRAY = Color.FromArgb(100, 100, 100)
    TEXT_PRIMARY = Color.FromArgb(50, 50, 50)
    TEXT_SECONDARY = Color.FromArgb(128, 128, 128)
    
    GRAPH_TEMP = Color.FromArgb(220, 60, 40)
    GRAPH_PSAT = Color.FromArgb(40, 120, 200)
    GRAPH_PVAP = Color.FromArgb(40, 180, 120)
    
    MATERIAL_COLORS = {
        'isolatie': Color.FromArgb(255, 230, 150),
        'beton': Color.FromArgb(180, 180, 180),
        'hout': Color.FromArgb(205, 170, 125),
        'steen': Color.FromArgb(200, 100, 100),
        'folie': Color.FromArgb(100, 150, 255),
        'gips': Color.FromArgb(245, 245, 245),
        'metaal': Color.FromArgb(160, 160, 180),
        'lucht': Color.FromArgb(220, 240, 255),
        'default': Color.FromArgb(200, 200, 200),
    }
    
    @staticmethod
    def get_material_color(material_name):
        name_lower = material_name.lower()
        if any(x in name_lower for x in ['lucht', 'spouw', 'air', 'cavity']):
            return Huisstijl.MATERIAL_COLORS['lucht']
        elif any(x in name_lower for x in ['isolatie', 'pir', 'pur', 'eps', 'xps', 'mineral', 'wool']):
            return Huisstijl.MATERIAL_COLORS['isolatie']
        elif any(x in name_lower for x in ['beton', 'concrete']):
            return Huisstijl.MATERIAL_COLORS['beton']
        elif any(x in name_lower for x in ['hout', 'wood', 'timber', 'osb', 'multiplex']):
            return Huisstijl.MATERIAL_COLORS['hout']
        elif any(x in name_lower for x in ['steen', 'brick', 'metsel', 'klinker']):
            return Huisstijl.MATERIAL_COLORS['steen']
        elif any(x in name_lower for x in ['folie', 'damp', 'membrane', 'pe']):
            return Huisstijl.MATERIAL_COLORS['folie']
        elif any(x in name_lower for x in ['gips', 'gypsum', 'plaster']):
            return Huisstijl.MATERIAL_COLORS['gips']
        elif any(x in name_lower for x in ['staal', 'steel', 'metaal', 'metal']):
            return Huisstijl.MATERIAL_COLORS['metaal']
        return Huisstijl.MATERIAL_COLORS['default']


# ==============================================================================
# DPI SCALING - FIXED FOR IRONPYTHON
# ==============================================================================
class DPIScaler:
    """
    Helper voor DPI-aware scaling.
    Gebruikt .NET Graphics.DpiX voor betrouwbare detectie in IronPython.
    """
    
    _base_dpi = 96.0
    _current_dpi = None
    _scale_factor = None
    
    @classmethod
    def _init_dpi(cls):
        """Initialiseer DPI via .NET Graphics (werkt in IronPython/Revit)"""
        if cls._current_dpi is None:
            try:
                # Via Graphics object van desktop - meest betrouwbaar in .NET
                g = Graphics.FromHwnd(IntPtr.Zero)
                cls._current_dpi = g.DpiX
                g.Dispose()
            except:
                cls._current_dpi = 96.0
            
            cls._scale_factor = cls._current_dpi / cls._base_dpi
    
    @classmethod
    def get_scale_factor(cls):
        """Haal scale factor: 1.0=100%, 1.5=150%, 2.0=200%"""
        cls._init_dpi()
        return cls._scale_factor
    
    @classmethod
    def get_dpi(cls):
        """Haal huidige DPI waarde"""
        cls._init_dpi()
        return cls._current_dpi
    
    @classmethod
    def scale(cls, value):
        """Schaal pixel waarde naar huidige DPI"""
        cls._init_dpi()
        return int(value * cls._scale_factor)
    
    @classmethod
    def scale_size(cls, width, height):
        """Schaal Size naar huidige DPI"""
        return Size(cls.scale(width), cls.scale(height))
    
    @classmethod
    def scale_point(cls, x, y):
        """Schaal Point naar huidige DPI"""
        return Point(cls.scale(x), cls.scale(y))


# ==============================================================================
# UI FACTORY
# ==============================================================================
class UIFactory:
    """Factory voor 3BM-gestylde UI controls"""
    
    FONT_TITLE = 18
    FONT_SUBTITLE = 14
    FONT_HEADING = 12
    FONT_NORMAL = 10
    FONT_SMALL = 9
    FONT_TINY = 8
    
    @staticmethod
    def create_label(text, font_size=None, bold=False, color=None, italic=False):
        lbl = Label()
        lbl.Text = text
        lbl.AutoSize = True
        
        size = font_size or UIFactory.FONT_NORMAL
        style = FontStyle.Regular
        if bold and italic:
            style = FontStyle.Bold | FontStyle.Italic
        elif bold:
            style = FontStyle.Bold
        elif italic:
            style = FontStyle.Italic
        
        lbl.Font = Font("Segoe UI", size, style)
        lbl.ForeColor = color or Huisstijl.TEXT_PRIMARY
        return lbl
    
    @staticmethod
    def create_button(text, width=100, height=40, style='primary'):
        btn = Button()
        btn.Text = text
        btn.Size = DPIScaler.scale_size(width, height)
        btn.Font = Font("Segoe UI", UIFactory.FONT_NORMAL, 
                       FontStyle.Bold if style == 'primary' else FontStyle.Regular)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 1
        
        if style == 'primary':
            btn.BackColor = Huisstijl.TEAL
            btn.ForeColor = Color.White
            btn.FlatAppearance.BorderColor = Huisstijl.TEAL
        elif style == 'secondary':
            btn.BackColor = Color.White
            btn.ForeColor = Huisstijl.VIOLET
            btn.FlatAppearance.BorderColor = Huisstijl.VIOLET
        elif style == 'warning':
            btn.BackColor = Huisstijl.YELLOW
            btn.ForeColor = Huisstijl.VIOLET
            btn.FlatAppearance.BorderColor = Huisstijl.YELLOW
        elif style == 'danger':
            btn.BackColor = Huisstijl.PEACH
            btn.ForeColor = Color.White
            btn.FlatAppearance.BorderColor = Huisstijl.PEACH
        elif style == 'icon':
            btn.Size = DPIScaler.scale_size(45, 40)
            btn.Font = Font("Segoe UI", 14)
            btn.BackColor = Color.White
            btn.ForeColor = Huisstijl.VIOLET
            btn.FlatAppearance.BorderColor = Huisstijl.MEDIUM_GRAY
        
        return btn
    
    @staticmethod
    def create_textbox(width=150, height=28, multiline=False, readonly=False):
        txt = TextBox()
        txt.Size = DPIScaler.scale_size(width, 28 if not multiline else height)
        txt.Font = Font("Segoe UI", UIFactory.FONT_NORMAL)
        txt.Multiline = multiline
        txt.ReadOnly = readonly
        if multiline:
            txt.ScrollBars = ScrollBars.Vertical
        if readonly:
            txt.BackColor = Huisstijl.LIGHT_GRAY
        return txt
    
    @staticmethod
    def create_combobox(width=200, items=None, editable=False):
        cmb = ComboBox()
        cmb.Size = DPIScaler.scale_size(width, 28)
        cmb.Font = Font("Segoe UI", UIFactory.FONT_NORMAL)
        cmb.DropDownStyle = ComboBoxStyle.DropDown if editable else ComboBoxStyle.DropDownList
        cmb.DropDownHeight = DPIScaler.scale(400)
        
        if items:
            for item in items:
                cmb.Items.Add(item)
            if cmb.Items.Count > 0:
                cmb.SelectedIndex = 0
        return cmb
    
    @staticmethod
    def create_checkbox(text, checked=False):
        chk = CheckBox()
        chk.Text = text
        chk.AutoSize = True
        chk.Font = Font("Segoe UI", UIFactory.FONT_NORMAL)
        chk.ForeColor = Huisstijl.TEXT_PRIMARY
        chk.Checked = checked
        return chk
    
    @staticmethod
    def create_groupbox(text, width=400, height=150):
        gb = GroupBox()
        gb.Text = text
        gb.Size = DPIScaler.scale_size(width, height)
        gb.Font = Font("Segoe UI", UIFactory.FONT_HEADING, FontStyle.Bold)
        gb.ForeColor = Huisstijl.VIOLET
        return gb
    
    @staticmethod
    def create_panel(width=400, height=150, border=False):
        pnl = Panel()
        pnl.Size = DPIScaler.scale_size(width, height)
        pnl.BackColor = Color.White
        if border:
            pnl.BorderStyle = BorderStyle.FixedSingle
        return pnl
    
    @staticmethod
    def create_datagridview(columns, width=800, height=300, allow_edit=False):
        grid = DataGridView()
        grid.Size = DPIScaler.scale_size(width, height)
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.ReadOnly = not allow_edit
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.BackgroundColor = Color.White
        grid.BorderStyle = BorderStyle.FixedSingle
        grid.RowHeadersVisible = False
        grid.EnableHeadersVisualStyles = False
        
        grid.ColumnHeadersDefaultCellStyle.BackColor = Huisstijl.VIOLET
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        grid.ColumnHeadersDefaultCellStyle.Font = Font("Segoe UI", UIFactory.FONT_SMALL, FontStyle.Bold)
        grid.ColumnHeadersHeight = DPIScaler.scale(32)
        grid.RowTemplate.Height = DPIScaler.scale(28)
        grid.AlternatingRowsDefaultCellStyle.BackColor = Huisstijl.LIGHT_GRAY
        
        for name, header, col_width in columns:
            col = DataGridViewTextBoxColumn()
            col.Name = name
            col.HeaderText = header
            col.Width = DPIScaler.scale(col_width)
            grid.Columns.Add(col)
        
        return grid


# ==============================================================================
# BASE FORM
# ==============================================================================
class BaseForm(Form):
    """Basis form met 3BM styling en DPI scaling"""
    
    MARGIN = 20
    HEADER_HEIGHT = 85
    ACCENT_HEIGHT = 5
    FOOTER_HEIGHT = 60
    
    def __init__(self, title, width=800, height=600, show_header=True, show_footer=True):
        self._title = title
        self._base_width = width
        self._base_height = height
        self._show_header = show_header
        self._show_footer = show_footer
        
        self._header_h = DPIScaler.scale(self.HEADER_HEIGHT) if show_header else 0
        self._accent_h = DPIScaler.scale(self.ACCENT_HEIGHT) if show_header else 0
        self._footer_h = DPIScaler.scale(self.FOOTER_HEIGHT) if show_footer else 0
        self._margin = DPIScaler.scale(self.MARGIN)
        
        self._setup_form()
        if show_header:
            self._create_header()
        if show_footer:
            self._create_footer()
        self._create_content_panel()
        
        self.Resize += self._on_form_resize
    
    def _setup_form(self):
        self.Text = self._title + " - 3BM"
        self.Size = DPIScaler.scale_size(self._base_width, self._base_height)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.White
        self.AutoScaleMode = AutoScaleMode.None
        self.MinimumSize = DPIScaler.scale_size(min(self._base_width, 500), min(self._base_height, 350))
    
    def _create_header(self):
        self.pnl_header = Panel()
        self.pnl_header.Location = Point(0, 0)
        self.pnl_header.Size = Size(self.ClientSize.Width, self._header_h)
        self.pnl_header.BackColor = Huisstijl.VIOLET
        self.pnl_header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_header)
        
        self.lbl_title = Label()
        self.lbl_title.Text = self._title
        self.lbl_title.Font = Font("Segoe UI", UIFactory.FONT_TITLE, FontStyle.Bold)
        self.lbl_title.ForeColor = Color.White
        self.lbl_title.AutoSize = True
        self.lbl_title.Location = Point(self._margin, DPIScaler.scale(15))
        self.pnl_header.Controls.Add(self.lbl_title)
        
        self.lbl_subtitle = Label()
        self.lbl_subtitle.Text = ""
        self.lbl_subtitle.Font = Font("Segoe UI", UIFactory.FONT_HEADING)
        self.lbl_subtitle.ForeColor = Huisstijl.TEAL
        self.lbl_subtitle.AutoSize = True
        self.lbl_subtitle.Location = Point(self._margin, DPIScaler.scale(50))
        self.pnl_header.Controls.Add(self.lbl_subtitle)
        
        self.pnl_accent = Panel()
        self.pnl_accent.Location = Point(0, self._header_h)
        self.pnl_accent.Size = Size(self.ClientSize.Width, self._accent_h)
        self.pnl_accent.BackColor = Huisstijl.TEAL
        self.pnl_accent.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_accent)
    
    def _create_footer(self):
        self.pnl_footer = Panel()
        self.pnl_footer.Size = Size(self.ClientSize.Width, self._footer_h)
        self.pnl_footer.Location = Point(0, self.ClientSize.Height - self._footer_h)
        self.pnl_footer.BackColor = Huisstijl.LIGHT_GRAY
        self.pnl_footer.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_footer)
        
        self.btn_close = UIFactory.create_button("Sluiten", 110, 40, 'primary')
        self.btn_close.Location = Point(
            self.pnl_footer.Width - self.btn_close.Width - self._margin,
            DPIScaler.scale(10)
        )
        self.btn_close.Anchor = AnchorStyles.Right | AnchorStyles.Top
        self.btn_close.Click += self._on_close_click
        self.pnl_footer.Controls.Add(self.btn_close)
    
    def _create_content_panel(self):
        content_top = self._header_h + self._accent_h
        content_height = self.ClientSize.Height - content_top - self._footer_h
        
        self.pnl_content = Panel()
        self.pnl_content.Location = Point(0, content_top)
        self.pnl_content.Size = Size(self.ClientSize.Width, content_height)
        self.pnl_content.BackColor = Color.White
        self.pnl_content.Padding = Padding(self._margin)
        self.pnl_content.AutoScroll = True
        self.pnl_content.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.pnl_content)
    
    def _on_form_resize(self, sender, args):
        if hasattr(self, 'pnl_content'):
            content_top = self._header_h + self._accent_h
            content_height = self.ClientSize.Height - content_top - self._footer_h
            self.pnl_content.Location = Point(0, content_top)
            self.pnl_content.Size = Size(self.ClientSize.Width, max(100, content_height))
    
    def set_subtitle(self, text):
        if hasattr(self, 'lbl_subtitle'):
            self.lbl_subtitle.Text = text
    
    def add_footer_button(self, text, style='secondary', click_handler=None, width=100):
        btn = UIFactory.create_button(text, width, 40, style)
        btn.Anchor = AnchorStyles.Right | AnchorStyles.Top
        
        existing_buttons = [c for c in self.pnl_footer.Controls if isinstance(c, Button)]
        total_width = sum(b.Width + DPIScaler.scale(10) for b in existing_buttons)
        
        btn.Location = Point(
            self.pnl_footer.Width - self._margin - total_width - btn.Width,
            DPIScaler.scale(10)
        )
        
        if click_handler:
            btn.Click += click_handler
        
        self.pnl_footer.Controls.Add(btn)
        return btn
    
    def _on_close_click(self, sender, args):
        self.Close()
    
    def show_info(self, message, title="Info"):
        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information)
    
    def show_warning(self, message, title="Waarschuwing"):
        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning)
    
    def show_error(self, message, title="Fout"):
        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error)
    
    def ask_confirm(self, message, title="Bevestigen"):
        result = MessageBox.Show(message, title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        return result == DialogResult.OK
    
    def save_file_dialog(self, filter="Alle bestanden (*.*)|*.*", filename=""):
        dialog = SaveFileDialog()
        dialog.Filter = filter
        dialog.FileName = filename
        if dialog.ShowDialog() == DialogResult.OK:
            return dialog.FileName
        return None


# ==============================================================================
# LAYOUT HELPERS
# ==============================================================================
class LayoutHelper:
    
    @staticmethod
    def stack_vertical(controls, start_y, spacing=10, x=None):
        current_y = DPIScaler.scale(start_y)
        scaled_spacing = DPIScaler.scale(spacing)
        
        for ctrl in controls:
            if x is not None:
                ctrl.Location = Point(DPIScaler.scale(x), current_y)
            else:
                ctrl.Location = Point(ctrl.Location.X, current_y)
            current_y += ctrl.Height + scaled_spacing
        return current_y
    
    @staticmethod
    def stack_horizontal(controls, start_x, y, spacing=10):
        current_x = DPIScaler.scale(start_x)
        scaled_y = DPIScaler.scale(y)
        scaled_spacing = DPIScaler.scale(spacing)
        
        for ctrl in controls:
            ctrl.Location = Point(current_x, scaled_y)
            current_x += ctrl.Width + scaled_spacing
        return current_x
    
    @staticmethod
    def create_form_row(label_text, input_control, label_width=120):
        panel = Panel()
        panel.Height = DPIScaler.scale(35)
        panel.AutoSize = False
        
        lbl = UIFactory.create_label(label_text)
        lbl.Location = Point(0, DPIScaler.scale(5))
        lbl.AutoSize = False
        lbl.Width = DPIScaler.scale(label_width)
        panel.Controls.Add(lbl)
        
        input_control.Location = Point(DPIScaler.scale(label_width + 10), 0)
        panel.Controls.Add(input_control)
        
        panel.Width = DPIScaler.scale(label_width + 10) + input_control.Width + DPIScaler.scale(10)
        return panel


# ==============================================================================
# STANDAARD DIALOGEN
# ==============================================================================
class MaterialSelectorDialog(BaseForm):
    
    def __init__(self, materials, current_name=""):
        super(MaterialSelectorDialog, self).__init__("Materiaal kiezen", 600, 300)
        
        self.materials = materials
        self.current_name = current_name
        self.selected_material = None
        self.selected_index = -1
        self._setup_ui()
    
    def _setup_ui(self):
        y = 10
        
        short_name = self.current_name[:40] + "..." if len(self.current_name) > 40 else self.current_name
        lbl_current = UIFactory.create_label(
            "Huidig: {}".format(short_name if short_name else "-"),
            color=Huisstijl.TEXT_SECONDARY
        )
        lbl_current.Location = DPIScaler.scale_point(10, y)
        self.pnl_content.Controls.Add(lbl_current)
        y += 35
        
        self.cmb_material = UIFactory.create_combobox(520)
        self.cmb_material.Location = DPIScaler.scale_point(10, y)
        
        for mat in self.materials:
            if isinstance(mat, tuple):
                self.cmb_material.Items.Add(" - ".join(str(x) for x in mat if x))
            else:
                self.cmb_material.Items.Add(str(mat))
        
        if self.cmb_material.Items.Count > 0:
            self.cmb_material.SelectedIndex = 0
        
        self.pnl_content.Controls.Add(self.cmb_material)
        self.btn_ok = self.add_footer_button("OK", 'primary', self._on_ok_click)
    
    def _on_ok_click(self, sender, args):
        self.selected_index = self.cmb_material.SelectedIndex
        if self.selected_index >= 0:
            self.selected_material = self.materials[self.selected_index]
        self.Close()
