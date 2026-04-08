# -*- coding: utf-8 -*-
"""
FilterCreator - Maak filters op basis van geselecteerd element
Genereert automatisch filternaam met SfB-code en omschrijving.
Werkt ook met elementen uit gelinkte modellen.
"""
__title__ = "Filter\nCreator"
__author__ = "3BM Bouwkunde"
__doc__ = "Maak een filter op basis van categorie en typename van geselecteerd element"

# .NET imports
import clr
clr.AddReference('System')
clr.AddReference('System.Drawing')
from System.Collections.Generic import List
from System.Drawing import Color

# Revit/pyRevit imports
from Autodesk.Revit.DB import (
    FilteredElementCollector, ParameterFilterElement, ElementFilter,
    ElementCategoryFilter, ElementParameterFilter, FilterRule,
    ParameterFilterRuleFactory, ElementId, BuiltInParameter,
    ParameterValueProvider, FilterStringRule, FilterStringEquals,
    SelectionFilterElement, Transaction, RevitLinkInstance,
    OverrideGraphicSettings, FillPatternElement, LinePatternElement,
    View, ViewType
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, forms, script

# UI imports
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl, LayoutHelper
from bm_logger import get_logger

# GEEN doc/uidoc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
uidoc = None
log = get_logger("FilterCreator")


# ==============================================================================
# SELECTION FILTER VOOR LINKED ELEMENTS
# ==============================================================================
class LinkableSelectionFilter(ISelectionFilter):
    """SelectionFilter die ook linked elements accepteert"""
    
    def AllowElement(self, elem):
        # Accepteer alles behalve RevitLinkInstance zelf
        return True
    
    def AllowReference(self, reference, position):
        # Accepteer alle references (inclusief linked)
        return True


# ==============================================================================
# NL-SfB CODES MAPPING
# ==============================================================================
SFB_CODES = {
    "Walls": ("21", "Wanden"), "Basic Wall": ("21", "Wanden"),
    "Curtain Wall": ("21", "Vliesgevels"), "Floors": ("23", "Vloeren"),
    "Roofs": ("27", "Daken"), "Ceilings": ("35", "Plafonds"),
    "Doors": ("31", "Deuren"), "Windows": ("31", "Ramen"),
    "Stairs": ("24", "Trappen"), "Railings": ("34", "Balustrades"),
    "Ramps": ("24", "Hellingen"), "Columns": ("22", "Kolommen"),
    "Structural Columns": ("22", "Kolommen"), "Structural Framing": ("22", "Liggers"),
    "Structural Foundations": ("16", "Funderingen"), "Furniture": ("72", "Meubilair"),
    "Plumbing Fixtures": ("52", "Sanitair"), "Mechanical Equipment": ("56", "Installaties"),
    "Electrical Equipment": ("61", "Elektra"), "Lighting Fixtures": ("63", "Verlichting"),
    "Generic Models": ("--", "Generiek"), "Pipes": ("52", "Leidingen"),
    "Ducts": ("56", "Kanalen"),
}

SFB_FALLBACK = {
    "OST_Walls": ("21", "Wanden"), "OST_Floors": ("23", "Vloeren"),
    "OST_Roofs": ("27", "Daken"), "OST_Ceilings": ("35", "Plafonds"),
    "OST_Doors": ("31", "Deuren"), "OST_Windows": ("31", "Ramen"),
    "OST_Stairs": ("24", "Trappen"), "OST_Columns": ("22", "Kolommen"),
    "OST_StructuralColumns": ("22", "Kolommen"), "OST_StructuralFraming": ("22", "Liggers"),
    "OST_StructuralFoundation": ("16", "Fundering"), "OST_Furniture": ("72", "Meubilair"),
    "OST_GenericModel": ("--", "Generiek"),
}


def get_sfb_suggestion(category_name, builtin_category=None):
    """Haal SfB code suggestie op basis van categorie"""
    if category_name in SFB_CODES:
        return SFB_CODES[category_name]
    if builtin_category:
        bic_name = str(builtin_category)
        if bic_name in SFB_FALLBACK:
            return SFB_FALLBACK[bic_name]
    return ("--", "Element")


def sanitize_name(name):
    """Maak naam geschikt voor filter"""
    if not name:
        return "onbekend"
    result = name
    for char in " -/\\:.,()[]":
        result = result.replace(char, "_" if char in " -/\\:." else "")
    while "__" in result:
        result = result.replace("__", "_")
    return result.strip("_")[:50]


def get_existing_filters():
    """Haal alle bestaande filternames op"""
    return [f.Name for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement)]


def filter_exists(name):
    return name in get_existing_filters()


def get_solid_fill_pattern_id():
    """Haal Solid Fill pattern ID op"""
    for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
        pattern = fp.GetFillPattern()
        if pattern and pattern.IsSolidFill:
            return fp.Id
    return ElementId.InvalidElementId


def get_view_template(view):
    """Haal view template op als view er een heeft"""
    template_id = view.ViewTemplateId
    if template_id and template_id != ElementId.InvalidElementId:
        return doc.GetElement(template_id)
    return None


def apply_filter_to_view(view, filter_element, color=None, visible=True):
    """Voeg filter toe aan view met optionele kleur override"""
    try:
        view.AddFilter(filter_element.Id)
        
        if color:
            override = OverrideGraphicSettings()
            solid_id = get_solid_fill_pattern_id()
            if solid_id != ElementId.InvalidElementId:
                override.SetSurfaceForegroundPatternId(solid_id)
                override.SetSurfaceBackgroundPatternId(solid_id)
                override.SetCutForegroundPatternId(solid_id)
                override.SetCutBackgroundPatternId(solid_id)
            
            override.SetSurfaceForegroundPatternColor(color)
            override.SetSurfaceBackgroundPatternColor(color)
            override.SetCutForegroundPatternColor(color)
            override.SetCutBackgroundPatternColor(color)
            override.SetProjectionLineColor(color)
            override.SetCutLineColor(color)
            
            view.SetFilterOverrides(filter_element.Id, override)
        
        view.SetFilterVisibility(filter_element.Id, visible)
        return True
    except Exception as e:
        log.warning("Kon filter niet toepassen: {}".format(e))
        return False


# ==============================================================================
# UI FORM - 520x820
# ==============================================================================
class FilterCreatorForm(BaseForm):
    """Hoofd UI voor FilterCreator"""
    
    def __init__(self, element_data, view_info):
        # Vergroot: 520 breed, 820 hoog
        super(FilterCreatorForm, self).__init__("Filter Creator", 520, 820)
        
        self.element_data = element_data
        self.view_info = view_info
        self.created_filter = None
        self.selected_color = Color.FromArgb(255, 100, 100)
        
        self.set_subtitle("Filter op basis van selectie")
        self._setup_ui()
    
    def _setup_ui(self):
        y = 12
        margin = 12
        label_width = 100
        content_width = 480  # Aangepast voor bredere form
        
        # === Element Info ===
        gb_info = UIFactory.create_groupbox("Element", content_width, 110)
        gb_info.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(gb_info)
        
        # Bron
        lbl_source = UIFactory.create_label("Bron:", bold=True, font_size=9)
        lbl_source.Location = DPIScaler.scale_point(12, 26)
        gb_info.Controls.Add(lbl_source)
        
        source_text = self.element_data.get('link_name') or 'Host model'
        source_color = Huisstijl.TEAL if self.element_data.get('is_linked') else Huisstijl.TEXT_PRIMARY
        lbl_source_val = UIFactory.create_label(str(source_text)[:45], color=source_color, font_size=9)
        lbl_source_val.Location = DPIScaler.scale_point(label_width, 26)
        gb_info.Controls.Add(lbl_source_val)
        
        # Categorie
        lbl_cat = UIFactory.create_label("Categorie:", bold=True, font_size=9)
        lbl_cat.Location = DPIScaler.scale_point(12, 52)
        gb_info.Controls.Add(lbl_cat)
        
        lbl_cat_val = UIFactory.create_label(self.element_data['category'], font_size=9)
        lbl_cat_val.Location = DPIScaler.scale_point(label_width, 52)
        gb_info.Controls.Add(lbl_cat_val)
        
        # Type
        lbl_type = UIFactory.create_label("Type:", bold=True, font_size=9)
        lbl_type.Location = DPIScaler.scale_point(12, 78)
        gb_info.Controls.Add(lbl_type)
        
        type_display = self.element_data['type_name'][:40] + "..." if len(self.element_data['type_name']) > 40 else self.element_data['type_name']
        lbl_type_val = UIFactory.create_label(type_display, font_size=9)
        lbl_type_val.Location = DPIScaler.scale_point(label_width, 78)
        gb_info.Controls.Add(lbl_type_val)
        
        y += 125
        
        # === Filter Config ===
        gb_config = UIFactory.create_groupbox("Filter", content_width, 160)
        gb_config.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(gb_config)
        
        # SfB Code
        lbl_sfb = UIFactory.create_label("SfB Code:", bold=True, font_size=9)
        lbl_sfb.Location = DPIScaler.scale_point(12, 30)
        gb_config.Controls.Add(lbl_sfb)
        
        sfb_items = [
            "21 - Wanden", "22 - Kolommen", "23 - Vloeren", "24 - Trappen",
            "27 - Daken", "31 - Deuren/ramen", "35 - Plafonds",
            "52 - Sanitair", "56 - Klimaat", "61 - Elektra", 
            "72 - Meubilair", "-- - Overig",
        ]
        self.cmb_sfb = UIFactory.create_combobox(160, sfb_items, editable=True)
        self.cmb_sfb.Location = DPIScaler.scale_point(label_width, 27)
        
        suggested_sfb = self.element_data['sfb_code']
        for i, item in enumerate(sfb_items):
            if item.startswith(suggested_sfb):
                self.cmb_sfb.SelectedIndex = i
                break
        self.cmb_sfb.TextChanged += self._on_input_changed
        gb_config.Controls.Add(self.cmb_sfb)
        
        # Omschrijving
        lbl_desc = UIFactory.create_label("Naam:", bold=True, font_size=9)
        lbl_desc.Location = DPIScaler.scale_point(12, 62)
        gb_config.Controls.Add(lbl_desc)
        
        self.txt_description = UIFactory.create_textbox(350)
        self.txt_description.Location = DPIScaler.scale_point(label_width, 59)
        self.txt_description.Text = self.element_data['suggested_name']
        self.txt_description.TextChanged += self._on_input_changed
        gb_config.Controls.Add(self.txt_description)
        
        # Filter type
        lbl_filter_type = UIFactory.create_label("Type:", bold=True, font_size=9)
        lbl_filter_type.Location = DPIScaler.scale_point(12, 94)
        gb_config.Controls.Add(lbl_filter_type)
        
        filter_types = ["Categorie + Type", "Alleen Categorie"]
        self.cmb_filter_type = UIFactory.create_combobox(220, filter_types)
        self.cmb_filter_type.Location = DPIScaler.scale_point(label_width, 91)
        gb_config.Controls.Add(self.cmb_filter_type)
        
        # Preview
        lbl_preview = UIFactory.create_label("Filternaam:", bold=True, font_size=9)
        lbl_preview.Location = DPIScaler.scale_point(12, 126)
        gb_config.Controls.Add(lbl_preview)
        
        self.lbl_preview_val = UIFactory.create_label("", color=Huisstijl.TEAL, bold=True, font_size=9)
        self.lbl_preview_val.Location = DPIScaler.scale_point(label_width, 126)
        self.lbl_preview_val.AutoSize = True
        gb_config.Controls.Add(self.lbl_preview_val)
        
        y += 175
        
        # === Toepassen ===
        gb_apply = UIFactory.create_groupbox("Toepassen", content_width, 190)
        gb_apply.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(gb_apply)
        
        # View/Template info
        view_text = "View: {}".format(self.view_info['view_name'][:35])
        if self.view_info['has_template']:
            view_text += "\nTemplate: {}".format(self.view_info['template_name'][:35])
        
        lbl_view_info = UIFactory.create_label(view_text, font_size=9, color=Huisstijl.TEXT_SECONDARY)
        lbl_view_info.Location = DPIScaler.scale_point(12, 26)
        lbl_view_info.AutoSize = True
        gb_apply.Controls.Add(lbl_view_info)
        
        # Toepassen opties
        apply_y = 60 if self.view_info['has_template'] else 50
        
        self.chk_apply_view = UIFactory.create_checkbox("Toevoegen aan actieve view", True)
        self.chk_apply_view.Location = DPIScaler.scale_point(12, apply_y)
        self.chk_apply_view.CheckedChanged += self._on_apply_changed
        gb_apply.Controls.Add(self.chk_apply_view)
        
        if self.view_info['has_template']:
            self.chk_apply_template = UIFactory.create_checkbox("Ook toevoegen aan template", True)
            self.chk_apply_template.Location = DPIScaler.scale_point(12, apply_y + 30)
            gb_apply.Controls.Add(self.chk_apply_template)
            apply_y += 30
        else:
            self.chk_apply_template = None
        
        # Kleur
        lbl_color = UIFactory.create_label("Kleur:", bold=True, font_size=9)
        lbl_color.Location = DPIScaler.scale_point(12, apply_y + 42)
        gb_apply.Controls.Add(lbl_color)
        
        self.color_buttons = []
        preset_colors = [
            Color.FromArgb(255, 100, 100),  # Rood
            Color.FromArgb(100, 200, 100),  # Groen
            Color.FromArgb(100, 150, 255),  # Blauw
            Color.FromArgb(255, 220, 100),  # Geel
            Color.FromArgb(255, 150, 80),   # Oranje
            Color.FromArgb(180, 120, 220),  # Paars
        ]
        
        from System.Windows.Forms import Button as WFButton, FlatStyle as WFFlatStyle
        for i, color in enumerate(preset_colors):
            btn = WFButton()
            btn.Text = ""
            btn.Size = DPIScaler.scale_size(30, 30)
            btn.Location = DPIScaler.scale_point(label_width + i * 36, apply_y + 38)
            btn.BackColor = color
            btn.FlatStyle = WFFlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.FlatAppearance.BorderColor = Huisstijl.MEDIUM_GRAY
            btn.Tag = color
            btn.Click += self._on_color_click
            gb_apply.Controls.Add(btn)
            self.color_buttons.append(btn)
        
        self.chk_no_color = UIFactory.create_checkbox("Geen kleur", False)
        self.chk_no_color.Location = DPIScaler.scale_point(label_width + 230, apply_y + 40)
        self.chk_no_color.CheckedChanged += self._on_no_color_changed
        gb_apply.Controls.Add(self.chk_no_color)
        
        y += 205
        
        # === Bestaande filters ===
        existing = get_existing_filters()
        sfb = self.element_data['sfb_code']
        similar = [f for f in existing if f.startswith(sfb)][:6]
        
        if similar:
            lbl_existing = UIFactory.create_label(
                "Vergelijkbaar: " + ", ".join(similar), 
                font_size=8, color=Huisstijl.TEXT_SECONDARY
            )
            lbl_existing.Location = DPIScaler.scale_point(margin, y)
            lbl_existing.MaximumSize = DPIScaler.scale_size(content_width, 50)
            self.pnl_content.Controls.Add(lbl_existing)
        
        # Footer
        self.add_footer_button("Annuleren", 'secondary', self._on_cancel, width=110)
        self.btn_create = self.add_footer_button("Aanmaken", 'primary', self._on_create, width=130)
        
        # Init
        self._update_preview()
        self._highlight_selected_color(self.color_buttons[0])
    
    def _on_input_changed(self, sender, args):
        self._update_preview()
    
    def _on_apply_changed(self, sender, args):
        enabled = self.chk_apply_view.Checked
        for btn in self.color_buttons:
            btn.Enabled = enabled and not self.chk_no_color.Checked
        self.chk_no_color.Enabled = enabled
        if self.chk_apply_template:
            self.chk_apply_template.Enabled = enabled
    
    def _on_no_color_changed(self, sender, args):
        for btn in self.color_buttons:
            btn.Enabled = not self.chk_no_color.Checked
    
    def _on_color_click(self, sender, args):
        self.selected_color = sender.Tag
        self._highlight_selected_color(sender)
    
    def _highlight_selected_color(self, selected_btn):
        from System.Windows.Forms import FlatStyle as WFFlatStyle
        for btn in self.color_buttons:
            if btn == selected_btn:
                btn.FlatStyle = WFFlatStyle.Flat
                btn.FlatAppearance.BorderSize = 2
                btn.FlatAppearance.BorderColor = Huisstijl.VIOLET
            else:
                btn.FlatStyle = WFFlatStyle.Flat
                btn.FlatAppearance.BorderSize = 1
                btn.FlatAppearance.BorderColor = Huisstijl.MEDIUM_GRAY
    
    def _update_preview(self):
        sfb_text = self.cmb_sfb.Text
        sfb_code = sfb_text.split(" - ")[0].strip() if " - " in sfb_text else sfb_text.strip()
        description = sanitize_name(self.txt_description.Text)
        
        if sfb_code and description:
            filter_name = "{}_{}".format(sfb_code, description)
        elif description:
            filter_name = description
        else:
            filter_name = "filter"
        
        if filter_exists(filter_name):
            self.lbl_preview_val.Text = filter_name + " (BESTAAT)"
            self.lbl_preview_val.ForeColor = Huisstijl.PEACH
            self.btn_create.Enabled = False
        else:
            self.lbl_preview_val.Text = filter_name
            self.lbl_preview_val.ForeColor = Huisstijl.TEAL
            self.btn_create.Enabled = True
    
    def _get_filter_name(self):
        sfb_text = self.cmb_sfb.Text
        sfb_code = sfb_text.split(" - ")[0].strip() if " - " in sfb_text else sfb_text.strip()
        description = sanitize_name(self.txt_description.Text)
        
        if sfb_code and description:
            return "{}_{}".format(sfb_code, description)
        return description or "filter_{}".format(self.element_data['category'])
    
    def _on_create(self, sender, args):
        filter_name = self._get_filter_name()
        filter_type_idx = self.cmb_filter_type.SelectedIndex
        apply_to_view = self.chk_apply_view.Checked
        apply_to_template = self.chk_apply_template.Checked if self.chk_apply_template else False
        use_color = not self.chk_no_color.Checked
        
        log.info("Creating filter: {} (apply_view={}, apply_template={})".format(
            filter_name, apply_to_view, apply_to_template))
        
        try:
            with Transaction(doc, "Filter: {}".format(filter_name)) as t:
                t.Start()
                
                cat_id = self.element_data['category_id']
                categories = List[ElementId]([cat_id])
                
                if filter_type_idx == 0:
                    type_name = self.element_data['type_name']
                    param_id = ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, type_name, True)
                    elem_filter = ElementParameterFilter(rule)
                    new_filter = ParameterFilterElement.Create(doc, filter_name, categories, elem_filter)
                else:
                    new_filter = ParameterFilterElement.Create(doc, filter_name, categories)
                
                # Kleur voorbereiden
                color = None
                if use_color:
                    from Autodesk.Revit.DB import Color as RevitColor
                    color = RevitColor(self.selected_color.R, self.selected_color.G, self.selected_color.B)
                
                # Toepassen
                applied_to = []
                active_view = doc.ActiveView
                
                if apply_to_view and active_view:
                    if active_view.ViewType in [ViewType.FloorPlan, ViewType.CeilingPlan, 
                                                 ViewType.Elevation, ViewType.Section,
                                                 ViewType.ThreeD, ViewType.Detail]:
                        if apply_filter_to_view(active_view, new_filter, color):
                            applied_to.append("view")
                
                if apply_to_template and self.view_info['template']:
                    template = self.view_info['template']
                    if apply_filter_to_view(template, new_filter, color):
                        applied_to.append("template")
                
                t.Commit()
                
                self.created_filter = new_filter
                log.info("Filter aangemaakt: {} -> {}".format(filter_name, applied_to))
                
                # Success
                msg = "Filter '{}' aangemaakt!".format(filter_name)
                if applied_to:
                    msg += "\n\nToegevoegd aan: {}".format(", ".join(applied_to))
                
                self.show_info(msg)
                self.Close()
                
        except Exception as e:
            log.error("Fout: {}".format(e), exc_info=True)
            self.show_error("Fout:\n{}".format(str(e)))
    
    def _on_cancel(self, sender, args):
        self.Close()


# ==============================================================================
# HELPERS
# ==============================================================================
def get_element_from_reference(reference):
    """
    Haal element op uit reference, ook voor linked elements.
    Werkt met PointOnElement - check LinkedElementId property.
    """
    try:
        linked_elem_id = reference.LinkedElementId
    except:
        linked_elem_id = ElementId.InvalidElementId
    
    host_elem = doc.GetElement(reference.ElementId)
    
    log.info("get_element_from_reference: ElementId={}, LinkedElementId={}, host_type={}".format(
        reference.ElementId, linked_elem_id, type(host_elem).__name__))
    
    # Check of LinkedElementId geldig is (dan is het linked)
    if linked_elem_id and linked_elem_id != ElementId.InvalidElementId:
        # host_elem zou de RevitLinkInstance moeten zijn
        if isinstance(host_elem, RevitLinkInstance):
            link_doc = host_elem.GetLinkDocument()
            if link_doc:
                linked_element = link_doc.GetElement(linked_elem_id)
                log.info("SUCCESS: Linked element {} in '{}'".format(
                    linked_element.Id if linked_element else "None", link_doc.Title))
                return linked_element, link_doc, link_doc.Title, True
            else:
                log.warning("Link document niet beschikbaar")
        else:
            log.warning("Host element is geen RevitLinkInstance: {}".format(type(host_elem).__name__))
    
    # Geen linked element, return host element
    log.info("Returning host element: {}".format(host_elem.Id if host_elem else None))
    return host_elem, doc, None, False


def get_element_info(element, source_doc=None, link_name=None, is_linked=False):
    """Verzamel info over geselecteerd element"""
    if source_doc is None:
        source_doc = doc
    
    category = element.Category
    cat_name = category.Name if category else "Onbekend"
    
    # Voor linked: zoek category in host doc via BuiltInCategory
    if is_linked and category:
        host_cat = None
        try:
            bic = category.BuiltInCategory
            # Zoek via BuiltInCategory (betrouwbaarder dan naam)
            host_cat = doc.Settings.Categories.get_Item(bic)
        except:
            pass
        
        if not host_cat:
            # Fallback naar naam-matching
            for cat in doc.Settings.Categories:
                if cat.Name == cat_name:
                    host_cat = cat
                    break
        
        cat_id = host_cat.Id if host_cat else ElementId.InvalidElementId
        log.info("Linked category '{}' -> host cat_id: {}".format(cat_name, cat_id))
    else:
        cat_id = category.Id if category else ElementId.InvalidElementId
    
    try:
        bic = category.BuiltInCategory if category else None
    except:
        bic = None
    
    type_id = element.GetTypeId()
    elem_type = source_doc.GetElement(type_id) if type_id != ElementId.InvalidElementId else None
    
    if elem_type:
        type_name_param = elem_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        type_name = type_name_param.AsString() if type_name_param else elem_type.Name
        try:
            fam_name = elem_type.FamilyName
        except:
            fam_name = ""
    else:
        type_name = "Onbekend"
        fam_name = ""
    
    sfb_code, _ = get_sfb_suggestion(cat_name, bic)
    suggested = sanitize_name(fam_name if fam_name and fam_name != cat_name else type_name)
    
    return {
        'element_id': element.Id,
        'category': cat_name,
        'category_id': cat_id,
        'type_name': type_name,
        'family_name': fam_name,
        'sfb_code': sfb_code,
        'suggested_name': suggested,
        'is_linked': is_linked,
        'link_name': link_name,
    }


def get_view_info():
    """Verzamel info over actieve view en template"""
    active_view = doc.ActiveView
    template = get_view_template(active_view) if active_view else None
    
    return {
        'view': active_view,
        'view_name': active_view.Name if active_view else "Geen",
        'has_template': template is not None,
        'template': template,
        'template_name': template.Name if template else None,
    }


# ==============================================================================
# LINKED MODEL HELPERS
# ==============================================================================
def get_linked_models():
    """Haal alle geladen linked models op"""
    links = []
    for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = link.GetLinkDocument()
        if link_doc:
            links.append({
                'instance': link,
                'doc': link_doc,
                'name': link_doc.Title
            })
    return links


def select_from_link(link_instance):
    """Selecteer een element uit een specifieke link"""
    try:
        sel_filter = LinkableSelectionFilter()
        ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            sel_filter,
            "Selecteer element in '{}'".format(link_instance.GetLinkDocument().Title)
        )
        
        # Bij LinkedElement is ElementId de link instance, LinkedElementId het element
        linked_elem_id = ref.LinkedElementId
        link_doc = link_instance.GetLinkDocument()
        
        if linked_elem_id and linked_elem_id != ElementId.InvalidElementId and link_doc:
            element = link_doc.GetElement(linked_elem_id)
            log.info("Linked element geselecteerd: {} in {}".format(
                element.Id, link_doc.Title))
            return element, link_doc, link_doc.Title, True
        
        return None, None, None, False
        
    except Exception as e:
        log.info("Linked selectie geannuleerd: {}".format(e))
        return None, None, None, False


def select_from_host():
    """Selecteer een element uit het host model"""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Selecteer element in host model"
        )
        element = doc.GetElement(ref.ElementId)
        
        # Check of het geen link instance is
        if isinstance(element, RevitLinkInstance):
            return None, None, None, False
            
        log.info("Host element geselecteerd: {}".format(element.Id))
        return element, doc, None, False
        
    except Exception as e:
        log.info("Host selectie geannuleerd: {}".format(e))
        return None, None, None, False


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    global doc, uidoc
    
    doc = revit.doc
    uidoc = revit.uidoc
    
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Filter Creator")
        return
    
    log.info("=" * 50)
    log.info("FilterCreator gestart")
    
    selection = uidoc.Selection.GetElementIds()
    
    element = None
    source_doc = doc
    link_name = None
    is_linked = False
    
    if not selection or selection.Count == 0:
        # Geen selectie - check of er links zijn
        links = get_linked_models()
        
        if links:
            # Er zijn linked models - vraag gebruiker waar te selecteren
            options = ["Host model"] + [link['name'] for link in links]
            choice = forms.SelectFromList.show(
                options,
                title="Selecteer bron",
                message="Waar wil je een element selecteren?",
                button_name="Selecteer"
            )
            
            if not choice:
                return
            
            if choice == "Host model":
                element, source_doc, link_name, is_linked = select_from_host()
            else:
                # Zoek de gekozen link
                for link in links:
                    if link['name'] == choice:
                        element, source_doc, link_name, is_linked = select_from_link(link['instance'])
                        break
        else:
            # Geen links - selecteer direct uit host
            element, source_doc, link_name, is_linked = select_from_host()
    else:
        # Gebruik huidige selectie
        elem_id = list(selection)[0]
        element = doc.GetElement(elem_id)
        log.info("Bestaande selectie: {}".format(elem_id))
        
        # Check of geselecteerde element een RevitLinkInstance is
        if isinstance(element, RevitLinkInstance):
            forms.alert("Selecteer een element IN het gelinkte model,\nniet het gelinkte model zelf.\n\nDeselecteer alles en run de tool opnieuw.", 
                       title="Filter Creator")
            return
    
    if not element:
        log.info("Geen element - gestopt")
        return
    
    element_data = get_element_info(element, source_doc, link_name, is_linked)
    view_info = get_view_info()
    
    log.info("Element: {} ({}) | Type: {} | cat_id: {}".format(
        element_data['category'], 
        "LINKED: " + str(link_name) if is_linked else "host",
        element_data['type_name'],
        element_data['category_id']))
    log.info("View: {} | Template: {}".format(
        view_info['view_name'], 
        view_info['template_name']))
    
    if element_data['category_id'] == ElementId.InvalidElementId:
        forms.alert("Categorie '{}' niet gevonden in host model.\n\nDit kan gebeuren als het gelinkte model een categorie gebruikt die niet in het host model bestaat.".format(
            element_data['category']), exitscript=True)
        return
    
    form = FilterCreatorForm(element_data, view_info)
    form.ShowDialog()


if __name__ == "__main__":
    main()
