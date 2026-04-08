# -*- coding: utf-8 -*-
"""
Schedule Export Tool
====================
Exporteer Revit schedules naar xlsx bestanden.
Met sets voor favoriete selecties en configuratie opties.
"""

__title__ = "Exp"
__author__ = "3BM Bouwkunde"
__doc__ = "Exporteer schedules naar Excel (.xlsx)"

import os
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Application, CheckedListBox, FolderBrowserDialog,
    AnchorStyles, BorderStyle, Padding, Panel, ScrollBars,
    DialogResult, SelectionMode
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# Revit imports
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ScheduleFieldType,
    SectionType, TableSectionData
)
from pyrevit import revit, forms, script

# UI imports - via lib folder
import sys
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), 'lib')
sys.path.insert(0, LIB_DIR)

from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl
from bm_logger import get_logger
from xlsx_helper import write_xlsx
import schedule_config as config

# GEEN doc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
log = get_logger("ScheduleExport")


# ==============================================================================
# SCHEDULE DATA EXTRACTION
# ==============================================================================
def get_all_schedules():
    """Haal alle schedules op uit het model"""
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    schedules = []
    
    for schedule in collector:
        # Skip templates en interne schedules
        if schedule.IsTemplate:
            continue
        if '<' in schedule.Name:  # Revision schedules etc.
            continue
        
        schedules.append({
            'name': schedule.Name,
            'id': schedule.Id,
            'element': schedule
        })
    
    return sorted(schedules, key=lambda x: x['name'].lower())


def extract_schedule_data(schedule):
    """
    Extraheer data uit een schedule.
    
    Returns:
        list: [[header1, header2, ...], [row1_val1, row1_val2, ...], ...]
    """
    table_data = schedule.GetTableData()
    section = table_data.GetSectionData(SectionType.Body)
    
    rows = []
    
    # Headers (eerste rij van body of aparte header sectie)
    header_section = table_data.GetSectionData(SectionType.Header)
    
    num_rows = section.NumberOfRows
    num_cols = section.NumberOfColumns
    
    # Extract alle rijen
    for row_idx in range(num_rows):
        row_data = []
        for col_idx in range(num_cols):
            try:
                cell_text = schedule.GetCellText(SectionType.Body, row_idx, col_idx)
                row_data.append(cell_text if cell_text else '')
            except:
                row_data.append('')
        rows.append(row_data)
    
    return rows


def get_current_schedule():
    """Haal huidige schedule op als actieve view een schedule is"""
    active_view = doc.ActiveView
    if isinstance(active_view, ViewSchedule) and not active_view.IsTemplate:
        return {
            'name': active_view.Name,
            'id': active_view.Id,
            'element': active_view
        }
    return None


# ==============================================================================
# EXPORT FORM
# ==============================================================================
class ScheduleExportForm(BaseForm):
    """Hoofd UI voor schedule export"""
    
    def __init__(self, schedules):
        super(ScheduleExportForm, self).__init__(
            "Schedule Export",
            width=500,
            height=680
        )
        
        self.all_schedules = schedules
        self.filtered_schedules = list(schedules)
        
        self.set_subtitle("{} schedules beschikbaar".format(len(schedules)))
        self._setup_ui()
        self._load_last_settings()
    
    def _setup_ui(self):
        """Bouw de UI op"""
        y = 10
        margin = 15
        full_width = 440
        combo_width = 300
        btn_small = 36
        
        # === SCHEDULES SECTIE ===
        lbl_schedules = UIFactory.create_label("Schedules", bold=True, font_size=UIFactory.FONT_HEADING)
        lbl_schedules.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(lbl_schedules)
        y += 28
        
        # Set selectie
        lbl_set = UIFactory.create_label("Set:")
        lbl_set.Location = DPIScaler.scale_point(margin, y + 5)
        self.pnl_content.Controls.Add(lbl_set)
        
        self.cmb_set = UIFactory.create_combobox(combo_width - 40, editable=False)
        self.cmb_set.Location = DPIScaler.scale_point(margin + 35, y)
        self.cmb_set.SelectedIndexChanged += self._set_changed
        self.pnl_content.Controls.Add(self.cmb_set)
        
        # Set buttons
        x_btn = margin + 35 + combo_width - 40 + 10
        self.btn_add_set = UIFactory.create_button("+", btn_small, 28, 'icon')
        self.btn_add_set.Location = DPIScaler.scale_point(x_btn, y)
        self.btn_add_set.Click += self._add_set_click
        self.pnl_content.Controls.Add(self.btn_add_set)
        
        self.btn_save_set = UIFactory.create_button("S", btn_small, 28, 'icon')
        self.btn_save_set.Location = DPIScaler.scale_point(x_btn + btn_small + 5, y)
        self.btn_save_set.Click += self._save_set_click
        self.pnl_content.Controls.Add(self.btn_save_set)
        
        self.btn_del_set = UIFactory.create_button("X", btn_small, 28, 'icon')
        self.btn_del_set.Location = DPIScaler.scale_point(x_btn + (btn_small + 5) * 2, y)
        self.btn_del_set.Click += self._delete_set_click
        self.pnl_content.Controls.Add(self.btn_del_set)
        y += 38
        
        # Search box
        self.txt_search = UIFactory.create_textbox(full_width)
        self.txt_search.Location = DPIScaler.scale_point(margin, y)
        self.txt_search.TextChanged += self._search_changed
        # Placeholder via GotFocus/LostFocus
        self.txt_search.Text = ""
        self.txt_search.ForeColor = Huisstijl.TEXT_SECONDARY
        self.pnl_content.Controls.Add(self.txt_search)
        
        # Clear button in search
        self.btn_clear_search = UIFactory.create_button("×", 28, 28, 'icon')
        self.btn_clear_search.Location = DPIScaler.scale_point(margin + full_width - 28, y)
        self.btn_clear_search.Click += self._clear_search_click
        self.pnl_content.Controls.Add(self.btn_clear_search)
        y += 38
        
        # Schedule checklist
        self.lst_schedules = CheckedListBox()
        self.lst_schedules.Location = DPIScaler.scale_point(margin, y)
        self.lst_schedules.Size = DPIScaler.scale_size(full_width, 180)
        self.lst_schedules.Font = Font("Segoe UI", UIFactory.FONT_NORMAL)
        self.lst_schedules.CheckOnClick = True
        self.lst_schedules.BorderStyle = BorderStyle.FixedSingle
        self.lst_schedules.ItemCheck += self._item_check_changed
        self.pnl_content.Controls.Add(self.lst_schedules)
        y += 188
        
        # Selected count en check buttons
        self.lbl_selected = UIFactory.create_label("Selected: 0", color=Huisstijl.TEXT_SECONDARY)
        self.lbl_selected.Location = DPIScaler.scale_point(margin, y + 5)
        self.pnl_content.Controls.Add(self.lbl_selected)
        
        self.btn_check_all = UIFactory.create_button("Check all visible", 120, 28, 'secondary')
        self.btn_check_all.Location = DPIScaler.scale_point(margin + 150, y)
        self.btn_check_all.Click += self._check_all_click
        self.pnl_content.Controls.Add(self.btn_check_all)
        
        self.btn_uncheck_all = UIFactory.create_button("Uncheck all visible", 130, 28, 'secondary')
        self.btn_uncheck_all.Location = DPIScaler.scale_point(margin + 280, y)
        self.btn_uncheck_all.Click += self._uncheck_all_click
        self.pnl_content.Controls.Add(self.btn_uncheck_all)
        y += 45
        
        # === EXPORT OPTIONS SECTIE ===
        lbl_options = UIFactory.create_label("Export options", bold=True, font_size=UIFactory.FONT_HEADING)
        lbl_options.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(lbl_options)
        y += 28
        
        # Configuration selectie
        lbl_config = UIFactory.create_label("Configuration:")
        lbl_config.Location = DPIScaler.scale_point(margin, y + 5)
        self.pnl_content.Controls.Add(lbl_config)
        
        self.cmb_config = UIFactory.create_combobox(combo_width - 90, editable=False)
        self.cmb_config.Location = DPIScaler.scale_point(margin + 90, y)
        self.cmb_config.SelectedIndexChanged += self._config_changed
        self.pnl_content.Controls.Add(self.cmb_config)
        
        # Config buttons
        x_btn = margin + 90 + combo_width - 90 + 10
        self.btn_add_config = UIFactory.create_button("+", btn_small, 28, 'icon')
        self.btn_add_config.Location = DPIScaler.scale_point(x_btn, y)
        self.btn_add_config.Click += self._add_config_click
        self.pnl_content.Controls.Add(self.btn_add_config)
        
        self.btn_save_config = UIFactory.create_button("S", btn_small, 28, 'icon')
        self.btn_save_config.Location = DPIScaler.scale_point(x_btn + btn_small + 5, y)
        self.btn_save_config.Click += self._save_config_click
        self.pnl_content.Controls.Add(self.btn_save_config)
        
        self.btn_del_config = UIFactory.create_button("X", btn_small, 28, 'icon')
        self.btn_del_config.Location = DPIScaler.scale_point(x_btn + (btn_small + 5) * 2, y)
        self.btn_del_config.Click += self._delete_config_click
        self.pnl_content.Controls.Add(self.btn_del_config)
        y += 38
        
        # Export folder
        lbl_folder = UIFactory.create_label("Export folder:")
        lbl_folder.Location = DPIScaler.scale_point(margin, y + 5)
        self.pnl_content.Controls.Add(lbl_folder)
        
        self.txt_folder = UIFactory.create_textbox(combo_width - 20)
        self.txt_folder.Location = DPIScaler.scale_point(margin + 90, y)
        self.txt_folder.Text = config.get_last_export_folder()
        self.pnl_content.Controls.Add(self.txt_folder)
        
        self.btn_browse = UIFactory.create_button("Select", 70, 28, 'secondary')
        self.btn_browse.Location = DPIScaler.scale_point(margin + 90 + combo_width - 20 + 10, y)
        self.btn_browse.Click += self._browse_click
        self.pnl_content.Controls.Add(self.btn_browse)
        y += 38
        
        # Separate files checkbox
        self.chk_separate = UIFactory.create_checkbox("Each schedule in a separate file")
        self.chk_separate.Location = DPIScaler.scale_point(margin, y)
        self.chk_separate.CheckedChanged += self._separate_changed
        self.pnl_content.Controls.Add(self.chk_separate)
        y += 30
        
        # File name prefix
        lbl_prefix = UIFactory.create_label("File name prefix:")
        lbl_prefix.Location = DPIScaler.scale_point(margin, y + 5)
        self.pnl_content.Controls.Add(lbl_prefix)
        
        self.txt_prefix = UIFactory.create_textbox(combo_width)
        self.txt_prefix.Location = DPIScaler.scale_point(margin + 105, y)
        self.txt_prefix.Enabled = False
        self.pnl_content.Controls.Add(self.txt_prefix)
        y += 38
        
        # File name
        lbl_filename = UIFactory.create_label("File name:")
        lbl_filename.Location = DPIScaler.scale_point(margin, y + 5)
        self.pnl_content.Controls.Add(lbl_filename)
        
        self.txt_filename = UIFactory.create_textbox(combo_width)
        self.txt_filename.Location = DPIScaler.scale_point(margin + 105, y)
        self.txt_filename.Text = doc.Title.replace('.rvt', '') if doc.Title else "export"
        self.pnl_content.Controls.Add(self.txt_filename)
        
        lbl_ext = UIFactory.create_label(".xlsx", color=Huisstijl.TEXT_SECONDARY)
        lbl_ext.Location = DPIScaler.scale_point(margin + 105 + combo_width + 5, y + 5)
        self.pnl_content.Controls.Add(lbl_ext)
        y += 38
        
        # Export title checkbox
        self.chk_title = UIFactory.create_checkbox("Export Title")
        self.chk_title.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(self.chk_title)
        
        # === FOOTER BUTTONS ===
        self.btn_export_current = self.add_footer_button("Export current", 'secondary', self._export_current_click, width=110)
        self.btn_export = self.add_footer_button("Export", 'primary', self._export_click, width=90)
        
        # Fill data
        self._fill_sets()
        self._fill_configs()
        self._fill_schedules()
    
    def _fill_sets(self):
        """Vul sets dropdown"""
        self.cmb_set.Items.Clear()
        self.cmb_set.Items.Add("(geen set)")
        
        for set_name in config.get_sets():
            self.cmb_set.Items.Add(set_name)
        
        # Selecteer laatst gebruikte
        last_set = config.get_last_set_name()
        if last_set:
            for i in range(self.cmb_set.Items.Count):
                if self.cmb_set.Items[i].ToString() == last_set:
                    self.cmb_set.SelectedIndex = i
                    return
        
        self.cmb_set.SelectedIndex = 0
    
    def _fill_configs(self):
        """Vul configurations dropdown"""
        self.cmb_config.Items.Clear()
        self.cmb_config.Items.Add("(default)")
        
        for cfg_name in config.get_configurations():
            self.cmb_config.Items.Add(cfg_name)
        
        # Selecteer laatst gebruikte
        last_config = config.get_last_configuration_name()
        if last_config:
            for i in range(self.cmb_config.Items.Count):
                if self.cmb_config.Items[i].ToString() == last_config:
                    self.cmb_config.SelectedIndex = i
                    return
        
        self.cmb_config.SelectedIndex = 0
    
    def _fill_schedules(self):
        """Vul schedule checklist"""
        self.lst_schedules.Items.Clear()
        
        for schedule in self.filtered_schedules:
            self.lst_schedules.Items.Add(schedule['name'])
        
        self._update_selected_count()
    
    def _load_last_settings(self):
        """Laad laatst gebruikte instellingen"""
        self.txt_folder.Text = config.get_last_export_folder()
        
        # Default prefix = modelnaam + underscore
        model_name = doc.Title
        # Verwijder .rvt extensie indien aanwezig
        if model_name.lower().endswith('.rvt'):
            model_name = model_name[:-4]
        self.txt_prefix.Text = model_name + "_"
    
    def _update_selected_count(self):
        """Update selected count label"""
        count = self.lst_schedules.CheckedItems.Count
        self.lbl_selected.Text = "Selected: {}".format(count)
    
    def _get_checked_schedules(self):
        """Haal lijst van geselecteerde schedules op"""
        checked = []
        for i in range(self.lst_schedules.Items.Count):
            if self.lst_schedules.GetItemChecked(i):
                name = self.lst_schedules.Items[i].ToString()
                # Vind schedule in all_schedules
                for sched in self.all_schedules:
                    if sched['name'] == name:
                        checked.append(sched)
                        break
        return checked
    
    # === EVENT HANDLERS ===
    def _set_changed(self, sender, args):
        """Set selectie gewijzigd"""
        if self.cmb_set.SelectedIndex <= 0:
            return
        
        set_name = self.cmb_set.SelectedItem.ToString()
        schedule_names = config.get_set(set_name)
        
        # Check matching schedules
        for i in range(self.lst_schedules.Items.Count):
            name = self.lst_schedules.Items[i].ToString()
            self.lst_schedules.SetItemChecked(i, name in schedule_names)
        
        self._update_selected_count()
    
    def _config_changed(self, sender, args):
        """Configuration selectie gewijzigd"""
        if self.cmb_config.SelectedIndex <= 0:
            return
        
        cfg_name = self.cmb_config.SelectedItem.ToString()
        cfg = config.get_configuration(cfg_name)
        
        # Apply configuration
        if cfg.get('export_folder'):
            self.txt_folder.Text = cfg['export_folder']
        self.chk_separate.Checked = cfg.get('separate_files', False)
        self.txt_prefix.Text = cfg.get('file_prefix', '')
        self.txt_filename.Text = cfg.get('filename', 'export')
        self.chk_title.Checked = cfg.get('include_title', False)
    
    def _search_changed(self, sender, args):
        """Search text gewijzigd"""
        search = self.txt_search.Text.lower().strip()
        
        if search:
            self.filtered_schedules = [
                s for s in self.all_schedules 
                if search in s['name'].lower()
            ]
        else:
            self.filtered_schedules = list(self.all_schedules)
        
        # Remember checked items
        checked_names = set()
        for i in range(self.lst_schedules.Items.Count):
            if self.lst_schedules.GetItemChecked(i):
                checked_names.add(self.lst_schedules.Items[i].ToString())
        
        # Refill
        self.lst_schedules.Items.Clear()
        for schedule in self.filtered_schedules:
            idx = self.lst_schedules.Items.Add(schedule['name'])
            if schedule['name'] in checked_names:
                self.lst_schedules.SetItemChecked(idx, True)
        
        self._update_selected_count()
    
    def _clear_search_click(self, sender, args):
        """Clear search"""
        self.txt_search.Text = ""
    
    def _item_check_changed(self, sender, args):
        """Item check gewijzigd"""
        # Delay update (ItemCheck fires before state changes)
        self._update_selected_count()
    
    def _check_all_click(self, sender, args):
        """Check all visible"""
        for i in range(self.lst_schedules.Items.Count):
            self.lst_schedules.SetItemChecked(i, True)
        self._update_selected_count()
    
    def _uncheck_all_click(self, sender, args):
        """Uncheck all visible"""
        for i in range(self.lst_schedules.Items.Count):
            self.lst_schedules.SetItemChecked(i, False)
        self._update_selected_count()
    
    def _separate_changed(self, sender, args):
        """Separate files checkbox gewijzigd"""
        self.txt_prefix.Enabled = self.chk_separate.Checked
        self.txt_filename.Enabled = not self.chk_separate.Checked
    
    def _browse_click(self, sender, args):
        """Browse voor export folder"""
        dialog = FolderBrowserDialog()
        dialog.Description = "Selecteer export folder"
        
        if self.txt_folder.Text and os.path.exists(self.txt_folder.Text):
            dialog.SelectedPath = self.txt_folder.Text
        
        if dialog.ShowDialog() == DialogResult.OK:
            self.txt_folder.Text = dialog.SelectedPath
            config.set_last_export_folder(dialog.SelectedPath)
    
    def _add_set_click(self, sender, args):
        """Nieuwe set toevoegen"""
        name = forms.ask_for_string(
            prompt="Naam voor nieuwe set:",
            title="Nieuwe Set"
        )
        if name:
            checked = [s['name'] for s in self._get_checked_schedules()]
            if config.save_set(name, checked):
                self._fill_sets()
                # Selecteer nieuwe set
                for i in range(self.cmb_set.Items.Count):
                    if self.cmb_set.Items[i].ToString() == name:
                        self.cmb_set.SelectedIndex = i
                        break
                self.show_info("Set '{}' opgeslagen met {} schedules.".format(name, len(checked)))
    
    def _save_set_click(self, sender, args):
        """Huidige set opslaan"""
        if self.cmb_set.SelectedIndex <= 0:
            self.show_warning("Selecteer eerst een set of maak een nieuwe aan.")
            return
        
        set_name = self.cmb_set.SelectedItem.ToString()
        checked = [s['name'] for s in self._get_checked_schedules()]
        
        if config.save_set(set_name, checked):
            self.show_info("Set '{}' bijgewerkt met {} schedules.".format(set_name, len(checked)))
    
    def _delete_set_click(self, sender, args):
        """Set verwijderen"""
        if self.cmb_set.SelectedIndex <= 0:
            return
        
        set_name = self.cmb_set.SelectedItem.ToString()
        if self.ask_confirm("Set '{}' verwijderen?".format(set_name)):
            config.delete_set(set_name)
            self._fill_sets()
    
    def _add_config_click(self, sender, args):
        """Nieuwe configuratie toevoegen"""
        name = forms.ask_for_string(
            prompt="Naam voor nieuwe configuratie:",
            title="Nieuwe Configuratie"
        )
        if name:
            options = {
                'export_folder': self.txt_folder.Text,
                'separate_files': self.chk_separate.Checked,
                'file_prefix': self.txt_prefix.Text,
                'filename': self.txt_filename.Text,
                'include_title': self.chk_title.Checked
            }
            if config.save_configuration(name, options):
                self._fill_configs()
                for i in range(self.cmb_config.Items.Count):
                    if self.cmb_config.Items[i].ToString() == name:
                        self.cmb_config.SelectedIndex = i
                        break
                self.show_info("Configuratie '{}' opgeslagen.".format(name))
    
    def _save_config_click(self, sender, args):
        """Huidige configuratie opslaan"""
        if self.cmb_config.SelectedIndex <= 0:
            self.show_warning("Selecteer eerst een configuratie of maak een nieuwe aan.")
            return
        
        cfg_name = self.cmb_config.SelectedItem.ToString()
        options = {
            'export_folder': self.txt_folder.Text,
            'separate_files': self.chk_separate.Checked,
            'file_prefix': self.txt_prefix.Text,
            'filename': self.txt_filename.Text,
            'include_title': self.chk_title.Checked
        }
        
        if config.save_configuration(cfg_name, options):
            self.show_info("Configuratie '{}' bijgewerkt.".format(cfg_name))
    
    def _delete_config_click(self, sender, args):
        """Configuratie verwijderen"""
        if self.cmb_config.SelectedIndex <= 0:
            return
        
        cfg_name = self.cmb_config.SelectedItem.ToString()
        if self.ask_confirm("Configuratie '{}' verwijderen?".format(cfg_name)):
            config.delete_configuration(cfg_name)
            self._fill_configs()
    
    def _export_current_click(self, sender, args):
        """Export huidige schedule (actieve view)"""
        current = get_current_schedule()
        if not current:
            self.show_warning("Actieve view is geen schedule.\nOpen een schedule en probeer opnieuw.")
            return
        
        self._do_export([current])
    
    def _export_click(self, sender, args):
        """Export geselecteerde schedules"""
        schedules = self._get_checked_schedules()
        if not schedules:
            self.show_warning("Selecteer minimaal één schedule.")
            return
        
        self._do_export(schedules)
    
    def _do_export(self, schedules):
        """Voer export uit"""
        export_folder = self.txt_folder.Text.strip()
        if not export_folder or not os.path.exists(export_folder):
            self.show_error("Selecteer een geldige export folder.")
            return
        
        config.set_last_export_folder(export_folder)
        
        try:
            if self.chk_separate.Checked:
                # Separate files
                prefix = self.txt_prefix.Text.strip()
                for schedule in schedules:
                    data = extract_schedule_data(schedule['element'])
                    
                    if self.chk_title.Checked:
                        data.insert(0, [schedule['name']])
                    
                    safe_name = "".join(c if c.isalnum() or c in ' -_' else '_' for c in schedule['name'])
                    filename = "{}{}".format(prefix, safe_name) if prefix else safe_name
                    filepath = os.path.join(export_folder, filename + '.xlsx')
                    
                    write_xlsx(filepath, data, schedule['name'])
                
                self.show_info("Geëxporteerd: {} bestanden naar\n{}".format(
                    len(schedules), export_folder))
            else:
                # Single file with multiple sheets
                filename = self.txt_filename.Text.strip() or 'export'
                filepath = os.path.join(export_folder, filename + '.xlsx')
                
                sheets_data = {}
                for schedule in schedules:
                    data = extract_schedule_data(schedule['element'])
                    
                    if self.chk_title.Checked:
                        data.insert(0, [schedule['name']])
                    
                    sheets_data[schedule['name']] = data
                
                write_xlsx(filepath, sheets_data)
                
                self.show_info("Geëxporteerd: {} schedules naar\n{}".format(
                    len(schedules), filepath))
            
            log.info("Export voltooid: {} schedules".format(len(schedules)))
            
        except Exception as e:
            log.error("Export fout: {}".format(e), exc_info=True)
            self.show_error("Export fout:\n{}".format(str(e)))


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    global doc
    
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Schedule Export")
        return
    
    log.info("Schedule Export gestart")
    
    # Haal schedules op
    schedules = get_all_schedules()
    
    if not schedules:
        forms.alert("Geen schedules gevonden in dit model.", exitscript=True)
    
    # Toon form
    form = ScheduleExportForm(schedules)
    form.ShowDialog()


if __name__ == '__main__':
    main()
