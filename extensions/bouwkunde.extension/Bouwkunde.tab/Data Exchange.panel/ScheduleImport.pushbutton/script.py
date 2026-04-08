# -*- coding: utf-8 -*-
"""
Schedule Import Tool
====================
Importeer data uit xlsx terug naar Revit schedules.
Matcht op ElementId en update alleen schrijfbare parameters.
"""

__title__ = "Imp"
__author__ = "3BM Bouwkunde"
__doc__ = "Importeer Excel data terug naar schedules"

import os
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Application, OpenFileDialog, AnchorStyles, 
    BorderStyle, DialogResult
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# Revit imports
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ScheduleFieldType,
    SectionType, TableSectionData, ScheduleField, Transaction,
    ElementId, StorageType
)
from pyrevit import revit, forms, script

# UI imports - via lib folder
import sys
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), 'lib')
sys.path.insert(0, LIB_DIR)

from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl
from bm_logger import get_logger
from xlsx_helper import read_xlsx, get_sheet_names
import schedule_config as config

# GEEN doc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
log = get_logger("ScheduleImport")


# ==============================================================================
# SCHEDULE ANALYSIS
# ==============================================================================
def get_all_schedules():
    """Haal alle schedules op uit het model"""
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    schedules = {}
    
    for schedule in collector:
        if schedule.IsTemplate:
            continue
        if '<' in schedule.Name:
            continue
        
        schedules[schedule.Name] = {
            'name': schedule.Name,
            'id': schedule.Id,
            'element': schedule
        }
    
    return schedules


def analyze_schedule(schedule):
    """
    Analyseer schedule structuur.
    
    Returns:
        dict met:
        - fields: lijst van {name, field_type, is_editable, column_index}
        - element_id_column: column index van ElementId (-1 als niet gevonden)
        - rows: aantal rijen
    """
    definition = schedule.Definition
    field_count = definition.GetFieldCount()
    
    fields = []
    element_id_column = -1
    
    for i in range(field_count):
        field = definition.GetField(i)
        field_info = {
            'name': field.GetName(),
            'field_type': field.FieldType,
            'is_editable': not field.IsCalculatedField,
            'column_index': i,
            'parameter_id': field.ParameterId if hasattr(field, 'ParameterId') else None,
            'field': field
        }
        
        # Check voor ElementId veld
        if field.GetName().lower() in ['elementid', 'element id', 'id']:
            element_id_column = i
        
        # Check of het een ElementId field type is
        if field.FieldType == ScheduleFieldType.ElementId:
            element_id_column = i
        
        fields.append(field_info)
    
    # Get row count
    table_data = schedule.GetTableData()
    section = table_data.GetSectionData(SectionType.Body)
    
    return {
        'fields': fields,
        'element_id_column': element_id_column,
        'rows': section.NumberOfRows,
        'columns': section.NumberOfColumns
    }


def get_schedule_elements(schedule, element_id_column):
    """
    Haal elementen op per rij in schedule.
    
    Returns:
        dict {row_index: element} of {element_id_str: element}
    """
    elements = {}
    
    table_data = schedule.GetTableData()
    section = table_data.GetSectionData(SectionType.Body)
    
    for row_idx in range(section.NumberOfRows):
        try:
            # Probeer ElementId uit cel te halen
            if element_id_column >= 0:
                cell_text = schedule.GetCellText(SectionType.Body, row_idx, element_id_column)
                if cell_text:
                    try:
                        elem_id = int(cell_text)
                        element = doc.GetElement(ElementId(elem_id))
                        if element:
                            elements[cell_text] = {
                                'element': element,
                                'row_index': row_idx
                            }
                    except:
                        pass
        except:
            pass
    
    return elements


def find_matching_column(xlsx_headers, schedule_fields):
    """
    Match xlsx kolommen met schedule velden.
    
    Returns:
        dict {xlsx_column_index: schedule_field_info}
    """
    matches = {}
    
    for xlsx_idx, header in enumerate(xlsx_headers):
        if not header:
            continue
        
        header_lower = str(header).lower().strip()
        
        for field in schedule_fields:
            field_name_lower = field['name'].lower().strip()
            
            if header_lower == field_name_lower:
                matches[xlsx_idx] = field
                break
    
    return matches


def update_element_parameter(element, param_name, value):
    """
    Update parameter waarde op element.
    
    Returns:
        tuple (success, message)
    """
    param = element.LookupParameter(param_name)
    
    if param is None:
        return False, "Parameter '{}' niet gevonden".format(param_name)
    
    if param.IsReadOnly:
        return False, "Parameter '{}' is read-only".format(param_name)
    
    try:
        storage_type = param.StorageType
        
        if storage_type == StorageType.String:
            param.Set(str(value) if value else '')
        elif storage_type == StorageType.Integer:
            param.Set(int(float(value)) if value else 0)
        elif storage_type == StorageType.Double:
            param.Set(float(value) if value else 0.0)
        elif storage_type == StorageType.ElementId:
            return False, "ElementId parameters worden niet aangepast"
        else:
            return False, "Onbekend storage type"
        
        return True, "OK"
        
    except Exception as e:
        return False, str(e)


# ==============================================================================
# IMPORT FORM
# ==============================================================================
class ScheduleImportForm(BaseForm):
    """Hoofd UI voor schedule import"""
    
    def __init__(self, schedules):
        super(ScheduleImportForm, self).__init__(
            "Schedule Import",
            width=550,
            height=550
        )
        
        self.schedules = schedules
        self.xlsx_data = None
        self.xlsx_path = None
        self.selected_sheet = None
        
        self.set_subtitle("Importeer xlsx data naar schedules")
        self._setup_ui()
    
    def _setup_ui(self):
        """Bouw de UI op"""
        y = 10
        margin = 15
        full_width = 490
        
        # === STAP 1: BESTAND SELECTEREN ===
        lbl_step1 = UIFactory.create_label("1. Selecteer xlsx bestand", bold=True, font_size=UIFactory.FONT_HEADING)
        lbl_step1.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(lbl_step1)
        y += 28
        
        self.txt_file = UIFactory.create_textbox(full_width - 90, readonly=True)
        self.txt_file.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(self.txt_file)
        
        self.btn_browse = UIFactory.create_button("Browse...", 80, 28, 'secondary')
        self.btn_browse.Location = DPIScaler.scale_point(margin + full_width - 80, y)
        self.btn_browse.Click += self._browse_click
        self.pnl_content.Controls.Add(self.btn_browse)
        y += 40
        
        # === STAP 2: SHEET SELECTEREN ===
        lbl_step2 = UIFactory.create_label("2. Selecteer sheet", bold=True, font_size=UIFactory.FONT_HEADING)
        lbl_step2.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(lbl_step2)
        y += 28
        
        self.cmb_sheet = UIFactory.create_combobox(full_width)
        self.cmb_sheet.Location = DPIScaler.scale_point(margin, y)
        self.cmb_sheet.Enabled = False
        self.cmb_sheet.SelectedIndexChanged += self._sheet_changed
        self.pnl_content.Controls.Add(self.cmb_sheet)
        y += 40
        
        # === STAP 3: TARGET SCHEDULE ===
        lbl_step3 = UIFactory.create_label("3. Selecteer target schedule", bold=True, font_size=UIFactory.FONT_HEADING)
        lbl_step3.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(lbl_step3)
        y += 28
        
        self.cmb_schedule = UIFactory.create_combobox(full_width)
        self.cmb_schedule.Location = DPIScaler.scale_point(margin, y)
        self.cmb_schedule.SelectedIndexChanged += self._schedule_changed
        
        for name in sorted(self.schedules.keys()):
            self.cmb_schedule.Items.Add(name)
        
        self.pnl_content.Controls.Add(self.cmb_schedule)
        y += 40
        
        # === PREVIEW / ANALYSE ===
        lbl_preview = UIFactory.create_label("Analyse", bold=True, font_size=UIFactory.FONT_HEADING)
        lbl_preview.Location = DPIScaler.scale_point(margin, y)
        self.pnl_content.Controls.Add(lbl_preview)
        y += 28
        
        self.txt_analysis = UIFactory.create_textbox(full_width, 180, multiline=True, readonly=True)
        self.txt_analysis.Location = DPIScaler.scale_point(margin, y)
        self.txt_analysis.Text = "Selecteer een xlsx bestand en schedule om de analyse te zien."
        self.pnl_content.Controls.Add(self.txt_analysis)
        y += 190
        
        # Waarschuwing
        lbl_warning = UIFactory.create_label(
            "⚠ Alleen schrijfbare parameters worden aangepast. Read-only en calculated velden worden overgeslagen.",
            font_size=UIFactory.FONT_SMALL,
            color=Huisstijl.TEXT_SECONDARY,
            italic=True
        )
        lbl_warning.Location = DPIScaler.scale_point(margin, y)
        lbl_warning.MaximumSize = DPIScaler.scale_size(full_width, 0)
        self.pnl_content.Controls.Add(lbl_warning)
        
        # === FOOTER ===
        self.btn_import = self.add_footer_button("Import", 'primary', self._import_click, width=100)
        self.btn_import.Enabled = False
    
    def _browse_click(self, sender, args):
        """Browse voor xlsx bestand"""
        dialog = OpenFileDialog()
        dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
        dialog.Title = "Selecteer xlsx bestand"
        
        last_folder = config.get_last_export_folder()
        if last_folder and os.path.exists(last_folder):
            dialog.InitialDirectory = last_folder
        
        if dialog.ShowDialog() == DialogResult.OK:
            self.xlsx_path = dialog.FileName
            self.txt_file.Text = dialog.FileName
            
            # Laad sheet namen
            try:
                sheets = get_sheet_names(dialog.FileName)
                self.cmb_sheet.Items.Clear()
                for sheet in sheets:
                    self.cmb_sheet.Items.Add(sheet)
                
                if self.cmb_sheet.Items.Count > 0:
                    self.cmb_sheet.Enabled = True
                    self.cmb_sheet.SelectedIndex = 0
                    
                    # Auto-match schedule naam
                    first_sheet = sheets[0]
                    for i in range(self.cmb_schedule.Items.Count):
                        if self.cmb_schedule.Items[i].ToString() == first_sheet:
                            self.cmb_schedule.SelectedIndex = i
                            break
                    
            except Exception as e:
                self.show_error("Fout bij lezen xlsx:\n{}".format(str(e)))
    
    def _sheet_changed(self, sender, args):
        """Sheet selectie gewijzigd"""
        self._update_analysis()
        
        # Auto-match schedule
        if self.cmb_sheet.SelectedIndex >= 0:
            sheet_name = self.cmb_sheet.SelectedItem.ToString()
            for i in range(self.cmb_schedule.Items.Count):
                if self.cmb_schedule.Items[i].ToString() == sheet_name:
                    self.cmb_schedule.SelectedIndex = i
                    break
    
    def _schedule_changed(self, sender, args):
        """Schedule selectie gewijzigd"""
        self._update_analysis()
    
    def _update_analysis(self):
        """Update analyse tekst"""
        if not self.xlsx_path or self.cmb_sheet.SelectedIndex < 0:
            self.txt_analysis.Text = "Selecteer een xlsx bestand en sheet."
            self.btn_import.Enabled = False
            return
        
        if self.cmb_schedule.SelectedIndex < 0:
            self.txt_analysis.Text = "Selecteer een target schedule."
            self.btn_import.Enabled = False
            return
        
        try:
            # Lees xlsx data
            sheet_name = self.cmb_sheet.SelectedItem.ToString()
            all_data = read_xlsx(self.xlsx_path)
            
            if sheet_name not in all_data:
                self.txt_analysis.Text = "Sheet '{}' niet gevonden in bestand.".format(sheet_name)
                self.btn_import.Enabled = False
                return
            
            xlsx_data = all_data[sheet_name]
            
            if not xlsx_data or len(xlsx_data) < 2:
                self.txt_analysis.Text = "Sheet bevat geen data (minimaal header + 1 rij nodig)."
                self.btn_import.Enabled = False
                return
            
            xlsx_headers = xlsx_data[0]
            xlsx_rows = len(xlsx_data) - 1
            
            # Analyseer schedule
            schedule_name = self.cmb_schedule.SelectedItem.ToString()
            schedule_info = self.schedules.get(schedule_name)
            
            if not schedule_info:
                self.txt_analysis.Text = "Schedule niet gevonden."
                self.btn_import.Enabled = False
                return
            
            schedule = schedule_info['element']
            analysis = analyze_schedule(schedule)
            
            # Match kolommen
            matches = find_matching_column(xlsx_headers, analysis['fields'])
            
            # Check ElementId kolom
            element_id_xlsx_col = -1
            for xlsx_idx, header in enumerate(xlsx_headers):
                if header and str(header).lower() in ['elementid', 'element id', 'id']:
                    element_id_xlsx_col = xlsx_idx
                    break
            
            # Build analysis text
            lines = []
            lines.append("=== XLSX BESTAND ===")
            lines.append("Sheet: {}".format(sheet_name))
            lines.append("Kolommen: {}".format(len(xlsx_headers)))
            lines.append("Data rijen: {}".format(xlsx_rows))
            lines.append("ElementId kolom: {}".format("Ja (kolom {})".format(element_id_xlsx_col + 1) if element_id_xlsx_col >= 0 else "NEE - matching niet mogelijk!"))
            lines.append("")
            lines.append("=== SCHEDULE ===")
            lines.append("Naam: {}".format(schedule_name))
            lines.append("Velden: {}".format(len(analysis['fields'])))
            lines.append("Rijen: {}".format(analysis['rows']))
            lines.append("")
            lines.append("=== KOLOM MATCHING ===")
            
            editable_matches = 0
            for xlsx_idx, field_info in matches.items():
                is_editable = field_info['is_editable'] and field_info['field_type'] != ScheduleFieldType.ElementId
                status = "✓ schrijfbaar" if is_editable else "✗ read-only/calculated"
                lines.append("  {} -> {} ({})".format(
                    xlsx_headers[xlsx_idx],
                    field_info['name'],
                    status
                ))
                if is_editable:
                    editable_matches += 1
            
            unmatched = [h for i, h in enumerate(xlsx_headers) if i not in matches and h]
            if unmatched:
                lines.append("")
                lines.append("Niet gematcht: {}".format(", ".join(str(h) for h in unmatched)))
            
            lines.append("")
            lines.append("=== RESULTAAT ===")
            
            if element_id_xlsx_col < 0:
                lines.append("❌ IMPORT NIET MOGELIJK: Geen ElementId kolom in xlsx")
                self.btn_import.Enabled = False
            elif editable_matches == 0:
                lines.append("❌ IMPORT NIET MOGELIJK: Geen schrijfbare kolommen gematcht")
                self.btn_import.Enabled = False
            else:
                lines.append("✓ {} schrijfbare kolommen kunnen worden geïmporteerd".format(editable_matches))
                self.btn_import.Enabled = True
            
            self.txt_analysis.Text = "\r\n".join(lines)
            
            # Store for import
            self.xlsx_data = xlsx_data
            self.selected_sheet = sheet_name
            
        except Exception as e:
            self.txt_analysis.Text = "Fout bij analyse:\n{}".format(str(e))
            self.btn_import.Enabled = False
            log.error("Analyse fout: {}".format(e), exc_info=True)
    
    def _import_click(self, sender, args):
        """Voer import uit"""
        if not self.xlsx_data or self.cmb_schedule.SelectedIndex < 0:
            return
        
        schedule_name = self.cmb_schedule.SelectedItem.ToString()
        schedule_info = self.schedules.get(schedule_name)
        
        if not schedule_info:
            self.show_error("Schedule niet gevonden.")
            return
        
        schedule = schedule_info['element']
        analysis = analyze_schedule(schedule)
        
        xlsx_headers = self.xlsx_data[0]
        xlsx_rows = self.xlsx_data[1:]
        
        # Vind ElementId kolom in xlsx
        element_id_xlsx_col = -1
        for idx, header in enumerate(xlsx_headers):
            if header and str(header).lower() in ['elementid', 'element id', 'id']:
                element_id_xlsx_col = idx
                break
        
        if element_id_xlsx_col < 0:
            self.show_error("Geen ElementId kolom gevonden in xlsx.")
            return
        
        # Match kolommen
        matches = find_matching_column(xlsx_headers, analysis['fields'])
        
        # Filter voor schrijfbare velden
        editable_matches = {
            k: v for k, v in matches.items() 
            if v['is_editable'] and v['field_type'] != ScheduleFieldType.ElementId
        }
        
        if not editable_matches:
            self.show_error("Geen schrijfbare kolommen om te importeren.")
            return
        
        # Bevestiging
        if not self.ask_confirm(
            "Import {} rijen naar schedule '{}'?\n\n"
            "Kolommen die worden aangepast:\n{}".format(
                len(xlsx_rows),
                schedule_name,
                "\n".join("  - {}".format(v['name']) for v in editable_matches.values())
            )
        ):
            return
        
        # Import uitvoeren
        updated = 0
        skipped = 0
        errors = []
        
        try:
            t = Transaction(doc, "Schedule Import")
            t.Start()
            
            for row_idx, row in enumerate(xlsx_rows):
                # Haal ElementId
                if element_id_xlsx_col >= len(row):
                    skipped += 1
                    continue
                
                elem_id_str = row[element_id_xlsx_col]
                if not elem_id_str:
                    skipped += 1
                    continue
                
                try:
                    elem_id = int(float(elem_id_str))
                    element = doc.GetElement(ElementId(elem_id))
                except:
                    skipped += 1
                    continue
                
                if not element:
                    skipped += 1
                    continue
                
                # Update parameters
                for xlsx_col, field_info in editable_matches.items():
                    if xlsx_col >= len(row):
                        continue
                    
                    value = row[xlsx_col]
                    param_name = field_info['name']
                    
                    success, msg = update_element_parameter(element, param_name, value)
                    
                    if success:
                        updated += 1
                    elif "read-only" not in msg.lower():
                        errors.append("Rij {}, {}: {}".format(row_idx + 2, param_name, msg))
            
            t.Commit()
            
            # Resultaat tonen
            result_msg = "Import voltooid!\n\n"
            result_msg += "• {} parameter waarden aangepast\n".format(updated)
            result_msg += "• {} rijen overgeslagen\n".format(skipped)
            
            if errors:
                result_msg += "\nFouten ({}):\n".format(len(errors))
                result_msg += "\n".join(errors[:10])
                if len(errors) > 10:
                    result_msg += "\n... en {} meer".format(len(errors) - 10)
            
            self.show_info(result_msg, "Import Resultaat")
            log.info("Import voltooid: {} updates, {} skipped".format(updated, skipped))
            
        except Exception as e:
            if t.HasStarted():
                t.RollBack()
            log.error("Import fout: {}".format(e), exc_info=True)
            self.show_error("Import fout:\n{}".format(str(e)))


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    global doc
    
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Schedule Import")
        return
    
    log.info("Schedule Import gestart")
    
    # Haal schedules op
    schedules = get_all_schedules()
    
    if not schedules:
        forms.alert("Geen schedules gevonden in dit model.", exitscript=True)
    
    # Toon form
    form = ScheduleImportForm(schedules)
    form.ShowDialog()


if __name__ == '__main__':
    main()
