# -*- coding: utf-8 -*-
"""
XLSX Helper Module
==================
Pure Python/IronPython implementatie voor xlsx lezen/schrijven.
Geen externe dependencies - werkt direct in pyRevit.

xlsx is een ZIP bestand met XML bestanden.
"""

import zipfile
import os
import re
from xml.etree import ElementTree as ET
from datetime import datetime

# XML namespaces voor xlsx
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'cp': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# Registreer namespaces
for prefix, uri in NS.items():
    ET.register_namespace('' if prefix == 'main' else prefix, uri)


class XlsxWriter:
    """Schrijf data naar xlsx bestand"""
    
    def __init__(self, filepath):
        self.filepath = filepath
        self.sheets = {}  # {sheet_name: [[row1], [row2], ...]}
    
    def add_sheet(self, name, data):
        """
        Voeg sheet toe met data.
        
        Args:
            name: Sheet naam
            data: 2D lijst [[header1, header2], [val1, val2], ...]
        """
        # Sanitize sheet name
        safe_name = re.sub(r'[\\/*?:\[\]]', '_', name)[:31]
        self.sheets[safe_name] = data
    
    def save(self):
        """Schrijf xlsx bestand"""
        with zipfile.ZipFile(self.filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml
            zf.writestr('[Content_Types].xml', self._create_content_types())
            
            # _rels/.rels
            zf.writestr('_rels/.rels', self._create_rels())
            
            # xl/workbook.xml
            zf.writestr('xl/workbook.xml', self._create_workbook())
            
            # xl/_rels/workbook.xml.rels
            zf.writestr('xl/_rels/workbook.xml.rels', self._create_workbook_rels())
            
            # xl/styles.xml
            zf.writestr('xl/styles.xml', self._create_styles())
            
            # Shared strings en sheets
            shared_strings = []
            string_index = {}
            
            for idx, (sheet_name, data) in enumerate(self.sheets.items(), 1):
                sheet_xml = self._create_sheet(data, shared_strings, string_index)
                zf.writestr('xl/worksheets/sheet{}.xml'.format(idx), sheet_xml)
            
            # xl/sharedStrings.xml
            if shared_strings:
                zf.writestr('xl/sharedStrings.xml', self._create_shared_strings(shared_strings))
    
    def _create_content_types(self):
        """Genereer [Content_Types].xml"""
        root = ET.Element('Types', xmlns=NS['cp'])
        
        ET.SubElement(root, 'Default', Extension='rels', 
                     ContentType='application/vnd.openxmlformats-package.relationships+xml')
        ET.SubElement(root, 'Default', Extension='xml', 
                     ContentType='application/xml')
        
        ET.SubElement(root, 'Override', PartName='/xl/workbook.xml',
                     ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
        ET.SubElement(root, 'Override', PartName='/xl/styles.xml',
                     ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')
        
        if any(self.sheets.values()):
            ET.SubElement(root, 'Override', PartName='/xl/sharedStrings.xml',
                         ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
        
        for idx in range(1, len(self.sheets) + 1):
            ET.SubElement(root, 'Override', 
                         PartName='/xl/worksheets/sheet{}.xml'.format(idx),
                         ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root)
    
    def _create_rels(self):
        """Genereer _rels/.rels"""
        root = ET.Element('Relationships', xmlns=NS['rel'])
        ET.SubElement(root, 'Relationship', Id='rId1',
                     Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                     Target='xl/workbook.xml')
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root)
    
    def _create_workbook(self):
        """Genereer xl/workbook.xml"""
        root = ET.Element('workbook', xmlns=NS['main'])
        root.set('{%s}r' % NS['r'], NS['r'])
        
        sheets = ET.SubElement(root, 'sheets')
        for idx, name in enumerate(self.sheets.keys(), 1):
            ET.SubElement(sheets, 'sheet', name=name, sheetId=str(idx),
                         **{'{%s}id' % NS['r']: 'rId{}'.format(idx)})
        
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root)
    
    def _create_workbook_rels(self):
        """Genereer xl/_rels/workbook.xml.rels"""
        root = ET.Element('Relationships', xmlns=NS['rel'])
        
        for idx in range(1, len(self.sheets) + 1):
            ET.SubElement(root, 'Relationship', Id='rId{}'.format(idx),
                         Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                         Target='worksheets/sheet{}.xml'.format(idx))
        
        rel_id = len(self.sheets) + 1
        ET.SubElement(root, 'Relationship', Id='rId{}'.format(rel_id),
                     Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                     Target='styles.xml')
        
        if any(self.sheets.values()):
            rel_id += 1
            ET.SubElement(root, 'Relationship', Id='rId{}'.format(rel_id),
                         Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                         Target='sharedStrings.xml')
        
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root)
    
    def _create_styles(self):
        """Genereer xl/styles.xml met header styling"""
        root = ET.Element('styleSheet', xmlns=NS['main'])
        
        # Fonts
        fonts = ET.SubElement(root, 'fonts', count='2')
        # Font 0: normaal
        font0 = ET.SubElement(fonts, 'font')
        ET.SubElement(font0, 'sz', val='11')
        ET.SubElement(font0, 'name', val='Calibri')
        # Font 1: bold voor headers
        font1 = ET.SubElement(fonts, 'font')
        ET.SubElement(font1, 'b')
        ET.SubElement(font1, 'sz', val='11')
        ET.SubElement(font1, 'name', val='Calibri')
        
        # Fills
        fills = ET.SubElement(root, 'fills', count='3')
        # Fill 0: none
        fill0 = ET.SubElement(fills, 'fill')
        ET.SubElement(fill0, 'patternFill', patternType='none')
        # Fill 1: gray125
        fill1 = ET.SubElement(fills, 'fill')
        ET.SubElement(fill1, 'patternFill', patternType='gray125')
        # Fill 2: header background (3BM Violet)
        fill2 = ET.SubElement(fills, 'fill')
        pf2 = ET.SubElement(fill2, 'patternFill', patternType='solid')
        ET.SubElement(pf2, 'fgColor', rgb='FF350E35')
        
        # Borders
        borders = ET.SubElement(root, 'borders', count='1')
        border = ET.SubElement(borders, 'border')
        ET.SubElement(border, 'left')
        ET.SubElement(border, 'right')
        ET.SubElement(border, 'top')
        ET.SubElement(border, 'bottom')
        ET.SubElement(border, 'diagonal')
        
        # Cell style xfs
        cellStyleXfs = ET.SubElement(root, 'cellStyleXfs', count='1')
        ET.SubElement(cellStyleXfs, 'xf', numFmtId='0', fontId='0', fillId='0', borderId='0')
        
        # Cell xfs
        cellXfs = ET.SubElement(root, 'cellXfs', count='2')
        # Style 0: normal
        ET.SubElement(cellXfs, 'xf', numFmtId='0', fontId='0', fillId='0', borderId='0', xfId='0')
        # Style 1: header (bold, violet background, white text)
        xf1 = ET.SubElement(cellXfs, 'xf', numFmtId='0', fontId='1', fillId='2', borderId='0', xfId='0')
        xf1.set('applyFont', '1')
        xf1.set('applyFill', '1')
        
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root)
    
    def _create_sheet(self, data, shared_strings, string_index):
        """Genereer worksheet XML"""
        root = ET.Element('worksheet', xmlns=NS['main'])
        
        # Bepaal dimensies
        if data:
            max_row = len(data)
            max_col = max(len(row) for row in data) if data else 0
            dim = 'A1:{}{}'.format(self._col_letter(max_col), max_row)
        else:
            dim = 'A1'
        
        ET.SubElement(root, 'dimension', ref=dim)
        
        sheet_data = ET.SubElement(root, 'sheetData')
        
        for row_idx, row in enumerate(data, 1):
            row_el = ET.SubElement(sheet_data, 'row', r=str(row_idx))
            
            for col_idx, value in enumerate(row, 1):
                cell_ref = '{}{}'.format(self._col_letter(col_idx), row_idx)
                cell = ET.SubElement(row_el, 'c', r=cell_ref)
                
                # Header row styling
                if row_idx == 1:
                    cell.set('s', '1')  # Style 1 = header
                
                if value is None:
                    continue
                
                # Type bepalen
                if isinstance(value, bool):
                    cell.set('t', 'b')
                    ET.SubElement(cell, 'v').text = '1' if value else '0'
                elif isinstance(value, (int, float)):
                    ET.SubElement(cell, 'v').text = str(value)
                else:
                    # String - gebruik shared strings
                    str_val = unicode(value) if hasattr(__builtins__, 'unicode') else str(value)
                    if str_val not in string_index:
                        string_index[str_val] = len(shared_strings)
                        shared_strings.append(str_val)
                    cell.set('t', 's')
                    ET.SubElement(cell, 'v').text = str(string_index[str_val])
        
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root)
    
    def _create_shared_strings(self, strings):
        """Genereer xl/sharedStrings.xml"""
        root = ET.Element('sst', xmlns=NS['main'])
        root.set('count', str(len(strings)))
        root.set('uniqueCount', str(len(strings)))
        
        for s in strings:
            si = ET.SubElement(root, 'si')
            t = ET.SubElement(si, 't')
            t.text = s
        
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root)
    
    def _col_letter(self, col_num):
        """Convert column number to Excel letter (1=A, 27=AA)"""
        result = ''
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result


class XlsxReader:
    """Lees data uit xlsx bestand"""
    
    def __init__(self, filepath):
        self.filepath = filepath
        self.shared_strings = []
        self.sheets = {}  # {sheet_name: [[row1], [row2], ...]}
    
    def read(self):
        """Lees xlsx bestand en return sheets dict"""
        with zipfile.ZipFile(self.filepath, 'r') as zf:
            # Lees shared strings
            try:
                ss_xml = zf.read('xl/sharedStrings.xml')
                self._parse_shared_strings(ss_xml)
            except KeyError:
                pass  # Geen shared strings
            
            # Lees workbook voor sheet namen
            wb_xml = zf.read('xl/workbook.xml')
            sheet_names = self._parse_workbook(wb_xml)
            
            # Lees workbook rels voor sheet files
            try:
                rels_xml = zf.read('xl/_rels/workbook.xml.rels')
                sheet_files = self._parse_workbook_rels(rels_xml)
            except KeyError:
                sheet_files = {}
            
            # Lees elke sheet
            for sheet_id, sheet_name in sheet_names.items():
                sheet_file = sheet_files.get(sheet_id, 'worksheets/sheet{}.xml'.format(sheet_id))
                try:
                    sheet_xml = zf.read('xl/' + sheet_file)
                    self.sheets[sheet_name] = self._parse_sheet(sheet_xml)
                except KeyError:
                    pass
        
        return self.sheets
    
    def _parse_shared_strings(self, xml_data):
        """Parse shared strings XML"""
        root = ET.fromstring(xml_data)
        ns = {'main': NS['main']}
        
        for si in root.findall('.//main:si', ns):
            t = si.find('.//main:t', ns)
            if t is not None and t.text:
                self.shared_strings.append(t.text)
            else:
                self.shared_strings.append('')
    
    def _parse_workbook(self, xml_data):
        """Parse workbook XML, return {rId: sheet_name}"""
        root = ET.fromstring(xml_data)
        ns = {'main': NS['main'], 'r': NS['r']}
        
        sheets = {}
        for sheet in root.findall('.//main:sheet', ns):
            name = sheet.get('name')
            r_id = sheet.get('{%s}id' % NS['r'])
            if r_id:
                # Extract nummer uit rId1, rId2, etc.
                match = re.search(r'rId(\d+)', r_id)
                if match:
                    sheets[r_id] = name
        
        return sheets
    
    def _parse_workbook_rels(self, xml_data):
        """Parse workbook rels, return {rId: target_file}"""
        root = ET.fromstring(xml_data)
        ns = {'rel': NS['rel']}
        
        rels = {}
        for rel in root.findall('.//rel:Relationship', ns):
            if rel is None:
                continue
            r_id = rel.get('Id')
            target = rel.get('Target')
            rel_type = rel.get('Type', '')
            if 'worksheet' in rel_type and r_id and target:
                rels[r_id] = target
        
        return rels
    
    def _parse_sheet(self, xml_data):
        """Parse worksheet XML, return 2D list"""
        root = ET.fromstring(xml_data)
        ns = {'main': NS['main']}
        
        rows = []
        max_col = 0
        
        for row in root.findall('.//main:row', ns):
            row_data = []
            for cell in row.findall('main:c', ns):
                cell_ref = cell.get('r', '')
                col_idx = self._col_index(cell_ref) - 1
                
                # Pad met None als er gaten zijn
                while len(row_data) < col_idx:
                    row_data.append(None)
                
                value = self._get_cell_value(cell, ns)
                row_data.append(value)
                max_col = max(max_col, len(row_data))
            
            rows.append(row_data)
        
        # Pad alle rijen tot zelfde lengte
        for row in rows:
            while len(row) < max_col:
                row.append(None)
        
        return rows
    
    def _get_cell_value(self, cell, ns):
        """Haal waarde uit cel"""
        cell_type = cell.get('t', '')
        v = cell.find('main:v', ns)
        
        if v is None or v.text is None:
            return None
        
        if cell_type == 's':
            # Shared string
            idx = int(v.text)
            return self.shared_strings[idx] if idx < len(self.shared_strings) else ''
        elif cell_type == 'b':
            # Boolean
            return v.text == '1'
        elif cell_type == 'str' or cell_type == 'inlineStr':
            # Inline string
            return v.text
        else:
            # Nummer
            try:
                if '.' in v.text:
                    return float(v.text)
                return int(v.text)
            except ValueError:
                return v.text
    
    def _col_index(self, cell_ref):
        """Convert cell reference to column index (A1 -> 1, B1 -> 2)"""
        col_str = ''.join(c for c in cell_ref if c.isalpha())
        result = 0
        for char in col_str.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result


# ==============================================================================
# CONVENIENCE FUNCTIONS
# ==============================================================================
def write_xlsx(filepath, data, sheet_name='Sheet1'):
    """
    Schrijf data naar xlsx bestand.
    
    Args:
        filepath: Pad naar xlsx bestand
        data: 2D lijst [[header1, header2], [val1, val2], ...]
              of dict {sheet_name: [[data]]}
        sheet_name: Sheet naam (alleen bij 2D lijst)
    """
    writer = XlsxWriter(filepath)
    
    if isinstance(data, dict):
        for name, sheet_data in data.items():
            writer.add_sheet(name, sheet_data)
    else:
        writer.add_sheet(sheet_name, data)
    
    writer.save()


def read_xlsx(filepath):
    """
    Lees xlsx bestand.
    
    Args:
        filepath: Pad naar xlsx bestand
    
    Returns:
        dict {sheet_name: [[row1], [row2], ...]}
    """
    reader = XlsxReader(filepath)
    return reader.read()


def get_sheet_names(filepath):
    """
    Haal sheet namen op uit xlsx bestand.
    
    Args:
        filepath: Pad naar xlsx bestand
    
    Returns:
        list met sheet namen
    """
    with zipfile.ZipFile(filepath, 'r') as zf:
        wb_xml = zf.read('xl/workbook.xml')
        root = ET.fromstring(wb_xml)
        ns = {'main': NS['main']}
        
        names = []
        for sheet in root.findall('.//main:sheet', ns):
            name = sheet.get('name')
            if name:
                names.append(name)
        
        return names
