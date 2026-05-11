# -*- coding: utf-8 -*-
"""Sheet Parameters Updater
Update tekening parameters voor meerdere sheets tegelijk.
"""
__title__ = "Sheet Prm"
__author__ = "3BM Bouwkunde"

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from wpf_template import WPFWindow, Huisstijl
from bm_logger import get_logger

import datetime

# GEEN doc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen
doc = None
log = get_logger("SheetParameters")


# ==============================================================================
# CONSTANTEN
# ==============================================================================
STATUS_OPTIONS = ["Definitief", "Voorlopig", "Concept", "Voor Akkoord"]
FASE_OPTIONS = ["OV", "DO", "SO", "VO", "TO", "UO"]
JA_NEE = ["Ja", "Nee"]
LEDEN = [
    "00_3BM_auteur", "01_auteur_MDVroegindeweij", "02_auteur_PMol",
    "03_auteur_JHCBongers", "05_auteur_JPDaane", "06_auteur_ATuk",
    "08_auteur_LNazaria", "10_auteur_JKolthof", "11_auteur_MPGStok",
    "12_auteur_AAli", "13_auteur_JdeKrijger", "14_auteur_LPost",
    "15_auteur_TvanZyl", "16_auteur_MHosseini"
]
WIJZIGING_LETTERS = ['A', 'B', 'C', 'D', 'E', 'F']


# ==============================================================================
# UI WINDOW
# ==============================================================================
class SheetParameterWindow(WPFWindow):
    """Sheet Parameters UI - WPF versie"""

    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        super(SheetParameterWindow, self).__init__(
            xaml_file, "Sheet Parameters Updater", width=1050, height=850
        )
        self.wijz_controls = []
        self._populate_combos()
        self._build_wijz_list()
        self._load_defaults()
        self._bind_events()

    def _populate_combos(self):
        """Vul alle comboboxen"""
        for item in STATUS_OPTIONS:
            self.combo_status.Items.Add(item)
        for item in FASE_OPTIONS:
            self.combo_fase.Items.Add(item)
        for item in JA_NEE:
            self.combo_std_schaal.Items.Add(item)
            self.combo_peil.Items.Add(item)
            self.combo_noord.Items.Add(item)
            self.combo_stempel.Items.Add(item)
        for item in LEDEN:
            self.combo_00_auteur.Items.Add(item)

    def _build_wijz_list(self):
        """Bouw lijst van wijziging control-tuples"""
        self.wijz_controls = [
            (self.chk_wijz_a, self.txt_wijz_a_datum, self.txt_wijz_a_omschr),
            (self.chk_wijz_b, self.txt_wijz_b_datum, self.txt_wijz_b_omschr),
            (self.chk_wijz_c, self.txt_wijz_c_datum, self.txt_wijz_c_omschr),
            (self.chk_wijz_d, self.txt_wijz_d_datum, self.txt_wijz_d_omschr),
            (self.chk_wijz_e, self.txt_wijz_e_datum, self.txt_wijz_e_omschr),
            (self.chk_wijz_f, self.txt_wijz_f_datum, self.txt_wijz_f_omschr),
        ]

    def _load_defaults(self):
        """Laad standaardwaarden"""
        today = datetime.datetime.now()
        self.txt_date.Text = today.strftime("%d-%m-%Y")
        self.txt_schaal.Text = "1:50"

        # Titleblock defaults — Ja=1 (standaard/zichtbaar)
        self.combo_std_schaal.SelectedIndex = 0  # Ja
        self.combo_peil.SelectedIndex = 0  # Ja
        self.combo_noord.SelectedIndex = 0  # Ja
        self.combo_stempel.SelectedIndex = 0  # Ja
        self.txt_kenmerk.Text = "2"
        self.txt_aantal_wijz.Text = "0"

        # Wijzigingen defaults
        for chk, datum, omschr in self.wijz_controls:
            omschr.Text = "-"

    def _bind_events(self):
        """Bind button events"""
        if self.btn_ok:
            self.btn_ok.Click += self._on_ok
        if self.btn_cancel:
            self.btn_cancel.Click += self._on_cancel

    def _on_ok(self, sender, args):
        self.close_ok()

    def _on_cancel(self, sender, args):
        self.close_cancel()

    def get_parameters(self):
        """Verzamel alle ingevulde parameters"""
        params = {}

        # Sheet parameters
        if self.chk_status.IsChecked == True and self.combo_status.SelectedIndex >= 0:
            params['status'] = self.combo_status.SelectedItem

        if self.chk_date.IsChecked == True and self.txt_date.Text.strip():
            params['issue_date'] = self.txt_date.Text.strip()

        if self.chk_schaal.IsChecked == True and self.txt_schaal.Text.strip():
            params['schaal'] = self.txt_schaal.Text.strip()

        if self.chk_fase.IsChecked == True and self.combo_fase.SelectedIndex >= 0:
            params['fase'] = self.combo_fase.SelectedItem

        # Wijzigingen
        for i, (chk, datum_ctrl, omschr_ctrl) in enumerate(self.wijz_controls):
            if chk.IsChecked == True:
                letter = chr(97 + i)  # a, b, c, d, e, f
                datum = datum_ctrl.Text.strip()
                omschr = omschr_ctrl.Text.strip()

                if datum:
                    params["wijziging_{}_datum".format(letter)] = datum
                if omschr and omschr != "-":
                    params["wijziging_{}_omschr".format(letter)] = omschr

        # Titleblock parameters — Ja=1, Nee=0 (Revit Yes/No convention)
        if self.chk_std_schaal.IsChecked == True and self.combo_std_schaal.SelectedIndex >= 0:
            params['std_schaal'] = 1 if self.combo_std_schaal.SelectedIndex == 0 else 0

        if self.chk_peil.IsChecked == True and self.combo_peil.SelectedIndex >= 0:
            params['v_peil'] = 1 if self.combo_peil.SelectedIndex == 0 else 0

        if self.chk_noord.IsChecked == True and self.combo_noord.SelectedIndex >= 0:
            params['noordpijl'] = 1 if self.combo_noord.SelectedIndex == 0 else 0

        if self.chk_kenmerk.IsChecked == True and self.txt_kenmerk.Text.strip():
            try:
                params['kenmerknummer'] = int(self.txt_kenmerk.Text.strip())
            except (ValueError, TypeError):
                pass

        if self.chk_stempel.IsChecked == True and self.combo_stempel.SelectedIndex >= 0:
            params['stempel'] = 1 if self.combo_stempel.SelectedIndex == 0 else 0

        if self.chk_aantal_wijz.IsChecked == True and self.txt_aantal_wijz.Text.strip():
            try:
                params['aantal_wijzigingen'] = int(self.txt_aantal_wijz.Text.strip())
            except (ValueError, TypeError):
                pass

        if self.chk_00_auteur.IsChecked == True and self.combo_00_auteur.SelectedIndex >= 0:
            params['00_3bm_auteur'] = self.combo_00_auteur.SelectedItem

        return params

    def get_filter_text(self):
        return self.txt_filter.Text.strip()


# ==============================================================================
# REVIT FUNCTIES
# ==============================================================================
def get_parameter_value(element, param_name):
    """Haal parameter waarde op"""
    param = element.LookupParameter(param_name)
    if param and param.HasValue:
        if param.StorageType == StorageType.String:
            return param.AsString()
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == StorageType.Double:
            return param.AsDouble()
    return None


def _resolve_param(element, param_name):
    """Zoek parameter op instance, val terug op symbol (type-parameter)."""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        return param
    symbol = getattr(element, 'Symbol', None)
    if symbol is not None:
        sym_param = symbol.LookupParameter(param_name)
        if sym_param and not sym_param.IsReadOnly:
            return sym_param
    return None


def set_parameter_value(element, param_name, value):
    """Set parameter waarde - probeert instance eerst, dan type."""
    param = _resolve_param(element, param_name)
    if not param:
        return False
    try:
        if param.StorageType == StorageType.String:
            param.Set(str(value))
        elif param.StorageType == StorageType.Integer:
            param.Set(int(value))
        elif param.StorageType == StorageType.Double:
            param.Set(float(value))
        else:
            return False
        return True
    except (ValueError, TypeError, Exception):
        return False


def filter_sheets_by_number(sheets, filter_text):
    """Filter sheets op nummer"""
    if not filter_text:
        return sheets

    filter_lower = filter_text.lower()
    return [s for s in sheets
            if get_parameter_value(s, "Sheet Number") and
               filter_lower in get_parameter_value(s, "Sheet Number").lower()]


def update_sheet_parameters(sheet, params):
    """Update sheet parameters"""
    updated = []

    mappings = [
        ('status', 'tekening_status'),
        ('issue_date', 'Sheet Issue Date'),
        ('schaal', 'tekening_schaal'),
        ('fase', 'tekening_fase'),
    ]

    for key, param_name in mappings:
        if params.get(key):
            if set_parameter_value(sheet, param_name, params[key]):
                updated.append(key)

    # Wijzigingen
    for letter in ['a', 'b', 'c', 'd', 'e', 'f']:
        datum_key = "wijziging_{}_datum".format(letter)
        omschr_key = "wijziging_{}_omschr".format(letter)

        if params.get(datum_key):
            if set_parameter_value(sheet, "wijziging_{}".format(letter), params[datum_key]):
                updated.append(datum_key)

        if params.get(omschr_key):
            if set_parameter_value(sheet, "wijziging_{}_omschrijving".format(letter), params[omschr_key]):
                updated.append(omschr_key)

    return updated


def update_titleblock_parameters(titleblock, params):
    """Update titleblock parameters. Returns (updated, failed) als param-namen.

    Mapping-waarde mag string of lijst zijn. Bij lijst: eerste param-naam die
    op instance/symbol gevonden wordt, wordt gezet. Backwards-compat met oudere
    titleblock-families.
    """
    updated = []
    failed = []

    mappings = [
        ('std_schaal', ['standaard_schaal']),
        ('v_peil', ['v_peil']),
        ('noordpijl', ['noordpijl']),
        ('kenmerknummer', ['kenmerknummer']),
        ('stempel', ['Stempel', 'stempel']),
        ('aantal_wijzigingen', ['aantal_wijzigingen', 'wijzigingen_op_tek']),
        ('00_3bm_auteur', ['00_3BM_auteur']),
    ]

    for key, candidates in mappings:
        if params.get(key) is None:
            continue
        success = False
        for name in candidates:
            if set_parameter_value(titleblock, name, params[key]):
                updated.append(key)
                success = True
                break
        if not success:
            failed.append(candidates[0])

    return updated, failed


def list_titleblock_param_names(titleblock):
    """Verzamel parameter-namen op instance + symbol voor diagnostiek."""
    names = set()
    try:
        for p in titleblock.Parameters:
            names.add("{} (instance)".format(p.Definition.Name))
    except Exception:
        pass
    symbol = getattr(titleblock, 'Symbol', None)
    if symbol is not None:
        try:
            for p in symbol.Parameters:
                names.add("{} (type)".format(p.Definition.Name))
        except Exception:
            pass
    return sorted(names)


def get_titleblock_from_sheet(sheet):
    """Haal titleblock van sheet"""
    collector = FilteredElementCollector(doc, sheet.Id)\
        .OfCategory(BuiltInCategory.OST_TitleBlocks)\
        .WhereElementIsNotElementType()

    titleblocks = list(collector)
    return titleblocks[0] if titleblocks else None


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    global doc

    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Sheet Parameters")
        return

    log.info("SheetParameters gestart")

    window = SheetParameterWindow()
    if not window.show_dialog():
        log.info("Geannuleerd door gebruiker")
        return

    params = window.get_parameters()
    filter_text = window.get_filter_text()

    if not params:
        forms.alert("Geen parameters geselecteerd om te updaten.", exitscript=True)

    # Verzamel sheets
    all_sheets = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Sheets)\
        .WhereElementIsNotElementType()\
        .ToElements()

    sheets = filter_sheets_by_number(all_sheets, filter_text)

    if not sheets:
        forms.alert("Geen sheets gevonden" +
                   (" met filter '{}'".format(filter_text) if filter_text else ""),
                   exitscript=True)

    # Bevestiging
    msg = "Wilt u {} sheet(s) updaten?".format(len(sheets))
    if filter_text:
        msg += "\n\nFilter: bevat '{}'".format(filter_text)
    msg += "\n\nGeselecteerde parameters: {}".format(len(params))

    if not forms.alert(msg, yes=True, no=True):
        return

    # Transaction
    output = script.get_output()
    output.print_md("## Sheet Parameters Update")
    output.print_md("---")

    with revit.Transaction("Update Sheet Parameters"):
        updated_sheets = 0
        updated_titleblocks = 0
        sheets_zonder_titleblock = 0
        all_failed = set()
        diagnostic_dumped = False

        for sheet in sheets:
            sheet_num = get_parameter_value(sheet, "Sheet Number")
            sheet_name = get_parameter_value(sheet, "Sheet Name")

            sheet_updates = update_sheet_parameters(sheet, params)
            if sheet_updates:
                updated_sheets += 1
                output.print_md("**{}** - {}: {}".format(
                    sheet_num, sheet_name, ", ".join(sheet_updates)))

            titleblock = get_titleblock_from_sheet(sheet)
            if not titleblock:
                sheets_zonder_titleblock += 1
                continue

            tb_updates, tb_failed = update_titleblock_parameters(titleblock, params)
            if tb_updates:
                updated_titleblocks += 1
            if tb_failed:
                all_failed.update(tb_failed)
                # Eerste keer dat NIETS van titleblock-params lukt: dump beschikbare namen
                if not diagnostic_dumped and not tb_updates:
                    diagnostic_dumped = True
                    output.print_md("---")
                    output.print_md(
                        "### Diagnose: titleblock op {} "
                        "({})".format(sheet_num, sheet_name))
                    output.print_md(
                        "Type: `{}`".format(titleblock.Symbol.FamilyName)
                        if getattr(titleblock, 'Symbol', None) else "")
                    output.print_md("**Beschikbare parameter-namen:**")
                    for name in list_titleblock_param_names(titleblock):
                        output.print_md("- `{}`".format(name))

    output.print_md("---")
    output.print_md("**{} sheets** bijgewerkt".format(updated_sheets))
    output.print_md("**{} titleblocks** bijgewerkt".format(updated_titleblocks))
    if sheets_zonder_titleblock:
        output.print_md(
            "**{} sheets zonder titleblock**".format(sheets_zonder_titleblock))
    if all_failed:
        output.print_md(
            "**Niet gevonden/read-only params:** {}".format(
                ", ".join(sorted(all_failed))))

    log.info("Voltooid: {} sheets, {} titleblocks, failed_params={}".format(
        updated_sheets, updated_titleblocks, sorted(all_failed)))


if __name__ == '__main__':
    main()
