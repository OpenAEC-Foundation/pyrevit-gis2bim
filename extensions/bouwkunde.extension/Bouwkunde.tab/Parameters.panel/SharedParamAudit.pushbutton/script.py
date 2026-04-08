# -*- coding: utf-8 -*-
"""Shared Parameter Audit — Vind ongebruikte parameters en vergelijk met SPF-bestanden.

Twee functies:
1. Scan alle shared parameters, classificeer gebruik (actief/schedule/filter/leeg/orphan)
2. Vergelijk project-parameters met meerdere .txt shared parameter bestanden
"""

__title__ = "Param\nAudit"
__author__ = "3BM Bouwkunde"
__doc__ = "Audit shared parameters: vind ongebruikte en vergelijk met SPF-bestanden"

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")
clr.AddReference("System.Data")

from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows.Markup import XamlReader
from System.Windows.Controls import DataGridTextColumn, DataGridCheckBoxColumn, DataGridLength
from System.Windows.Data import Binding
import System.Data

from pyrevit import revit, DB, forms

import os
import sys
import codecs

sys.path.append(os.path.join(os.path.dirname(__file__), "..", "..", "..", "lib"))
from bm_logger import get_logger

log = get_logger("SharedParamAudit")


# =============================================================================
# CONSTANTEN
# =============================================================================
STATUS_ACTIVE = "Actief"
STATUS_SCHEDULE_FILTER = "Schedule/Filter"
STATUS_FAMILY = "Familie-parameter"
STATUS_BOUND_EMPTY = "Gebonden (leeg)"
STATUS_UNBOUND = "Ongebonden (orphan)"

ALL_STATUSES = [
    STATUS_ACTIVE,
    STATUS_SCHEDULE_FILTER,
    STATUS_FAMILY,
    STATUS_BOUND_EMPTY,
    STATUS_UNBOUND,
]

COMPARE_MATCH = "Match"
COMPARE_ORPHAN = "Orphan (project)"
COMPARE_NOT_LOADED = "Niet geladen"
COMPARE_NAME_DIFF = "Naam verschil"
COMPARE_DUPLICATE = "Duplicaat"

ALL_COMPARE_RESULTS = [
    COMPARE_MATCH,
    COMPARE_ORPHAN,
    COMPARE_NOT_LOADED,
    COMPARE_NAME_DIFF,
    COMPARE_DUPLICATE,
]

MAX_ELEMENTS_TO_CHECK = 2000

FILTER_ALL = "(alle)"
FILTER_NON_ACTIVE = "(niet-actief)"


# =============================================================================
# DATA CLASSES
# =============================================================================
class ParamInfo:
    """Audit-resultaat voor één shared parameter."""

    def __init__(self, element_id, name, guid, group_name=""):
        self.element_id = element_id
        self.name = name
        self.guid = guid
        self.group_name = group_name
        self.bound_categories = []
        self.is_type_binding = False
        self.in_schedules = False
        self.in_filters = False
        self.has_values = False
        self.in_families = False
        self.status = STATUS_UNBOUND
        self.selected = False


class CompareResult:
    """Vergelijkingsresultaat voor één GUID."""

    def __init__(self, guid, project_name="", file_name="", result_type="", source_file=""):
        self.guid = guid
        self.project_name = project_name
        self.file_name = file_name
        self.result_type = result_type
        self.source_file = source_file


# =============================================================================
# PARAMETER AUDITOR
# =============================================================================
class ParameterAuditor:
    """Verzamelt en classificeert alle shared parameters in het project."""

    def __init__(self, doc):
        self.doc = doc
        self.params = []
        self._schedule_param_ids = set()
        self._filter_param_ids = set()

    def collect_all(self):
        """Voer volledige audit uit en return lijst ParamInfo."""
        self.params = []
        self._collect_shared_params()

        if not self.params:
            return self.params

        self._collect_bindings()
        self._collect_schedule_usage()
        self._collect_filter_usage()
        self._check_values()
        self._classify()
        return self.params

    def _collect_shared_params(self):
        """Verzamel alle SharedParameterElement objecten."""
        collector = DB.FilteredElementCollector(self.doc)
        shared_params = collector.OfClass(DB.SharedParameterElement).ToElements()

        for sp in shared_params:
            try:
                guid_str = str(sp.GuidValue)
                definition = sp.GetDefinition()
                name = definition.Name if definition else sp.Name
                group_name = ""
                try:
                    group_name = str(definition.GetGroupTypeId().TypeId) if definition else ""
                except Exception:
                    pass

                info = ParamInfo(
                    element_id=sp.Id,
                    name=name,
                    guid=guid_str,
                    group_name=group_name,
                )
                self.params.append(info)
            except Exception as ex:
                log.warning("Kan shared param niet lezen: {}".format(ex))

        log.info("Gevonden: {} shared parameters".format(len(self.params)))

    def _collect_bindings(self):
        """Check welke parameters gebonden zijn aan categorieën."""
        binding_map = self.doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()

        # Bouw lookup: definition name -> binding info
        bound_names = {}
        while iterator.MoveNext():
            try:
                definition = iterator.Key
                binding = iterator.Current
                param_name = definition.Name

                categories = []
                is_type = isinstance(binding, DB.TypeBinding)
                cat_set = binding.Categories
                cat_iter = cat_set.ForwardIterator()
                while cat_iter.MoveNext():
                    try:
                        categories.append(cat_iter.Current.Name)
                    except Exception:
                        pass

                bound_names[param_name] = {
                    "categories": categories,
                    "is_type": is_type,
                }
            except Exception as ex:
                log.debug("Binding iterator fout: {}".format(ex))

        # Koppel aan ParamInfo
        for param in self.params:
            if param.name in bound_names:
                info = bound_names[param.name]
                param.bound_categories = info["categories"]
                param.is_type_binding = info["is_type"]

    def _collect_schedule_usage(self):
        """Zoek parameters die in schedules gebruikt worden."""
        self._schedule_param_ids = set()

        collector = DB.FilteredElementCollector(self.doc)
        schedules = collector.OfClass(DB.ViewSchedule).ToElements()

        for schedule in schedules:
            try:
                definition = schedule.Definition
                field_ids = definition.GetFieldOrder()
                for field_id in field_ids:
                    try:
                        field = definition.GetField(field_id)
                        param_id = field.ParameterId
                        if param_id and param_id != DB.ElementId.InvalidElementId:
                            self._schedule_param_ids.add(param_id.IntegerValue)
                    except Exception:
                        pass
            except Exception as ex:
                log.debug("Schedule check fout: {}".format(ex))

        log.info("Parameters in schedules: {}".format(len(self._schedule_param_ids)))

        # Markeer ParamInfo
        for param in self.params:
            if param.element_id.IntegerValue in self._schedule_param_ids:
                param.in_schedules = True

    def _collect_filter_usage(self):
        """Zoek parameters die in view filters gebruikt worden."""
        self._filter_param_ids = set()

        collector = DB.FilteredElementCollector(self.doc)
        filter_elements = collector.OfClass(DB.ParameterFilterElement).ToElements()

        for pfe in filter_elements:
            try:
                ef = pfe.GetElementFilter()
                if ef:
                    self._extract_filter_param_ids(ef)
            except Exception as ex:
                log.debug("Filter check fout: {}".format(ex))

        log.info("Parameters in filters: {}".format(len(self._filter_param_ids)))

        # Markeer ParamInfo
        for param in self.params:
            if param.element_id.IntegerValue in self._filter_param_ids:
                param.in_filters = True

    def _extract_filter_param_ids(self, element_filter):
        """Recursief parameter IDs uit ElementFilter halen."""
        try:
            # FilterRule-gebaseerde filters
            if hasattr(element_filter, "GetRules"):
                for rule in element_filter.GetRules():
                    try:
                        param_id = rule.GetRuleParameter()
                        if param_id and param_id != DB.ElementId.InvalidElementId:
                            self._filter_param_ids.add(param_id.IntegerValue)
                    except Exception:
                        pass

            # Logische filters (AND/OR) recursief uitpakken
            if hasattr(element_filter, "GetFilters"):
                for sub_filter in element_filter.GetFilters():
                    self._extract_filter_param_ids(sub_filter)
        except Exception:
            pass

    def _check_values(self):
        """Single-pass check of parameters waarden hebben op elementen.

        Checkt ALLE params (ook ongebonden) om familie-parameters te detecteren.
        LookupParameter vindt ook params die via families geladen zijn.
        """
        params_to_check = [p for p in self.params if not p.has_values]

        if not params_to_check:
            return

        # Verzamel instance elementen
        collector = DB.FilteredElementCollector(self.doc)
        collector.WhereElementIsNotElementType()
        all_elements = list(collector.ToElements())

        # Verzamel type elementen
        type_collector = DB.FilteredElementCollector(self.doc)
        type_collector.WhereElementIsElementType()
        type_elements = list(type_collector.ToElements())

        # Beperk aantal elementen
        elements_to_check = all_elements[:MAX_ELEMENTS_TO_CHECK]
        types_to_check = type_elements[:MAX_ELEMENTS_TO_CHECK]

        log.info(
            "Waarden checken: {} params, {} instances, {} types".format(
                len(params_to_check), len(elements_to_check), len(types_to_check)
            )
        )

        # Splits in gebonden (project) en ongebonden (potentieel familie)
        bound_params = [p for p in params_to_check if p.bound_categories]
        unbound_params = [p for p in params_to_check if not p.bound_categories]

        # Pass 1: check instance elementen
        unchecked_bound = list(bound_params)
        unchecked_unbound = list(unbound_params)

        for elem in elements_to_check:
            if not unchecked_bound and not unchecked_unbound:
                break

            # Check gebonden params op waarden
            still_unchecked = []
            for pi in unchecked_bound:
                try:
                    rp = elem.LookupParameter(pi.name)
                    if rp and rp.HasValue:
                        pi.has_values = True
                    else:
                        still_unchecked.append(pi)
                except Exception:
                    still_unchecked.append(pi)
            unchecked_bound = still_unchecked

            # Check ongebonden params: als LookupParameter iets vindt
            # is het een familie-parameter
            still_unchecked = []
            for pi in unchecked_unbound:
                try:
                    rp = elem.LookupParameter(pi.name)
                    if rp:
                        pi.in_families = True
                        if rp.HasValue:
                            pi.has_values = True
                    else:
                        still_unchecked.append(pi)
                except Exception:
                    still_unchecked.append(pi)
            unchecked_unbound = still_unchecked

        # Pass 2: check type elementen
        unchecked_bound_t = [p for p in bound_params if p.is_type_binding and not p.has_values]
        unchecked_unbound_t = list(unchecked_unbound)  # restant ongebonden

        for elem in types_to_check:
            if not unchecked_bound_t and not unchecked_unbound_t:
                break

            still_unchecked = []
            for pi in unchecked_bound_t:
                try:
                    rp = elem.LookupParameter(pi.name)
                    if rp and rp.HasValue:
                        pi.has_values = True
                    else:
                        still_unchecked.append(pi)
                except Exception:
                    still_unchecked.append(pi)
            unchecked_bound_t = still_unchecked

            still_unchecked = []
            for pi in unchecked_unbound_t:
                try:
                    rp = elem.LookupParameter(pi.name)
                    if rp:
                        pi.in_families = True
                        if rp.HasValue:
                            pi.has_values = True
                    else:
                        still_unchecked.append(pi)
                except Exception:
                    still_unchecked.append(pi)
            unchecked_unbound_t = still_unchecked

        family_count = sum(1 for p in self.params if p.in_families)
        log.info("Familie-parameters gedetecteerd: {}".format(family_count))

    def _classify(self):
        """Bepaal status per parameter."""
        for param in self.params:
            if param.has_values:
                param.status = STATUS_ACTIVE
            elif param.in_schedules or param.in_filters:
                param.status = STATUS_SCHEDULE_FILTER
            elif param.in_families:
                param.status = STATUS_FAMILY
            elif param.bound_categories:
                param.status = STATUS_BOUND_EMPTY
            else:
                param.status = STATUS_UNBOUND


# =============================================================================
# FILE COMPARER
# =============================================================================
class FileComparer:
    """Vergelijkt project shared parameters met SPF-bestanden."""

    def __init__(self, doc):
        self.doc = doc
        self.app = doc.Application

    def read_spf_file(self, file_path):
        """Lees SPF-bestand en return dict {guid_str: name}."""
        result = {}
        original_spf = self.app.SharedParametersFilename

        try:
            self.app.SharedParametersFilename = file_path
            spf = self.app.OpenSharedParameterFile()

            if not spf:
                log.warning("Kan SPF niet openen: {}".format(file_path))
                return result

            for group in spf.Groups:
                for definition in group.Definitions:
                    try:
                        guid_str = str(definition.GUID)
                        result[guid_str] = definition.Name
                    except Exception:
                        pass
        except Exception as ex:
            log.error("Fout bij lezen SPF {}: {}".format(file_path, ex))
        finally:
            # Altijd origineel herstellen
            try:
                self.app.SharedParametersFilename = original_spf if original_spf else ""
            except Exception:
                pass

        log.info("SPF {}: {} parameters".format(os.path.basename(file_path), len(result)))
        return result

    def compare(self, spf_paths, project_params):
        """Vergelijk project params met meerdere SPF-bestanden.

        Args:
            spf_paths: lijst van bestandspaden
            project_params: lijst van ParamInfo objecten

        Returns:
            lijst van CompareResult
        """
        results = []

        # Lees alle bestanden
        file_data = {}  # {path: {guid: name}}
        for path in spf_paths:
            file_data[path] = self.read_spf_file(path)

        # Bouw project GUID lookup
        project_guids = {}  # {guid: ParamInfo}
        for p in project_params:
            project_guids[p.guid] = p

        # Verzamel alle file GUIDs met bron
        all_file_guids = {}  # {guid: [(path, name), ...]}
        for path, data in file_data.items():
            for guid, name in data.items():
                if guid not in all_file_guids:
                    all_file_guids[guid] = []
                all_file_guids[guid].append((path, name))

        # Check duplicaten (GUID in meerdere bestanden)
        duplicate_guids = set()
        for guid, sources in all_file_guids.items():
            if len(sources) > 1:
                duplicate_guids.add(guid)

        matched_project_guids = set()

        # Vergelijk: bestanden → project
        for guid, sources in all_file_guids.items():
            file_name = sources[0][1]
            source_files = ", ".join(os.path.basename(s[0]) for s in sources)

            if guid in project_guids:
                matched_project_guids.add(guid)
                proj_param = project_guids[guid]

                if guid in duplicate_guids:
                    results.append(
                        CompareResult(
                            guid=guid,
                            project_name=proj_param.name,
                            file_name=file_name,
                            result_type=COMPARE_DUPLICATE,
                            source_file=source_files,
                        )
                    )
                elif proj_param.name != file_name:
                    results.append(
                        CompareResult(
                            guid=guid,
                            project_name=proj_param.name,
                            file_name=file_name,
                            result_type=COMPARE_NAME_DIFF,
                            source_file=source_files,
                        )
                    )
                else:
                    results.append(
                        CompareResult(
                            guid=guid,
                            project_name=proj_param.name,
                            file_name=file_name,
                            result_type=COMPARE_MATCH,
                            source_file=source_files,
                        )
                    )
            else:
                # In bestand maar niet in project
                results.append(
                    CompareResult(
                        guid=guid,
                        project_name="",
                        file_name=file_name,
                        result_type=COMPARE_NOT_LOADED,
                        source_file=source_files,
                    )
                )

        # Orphans: in project maar in geen enkel bestand
        for guid, param in project_guids.items():
            if guid not in matched_project_guids:
                results.append(
                    CompareResult(
                        guid=guid,
                        project_name=param.name,
                        file_name="",
                        result_type=COMPARE_ORPHAN,
                        source_file="",
                    )
                )

        return results


# =============================================================================
# DELETE HELPER
# =============================================================================
def delete_shared_parameters(doc, param_infos):
    """Verwijder geselecteerde shared parameters.

    Args:
        doc: Revit document
        param_infos: lijst van ParamInfo objecten

    Returns:
        tuple (deleted_count, error_messages)
    """
    deleted = 0
    errors = []
    ids_to_delete = [p.element_id for p in param_infos]

    with DB.Transaction(doc, "Shared Parameters verwijderen") as t:
        t.Start()
        for eid in ids_to_delete:
            try:
                doc.Delete(eid)
                deleted += 1
            except Exception as ex:
                errors.append(str(ex))
        t.Commit()

    return deleted, errors


# =============================================================================
# WPF WINDOW
# =============================================================================
class SharedParamAuditWindow(Window):
    """Hoofdvenster voor Shared Parameter Audit."""

    def __init__(self, doc):
        Window.__init__(self)
        self.doc = doc
        self.auditor = ParameterAuditor(doc)
        self.comparer = FileComparer(doc)

        self._params = []          # Huidige scan resultaten
        self._compare_results = [] # Huidige vergelijkingsresultaten
        self._spf_paths = []       # Geselecteerde SPF-bestanden

        self._load_xaml()
        self._setup_filters()
        self._setup_columns()
        self._bind_events()

    # -----------------------------------------------------------------
    # XAML laden
    # -----------------------------------------------------------------
    def _load_xaml(self):
        """Laad XAML layout en bind elementen."""
        xaml_path = os.path.join(os.path.dirname(__file__), "UI.xaml")

        with open(xaml_path, "r") as f:
            xaml_content = f.read()

        reader = StringReader(xaml_content)
        loaded = XamlReader.Load(SysXmlReader.Create(reader))

        # Transfer window properties
        self.Title = loaded.Title
        self.Width = loaded.Width
        self.Height = loaded.Height
        self.WindowStartupLocation = loaded.WindowStartupLocation
        self.ResizeMode = loaded.ResizeMode
        self.Background = loaded.Background
        self.Content = loaded.Content

        # Bind named elements
        element_names = [
            # Header / Footer
            "txt_subtitle", "txt_status",
            # Tab control
            "tab_main",
            # Tab 1: Audit
            "txt_summary", "txt_filter_name", "cmb_filter_status",
            "btn_scan", "grid_params",
            "btn_select_orphans", "btn_select_all", "btn_deselect_all", "txt_selection_count", "btn_delete",
            # Tab 2: Compare
            "lst_spf_files", "btn_browse_spf", "btn_remove_spf",
            "txt_compare_summary", "txt_filter_compare", "cmb_filter_result",
            "btn_compare", "grid_compare", "btn_export_csv",
            # Footer
            "btn_close",
        ]

        for name in element_names:
            element = loaded.FindName(name)
            if element:
                setattr(self, name, element)
            else:
                log.warning("Element niet gevonden: {}".format(name))

    # -----------------------------------------------------------------
    # Setup
    # -----------------------------------------------------------------
    def _setup_filters(self):
        """Vul filter ComboBoxen."""
        # Tab 1: Status filter
        self.cmb_filter_status.Items.Clear()
        self.cmb_filter_status.Items.Add(FILTER_ALL)
        self.cmb_filter_status.Items.Add(FILTER_NON_ACTIVE)
        for status in ALL_STATUSES:
            self.cmb_filter_status.Items.Add(status)
        self.cmb_filter_status.SelectedIndex = 0

        # Tab 2: Result filter
        self.cmb_filter_result.Items.Clear()
        self.cmb_filter_result.Items.Add(FILTER_ALL)
        for result_type in ALL_COMPARE_RESULTS:
            self.cmb_filter_result.Items.Add(result_type)
        self.cmb_filter_result.SelectedIndex = 0

    def _set_filter_non_active(self):
        """Zet status filter op '(niet-actief)' na scan."""
        for i in range(self.cmb_filter_status.Items.Count):
            if str(self.cmb_filter_status.Items[i]) == FILTER_NON_ACTIVE:
                self.cmb_filter_status.SelectedIndex = i
                return

    def _setup_columns(self):
        """Stel DataGrid kolommen in."""
        # Tab 1: Audit grid
        grid = self.grid_params
        grid.Columns.Clear()

        col_sel = DataGridCheckBoxColumn()
        col_sel.Header = ""
        col_sel.Binding = Binding("Geselecteerd")
        col_sel.Width = DataGridLength(32)
        col_sel.IsReadOnly = False
        grid.Columns.Add(col_sel)

        col_name = DataGridTextColumn()
        col_name.Header = "Naam"
        col_name.Binding = Binding("Naam")
        col_name.Width = DataGridLength(200)
        grid.Columns.Add(col_name)

        col_guid = DataGridTextColumn()
        col_guid.Header = "GUID"
        col_guid.Binding = Binding("GUID")
        col_guid.Width = DataGridLength(240)
        grid.Columns.Add(col_guid)

        col_status = DataGridTextColumn()
        col_status.Header = "Status"
        col_status.Binding = Binding("Status")
        col_status.Width = DataGridLength(130)
        grid.Columns.Add(col_status)

        col_cats = DataGridTextColumn()
        col_cats.Header = "Categorieën"
        col_cats.Binding = Binding("Categorieen")
        col_cats.Width = DataGridLength(180)
        grid.Columns.Add(col_cats)

        col_use = DataGridTextColumn()
        col_use.Header = "Gebruik"
        col_use.Binding = Binding("Gebruik")
        col_use.Width = DataGridLength(100)
        grid.Columns.Add(col_use)

        # Checkbox kolom bewerkbaar maken
        grid.IsReadOnly = False

        # Tab 2: Compare grid
        grid2 = self.grid_compare
        grid2.Columns.Clear()

        col_pn = DataGridTextColumn()
        col_pn.Header = "Project Naam"
        col_pn.Binding = Binding("ProjectNaam")
        col_pn.Width = DataGridLength(180)
        grid2.Columns.Add(col_pn)

        col_fn = DataGridTextColumn()
        col_fn.Header = "Bestand Naam"
        col_fn.Binding = Binding("BestandNaam")
        col_fn.Width = DataGridLength(180)
        grid2.Columns.Add(col_fn)

        col_g2 = DataGridTextColumn()
        col_g2.Header = "GUID"
        col_g2.Binding = Binding("GUID")
        col_g2.Width = DataGridLength(240)
        grid2.Columns.Add(col_g2)

        col_res = DataGridTextColumn()
        col_res.Header = "Resultaat"
        col_res.Binding = Binding("Resultaat")
        col_res.Width = DataGridLength(130)
        grid2.Columns.Add(col_res)

        col_src = DataGridTextColumn()
        col_src.Header = "Bronbestand"
        col_src.Binding = Binding("Bronbestand")
        col_src.Width = DataGridLength(150)
        grid2.Columns.Add(col_src)

    def _bind_events(self):
        """Koppel event handlers."""
        self.btn_scan.Click += self._on_scan
        self.btn_close.Click += self._on_close
        self.btn_select_orphans.Click += self._on_select_orphans
        self.btn_select_all.Click += self._on_select_all
        self.btn_deselect_all.Click += self._on_deselect_all
        self.btn_delete.Click += self._on_delete

        self.txt_filter_name.TextChanged += self._on_filter_params
        self.cmb_filter_status.SelectionChanged += self._on_filter_params

        self.btn_browse_spf.Click += self._on_browse_spf
        self.btn_remove_spf.Click += self._on_remove_spf
        self.btn_compare.Click += self._on_compare
        self.btn_export_csv.Click += self._on_export_csv

        self.txt_filter_compare.TextChanged += self._on_filter_compare
        self.cmb_filter_result.SelectionChanged += self._on_filter_compare

    # -----------------------------------------------------------------
    # Tab 1: Scan + Display
    # -----------------------------------------------------------------
    def _on_scan(self, sender, args):
        """Start parameter audit."""
        self.txt_status.Text = "Scannen..."
        try:
            self._params = self.auditor.collect_all()

            # Sorteer: orphans eerst, dan leeg, dan rest
            status_order = {
                STATUS_UNBOUND: 0,
                STATUS_BOUND_EMPTY: 1,
                STATUS_SCHEDULE_FILTER: 2,
                STATUS_FAMILY: 3,
                STATUS_ACTIVE: 4,
            }
            self._params.sort(key=lambda p: (status_order.get(p.status, 9), p.name))

            self._update_summary()

            # Filter standaard op niet-actief (de interessante parameters)
            non_active = [p for p in self._params if p.status != STATUS_ACTIVE]
            if non_active:
                # Zet combobox op "(niet-actief)" — we voegen die optie toe
                self._set_filter_non_active()
                self._populate_audit_grid(non_active)
                self.txt_status.Text = "{} totaal, {} niet-actief getoond".format(
                    len(self._params), len(non_active)
                )
            else:
                self._populate_audit_grid(self._params)
                self.txt_status.Text = "{} parameters, allemaal actief".format(len(self._params))
        except Exception as ex:
            log.exception("Scan fout")
            self.txt_status.Text = "Fout bij scannen"
            MessageBox.Show(
                "Fout bij scannen:\n{}".format(ex),
                "Scan Fout",
                MessageBoxButton.OK,
                MessageBoxImage.Error,
            )

    def _populate_audit_grid(self, params):
        """Vul audit DataGrid met DataTable."""
        dt = System.Data.DataTable("AuditParams")
        dt.Columns.Add("Geselecteerd", clr.GetClrType(System.Boolean))
        dt.Columns.Add("Naam", clr.GetClrType(System.String))
        dt.Columns.Add("GUID", clr.GetClrType(System.String))
        dt.Columns.Add("Status", clr.GetClrType(System.String))
        dt.Columns.Add("Categorieen", clr.GetClrType(System.String))
        dt.Columns.Add("Gebruik", clr.GetClrType(System.String))

        for p in params:
            row = dt.NewRow()
            row["Geselecteerd"] = p.selected
            row["Naam"] = p.name
            row["GUID"] = p.guid
            row["Status"] = p.status
            row["Categorieen"] = ", ".join(p.bound_categories) if p.bound_categories else "-"

            usage_parts = []
            if p.in_schedules:
                usage_parts.append("Schedule")
            if p.in_filters:
                usage_parts.append("Filter")
            if p.has_values:
                usage_parts.append("Waarden")
            row["Gebruik"] = ", ".join(usage_parts) if usage_parts else "-"

            dt.Rows.Add(row)

        self.grid_params.ItemsSource = dt.DefaultView
        self.btn_delete.IsEnabled = len(params) > 0

    def _update_summary(self):
        """Update samenvatting balk."""
        if not self._params:
            self.txt_summary.Text = "Geen shared parameters gevonden"
            return

        counts = {}
        for status in ALL_STATUSES:
            counts[status] = 0
        for p in self._params:
            counts[p.status] = counts.get(p.status, 0) + 1

        self.txt_summary.Text = (
            "Totaal: {}  |  Actief: {}  |  Familie: {}  |  Schedule/Filter: {}  |  Leeg: {}  |  Orphan: {}".format(
                len(self._params),
                counts[STATUS_ACTIVE],
                counts[STATUS_FAMILY],
                counts[STATUS_SCHEDULE_FILTER],
                counts[STATUS_BOUND_EMPTY],
                counts[STATUS_UNBOUND],
            )
        )

    # -----------------------------------------------------------------
    # Tab 1: Filtering
    # -----------------------------------------------------------------
    def _on_filter_params(self, sender, args):
        """Filter audit grid op naam en status."""
        if not self._params:
            return

        name_filter = self.txt_filter_name.Text.strip().lower()
        status_sel = str(self.cmb_filter_status.SelectedItem) if self.cmb_filter_status.SelectedItem else FILTER_ALL

        filtered = []
        for p in self._params:
            if name_filter and name_filter not in p.name.lower():
                continue
            if status_sel == FILTER_NON_ACTIVE and p.status == STATUS_ACTIVE:
                continue
            elif status_sel != FILTER_ALL and status_sel != FILTER_NON_ACTIVE and p.status != status_sel:
                continue
            filtered.append(p)

        self._populate_audit_grid(filtered)

    # -----------------------------------------------------------------
    # Tab 1: Selectie
    # -----------------------------------------------------------------
    def _on_select_orphans(self, sender, args):
        """Selecteer alleen ongebonden (orphan) rijen."""
        view = self.grid_params.ItemsSource
        if not view:
            return
        for row_view in view:
            row_view["Geselecteerd"] = (str(row_view["Status"]) == STATUS_UNBOUND)
        self._update_selection_count()

    def _on_select_all(self, sender, args):
        """Selecteer alle zichtbare rijen."""
        view = self.grid_params.ItemsSource
        if not view:
            return
        for row_view in view:
            row_view["Geselecteerd"] = True
        self._update_selection_count()

    def _on_deselect_all(self, sender, args):
        """Deselecteer alle rijen."""
        view = self.grid_params.ItemsSource
        if not view:
            return
        for row_view in view:
            row_view["Geselecteerd"] = False
        self._update_selection_count()

    def _update_selection_count(self):
        """Teller geselecteerde rijen."""
        view = self.grid_params.ItemsSource
        if not view:
            self.txt_selection_count.Text = ""
            return
        count = 0
        for row_view in view:
            if row_view["Geselecteerd"]:
                count += 1
        self.txt_selection_count.Text = "{} geselecteerd".format(count) if count else ""

    # -----------------------------------------------------------------
    # Tab 1: Verwijderen
    # -----------------------------------------------------------------
    def _on_delete(self, sender, args):
        """Verwijder geselecteerde parameters."""
        # Verzamel geselecteerde GUIDs uit grid
        view = self.grid_params.ItemsSource
        if not view:
            return

        selected_guids = set()
        for row_view in view:
            if row_view["Geselecteerd"]:
                selected_guids.add(str(row_view["GUID"]))

        if not selected_guids:
            MessageBox.Show(
                "Selecteer eerst parameters om te verwijderen.",
                "Geen selectie",
                MessageBoxButton.OK,
                MessageBoxImage.Information,
            )
            return

        # Zoek bijbehorende ParamInfo objecten
        to_delete = [p for p in self._params if p.guid in selected_guids]

        # Bevestiging
        has_active = any(p.status == STATUS_ACTIVE for p in to_delete)
        warning = ""
        if has_active:
            warning = (
                "\n\nLET OP: {} parameter(s) hebben actieve waarden!\n"
                "Verwijderen kan niet ongedaan gemaakt worden (behalve via Undo)."
            ).format(sum(1 for p in to_delete if p.status == STATUS_ACTIVE))

        confirm = MessageBox.Show(
            "{} parameter(s) verwijderen?{}".format(len(to_delete), warning),
            "Bevestig verwijdering",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
        )

        if confirm != MessageBoxResult.Yes:
            return

        # Verwijder
        deleted, errors = delete_shared_parameters(self.doc, to_delete)

        if errors:
            log.warning("Delete fouten: {}".format(errors))

        self.txt_status.Text = "{} verwijderd, {} fouten".format(deleted, len(errors))

        # Re-scan
        self._on_scan(None, None)

    # -----------------------------------------------------------------
    # Tab 2: Bestanden beheren
    # -----------------------------------------------------------------
    def _on_browse_spf(self, sender, args):
        """Voeg SPF-bestanden toe via bestandsdialoog."""
        from Microsoft.Win32 import OpenFileDialog

        dlg = OpenFileDialog()
        dlg.Title = "Selecteer Shared Parameter Bestand(en)"
        dlg.Filter = "Shared Parameter Files (*.txt)|*.txt|All Files (*.*)|*.*"
        dlg.Multiselect = True

        if dlg.ShowDialog():
            for path in dlg.FileNames:
                if path not in self._spf_paths:
                    self._spf_paths.append(path)
                    self.lst_spf_files.Items.Add(os.path.basename(path))

    def _on_remove_spf(self, sender, args):
        """Verwijder geselecteerd bestand uit lijst."""
        idx = self.lst_spf_files.SelectedIndex
        if idx >= 0 and idx < len(self._spf_paths):
            self._spf_paths.pop(idx)
            self.lst_spf_files.Items.RemoveAt(idx)

    # -----------------------------------------------------------------
    # Tab 2: Vergelijking
    # -----------------------------------------------------------------
    def _on_compare(self, sender, args):
        """Start vergelijking."""
        if not self._spf_paths:
            MessageBox.Show(
                "Voeg eerst SPF-bestanden toe.",
                "Geen bestanden",
                MessageBoxButton.OK,
                MessageBoxImage.Information,
            )
            return

        self.txt_status.Text = "Vergelijken..."

        try:
            # Zorg dat we actuele project params hebben
            if not self._params:
                self._params = self.auditor.collect_all()

            self._compare_results = self.comparer.compare(self._spf_paths, self._params)
            self._populate_compare_grid(self._compare_results)
            self._update_compare_summary()
            self.btn_export_csv.IsEnabled = len(self._compare_results) > 0
            self.txt_status.Text = "{} vergelijkingsresultaten".format(len(self._compare_results))
        except Exception as ex:
            log.exception("Vergelijk fout")
            self.txt_status.Text = "Fout bij vergelijken"
            MessageBox.Show(
                "Fout bij vergelijken:\n{}".format(ex),
                "Vergelijk Fout",
                MessageBoxButton.OK,
                MessageBoxImage.Error,
            )

    def _populate_compare_grid(self, results):
        """Vul compare DataGrid."""
        dt = System.Data.DataTable("CompareResults")
        dt.Columns.Add("ProjectNaam", clr.GetClrType(System.String))
        dt.Columns.Add("BestandNaam", clr.GetClrType(System.String))
        dt.Columns.Add("GUID", clr.GetClrType(System.String))
        dt.Columns.Add("Resultaat", clr.GetClrType(System.String))
        dt.Columns.Add("Bronbestand", clr.GetClrType(System.String))

        for r in results:
            row = dt.NewRow()
            row["ProjectNaam"] = r.project_name if r.project_name else "-"
            row["BestandNaam"] = r.file_name if r.file_name else "-"
            row["GUID"] = r.guid
            row["Resultaat"] = r.result_type
            row["Bronbestand"] = r.source_file if r.source_file else "-"
            dt.Rows.Add(row)

        self.grid_compare.ItemsSource = dt.DefaultView

    def _update_compare_summary(self):
        """Update vergelijkings-samenvatting."""
        if not self._compare_results:
            self.txt_compare_summary.Text = "Geen resultaten"
            return

        counts = {}
        for rt in ALL_COMPARE_RESULTS:
            counts[rt] = 0
        for r in self._compare_results:
            counts[r.result_type] = counts.get(r.result_type, 0) + 1

        self.txt_compare_summary.Text = (
            "Match: {}  |  Orphan: {}  |  Niet geladen: {}  |  Naam verschil: {}  |  Duplicaat: {}".format(
                counts[COMPARE_MATCH],
                counts[COMPARE_ORPHAN],
                counts[COMPARE_NOT_LOADED],
                counts[COMPARE_NAME_DIFF],
                counts[COMPARE_DUPLICATE],
            )
        )

    # -----------------------------------------------------------------
    # Tab 2: Filtering
    # -----------------------------------------------------------------
    def _on_filter_compare(self, sender, args):
        """Filter compare grid."""
        if not self._compare_results:
            return

        name_filter = self.txt_filter_compare.Text.strip().lower()
        result_sel = str(self.cmb_filter_result.SelectedItem) if self.cmb_filter_result.SelectedItem else FILTER_ALL

        filtered = []
        for r in self._compare_results:
            searchable = (r.project_name + " " + r.file_name).lower()
            if name_filter and name_filter not in searchable:
                continue
            if result_sel != FILTER_ALL and r.result_type != result_sel:
                continue
            filtered.append(r)

        self._populate_compare_grid(filtered)

    # -----------------------------------------------------------------
    # Tab 2: CSV Export
    # -----------------------------------------------------------------
    def _on_export_csv(self, sender, args):
        """Exporteer vergelijkingsresultaten naar CSV."""
        if not self._compare_results:
            return

        from Microsoft.Win32 import SaveFileDialog

        dlg = SaveFileDialog()
        dlg.Title = "Exporteer vergelijkingsresultaten"
        dlg.Filter = "CSV bestanden (*.csv)|*.csv"
        dlg.DefaultExt = ".csv"
        dlg.FileName = "shared_param_vergelijking.csv"

        if not dlg.ShowDialog():
            return

        try:
            with codecs.open(dlg.FileName, "w", encoding="utf-8-sig") as f:
                # Header
                f.write("Project Naam;Bestand Naam;GUID;Resultaat;Bronbestand\n")

                for r in self._compare_results:
                    f.write("{};{};{};{};{}\n".format(
                        r.project_name or "-",
                        r.file_name or "-",
                        r.guid,
                        r.result_type,
                        r.source_file or "-",
                    ))

            self.txt_status.Text = "Geexporteerd naar: {}".format(os.path.basename(dlg.FileName))
            log.info("CSV export: {}".format(dlg.FileName))
        except Exception as ex:
            log.exception("CSV export fout")
            MessageBox.Show(
                "Fout bij exporteren:\n{}".format(ex),
                "Export Fout",
                MessageBoxButton.OK,
                MessageBoxImage.Error,
            )

    # -----------------------------------------------------------------
    # Sluiten
    # -----------------------------------------------------------------
    def _on_close(self, sender, args):
        """Sluit venster."""
        self.Close()


# =============================================================================
# MAIN
# =============================================================================
def main():
    """Entry point."""
    log.info("=== SharedParamAudit gestart ===")
    log.log_revit_info()

    doc = revit.doc
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Shared Parameter Audit")
        return

    try:
        window = SharedParamAuditWindow(doc)
        window.ShowDialog()
    except Exception as ex:
        log.exception("SharedParamAudit fout")
        forms.alert(
            "Fout bij laden Shared Parameter Audit:\n\n{}".format(ex),
            title="Fout",
        )

    log.info("=== SharedParamAudit afgesloten ===")


if __name__ == "__main__":
    main()
