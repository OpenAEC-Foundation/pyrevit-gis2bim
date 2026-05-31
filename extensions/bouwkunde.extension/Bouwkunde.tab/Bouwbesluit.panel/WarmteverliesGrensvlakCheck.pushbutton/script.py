# -*- coding: utf-8 -*-
"""Warmteverlies Grensvlak-check — visuele controle van SEGC-grensvlakken.

Rendert per ruimte de grensvlakken die de warmteverlies-exporter geometrisch
ziet (uit SpatialElementGeometryCalculator) als gekleurde DirectShapes
(Generic Models) in een dedicated 3D view. Dit is een visuele controle vooraf,
voordat de JSON-export wordt gedraaid.

Eén knop met huisstijl-dialog: 'Tonen' rendert (wist eerst oude shapes),
'Wissen' verwijdert de WV_BND controle-shapes.

Kleurcode:
- dak/plafond = rood, wand = geel, vloer = groen, openingen/vlies = blauw

IronPython 2.7 — geen f-strings, geen type hints.
"""

__title__ = "Grensvlak\nCheck"
__author__ = "3BM Bouwkunde"
__doc__ = (
    "Render de SEGC-grensvlakken per ruimte als gekleurde DirectShapes "
    "ter visuele controle voor de warmteverlies-export. De dialog biedt "
    "ook een 'Wissen'-knop om de WV_BND controle-shapes te verwijderen."
)

import os
import sys

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    View3D,
    ViewFamilyType,
    ViewFamily,
    ViewDetailLevel,
    DisplayStyle,
    Transaction,
    ElementId,
    BuiltInCategory,
    FilteredElementCollector,
    DirectShape,
)

# Lib pad toevoegen
sys.path.append(os.path.join(
    os.path.dirname(__file__), "..", "..", "..", "lib"
))

from wpf_template import WPFWindow

from warmteverlies.room_collector import collect_rooms
from warmteverlies.room_function_mapper import map_all_rooms
from warmteverlies.boundary_preview import (
    ensure_materials,
    render_room_boundaries,
    clear_boundary_shapes,
)

# =============================================================================
# Constanten
# =============================================================================
VIEW_NAME = "WV - Grensvlak-check"
COMMENTS_PREFIX = "WV_BND"
DEFAULT_MIN_FACE_AREA_M2 = 0.5


# =============================================================================
# Helpers — model-inspectie (read-only)
# =============================================================================
def _count_existing_shapes(doc):
    """Tel het aantal WV_BND DirectShapes in het model (read-only)."""
    count = 0
    collector = (
        FilteredElementCollector(doc)
        .OfClass(DirectShape)
        .WhereElementIsNotElementType()
    )
    for ds in collector:
        try:
            param = ds.LookupParameter("Comments")
            if param is None or not param.HasValue:
                continue
            value = param.AsString()
            if value and value.startswith(COMMENTS_PREFIX):
                count += 1
        except Exception:
            continue
    return count


# =============================================================================
# View management
# =============================================================================
def _delete_existing_view(doc, view_name):
    """Verwijder bestaande view met deze naam (geen templates)."""
    collector = (
        FilteredElementCollector(doc)
        .OfClass(View3D)
        .WhereElementIsNotElementType()
    )
    for view in collector:
        if view.IsTemplate:
            continue
        if view.Name == view_name:
            doc.Delete(view.Id)
            return


def _create_3d_view(doc, view_name):
    """Maak een nieuwe isometrische 3D view aan (Shading, Fine)."""
    vft_collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    view_family_type_id = None
    for vft in vft_collector:
        if vft.ViewFamily == ViewFamily.ThreeDimensional:
            view_family_type_id = vft.Id
            break

    if view_family_type_id is None:
        raise Exception("Geen 3D ViewFamilyType gevonden in het model.")

    view_3d = View3D.CreateIsometric(doc, view_family_type_id)
    view_3d.Name = view_name
    view_3d.DetailLevel = ViewDetailLevel.Fine
    try:
        view_3d.DisplayStyle = DisplayStyle.Shading
    except Exception:
        pass
    return view_3d


def _ensure_generic_models_visible(doc, view):
    """Zorg dat de Generic Models categorie zichtbaar is in de view."""
    try:
        cat_id = ElementId(BuiltInCategory.OST_GenericModel)
        if view.CanCategoryBeHidden(cat_id):
            view.SetCategoryHidden(cat_id, False)
    except Exception:
        pass


# =============================================================================
# Rapportage
# =============================================================================
def _print_summary(output, stats, params):
    """Print een samenvatting van de render-statistieken."""
    output.print_md("---")
    output.print_md("## Resultaat")
    output.print_md(
        "Parameters: min. vlakgrootte **{0} m2**, "
        "openingen **{1}**, slivers verbergen **{2}**, "
        "alleen verwarmd **{3}**".format(
            params["min_face_area_m2"],
            "aan" if params["show_openings"] else "uit",
            "aan" if params["hide_hostless_slivers"] else "uit",
            "aan" if params["heated_only"] else "uit",
        )
    )
    output.print_md(
        "Ruimten verwerkt: **{0}** &middot; overgeslagen: {1} "
        "&middot; mislukt: {2}".format(
            stats["rooms_processed"],
            stats["rooms_skipped"],
            stats["rooms_failed"],
        )
    )
    output.print_table(
        [
            ["Dak / plafond (rood)", stats["top"]],
            ["Vloer (groen)", stats["bot"]],
            ["Wand (geel)", stats["wall"]],
            ["Netto wandoppervlak (m2)", round(stats["netto_wall_m2"], 2)],
            ["Gaten gesneden (openingen uit wand)", stats["holes_cut"]],
            ["Vliesgevel (blauw)", stats["open"]],
            ["Openingen deur/raam (blauw)", stats["openings"]],
            ["Host-loze slivers verborgen", stats["slivers_hidden"]],
            ["Faces mislukt", stats["faces_failed"]],
        ],
        columns=["Type", "Aantal"],
    )


# =============================================================================
# WPF Help-venster
# =============================================================================
class HelpWindow(WPFWindow):
    """Huisstijl-uitleg-venster (puur UI, geen Revit-transacties).

    Laadt help.xaml via hetzelfde Window-root-transfer-patroon als de
    hoofddialog en toont scrollbare uitleg met een Sluiten-knop.
    """

    def __init__(self):
        super(HelpWindow, self).__init__(
            xaml_file=None,
            title="Wat doet de Grensvlak-check?",
            width=640,
            height=700,
        )
        self._load_layout()
        if getattr(self, "btn_help_close", None) is not None:
            self.btn_help_close.Click += self._on_close

    def _load_layout(self):
        """Laad help.xaml en bind de named elementen."""
        from System.IO import StringReader
        from System.Xml import XmlReader as SysXmlReader
        from System.Windows.Markup import XamlReader

        xaml_path = os.path.join(os.path.dirname(__file__), "help.xaml")
        with open(xaml_path, "r") as f:
            xaml_content = f.read()

        loaded = XamlReader.Load(SysXmlReader.Create(StringReader(xaml_content)))

        self.Title = loaded.Title
        self.Width = loaded.Width
        self.Height = loaded.Height
        self.WindowStartupLocation = loaded.WindowStartupLocation
        self.ResizeMode = loaded.ResizeMode
        self.Background = loaded.Background
        self.Content = loaded.Content

        for name in ("btn_help_close",):
            element = loaded.FindName(name)
            if element is not None:
                setattr(self, name, element)

    def _on_close(self, sender, args):
        self.Close()


# =============================================================================
# WPF Dialog
# =============================================================================
class GrensvlakCheckWindow(WPFWindow):
    """Huisstijl-dialog voor de Grensvlak-check.

    Zet bij klik self.action ('tonen' / 'wissen' / None) + self.params en
    sluit. Er gebeurt GEEN Revit-transactie in een event-handler; het
    aanroepende script leest de actie na show_dialog() en voert die uit.
    """

    def __init__(self, doc):
        # Base init zonder xaml_file: we laden de XAML zelf (de root is een
        # <Window>, die transfereren we naar dit window — net als de
        # bewezen SharedParamAudit-aanpak). De base laadt wel de huisstijl.
        super(GrensvlakCheckWindow, self).__init__(
            xaml_file=None,
            title="Warmteverlies - Grensvlak-check",
            width=480,
            height=430,
        )

        self.doc = doc
        self.action = None
        self.params = None

        self._load_layout()
        self._populate_existing_count()
        self._bind_events()

    # -----------------------------------------------------------------
    # XAML laden (Window-root → transfer, zelfde patroon als SharedParamAudit)
    # -----------------------------------------------------------------
    def _load_layout(self):
        """Laad de UI.xaml layout en bind de named elementen."""
        from System.IO import StringReader
        from System.Xml import XmlReader as SysXmlReader
        from System.Windows.Markup import XamlReader

        xaml_path = os.path.join(os.path.dirname(__file__), "UI.xaml")
        with open(xaml_path, "r") as f:
            xaml_content = f.read()

        loaded = XamlReader.Load(SysXmlReader.Create(StringReader(xaml_content)))

        # Window-eigenschappen + content overnemen
        self.Title = loaded.Title
        self.Width = loaded.Width
        self.Height = loaded.Height
        self.WindowStartupLocation = loaded.WindowStartupLocation
        self.ResizeMode = loaded.ResizeMode
        self.Background = loaded.Background
        self.Content = loaded.Content

        element_names = [
            "txt_min_area",
            "chk_openings",
            "chk_slivers",
            "chk_heated",
            "txt_existing_count",
            "btn_show",
            "btn_clear",
            "btn_cancel",
            "btn_help",
        ]
        for name in element_names:
            element = loaded.FindName(name)
            if element is not None:
                setattr(self, name, element)

    def _populate_existing_count(self):
        """Toon het huidige aantal WV_BND shapes in het model."""
        try:
            count = _count_existing_shapes(self.doc)
        except Exception:
            count = None

        if count is None:
            self.txt_existing_count.Text = (
                "Huidig in model: onbekend aantal WV_BND controle-shapes"
            )
        elif count == 0:
            self.txt_existing_count.Text = (
                "Huidig in model: geen WV_BND controle-shapes"
            )
        else:
            self.txt_existing_count.Text = (
                "Huidig in model: {0} WV_BND controle-shapes".format(count)
            )

    def _bind_events(self):
        """Koppel knop-events."""
        self.btn_show.Click += self._on_show
        self.btn_clear.Click += self._on_clear
        self.btn_cancel.Click += self._on_cancel
        if getattr(self, "btn_help", None) is not None:
            self.btn_help.Click += self._on_help

    # -----------------------------------------------------------------
    # Help-venster (genest modaal, puur UI — geen Revit-transactie)
    # -----------------------------------------------------------------
    def _on_help(self, sender, args):
        """Open het help/uitleg-venster bovenop deze dialog."""
        try:
            help_window = HelpWindow()
            help_window.Owner = self
            help_window.ShowDialog()
        except Exception as ex:
            self.show_error(
                "Kon het uitleg-venster niet openen:\n{0}".format(str(ex)),
                title="Fout",
            )

    # -----------------------------------------------------------------
    # Parameters lezen
    # -----------------------------------------------------------------
    def _read_params(self):
        """Lees de inputs uit de dialog naar een params-dict."""
        raw = self.txt_min_area.Text if self.txt_min_area.Text else ""
        try:
            min_area = float(raw.replace(",", ".").strip())
        except Exception:
            min_area = DEFAULT_MIN_FACE_AREA_M2

        return {
            "min_face_area_m2": min_area,
            "show_openings": bool(self.chk_openings.IsChecked),
            "hide_hostless_slivers": bool(self.chk_slivers.IsChecked),
            "heated_only": bool(self.chk_heated.IsChecked),
        }

    # -----------------------------------------------------------------
    # Event-handlers — GEEN transactie hier, alleen actie vastleggen + sluiten
    # -----------------------------------------------------------------
    def _on_show(self, sender, args):
        self.action = "tonen"
        self.params = self._read_params()
        self.close_ok()

    def _on_clear(self, sender, args):
        self.action = "wissen"
        self.params = None
        self.close_ok()

    def _on_cancel(self, sender, args):
        self.action = None
        self.params = None
        self.close_cancel()


# =============================================================================
# Acties (buiten de WPF-event-handlers, met transacties)
# =============================================================================
def _do_render(doc, params):
    """Voer de render-flow uit: wis oude shapes, render, maak/activeer view."""
    output = script.get_output()
    output.print_md("## Warmteverlies — Grensvlak-check")

    # --- Rooms ophalen ---
    output.print_md("**Stap 1:** Rooms verzamelen...")
    rooms = collect_rooms(doc)
    if not rooms:
        forms.alert(
            "Geen rooms gevonden in het model.\n"
            "Plaats rooms via Architecture > Room.",
            title="Geen Rooms",
        )
        return

    rooms = map_all_rooms(rooms)
    output.print_md("Gevonden: **{0}** rooms".format(len(rooms)))

    # --- Render in transactie (eerst oude shapes wissen) ---
    output.print_md("**Stap 2:** Grensvlakken renderen...")
    stats = None
    t = Transaction(doc, "WV - Grensvlak-check renderen")
    t.Start()
    try:
        removed = clear_boundary_shapes(doc)
        if removed:
            output.print_md(
                "Oude controle-shapes verwijderd: **{0}**".format(removed)
            )
        material_ids = ensure_materials(doc)
        stats = render_room_boundaries(
            doc, rooms, material_ids, params, output=output
        )
        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(
            "Fout bij renderen grensvlakken:\n{0}".format(str(ex)),
            title="Fout",
        )
        return

    # --- View aanmaken ---
    output.print_md("**Stap 3:** 3D view aanmaken...")
    t2 = Transaction(doc, "WV - Grensvlak-check view")
    t2.Start()
    view_3d = None
    try:
        _delete_existing_view(doc, VIEW_NAME)
        view_3d = _create_3d_view(doc, VIEW_NAME)
        _ensure_generic_models_visible(doc, view_3d)
        t2.Commit()
    except Exception as ex:
        if t2.HasStarted() and not t2.HasEnded():
            t2.RollBack()
        forms.alert(
            "Fout bij aanmaken view:\n{0}".format(str(ex)),
            title="Fout",
        )
        return

    # --- View activeren ---
    if view_3d is not None:
        try:
            revit.uidoc.ActiveView = view_3d
        except Exception:
            output.print_md(
                "*View kon niet automatisch geactiveerd worden. "
                "Open '{0}' handmatig.*".format(VIEW_NAME)
            )

    # --- Samenvatting ---
    _print_summary(output, stats, params)
    output.print_md(
        "Heropen deze knop en kies **Wissen** om de WV_BND controle-shapes "
        "te verwijderen."
    )


def _do_clear(doc):
    """Verwijder alle WV_BND grensvlak-DirectShapes."""
    output = script.get_output()
    output.print_md("## Warmteverlies — Grensvlak wissen")

    count = 0
    t = Transaction(doc, "WV - Grensvlak-check wissen")
    t.Start()
    try:
        count = clear_boundary_shapes(doc)
        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(
            "Fout bij verwijderen grensvlakken:\n{0}".format(str(ex)),
            title="Fout",
        )
        return

    output.print_md(
        "Verwijderd: **{0}** WV_BND controle-shapes.".format(count)
    )


# =============================================================================
# Hoofdfunctie
# =============================================================================
def run_grensvlak_check(doc):
    """Open de huisstijl-dialog en voer de gekozen actie uit."""
    window = GrensvlakCheckWindow(doc)
    window.show_dialog()

    if window.action == "tonen":
        _do_render(doc, window.params)
    elif window.action == "wissen":
        _do_clear(doc)
    # None / Annuleren -> niets doen


# =============================================================================
# Entry point
# =============================================================================
if __name__ == "__main__":
    doc = revit.doc
    if doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run_grensvlak_check(doc)
