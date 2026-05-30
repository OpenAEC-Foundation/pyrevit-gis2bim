# -*- coding: utf-8 -*-
"""Warmteverlies Grensvlak-check — visuele controle van SEGC-grensvlakken.

Rendert per ruimte de grensvlakken die de warmteverlies-exporter geometrisch
ziet (uit SpatialElementGeometryCalculator) als gekleurde DirectShapes
(Generic Models) in een dedicated 3D view. Dit is een visuele controle vooraf,
voordat de JSON-export wordt gedraaid.

Kleurcode:
- dak/plafond = rood, wand = geel, vloer = groen, openingen/vlies = blauw

IronPython 2.7 — geen f-strings, geen type hints.
"""

__title__ = "Grensvlak\nCheck"
__author__ = "3BM Bouwkunde"
__doc__ = (
    "Render de SEGC-grensvlakken per ruimte als gekleurde DirectShapes "
    "ter visuele controle voor de warmteverlies-export"
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
)

# Lib pad toevoegen
sys.path.append(os.path.join(
    os.path.dirname(__file__), "..", "..", "..", "lib"
))

from warmteverlies.room_collector import collect_rooms
from warmteverlies.room_function_mapper import map_all_rooms
from warmteverlies.boundary_preview import (
    ensure_materials,
    render_room_boundaries,
)

# =============================================================================
# Constanten
# =============================================================================
VIEW_NAME = "WV - Grensvlak-check"

# Parameter-labels voor de multiselect
OPT_SHOW_OPENINGS = "Openingen tonen (deuren/ramen/vliesgevels)"
OPT_HIDE_SLIVERS = "Host-loze slivers verbergen"
OPT_HEATED_ONLY = "Alleen verwarmde ruimten"

DEFAULT_MIN_FACE_AREA_M2 = 0.10


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
# Parameters verzamelen
# =============================================================================
def _gather_parameters():
    """Verzamel de 4 instelbare parameters via pyRevit forms.

    Returns:
        dict of None (bij annuleren)
    """
    # Stap 1: booleaanse opties via multiselect (default alle drie aan)
    options = [OPT_SHOW_OPENINGS, OPT_HIDE_SLIVERS, OPT_HEATED_ONLY]
    selected = forms.SelectFromList.show(
        options,
        title="Grensvlak-check opties",
        multiselect=True,
        button_name="Volgende",
    )
    if selected is None:
        return None

    selected = list(selected)

    # Stap 2: minimum vlakgrootte
    min_area_str = forms.ask_for_string(
        default=str(DEFAULT_MIN_FACE_AREA_M2),
        prompt="Minimum vlakgrootte (m2) — kleinere vlakken overslaan:",
        title="Grensvlak-check opties",
    )
    if min_area_str is None:
        return None

    try:
        min_area = float(min_area_str.replace(",", "."))
    except Exception:
        min_area = DEFAULT_MIN_FACE_AREA_M2

    return {
        "min_face_area_m2": min_area,
        "show_openings": OPT_SHOW_OPENINGS in selected,
        "hide_hostless_slivers": OPT_HIDE_SLIVERS in selected,
        "heated_only": OPT_HEATED_ONLY in selected,
    }


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
            ["Vliesgevel (blauw)", stats["open"]],
            ["Openingen deur/raam (blauw)", stats["openings"]],
            ["Host-loze slivers verborgen", stats["slivers_hidden"]],
            ["Faces mislukt", stats["faces_failed"]],
        ],
        columns=["Type", "Aantal"],
    )


# =============================================================================
# Hoofdfunctie
# =============================================================================
def run_grensvlak_check(doc):
    """Render de SEGC-grensvlakken als gekleurde DirectShapes."""
    output = script.get_output()
    output.print_md("## Warmteverlies — Grensvlak-check")

    # --- Parameters ---
    params = _gather_parameters()
    if params is None:
        return

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

    # --- Render in transactie ---
    output.print_md("**Stap 2:** Grensvlakken renderen...")
    stats = None
    t = Transaction(doc, "WV - Grensvlak-check renderen")
    t.Start()
    try:
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
        "Gebruik **Grensvlak Wis** om de WV_BND DirectShapes te verwijderen."
    )


# =============================================================================
# Entry point
# =============================================================================
if __name__ == "__main__":
    doc = revit.doc
    if doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run_grensvlak_check(doc)
