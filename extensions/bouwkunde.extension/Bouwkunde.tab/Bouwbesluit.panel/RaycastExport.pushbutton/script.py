# -*- coding: utf-8 -*-
"""Raycast Export -- Thermische schil export via raycast scanning.

Scant alle verwarmde ruimten met ReferenceIntersector voor
cross-model laagopbouw detectie. Exporteert Thermal Import JSON
voor warmteverlies.open-aec.com/import/thermal.

IronPython 2.7 -- geen f-strings, geen type hints.
"""

__title__ = "Raycast\nExport"
__author__ = "3BM Bouwkunde"
__doc__ = "Exporteer thermische schil (raycast) naar JSON voor warmteverlies import"

from pyrevit import revit, DB, forms, script

import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import DialogResult, SaveFileDialog

import json
import os
import sys
import time

from warmteverlies.room_collector import collect_rooms
from warmteverlies.room_function_mapper import map_all_rooms
from warmteverlies.raycast_scanner import scan_room_boundaries
from warmteverlies.thermal_json_builder import build_thermal_import


# =============================================================================
# SelectFromList wrapper class
# =============================================================================
class RoomItem(object):
    """Wrapper voor forms.SelectFromList met custom string representatie."""

    def __init__(self, room_data):
        self.room_data = room_data
        self.name = "{0} - {1} ({2})".format(
            room_data.get("number", "?"),
            room_data.get("name", "Onbekend"),
            room_data.get("level_name", ""),
        )

    def __str__(self):
        return self.name

    def __repr__(self):
        return self.name


# =============================================================================
# 3D View aanmaken (Bug E.5a: altijd verse view zonder view template)
# =============================================================================
# Categorieen die zichtbaar MOETEN zijn voor een betrouwbare raycast-scan.
# Als een view template OST_Doors/OST_Windows/OST_CurtainWallPanels verbergt,
# vindt de scanner 0 openings (Bug E.5a).
_RAYCAST_VISIBLE_CATEGORIES = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_CurtainWallPanels,
    DB.BuiltInCategory.OST_Rooms,
]


def _get_or_create_raycast_view(doc):
    """Maak een schone 3D view specifiek voor raycast scanning.

    Forceer alle relevante categorieen zichtbaar, Fine detail, geen
    view template. Retourneert None als er geen 3D ViewFamilyType
    beschikbaar is.

    De caller is verantwoordelijk voor opruimen van de view na de
    scan (binnen eigen Transaction). Zie Bug E.5a.

    Args:
        doc: Autodesk.Revit.DB.Document

    Returns:
        View3D of None als aanmaken mislukt
    """
    vt_collector = DB.FilteredElementCollector(doc).OfClass(
        DB.ViewFamilyType
    )
    three_d_type = None
    for vt in vt_collector:
        if vt.ViewFamily == DB.ViewFamily.ThreeDimensional:
            three_d_type = vt
            break

    if three_d_type is None:
        return None

    view_name = "RAYCAST_SCAN_{0}".format(int(time.time()))

    t = DB.Transaction(doc, "Create Raycast View")
    t.Start()
    try:
        new_view = DB.View3D.CreateIsometric(doc, three_d_type.Id)
        try:
            new_view.Name = view_name
        except Exception:
            pass

        # Ontkoppel view template (anders dwingt template categorieen
        # weer uit, wat Bug E.5a veroorzaakt).
        try:
            new_view.ViewTemplateId = DB.ElementId.InvalidElementId
        except Exception:
            pass

        # Forceer Fine detail level voor betrouwbare geometrie
        try:
            new_view.DetailLevel = DB.ViewDetailLevel.Fine
        except Exception:
            pass

        # Forceer relevante categorieen zichtbaar
        for bic in _RAYCAST_VISIBLE_CATEGORIES:
            try:
                cat = DB.Category.GetCategory(doc, bic)
                if cat is not None:
                    new_view.SetCategoryHidden(cat.Id, False)
            except Exception:
                pass

        t.Commit()
        return new_view
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise


def _delete_raycast_view(doc, view, output=None):
    """Ruim een tijdelijke raycast view op in eigen Transaction.

    Faalt niet als de delete niet lukt — logt alleen een waarschuwing
    zodat de export niet crasht op cleanup.

    Args:
        doc: Autodesk.Revit.DB.Document
        view: View3D om te verwijderen
        output: pyRevit output object voor logging (optioneel)
    """
    if view is None:
        return
    try:
        view_id = view.Id
    except Exception:
        return

    cleanup_t = DB.Transaction(doc, "Cleanup Raycast View")
    cleanup_t.Start()
    try:
        doc.Delete(view_id)
        cleanup_t.Commit()
    except Exception as ex:
        try:
            if cleanup_t.HasStarted() and not cleanup_t.HasEnded():
                cleanup_t.RollBack()
        except Exception:
            pass
        if output is not None:
            try:
                output.print_md(
                    "  *Waarschuwing: opruimen raycast view "
                    "mislukt: {0}*".format(str(ex))
                )
            except Exception:
                pass


# =============================================================================
# Room element lookup opbouwen
# =============================================================================
def _build_room_element_lookup(doc):
    """Bouw een dict van element_id -> Room element voor snelle lookup."""
    lookup = {}
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
    )
    for room_elem in collector:
        lookup[room_elem.Id.IntegerValue] = room_elem
    return lookup


# =============================================================================
# Hoofdlogica
# =============================================================================
def run_raycast_export(doc):
    """Voer de volledige raycast export workflow uit."""
    output = script.get_output()
    output.print_md("## Raycast Export")
    output.print_md("*Thermische schil export via raycast scanning*")
    output.print_md("---")

    # Stap 1: Rooms verzamelen
    output.print_md("**Stap 1:** Rooms verzamelen...")
    rooms = collect_rooms(doc)
    if not rooms:
        forms.alert(
            "Geen ruimten gevonden in het model.\n"
            "Plaats rooms via Architecture > Room.",
            title="Raycast Export",
        )
        return

    output.print_md("Gevonden: **{0}** rooms".format(len(rooms)))

    # Stap 2: Functies mappen
    output.print_md("**Stap 2:** Functies toewijzen...")
    rooms = map_all_rooms(rooms)

    heated_rooms = [r for r in rooms if r.get("is_heated", False)]
    if not heated_rooms:
        forms.alert(
            "Geen verwarmde ruimten gevonden.\n"
            "Controleer of de ruimtenamen correct zijn ingevuld.",
            title="Raycast Export",
        )
        return

    output.print_md(
        "Verwarmde ruimten: **{0}** van {1}".format(
            len(heated_rooms), len(rooms)
        )
    )

    # Stap 3: Room selectie
    output.print_md("**Stap 3:** Ruimteselectie...")
    room_items = [RoomItem(r) for r in heated_rooms]
    selected_items = forms.SelectFromList.show(
        room_items,
        title="Selecteer ruimten voor raycast scan",
        multiselect=True,
        button_name="Scan",
    )

    if not selected_items:
        output.print_md("*Geen ruimten geselecteerd, export geannuleerd.*")
        return

    # Map terug naar room data dicts
    selected_names = set(str(item) for item in selected_items)
    selected_room_data = [
        item.room_data
        for item in room_items
        if str(item) in selected_names
    ]

    output.print_md(
        "Geselecteerd: **{0}** ruimten".format(len(selected_room_data))
    )

    # Stap 4: Verse 3D view aanmaken (Bug E.5a).
    # Voorkomt dat een bestaande view template OST_Doors/Windows/
    # CurtainWallPanels verbergt en de scan 0 openings vindt.
    output.print_md("**Stap 4:** Verse raycast 3D view aanmaken...")
    try:
        view3d = _get_or_create_raycast_view(doc)
    except Exception as ex:
        forms.alert(
            "Kon geen raycast 3D view aanmaken.\n"
            "Reden: {0}".format(str(ex)),
            title="Raycast Export",
        )
        return

    if view3d is None:
        forms.alert(
            "Geen 3D ViewFamilyType beschikbaar in dit model.\n"
            "Controleer de view types in het project.",
            title="Raycast Export",
        )
        return

    output.print_md(
        "Raycast view aangemaakt: **{0}**".format(view3d.Name)
    )

    # Room element lookup voor scan
    room_elements = _build_room_element_lookup(doc)

    # Scan + cleanup in try/finally: de tijdelijke view MOET altijd
    # worden opgeruimd, ook bij vroege return of exception.
    scan_results = {}
    try:
        # Stap 5: Raycast scan met progress bar
        output.print_md("**Stap 5:** Raycast scanning...")
        scan_count = len(selected_room_data)

        with forms.ProgressBar(
            title="Raycast scanning... ({value} van {max_value})",
            cancellable=True,
        ) as pb:
            for i, room_data in enumerate(selected_room_data):
                if pb.cancelled:
                    output.print_md(
                        "*Scan geannuleerd door gebruiker.*"
                    )
                    break

                pb.update_progress(i, scan_count)

                room_id = room_data["element_id"]
                room_elem = room_elements.get(room_id)
                if room_elem is None:
                    output.print_md(
                        "**Waarschuwing:** Room {0} niet gevonden "
                        "als element, overgeslagen".format(
                            room_data.get("name", "?")
                        )
                    )
                    continue

                try:
                    result = scan_room_boundaries(
                        doc, room_elem, view3d, rooms
                    )
                    scan_results[room_id] = result

                    n_constructions = len(
                        result.get("constructions", [])
                    )
                    n_openings = len(result.get("openings", []))
                    n_connections = len(
                        result.get("open_connections", [])
                    )

                    output.print_md(
                        "  - **{0}**: {1} constructies, "
                        "{2} openingen, {3} open verbindingen".format(
                            room_data.get("name", "?"),
                            n_constructions,
                            n_openings,
                            n_connections,
                        )
                    )
                except Exception as ex:
                    output.print_md(
                        "  - **FOUT bij {0}:** {1}".format(
                            room_data.get("name", "?"), str(ex)
                        )
                    )
    finally:
        # Altijd opruimen — ook bij exception of vroege return.
        _delete_raycast_view(doc, view3d, output=output)

    if not scan_results:
        forms.alert(
            "Geen scan resultaten verkregen.\n"
            "Controleer of de 3D view alle elementen bevat.",
            title="Raycast Export",
        )
        return

    # Stap 6: JSON opbouwen
    output.print_md("**Stap 6:** JSON opbouwen...")

    project_name = doc.Title or "Onbekend project"
    thermal_json = build_thermal_import(
        project_name, selected_room_data, scan_results
    )

    # Samenvatting
    total_rooms = len(
        [r for r in thermal_json.get("rooms", [])
         if r.get("room_type") == "heated"]
    )
    total_constructions = len(
        thermal_json.get("constructions", [])
    )
    total_openings = len(thermal_json.get("openings", []))
    total_connections = len(
        thermal_json.get("open_connections", [])
    )

    output.print_md("---")
    output.print_md("## Scan compleet")
    output.print_md("- **Verwarmde ruimten:** {0}".format(total_rooms))
    output.print_md("- **Constructies:** {0}".format(total_constructions))
    output.print_md("- **Openingen:** {0}".format(total_openings))
    output.print_md(
        "- **Open verbindingen:** {0}".format(total_connections)
    )

    # Stap 7: Opslaan
    output.print_md("**Stap 7:** JSON opslaan...")

    dlg = SaveFileDialog()
    dlg.Filter = "JSON bestanden (*.json)|*.json"
    dlg.DefaultExt = ".json"
    dlg.FileName = "{0}_thermal_import.json".format(
        project_name.replace(" ", "_")
    )

    if dlg.ShowDialog() == DialogResult.OK:
        file_path = dlg.FileName

        try:
            with open(file_path, "w") as f:
                json.dump(thermal_json, f, indent=2, ensure_ascii=False)

            output.print_md("---")
            output.print_md("**Opgeslagen:** `{0}`".format(file_path))
            output.print_md(
                "\nImporteer dit bestand op "
                "[warmteverlies.open-aec.com]"
                "(https://warmteverlies.open-aec.com/import/thermal) "
                "via **Importeer thermische schil**."
            )

            forms.alert(
                "Export geslaagd!\n\n"
                "{0} ruimten, {1} constructies, "
                "{2} openingen, {3} open verbindingen\n\n"
                "Importeer op warmteverlies.open-aec.com/import/thermal"
                .format(
                    total_rooms,
                    total_constructions,
                    total_openings,
                    total_connections,
                ),
                title="Raycast Export",
            )
        except Exception as ex:
            forms.alert(
                "Fout bij opslaan:\n{0}".format(str(ex)),
                title="Export Fout",
            )
            output.print_md("**FOUT:** {0}".format(str(ex)))
    else:
        output.print_md("*Opslaan geannuleerd.*")


# =============================================================================
# Entry point
# =============================================================================
if __name__ == "__main__":
    doc = revit.doc
    if doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run_raycast_export(doc)
