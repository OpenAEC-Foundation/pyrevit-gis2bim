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
# 3D View zoeken of aanmaken
# =============================================================================
def _get_or_create_3d_view(doc):
    """Zoek een bruikbare 3D view zonder section box.

    Zoekt eerst een bestaande 3D view. Als die niet beschikbaar is,
    wordt een tijdelijke isometrische view aangemaakt.

    Returns:
        View3D of None als aanmaken mislukt
    """
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View3D)
    for view in collector:
        if not view.IsTemplate and not view.IsSectionBoxActive:
            return view

    # Fallback: maak tijdelijke 3D view
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

    with revit.Transaction("Raycast Export - Temp 3D View"):
        new_view = DB.View3D.CreateIsometric(doc, three_d_type.Id)
        return new_view


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

    # Stap 4: 3D view opzoeken
    output.print_md("**Stap 4:** 3D view opzoeken...")
    view3d = _get_or_create_3d_view(doc)
    if view3d is None:
        forms.alert(
            "Geen 3D view beschikbaar en aanmaken is mislukt.\n"
            "Maak handmatig een 3D view aan.",
            title="Raycast Export",
        )
        return

    output.print_md("Gebruik 3D view: **{0}**".format(view3d.Name))

    # Room element lookup voor scan
    room_elements = _build_room_element_lookup(doc)

    # Stap 5: Raycast scan met progress bar
    output.print_md("**Stap 5:** Raycast scanning...")
    scan_results = {}
    scan_count = len(selected_room_data)

    with forms.ProgressBar(
        title="Raycast scanning... ({value} van {max_value})",
        cancellable=True,
    ) as pb:
        for i, room_data in enumerate(selected_room_data):
            if pb.cancelled:
                output.print_md("*Scan geannuleerd door gebruiker.*")
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
