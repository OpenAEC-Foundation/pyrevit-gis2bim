# -*- coding: utf-8 -*-
"""Warmteverlies Catalogus Export — ThermalImport JSON voor open-heatloss-studio.

Exporteert de WV_BND catalogus + rooms naar een ThermalImport JSON bestand
dat de warmteverlies-tool kan inlezen voor U-waarde-consolidatie en berekening.

IronPython 2.7 — geen f-strings, geen type hints.
"""

__title__ = "Catalogus\nExport"
__author__ = "3BM Bouwkunde"
__doc__ = (
    "Exporteer de WV_BND catalogus naar ThermalImport JSON voor "
    "open-heatloss-studio. Genereert rooms + dichte constructies "
    "met laagopbouw uit de grensvlak-check."
)

import os
import sys
import json

from pyrevit import revit, forms, script

# Lib pad toevoegen
sys.path.append(os.path.join(
    os.path.dirname(__file__), "..", "..", "..", "lib"
))

from warmteverlies.room_collector import collect_rooms
from warmteverlies.room_function_mapper import map_all_rooms
from warmteverlies.catalog_export import build_catalog_thermal_import


def run_catalog_export():
    """Export de warmteverlies catalogus naar ThermalImport JSON."""
    output = script.get_output()
    doc = revit.doc

    if doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
        return

    output.print_md("## Warmteverlies — Catalogus Export")

    # --- Rooms ophalen ---
    output.print_md("**Stap 1:** Rooms verzamelen...")
    rooms = collect_rooms(doc)
    if not rooms:
        forms.alert(
            "Geen rooms gevonden in het model.\n"
            "Plaats eerst rooms en voer de Grensvlak-check uit.",
            title="Geen Rooms"
        )
        return

    rooms = map_all_rooms(rooms)
    output.print_md("Gevonden: **{0}** rooms".format(len(rooms)))

    # --- Export JSON bouwen ---
    output.print_md("**Stap 2:** ThermalImport JSON bouwen...")
    try:
        thermal_data = build_catalog_thermal_import(doc, rooms)
    except Exception as ex:
        forms.alert(
            "Fout bij bouwen van export data:\n{0}".format(str(ex)),
            title="Export Fout"
        )
        return

    # Rapporteer tellingen
    n_rooms = len(thermal_data.get("rooms", []))
    n_constructions = len(thermal_data.get("constructions", []))
    n_openings = len(thermal_data.get("openings", []))
    debug_info = thermal_data.get("_debug", {})
    warnings = debug_info.get("warnings", [])
    n_fallbacks = debug_info.get("fallback_constructions", 0)

    output.print_md(
        "Export data: **{0}** rooms, **{1}** constructies, **{2}** openingen".format(
            n_rooms, n_constructions, n_openings
        )
    )

    if n_fallbacks > 0:
        output.print_md("Fallback constructies: **{0}** (openingen zonder host-wand)".format(n_fallbacks))

    if warnings:
        output.print_md("**Waarschuwingen:**")
        for warning in warnings:
            output.print_md("- {0}".format(warning))

    # --- Bestand opslaan ---
    output.print_md("**Stap 3:** JSON bestand opslaan...")

    # Standaard bestandsnaam
    project_name = thermal_data.get("project_name", "thermal_import")
    # Sanitize filename
    safe_name = "".join(c for c in project_name if c.isalnum() or c in " -_")
    default_name = "{0}_thermal_import".format(safe_name)

    # Bestand kiezen
    filepath = forms.save_file(
        file_ext='json',
        default_name=default_name
    )

    if filepath is None:
        output.print_md("*Export geannuleerd door gebruiker.*")
        return

    # JSON schrijven
    try:
        # Remove debug info from final export
        export_data = dict(thermal_data)
        if "_debug" in export_data:
            del export_data["_debug"]

        with open(filepath, 'w') as f:
            json.dump(export_data, f, indent=2)

        output.print_md(
            "**Succes!** Export opgeslagen: `{0}`".format(filepath)
        )
        output.print_md(
            "Upload dit bestand naar open-heatloss-studio voor "
            "warmteverlies-berekening."
        )

    except Exception as ex:
        forms.alert(
            "Fout bij schrijven van JSON bestand:\n{0}".format(str(ex)),
            title="Bestand Fout"
        )
        return


# =============================================================================
# Entry point
# =============================================================================
if __name__ == "__main__":
    run_catalog_export()