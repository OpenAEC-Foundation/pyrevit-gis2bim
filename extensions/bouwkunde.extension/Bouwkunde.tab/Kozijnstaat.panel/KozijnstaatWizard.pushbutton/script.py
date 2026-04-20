# -*- coding: utf-8 -*-
"""Kozijnstaat - Wizard.

Doorloopt de 4 deelstappen sequentieel. Voor elke stap vraagt de
wizard bevestiging zodat de gebruiker kan skippen of stoppen.

IronPython 2.7.
"""

__title__ = "Kozijnstaat\nWizard"
__author__ = "3BM Bouwkunde"
__doc__ = "Doorloop de volledige Kozijnstaat-workflow in 4 stappen"

import imp
import os
import sys

SCRIPT_DIR = os.path.dirname(__file__)
PANEL_DIR = os.path.dirname(SCRIPT_DIR)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(PANEL_DIR))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import revit, forms, script


STEPS = [
    ("Stap 1: Create", "KozijnstaatCreate.pushbutton"),
    ("Stap 2: Maatvoeren", "KozijnstaatMaatvoering.pushbutton"),
    ("Stap 3: Glas taggen", "KozijnstaatGlasTag.pushbutton"),
    ("Stap 4: Aantallen tellen", "KozijnstaatAantallen.pushbutton"),
]


def _load_step(folder_name):
    """Laad script.py uit een zusterknop via imp.load_source."""
    script_path = os.path.join(PANEL_DIR, folder_name, "script.py")
    if not os.path.isfile(script_path):
        return None
    mod_name = "kozijnstaat_wizard_" + folder_name.replace(".", "_")
    return imp.load_source(mod_name, script_path)


def run():
    output = script.get_output()
    output.print_md("## Kozijnstaat - Wizard")
    output.print_md(
        "Doorloopt **{0}** stappen. Bevestig per stap Uitvoeren / "
        "Skip / Stop.".format(len(STEPS))
    )

    executed = []
    skipped = []
    for label, folder in STEPS:
        output.print_md("---")
        output.print_md("### {0}".format(label))

        choice = forms.alert(
            "{0}\n\nUitvoeren?".format(label),
            options=["Uitvoeren", "Skip", "Stop wizard"],
        )
        if choice is None or choice == "Stop wizard":
            output.print_md("*Wizard gestopt door gebruiker.*")
            break
        if choice == "Skip":
            output.print_md("*Overgeslagen.*")
            skipped.append(label)
            continue

        mod = _load_step(folder)
        if mod is None:
            output.print_md(
                "**FOUT:** script niet gevonden voor '{0}'".format(folder)
            )
            continue

        try:
            mod.run()
            executed.append(label)
        except Exception as ex:
            output.print_md(
                "**FOUT in '{0}':** {1}".format(label, ex)
            )
            cont = forms.alert(
                "Stap '{0}' gaf een fout:\n{1}\n\nDoorgaan met "
                "volgende stap?".format(label, ex),
                yes=True, no=True,
            )
            if not cont:
                break

    output.print_md("---")
    output.print_md("## Wizard klaar")
    output.print_md(
        "**Uitgevoerd:** {0}".format(
            ", ".join(executed) if executed else "geen"
        )
    )
    output.print_md(
        "**Overgeslagen:** {0}".format(
            ", ".join(skipped) if skipped else "geen"
        )
    )


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run()
