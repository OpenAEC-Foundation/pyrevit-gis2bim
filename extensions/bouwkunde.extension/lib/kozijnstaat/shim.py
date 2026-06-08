# -*- coding: utf-8 -*-
"""Shim-loader voor profiel-hergebruik van Kozijnstaat-pushbuttons.

De Deurstaat-knoppen hergebruiken de Kozijnstaat-button-logica zonder
code-duplicatie: deze helper lokaliseert het bron-script in de
Kozijnstaat.panel en voert zijn run(profile) uit met profile="deur".

Werkt omdat alle gedeelde scripts hun actie in een run(profile)-functie
hebben met een `if __name__ == "__main__"`-guard — imp.load_source draait
de module onder een eigen naam (niet "__main__"), dus de actie start pas
als deze helper run(profile) expliciet aanroept.

IronPython 2.7.
"""

import imp
import os
import sys

_THIS_DIR = os.path.dirname(__file__)       # .../lib/kozijnstaat
_LIB_DIR = os.path.dirname(_THIS_DIR)       # .../lib
_EXT_DIR = os.path.dirname(_LIB_DIR)        # .../bouwkunde.extension
_KOZIJN_PANEL = os.path.join(
    _EXT_DIR, "Bouwkunde.tab", "Kozijnstaat.panel",
)


def run_shared(button_dirname, profile):
    """Laad Kozijnstaat.panel/<button_dirname>/script.py en draai run(profile).

    Args:
        button_dirname: map-naam van de gedeelde pushbutton, bv.
            "KozijnstaatCreate.pushbutton".
        profile: "kozijn" of "deur".

    Raises:
        IOError als het bron-script ontbreekt.
        AttributeError als het bron-script geen run() heeft.
    """
    if _LIB_DIR not in sys.path:
        sys.path.insert(0, _LIB_DIR)

    script_path = os.path.join(_KOZIJN_PANEL, button_dirname, "script.py")
    if not os.path.isfile(script_path):
        raise IOError(
            "Gedeeld script niet gevonden: {0}".format(script_path)
        )

    mod_name = "shared_" + button_dirname.replace(".", "_").replace("-", "_")
    mod = imp.load_source(mod_name, script_path)
    if not hasattr(mod, "run"):
        raise AttributeError(
            "Gedeeld script {0} heeft geen run().".format(button_dirname)
        )
    mod.run(profile)
