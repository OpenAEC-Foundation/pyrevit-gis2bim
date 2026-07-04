# -*- coding: utf-8 -*-
"""
Regels op de kaart - GIS2BIM
=============================

Vraag de geldende regels op de projectlocatie op uit het DSO-LV
(Ozon Presenteren API) - dezelfde bron als "Regels op de kaart"
in het Omgevingsloket. Toont welke regelingen (omgevingsplan,
omgevingsverordening, waterschapsverordening) gelden en welke
artikelen/tekstdelen een werkingsgebied op de locatie hebben.

Resultaten worden getoond in het output-venster en weggeschreven
als JSON naar %TEMP%\\3bm_exchange\\ voor rapportgeneratie.

Vereist een DSO API-key (gratis via developer.omgevingswet.overheid.nl).
"""

__title__ = "Regels op\nde kaart"
__author__ = "OpenAEC Foundation"
__doc__ = ("Geldende regels op de projectlocatie uit het DSO "
           "(Omgevingsloket 'Regels op de kaart')\n\n"
           "Vereist een DSO API-key van developer.omgevingswet.overheid.nl")

from pyrevit import revit, script, forms

import sys
import os
import json
import traceback

# Voeg lib folder toe aan path
extension_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(extension_path, "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from bm_logger import get_logger

log = get_logger("RegelsOpDeKaart")

GIS2BIM_LOADED = False
IMPORT_ERROR = ""
try:
    from gis2bim import config
    from gis2bim.api.dso import (
        DSOClient,
        DSOAPIError,
        CONFIG_KEY_API_KEY,
        CONFIG_KEY_ENVIRONMENT,
        DEFAULT_ENVIRONMENT,
    )
    from gis2bim.revit.location import get_project_location_rd
    GIS2BIM_LOADED = True
except ImportError as e:
    IMPORT_ERROR = str(e)
    log("Import error: {0}".format(e))
    log(traceback.format_exc())

output = script.get_output()

EXCHANGE_DIR = os.path.join(os.environ.get("TEMP", ""), "3bm_exchange")
EXPORT_FILE = "dso_regels_op_locatie.json"

# Maximaal aantal getoonde annotaties per regeling in het output-venster
MAX_TOON_ANNOTATIES = 100


# ----------------------------------------------------------------------
# Defensieve veld-extractie (API-responsestructuur kan per object varieren)
# ----------------------------------------------------------------------

def _first(waarde, keys):
    """Geef de eerste aanwezige, niet-lege key uit een dict."""
    if not isinstance(waarde, dict):
        return None
    for key in keys:
        if waarde.get(key):
            return waarde[key]
    return None


def _als_tekst(waarde):
    """Maak van een API-veld (string of dict met waarde/naam/code) een tekst."""
    if waarde is None:
        return ""
    if isinstance(waarde, dict):
        return _als_tekst(_first(waarde, ["waarde", "naam", "omschrijving", "code"]))
    if isinstance(waarde, list):
        return ", ".join([_als_tekst(v) for v in waarde if v])
    try:
        return u"{0}".format(waarde)
    except Exception:
        return str(waarde)


def regeling_titel(reg):
    """Beste beschikbare titel van een regeling."""
    return _als_tekst(_first(reg, ["citeerTitel", "officieleTitel", "identificatie"]))


def regeling_type(reg):
    """Documenttype van een regeling (bijv. Omgevingsplan)."""
    return _als_tekst(reg.get("type"))


def regeling_bevoegd_gezag(reg):
    """Naam van het bevoegd gezag dat de regeling aanleverde."""
    bg = _first(reg, ["aangeleverdDoorEen", "bevoegdGezag", "aangeleverdDoor"])
    return _als_tekst(bg)


def annotatie_label(annotatie):
    """Leesbaar label voor een regeltekst-/divisieannotatie."""
    kruimelpad = annotatie.get("documentKruimelpad")
    if isinstance(kruimelpad, list) and kruimelpad:
        delen = []
        for kruimel in kruimelpad:
            tekst = _als_tekst(_first(kruimel, ["nummer", "label", "opschrift"]))
            if tekst:
                delen.append(tekst)
        if delen:
            return " > ".join(delen)
    label = _first(annotatie, ["opschrift", "omschrijving", "wId", "identificatie"])
    return _als_tekst(label)


class RegelingItem(forms.TemplateListItem):
    """Regeling-item voor de selectielijst."""

    @property
    def name(self):
        naam = regeling_titel(self.item)
        doc_type = regeling_type(self.item)
        if doc_type:
            return u"[{0}] {1}".format(doc_type, naam)
        return naam


# ----------------------------------------------------------------------
# Hoofdflow
# ----------------------------------------------------------------------

def get_api_key_interactief():
    """Haal de DSO API-key uit config, of vraag hem eenmalig aan de gebruiker."""
    key = config.get_api_key(CONFIG_KEY_API_KEY)
    if key:
        return key
    key = forms.ask_for_string(
        default="",
        prompt=("Voer je DSO API-key in (aangevraagd via "
                "developer.omgevingswet.overheid.nl).\n"
                "De key wordt lokaal opgeslagen in %APPDATA%\\GIS2BIM\\config.json."),
        title="DSO API-key",
    )
    if key:
        key = key.strip()
        config.set_api_key(key, CONFIG_KEY_API_KEY)
    return key


def exporteer_json(locatie, regelingen, details):
    """Schrijf resultaten naar %TEMP%\\3bm_exchange voor rapportgeneratie."""
    try:
        if not os.path.exists(EXCHANGE_DIR):
            os.makedirs(EXCHANGE_DIR)
        pad = os.path.join(EXCHANGE_DIR, EXPORT_FILE)
        data = {
            "bron": "DSO-LV Ozon Presenteren API v8",
            "locatie_rd": {"x": locatie["rd_x"], "y": locatie["rd_y"]},
            "regelingen": regelingen,
            "regels_per_regeling": details,
        }
        with open(pad, "w") as f:
            json.dump(data, f, indent=2)
        return pad
    except Exception as e:
        log("JSON-export mislukt: {0}".format(e))
        return None


def main():
    if not GIS2BIM_LOADED:
        forms.alert(
            "GIS2BIM modules niet geladen:\n{0}".format(IMPORT_ERROR),
            title="Regels op de kaart",
        )
        return

    doc = revit.doc
    if doc is None:
        forms.alert("Geen actief Revit-document.", title="Regels op de kaart")
        return

    # 1. Projectlocatie in RD
    locatie = get_project_location_rd(doc)
    if not locatie or not locatie.get("rd_x"):
        forms.alert(
            "Geen projectlocatie gevonden.\n\n"
            "Stel eerst de locatie in via GIS2BIM > Setup > Locatie Instellen.",
            title="Regels op de kaart",
        )
        return
    rd_x = locatie["rd_x"]
    rd_y = locatie["rd_y"]

    # 2. API-key
    api_key = get_api_key_interactief()
    if not api_key:
        return

    omgeving = config.load_config().get(CONFIG_KEY_ENVIRONMENT, DEFAULT_ENVIRONMENT)
    client = DSOClient(api_key=api_key, environment=omgeving)

    output.print_md("# Regels op de kaart (DSO)")
    output.print_md("**Locatie (RD):** {0:.1f}, {1:.1f} &nbsp;|&nbsp; "
                    "**Omgeving:** {2}".format(rd_x, rd_y, omgeving))

    # 3. Regelingen zoeken op punt
    try:
        with forms.ProgressBar(title="Regelingen zoeken in DSO...") as pb:
            pb.update_progress(1, 2)
            regelingen = client.zoek_regelingen(rd_x, rd_y)
    except DSOAPIError as e:
        if e.is_auth_error:
            # Key verwijderen zodat er opnieuw gevraagd wordt
            config.set_api_key("", CONFIG_KEY_API_KEY)
        forms.alert(str(e), title="Regels op de kaart")
        return
    except Exception as e:
        log(traceback.format_exc())
        forms.alert("Onverwachte fout: {0}".format(e), title="Regels op de kaart")
        return

    if not regelingen:
        forms.alert(
            "Geen regelingen gevonden op deze locatie.\n"
            "Controleer of de projectlocatie klopt (RD: {0:.0f}, {1:.0f}).".format(
                rd_x, rd_y),
            title="Regels op de kaart",
        )
        return

    output.print_md("## Geldende regelingen ({0})".format(len(regelingen)))
    rows = []
    for reg in regelingen:
        rows.append([
            regeling_type(reg) or "-",
            regeling_titel(reg) or "-",
            regeling_bevoegd_gezag(reg) or "-",
        ])
    output.print_table(
        table_data=rows,
        columns=["Type", "Titel", "Bevoegd gezag"],
    )

    # 4. Selectie voor detail-opvraging
    selectie = forms.SelectFromList.show(
        [RegelingItem(reg) for reg in regelingen],
        title="Regels ophalen voor welke regelingen?",
        multiselect=True,
        button_name="Regels ophalen",
    )
    if not selectie:
        exporteer_json(locatie, regelingen, {})
        return

    # 5. Per regeling de geldende regelteksten/divisies ophalen
    details = {}
    with forms.ProgressBar(title="Regels ophalen... ({value} van {max_value})") as pb:
        for i, reg in enumerate(selectie):
            pb.update_progress(i + 1, len(selectie))
            identificatie = reg.get("identificatie", "")
            titel = regeling_titel(reg)
            output.print_md("## {0}".format(titel))
            try:
                annotaties = client.zoek_regelteksten(identificatie, rd_x, rd_y)
                vorm = "artikelen"
                if not annotaties:
                    annotaties = client.zoek_divisies(identificatie, rd_x, rd_y)
                    vorm = "tekstdelen"
            except DSOAPIError as e:
                output.print_md("*Fout bij ophalen: {0}*".format(e))
                continue

            details[identificatie] = annotaties
            if not annotaties:
                output.print_md("*Geen regels met werkingsgebied op deze locatie.*")
                continue

            output.print_md("**{0} {1} van toepassing:**".format(
                len(annotaties), vorm))
            for annotatie in annotaties[:MAX_TOON_ANNOTATIES]:
                output.print_md("- {0}".format(annotatie_label(annotatie)))
            if len(annotaties) > MAX_TOON_ANNOTATIES:
                output.print_md("*... en {0} meer (zie JSON-export).*".format(
                    len(annotaties) - MAX_TOON_ANNOTATIES))

    # 6. Export voor rapportage
    pad = exporteer_json(locatie, regelingen, details)
    if pad:
        output.print_md("---")
        output.print_md("**JSON-export:** `{0}`".format(pad))
    log("Klaar: {0} regelingen, {1} met detail".format(
        len(regelingen), len(details)))


if __name__ == "__main__":
    main()
