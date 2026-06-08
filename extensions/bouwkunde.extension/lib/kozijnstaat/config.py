# -*- coding: utf-8 -*-
"""Kozijnstaat configuratie - family namen en layout parameters.

Opslag in user_config.json naast deze module. Defaults bevinden zich
in DEFAULTS; user_config.json overschrijft alleen gezette keys.
"""

import json
import os

CONFIG_DIR = os.path.dirname(__file__)
CONFIG_FILE = os.path.join(CONFIG_DIR, "user_config.json")

DEFAULTS = {
    # Element-categorie waarin de tool zoekt. "OST_Windows" voor de
    # kozijnstaat, "OST_Doors" voor de deurstaat. Wordt door
    # family_collector.resolve_category() omgezet naar BuiltInCategory.
    "element_category": "OST_Windows",

    # UI-label voor output-titels en dialogen ("Kozijnstaat"/"Deurstaat").
    "tool_label": "Kozijnstaat",
    # Enkelvoud van het element, voor veld-labels ("Kozijn"/"Deur").
    "element_label": "Kozijn",

    # Family namen — 3BM-conventie 2025: "31_"-prefix, "_v4" suffix.
    # Per project override-baar via user_config.json in deze lib-dir.
    "kozijn_family": "31_kozijn",
    "glas_tag_family": "GEN_glas_v4",
    "kozijn_tag_family": "31_TAG_wi_kozijnstaat_window",

    # Canvas wall (voor auto-generate)
    "canvas_wall_type": "Generic - 200mm",
    "canvas_wall_level": None,          # None = eerste level in project
    "canvas_wall_name": "3BM_Kozijnstaat_Canvas",

    # Grid layout
    "grid_rows": 6,                     # legacy — niet meer gebruikt door Create (auto)
    "grid_cols": 8,                     # legacy — niet meer gebruikt door Create (auto)
    "padding_mm": 500.0,                # marge tussen cellen (auto-grid)
    "tag_offset_mm": -1000.0,           # Z-offset voor tag-positie t.o.v. kozijn (legacy)
    "glas_tag_h_offset_mm": 50.0,       # H-offset glas-tag vanaf bottom-left hoek (langs view-right)
    "glas_tag_v_offset_mm": 500.0,      # V-offset glas-tag vanaf bottom-left hoek (langs view-up)
    "kozijn_tag_v_offset_mm": -500.0,   # V-offset kozijn-tag t.o.v. kozijn-bottom (langs view-up)
    "glas_tag_offset_x_mm": -500.0,     # legacy — niet meer gebruikt
    "glas_tag_offset_y_mm": 500.0,      # legacy — niet meer gebruikt

    # Maatvoering named references (voor family "3BM_kozijn")
    "detail_h_refs": [
        "Right",
        "vakvulling_a1_l", "vakvulling_a1_r",
        "vakvulling_b1_l", "vakvulling_b1_r",
        "vakvulling_c1_l", "vakvulling_c1_r",
        "vakvulling_d1_l", "vakvulling_d1_r",
        "vakvulling_e1_l", "vakvulling_e1_r",
        "vakvulling_f1_l", "vakvulling_f1_r",
        "Left",
    ],
    "detail_v_refs": [
        "Sill",
        "vakvulling_a1_o", "vakvulling_a1_b",
        "vakvulling_a2_o", "vakvulling_a2_b",
        "Head",
    ],
    "main_h_refs": ["Right", "Left"],
    "main_v_refs": ["Sill", "Head"],

    # Parameter namen (voor aantallen_tellen).
    # 3BM-conventie (post 2025-05 cleanup):
    #   aantal_getekend  = aantal NIET-gespiegelde instances per type
    #   aantal_gespiegeld = aantal gespiegelde instances per type
    # Som = totaal aantal van dit type in het model.
    "param_aantal": "aantal_getekend",
    "param_aantal_gespiegeld": "aantal_gespiegeld",

    # Sortering: type-parameter waarop Create de kozijnen sorteert.
    # Leeg ("") = fallback naar family- + type-naam.
    "param_merk": "kozijnmerk",

    # Instance-parameters die Create overneemt van de eerst-gevonden
    # model-instance naar de canvas-instance (optie A: 1 stelkozijn per
    # type). 3BM-conventie: stelkozijn = sparing_type (ElementId ref).
    # Bij projecten waar 1 type meerdere stelkozijnen heeft, toont de
    # kozijnstaat de waarde van de EERSTE gevonden instance.
    "instance_params_to_copy": [
        "sparing_type",
        "aftimmering_type",
        "latei_type_cw",
        "vensterbank_type_cw",
        "waterslag_type_cw",
    ],

    # Workset waarin Create de canvas-kozijnen plaatst. Tegelijk wordt
    # deze workset gebruikt om kozijnstaat-instances UIT te sluiten als
    # source bij optie-A copy (anders zou een herhaalde run de canvas
    # van zichzelf laten kopieren). Leeg ("") = workset-feature uit.
    "kozijnstaat_workset_name": "kozijnstaat",

    # True (default): plaats alleen kozijn-TYPES die ergens in het model
    # daadwerkelijk geinstanceerd zijn (buiten de kozijnstaat-workset).
    # Geladen-maar-ongebruikte types worden overgeslagen. Zet op False
    # om alle types op canvas te plaatsen (bv. voor library-overzichten).
    "only_placed_types": True,

    # Legend (POC)
    # Naam van een handmatig aangemaakte template Legend-view die
    # minstens 1 LegendComponent bevat. Vereist omdat Legend-views en
    # LegendComponents niet via de Revit API from-scratch gemaakt
    # kunnen worden — alleen gedupliceerd/gekopieerd.
    "template_legend_name": "TEMPLATE_kozijn_legend",
    "legend_scale": 20,                 # 1:20
    "legend_spacing_ft": 3.0,           # afstand tussen views in feet

    # Filter
    "name_filter_contains": "kozijn",
}


# ---------------------------------------------------------------------------
# Profielen — kozijnstaat (Windows) en deurstaat (Doors) delen dezelfde
# lib. Per profiel een eigen DEFAULTS-set en een eigen user-config-file,
# zodat de profielen elkaars overrides niet overschrijven.
#
# DEUR_OVERRIDES bevat ALLEEN de keys die voor deuren afwijken. Family-
# namen, tag-family en maatvoering-references zijn PLACEHOLDERS — de
# gebruiker zet de juiste 3BM-deurconventie via de Deurstaat-Config UI.
# ---------------------------------------------------------------------------

DEUR_OVERRIDES = {
    "element_category": "OST_Doors",
    "tool_label": "Deurstaat",
    "element_label": "Deur",
    "kozijn_family": "31_deur",        # placeholder — via Config zetten
    "name_filter_contains": "deur",
    "glas_tag_family": "",             # deuren hebben geen glas-tag
    "kozijn_tag_family": "31_TAG_de_deurstaat_door",  # placeholder
    "kozijnstaat_workset_name": "deurstaat",
    "param_merk": "deurmerk",          # placeholder — via Config zetten
    # Maatvoering — named references van de 3BM-deurfamily
    # "32_DO_binnenkozijn_woning", geordend LINKS -> RECHTS zodat de
    # maatketen klopt. LET OP de spelling in de family: "sponing_links"
    # (1 n) vs "sponning_rechts" (2 n). Deze family heeft GEEN named
    # verticale references (Sill/Head bestaan niet), dus verticale
    # maatvoering staat uit tot de family promoted head/sill-planes heeft.
    "detail_h_refs": [
        "Left",
        "sponing_links", "dagmaat links",
        "dagmaat rechts", "sponning_rechts",
        "Right",
    ],
    "detail_v_refs": [],
    "main_h_refs": ["Left", "Right"],
    "main_v_refs": [],
}


def _build_profile_defaults(overrides):
    """DEFAULTS-kopie met profiel-overrides erover."""
    d = dict(DEFAULTS)
    d.update(overrides)
    return d


PROFILES = {
    "kozijn": {
        "defaults": dict(DEFAULTS),
        "config_file": os.path.join(CONFIG_DIR, "user_config.json"),
    },
    "deur": {
        "defaults": _build_profile_defaults(DEUR_OVERRIDES),
        "config_file": os.path.join(CONFIG_DIR, "user_config_deur.json"),
    },
}

DEFAULT_PROFILE = "kozijn"


def _profile(profile):
    """Resolve profielnaam naar de profiel-dict (fallback kozijn)."""
    return PROFILES.get(profile or DEFAULT_PROFILE, PROFILES[DEFAULT_PROFILE])


def profile_defaults(profile=DEFAULT_PROFILE):
    """Defaults-set (zonder user-overrides) voor een profiel."""
    return dict(_profile(profile)["defaults"])


def load_config(profile=DEFAULT_PROFILE):
    """Laad config - merged profiel-defaults + user overrides.

    Args:
        profile: "kozijn" (default, backward-compat) of "deur".
    """
    prof = _profile(profile)
    cfg = dict(prof["defaults"])
    config_file = prof["config_file"]
    if os.path.isfile(config_file):
        try:
            f = open(config_file, "r")
            try:
                user = json.load(f)
            finally:
                f.close()
            if isinstance(user, dict):
                cfg.update(user)
        except (ValueError, IOError):
            pass
    return cfg


def save_config(cfg, profile=DEFAULT_PROFILE):
    """Sla config op naar het user-config-file van het profiel.

    Slaat alleen keys op die afwijken van de profiel-defaults om de
    file klein te houden en future default-wijzigingen door te laten
    werken.
    """
    prof = _profile(profile)
    defaults = prof["defaults"]
    diff = {}
    for k, v in cfg.items():
        if k not in defaults or defaults[k] != v:
            diff[k] = v

    f = open(prof["config_file"], "w")
    try:
        json.dump(diff, f, indent=2, ensure_ascii=False)
    finally:
        f.close()


def reset_config(profile=DEFAULT_PROFILE):
    """Verwijder het user-config-file zodat profiel-defaults weer gelden."""
    config_file = _profile(profile)["config_file"]
    if os.path.isfile(config_file):
        os.remove(config_file)
