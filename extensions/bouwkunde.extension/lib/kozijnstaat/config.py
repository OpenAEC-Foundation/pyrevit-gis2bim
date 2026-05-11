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


def load_config():
    """Laad config - merged defaults + user overrides."""
    cfg = dict(DEFAULTS)
    if os.path.isfile(CONFIG_FILE):
        try:
            f = open(CONFIG_FILE, "r")
            try:
                user = json.load(f)
            finally:
                f.close()
            if isinstance(user, dict):
                cfg.update(user)
        except (ValueError, IOError):
            pass
    return cfg


def save_config(cfg):
    """Sla config op naar user_config.json.

    Slaat alleen keys op die afwijken van DEFAULTS om de file klein
    te houden en future default-wijzigingen door te laten werken.
    """
    diff = {}
    for k, v in cfg.items():
        if k not in DEFAULTS or DEFAULTS[k] != v:
            diff[k] = v

    f = open(CONFIG_FILE, "w")
    try:
        json.dump(diff, f, indent=2, ensure_ascii=False)
    finally:
        f.close()


def reset_config():
    """Verwijder user_config.json zodat defaults weer gelden."""
    if os.path.isfile(CONFIG_FILE):
        os.remove(CONFIG_FILE)
