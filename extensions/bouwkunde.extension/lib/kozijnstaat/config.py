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
    # Family namen
    "kozijn_family": "3BM_kozijn",
    "glas_tag_family": "GEN_glas_v3",

    # Canvas wall (voor auto-generate)
    "canvas_wall_type": "Generic - 200mm",
    "canvas_wall_level": None,          # None = eerste level in project
    "canvas_wall_name": "3BM_Kozijnstaat_Canvas",

    # Grid layout
    "grid_rows": 6,
    "grid_cols": 8,
    "tag_offset_mm": -1000.0,           # Z-offset voor tag-positie t.o.v. kozijn
    "glas_tag_offset_x_mm": -500.0,     # X-offset glas-tag t.o.v. BoundingBox-midden
    "glas_tag_offset_y_mm": 500.0,      # Y-offset glas-tag

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

    # Parameter namen (voor aantallen_tellen)
    "param_aantal": "aantal",
    "param_aantal_gespiegeld": "aantal_gespiegeld",

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
