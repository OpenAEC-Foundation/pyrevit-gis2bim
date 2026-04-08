# -*- coding: utf-8 -*-
"""
GIS2BIM Configuration
=====================

Simpele JSON config opslag in %APPDATA%\\GIS2BIM\\config.json.
Gebruikt voor API keys en andere tool-instellingen.
"""

import os
import json

CONFIG_DIR = os.path.join(os.environ.get("APPDATA", ""), "GIS2BIM")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")


def load_config():
    """Laad configuratie uit JSON bestand.

    Returns:
        dict: Configuratie dictionary, leeg dict als bestand niet bestaat.
    """
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except (ValueError, IOError):
        return {}


def save_config(data):
    """Sla configuratie op naar JSON bestand.

    Args:
        data: Dictionary met configuratie.
    """
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=2)


def get_api_key(key_name="google_streetview_api_key"):
    """Haal API key op uit configuratie.

    Args:
        key_name: Naam van de key in config (default: google_streetview_api_key).

    Returns:
        str: API key of lege string als niet gevonden.
    """
    config = load_config()
    return config.get(key_name, "")


def set_api_key(key_value, key_name="google_streetview_api_key"):
    """Sla API key op in configuratie.

    Args:
        key_value: De API key waarde.
        key_name: Naam van de key in config (default: google_streetview_api_key).
    """
    config = load_config()
    config[key_name] = key_value
    save_config(config)
