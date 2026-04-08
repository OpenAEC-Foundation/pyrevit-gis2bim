# -*- coding: utf-8 -*-
"""
Schedule Config Module
======================
Configuratie opslag voor Schedule Export/Import tools.
Slaat sets en export opties op in JSON.
"""

import os
import json

# Config bestand locatie (in lib folder)
CONFIG_DIR = os.path.dirname(__file__)
CONFIG_FILE = os.path.join(CONFIG_DIR, 'schedule_export_config.json')


def _load_config():
    """Laad configuratie uit JSON"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                content = f.read()
                # Handle BOM
                if content.startswith('\xef\xbb\xbf'):
                    content = content[3:]
                return json.loads(content)
        except:
            pass
    return {
        'sets': {},
        'configurations': {},
        'last_export_folder': '',
        'last_set': '',
        'last_configuration': ''
    }


def _save_config(config):
    """Sla configuratie op naar JSON"""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json_str = json.dumps(config, indent=2, ensure_ascii=False)
            if hasattr(json_str, 'encode'):
                f.write(json_str.encode('utf-8'))
            else:
                f.write(json_str)
        return True
    except:
        return False


# ==============================================================================
# SETS MANAGEMENT
# ==============================================================================
def get_sets():
    """Haal alle set namen op"""
    config = _load_config()
    return sorted(config.get('sets', {}).keys())


def get_set(name):
    """
    Haal schedule namen voor een set op.
    
    Returns:
        list met schedule namen, of lege lijst
    """
    config = _load_config()
    return config.get('sets', {}).get(name, [])


def save_set(name, schedule_names):
    """
    Sla set op met schedule namen.
    
    Args:
        name: Set naam
        schedule_names: lijst met schedule namen
    """
    config = _load_config()
    if 'sets' not in config:
        config['sets'] = {}
    config['sets'][name] = list(schedule_names)
    config['last_set'] = name
    return _save_config(config)


def delete_set(name):
    """Verwijder set"""
    config = _load_config()
    if name in config.get('sets', {}):
        del config['sets'][name]
        return _save_config(config)
    return False


def rename_set(old_name, new_name):
    """Hernoem set"""
    config = _load_config()
    if old_name in config.get('sets', {}):
        config['sets'][new_name] = config['sets'].pop(old_name)
        if config.get('last_set') == old_name:
            config['last_set'] = new_name
        return _save_config(config)
    return False


# ==============================================================================
# CONFIGURATIONS MANAGEMENT
# ==============================================================================
def get_configurations():
    """Haal alle configuratie namen op"""
    config = _load_config()
    return sorted(config.get('configurations', {}).keys())


def get_configuration(name):
    """
    Haal configuratie op.
    
    Returns:
        dict met configuratie opties, of defaults
    """
    config = _load_config()
    defaults = {
        'export_folder': '',
        'separate_files': False,
        'file_prefix': '',
        'filename': 'export',
        'include_title': False
    }
    stored = config.get('configurations', {}).get(name, {})
    defaults.update(stored)
    return defaults


def save_configuration(name, options):
    """
    Sla configuratie op.
    
    Args:
        name: Configuratie naam
        options: dict met opties
    """
    config = _load_config()
    if 'configurations' not in config:
        config['configurations'] = {}
    config['configurations'][name] = dict(options)
    config['last_configuration'] = name
    return _save_config(config)


def delete_configuration(name):
    """Verwijder configuratie"""
    config = _load_config()
    if name in config.get('configurations', {}):
        del config['configurations'][name]
        return _save_config(config)
    return False


# ==============================================================================
# LAST USED
# ==============================================================================
def get_last_export_folder():
    """Haal laatst gebruikte export folder op"""
    config = _load_config()
    folder = config.get('last_export_folder', '')
    if folder and os.path.exists(folder):
        return folder
    return os.path.expanduser('~\\Desktop')


def set_last_export_folder(folder):
    """Sla laatst gebruikte export folder op"""
    config = _load_config()
    config['last_export_folder'] = folder
    return _save_config(config)


def get_last_set_name():
    """Haal laatst gebruikte set naam op"""
    config = _load_config()
    return config.get('last_set', '')


def get_last_configuration_name():
    """Haal laatst gebruikte configuratie naam op"""
    config = _load_config()
    return config.get('last_configuration', '')
