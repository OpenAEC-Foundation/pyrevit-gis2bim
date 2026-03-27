# -*- coding: utf-8 -*-
"""
Categorie Definities - GebiedsAnalyse
======================================

Definieert voorziening-categorieen met OSM tags, scoringsringen
en kleurcodes voor de gebiedswaarde-analyse.

Elke categorie bevat:
    label:      Weergavenaam in de UI
    osm_tags:   Overpass QL fragments voor deze categorie
    max_score:  Maximale score als POI direct naast gridpunt ligt
    rings:      Afstandsringen in meters (binnen->buiten)
    color:      RGB tuple voor UI weergave
    enabled:    Standaard aan/uit
"""

# Bereikprofielen: max afstand -> ringafstanden in meters
RING_PROFILES = {
    500:  [50, 100, 200, 500],
    1000: [100, 200, 500, 1000],
    2000: [200, 500, 1000, 2000],
    5000: [500, 1000, 2000, 5000],
}

RING_PROFILE_LABELS = {
    500:  "500 m",
    1000: "1.000 m",
    2000: "2.000 m",
    5000: "5.000 m",
}


CATEGORIES = {
    "school": {
        "label": "School / onderwijs",
        "osm_tags": [
            'node["amenity"="school"]', 'way["amenity"="school"]',
            'node["amenity"="kindergarten"]', 'way["amenity"="kindergarten"]',
        ],
        "max_score": 1.0,
        "ring_profile": 5000,
        "rings": [500, 1000, 2000, 5000],
        "color": (69, 182, 168),
        "enabled": True,
    },
    "bus": {
        "label": "Bushalte",
        "osm_tags": [
            'node["highway"="bus_stop"]',
            'node["public_transport"="platform"]["bus"="yes"]',
        ],
        "max_score": 0.5,
        "ring_profile": 500,
        "rings": [50, 100, 200, 500],
        "color": (239, 189, 117),
        "enabled": True,
    },
    "trein": {
        "label": "Treinstation",
        "osm_tags": [
            'node["railway"="station"]', 'node["railway"="halt"]',
            'way["railway"="station"]',
        ],
        "max_score": 1.0,
        "ring_profile": 2000,
        "rings": [200, 500, 1000, 2000],
        "color": (160, 28, 72),
        "enabled": True,
    },
    "park": {
        "label": "Park / groen",
        "osm_tags": [
            'way["leisure"="park"]', 'relation["leisure"="park"]',
            'way["landuse"="forest"]',
        ],
        "max_score": 0.5,
        "ring_profile": 1000,
        "rings": [100, 200, 500, 1000],
        "color": (29, 158, 117),
        "enabled": True,
    },
    "supermarkt": {
        "label": "Supermarkt / winkels",
        "osm_tags": [
            'node["shop"="supermarket"]', 'way["shop"="supermarket"]',
            'node["shop"="convenience"]',
        ],
        "max_score": 0.75,
        "ring_profile": 2000,
        "rings": [200, 500, 1000, 2000],
        "color": (55, 138, 221),
        "enabled": True,
    },
    "ziekenhuis": {
        "label": "Ziekenhuis",
        "osm_tags": [
            'node["amenity"="hospital"]', 'way["amenity"="hospital"]',
        ],
        "max_score": 0.75,
        "ring_profile": 5000,
        "rings": [500, 1000, 2000, 5000],
        "color": (219, 76, 64),
        "enabled": True,
    },
    "huisarts": {
        "label": "Huisarts",
        "osm_tags": [
            'node["amenity"="doctors"]', 'node["healthcare"="doctor"]',
            'way["amenity"="doctors"]',
        ],
        "max_score": 0.5,
        "ring_profile": 2000,
        "rings": [200, 500, 1000, 2000],
        "color": (127, 77, 157),
        "enabled": True,
    },
}


PRESETS = {
    "Woningbouw": {
        "school": {"enabled": True, "max_score": 1.0, "ring_profile": 5000},
        "bus": {"enabled": True, "max_score": 0.5, "ring_profile": 500},
        "trein": {"enabled": True, "max_score": 1.0, "ring_profile": 2000},
        "park": {"enabled": True, "max_score": 0.5, "ring_profile": 1000},
        "supermarkt": {"enabled": True, "max_score": 0.75, "ring_profile": 2000},
        "ziekenhuis": {"enabled": True, "max_score": 0.75, "ring_profile": 5000},
        "huisarts": {"enabled": True, "max_score": 0.5, "ring_profile": 2000},
    },
    "Kantoor": {
        "school": {"enabled": False, "max_score": 0.0, "ring_profile": 5000},
        "bus": {"enabled": True, "max_score": 0.75, "ring_profile": 500},
        "trein": {"enabled": True, "max_score": 1.0, "ring_profile": 1000},
        "park": {"enabled": True, "max_score": 0.25, "ring_profile": 1000},
        "supermarkt": {"enabled": True, "max_score": 0.5, "ring_profile": 1000},
        "ziekenhuis": {"enabled": False, "max_score": 0.0, "ring_profile": 5000},
        "huisarts": {"enabled": False, "max_score": 0.0, "ring_profile": 2000},
    },
}


def get_max_ring(categories):
    """Bepaal de maximale ring-afstand van alle ingeschakelde categorieen.

    Args:
        categories: Dict van categorie-id -> categorie dict (met 'rings' en 'enabled')

    Returns:
        Maximale ringafstand in meters
    """
    max_ring = 0
    for cat in categories.values():
        if cat.get("enabled", True) and cat.get("rings"):
            ring_max = max(cat["rings"])
            if ring_max > max_ring:
                max_ring = ring_max
    return max_ring


def get_all_osm_tags(categories):
    """Verzamel alle OSM tags van ingeschakelde categorieen.

    Args:
        categories: Dict van categorie-id -> categorie dict

    Returns:
        Lijst van unieke Overpass QL tag fragments
    """
    all_tags = []
    for cat in categories.values():
        if not cat.get("enabled", True):
            continue
        for tag in cat.get("osm_tags", []):
            if tag not in all_tags:
                all_tags.append(tag)
    return all_tags


def apply_ring_profile(category):
    """Update de rings van een categorie op basis van het ring_profile.

    Args:
        category: Categorie dict met 'ring_profile' key

    Returns:
        Dezelfde categorie dict met bijgewerkte 'rings'
    """
    profile_key = category.get("ring_profile")
    if profile_key and profile_key in RING_PROFILES:
        category["rings"] = list(RING_PROFILES[profile_key])
    return category


def apply_preset(preset_name):
    """Pas een preset toe op de standaard categorieen.

    Args:
        preset_name: Naam van het preset ("Woningbouw", "Kantoor")

    Returns:
        Kopie van CATEGORIES met preset waarden toegepast, of None
    """
    preset = PRESETS.get(preset_name)
    if not preset:
        return None

    import copy
    result = copy.deepcopy(CATEGORIES)
    for cat_id, overrides in preset.items():
        if cat_id in result:
            result[cat_id].update(overrides)
            apply_ring_profile(result[cat_id])
    return result
