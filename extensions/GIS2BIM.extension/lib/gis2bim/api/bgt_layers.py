# -*- coding: utf-8 -*-
"""
BGT Layer Configuraties
=======================

Voorgedefinieerde BGT (Basisregistratie Grootschalige Topografie) layer configuraties.
Gebruikt de OGC API Features van PDOK.

API: https://api.pdok.nl/lv/bgt/ogc/v1/

Gebruik:
    from gis2bim.api.bgt_layers import BGT_LAYERS, get_bgt_layer

    wegdeel = get_bgt_layer("wegdeel")
    all_terrain = get_bgt_layers_by_category("terrein")
"""

from .ogc_api import OGCAPICollection


# =============================================================================
# PDOK BGT OGC API Features
# =============================================================================

BGT_API_URL = "https://api.pdok.nl/lv/bgt/ogc/v1"


# =============================================================================
# Verharding & Wegen (vlakken)
# =============================================================================

BGT_WEGDEEL = OGCAPICollection(
    collection_id="wegdeel",
    name="Wegdeel",
    geometry_type="polygon",
    category="verharding",
    filled_region_type="bgt_wegdeel",
    enabled=True
)

BGT_ONDERSTEUNENDWEGDEEL = OGCAPICollection(
    collection_id="ondersteunendwegdeel",
    name="Ondersteunend wegdeel",
    geometry_type="polygon",
    category="verharding",
    filled_region_type="bgt_ondersteunendwegdeel",
    enabled=True
)

BGT_OVERBRUGGINGSDEEL = OGCAPICollection(
    collection_id="overbruggingsdeel",
    name="Overbruggingsdeel",
    geometry_type="polygon",
    category="verharding",
    filled_region_type="bgt_overbruggingsdeel",
    enabled=False
)


# =============================================================================
# Terrein (vlakken)
# =============================================================================

BGT_BEGROEIDTERREINDEEL = OGCAPICollection(
    collection_id="begroeidterreindeel",
    name="Begroeid terreindeel",
    geometry_type="polygon",
    category="terrein",
    filled_region_type="bgt_begroeidterreindeel",
    enabled=True
)

BGT_ONBEGROEIDTERREINDEEL = OGCAPICollection(
    collection_id="onbegroeidterreindeel",
    name="Onbegroeid terreindeel",
    geometry_type="polygon",
    category="terrein",
    filled_region_type="bgt_onbegroeidterreindeel",
    enabled=True
)


# =============================================================================
# Water (vlakken)
# =============================================================================

BGT_WATERDEEL = OGCAPICollection(
    collection_id="waterdeel",
    name="Waterdeel",
    geometry_type="polygon",
    category="water",
    filled_region_type="bgt_waterdeel",
    enabled=True
)

BGT_ONDERSTEUNENDWATERDEEL = OGCAPICollection(
    collection_id="ondersteunendwaterdeel",
    name="Ondersteunend waterdeel",
    geometry_type="polygon",
    category="water",
    filled_region_type="bgt_ondersteunendwaterdeel",
    enabled=False
)


# =============================================================================
# Gebouwen & Bouwwerken (vlakken)
# =============================================================================

BGT_PAND = OGCAPICollection(
    collection_id="pand",
    name="Pand",
    geometry_type="polygon",
    category="gebouwen",
    filled_region_type="bgt_pand",
    enabled=True
)

BGT_OVERIGBOUWWERK = OGCAPICollection(
    collection_id="overigbouwwerk",
    name="Overig bouwwerk",
    geometry_type="polygon",
    category="gebouwen",
    filled_region_type="bgt_overigbouwwerk",
    enabled=False
)


# =============================================================================
# Scheidingen (lijnen en vlakken)
# =============================================================================

BGT_SCHEIDING_LIJN = OGCAPICollection(
    collection_id="scheiding_lijn",
    name="Scheiding (lijn)",
    geometry_type="line",
    category="scheidingen",
    line_style="bgt_scheiding",
    enabled=True
)

BGT_SCHEIDING_VLAK = OGCAPICollection(
    collection_id="scheiding_vlak",
    name="Scheiding (vlak)",
    geometry_type="polygon",
    category="scheidingen",
    filled_region_type="bgt_scheiding",
    enabled=False
)


# =============================================================================
# Kunstwerken (lijnen, punten, vlakken)
# =============================================================================

BGT_KUNSTWERKDEEL_LIJN = OGCAPICollection(
    collection_id="kunstwerkdeel_lijn",
    name="Kunstwerkdeel (lijn)",
    geometry_type="line",
    category="kunstwerken",
    line_style="bgt_kunstwerk",
    enabled=False
)

BGT_KUNSTWERKDEEL_VLAK = OGCAPICollection(
    collection_id="kunstwerkdeel_vlak",
    name="Kunstwerkdeel (vlak)",
    geometry_type="polygon",
    category="kunstwerken",
    filled_region_type="bgt_kunstwerk",
    enabled=False
)


# =============================================================================
# Spoor
# =============================================================================

BGT_SPOOR = OGCAPICollection(
    collection_id="spoor",
    name="Spoor",
    geometry_type="line",
    category="infrastructuur",
    line_style="bgt_spoor",
    enabled=False
)


# =============================================================================
# Functioneel gebied & Openbare ruimte
# =============================================================================

BGT_FUNCTIONEELGEBIED = OGCAPICollection(
    collection_id="functioneelgebied",
    name="Functioneel gebied",
    geometry_type="polygon",
    category="gebieden",
    filled_region_type="bgt_functioneelgebied",
    enabled=False
)

BGT_OPENBARERUIMTE = OGCAPICollection(
    collection_id="openbareruimte",
    name="Openbare ruimte",
    geometry_type="polygon",
    category="gebieden",
    filled_region_type="bgt_openbareruimte",
    enabled=False
)


# =============================================================================
# 3D Elementen (punten)
# =============================================================================

BGT_VEGETATIEOBJECT = OGCAPICollection(
    collection_id="vegetatieobject_punt",
    name="Vegetatieobject (bomen)",
    geometry_type="point",
    category="3d_elementen",
    enabled=True
)

BGT_PAAL = OGCAPICollection(
    collection_id="paal",
    name="Paal (lantaarnpalen)",
    geometry_type="point",
    category="3d_elementen",
    enabled=True
)


# =============================================================================
# Kruinlijnen (lijnen)
# =============================================================================

BGT_BEGROEIDTERREINDEEL_KRUINLIJN = OGCAPICollection(
    collection_id="begroeidterreindeel_kruinlijn",
    name="Kruinlijn begroeid",
    geometry_type="line",
    category="kruinlijnen",
    line_style="bgt_kruinlijn",
    enabled=False
)

BGT_ONBEGROEIDTERREINDEEL_KRUINLIJN = OGCAPICollection(
    collection_id="onbegroeidterreindeel_kruinlijn",
    name="Kruinlijn onbegroeid",
    geometry_type="line",
    category="kruinlijnen",
    line_style="bgt_kruinlijn",
    enabled=False
)

BGT_ONDERSTEUNENDWEGDEEL_KRUINLIJN = OGCAPICollection(
    collection_id="ondersteunendwegdeel_kruinlijn",
    name="Kruinlijn ondersteunend wegdeel",
    geometry_type="line",
    category="kruinlijnen",
    line_style="bgt_kruinlijn",
    enabled=False
)


# =============================================================================
# Layer Registry
# =============================================================================

BGT_LAYERS = {
    # Verharding
    "wegdeel": BGT_WEGDEEL,
    "ondersteunendwegdeel": BGT_ONDERSTEUNENDWEGDEEL,
    "overbruggingsdeel": BGT_OVERBRUGGINGSDEEL,

    # Terrein
    "begroeidterreindeel": BGT_BEGROEIDTERREINDEEL,
    "onbegroeidterreindeel": BGT_ONBEGROEIDTERREINDEEL,

    # Water
    "waterdeel": BGT_WATERDEEL,
    "ondersteunendwaterdeel": BGT_ONDERSTEUNENDWATERDEEL,

    # Gebouwen
    "pand": BGT_PAND,
    "overigbouwwerk": BGT_OVERIGBOUWWERK,

    # Scheidingen
    "scheiding_lijn": BGT_SCHEIDING_LIJN,
    "scheiding_vlak": BGT_SCHEIDING_VLAK,

    # Kunstwerken
    "kunstwerkdeel_lijn": BGT_KUNSTWERKDEEL_LIJN,
    "kunstwerkdeel_vlak": BGT_KUNSTWERKDEEL_VLAK,

    # Infrastructuur
    "spoor": BGT_SPOOR,

    # Gebieden
    "functioneelgebied": BGT_FUNCTIONEELGEBIED,
    "openbareruimte": BGT_OPENBARERUIMTE,

    # Kruinlijnen
    "begroeidterreindeel_kruinlijn": BGT_BEGROEIDTERREINDEEL_KRUINLIJN,
    "onbegroeidterreindeel_kruinlijn": BGT_ONBEGROEIDTERREINDEEL_KRUINLIJN,
    "ondersteunendwegdeel_kruinlijn": BGT_ONDERSTEUNENDWEGDEEL_KRUINLIJN,

    # 3D Elementen
    "vegetatieobject_punt": BGT_VEGETATIEOBJECT,
    "paal": BGT_PAAL,
}


# Layers per categorie
LAYER_CATEGORIES = {
    "verharding": ["wegdeel", "ondersteunendwegdeel", "overbruggingsdeel"],
    "terrein": ["begroeidterreindeel", "onbegroeidterreindeel"],
    "water": ["waterdeel", "ondersteunendwaterdeel"],
    "gebouwen": ["pand", "overigbouwwerk"],
    "scheidingen": ["scheiding_lijn", "scheiding_vlak"],
    "kunstwerken": ["kunstwerkdeel_lijn", "kunstwerkdeel_vlak"],
    "infrastructuur": ["spoor"],
    "gebieden": ["functioneelgebied", "openbareruimte"],
    "kruinlijnen": [
        "begroeidterreindeel_kruinlijn",
        "onbegroeidterreindeel_kruinlijn",
        "ondersteunendwegdeel_kruinlijn"
    ],
    "3d_elementen": ["vegetatieobject_punt", "paal"],
}


# Standaard actieve layers
DEFAULT_LAYERS = [
    "wegdeel",
    "ondersteunendwegdeel",
    "begroeidterreindeel",
    "onbegroeidterreindeel",
    "waterdeel",
    "pand",
    "scheiding_lijn",
]


def get_bgt_layer(layer_id):
    """
    Haal een layer configuratie op via ID.

    Args:
        layer_id: Layer identifier (bijv. "wegdeel")

    Returns:
        OGCAPICollection object of None
    """
    return BGT_LAYERS.get(layer_id)


def get_all_bgt_layers():
    """
    Haal alle beschikbare layers op.

    Returns:
        Dictionary met layer_id -> OGCAPICollection
    """
    return BGT_LAYERS.copy()


def get_active_bgt_layers():
    """
    Haal alle standaard actieve layers op.

    Returns:
        Lijst van OGCAPICollection objecten
    """
    return [layer for layer in BGT_LAYERS.values() if layer.enabled]


def get_default_bgt_layers():
    """
    Haal de default layer objecten op.

    Returns:
        Lijst van OGCAPICollection objecten
    """
    return [BGT_LAYERS[lid] for lid in DEFAULT_LAYERS if lid in BGT_LAYERS]


def get_bgt_layers_by_category(category):
    """
    Haal layers op per categorie.

    Args:
        category: Categorie naam ("verharding", "terrein", "water", etc.)

    Returns:
        Lijst van OGCAPICollection objecten
    """
    layer_ids = LAYER_CATEGORIES.get(category, [])
    return [BGT_LAYERS[lid] for lid in layer_ids if lid in BGT_LAYERS]


def get_bgt_layer_info():
    """
    Haal layer informatie op voor UI.

    Returns:
        Lijst van dictionaries met layer info
    """
    info = []
    for layer_id, layer in BGT_LAYERS.items():
        info.append({
            "id": layer_id,
            "name": layer.name,
            "category": layer.category,
            "enabled": layer.enabled,
            "geometry_type": layer.geometry_type,
            "is_polygon": layer.geometry_type in ("polygon", "multipolygon"),
            "is_line": layer.geometry_type in ("line", "multiline"),
        })
    return info
