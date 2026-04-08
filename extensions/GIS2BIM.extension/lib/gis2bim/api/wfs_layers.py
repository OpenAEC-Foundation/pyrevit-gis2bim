# -*- coding: utf-8 -*-
"""
WFS Layer Configuraties
=======================

Voorgedefinieerde WFS layer configuraties voor Nederlandse geodata.

PDOK (Publieke Dienstverlening Op de Kaart) services:
- Kadastrale Kaart WFS v5: perceelgrenzen, perceelnummers, straatnamen
- BAG WFS: huisnummers, panden
- BGT WFS: grootschalige topografie (toekomstig)

Gebruik:
    from gis2bim.api.wfs_layers import PDOK_LAYERS, get_layer

    percelen_layer = get_layer("kadaster_percelen")
    alle_actieve = get_active_layers()
"""

from .wfs import WFSLayer


# =============================================================================
# PDOK Kadastrale Kaart WFS v5
# =============================================================================

KADASTER_WFS_URL = "https://service.pdok.nl/kadaster/kadastralekaart/wfs/v5_0"

KADASTER_PERCELEN = WFSLayer(
    name="Perceelgrenzen",
    wfs_url=KADASTER_WFS_URL,
    layer_name="kadastralekaartv5:Perceel",
    geometry_type="polygon",
    crs="EPSG:28992",
    label_field="perceelnummer",
    style_color="#2196F3",  # Blauw
    enabled=True,
    category="kadaster"
)

KADASTER_PERCEELNUMMERS = WFSLayer(
    name="Perceelnummers",
    wfs_url=KADASTER_WFS_URL,
    layer_name="kadastralekaartv5:Perceel",
    geometry_type="polygon",  # Geometry voor positie
    crs="EPSG:28992",
    label_field="perceelnummer",
    style_color="#1976D2",
    enabled=True,
    category="kadaster"
)

KADASTER_STRAATNAMEN = WFSLayer(
    name="Straatnamen",
    wfs_url=KADASTER_WFS_URL,
    layer_name="kadastralekaartv5:OpenbareRuimteNaam",
    geometry_type="point",
    crs="EPSG:28992",
    label_field="tekst",
    style_color="#4CAF50",  # Groen
    enabled=True,
    category="kadaster"
)


# =============================================================================
# PDOK BAG WFS
# =============================================================================

BAG_WFS_URL = "https://service.pdok.nl/kadaster/bag/wfs/v2_0"

# NOTE: bag:nummeraanduiding bestaat NIET als WFS layer.
# Huisnummers zitten in bag:verblijfsobject (huisnummer, huisletter, toevoeging).
BAG_HUISNUMMERS = WFSLayer(
    name="Huisnummers",
    wfs_url=BAG_WFS_URL,
    layer_name="bag:verblijfsobject",
    geometry_type="point",
    crs="EPSG:28992",
    label_field="huisnummer",
    style_color="#FF9800",  # Oranje
    enabled=True,
    category="bag"
)

BAG_PAND = WFSLayer(
    name="Panden (BAG)",
    wfs_url=BAG_WFS_URL,
    layer_name="bag:pand",
    geometry_type="polygon",
    crs="EPSG:28992",
    label_field=None,
    style_color="#9C27B0",  # Paars
    enabled=False,  # Standaard uit
    category="bag"
)


# =============================================================================
# PDOK BGT WFS (Basisregistratie Grootschalige Topografie)
# =============================================================================

BGT_WFS_URL = "https://service.pdok.nl/lv/bgt/wfs/v1_0"

BGT_WEGDEEL = WFSLayer(
    name="Wegdelen",
    wfs_url=BGT_WFS_URL,
    layer_name="bgt:wegdeel",
    geometry_type="polygon",
    crs="EPSG:28992",
    label_field=None,
    style_color="#795548",  # Bruin
    enabled=False,
    category="bgt"
)

BGT_PAND = WFSLayer(
    name="Panden (BGT)",
    wfs_url=BGT_WFS_URL,
    layer_name="bgt:pand",
    geometry_type="polygon",
    crs="EPSG:28992",
    label_field=None,
    style_color="#607D8B",  # Blauwgrijs
    enabled=False,
    category="bgt"
)


# =============================================================================
# Layer Registry
# =============================================================================

PDOK_LAYERS = {
    # Kadaster
    "kadaster_percelen": KADASTER_PERCELEN,
    "kadaster_perceelnummers": KADASTER_PERCEELNUMMERS,
    "kadaster_straatnamen": KADASTER_STRAATNAMEN,

    # BAG
    "bag_huisnummers": BAG_HUISNUMMERS,
    "bag_panden": BAG_PAND,

    # BGT (NOTE: WFS endpoint geeft 404, PDOK migreert naar OGC API Features)
    "bgt_wegdelen": BGT_WEGDEEL,
    "bgt_panden": BGT_PAND,
}

# Default layers voor de WFS tool
DEFAULT_LAYERS = [
    "kadaster_percelen",
    "kadaster_perceelnummers",
    "kadaster_straatnamen",
    "bag_huisnummers",
]

# Layers per categorie
LAYER_CATEGORIES = {
    "kadaster": ["kadaster_percelen", "kadaster_perceelnummers", "kadaster_straatnamen"],
    "bag": ["bag_huisnummers", "bag_panden"],
    "bgt": ["bgt_wegdelen", "bgt_panden"],
}


def get_layer(layer_id):
    """
    Haal een layer configuratie op via ID.

    Args:
        layer_id: Layer identifier (bijv. "kadaster_percelen")

    Returns:
        WFSLayer object of None
    """
    return PDOK_LAYERS.get(layer_id)


def get_all_layers():
    """
    Haal alle beschikbare layers op.

    Returns:
        Dictionary met layer_id -> WFSLayer
    """
    return PDOK_LAYERS.copy()


def get_active_layers():
    """
    Haal alle standaard actieve layers op.

    Returns:
        Lijst van WFSLayer objecten
    """
    return [layer for layer in PDOK_LAYERS.values() if layer.enabled]


def get_default_layers():
    """
    Haal de default layer objecten op.

    Returns:
        Lijst van WFSLayer objecten
    """
    return [PDOK_LAYERS[lid] for lid in DEFAULT_LAYERS if lid in PDOK_LAYERS]


def get_layers_by_category(category):
    """
    Haal layers op per categorie.

    Args:
        category: Categorie naam ("kadaster", "bag", "bgt")

    Returns:
        Lijst van WFSLayer objecten
    """
    layer_ids = LAYER_CATEGORIES.get(category, [])
    return [PDOK_LAYERS[lid] for lid in layer_ids if lid in PDOK_LAYERS]


def get_layer_info():
    """
    Haal layer informatie op voor UI.

    Returns:
        Lijst van dictionaries met layer info
    """
    info = []
    for layer_id, layer in PDOK_LAYERS.items():
        info.append({
            "id": layer_id,
            "name": layer.name,
            "category": layer.category,
            "enabled": layer.enabled,
            "geometry_type": layer.geometry_type,
            "has_labels": layer.label_field is not None
        })
    return info
