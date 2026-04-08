# -*- coding: utf-8 -*-
"""
OSM Layer Configuraties
=======================

Voorgedefinieerde OpenStreetMap layer configuraties voor de Overpass API.
Alleen polygon layers - voor filled regions in Revit.

Naamconventie filled region types (conform Dynamo):
    osm_{feature}_{subtype}
    Bijv: osm_building_residential, osm_landuse_forest, osm_natural_wood

Fallback volgorde:
    1. osm_building_residential  (specifiek)
    2. osm_building              (generiek)
    3. Default FilledRegionType   (laatste fallback)

Gebruik:
    from gis2bim.api.osm_layers import OSM_LAYERS, get_osm_layer

    buildings = get_osm_layer("building")
    query_tags = buildings.query_tags  # ['way["building"]', 'relation["building"]']
"""


class OverpassLayer(object):
    """
    Configuratie voor een OSM/Overpass data layer.

    Attributes:
        layer_id: Unieke identifier
        name: Weergavenaam voor UI
        query_tags: Lijst van Overpass QL fragments
        geometry_type: Verwacht type ("polygon" of "line")
        category: Categorie voor UI groepering
        filled_region_type: Naam van FilledRegionType in Revit (voor vlakken)
        line_style: Naam van line style in Revit (voor lijnen)
        tag_key: OSM tag key voor subtype (bijv. "building", "landuse")
                 Wordt gebruikt om per-subtype filled region types te zoeken:
                 osm_building_residential, osm_landuse_forest, etc.
        enabled: Of de layer standaard actief is
    """

    def __init__(self, layer_id, name, query_tags, geometry_type="polygon",
                 category="osm", filled_region_type=None, line_style=None,
                 tag_key=None, enabled=True):
        self.layer_id = layer_id
        self.name = name
        self.query_tags = query_tags
        self.geometry_type = geometry_type
        self.category = category
        self.filled_region_type = filled_region_type
        self.line_style = line_style
        self.tag_key = tag_key
        self.enabled = enabled

    def __repr__(self):
        return "OverpassLayer('{0}', '{1}')".format(
            self.layer_id, self.name
        )


# =============================================================================
# Gebouwen
# =============================================================================

OSM_BUILDING = OverpassLayer(
    layer_id="building",
    name="Gebouwen",
    query_tags=[
        'way["building"]',
        'relation["building"]',
    ],
    geometry_type="polygon",
    category="gebouwen",
    filled_region_type="osm_building",
    tag_key="building",
    enabled=True
)


# =============================================================================
# Water
# =============================================================================

OSM_WATER = OverpassLayer(
    layer_id="water",
    name="Water",
    query_tags=[
        'way["natural"="water"]',
        'relation["natural"="water"]',
        'way["water"]',
        'relation["water"]',
        'way["waterway"~"riverbank|dock|canal"]',
    ],
    geometry_type="polygon",
    category="water",
    filled_region_type="osm_water",
    tag_key="water",
    enabled=True
)


# =============================================================================
# Terrein / Landgebruik
# =============================================================================

OSM_LANDUSE = OverpassLayer(
    layer_id="landuse",
    name="Landgebruik",
    query_tags=[
        'way["landuse"]',
        'relation["landuse"]',
    ],
    geometry_type="polygon",
    category="terrein",
    filled_region_type="osm_landuse",
    tag_key="landuse",
    enabled=True
)

OSM_NATURAL = OverpassLayer(
    layer_id="natural",
    name="Natuur",
    query_tags=[
        'way["natural"]',
        'relation["natural"]',
    ],
    geometry_type="polygon",
    category="terrein",
    filled_region_type="osm_natural",
    tag_key="natural",
    enabled=True
)


# =============================================================================
# Voorzieningen
# =============================================================================

OSM_AMENITY = OverpassLayer(
    layer_id="amenity",
    name="Voorzieningen",
    query_tags=[
        'way["amenity"]',
        'relation["amenity"]',
    ],
    geometry_type="polygon",
    category="voorzieningen",
    filled_region_type="osm_amenity",
    tag_key="amenity",
    enabled=True
)

OSM_LEISURE = OverpassLayer(
    layer_id="leisure",
    name="Recreatie (parken, sport)",
    query_tags=[
        'way["leisure"]',
        'relation["leisure"]',
    ],
    geometry_type="polygon",
    category="voorzieningen",
    filled_region_type="osm_leisure",
    tag_key="leisure",
    enabled=True
)


# =============================================================================
# Layer Registry
# =============================================================================

OSM_LAYERS = {
    "building": OSM_BUILDING,
    "water": OSM_WATER,
    "landuse": OSM_LANDUSE,
    "natural": OSM_NATURAL,
    "amenity": OSM_AMENITY,
    "leisure": OSM_LEISURE,
}

LAYER_CATEGORIES = {
    "gebouwen": ["building"],
    "water": ["water"],
    "terrein": ["landuse", "natural"],
    "voorzieningen": ["amenity", "leisure"],
}

DEFAULT_LAYERS = [
    "building",
    "water",
    "landuse",
    "natural",
    "amenity",
    "leisure",
]


def get_osm_layer(layer_id):
    """
    Haal een layer configuratie op via ID.

    Args:
        layer_id: Layer identifier (bijv. "building")

    Returns:
        OverpassLayer object of None
    """
    return OSM_LAYERS.get(layer_id)


def get_all_osm_layers():
    """Haal alle beschikbare layers op."""
    return OSM_LAYERS.copy()


def get_default_osm_layers():
    """Haal de default layer objecten op."""
    return [OSM_LAYERS[lid] for lid in DEFAULT_LAYERS if lid in OSM_LAYERS]


def get_osm_layers_by_category(category):
    """Haal layers op per categorie."""
    layer_ids = LAYER_CATEGORIES.get(category, [])
    return [OSM_LAYERS[lid] for lid in layer_ids if lid in OSM_LAYERS]
