# -*- coding: utf-8 -*-
"""
Generieke WFS Client
====================

Een flexibele WFS 2.0 client voor het ophalen van geo-data.
Ondersteunt GeoJSON output en configureerbare layers.

Gebruik:
    from gis2bim.api.wfs import WFSClient, WFSLayer

    layer = WFSLayer(
        name="Perceelgrenzen",
        wfs_url="https://service.pdok.nl/kadaster/kadastralekaart/wfs/v5_0",
        layer_name="kadastralekaartv5:Perceel",
        geometry_type="polygon",
        crs="EPSG:28992"
    )

    client = WFSClient()
    features = client.get_features(layer, bbox=(155000, 463000, 155500, 463500))
"""

import json

try:
    import requests
except ImportError:
    requests = None


class WFSLayer(object):
    """
    Configuratie voor een WFS layer.

    Attributes:
        name: Weergavenaam voor de layer
        wfs_url: WFS service endpoint URL
        layer_name: TypeName in WFS (bijv. "namespace:LayerName")
        geometry_type: Type geometry ("polygon", "point", "line", "multipolygon")
        crs: Coordinaat referentie systeem (bijv. "EPSG:28992")
        label_field: Optioneel veld voor labels
        style_color: Optionele kleurcode voor rendering
        enabled: Of de layer standaard actief is
    """

    def __init__(self, name, wfs_url, layer_name, geometry_type="polygon",
                 crs="EPSG:28992", label_field=None, style_color=None,
                 enabled=True, category="default"):
        self.name = name
        self.wfs_url = wfs_url
        self.layer_name = layer_name
        self.geometry_type = geometry_type
        self.crs = crs
        self.label_field = label_field
        self.style_color = style_color
        self.enabled = enabled
        self.category = category

    def __repr__(self):
        return "WFSLayer('{0}', '{1}')".format(self.name, self.layer_name)


class WFSFeature(object):
    """
    Container voor een WFS feature met geometry en properties.

    Attributes:
        geometry: Lijst van coordinaten (afhankelijk van type)
        geometry_type: Type geometry ("polygon", "point", etc.)
        properties: Dictionary met feature eigenschappen
        label: Optionele label tekst
        label_position: Optionele (x, y) voor label positie
        label_rotation: Optionele rotatie in graden
    """

    def __init__(self, geometry, geometry_type, properties=None):
        self.geometry = geometry
        self.geometry_type = geometry_type
        self.properties = properties or {}
        self.label = None
        self.label_position = None
        self.label_rotation = 0.0

    def __repr__(self):
        return "WFSFeature({0}, props={1})".format(
            self.geometry_type,
            len(self.properties)
        )


class WFSClient(object):
    """
    Generieke WFS 2.0 client voor GeoJSON output.

    Ondersteunt:
    - GetFeature requests met BBOX filter
    - GeoJSON output parsing
    - Polygon, MultiPolygon, Point en LineString geometries
    """

    def __init__(self, timeout=30):
        if requests is None:
            raise ImportError("requests library required. pip install requests")
        self.timeout = timeout

    def get_features(self, layer, bbox, max_features=10000):
        """
        Haal features op voor een layer binnen een bounding box.

        Args:
            layer: WFSLayer configuratie
            bbox: Tuple (xmin, ymin, xmax, ymax) in layer CRS
            max_features: Maximum aantal features (default 10000)

        Returns:
            Lijst van WFSFeature objecten
        """
        url = self._build_url(layer, bbox, max_features)

        try:
            response = requests.get(url, timeout=self.timeout)
            response.raise_for_status()

            data = response.json()
            return self._parse_geojson(data, layer)

        except requests.exceptions.RequestException as e:
            print("WFS request error for {0}: {1}".format(layer.name, e))
            return []
        except ValueError as e:
            print("WFS JSON parse error for {0}: {1}".format(layer.name, e))
            return []

    def get_capabilities(self, wfs_url):
        """
        Haal beschikbare layers op van een WFS service.

        Args:
            wfs_url: WFS service endpoint URL

        Returns:
            Lijst van layer namen (TypeNames)
        """
        url = "{0}?service=WFS&version=2.0.0&request=GetCapabilities".format(wfs_url)

        try:
            response = requests.get(url, timeout=self.timeout)
            response.raise_for_status()

            # Parse XML - eenvoudige string search voor FeatureType names
            content = response.text
            layers = []

            # Zoek naar <Name> tags binnen <FeatureType> blokken
            import re
            pattern = r"<(?:wfs:)?FeatureType[^>]*>.*?<(?:wfs:)?Name>([^<]+)</(?:wfs:)?Name>"
            matches = re.findall(pattern, content, re.DOTALL)

            for match in matches:
                if match not in layers:
                    layers.append(match)

            return layers

        except Exception as e:
            print("GetCapabilities error: {0}".format(e))
            return []

    def _build_url(self, layer, bbox, max_features):
        """Bouw WFS GetFeature URL."""
        # BBOX format: xmin,ymin,xmax,ymax,CRS
        return (
            "{base}?service=WFS&version=2.0.0&request=GetFeature"
            "&typeName={layer}"
            "&bbox={xmin},{ymin},{xmax},{ymax},{crs}"
            "&count={count}"
            "&outputFormat=application/json"
        ).format(
            base=layer.wfs_url,
            layer=layer.layer_name,
            xmin=bbox[0],
            ymin=bbox[1],
            xmax=bbox[2],
            ymax=bbox[3],
            crs=layer.crs,
            count=max_features
        )

    def _parse_geojson(self, data, layer):
        """Parse GeoJSON response naar WFSFeature objecten."""
        features = []

        for gj_feature in data.get("features", []):
            props = gj_feature.get("properties", {})
            geom = gj_feature.get("geometry", {})

            if not geom:
                continue

            geom_type = geom.get("type", "").lower()
            coords = geom.get("coordinates", [])

            # Parse geometry
            parsed_geom = self._parse_geometry(geom_type, coords)
            if parsed_geom is None:
                continue

            # Maak feature
            feature = WFSFeature(
                geometry=parsed_geom,
                geometry_type=self._normalize_geometry_type(geom_type),
                properties=props
            )

            # Voeg label info toe indien geconfigureerd
            if layer.label_field:
                feature.label = str(props.get(layer.label_field, ""))

            # Check voor specifieke label positie velden (PDOK specifiek)
            self._extract_label_position(feature, props, layer)

            features.append(feature)

        return features

    def _parse_geometry(self, geom_type, coords):
        """Parse GeoJSON coordinates naar simpele tuple lijsten."""
        if geom_type == "point":
            if len(coords) >= 2:
                return (coords[0], coords[1])
            return None

        elif geom_type == "linestring":
            return [(c[0], c[1]) for c in coords if len(c) >= 2]

        elif geom_type == "polygon":
            # Eerste ring is outer boundary
            if coords and coords[0]:
                return [(c[0], c[1]) for c in coords[0] if len(c) >= 2]
            return None

        elif geom_type == "multipolygon":
            # Return alle polygonen als lijst van rings
            result = []
            for polygon in coords:
                if polygon and polygon[0]:
                    ring = [(c[0], c[1]) for c in polygon[0] if len(c) >= 2]
                    if ring:
                        result.append(ring)
            return result if result else None

        elif geom_type == "multilinestring":
            result = []
            for line in coords:
                parsed_line = [(c[0], c[1]) for c in line if len(c) >= 2]
                if parsed_line:
                    result.append(parsed_line)
            return result if result else None

        elif geom_type == "multipoint":
            return [(c[0], c[1]) for c in coords if len(c) >= 2]

        return None

    def _normalize_geometry_type(self, geom_type):
        """Normaliseer geometry type naar standaard namen."""
        type_map = {
            "point": "point",
            "multipoint": "multipoint",
            "linestring": "line",
            "multilinestring": "multiline",
            "polygon": "polygon",
            "multipolygon": "multipolygon"
        }
        return type_map.get(geom_type.lower(), geom_type)

    def _extract_label_position(self, feature, props, layer):
        """Extraheer label positie uit PDOK-specifieke velden."""
        # PDOK Kadaster: perceelnummer positie
        if "perceelnummerPlaatscoordinaatX" in props:
            x = props.get("perceelnummerPlaatscoordinaatX")
            y = props.get("perceelnummerPlaatscoordinaatY")
            if x is not None and y is not None:
                feature.label_position = (float(x), float(y))
                feature.label_rotation = float(props.get("perceelnummerRotatie", 0))
                if not feature.label:
                    feature.label = str(props.get("perceelnummer", ""))

        # PDOK BAG: nummeraanduiding positie
        elif "plaatscoordinaat" in str(props.keys()).lower():
            # Variaties in BAG veldnamen
            for key in props:
                if "plaatscoordinaatx" in key.lower():
                    x_key = key
                    y_key = key.replace("X", "Y").replace("x", "y")
                    if y_key in props:
                        feature.label_position = (
                            float(props.get(x_key, 0)),
                            float(props.get(y_key, 0))
                        )
                        break

        # PDOK: hoek veld voor rotatie (straatnamen etc.)
        if "hoek" in props:
            feature.label_rotation = float(props.get("hoek", 0))


def get_wfs_data(layers, bbox, parallel=False):
    """
    Utility functie om data van meerdere layers op te halen.

    Args:
        layers: Lijst van WFSLayer objecten
        bbox: Bounding box tuple (xmin, ymin, xmax, ymax)
        parallel: Of parallel ophalen (nog niet geimplementeerd)

    Returns:
        Dictionary met layer.name -> lijst van features
    """
    client = WFSClient()
    results = {}

    for layer in layers:
        if layer.enabled:
            features = client.get_features(layer, bbox)
            results[layer.name] = features
            print("WFS: {0} - {1} features opgehaald".format(
                layer.name, len(features)
            ))

    return results
