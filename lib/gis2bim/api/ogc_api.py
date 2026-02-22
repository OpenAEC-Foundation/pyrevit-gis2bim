# -*- coding: utf-8 -*-
"""
OGC API Features Client
=======================

Client voor OGC API Features (de opvolger van WFS).
Specifiek voor PDOK BGT en andere Nederlandse geodata.

Gebruik:
    from gis2bim.api.ogc_api import OGCAPIClient, OGCAPICollection

    client = OGCAPIClient("https://api.pdok.nl/lv/bgt/ogc/v1")
    features = client.get_features("wegdeel", bbox=(155000, 463000, 155500, 463500))
"""

import json

try:
    import urllib2
    from urllib import urlencode
except ImportError:
    # Python 3
    import urllib.request as urllib2
    from urllib.parse import urlencode


class OGCAPICollection(object):
    """
    Configuratie voor een OGC API Features collection.

    Attributes:
        collection_id: ID van de collection in de API
        name: Weergavenaam voor UI
        geometry_type: Type geometry ("polygon", "line", "point")
        category: Categorie voor groepering
        filled_region_type: Naam van FilledRegionType in Revit (voor vlakken)
        line_style: Naam van line style in Revit (voor lijnen)
        boundary_line_style: Naam van lijnstijl voor filled region randen
        enabled: Of de layer standaard actief is
    """

    def __init__(self, collection_id, name, geometry_type="polygon",
                 category="bgt", filled_region_type=None, line_style=None,
                 boundary_line_style=None, enabled=True):
        self.collection_id = collection_id
        self.name = name
        self.geometry_type = geometry_type
        self.category = category
        self.filled_region_type = filled_region_type
        self.line_style = line_style
        self.boundary_line_style = boundary_line_style
        self.enabled = enabled

    def __repr__(self):
        return "OGCAPICollection('{0}', '{1}')".format(
            self.collection_id, self.name
        )


class OGCAPIFeature(object):
    """
    Container voor een OGC API feature met geometry en properties.

    Attributes:
        geometry: Lijst van coordinaten
        geometry_type: Type geometry ("polygon", "point", etc.)
        properties: Dictionary met feature eigenschappen
    """

    def __init__(self, geometry, geometry_type, properties=None):
        self.geometry = geometry
        self.geometry_type = geometry_type
        self.properties = properties or {}

    def __repr__(self):
        return "OGCAPIFeature({0}, props={1})".format(
            self.geometry_type,
            len(self.properties)
        )


class OGCAPIClient(object):
    """
    Client voor OGC API Features services.

    Ondersteunt:
    - GetItems requests met BBOX filter
    - GeoJSON output parsing
    - Automatische paginatie
    - Polygon, MultiPolygon, Point en LineString geometries
    """

    def __init__(self, base_url, timeout=30):
        """
        Initialiseer de client.

        Args:
            base_url: Basis URL van de OGC API (bijv. https://api.pdok.nl/lv/bgt/ogc/v1)
            timeout: Timeout in seconden
        """
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout

    def get_features(self, collection_id, bbox, limit=1000, max_features=10000, crs="EPSG:28992"):
        """
        Haal features op voor een collection binnen een bounding box.

        Args:
            collection_id: ID van de collection (bijv. "wegdeel")
            bbox: Tuple (xmin, ymin, xmax, ymax) in RD coordinaten
            limit: Features per request (max 1000 voor PDOK)
            max_features: Maximum totaal aantal features
            crs: Coordinaat referentie systeem

        Returns:
            Lijst van OGCAPIFeature objecten
        """
        all_features = []
        offset = 0

        while len(all_features) < max_features:
            url = self._build_url(collection_id, bbox, limit, offset, crs)

            try:
                response = self._make_request(url)
                data = json.loads(response)

                features = self._parse_geojson(data)
                all_features.extend(features)

                # Check of er meer features zijn
                number_returned = data.get("numberReturned", 0)
                if number_returned < limit:
                    # Geen volgende pagina
                    break

                offset += limit

            except Exception as e:
                print("OGC API request error for {0}: {1}".format(collection_id, e))
                break

        return all_features[:max_features]

    def get_collections(self):
        """
        Haal beschikbare collections op van de API.

        Returns:
            Lijst van collection IDs
        """
        url = "{0}/collections?f=json".format(self.base_url)

        try:
            response = self._make_request(url)
            data = json.loads(response)

            collections = []
            for coll in data.get("collections", []):
                collections.append(coll.get("id", ""))

            return collections

        except Exception as e:
            print("Get collections error: {0}".format(e))
            return []

    def _build_url(self, collection_id, bbox, limit, offset, crs):
        """Bouw OGC API Features URL."""
        # CRS URI voor RD (EPSG:28992)
        crs_code = crs.replace("EPSG:", "")
        crs_uri = "http://www.opengis.net/def/crs/EPSG/0/{0}".format(crs_code)

        # Bouw URL handmatig om encoding problemen te voorkomen
        # (urlencode encodet de slashes in de CRS URI verkeerd)
        url = (
            "{base}/collections/{collection}/items"
            "?f=json"
            "&bbox={xmin},{ymin},{xmax},{ymax}"
            "&bbox-crs={crs}"
            "&crs={crs}"
            "&limit={limit}"
        ).format(
            base=self.base_url,
            collection=collection_id,
            xmin=bbox[0],
            ymin=bbox[1],
            xmax=bbox[2],
            ymax=bbox[3],
            crs=crs_uri,
            limit=limit
        )

        if offset > 0:
            url += "&offset={0}".format(offset)

        return url

    def _make_request(self, url):
        """Maak HTTP request en return response text."""
        request = urllib2.Request(url)
        request.add_header("Accept", "application/json")
        request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

        response = urllib2.urlopen(request, timeout=self.timeout)
        return response.read().decode("utf-8")

    def _parse_geojson(self, data):
        """Parse GeoJSON response naar OGCAPIFeature objecten."""
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

            feature = OGCAPIFeature(
                geometry=parsed_geom,
                geometry_type=self._normalize_geometry_type(geom_type),
                properties=props
            )

            features.append(feature)

        return features

    def _parse_geometry(self, geom_type, coords):
        """Parse GeoJSON coordinates naar simpele tuple lijsten.

        Voor polygonen worden ALLE rings teruggegeven (outer + holes):
        - polygon: [[outer_ring], [hole1], [hole2], ...]
        - multipolygon: [[[outer1], [hole1a]], [[outer2], [hole2a]], ...]
        """
        if geom_type == "point":
            if len(coords) >= 2:
                return (coords[0], coords[1])
            return None

        elif geom_type == "linestring":
            return [(c[0], c[1]) for c in coords if len(c) >= 2]

        elif geom_type == "polygon":
            # Return ALLE rings: outer + holes
            # coords = [outer_ring, hole1, hole2, ...]
            rings = []
            for ring_coords in coords:
                if ring_coords:
                    ring = [(c[0], c[1]) for c in ring_coords if len(c) >= 2]
                    if ring:
                        rings.append(ring)
            return rings if rings else None

        elif geom_type == "multipolygon":
            # Return alle polygonen, elk met hun rings
            # coords = [polygon1, polygon2, ...]
            # polygon = [outer_ring, hole1, hole2, ...]
            result = []
            for polygon_coords in coords:
                rings = []
                for ring_coords in polygon_coords:
                    if ring_coords:
                        ring = [(c[0], c[1]) for c in ring_coords if len(c) >= 2]
                        if ring:
                            rings.append(ring)
                if rings:
                    result.append(rings)
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
