# -*- coding: utf-8 -*-
"""
Overpass API Client
===================

Client voor de OpenStreetMap Overpass API.
Haalt OSM data op via Overpass QL queries en parseert
nodes, ways en relations naar bruikbare geometrie.

API: https://overpass-api.de/api/interpreter

Gebruik:
    from gis2bim.api.overpass import OverpassClient

    client = OverpassClient()
    features = client.get_features(
        bbox_wgs84=(51.9, 4.4, 52.0, 4.5),
        query_tags=['way["building"]', 'relation["building"]']
    )
"""

import json

try:
    import urllib2
    from urllib import urlencode
except ImportError:
    import urllib.request as urllib2
    from urllib.parse import urlencode


class OverpassFeature(object):
    """
    Container voor een OSM feature met geometry en tags.

    Attributes:
        geometry: Coordinaten - afhankelijk van geometry_type:
            - polygon: lijst van ringen, elke ring is [(lat, lon), ...]
            - line: [(lat, lon), ...]
            - point: (lat, lon)
        geometry_type: "polygon", "multipolygon", "line" of "point"
        tags: Dictionary met OSM tags
        osm_id: OSM element ID
    """

    def __init__(self, geometry, geometry_type, tags=None, osm_id=None):
        self.geometry = geometry
        self.geometry_type = geometry_type
        self.tags = tags or {}
        self.osm_id = osm_id

    def __repr__(self):
        return "OverpassFeature({0}, id={1}, tags={2})".format(
            self.geometry_type, self.osm_id, len(self.tags)
        )


class OverpassClient(object):
    """
    Client voor de OpenStreetMap Overpass API.

    Ondersteunt:
    - Overpass QL queries met bounding box
    - JSON output parsing
    - Node, Way en Relation geometrie
    - Polygon en line geometrie met holes (multipolygon relations)
    """

    OVERPASS_URL = "https://overpass-api.de/api/interpreter"

    def __init__(self, timeout=30, log_func=None):
        """
        Initialiseer de client.

        Args:
            timeout: Query timeout in seconden (Overpass server-side)
            log_func: Optionele log functie voor foutmeldingen
        """
        self.timeout = timeout
        self._log = log_func or (lambda msg: None)

    def get_features(self, bbox_wgs84, query_tags, as_polygon=True):
        """
        Haal features op uit OpenStreetMap via Overpass API.

        Gebruikt 'out geom;' voor directe geometrie in de response,
        wat veel sneller is dan 'out body; >; out skel qt;' omdat
        geen aparte node-resolutie nodig is.

        Args:
            bbox_wgs84: Tuple (south, west, north, east) in WGS84 graden
            query_tags: Lijst van Overpass QL fragments,
                        bijv. ['way["building"]', 'relation["building"]']
            as_polygon: Of gesloten ways als polygonen behandeld moeten worden

        Returns:
            Lijst van OverpassFeature objecten
        """
        south, west, north, east = bbox_wgs84
        bbox_str = "{0},{1},{2},{3}".format(south, west, north, east)

        # Bouw Overpass QL query
        fragments = ""
        for tag_query in query_tags:
            fragments += "  {0}({1});\n".format(tag_query, bbox_str)

        # out geom: geometrie direct in response (geen aparte node lookup)
        query = (
            "[out:json][timeout:{timeout}];\n"
            "(\n"
            "{fragments}"
            ");\n"
            "out geom;"
        ).format(
            timeout=self.timeout,
            fragments=fragments
        )

        data = self._execute_query(query)
        if not data:
            return []

        return self._parse_response_geom(data, as_polygon)

    def get_pois(self, bbox_wgs84, query_tags):
        """
        Haal POI punten op (nodes + way-centers) via Overpass API.

        Gebruikt 'out center;' zodat ways een center-coordinaat krijgen.
        Parseert zowel node als way elementen.

        Args:
            bbox_wgs84: Tuple (south, west, north, east) in WGS84 graden
            query_tags: Lijst van Overpass QL fragments,
                        bijv. ['node["amenity"="school"]', 'way["amenity"="school"]']

        Returns:
            Lijst van dicts: [{"lat": float, "lon": float, "tags": dict, "osm_id": int}, ...]
        """
        south, west, north, east = bbox_wgs84
        bbox_str = "{0},{1},{2},{3}".format(south, west, north, east)

        # Bouw Overpass QL query
        fragments = ""
        for tag_query in query_tags:
            fragments += "  {0}({1});\n".format(tag_query, bbox_str)

        # out center: ways krijgen een center coordinaat
        query = (
            "[out:json][timeout:{timeout}];\n"
            "(\n"
            "{fragments}"
            ");\n"
            "out center;"
        ).format(
            timeout=self.timeout,
            fragments=fragments
        )

        data = self._execute_query(query)
        if not data:
            return []

        return self._parse_response_pois(data)

    def _parse_response_pois(self, data):
        """
        Parse Overpass JSON response met 'out center' naar POI punten.

        Nodes: direct lat/lon.
        Ways: center veld (lat/lon).
        """
        elements = data.get("elements", [])
        pois = []

        for elem in elements:
            etype = elem.get("type")
            eid = elem.get("id")
            tags = elem.get("tags", {})

            lat = None
            lon = None

            if etype == "node":
                lat = elem.get("lat")
                lon = elem.get("lon")
            elif etype == "way":
                center = elem.get("center")
                if center:
                    lat = center.get("lat")
                    lon = center.get("lon")
            elif etype == "relation":
                center = elem.get("center")
                if center:
                    lat = center.get("lat")
                    lon = center.get("lon")

            if lat is not None and lon is not None:
                pois.append({
                    "lat": lat,
                    "lon": lon,
                    "tags": tags,
                    "osm_id": eid,
                })

        self._log("POIs parsed: {0} elements -> {1} points".format(
            len(elements), len(pois)))
        return pois

    def _execute_query(self, query):
        """Voer een Overpass QL query uit via HTTP POST."""
        try:
            self._log("Overpass query ({0} tekens)".format(len(query)))

            post_data = urlencode({"data": query})
            # IronPython 2.7: encode to bytes
            if hasattr(post_data, 'encode'):
                post_data = post_data.encode("utf-8")

            request = urllib2.Request(self.OVERPASS_URL, data=post_data)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            request.add_header("Content-Type",
                               "application/x-www-form-urlencoded")

            response = urllib2.urlopen(request, timeout=self.timeout + 15)
            text = response.read().decode("utf-8")
            self._log("Response ontvangen: {0} bytes".format(len(text)))
            return json.loads(text)

        except Exception as e:
            self._log("Overpass API error: {0}".format(e))
            return None

    def _parse_response_geom(self, data, as_polygon):
        """
        Parse Overpass JSON response met 'out geom' naar OverpassFeature objecten.

        Met 'out geom' bevat elk element direct zijn geometry:
        - Ways: 'geometry' array met {lat, lon} objecten
        - Relations: 'members' met per member een 'geometry' array
        """
        elements = data.get("elements", [])
        features = []

        for elem in elements:
            etype = elem.get("type")
            eid = elem.get("id")
            tags = elem.get("tags", {})

            if not tags:
                continue

            if etype == "way":
                # Way geometry zit direct in 'geometry' array
                geom_data = elem.get("geometry", [])
                coords = [(pt["lat"], pt["lon"]) for pt in geom_data
                          if "lat" in pt and "lon" in pt]

                if len(coords) < 2:
                    continue

                # Bepaal of het een gesloten polygon is
                is_closed = (len(coords) >= 4 and
                             coords[0][0] == coords[-1][0] and
                             coords[0][1] == coords[-1][1])

                if as_polygon and is_closed:
                    features.append(OverpassFeature(
                        geometry=[coords],
                        geometry_type="polygon",
                        tags=tags,
                        osm_id=eid
                    ))
                else:
                    features.append(OverpassFeature(
                        geometry=coords,
                        geometry_type="line",
                        tags=tags,
                        osm_id=eid
                    ))

            elif etype == "relation":
                rel_type = tags.get("type", "")
                if rel_type != "multipolygon":
                    continue

                members = elem.get("members", [])
                outer_rings = []
                inner_rings = []

                for member in members:
                    if member.get("type") != "way":
                        continue

                    role = member.get("role", "outer")
                    geom_data = member.get("geometry", [])
                    coords = [(pt["lat"], pt["lon"]) for pt in geom_data
                              if "lat" in pt and "lon" in pt]

                    if len(coords) < 3:
                        continue

                    if role == "inner":
                        inner_rings.append(coords)
                    else:
                        outer_rings.append(coords)

                if not outer_rings:
                    continue

                if len(outer_rings) == 1:
                    rings = [outer_rings[0]] + inner_rings
                    features.append(OverpassFeature(
                        geometry=rings,
                        geometry_type="polygon",
                        tags=tags,
                        osm_id=eid
                    ))
                else:
                    for outer in outer_rings:
                        features.append(OverpassFeature(
                            geometry=[outer] + inner_rings,
                            geometry_type="polygon",
                            tags=tags,
                            osm_id=eid
                        ))

        self._log("Parsed: {0} elements -> {1} features".format(
            len(elements), len(features)))
        return features
