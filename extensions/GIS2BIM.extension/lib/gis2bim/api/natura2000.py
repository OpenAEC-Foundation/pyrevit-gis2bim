# -*- coding: utf-8 -*-
"""
Natura 2000 WFS Client
=======================

Client voor het ophalen van Natura 2000 gebieden via de PDOK WFS service
en het berekenen van de minimale afstand tot het dichtstbijzijnde gebied.

Gebruik:
    from gis2bim.api.natura2000 import Natura2000Client

    client = Natura2000Client()
    result = client.get_natura2000_info((155000, 463000), 15000)
    print(result.nearest_name)      # "Veluwe"
    print(result.min_distance)      # 3245.7
    print(result.area_names)        # ["Veluwe", "Rijntakken"]
"""

import json
import math

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2


# WFS endpoint PDOK Natura 2000
NATURA2000_WFS_URL = "https://service.pdok.nl/rvo/natura2000/wfs/v1_0"
NATURA2000_LAYER = "natura2000"


class Natura2000Area(object):
    """
    Container voor een Natura 2000 gebied.

    Attributes:
        name: Naam van het gebied (bijv. "Veluwe")
        nr: Gebiedsnummer (bijv. "57")
        bescherming: Type bescherming (bijv. "Habitatrichtlijn")
        polygons: Lijst van polygon rings, elke ring is een lijst
                  van (x, y) tuples in RD coordinaten
    """

    def __init__(self, name, nr="", bescherming="", polygons=None):
        self.name = name
        self.nr = nr
        self.bescherming = bescherming
        self.polygons = polygons or []

    def __repr__(self):
        return "Natura2000Area('{0}', nr='{1}')".format(self.name, self.nr)


class Natura2000Result(object):
    """
    Aggregaat resultaat van een Natura 2000 query.

    Attributes:
        areas: Lijst van Natura2000Area objecten
        min_distance: Minimale afstand in meters tot dichtstbijzijnd gebied
        nearest_name: Naam van het dichtstbijzijnde gebied
        area_names: Lijst van namen van gevonden gebieden (uniek)
    """

    def __init__(self, areas=None, min_distance=None, nearest_name="",
                 area_names=None):
        self.areas = areas or []
        self.min_distance = min_distance
        self.nearest_name = nearest_name
        self.area_names = area_names or []

    def __repr__(self):
        return "Natura2000Result(areas={0}, min_distance={1}, nearest='{2}')".format(
            len(self.areas), self.min_distance, self.nearest_name
        )


class Natura2000Client(object):
    """
    WFS client voor PDOK Natura 2000 gebieden.

    Haalt Natura 2000 gebieden op via de WFS service van PDOK
    en berekent de minimale afstand tot het dichtstbijzijnde gebied.
    """

    def __init__(self, timeout=60):
        self.timeout = timeout
        self.wfs_url = NATURA2000_WFS_URL
        self.layer = NATURA2000_LAYER

    def get_areas(self, bbox):
        """
        Haal Natura 2000 gebieden op binnen een bounding box.

        Args:
            bbox: Tuple (xmin, ymin, xmax, ymax) in RD coordinaten

        Returns:
            Lijst van Natura2000Area objecten
        """
        url = self._build_url(bbox)

        try:
            request = urllib2.Request(url)
            request.add_header("Accept", "application/json")
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

            response = urllib2.urlopen(request, timeout=self.timeout)
            data = json.loads(response.read().decode("utf-8"))

            return self._parse_response(data)

        except Exception as e:
            print("Natura 2000 WFS request error: {0}".format(e))
            return []

    def calculate_distance(self, point, areas):
        """
        Bereken minimale afstand van een punt tot Natura 2000 gebieden.

        Gebruikt punt-in-polygon test (ray casting) en punt-naar-segment
        afstand berekening. Alle berekeningen in RD meters (Euclidisch).

        Args:
            point: Tuple (x, y) in RD coordinaten
            areas: Lijst van Natura2000Area objecten

        Returns:
            Tuple (min_distance, nearest_area_name)
            min_distance is 0.0 als het punt binnen een gebied ligt
        """
        px, py = point
        min_dist = float("inf")
        nearest_name = ""

        for area in areas:
            for ring in area.polygons:
                if len(ring) < 3:
                    continue

                # Punt-in-polygon test (ray casting)
                if self._point_in_polygon(px, py, ring):
                    return (0.0, area.name)

                # Punt-naar-ring afstand
                dist = self._point_to_ring_distance(px, py, ring)
                if dist < min_dist:
                    min_dist = dist
                    nearest_name = area.name

        if min_dist == float("inf"):
            return (None, "")

        return (min_dist, nearest_name)

    def get_natura2000_info(self, center_rd, search_radius=15000):
        """
        Convenience methode: haal gebieden op en bereken afstand.

        Args:
            center_rd: Tuple (x, y) in RD coordinaten
            search_radius: Zoekstraal in meters (default 15000 = 15km)

        Returns:
            Natura2000Result object
        """
        cx, cy = center_rd

        # Bbox is search_radius * 2 vierkant rondom center
        bbox_size = search_radius * 2
        bbox = (
            cx - search_radius,
            cy - search_radius,
            cx + search_radius,
            cy + search_radius
        )

        areas = self.get_areas(bbox)

        if not areas:
            return Natura2000Result(
                areas=[],
                min_distance=None,
                nearest_name="",
                area_names=[]
            )

        # Unieke namen
        seen = set()
        area_names = []
        for a in areas:
            if a.name not in seen:
                seen.add(a.name)
                area_names.append(a.name)

        # Afstand berekenen
        min_distance, nearest_name = self.calculate_distance(center_rd, areas)

        return Natura2000Result(
            areas=areas,
            min_distance=min_distance,
            nearest_name=nearest_name,
            area_names=area_names
        )

    # ---- Private methods ----

    def _build_url(self, bbox):
        """Bouw WFS GetFeature URL voor Natura 2000."""
        return (
            "{base}?service=WFS&version=2.0.0&request=GetFeature"
            "&typeName={layer}"
            "&bbox={xmin},{ymin},{xmax},{ymax},EPSG:28992"
            "&count=5000"
            "&outputFormat=application/json"
        ).format(
            base=self.wfs_url,
            layer=self.layer,
            xmin=bbox[0],
            ymin=bbox[1],
            xmax=bbox[2],
            ymax=bbox[3]
        )

    def _parse_response(self, data):
        """Parse GeoJSON response naar Natura2000Area objecten."""
        areas = []

        for feature in data.get("features", []):
            props = feature.get("properties", {})
            geom = feature.get("geometry", {})

            if not geom:
                continue

            name = props.get("naamN2K", props.get("naam", "Onbekend"))
            nr = str(props.get("nr", ""))
            bescherming = props.get("beschermin", "")

            polygons = self._extract_polygons(geom)

            if polygons:
                area = Natura2000Area(
                    name=name,
                    nr=nr,
                    bescherming=bescherming,
                    polygons=polygons
                )
                areas.append(area)

        return areas

    def _extract_polygons(self, geom):
        """Extraheer polygon rings uit GeoJSON geometry.

        Ondersteunt Polygon en MultiPolygon geometrieen.

        Returns:
            Lijst van rings (lijst van (x, y) tuples)
        """
        geom_type = geom.get("type", "")
        coords = geom.get("coordinates", [])
        rings = []

        if geom_type == "Polygon":
            # coords = [outer_ring, hole1, hole2, ...]
            # We gebruiken alleen de outer ring voor afstandsberekening
            for ring_coords in coords:
                ring = [(c[0], c[1]) for c in ring_coords]
                if ring:
                    rings.append(ring)

        elif geom_type == "MultiPolygon":
            # coords = [polygon1, polygon2, ...]
            for polygon_coords in coords:
                for ring_coords in polygon_coords:
                    ring = [(c[0], c[1]) for c in ring_coords]
                    if ring:
                        rings.append(ring)

        return rings

    def _point_in_polygon(self, px, py, ring):
        """Ray casting punt-in-polygon test.

        Args:
            px, py: Punt coordinaten
            ring: Lijst van (x, y) tuples

        Returns:
            True als punt binnen de polygon ligt
        """
        n = len(ring)
        inside = False

        j = n - 1
        for i in range(n):
            xi, yi = ring[i]
            xj, yj = ring[j]

            if ((yi > py) != (yj > py)) and \
               (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
                inside = not inside

            j = i

        return inside

    def _point_to_segment_distance(self, px, py, x1, y1, x2, y2):
        """Bereken minimale afstand van punt naar lijnsegment.

        Args:
            px, py: Punt coordinaten
            x1, y1, x2, y2: Segment eindpunten

        Returns:
            Afstand in meters (RD eenheden)
        """
        dx = x2 - x1
        dy = y2 - y1
        length_sq = dx * dx + dy * dy

        if length_sq == 0:
            # Segment is een punt
            return math.sqrt((px - x1) ** 2 + (py - y1) ** 2)

        # Projectie van punt op de lijn (parameter t)
        t = ((px - x1) * dx + (py - y1) * dy) / length_sq
        t = max(0, min(1, t))

        # Dichtstbijzijnd punt op het segment
        nearest_x = x1 + t * dx
        nearest_y = y1 + t * dy

        return math.sqrt((px - nearest_x) ** 2 + (py - nearest_y) ** 2)

    def _point_to_ring_distance(self, px, py, ring):
        """Bereken minimale afstand van punt naar polygon ring.

        Args:
            px, py: Punt coordinaten
            ring: Lijst van (x, y) tuples

        Returns:
            Minimale afstand in meters
        """
        min_dist = float("inf")
        n = len(ring)

        for i in range(n):
            x1, y1 = ring[i]
            x2, y2 = ring[(i + 1) % n]

            dist = self._point_to_segment_distance(px, py, x1, y1, x2, y2)
            if dist < min_dist:
                min_dist = dist

        return min_dist
