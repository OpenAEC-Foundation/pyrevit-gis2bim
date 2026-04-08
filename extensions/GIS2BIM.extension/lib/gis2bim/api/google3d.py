# -*- coding: utf-8 -*-
"""
Google 3D Tiles API Client
============================

Download fotogrammetrische 3D meshes via de Google Map Tiles API.
Tiles worden geserveerd als GLB (binary glTF 2.0) bestanden.

Vereist een Google API key met Map Tiles API ingeschakeld.

Geen session token nodig voor 3D tiles — alleen API key.
De session parameter zit al in de child tile URIs uit root.json.

Gebruik:
    from gis2bim.api.google3d import Google3DClient, Google3DError

    client = Google3DClient(api_key="AIza...")
    meshes = client.get_meshes_rd(155000, 463000, bbox_size=200)
"""

import json
import math
import os
import tempfile

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2


class Google3DError(Exception):
    """Fout bij de Google 3D Tiles API."""
    pass


# =============================================================================
# WGS84 Ellipsoid constanten
# =============================================================================

WGS84_A = 6378137.0                # Semi-major axis (m)
WGS84_F = 1.0 / 298.257223563      # Flattening
WGS84_B = WGS84_A * (1.0 - WGS84_F)  # Semi-minor axis
WGS84_E2 = 2.0 * WGS84_F - WGS84_F ** 2  # Eccentricity squared


# =============================================================================
# ECEF <-> WGS84 conversies
# =============================================================================

def wgs84_to_ecef(lat_deg, lon_deg, h=0.0):
    """Converteer WGS84 (lat, lon, hoogte) naar ECEF (x, y, z).

    Args:
        lat_deg: Latitude in graden
        lon_deg: Longitude in graden
        h: Hoogte boven ellipsoide in meters

    Returns:
        (x, y, z) tuple in meters
    """
    lat = math.radians(lat_deg)
    lon = math.radians(lon_deg)

    sin_lat = math.sin(lat)
    cos_lat = math.cos(lat)
    sin_lon = math.sin(lon)
    cos_lon = math.cos(lon)

    n = WGS84_A / math.sqrt(1.0 - WGS84_E2 * sin_lat ** 2)

    x = (n + h) * cos_lat * cos_lon
    y = (n + h) * cos_lat * sin_lon
    z = (n * (1.0 - WGS84_E2) + h) * sin_lat

    return (x, y, z)


def ecef_to_wgs84(x, y, z):
    """Converteer ECEF (x, y, z) naar WGS84 (lat, lon, hoogte).

    Gebruikt iteratieve Bowring methode (convergeert in 2-3 iteraties).

    Args:
        x, y, z: ECEF coordinaten in meters

    Returns:
        (lat_deg, lon_deg, h) tuple, lat/lon in graden, h in meters
    """
    lon = math.atan2(y, x)
    p = math.sqrt(x ** 2 + y ** 2)

    # Initieel: lat schatting (Bowring)
    lat = math.atan2(z, p * (1.0 - WGS84_E2))

    for _ in range(5):
        sin_lat = math.sin(lat)
        n = WGS84_A / math.sqrt(1.0 - WGS84_E2 * sin_lat ** 2)
        lat = math.atan2(z + WGS84_E2 * n * sin_lat, p)

    sin_lat = math.sin(lat)
    cos_lat = math.cos(lat)
    n = WGS84_A / math.sqrt(1.0 - WGS84_E2 * sin_lat ** 2)

    if abs(cos_lat) > 1e-10:
        h = p / cos_lat - n
    else:
        h = abs(z) - WGS84_B

    return (math.degrees(lat), math.degrees(lon), h)


def ecef_distance(p1, p2):
    """Euclidische afstand tussen twee ECEF punten.

    Args:
        p1, p2: (x, y, z) tuples

    Returns:
        Afstand in meters
    """
    return math.sqrt(
        (p1[0] - p2[0]) ** 2 +
        (p1[1] - p2[1]) ** 2 +
        (p1[2] - p2[2]) ** 2
    )


# =============================================================================
# Google 3D Tiles Client
# =============================================================================

TILES_BASE_URL = "https://tile.googleapis.com"
TILES_ROOT_PATH = "/v1/3dtiles/root.json"


class Google3DClient(object):
    """Client voor de Google Map Tiles API (3D Tiles).

    Download fotogrammetrische 3D mesh tiles voor een gegeven locatie.
    Geen session token nodig — de session zit in de child URIs.
    """

    def __init__(self, api_key, timeout=60):
        """
        Args:
            api_key: Google API key met Map Tiles API ingeschakeld.
            timeout: HTTP timeout in seconden.
        """
        self.api_key = api_key
        self.timeout = timeout

    # =========================================================================
    # Tileset traversal
    # =========================================================================

    def get_tiles_for_location(self, lat, lon, radius_m=150,
                               max_geometric_error=20.0,
                               max_tiles=50,
                               progress_callback=None):
        """Zoek en download 3D tiles voor een locatie.

        Traverseert de 3D Tiles tileset boom en selecteert tiles
        die het opgegeven gebied overlappen.

        Args:
            lat: Latitude (WGS84 graden)
            lon: Longitude (WGS84 graden)
            radius_m: Zoekradius in meters
            max_geometric_error: Maximale geometric error
                (lagere waarde = meer detail, meer tiles)
            max_tiles: Maximum aantal tiles om te downloaden
            progress_callback: Optionele callback(str) voor voortgang

        Returns:
            Lijst van paden naar gedownloade GLB bestanden

        Raises:
            Google3DError: Bij API fouten
        """
        center_ecef = wgs84_to_ecef(lat, lon, 0.0)

        # Haal root tileset op
        if progress_callback:
            progress_callback("Root tileset ophalen...")

        root_url = "{0}{1}?key={2}".format(
            TILES_BASE_URL, TILES_ROOT_PATH, self.api_key)
        root_tileset = self._fetch_json(root_url)

        # Traverseer de boom
        if progress_callback:
            progress_callback("Tiles zoeken in zoekgebied...")

        content_uris = []
        self._traverse_tileset(
            root_tileset.get("root", {}),
            center_ecef,
            radius_m,
            max_geometric_error,
            max_tiles,
            content_uris,
        )

        if not content_uris:
            raise Google3DError(
                "Geen 3D tiles gevonden voor deze locatie. "
                "Mogelijk geen Google 3D dekking in dit gebied.")

        # Download tiles
        glb_paths = []
        for i, uri in enumerate(content_uris):
            if progress_callback:
                progress_callback("Tile downloaden {0}/{1}...".format(
                    i + 1, len(content_uris)))

            try:
                path = self._download_glb(uri)
                glb_paths.append(path)
            except Exception:
                pass  # Skip ongeldige tiles

        if not glb_paths:
            raise Google3DError(
                "Kon geen tiles downloaden (netwerk of API fout)")

        return glb_paths

    def _traverse_tileset(self, tile, center_ecef, radius_m,
                          max_error, max_tiles, result):
        """Recursieve tileset traversal.

        Args:
            tile: Tile dict uit tileset JSON
            center_ecef: Zoekcentrum in ECEF
            radius_m: Zoekradius
            max_error: Maximale geometric error threshold
            max_tiles: Maximum resultaten
            result: Lijst waar content URIs aan worden toegevoegd
        """
        if len(result) >= max_tiles:
            return

        # Check bounding volume
        bv = tile.get("boundingVolume", {})
        if not self._intersects_sphere(bv, center_ecef, radius_m):
            return

        geom_error = tile.get("geometricError", float("inf"))
        children = tile.get("children", [])
        content = tile.get("content", {})
        content_uri = content.get("uri") or content.get("url")

        refine = tile.get("refine", "REPLACE").upper()

        if children and geom_error > max_error:
            # Recursief naar kinderen voor meer detail
            if refine == "ADD" and content_uri:
                full_uri = self._resolve_uri(content_uri)
                result.append(full_uri)

            for child in children:
                if len(result) >= max_tiles:
                    break

                child_content = child.get("content", {})
                child_uri = (child_content.get("uri")
                             or child_content.get("url"))

                if child_uri and child_uri.endswith(".json"):
                    # Extern tileset — volg de link
                    try:
                        child_url = self._resolve_uri(child_uri)
                        child_tileset = self._fetch_json(child_url)
                        child_root = child_tileset.get("root", child_tileset)
                        self._traverse_tileset(
                            child_root, center_ecef,
                            radius_m, max_error, max_tiles, result)
                    except Exception:
                        pass
                else:
                    self._traverse_tileset(
                        child, center_ecef,
                        radius_m, max_error, max_tiles, result)
        else:
            # Blad tile of voldoende detail: download content
            if content_uri:
                full_uri = self._resolve_uri(content_uri)
                result.append(full_uri)

    def _intersects_sphere(self, bounding_volume, center_ecef, radius_m):
        """Check of een bounding volume overlapt met een zoeksphere.

        Ondersteunt 'box', 'sphere', en 'region' bounding volumes.

        Args:
            bounding_volume: glTF bounding volume dict
            center_ecef: Zoekcentrum (x, y, z) in ECEF
            radius_m: Zoekradius in meters

        Returns:
            True als er overlap is
        """
        if "box" in bounding_volume:
            box = bounding_volume["box"]
            # box = [cx, cy, cz, xx, xy, xz, yx, yy, yz, zx, zy, zz]
            if len(box) >= 3:
                tile_center = (box[0], box[1], box[2])
                tile_radius = 0.0
                if len(box) >= 12:
                    for i in range(3):
                        axis_len = math.sqrt(
                            box[3 + i * 3] ** 2 +
                            box[4 + i * 3] ** 2 +
                            box[5 + i * 3] ** 2
                        )
                        tile_radius += axis_len
                else:
                    tile_radius = 10000

                dist = ecef_distance(tile_center, center_ecef)
                return dist < (radius_m + tile_radius)

        elif "sphere" in bounding_volume:
            sphere = bounding_volume["sphere"]
            if len(sphere) >= 4:
                tile_center = (sphere[0], sphere[1], sphere[2])
                tile_radius = sphere[3]
                dist = ecef_distance(tile_center, center_ecef)
                return dist < (radius_m + tile_radius)

        elif "region" in bounding_volume:
            region = bounding_volume["region"]
            if len(region) >= 4:
                center_lat = (region[1] + region[3]) / 2.0
                center_lon = (region[0] + region[2]) / 2.0
                tile_ecef = wgs84_to_ecef(
                    math.degrees(center_lat),
                    math.degrees(center_lon), 0.0)
                lat_span = abs(region[3] - region[1])
                lon_span = abs(region[2] - region[0])
                tile_radius = max(lat_span, lon_span) * WGS84_A
                dist = ecef_distance(tile_ecef, center_ecef)
                return dist < (radius_m + tile_radius)

        # Onbekend type: neem aan dat het overlapt (conservatief)
        return True

    # =========================================================================
    # HTTP helpers
    # =========================================================================

    def _resolve_uri(self, uri):
        """Resolve een tile URI naar volledige URL met API key.

        Child URIs uit Google 3D Tiles zijn absolute paden (beginnen met /)
        en bevatten al een ?session= parameter. We moeten alleen het
        base domain prependen en &key= toevoegen.

        Args:
            uri: URI uit tile content (bv. "/v1/3dtiles/datasets/CgA/files/xxx.glb?session=abc")

        Returns:
            Volledige URL met API key
        """
        if uri.startswith("http"):
            full_url = uri
        elif uri.startswith("/"):
            full_url = "{0}{1}".format(TILES_BASE_URL, uri)
        else:
            # Relatief pad — prepend base
            full_url = "{0}/v1/3dtiles/{1}".format(TILES_BASE_URL, uri)

        # Voeg API key toe
        separator = "&" if "?" in full_url else "?"
        full_url = "{0}{1}key={2}".format(full_url, separator, self.api_key)

        return full_url

    def _fetch_json(self, url):
        """Haal JSON data op van URL.

        Args:
            url: Volledige URL

        Returns:
            Geparsed JSON dict

        Raises:
            Google3DError: Bij HTTP fouten
        """
        request = urllib2.Request(url)
        request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

        try:
            response = urllib2.urlopen(request, timeout=self.timeout)
            return json.loads(response.read().decode("utf-8"))
        except urllib2.HTTPError as e:
            error_body = ""
            try:
                error_body = e.read().decode("utf-8")
            except Exception:
                pass
            raise Google3DError(
                "HTTP {0}: {1}".format(e.code, error_body))
        except Exception as e:
            raise Google3DError(
                "Fout bij ophalen data: {0}".format(e))

    def _download_glb(self, url):
        """Download een GLB tile naar een temp bestand.

        Args:
            url: Volledige URL naar GLB content

        Returns:
            Pad naar gedownload temp bestand

        Raises:
            Google3DError: Bij download fouten
        """
        request = urllib2.Request(url)
        request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

        try:
            response = urllib2.urlopen(request, timeout=self.timeout)
            data = response.read()
        except Exception as e:
            raise Google3DError(
                "Fout bij downloaden tile: {0}".format(e))

        if len(data) < 12:
            raise Google3DError("Ongeldig GLB bestand (te klein)")

        fd, tmp_path = tempfile.mkstemp(
            suffix=".glb", prefix="gis2bim_g3d_")
        try:
            os.write(fd, data)
        finally:
            os.close(fd)

        return tmp_path
