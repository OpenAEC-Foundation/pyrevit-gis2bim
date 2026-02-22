# -*- coding: utf-8 -*-
"""
PDOK 3D Basisvoorziening Client
================================

Client voor het downloaden van CityJSON data van de PDOK 3D Basisvoorziening.

Workflow:
1. Bereken tile ID op basis van RD coordinaten (2x2km grid)
2. Query PDOK OGC API voor download link
3. Download CityJSON ZIP bestand
4. Extract .json/.city.json uit ZIP

Gebruik:
    from gis2bim.api.pdok3d import PDOK3DClient, PDOK3DError

    client = PDOK3DClient()
    url = client.get_tile_url(155000, 463000)
    filepath = client.download_cityjson(url)
"""

import json
import math
import os
import tempfile
import zipfile

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2

from gis2bim.coordinates import rd_to_wgs84


class PDOK3DError(Exception):
    """Fout bij het downloaden of verwerken van PDOK 3D data."""
    pass


class PDOK3DClient(object):
    """Client voor PDOK 3D Basisvoorziening tile download.

    Tiles zijn 2x2km in RD coordinaten.

    Attributes:
        OGC_BASE_URL: Basis URL van de PDOK OGC API
        COLLECTIONS: Beschikbare data collecties
    """

    OGC_BASE_URL = (
        "https://api.pdok.nl/kadaster/3d-basisvoorziening/ogc/v1_0"
    )

    COLLECTIONS = {
        "gebouwen": "basisbestand_gebouwen",
        "gebouwen_terreinen": "basisbestand_gebouwen_terreinen",
    }

    # Direct download URL patroon
    DOWNLOAD_BASE_URL = (
        "https://download.pdok.nl/kadaster/basisvoorziening-3d/v1_0"
    )

    def __init__(self, timeout=120):
        """Initialiseer de client.

        Args:
            timeout: Download timeout in seconden
        """
        self.timeout = timeout

    def get_tile_id(self, rd_x, rd_y):
        """Bereken tile ID voor RD coordinaten.

        Tiles zijn 2x2km. Tile ID = "{tile_x}_{tile_y}".

        Args:
            rd_x: RD X coordinaat in meters
            rd_y: RD Y coordinaat in meters

        Returns:
            Tile ID string (bijv. "106000_434000")
        """
        tile_x = int(math.floor(rd_x / 2000.0)) * 2000
        tile_y = int(math.floor(rd_y / 2000.0)) * 2000
        return "{0}_{1}".format(tile_x, tile_y)

    def get_tile_url(self, rd_x, rd_y, collection="gebouwen_terreinen",
                     year=2024):
        """Bereken tile ID en haal download URL op via PDOK API.

        Probeert eerst de directe download URL.
        Fallback: query OGC API voor download link.

        Args:
            rd_x: RD X coordinaat in meters
            rd_y: RD Y coordinaat in meters
            collection: "gebouwen" of "gebouwen_terreinen"
            year: Jaar van de dataset (default 2024)

        Returns:
            Download URL voor CityJSON ZIP

        Raises:
            PDOK3DError: Bij fouten
        """
        tile_id = self.get_tile_id(rd_x, rd_y)
        col_name = self.COLLECTIONS.get(collection, collection)

        # Probeer directe download URL
        direct_url = self._build_direct_url(tile_id, col_name, year)

        # Verifieer dat URL bereikbaar is met HEAD request
        if self._check_url_exists(direct_url):
            return direct_url

        # Fallback: query OGC API
        return self._query_ogc_api(rd_x, rd_y, col_name)

    def get_neighboring_tile_urls(self, rd_x, rd_y,
                                  collection="gebouwen_terreinen",
                                  year=2024):
        """Haal URLs op voor de tile + directe buren (3x3 grid).

        Nuttig wanneer het punt dichtbij een tile-grens ligt.

        Args:
            rd_x: RD X coordinaat
            rd_y: RD Y coordinaat
            collection: Data collectie
            year: Dataset jaar

        Returns:
            Lijst van (tile_id, url) tuples
        """
        col_name = self.COLLECTIONS.get(collection, collection)
        results = []

        # Alleen de centrale tile
        tile_id = self.get_tile_id(rd_x, rd_y)
        url = self._build_direct_url(tile_id, col_name, year)
        results.append((tile_id, url))

        return results

    def download_cityjson(self, download_url, progress_callback=None):
        """Download CityJSON ZIP en extract het JSON bestand.

        Gebruikt caching: als het ZIP bestand al in temp staat,
        wordt het niet opnieuw gedownload.

        Args:
            download_url: URL naar CityJSON ZIP bestand
            progress_callback: Optionele callback(message) voor progress

        Returns:
            Pad naar het uitgepakte CityJSON bestand

        Raises:
            PDOK3DError: Bij download of uitpak fouten
        """
        temp_dir = tempfile.gettempdir()

        # Bepaal bestandsnaam uit URL
        url_filename = download_url.rstrip("/").split("/")[-1]
        if not url_filename.endswith(".zip"):
            url_filename = "pdok3d_tile.zip"

        zip_path = os.path.join(temp_dir, url_filename)

        # Download ZIP (met caching)
        if not (os.path.exists(zip_path)
                and os.path.getsize(zip_path) > 1000):
            if progress_callback:
                progress_callback("CityJSON downloaden...")
            self._download_file(download_url, zip_path)

        # Extract CityJSON uit ZIP
        if progress_callback:
            progress_callback("CityJSON uitpakken...")

        return self._extract_cityjson(zip_path, temp_dir)

    # =========================================================================
    # URL constructie
    # =========================================================================

    def _build_direct_url(self, tile_id, collection, year):
        """Bouw directe download URL.

        Patroon:
        https://download.pdok.nl/kadaster/basisvoorziening-3d/v1_0/
        {year}/{collection}/{tile_id}.city.json.zip

        Args:
            tile_id: Tile ID (bijv. "106000_434000")
            collection: Collectie naam
            year: Dataset jaar

        Returns:
            Download URL string
        """
        return (
            "{base}/{year}/{collection}/{tile_id}.city.json.zip"
        ).format(
            base=self.DOWNLOAD_BASE_URL,
            year=year,
            collection=collection,
            tile_id=tile_id,
        )

    def _check_url_exists(self, url):
        """Controleer of een URL bereikbaar is (HEAD request).

        Returns:
            True als de URL een 200 response geeft
        """
        try:
            request = urllib2.Request(url)
            request.get_method = lambda: "HEAD"
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            response = urllib2.urlopen(request, timeout=10)
            return response.getcode() == 200
        except Exception:
            return False

    # =========================================================================
    # OGC API fallback
    # =========================================================================

    def _query_ogc_api(self, rd_x, rd_y, collection):
        """Query PDOK OGC API voor tile download URL.

        Converteert RD naar WGS84 bbox voor de API query.

        Args:
            rd_x: RD X coordinaat
            rd_y: RD Y coordinaat
            collection: Collectie naam

        Returns:
            Download URL string

        Raises:
            PDOK3DError: Bij API fouten
        """
        # Maak een kleine bbox rond het punt (100m)
        margin = 50
        lat_min, lon_min = rd_to_wgs84(rd_x - margin, rd_y - margin)
        lat_max, lon_max = rd_to_wgs84(rd_x + margin, rd_y + margin)

        # PDOK bbox parameter: lon_min,lat_min,lon_max,lat_max (WGS84)
        bbox_str = "{0},{1},{2},{3}".format(
            lon_min, lat_min, lon_max, lat_max)

        url = (
            "{base}/collections/{collection}/items"
            "?f=json"
            "&bbox={bbox}"
            "&limit=10"
        ).format(
            base=self.OGC_BASE_URL,
            collection=collection,
            bbox=bbox_str,
        )

        try:
            request = urllib2.Request(url)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            request.add_header("Accept", "application/json")
            response = urllib2.urlopen(request, timeout=self.timeout)
            data = json.loads(response.read().decode("utf-8"))
        except Exception as e:
            raise PDOK3DError(
                "PDOK API query mislukt: {0}".format(e))

        # Zoek download_link in features
        features = data.get("features", [])
        for feature in features:
            props = feature.get("properties", {})
            download_link = props.get("download_link", "")
            if download_link:
                return download_link

        raise PDOK3DError(
            "Geen CityJSON tile gevonden voor locatie "
            "RD ({0:.0f}, {1:.0f}).\n"
            "Controleer of de locatie in Nederland ligt.".format(
                rd_x, rd_y))

    # =========================================================================
    # Download en extractie
    # =========================================================================

    def _download_file(self, url, filepath):
        """Download een bestand in chunks.

        Args:
            url: Download URL
            filepath: Lokaal bestandspad

        Raises:
            PDOK3DError: Bij download fouten
        """
        try:
            request = urllib2.Request(url)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            response = urllib2.urlopen(request, timeout=self.timeout)

            with open(filepath, "wb") as f:
                while True:
                    chunk = response.read(65536)  # 64KB chunks
                    if not chunk:
                        break
                    f.write(chunk)

        except Exception as e:
            # Verwijder incompleet bestand
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except OSError:
                    pass
            self._raise_download_error(e)

    def _extract_cityjson(self, zip_path, output_dir):
        """Extract CityJSON bestand uit ZIP.

        Zoekt naar bestanden met .json of .city.json extensie.

        Args:
            zip_path: Pad naar ZIP bestand
            output_dir: Map voor extractie

        Returns:
            Pad naar uitgepakt CityJSON bestand

        Raises:
            PDOK3DError: Bij uitpak fouten
        """
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                # Zoek CityJSON bestand
                # Voorkeur: .city.json, dan .json
                candidates = []
                for name in zf.namelist():
                    lower = name.lower()
                    if lower.endswith(".city.json"):
                        candidates.insert(0, name)  # Voorkeur
                    elif lower.endswith(".json"):
                        candidates.append(name)

                if not candidates:
                    raise PDOK3DError(
                        "Geen CityJSON bestand gevonden in ZIP: {0}".format(
                            os.path.basename(zip_path)))

                # Extract eerste candidate - stream in chunks
                # om OutOfMemoryException te voorkomen
                chosen = candidates[0]
                out_filename = os.path.basename(chosen)
                out_path = os.path.join(output_dir, out_filename)

                with zf.open(chosen) as src, open(out_path, "wb") as dst:
                    while True:
                        chunk = src.read(262144)  # 256KB chunks
                        if not chunk:
                            break
                        dst.write(chunk)

                return out_path

        except zipfile.BadZipfile:
            raise PDOK3DError(
                "Ongeldig ZIP bestand: {0}".format(
                    os.path.basename(zip_path)))
        except PDOK3DError:
            raise
        except Exception as e:
            raise PDOK3DError(
                "Fout bij uitpakken ZIP: {0}".format(e))

    def _raise_download_error(self, e):
        """Vertaal download exceptie naar PDOK3DError."""
        error_str = str(e).lower()
        if "timeout" in error_str or "timed out" in error_str:
            raise PDOK3DError(
                "Download timeout. Het bestand is mogelijk te groot.\n"
                "Probeer 'Alleen gebouwen' in plaats van 'Gebouwen + Terreinen'."
            )
        elif "404" in str(e) or "not found" in error_str:
            raise PDOK3DError(
                "CityJSON data niet gevonden voor dit gebied.\n"
                "Controleer of de locatie in Nederland ligt."
            )
        else:
            raise PDOK3DError(
                "Download mislukt: {0}".format(e))
