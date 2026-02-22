# -*- coding: utf-8 -*-
"""
BAG3D Client
============

Client voor het downloaden van 3DBAG (3D gebouwmodellen) data.

Workflow:
1. WFS query met bbox -> tile IDs + OBJ download URLs
2. Download OBJ ZIP bestanden
3. Uitpakken: {tile}-{lod}-3d.obj

Gebruik:
    from gis2bim.api.bag3d import BAG3DClient

    client = BAG3DClient()
    tiles = client.get_tiles(bbox)
    obj_paths = client.download_tiles(tiles, lod="lod22")
"""

import os
import re
import tempfile
import zipfile

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2

try:
    import json
except ImportError:
    json = None

try:
    import xml.etree.ElementTree as ET
except ImportError:
    ET = None


class BAG3DError(Exception):
    """Fout bij het downloaden of verwerken van BAG3D data."""
    pass


class BAG3DTile(object):
    """Representatie van een 3DBAG tile.

    Attributes:
        tile_id: Tile identifier (bv. "10-282-562")
        obj_download_url: URL naar OBJ ZIP bestand
    """

    def __init__(self, tile_id, obj_download_url):
        self.tile_id = tile_id
        self.obj_download_url = obj_download_url

    def __repr__(self):
        return "BAG3DTile({0})".format(self.tile_id)


class BAG3DClient(object):
    """Client voor 3DBAG tile download.

    Attributes:
        WFS_URL: Basis URL van de 3DBAG WFS service
    """

    WFS_URL = "https://data.3dbag.nl/api/BAG3D/wfs"

    # LoD bestandsnaam patronen in ZIP (3DBAG gebruikt uppercase)
    LOD_PATTERNS = {
        "lod12": "-LoD12-3D.obj",
        "lod13": "-LoD13-3D.obj",
        "lod22": "-LoD22-3D.obj",
    }

    def __init__(self, timeout=60):
        """Initialiseer de client.

        Args:
            timeout: Download timeout in seconden
        """
        self.timeout = timeout

    def get_tiles(self, bbox):
        """Query WFS voor tiles die de bbox bedekken.

        Args:
            bbox: (xmin, ymin, xmax, ymax) in RD (EPSG:28992)

        Returns:
            Lijst van BAG3DTile objecten

        Raises:
            BAG3DError: Bij query fouten of geen tiles gevonden
        """
        xmin, ymin, xmax, ymax = bbox

        if xmin >= xmax or ymin >= ymax:
            raise BAG3DError("Ongeldige bounding box: {0}".format(bbox))

        url = (
            "{base}?SERVICE=WFS&VERSION=2.0.0&REQUEST=GetFeature"
            "&typeName=BAG3D:Tiles"
            "&bbox={xmin},{ymin},{xmax},{ymax}"
        ).format(
            base=self.WFS_URL,
            xmin=xmin,
            ymin=ymin,
            xmax=xmax,
            ymax=ymax,
        )

        try:
            request = urllib2.Request(url)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            response = urllib2.urlopen(request, timeout=self.timeout)
            xml_data = response.read()
        except Exception as e:
            self._raise_download_error(e)

        tiles = self._parse_wfs_response(xml_data)

        if not tiles:
            raise BAG3DError(
                "Geen 3DBAG tiles gevonden voor dit gebied.\n"
                "Controleer of de locatie in Nederland ligt."
            )

        return tiles

    def download_tiles(self, tiles, lod="lod22", progress_callback=None):
        """Download OBJ bestanden uit 3DBAG tile ZIPs.

        Args:
            tiles: Lijst van BAG3DTile objecten (van get_tiles())
            lod: Detail niveau ("lod12", "lod13", "lod22")
            progress_callback: Optionele callback(message) voor progress updates

        Returns:
            Lijst van paden naar uitgepakte OBJ bestanden

        Raises:
            BAG3DError: Bij download of uitpak fouten
        """
        if lod not in self.LOD_PATTERNS:
            raise BAG3DError(
                "Ongeldig LoD: '{0}'. Kies 'lod12', 'lod13' of 'lod22'.".format(lod)
            )

        lod_suffix = self.LOD_PATTERNS[lod]
        obj_paths = []
        temp_dir = tempfile.gettempdir()

        for i, tile in enumerate(tiles):
            if progress_callback:
                progress_callback("Downloaden tile {0}/{1}: {2}...".format(
                    i + 1, len(tiles), tile.tile_id))

            # Download ZIP (tile_id bevat '/', vervang door '-' voor bestandsnaam)
            safe_tile_id = tile.tile_id.replace("/", "-")
            zip_filename = "{0}-obj.zip".format(safe_tile_id)
            zip_path = os.path.join(temp_dir, zip_filename)

            # Skip als al gedownload (cache)
            if not (os.path.exists(zip_path) and os.path.getsize(zip_path) > 1000):
                self._download_file_streamed(tile.obj_download_url, zip_path)

            # Uitpakken en juiste OBJ zoeken
            if progress_callback:
                progress_callback("Uitpakken tile {0}/{1}: {2}...".format(
                    i + 1, len(tiles), tile.tile_id))

            obj_path = self._extract_obj_from_zip(zip_path, lod_suffix, temp_dir)
            if obj_path:
                obj_paths.append(obj_path)

        if not obj_paths:
            raise BAG3DError(
                "Geen OBJ bestanden gevonden voor LoD {0}.\n"
                "Mogelijk bevat dit gebied geen gebouwen.".format(lod)
            )

        return obj_paths

    # =========================================================================
    # WFS Response parsing
    # =========================================================================

    def _parse_wfs_response(self, xml_data):
        """Parse WFS XML response naar lijst van BAG3DTile objecten.

        Probeert eerst ElementTree XML parsing, dan regex fallback.

        Args:
            xml_data: Raw XML bytes van WFS response

        Returns:
            Lijst van BAG3DTile objecten
        """
        # Decodeer bytes naar string
        if isinstance(xml_data, bytes):
            try:
                xml_str = xml_data.decode("utf-8")
            except UnicodeDecodeError:
                xml_str = xml_data.decode("latin-1")
        else:
            xml_str = xml_data

        # Probeer XML parsing
        tiles = self._parse_wfs_xml(xml_str)
        if tiles:
            return tiles

        # Fallback: regex parsing
        return self._parse_wfs_regex(xml_str)

    def _parse_wfs_xml(self, xml_str):
        """Parse WFS response met ElementTree."""
        if ET is None:
            return []

        try:
            root = ET.fromstring(xml_str.encode("utf-8"))
        except ET.ParseError:
            return []

        tiles = []

        # Zoek alle elementen, ongeacht namespace
        for elem in root.iter():
            tag = elem.tag
            # Strip namespace
            if "}" in tag:
                tag = tag.split("}", 1)[1]

            if tag == "Tiles":
                tile_id = None
                obj_url = None

                for child in elem.iter():
                    child_tag = child.tag
                    if "}" in child_tag:
                        child_tag = child_tag.split("}", 1)[1]

                    if child_tag == "tile_id" and child.text:
                        tile_id = child.text.strip()
                    elif child_tag == "obj_download" and child.text:
                        obj_url = child.text.strip()

                if tile_id and obj_url:
                    tiles.append(BAG3DTile(tile_id, obj_url))

        return tiles

    def _parse_wfs_regex(self, xml_str):
        """Parse WFS response met regex (fallback)."""
        tiles = []

        # Zoek tile_id waarden
        tile_ids = re.findall(r"<[^>]*tile_id[^>]*>([^<]+)</", xml_str)
        # Zoek obj_download URLs
        obj_urls = re.findall(r"<[^>]*obj_download[^>]*>([^<]+)</", xml_str)

        # Koppel tile_ids aan obj_urls
        for idx in range(min(len(tile_ids), len(obj_urls))):
            tile_id = tile_ids[idx].strip()
            obj_url = obj_urls[idx].strip()
            if tile_id and obj_url:
                tiles.append(BAG3DTile(tile_id, obj_url))

        return tiles

    # =========================================================================
    # BAG Pand attributen
    # =========================================================================

    BAG_WFS_URL = "https://service.pdok.nl/kadaster/bag/wfs/v2_0"

    def get_bouwjaren(self, bbox):
        """Query BAG WFS voor bouwjaren van panden in een bounding box.

        Args:
            bbox: (xmin, ymin, xmax, ymax) in RD (EPSG:28992)

        Returns:
            Dict mapping pand_id -> bouwjaar (int).
            Pand_id is de BAG identificatie (bv. "0363100012345678").
            Retourneert lege dict bij fouten.
        """
        xmin, ymin, xmax, ymax = bbox

        url = (
            "{base}?service=WFS&version=2.0.0&request=GetFeature"
            "&typeName=bag:pand"
            "&bbox={xmin},{ymin},{xmax},{ymax}"
            "&outputFormat=application%2Fjson"
            "&propertyName=identificatie,oorspronkelijkBouwjaar"
        ).format(
            base=self.BAG_WFS_URL,
            xmin=xmin, ymin=ymin, xmax=xmax, ymax=ymax,
        )

        try:
            request = urllib2.Request(url)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            response = urllib2.urlopen(request, timeout=self.timeout)
            data = response.read()
        except Exception:
            return {}

        return self._parse_bouwjaren_json(data)

    def _parse_bouwjaren_json(self, data):
        """Parse GeoJSON response naar pand_id -> bouwjaar dict."""
        if not json:
            return self._parse_bouwjaren_regex(data)

        try:
            geojson = json.loads(data)
            result = {}
            for feature in geojson.get("features", []):
                props = feature.get("properties", {})
                pand_id = props.get("identificatie", "")
                bouwjaar = props.get("oorspronkelijkBouwjaar")
                if pand_id and bouwjaar:
                    try:
                        result[str(pand_id)] = int(bouwjaar)
                    except (ValueError, TypeError):
                        pass
            return result
        except Exception:
            return self._parse_bouwjaren_regex(data)

    def _parse_bouwjaren_regex(self, data):
        """Parse bouwjaar data met regex (fallback)."""
        if isinstance(data, bytes):
            data = data.decode("utf-8", errors="replace")

        result = {}
        # Zoek identificatie + bouwjaar paren
        ids = re.findall(
            r'"identificatie"\s*:\s*"([^"]+)"', data)
        years = re.findall(
            r'"oorspronkelijkBouwjaar"\s*:\s*(\d+)', data)

        for i in range(min(len(ids), len(years))):
            try:
                result[ids[i]] = int(years[i])
            except (ValueError, IndexError):
                pass

        return result

    # =========================================================================
    # ZIP extractie
    # =========================================================================

    def _extract_obj_from_zip(self, zip_path, lod_suffix, output_dir):
        """Zoek en extraheer het juiste OBJ bestand uit een ZIP.

        Args:
            zip_path: Pad naar ZIP bestand
            lod_suffix: LoD suffix om te zoeken (bv. "-lod22-3d.obj")
            output_dir: Map om naar uit te pakken

        Returns:
            Pad naar uitgepakt OBJ bestand, of None als niet gevonden
        """
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                # Zoek bestand met juiste LoD suffix
                for name in zf.namelist():
                    if name.endswith(lod_suffix):
                        # Uitpakken naar output directory
                        obj_filename = os.path.basename(name)
                        obj_path = os.path.join(output_dir, obj_filename)
                        # Lees uit ZIP en schrijf
                        data = zf.read(name)
                        with open(obj_path, "wb") as f:
                            f.write(data)
                        return obj_path

                # Fallback: zoek willekeurig OBJ bestand met .obj extensie
                for name in zf.namelist():
                    if name.lower().endswith(".obj"):
                        obj_filename = os.path.basename(name)
                        obj_path = os.path.join(output_dir, obj_filename)
                        data = zf.read(name)
                        with open(obj_path, "wb") as f:
                            f.write(data)
                        return obj_path

        except zipfile.BadZipfile:
            raise BAG3DError(
                "Ongeldig ZIP bestand: {0}".format(os.path.basename(zip_path))
            )
        except Exception as e:
            raise BAG3DError(
                "Fout bij uitpakken ZIP: {0}".format(e)
            )

        return None

    # =========================================================================
    # Download
    # =========================================================================

    def _download_file_streamed(self, url, filepath):
        """Download een bestand in chunks naar lokaal pad."""
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

    def _raise_download_error(self, e):
        """Vertaal download exceptie naar BAG3DError."""
        error_type = type(e).__name__
        error_str = str(e).lower()
        if "timeout" in error_str or "timed out" in error_str:
            raise BAG3DError(
                "Download timeout. Probeer een kleiner gebied."
            )
        elif "404" in str(e) or "not found" in error_str:
            raise BAG3DError(
                "3DBAG data niet gevonden voor dit gebied.\n"
                "Controleer of de locatie in Nederland ligt."
            )
        else:
            raise BAG3DError(
                "Download mislukt ({0}): {1}".format(error_type, str(e))
            )
