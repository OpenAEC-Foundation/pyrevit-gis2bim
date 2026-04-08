# -*- coding: utf-8 -*-
"""
ArcGIS WMTS Tile Client
========================

Download en stitch WMTS tiles van de ArcGIS Historische Tijdreis service.
Gebruikt voor het downloaden van historische kaarten van Nederland.

Bron: https://tiles.arcgis.com/tiles/nSZVuSZjHpEZZbRo/arcgis/rest/services/
Beschikbare jaren: 1815 - 2019

Gebruik:
    from gis2bim.api.wmts_tiles import ArcGISTileClient, TIJDREIS_YEARS

    client = ArcGISTileClient()
    bbox = (154500, 462500, 155500, 463500)
    image_path = client.download_image("Historische_tijdreis_1900", bbox)
"""

import os
import math
import json
import hashlib
import tempfile

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2

# .NET imports voor image stitching
import clr
clr.AddReference("System.Drawing")
from System.Drawing import Bitmap, Graphics
from System.Drawing.Imaging import ImageFormat
from System.Net import WebRequest


ARCGIS_BASE = (
    "https://tiles.arcgis.com/tiles/nSZVuSZjHpEZZbRo"
    "/arcgis/rest/services"
)

# Alle beschikbare jaren in de Historische Tijdreis service
TIJDREIS_YEARS = [
    "1815", "1820", "1821", "1823_1829", "1830_1849", "1850",
    "1857", "1858", "1861", "1862", "1865", "1866", "1868", "1870", "1871",
    "1872", "1880", "1883", "1886", "1888", "1889", "1893", "1899", "1900",
    "1901", "1902", "1904", "1908", "1909", "1910", "1912", "1915", "1918",
    "1919", "1920", "1922", "1925", "1929", "1931", "1935", "1937", "1940",
    "1942", "1943", "1947", "1948", "1949", "1950", "1951", "1952", "1953",
    "1955", "1962", "1963", "1965", "1970", "1971", "1973", "1975", "1976",
    "1978", "1980", "1984", "1988", "1990", "1994", "1995", "1996", "1997",
    "1999", "2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007",
    "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016",
    "2017", "2018", "2019",
]


def _build_layer_name(year):
    """Bouw ArcGIS layer naam voor een jaar.

    Args:
        year: Jaar string (bijv. "1900" of "1823_1829")

    Returns:
        str: Volledige layer naam
    """
    return "Historische_tijdreis_{0}".format(year)


class ArcGISTileClient(object):
    """Client voor ArcGIS tiled map services (Historische Tijdreis).

    Download WMTS tiles, stitch ze samen met System.Drawing,
    en sla het resultaat op als PNG.
    """

    def __init__(self, timeout=30):
        self.timeout = timeout
        self._tile_info_cache = {}

    def _download_json(self, url):
        """Download en parse JSON van een URL."""
        request = urllib2.Request(url)
        request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
        response = urllib2.urlopen(request, timeout=self.timeout)
        data = response.read()
        if isinstance(data, bytes):
            data = data.decode("utf-8")
        return json.loads(data)

    def _get_tile_info(self, layer_name):
        """Haal tile info op via ArcGIS REST API.

        Args:
            layer_name: Volledige layer naam

        Returns:
            dict met origin, lods, tile dimensions
        """
        if layer_name in self._tile_info_cache:
            return self._tile_info_cache[layer_name]

        url = "{0}/{1}/MapServer?f=json".format(ARCGIS_BASE, layer_name)
        data = self._download_json(url)

        tile_info_raw = data.get("tileInfo", {})
        origin = tile_info_raw.get("origin", {})
        lods_raw = tile_info_raw.get("lods", [])

        tile_info = {
            "origin_x": float(origin.get("x", -285401.92)),
            "origin_y": float(origin.get("y", 903401.92)),
            "tile_width": int(tile_info_raw.get("cols", 256)),
            "tile_height": int(tile_info_raw.get("rows", 256)),
            "lods": [
                {
                    "level": int(lod["level"]),
                    "resolution": float(lod["resolution"]),
                    "scale": float(lod["scale"]),
                }
                for lod in lods_raw
            ],
        }

        self._tile_info_cache[layer_name] = tile_info
        return tile_info

    def _pick_lod(self, lods, target_resolution):
        """Kies het beste zoom level.

        Pikt het fijnste level waarvan de resolutie <= target is.

        Args:
            lods: Lijst van LOD dicts (gesorteerd grof -> fijn)
            target_resolution: Gewenste resolutie in m/pixel

        Returns:
            LOD dict
        """
        selected = lods[0]
        for lod in lods:
            if lod["resolution"] <= target_resolution:
                selected = lod
                break
            selected = lod
        return selected

    def _calc_tiles(self, tile_info, lod, bbox):
        """Bereken welke tiles de bounding box bedekken.

        Args:
            tile_info: Tile info dict
            lod: Gekozen LOD dict
            bbox: Tuple (xmin, ymin, xmax, ymax)

        Returns:
            dict met tile ranges en crop offsets
        """
        x_min, y_min, x_max, y_max = bbox

        origin_x = tile_info["origin_x"]
        origin_y = tile_info["origin_y"]
        tile_pixels = tile_info["tile_width"]
        resolution = lod["resolution"]
        tile_size = tile_pixels * resolution

        # Tile coordinaten berekenen (origin is linksboven)
        col_min_f = (x_min - origin_x) / tile_size
        row_min_f = (origin_y - y_max) / tile_size
        col_max_f = (x_max - origin_x) / tile_size
        row_max_f = (origin_y - y_min) / tile_size

        col_min = int(math.floor(col_min_f))
        row_min = int(math.floor(row_min_f))
        col_max = int(math.floor(col_max_f))
        row_max = int(math.floor(row_max_f))

        # Crop offsets in pixels
        left_crop = int((col_min_f - col_min) * tile_pixels)
        top_crop = int((row_min_f - row_min) * tile_pixels)
        right_crop = int((1.0 - (col_max_f - col_max)) * tile_pixels)
        bottom_crop = int((1.0 - (row_max_f - row_max)) * tile_pixels)

        return {
            "row_min": row_min,
            "row_max": row_max,
            "col_min": col_min,
            "col_max": col_max,
            "tile_pixels": tile_pixels,
            "level": lod["level"],
            "top_crop": top_crop,
            "left_crop": left_crop,
            "bottom_crop": bottom_crop,
            "right_crop": right_crop,
        }

    def _download_tile_bitmap(self, url):
        """Download een tile als System.Drawing.Bitmap."""
        request = WebRequest.Create(url)
        request.UserAgent = "GIS2BIM-pyRevit/1.0"
        request.Timeout = self.timeout * 1000
        response = request.GetResponse()
        stream = response.GetResponseStream()
        bitmap = Bitmap(stream)
        stream.Close()
        response.Close()
        return bitmap

    def _stitch_tiles(self, layer_name, tiles_info):
        """Download en stitch tiles tot een afbeelding.

        Args:
            layer_name: ArcGIS layer naam
            tiles_info: dict van _calc_tiles()

        Returns:
            System.Drawing.Bitmap
        """
        row_min = tiles_info["row_min"]
        row_max = tiles_info["row_max"]
        col_min = tiles_info["col_min"]
        col_max = tiles_info["col_max"]
        pixels = tiles_info["tile_pixels"]
        level = tiles_info["level"]

        row_count = row_max - row_min + 1
        col_count = col_max - col_min + 1

        # Finale afbeelding grootte na cropping
        total_w = col_count * pixels - tiles_info["left_crop"] - tiles_info["right_crop"]
        total_h = row_count * pixels - tiles_info["top_crop"] - tiles_info["bottom_crop"]

        combined = Bitmap(int(max(total_w, 1)), int(max(total_h, 1)))
        graphics = Graphics.FromImage(combined)

        for r in range(row_min, row_max + 1):
            for c in range(col_min, col_max + 1):
                url = "{0}/{1}/MapServer/tile/{2}/{3}/{4}".format(
                    ARCGIS_BASE, layer_name, level, r, c
                )
                try:
                    tile = self._download_tile_bitmap(url)
                    x = (c - col_min) * pixels - tiles_info["left_crop"]
                    y = (r - row_min) * pixels - tiles_info["top_crop"]
                    graphics.DrawImage(tile, x, y, pixels, pixels)
                    tile.Dispose()
                except Exception:
                    pass  # Skip failed tiles (leeg gebied)

        graphics.Dispose()
        return combined

    def get_sample_hash(self, year, bbox):
        """Download een enkele sample-tile en geef de MD5 hash terug.

        Gebruikt om te detecteren of twee jaren hetzelfde beeldmateriaal
        tonen voor een bepaalde locatie.

        Args:
            year: Jaar string (bijv. "1900")
            bbox: Tuple (xmin, ymin, xmax, ymax) in RD

        Returns:
            str: MD5 hex digest van de tile, of None bij fout
        """
        try:
            layer_name = _build_layer_name(year)
            tile_info = self._get_tile_info(layer_name)
            # Gebruik een middelmatige resolutie voor snelle download
            lod = self._pick_lod(tile_info["lods"], 2.0)
            tiles = self._calc_tiles(tile_info, lod, bbox)

            # Download alleen de centrale tile
            center_row = (tiles["row_min"] + tiles["row_max"]) // 2
            center_col = (tiles["col_min"] + tiles["col_max"]) // 2
            url = "{0}/{1}/MapServer/tile/{2}/{3}/{4}".format(
                ARCGIS_BASE, layer_name, tiles["level"],
                center_row, center_col
            )

            request = urllib2.Request(url)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            response = urllib2.urlopen(request, timeout=self.timeout)
            data = response.read()
            return hashlib.md5(data).hexdigest()
        except Exception:
            return None

    def download_image(self, year, bbox, target_resolution=None):
        """Download historische kaart als afbeelding.

        Args:
            year: Jaar string (bijv. "1900")
            bbox: Tuple (xmin, ymin, xmax, ymax) in EPSG:28992 (RD)
            target_resolution: Resolutie in m/pixel.
                               Indien None, wordt automatisch bepaald
                               op basis van bbox grootte (~2000px).

        Returns:
            str: Pad naar opgeslagen PNG afbeelding

        Raises:
            Exception: Bij download- of verwerkingsfouten
        """
        layer_name = _build_layer_name(year)

        # Automatische resolutie: ~2000 pixels voor de langste zijde
        if target_resolution is None:
            bbox_w = bbox[2] - bbox[0]
            bbox_h = bbox[3] - bbox[1]
            target_resolution = max(bbox_w, bbox_h) / 2000.0

        tile_info = self._get_tile_info(layer_name)
        lod = self._pick_lod(tile_info["lods"], target_resolution)
        tiles = self._calc_tiles(tile_info, lod, bbox)

        combined = self._stitch_tiles(layer_name, tiles)

        fd, tmp_path = tempfile.mkstemp(suffix=".png", prefix="gis2bim_topo_")
        os.close(fd)
        combined.Save(tmp_path, ImageFormat.Png)
        combined.Dispose()

        return tmp_path
