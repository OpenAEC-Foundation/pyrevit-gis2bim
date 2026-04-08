# -*- coding: utf-8 -*-
"""
AHN Client
==========

Client voor het downloaden van AHN (Actueel Hoogtebestand Nederland) data.

Twee methoden:
1. WCS (GeoTIFF) - Snel, klein bestand, geen externe tools
2. LAZ (Puntenwolk) - Hogere dichtheid, vereist LAStools

Gebruik:
    from gis2bim.api.ahn import AHNClient

    client = AHNClient()

    # Methode 1: WCS GeoTIFF
    tiff_path = client.download_geotiff(bbox, coverage="dtm")

    # Methode 2: LAZ Puntenwolk
    laz_paths = client.download_laz_tiles(bbox)
    result_path, fmt = client.process_laz(laz_paths[0], bbox=bbox)
"""

import os
import tempfile
import subprocess

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2


class AHNError(Exception):
    """Fout bij het downloaden of verwerken van AHN data."""
    pass


class AHNClient(object):
    """Client voor PDOK AHN data (WCS en LAZ tiles).

    Attributes:
        WCS_URL: Basis URL van de PDOK AHN WCS service
        LAZ_BASE_URL: Basis URL voor AHN5 COPC LAZ tiles
        TILE_SIZE: Grootte van LAZ tiles in meters (1km)
    """

    WCS_URL = "https://service.pdok.nl/rws/ahn/wcs/v1_0"
    LAZ_BASE_URL = "https://fsn1.your-objectstorage.com/hwh-ahn/AHN5_KM/01_LAZ"
    TILE_SIZE = 1000  # 1km tiles

    COVERAGES = {
        "dtm": "dtm_05m",
        "dsm": "dsm_05m",
    }

    # Meegeleverde LAStools in extensie bin map (eerste prioriteit)
    _BUNDLED_LASTOOLS = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(
            os.path.dirname(os.path.abspath(__file__))))),
        "bin", "lastools"
    )

    # Veelvoorkomende LAStools installatiepaden op Windows
    LASTOOLS_SEARCH_PATHS = [
        _BUNDLED_LASTOOLS,
        r"C:\LAStools\bin",
        r"C:\Program Files\LAStools\bin",
        r"C:\Program Files (x86)\LAStools\bin",
    ]

    def __init__(self, timeout=60):
        """Initialiseer de client.

        Args:
            timeout: Download timeout in seconden (60 voor WCS, 300 voor LAZ)
        """
        self.timeout = timeout

    # =========================================================================
    # WCS (GeoTIFF) methoden
    # =========================================================================

    def download_geotiff(self, bbox, coverage="dtm", resolution=0.5):
        """Download AHN data als GeoTIFF via WCS.

        Args:
            bbox: (xmin, ymin, xmax, ymax) in RD (EPSG:28992)
            coverage: "dtm" (maaiveld) of "dsm" (oppervlakte)
            resolution: meters per pixel (0.5 = native AHN resolutie)

        Returns:
            Pad naar gedownload GeoTIFF bestand in temp folder

        Raises:
            AHNError: Bij download fouten of ongeldige parameters
        """
        coverage_id = self.COVERAGES.get(coverage)
        if not coverage_id:
            raise AHNError(
                "Ongeldig coverage type: '{0}'. Kies 'dtm' of 'dsm'.".format(coverage)
            )

        xmin, ymin, xmax, ymax = bbox

        if xmin >= xmax or ymin >= ymax:
            raise AHNError("Ongeldige bounding box: {0}".format(bbox))

        width_m = xmax - xmin
        height_m = ymax - ymin
        width_px = int(round(width_m / resolution))
        height_px = int(round(height_m / resolution))

        if width_px <= 0 or height_px <= 0:
            raise AHNError(
                "Bounding box te klein voor resolutie {0}m".format(resolution)
            )

        url = self._build_wcs_url(coverage_id, bbox, width_px, height_px)

        temp_dir = tempfile.gettempdir()
        filename = "ahn_{0}_{1}x{2}.tif".format(coverage, width_px, height_px)
        filepath = os.path.join(temp_dir, filename)

        self._download_file(url, filepath)

        file_size = os.path.getsize(filepath)
        if file_size < 100:
            try:
                with open(filepath, "r") as f:
                    content = f.read(500)
                if "Exception" in content or "error" in content.lower():
                    raise AHNError(
                        "WCS server fout: {0}".format(content[:200])
                    )
            except UnicodeDecodeError:
                pass
            raise AHNError(
                "Download te klein ({0} bytes). Mogelijk geen data beschikbaar.".format(
                    file_size)
            )

        return filepath

    def estimate_points(self, bbox, resolution=0.5):
        """Schat het aantal punten voor een bbox en resolutie (WCS).

        Returns:
            Geschat aantal punten (width * height pixels)
        """
        xmin, ymin, xmax, ymax = bbox
        width_px = int(round((xmax - xmin) / resolution))
        height_px = int(round((ymax - ymin) / resolution))
        return width_px * height_px

    # =========================================================================
    # LAZ (Puntenwolk) methoden
    # =========================================================================

    def get_laz_tile_urls(self, bbox):
        """Bereken welke LAZ tiles de bbox bedekken.

        Args:
            bbox: (xmin, ymin, xmax, ymax) in RD

        Returns:
            Lijst van (url, filename, tile_bbox) tuples
        """
        xmin, ymin, xmax, ymax = bbox

        # Bereken tile indices (afronden naar beneden op 1km grid)
        x_start = int(xmin // self.TILE_SIZE) * self.TILE_SIZE
        y_start = int(ymin // self.TILE_SIZE) * self.TILE_SIZE
        x_end = int(xmax // self.TILE_SIZE) * self.TILE_SIZE
        y_end = int(ymax // self.TILE_SIZE) * self.TILE_SIZE

        tiles = []
        x = x_start
        while x <= x_end:
            y = y_start
            while y <= y_end:
                filename = "AHN5_C_{0}_{1}.COPC.LAZ".format(x, y)
                url = "{0}/{1}".format(self.LAZ_BASE_URL, filename)
                tile_bbox = (x, y, x + self.TILE_SIZE, y + self.TILE_SIZE)
                tiles.append((url, filename, tile_bbox))
                y += self.TILE_SIZE
            x += self.TILE_SIZE

        return tiles

    def download_laz_tiles(self, bbox, progress_callback=None):
        """Download LAZ tiles die de bbox bedekken.

        Args:
            bbox: (xmin, ymin, xmax, ymax) in RD
            progress_callback: Optionele callback(message) voor progress updates

        Returns:
            Lijst van paden naar gedownloade LAZ bestanden

        Raises:
            AHNError: Bij download fouten
        """
        tiles = self.get_laz_tile_urls(bbox)

        if not tiles:
            raise AHNError("Geen AHN tiles gevonden voor bbox")

        paths = []
        temp_dir = tempfile.gettempdir()

        for i, (url, filename, tile_bbox) in enumerate(tiles):
            filepath = os.path.join(temp_dir, filename)

            if progress_callback:
                progress_callback("Downloaden tile {0}/{1}: {2}...".format(
                    i + 1, len(tiles), filename))

            # Skip als al gedownload (cache)
            if os.path.exists(filepath) and os.path.getsize(filepath) > 1000:
                paths.append(filepath)
                continue

            self._download_file_streamed(url, filepath)
            paths.append(filepath)

        return paths

    def find_lastools(self):
        """Zoek LAStools executables op het systeem.

        Zoekt in PATH en veelvoorkomende installatiepaden.
        Ondersteunt zowel 32-bit (las2txt.exe) als 64-bit (las2txt64.exe) varianten.

        Returns:
            dict met tool paden: {"laszip": path, "las2las": path, "las2txt": path}
            Ontbrekende tools hebben None als waarde.
        """
        tools = {"laszip": None, "las2las": None, "las2txt": None}

        # Zoek in PATH en bekende locaties
        path_dirs = os.environ.get("PATH", "").split(os.pathsep)
        search_dirs = list(path_dirs) + self.LASTOOLS_SEARCH_PATHS

        for dir_path in search_dirs:
            if not os.path.isdir(dir_path):
                continue
            for tool_name in list(tools.keys()):
                if tools[tool_name]:
                    continue  # Al gevonden
                # Probeer eerst 64-bit variant, dan 32-bit
                for suffix in ["64.exe", ".exe"]:
                    exe_path = os.path.join(dir_path, tool_name + suffix)
                    if os.path.isfile(exe_path):
                        tools[tool_name] = exe_path
                        break

        return tools

    def process_laz(self, laz_path, output_base, bbox=None, tools=None,
                    keep_every_nth=None, keep_classification=None):
        """Verwerk LAZ bestand met LAStools.

        Probeert in volgorde:
        1. las2txt: directe conversie naar XYZ tekst met spatial filter
        2. las2las: conversie naar LAS met spatial filter
        3. laszip: decompressie naar LAS (zonder spatial filter)

        Args:
            laz_path: Pad naar .laz bestand
            output_base: Basis pad voor output (zonder extensie)
            bbox: Optioneel (xmin, ymin, xmax, ymax) voor spatial filter
            tools: Optioneel dict van tool paden (van find_lastools())
            keep_every_nth: Optioneel int, bewaar elke Nde punt (thinning)
            keep_classification: Optioneel lijst van classificatie codes
                (bv. [2, 9] voor ground + water bij DTM)

        Returns:
            (output_path, format) tuple waar format "xyz" of "las" is

        Raises:
            AHNError: Als geen LAStools gevonden of verwerking faalt
        """
        if tools is None:
            tools = self.find_lastools()

        # Gemeenschappelijke filter argumenten
        extra_args = []
        if bbox:
            extra_args.extend(["-keep_xy",
                               str(bbox[0]), str(bbox[1]),
                               str(bbox[2]), str(bbox[3])])
        if keep_classification:
            extra_args.append("-keep_classification")
            for cls in keep_classification:
                extra_args.append(str(cls))
        if keep_every_nth and keep_every_nth >= 2:
            extra_args.extend(["-keep_every_nth", str(keep_every_nth)])

        # Optie 1: las2txt - directe LAZ naar XYZ tekst met spatial filter
        if tools.get("las2txt"):
            xyz_path = output_base + ".xyz"
            cmd = [tools["las2txt"], "-i", laz_path, "-o", xyz_path,
                   "-parse", "xyz", "-sep", "space"]
            cmd.extend(extra_args)

            if self._run_tool(cmd):
                if os.path.exists(xyz_path) and os.path.getsize(xyz_path) > 0:
                    return (xyz_path, "xyz")

        # Optie 2: las2las - LAZ naar LAS met spatial filter
        if tools.get("las2las"):
            las_path = output_base + ".las"
            cmd = [tools["las2las"], "-i", laz_path, "-o", las_path]
            cmd.extend(extra_args)

            if self._run_tool(cmd):
                if os.path.exists(las_path) and os.path.getsize(las_path) > 0:
                    return (las_path, "las")

        # Optie 3: laszip - enkel decompressie (geen spatial filter)
        if tools.get("laszip"):
            las_path = output_base + ".las"
            cmd = [tools["laszip"], "-i", laz_path, "-o", las_path]

            if self._run_tool(cmd):
                if os.path.exists(las_path) and os.path.getsize(las_path) > 0:
                    return (las_path, "las")

        raise AHNError(
            "Geen LAStools gevonden of verwerking mislukt.\n\n"
            "Installeer LAStools en zorg dat las2txt.exe, las2las.exe "
            "of laszip.exe beschikbaar is in:\n"
            "- PATH omgevingsvariabele\n"
            "- C:\\LAStools\\bin\\\n\n"
            "Download: https://rapidlasso.de/lastools/"
        )

    def get_lastools_status(self):
        """Controleer welke LAStools beschikbaar zijn.

        Returns:
            dict met {"available": bool, "tools": dict, "message": str}
        """
        tools = self.find_lastools()
        available = any(v is not None for v in tools.values())

        if not available:
            message = "Geen LAStools gevonden"
        else:
            found = [name for name, path in tools.items() if path]
            message = "Gevonden: {0}".format(", ".join(found))

        return {
            "available": available,
            "tools": tools,
            "message": message,
        }

    # =========================================================================
    # Private methoden
    # =========================================================================

    def _build_wcs_url(self, coverage_id, bbox, width_px, height_px):
        """Bouw WCS GetCoverage URL."""
        xmin, ymin, xmax, ymax = bbox

        url = (
            "{base}?SERVICE=WCS&VERSION=1.0.0&REQUEST=GetCoverage"
            "&FORMAT=GEOTIFF"
            "&COVERAGE={coverage}"
            "&BBOX={xmin},{ymin},{xmax},{ymax}"
            "&CRS=EPSG:28992&RESPONSE_CRS=EPSG:28992"
            "&WIDTH={width}&HEIGHT={height}"
        ).format(
            base=self.WCS_URL,
            coverage=coverage_id,
            xmin=xmin,
            ymin=ymin,
            xmax=xmax,
            ymax=ymax,
            width=width_px,
            height=height_px,
        )

        return url

    def _download_file(self, url, filepath):
        """Download een klein bestand (WCS GeoTIFF) naar lokaal pad."""
        try:
            request = urllib2.Request(url)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            response = urllib2.urlopen(request, timeout=self.timeout)
            data = response.read()

            with open(filepath, "wb") as f:
                f.write(data)

        except Exception as e:
            self._raise_download_error(e)

    def _download_file_streamed(self, url, filepath):
        """Download een groot bestand (LAZ tile) in chunks naar lokaal pad."""
        try:
            request = urllib2.Request(url)
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")
            response = urllib2.urlopen(request, timeout=300)  # 5 min timeout

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
        """Vertaal download exceptie naar AHNError."""
        error_type = type(e).__name__
        error_str = str(e).lower()
        if "timeout" in error_str or "timed out" in error_str:
            raise AHNError(
                "Download timeout. Probeer een kleiner gebied "
                "of lagere resolutie."
            )
        elif "404" in str(e) or "not found" in error_str:
            raise AHNError(
                "AHN data niet gevonden voor dit gebied. "
                "Controleer of de locatie in Nederland ligt."
            )
        else:
            raise AHNError(
                "Download mislukt ({0}): {1}".format(error_type, str(e))
            )

    def _run_tool(self, cmd):
        """Voer een LAStools commando uit.

        Args:
            cmd: Lijst met commando en argumenten

        Returns:
            True als het commando succesvol was, anders False
        """
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            stdout, stderr = proc.communicate()
            return proc.returncode == 0
        except Exception:
            return False
