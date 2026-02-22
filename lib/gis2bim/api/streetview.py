# -*- coding: utf-8 -*-
"""
Google Street View Static API Client
=====================================

Download Street View images via de Google Street View Static API.
Elke aanroep retourneert een JPEG image voor een specifieke heading.

Vereist een Google API key met Street View Static API ingeschakeld.
"""

import os
import tempfile

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2


STREETVIEW_BASE_URL = "https://maps.googleapis.com/maps/api/streetview"


class StreetViewClient(object):
    """Client voor de Google Street View Static API."""

    def __init__(self, api_key, timeout=30):
        """
        Args:
            api_key: Google API key met Street View Static API access.
            timeout: HTTP timeout in seconden.
        """
        self.api_key = api_key
        self.timeout = timeout

    def _build_url(self, lat, lon, heading, fov=75, pitch=0,
                   width=1024, height=1024):
        """Bouw Google Street View Static API URL.

        Args:
            lat: Latitude (WGS84).
            lon: Longitude (WGS84).
            heading: Kompasrichting 0-360 (0=Noord, 90=Oost, etc.).
            fov: Field of view in graden (default 75).
            pitch: Camera pitch in graden (default 0).
            width: Breedte in pixels (max 640 gratis, 2048 premium).
            height: Hoogte in pixels.

        Returns:
            str: Volledige API URL.
        """
        params = [
            "size={0}x{1}".format(width, height),
            "location={0},{1}".format(lat, lon),
            "heading={0}".format(heading),
            "pitch={0}".format(pitch),
            "fov={0}".format(fov),
            "key={0}".format(self.api_key),
        ]
        return "{0}?{1}".format(STREETVIEW_BASE_URL, "&".join(params))

    def download_image(self, lat, lon, heading, fov=75, pitch=0,
                       width=1024, height=1024):
        """Download een Street View image.

        Args:
            lat: Latitude (WGS84).
            lon: Longitude (WGS84).
            heading: Kompasrichting 0-360.
            fov: Field of view in graden.
            pitch: Camera pitch in graden.
            width: Breedte in pixels.
            height: Hoogte in pixels.

        Returns:
            str: Pad naar gedownload .jpg temp bestand.

        Raises:
            Exception: Bij download- of API-fouten.
        """
        url = self._build_url(lat, lon, heading, fov, pitch, width, height)

        request = urllib2.Request(url)
        request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

        response = urllib2.urlopen(request, timeout=self.timeout)
        data = response.read()

        # Controleer of het een geldig JPEG is (begint met FF D8)
        if len(data) < 2:
            raise Exception("Lege response van Google Street View API")

        # Google retourneert soms een grijs beeld als er geen coverage is.
        # We controleren op content-type header.
        content_type = response.headers.get("Content-Type", "")
        if "image" not in content_type:
            raise Exception(
                "Onverwacht response type: {0}".format(content_type)
            )

        # Schrijf naar temp bestand
        fd, tmp_path = tempfile.mkstemp(suffix=".jpg", prefix="gis2bim_sv_")
        try:
            os.write(fd, data)
        finally:
            os.close(fd)

        return tmp_path
