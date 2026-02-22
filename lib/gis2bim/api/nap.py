# -*- coding: utf-8 -*-
"""
NAP Peilmerken WFS Client
==========================

Client voor het ophalen van NAP peilmerken via de Rijkswaterstaat WFS service.
NAP peilmerken zijn officiele hoogtepunten die als referentie dienen in bouwprojecten.

Gebruik:
    from gis2bim.api.nap import NAPClient

    client = NAPClient()
    peilmerken = client.get_peilmerken((155000, 463000, 156000, 464000))
    for pm in peilmerken:
        print(pm.hoogte_label)  # "+2.81 NAP"
"""

import json

try:
    import urllib2
    from urllib import urlencode
except ImportError:
    import urllib.request as urllib2
    from urllib.parse import urlencode


# WFS endpoint Rijkswaterstaat NAP peilmerken
NAP_WFS_URL = "https://geo.rijkswaterstaat.nl/services/ogc/gdr/nap/ows"
NAP_LAYER = "punten_actueel"


class NAPPeilmerk(object):
    """
    Container voor een NAP peilmerk.

    Attributes:
        puntnummer: Uniek identificatienummer van het peilmerk
        hoogte: Hoogte in meters t.o.v. NAP
        x_rd: RD X coordinaat in meters
        y_rd: RD Y coordinaat in meters
        omschrijving: Beschrijving van de locatie
        projectdatum: Datum van de laatste meting
        bereikbaar: Of het peilmerk bereikbaar is ("J" of "N")
        status: Status van het peilmerk (bijv. "ACTUEEL")
    """

    def __init__(self, puntnummer, hoogte, x_rd, y_rd,
                 omschrijving="", projectdatum="", bereikbaar="J",
                 status="ACTUEEL"):
        self.puntnummer = puntnummer
        self.hoogte = hoogte
        self.x_rd = x_rd
        self.y_rd = y_rd
        self.omschrijving = omschrijving
        self.projectdatum = projectdatum
        self.bereikbaar = bereikbaar
        self.status = status

    @property
    def hoogte_label(self):
        """Formateer hoogte als leesbaar NAP label.

        Returns:
            String zoals "+2.81 NAP" of "-0.45 NAP"
        """
        if self.hoogte >= 0:
            return "+{0:.2f} NAP".format(self.hoogte)
        else:
            return "{0:.2f} NAP".format(self.hoogte)

    def __repr__(self):
        return "NAPPeilmerk('{0}', {1})".format(
            self.puntnummer, self.hoogte_label
        )


class NAPClient(object):
    """
    WFS client voor Rijkswaterstaat NAP peilmerken.

    Haalt actuele NAP peilmerken op via de WFS service van
    Rijkswaterstaat. Retourneert NAPPeilmerk objecten met
    hoogte, locatie en metadata.
    """

    def __init__(self, timeout=30):
        self.timeout = timeout
        self.wfs_url = NAP_WFS_URL
        self.layer = NAP_LAYER

    def get_peilmerken(self, bbox, alleen_bereikbaar=False):
        """
        Haal NAP peilmerken op binnen een bounding box.

        Args:
            bbox: Tuple (xmin, ymin, xmax, ymax) in RD coordinaten
            alleen_bereikbaar: Als True, filter op bereikbare merken

        Returns:
            Lijst van NAPPeilmerk objecten
        """
        url = self._build_url(bbox)

        try:
            request = urllib2.Request(url)
            request.add_header("Accept", "application/json")
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

            response = urllib2.urlopen(request, timeout=self.timeout)
            data = json.loads(response.read().decode("utf-8"))

            return self._parse_response(data, alleen_bereikbaar)

        except Exception as e:
            print("NAP WFS request error: {0}".format(e))
            return []

    def _build_url(self, bbox):
        """Bouw WFS GetFeature URL voor NAP peilmerken."""
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

    def _parse_response(self, data, alleen_bereikbaar):
        """Parse GeoJSON response naar NAPPeilmerk objecten."""
        peilmerken = []

        for feature in data.get("features", []):
            props = feature.get("properties", {})
            geom = feature.get("geometry", {})

            if not geom:
                continue

            # Status check
            status = props.get("status", "")
            if status != "ACTUEEL":
                continue

            # Coordinaten uit geometry
            coords = geom.get("coordinates", [])
            if len(coords) < 2:
                continue

            x_rd = coords[0]
            y_rd = coords[1]

            # Hoogte ophalen
            hoogte = props.get("hoogte", None)
            if hoogte is None:
                continue

            try:
                hoogte = float(hoogte)
            except (ValueError, TypeError):
                continue

            # Bereikbaarheid
            bereikbaar = props.get("bereikbaar", "J")
            if alleen_bereikbaar and bereikbaar != "J":
                continue

            peilmerk = NAPPeilmerk(
                puntnummer=str(props.get("puntnummer", "")),
                hoogte=hoogte,
                x_rd=x_rd,
                y_rd=y_rd,
                omschrijving=str(props.get("omschrijving", "")),
                projectdatum=str(props.get("projectdatum", "")),
                bereikbaar=bereikbaar,
                status=status
            )
            peilmerken.append(peilmerk)

        return peilmerken
