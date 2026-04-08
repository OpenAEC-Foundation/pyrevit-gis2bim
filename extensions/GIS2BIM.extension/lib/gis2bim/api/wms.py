# -*- coding: utf-8 -*-
"""
WMS Client + Laagdefinities
============================

WMS GetMap client voor het downloaden van rasterkaarten
van Nederlandse geo-services (PDOK, RIVM, Risicokaart, CBS).

Gebaseerd op WMS URLs uit GIS2BIM Dynamo:
- GIS2BIM_NetherlandsGeoservicesLibrary.dyf
- GIS2BIM_NetherlandsGeoservicesLibrary_RIVM.dyf

Gebruik:
    from gis2bim.api.wms import WMSClient, WMS_LAYERS

    client = WMSClient()
    bbox = (154500, 462500, 155500, 463500)
    url = client.build_getmap_url(WMS_LAYERS["luchtfoto_actueel"], bbox)
    client.download_image(WMS_LAYERS["luchtfoto_actueel"], bbox, "/tmp/luchtfoto.png")
"""

import os
import tempfile

try:
    # Python 2 (IronPython)
    import urllib2
    _PY2 = True
except ImportError:
    # Python 3
    import urllib.request as urllib2
    _PY2 = False


# =============================================================================
# WMS Laagdefinities
# =============================================================================

WMS_LAYERS = {
    # --- Achtergrond ---
    "luchtfoto_actueel": {
        "key": "luchtfoto_actueel",
        "name": "Luchtfoto (meest recent)",
        "category": "Achtergrond",
        "view_name": "gis2bim_luchtfoto",
        "base_url": "https://service.pdok.nl/hwh/luchtfotorgb/wms/v1_0",
        "layers": "Actueel_orthoHR",
        "styles": "",
        "width": 2500,
        "height": 2500,
        "crs": "EPSG:28992",
        "format": "image/jpeg",
        "version": "1.3.0",
        "transparent": False,
    },

    # --- Ruimtelijke Plannen ---
    # URL gemigreerd: /plu/ruimtelijkeplannen/wms/v6_0 -> /kadaster/ruimtelijke-plannen/wms/v1_0
    # Laagnamen: BP:Enkelbestemming -> enkelbestemming (zonder prefix)
    "enkelbestemming": {
        "key": "enkelbestemming",
        "name": "Enkelbestemming",
        "category": "Ruimtelijke Plannen",
        "view_name": "gis2bim_enkelbestemming",
        "base_url": "https://service.pdok.nl/kadaster/ruimtelijke-plannen/wms/v1_0",
        "layers": "enkelbestemming",
        "styles": "enkelbestemming",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "dubbelbestemming": {
        "key": "dubbelbestemming",
        "name": "Dubbelbestemming",
        "category": "Ruimtelijke Plannen",
        "view_name": "gis2bim_dubbelbestemming",
        "base_url": "https://service.pdok.nl/kadaster/ruimtelijke-plannen/wms/v1_0",
        "layers": "dubbelbestemming",
        "styles": "dubbelbestemming",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "bouwvlak": {
        "key": "bouwvlak",
        "name": "Bouwvlak",
        "category": "Ruimtelijke Plannen",
        "view_name": "gis2bim_bouwvlak",
        "base_url": "https://service.pdok.nl/kadaster/ruimtelijke-plannen/wms/v1_0",
        "layers": "bouwvlak",
        "styles": "bouwvlak",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },

    # --- Milieu - Geluid ---
    # URL gemigreerd: geodata.rivm.nl/geoserver -> data.rivm.nl/geo
    # Laagnamen: _actueel aliassen (updaten automatisch bij nieuwe data)
    "geluid_alle_bronnen": {
        "key": "geluid_alle_bronnen",
        "name": "Alle bronnen (Lden)",
        "category": "Milieu - Geluid",
        "view_name": "gis2bim_geluid_alle_bronnen",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_Geluid_lden_allebronnen_actueel",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "geluid_wegverkeer": {
        "key": "geluid_wegverkeer",
        "name": "Wegverkeer (Lden)",
        "category": "Milieu - Geluid",
        "view_name": "gis2bim_geluid_wegverkeer",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_Geluid_lden_wegverkeer_actueel",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "geluid_trein": {
        "key": "geluid_trein",
        "name": "Treinverkeer (Lden)",
        "category": "Milieu - Geluid",
        "view_name": "gis2bim_geluid_trein",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_Geluid_lden_treinverkeer_actueel",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },

    # --- Milieu - Lucht ---
    # _actueel aliassen voor PM25, PM10, NO2 (updaten automatisch)
    # EC heeft geen _actueel alias, meest recente versie gebruikt
    "lucht_pm25": {
        "key": "lucht_pm25",
        "name": "Fijnstof PM2.5",
        "category": "Milieu - Lucht",
        "view_name": "gis2bim_lucht_pm25",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_jaargemiddeld_PM25_actueel",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "lucht_pm10": {
        "key": "lucht_pm10",
        "name": "Fijnstof PM10",
        "category": "Milieu - Lucht",
        "view_name": "gis2bim_lucht_pm10",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_jaargemiddeld_PM10_actueel",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "lucht_no2": {
        "key": "lucht_no2",
        "name": "Stikstofdioxide NO2",
        "category": "Milieu - Lucht",
        "view_name": "gis2bim_lucht_no2",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_jaargemiddeld_NO2_actueel",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "lucht_ec": {
        "key": "lucht_ec",
        "name": "Roet (EC)",
        "category": "Milieu - Lucht",
        "view_name": "gis2bim_lucht_ec",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_nsl_20240401_gm_EC2022",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },

    # --- Milieu - Overig ---
    "geurhinder_openhaarden": {
        "key": "geurhinder_openhaarden",
        "name": "Geurhinder openhaarden",
        "category": "Milieu - Overig",
        "view_name": "gis2bim_geurhinder",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "rivm_20180905_v_geurhinder_open_haarden",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },

    # --- Water ---
    "beschermingszones": {
        "key": "beschermingszones",
        "name": "Dijkbeschermingszones",
        "category": "Water",
        "view_name": "gis2bim_dijkbescherming",
        "base_url": "https://service.pdok.nl/hwh/waterschappen-zoneringen-imwa/wms/v2_0",
        "layers": "beschermingszone",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "overstromingsrisico": {
        "key": "overstromingsrisico",
        "name": "Overstromingsrisico",
        "category": "Water",
        "view_name": "gis2bim_overstromingsrisico",
        "base_url": "https://data.rivm.nl/geo/alo/wms",
        "layers": "20231201_kans_overstroming",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },

    # --- Veiligheid ---
    # rfrisk.nl is opgeheven, vervangen door REV-portaal
    "pr_contour_ev": {
        "key": "pr_contour_ev",
        "name": "PR 10-6 contour (EV activiteiten)",
        "category": "Veiligheid",
        "view_name": "gis2bim_pr_contour",
        "base_url": "https://rev-portaal.nl/geoserver/wms",
        "layers": "rev_public:ev_pr10_6",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
    "transportroutes": {
        "key": "transportroutes",
        "name": "Basisnet transportroutes (PR 10-6)",
        "category": "Veiligheid",
        "view_name": "gis2bim_transportroutes",
        "base_url": "https://rev-portaal.nl/geoserver/wms",
        "layers": "rev_public:bn_pr10_6",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },

    # --- Natuur ---
    "natura2000": {
        "key": "natura2000",
        "name": "Natura 2000",
        "category": "Natuur",
        "view_name": "gis2bim_natura2000",
        "base_url": "https://service.pdok.nl/rvo/natura2000/wms/v1_0",
        "layers": "natura2000",
        "styles": "",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },

    # --- Statistiek ---
    # URL gemigreerd: jaar-specifiek endpoint, laagnaam gewijzigd
    "woz_waarde": {
        "key": "woz_waarde",
        "name": "WOZ-waarde (CBS 100m)",
        "category": "Statistiek",
        "view_name": "gis2bim_woz_waarde",
        "base_url": "https://service.pdok.nl/cbs/vierkantstatistieken100m/2024/wms/v1_0",
        "layers": "vierkant_100m",
        "styles": "cbsvierkant100m_gemiddelde_woz_waarde_woning",
        "width": 3000,
        "height": 3000,
        "crs": "EPSG:28992",
        "format": "image/png",
        "version": "1.3.0",
        "transparent": True,
    },
}


# Geordende lijst van categorieen voor UI
WMS_CATEGORIES = [
    "Achtergrond",
    "Ruimtelijke Plannen",
    "Milieu - Geluid",
    "Milieu - Lucht",
    "Milieu - Overig",
    "Water",
    "Veiligheid",
    "Natuur",
    "Statistiek",
]


def get_layers_by_category():
    """Geeft WMS_LAYERS gegroepeerd per categorie.

    Returns:
        OrderedDict-achtige lijst van (category, [layer_dicts])
    """
    result = []
    for cat in WMS_CATEGORIES:
        layers = [l for l in WMS_LAYERS.values() if l["category"] == cat]
        if layers:
            result.append((cat, layers))
    return result


def get_layer(key):
    """Haal een laag op per key.

    Args:
        key: Unieke laag-id (bv. "luchtfoto_actueel")

    Returns:
        Layer dict of None
    """
    return WMS_LAYERS.get(key)


# =============================================================================
# WMS Client
# =============================================================================

class WMSClient(object):
    """WMS GetMap client voor het downloaden van kaartbeelden."""

    def __init__(self, timeout=60):
        self.timeout = timeout
        self.user_agent = "GIS2BIM-pyRevit/1.0"

    def build_getmap_url(self, layer, bbox):
        """Bouw een volledige WMS GetMap URL.

        Args:
            layer: Layer dict uit WMS_LAYERS
            bbox: Tuple (xmin, ymin, xmax, ymax) in EPSG:28992

        Returns:
            Volledige GetMap URL string
        """
        params = [
            "service=WMS",
            "request=GetMap",
            "VERSION={0}".format(layer.get("version", "1.3.0")),
            "LAYERS={0}".format(layer["layers"]),
            "STYLES={0}".format(layer.get("styles", "")),
            "WIDTH={0}".format(layer.get("width", 3000)),
            "HEIGHT={0}".format(layer.get("height", 3000)),
            "FORMAT={0}".format(layer.get("format", "image/png")),
            "CRS={0}".format(layer.get("crs", "EPSG:28992")),
            "BBOX={0},{1},{2},{3}".format(bbox[0], bbox[1], bbox[2], bbox[3]),
        ]

        if layer.get("transparent", False):
            params.append("TRANSPARENT=TRUE")

        base = layer["base_url"]
        separator = "&" if "?" in base else "?"
        return base + separator + "&".join(params)

    def download_image(self, layer, bbox, output_path=None):
        """Download een WMS kaartbeeld als PNG.

        Args:
            layer: Layer dict uit WMS_LAYERS
            bbox: Tuple (xmin, ymin, xmax, ymax) in EPSG:28992
            output_path: Pad voor het output bestand. Indien None,
                         wordt een temp bestand aangemaakt.

        Returns:
            Pad naar het gedownloade bestand

        Raises:
            IOError: Bij download fouten
        """
        url = self.build_getmap_url(layer, bbox)

        if output_path is None:
            suffix = ".png"
            fmt = layer.get("format", "image/png")
            if "jpeg" in fmt or "jpg" in fmt:
                suffix = ".jpg"
            fd, output_path = tempfile.mkstemp(
                suffix=suffix,
                prefix="gis2bim_wms_"
            )
            os.close(fd)

        request = urllib2.Request(url)
        request.add_header("User-Agent", self.user_agent)

        try:
            response = urllib2.urlopen(request, timeout=self.timeout)
            data = response.read()

            # Check of het een geldig beeld is (niet een XML foutmelding)
            if data[:5] == b"<?xml" or data[:5] == b"<Serv":
                error_text = data[:500]
                if not isinstance(error_text, str):
                    error_text = error_text.decode("utf-8", errors="replace")
                raise IOError(
                    "WMS server fout voor {0}: {1}".format(
                        layer["name"], error_text
                    )
                )

            with open(output_path, "wb") as f:
                f.write(data)

            return output_path

        except Exception as e:
            # Opruimen bij fout
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except Exception:
                    pass
            raise IOError(
                "Download mislukt voor {0}: {1}".format(layer["name"], str(e))
            )
