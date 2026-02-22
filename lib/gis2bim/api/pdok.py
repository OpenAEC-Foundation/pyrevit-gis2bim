# -*- coding: utf-8 -*-
"""
PDOK API Clients
================

Clients voor PDOK (Publieke Dienstverlening Op de Kaart) services:
- Locatieserver: Geocoding (adres -> coördinaten)
- BGT: Basisregistratie Grootschalige Topografie
- WMTS: Luchtfoto's en kaartlagen

API Documentatie:
    https://api.pdok.nl/
"""

import json
import time
import os

try:
    import requests
except ImportError:
    requests = None


# PDOK API Endpoints
class PDOKEndpoints:
    """PDOK API endpoint URLs."""
    LOCATIE_FREE = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free"
    LOCATIE_SUGGEST = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/suggest"
    LOCATIE_LOOKUP = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/lookup"
    BGT_DOWNLOAD = "https://api.pdok.nl/lv/bgt/download/v1_0/full/custom"
    BGT_STATUS = "https://api.pdok.nl/lv/bgt/download/v1_0/full/custom/{id}/status"
    LUCHTFOTO_WMTS = "https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0"


# BGT Feature Types
BGT_FEATURES = [
    "bak", "begroeidterreindeel", "bord", "buurt", "functioneelgebied",
    "gebouwinstallatie", "installatie", "kast", "kunstwerkdeel", "mast",
    "onbegroeidterreindeel", "ondersteunendwaterdeel", "ondersteunendwegdeel",
    "ongeclassificeerdobject", "openbareruimte", "openbareruimtelabel",
    "overbruggingsdeel", "overigbouwwerk", "overigescheiding", "paal",
    "pand", "plaatsbepalingspunt", "put", "scheiding", "sensor", "spoor",
    "stadsdeel", "straatmeubilair", "tunneldeel", "vegetatieobject",
    "waterdeel", "waterinrichtingselement", "waterschap", "wegdeel",
    "weginrichtingselement", "wijk"
]


# Windgebied mapping per provincie
WINDGEBIED_MAPPING = {
    "Zuid-Holland": 2,
    "Noord-Holland": "1_2",
    "Zeeland": 2,
    "Noord-Brabant": 3,
    "Limburg": 3,
    "Gelderland": 3,
    "Utrecht": 3,
    "Flevoland": 2,
    "Overijssel": 3,
    "Friesland": 2,
    "Groningen": 2,
    "Drenthe": 3,
}


class LocationData:
    """Container voor locatiegegevens."""
    
    def __init__(self, rd_x, rd_y, lat, lon, postcode, gemeente, provincie,
                 waterschap, kadaster_gemeente, kadaster_sectie, kadaster_perceel,
                 windgebied, url):
        self.rd_x = rd_x
        self.rd_y = rd_y
        self.lat = lat
        self.lon = lon
        self.postcode = postcode
        self.gemeente = gemeente
        self.provincie = provincie
        self.waterschap = waterschap
        self.kadaster_gemeente = kadaster_gemeente
        self.kadaster_sectie = kadaster_sectie
        self.kadaster_perceel = kadaster_perceel
        self.windgebied = windgebied
        self.url = url
    
    def to_dict(self):
        """Converteer naar dictionary."""
        return {
            "rd_x": self.rd_x,
            "rd_y": self.rd_y,
            "lat": self.lat,
            "lon": self.lon,
            "postcode": self.postcode,
            "gemeente": self.gemeente,
            "provincie": self.provincie,
            "waterschap": self.waterschap,
            "kadaster_gemeente": self.kadaster_gemeente,
            "kadaster_sectie": self.kadaster_sectie,
            "kadaster_perceel": self.kadaster_perceel,
            "windgebied": self.windgebied,
            "url": self.url,
        }


class PDOKLocatie:
    """
    Client voor PDOK Locatieserver API.
    
    Example:
        loc = PDOKLocatie()
        data = loc.search_address("Amsterdam", "Dam", "1")
        print(data.rd_x, data.rd_y)
    """
    
    def __init__(self):
        if requests is None:
            raise ImportError("requests library required. pip install requests")
    
    def search_address(self, city, street, housenumber):
        """
        Zoek locatie op basis van adres.
        
        Args:
            city: Plaatsnaam
            street: Straatnaam
            housenumber: Huisnummer
            
        Returns:
            LocationData object of None bij geen resultaat
        """
        query = "{0} and {1} and {2}".format(city, street, housenumber)
        url = "{0}?wt=json&q={1}".format(PDOKEndpoints.LOCATIE_FREE, query)
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            if data.get("response", {}).get("numFound", 0) == 0:
                return None
            
            doc = data["response"]["docs"][0]
            return self._parse_location_doc(doc, url)
            
        except Exception as e:
            print("Error searching address: {0}".format(e))
            return None
    
    def search_postcode(self, postcode, housenumber):
        """
        Zoek locatie op basis van postcode en huisnummer.
        
        Args:
            postcode: Postcode (bijv. "1012JS")
            housenumber: Huisnummer
            
        Returns:
            LocationData object of None bij geen resultaat
        """
        query = "{0} and {1}".format(postcode, housenumber)
        url = "{0}?wt=json&q={1}".format(PDOKEndpoints.LOCATIE_FREE, query)
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            if data.get("response", {}).get("numFound", 0) == 0:
                return None
            
            doc = data["response"]["docs"][0]
            return self._parse_location_doc(doc, url)
            
        except Exception as e:
            print("Error searching postcode: {0}".format(e))
            return None
    
    def _parse_location_doc(self, doc, url):
        """Parse PDOK response document naar LocationData."""
        # RD coördinaten
        centroid = doc.get("centroide_rd", "POINT(0 0)")
        rd_coords = centroid.replace("POINT(", "").replace(")", "").split()
        rd_x = float(rd_coords[0])
        rd_y = float(rd_coords[1])
        
        # WGS84 coördinaten
        centroid_ll = doc.get("centroide_ll", "POINT(0 0)")
        ll_coords = centroid_ll.replace("POINT(", "").replace(")", "").split()
        lon = float(ll_coords[0])
        lat = float(ll_coords[1])
        
        # Kadaster info
        kadaster_raw = doc.get("gekoppeld_perceel", ["--"])
        if kadaster_raw:
            kadaster = kadaster_raw[0].split("-") if kadaster_raw[0] else ["", "", ""]
        else:
            kadaster = ["", "", ""]
        
        # Provincie en windgebied
        provincie = doc.get("provincienaam", "")
        windgebied = WINDGEBIED_MAPPING.get(provincie, 3)
        if isinstance(windgebied, str):
            # Noord-Holland: bepaal obv latitude
            windgebied = 1 if lat > 52.5 else 2
        
        # Straatnaam uit verschillende velden proberen
        straatnaam = doc.get("straatnaam", "")
        if not straatnaam:
            straatnaam = doc.get("openbareruimtenaam", "")
        if not straatnaam:
            # Probeer uit weergavenaam te halen (bijv. "Damstraat 1, 1012JS Amsterdam")
            weergave = doc.get("weergavenaam", "")
            if weergave and "," in weergave:
                straatnaam = weergave.split(",")[0].rsplit(" ", 1)[0]  # Verwijder huisnummer
        
        # Huisnummer
        huisnummer = doc.get("huisnummer", "")
        huisletter = doc.get("huisletter", "")
        huisnummertoevoeging = doc.get("huisnummertoevoeging", "")
        
        volledig_huisnummer = str(huisnummer)
        if huisletter:
            volledig_huisnummer += huisletter
        if huisnummertoevoeging:
            volledig_huisnummer += "-" + str(huisnummertoevoeging)
        
        result = LocationData(
            rd_x=rd_x,
            rd_y=rd_y,
            lat=lat,
            lon=lon,
            postcode=doc.get("postcode", ""),
            gemeente=doc.get("gemeentenaam", ""),
            provincie=provincie,
            waterschap=doc.get("waterschapsnaam", ""),
            kadaster_gemeente=kadaster[0] if len(kadaster) > 0 else "",
            kadaster_sectie=kadaster[1] if len(kadaster) > 1 else "",
            kadaster_perceel=kadaster[2] if len(kadaster) > 2 else "",
            windgebied=windgebied,
            url=url,
        )
        
        # Extra velden toevoegen (niet in constructor)
        result.straatnaam = straatnaam
        result.huisnummer = volledig_huisnummer
        
        return result
    
    def suggest(self, query, rows=10):
        """
        Suggesties voor een zoekterm.
        
        Args:
            query: Zoekterm
            rows: Maximum aantal resultaten
            
        Returns:
            Lijst met suggesties
        """
        url = "{0}?q={1}&rows={2}".format(PDOKEndpoints.LOCATIE_SUGGEST, query, rows)
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()
            return data.get("response", {}).get("docs", [])
        except Exception as e:
            print("Error getting suggestions: {0}".format(e))
            return []
    
    def search_rd_coordinates(self, rd_x, rd_y, distance=100):
        """
        Reverse geocoding: zoek locatie op basis van RD coördinaten.
        
        Args:
            rd_x: RD X coördinaat
            rd_y: RD Y coördinaat
            distance: Zoekradius in meters (default: 100)
            
        Returns:
            LocationData object of None bij geen resultaat
        """
        print("PDOK reverse geocoding: RD {0}, {1}".format(rd_x, rd_y))
        
        # PDOK Locatieserver ondersteunt geometrie filter via fq parameter
        # We maken een bounding box filter rond het punt
        xmin = rd_x - distance
        ymin = rd_y - distance
        xmax = rd_x + distance
        ymax = rd_y + distance
        
        # Gebruik bq (boost query) met afstand tot punt voor sortering
        # en fq (filter query) met bbox voor filtering
        url = (
            "{0}?wt=json&rows=1&fl=*"
            "&fq=type:adres"
            "&fq=centroide_rd:[{1},{2} TO {3},{4}]"
        ).format(
            PDOKEndpoints.LOCATIE_FREE,
            xmin, ymin, xmax, ymax
        )
        
        print("PDOK URL: {0}".format(url))
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            num_found = data.get("response", {}).get("numFound", 0)
            print("PDOK response: {0} resultaten gevonden".format(num_found))
            
            if num_found == 0:
                # Probeer met grotere radius
                print("PDOK: geen resultaat, probeer grotere radius...")
                return self._search_by_point_expanded(rd_x, rd_y)
            
            doc = data["response"]["docs"][0]
            print("PDOK gevonden: {0}, {1}".format(
                doc.get("gemeentenaam", "?"), 
                doc.get("postcode", "?")
            ))
            
            result = self._parse_location_doc(doc, url)
            # Overschrijf coördinaten met de originele waarden
            result.rd_x = rd_x
            result.rd_y = rd_y
            return result
            
        except Exception as e:
            print("PDOK error in reverse geocoding: {0}".format(e))
            return None
    
    def _search_by_point_expanded(self, rd_x, rd_y):
        """
        Fallback: zoek met grotere radius (500m).
        """
        distance = 500
        xmin = rd_x - distance
        ymin = rd_y - distance
        xmax = rd_x + distance
        ymax = rd_y + distance
        
        url = (
            "{0}?wt=json&rows=1&fl=*"
            "&fq=type:adres"
            "&fq=centroide_rd:[{1},{2} TO {3},{4}]"
        ).format(
            PDOKEndpoints.LOCATIE_FREE,
            xmin, ymin, xmax, ymax
        )
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            if data.get("response", {}).get("numFound", 0) == 0:
                print("PDOK: ook geen resultaat met 500m radius")
                return None
            
            doc = data["response"]["docs"][0]
            result = self._parse_location_doc(doc, url)
            result.rd_x = rd_x
            result.rd_y = rd_y
            return result
            
        except Exception as e:
            print("PDOK error in expanded search: {0}".format(e))
            return None


class PDOKBGT:
    """
    Client voor PDOK BGT Download API.
    
    Example:
        bgt = PDOKBGT()
        files = bgt.download_bbox(155000, 463000, 156000, 464000, "output/")
    """
    
    def __init__(self, timeout=120):
        if requests is None:
            raise ImportError("requests library required")
        self.timeout = timeout
    
    def download_bbox(self, xmin, ymin, xmax, ymax, output_folder, 
                      features=None, format="gmllight"):
        """
        Download BGT data voor een bounding box.
        
        Args:
            xmin, ymin, xmax, ymax: Bounding box in RD coördinaten
            output_folder: Output folder voor ZIP bestand
            features: Lijst van feature types (default: alle)
            format: Output formaat ("gmllight", "citygml")
            
        Returns:
            Dict met download status en bestandslocaties
        """
        if features is None:
            features = BGT_FEATURES
        
        # Create polygon WKT
        polygon_wkt = (
            "POLYGON(({0} {1}, {2} {1}, "
            "{2} {3}, {0} {3}, {0} {1}))".format(xmin, ymin, xmax, ymax)
        )
        
        # Start download request
        payload = {
            "featuretypes": features,
            "format": format,
            "geofilter": polygon_wkt
        }
        
        headers = {
            "accept": "application/json",
            "Content-Type": "application/json"
        }
        
        try:
            # POST request to start download
            response = requests.post(
                PDOKEndpoints.BGT_DOWNLOAD,
                headers=headers,
                json=payload,
                timeout=30
            )
            response.raise_for_status()
            download_id = response.json()["downloadRequestId"]
            
            # Poll for completion
            status_url = PDOKEndpoints.BGT_STATUS.format(id=download_id)
            download_url = None
            
            start_time = time.time()
            while time.time() - start_time < self.timeout:
                status_resp = requests.get(status_url, timeout=10)
                status_data = status_resp.json()
                
                if status_data["status"] == "COMPLETED":
                    download_url = "https://api.pdok.nl" + status_data["_links"]["download"]["href"]
                    break
                elif status_data["status"] == "ERROR":
                    return {"status": "error", "message": status_data.get("message", "Unknown error")}
                
                time.sleep(2)
            
            if download_url is None:
                return {"status": "timeout", "message": "Download timed out"}
            
            # Download ZIP file
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            zip_path = os.path.join(output_folder, "bgt_{0}.zip".format(download_id))
            
            zip_response = requests.get(download_url, timeout=60)
            with open(zip_path, "wb") as f:
                f.write(zip_response.content)
            
            return {
                "status": "success",
                "download_id": download_id,
                "zip_file": zip_path,
                "features": features
            }
            
        except Exception as e:
            return {"status": "error", "message": str(e)}


# =============================================================================
# Kadastrale Kaart WFS
# =============================================================================


class PerceelData:
    """Container voor perceelgegevens."""
    
    def __init__(self, geometry, perceelnummer=None, sectie=None, 
                 gemeentecode=None, oppervlakte=None):
        self.geometry = geometry  # List of (x, y) tuples
        self.perceelnummer = perceelnummer
        self.sectie = sectie
        self.gemeentecode = gemeentecode
        self.oppervlakte = oppervlakte


class PerceelAnnotatie:
    """Container voor perceelnummer annotatie met positie en rotatie."""
    
    def __init__(self, tekst, x, y, rotatie=0.0, delta_x=0.0, delta_y=0.0):
        self.tekst = tekst
        self.x = x  # Plaatscoordinaat X
        self.y = y  # Plaatscoordinaat Y
        self.rotatie = rotatie  # In graden
        self.delta_x = delta_x  # Verschuiving
        self.delta_y = delta_y  # Verschuiving


class OpenbareRuimteNaam:
    """Container voor straatnaam met positie en rotatie."""
    
    def __init__(self, tekst, x, y, rotatie=0.0):
        self.tekst = tekst
        self.x = x
        self.y = y
        self.rotatie = rotatie


class Nummeraanduiding:
    """Container voor huisnummer met positie en rotatie."""
    
    def __init__(self, huisnummer, huisletter, toevoeging, x, y, rotatie=0.0):
        self.huisnummer = huisnummer
        self.huisletter = huisletter or ""
        self.toevoeging = toevoeging or ""
        self.x = x
        self.y = y
        self.rotatie = rotatie
    
    @property
    def volledig(self):
        """Geef volledig huisnummer terug."""
        result = str(self.huisnummer)
        if self.huisletter:
            result += self.huisletter
        if self.toevoeging:
            result += "-" + str(self.toevoeging)
        return result


class PDOKKadaster:
    """
    Client voor PDOK Kadastrale Kaart WFS API v5.
    
    Haalt perceelgrenzen, perceelnummers, straatnamen en huisnummers op
    via de WFS 2.0 interface.
    
    Example:
        kad = PDOKKadaster()
        percelen = kad.get_percelen(155000, 463000, 155500, 463500)
        annotaties = kad.get_perceel_annotaties(155000, 463000, 155500, 463500)
    """
    
    WFS_URL = "https://service.pdok.nl/kadaster/kadastralekaart/wfs/v5_0"
    
    # Layer names - NOTE: niet alle layers bestaan in v5
    LAYER_PERCEEL = "kadastralekaartv5:Perceel"
    LAYER_OPENBARE_RUIMTE = "kadastralekaartv5:OpenbareRuimteNaam"
    # LAYER_NUMMERAANDUIDING bestaat NIET in v5 - gebruik BAG WFS
    
    def __init__(self):
        if requests is None:
            raise ImportError("requests library required. pip install requests")
    
    def _build_wfs_url(self, layer, bbox, max_features=10000):
        """Bouw WFS GetFeature URL."""
        # BBOX format: xmin,ymin,xmax,ymax,CRS (correct WFS 2.0 volgorde)
        return (
            "{base}?service=WFS&version=2.0.0&request=GetFeature"
            "&typeName={layer}"
            "&bbox={xmin},{ymin},{xmax},{ymax},EPSG:28992"
            "&count={count}"
            "&outputFormat=application/json"
        ).format(
            base=self.WFS_URL,
            layer=layer,
            xmin=bbox[0],
            ymin=bbox[1],
            xmax=bbox[2],
            ymax=bbox[3],
            count=max_features
        )
    
    def get_percelen(self, xmin, ymin, xmax, ymax):
        """
        Haal perceelgrenzen op binnen bounding box.
        
        Args:
            xmin, ymin, xmax, ymax: Bounding box in RD coördinaten
            
        Returns:
            Lijst van PerceelData objecten
        """
        url = self._build_wfs_url(self.LAYER_PERCEEL, (xmin, ymin, xmax, ymax))
        
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            percelen = []
            
            for feature in data.get("features", []):
                props = feature.get("properties", {})
                geom = feature.get("geometry", {})
                
                # Parse geometry - kan Polygon of MultiPolygon zijn
                coords = []
                if geom.get("type") == "Polygon":
                    # Eerste ring is outer boundary
                    coords = [(c[0], c[1]) for c in geom.get("coordinates", [[]])[0]]
                elif geom.get("type") == "MultiPolygon":
                    # Pak eerste polygon, eerste ring
                    multi = geom.get("coordinates", [[[]]])
                    if multi and multi[0]:
                        coords = [(c[0], c[1]) for c in multi[0][0]]
                
                if not coords:
                    continue
                
                percelen.append(PerceelData(
                    geometry=coords,
                    perceelnummer=props.get("perceelnummer"),
                    sectie=props.get("sectie"),
                    gemeentecode=props.get("AKRKadastraleGemeenteCode"),
                    oppervlakte=props.get("kadastraleGrootte")
                ))
            
            return percelen
            
        except Exception as e:
            print("Error fetching percelen: {0}".format(e))
            return []
    
    def get_perceel_annotaties(self, xmin, ymin, xmax, ymax):
        """
        Haal perceelnummer annotaties op met positie en rotatie.
        
        NOTE: Annotatie-data zit in de Perceel layer zelf, niet in aparte layer.
        
        Args:
            xmin, ymin, xmax, ymax: Bounding box in RD coördinaten
            
        Returns:
            Lijst van PerceelAnnotatie objecten
        """
        # Annotatie-data zit in de Perceel layer
        url = self._build_wfs_url(self.LAYER_PERCEEL, (xmin, ymin, xmax, ymax))
        
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            annotaties = []
            
            for feature in data.get("features", []):
                props = feature.get("properties", {})
                
                tekst = props.get("perceelnummer")
                if not tekst:
                    continue
                
                # Annotatie positie is direct in properties
                x = props.get("perceelnummerPlaatscoordinaatX")
                y = props.get("perceelnummerPlaatscoordinaatY")
                
                if x is None or y is None:
                    continue
                
                annotaties.append(PerceelAnnotatie(
                    tekst=str(tekst),
                    x=float(x),
                    y=float(y),
                    rotatie=float(props.get("perceelnummerRotatie", 0)),
                    delta_x=float(props.get("perceelnummerVerschuivingDeltaX", 0)),
                    delta_y=float(props.get("perceelnummerVerschuivingDeltaY", 0)),
                ))
            
            return annotaties
            
        except Exception as e:
            print("Error fetching perceel annotaties: {0}".format(e))
            return []
    
    def get_straatnamen(self, xmin, ymin, xmax, ymax):
        """
        Haal straatnamen op met positie en rotatie.
        
        Args:
            xmin, ymin, xmax, ymax: Bounding box in RD coördinaten
            
        Returns:
            Lijst van OpenbareRuimteNaam objecten
        """
        url = self._build_wfs_url(self.LAYER_OPENBARE_RUIMTE, (xmin, ymin, xmax, ymax))
        
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            straatnamen = []
            
            for feature in data.get("features", []):
                props = feature.get("properties", {})
                geom = feature.get("geometry", {})
                
                # Veldnaam is 'tekst', niet 'openbareRuimteNaam'
                tekst = props.get("tekst")
                if not tekst:
                    continue
                
                # Coördinaten uit Point geometry
                coords = geom.get("coordinates", [0, 0])
                x = coords[0] if len(coords) > 0 else 0
                y = coords[1] if len(coords) > 1 else 0
                
                straatnamen.append(OpenbareRuimteNaam(
                    tekst=tekst,
                    x=float(x),
                    y=float(y),
                    rotatie=float(props.get("hoek", 0)),
                ))
            
            return straatnamen
            
        except Exception as e:
            print("Error fetching straatnamen: {0}".format(e))
            return []
    
    def get_huisnummers(self, xmin, ymin, xmax, ymax):
        """
        Haal huisnummers op met positie en rotatie.
        
        NOTE: Nummeraanduiding layer bestaat NIET in Kadastrale Kaart WFS v5.
        Huisnummers moeten via BAG WFS worden opgehaald.
        
        Args:
            xmin, ymin, xmax, ymax: Bounding box in RD coördinaten
            
        Returns:
            Lege lijst (layer niet beschikbaar in v5)
        """
        print("WARNING: Nummeraanduiding layer niet beschikbaar in Kadastrale Kaart v5")
        print("         Gebruik BAG WFS voor huisnummers")
        return []
    
    def get_all_data(self, xmin, ymin, xmax, ymax):
        """
        Haal alle kadaster data op in één keer.
        
        Returns:
            Dict met percelen, annotaties, straatnamen, huisnummers
        """
        return {
            "percelen": self.get_percelen(xmin, ymin, xmax, ymax),
            "annotaties": self.get_perceel_annotaties(xmin, ymin, xmax, ymax),
            "straatnamen": self.get_straatnamen(xmin, ymin, xmax, ymax),
            "huisnummers": self.get_huisnummers(xmin, ymin, xmax, ymax),
        }


class PDOKWMTS:
    """
    Client voor PDOK WMTS services.
    
    Example:
        wmts = PDOKWMTS()
        image = wmts.download_aerial_image(155000, 463000, 500, 500, "output.png")
    """
    
    LAYERS = {
        "luchtfoto_actueel": "Actueel_orthoHR",
        "luchtfoto_2022": "2022_orthoHR",
        "luchtfoto_2021": "2021_orthoHR",
        "luchtfoto_2020": "2020_orthoHR",
        "luchtfoto_2019": "2019_orthoHR",
        "luchtfoto_2018": "2018_orthoHR",
        "luchtfoto_2017": "2017_orthoHR",
        "luchtfoto_2016": "2016_orthoHR",
    }
    
    def __init__(self):
        if requests is None:
            raise ImportError("requests library required")
    
    def get_aerial_image_url(self, xmin, ymin, xmax, ymax, 
                              width=2000, height=2000, layer="luchtfoto_actueel"):
        """
        Genereer URL voor luchtfoto WMS request.
        """
        layer_name = self.LAYERS.get(layer, self.LAYERS["luchtfoto_actueel"])
        
        return (
            "https://service.pdok.nl/hwh/luchtfotorgb/wms/v1_0"
            "?&request=GetMap&VERSION=1.3.0&STYLES="
            "&layers={0}"
            "&width={1}&height={2}"
            "&format=image/png&crs=EPSG:28992"
            "&bbox={3},{4},{5},{6}".format(
                layer_name, width, height, xmin, ymin, xmax, ymax
            )
        )
    
    def download_aerial_image(self, xmin, ymin, xmax, ymax, output_path,
                               width=2000, height=2000, layer="luchtfoto_actueel"):
        """
        Download luchtfoto naar bestand.
        
        Returns:
            True bij succes
        """
        url = self.get_aerial_image_url(xmin, ymin, xmax, ymax, width, height, layer)
        
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            folder = os.path.dirname(output_path)
            if folder and not os.path.exists(folder):
                os.makedirs(folder)
            
            with open(output_path, "wb") as f:
                f.write(response.content)
            return True
            
        except Exception as e:
            print("Error downloading image: {0}".format(e))
            return False
