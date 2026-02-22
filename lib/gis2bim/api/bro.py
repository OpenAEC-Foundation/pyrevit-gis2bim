# -*- coding: utf-8 -*-
"""
BRO (Basisregistratie Ondergrond) API Client
=============================================

Client voor het ophalen van geotechnische data uit de BRO:
- CPT (Cone Penetration Test / Sondering)
- BHR-GT (Borehole Research Geotechnical / Boring)

Data wordt opgehaald via de publieke BRO REST API.
Zoeken gaat via JSON met WGS84 bounding box.
Responses zijn XML.

Gebruik:
    from gis2bim.api.bro import BROClient

    client = BROClient()
    cpt_list = client.search_cpt(155000, 463000, 500)
    for cpt in cpt_list:
        print(cpt.bro_id, cpt.einddiepte)

    bhr_ids = client.search_bhr(155000, 463000, 500)
    for bro_id in bhr_ids:
        bhr = client.get_bhr(bro_id)
        for laag in bhr.grondlagen:
            print(laag.grondsoort, laag.dikte)
"""

import json
import datetime
import xml.etree.ElementTree as ET

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2

from gis2bim.coordinates import rd_to_wgs84


# =============================================================================
# Data klassen
# =============================================================================

class CPTData(object):
    """Container voor CPT (sondering) data.

    Attributes:
        bro_id: BRO identificatie (bijv. CPT000000012345)
        rd_x: RD X coordinaat in meters
        rd_y: RD Y coordinaat in meters
        nap_hoogte: Maaiveld hoogte in meters t.o.v. NAP
        einddiepte: Diepte van de sondering in meters onder maaiveld
    """

    def __init__(self, bro_id, rd_x, rd_y, nap_hoogte=0.0, einddiepte=0.0,
                 metingen=None):
        self.bro_id = bro_id
        self.rd_x = rd_x
        self.rd_y = rd_y
        self.nap_hoogte = nap_hoogte
        self.einddiepte = einddiepte
        self.metingen = metingen or []

    def __repr__(self):
        return "CPTData('{0}', diepte={1:.1f}m)".format(
            self.bro_id, self.einddiepte)


class Grondlaag(object):
    """Container voor een grondlaag in een boring.

    Attributes:
        bovenkant: Bovenkant in meters onder maaiveld
        onderkant: Onderkant in meters onder maaiveld
        grondsoort: Geclassificeerde grondsoort (klei, zand, etc.)
        beschrijving: Originele beschrijving uit BRO
    """

    def __init__(self, bovenkant, onderkant, grondsoort="onbekend",
                 beschrijving=""):
        self.bovenkant = bovenkant
        self.onderkant = onderkant
        self.grondsoort = grondsoort
        self.beschrijving = beschrijving

    @property
    def dikte(self):
        """Dikte van de laag in meters."""
        return self.onderkant - self.bovenkant

    def __repr__(self):
        return "Grondlaag({0:.1f}-{1:.1f}m, {2})".format(
            self.bovenkant, self.onderkant, self.grondsoort)


class BHRData(object):
    """Container voor BHR-GT (boring) data.

    Attributes:
        bro_id: BRO identificatie (bijv. BHR000000012345)
        rd_x: RD X coordinaat in meters
        rd_y: RD Y coordinaat in meters
        nap_hoogte: Maaiveld hoogte in meters t.o.v. NAP
        grondlagen: Lijst van Grondlaag objecten
    """

    def __init__(self, bro_id, rd_x, rd_y, nap_hoogte=0.0, grondlagen=None):
        self.bro_id = bro_id
        self.rd_x = rd_x
        self.rd_y = rd_y
        self.nap_hoogte = nap_hoogte
        self.grondlagen = grondlagen or []

    @property
    def einddiepte(self):
        """Totale diepte in meters onder maaiveld."""
        if not self.grondlagen:
            return 0.0
        return max(laag.onderkant for laag in self.grondlagen)

    def __repr__(self):
        return "BHRData('{0}', {1} lagen, {2:.1f}m diep)".format(
            self.bro_id, len(self.grondlagen), self.einddiepte)


# =============================================================================
# Grondsoort kleuren (NEN 5104 / ISO 14688)
# =============================================================================

GRONDSOORT_KLEUREN = {
    "klei":     (100, 180, 100),   # Groen
    "zand":     (230, 210, 120),   # Geel
    "grind":    (220, 165, 80),    # Oranje
    "veen":     (100, 70, 40),     # Donkerbruin
    "leem":     (190, 170, 130),   # Lichtbruin
    "silt":     (190, 170, 130),   # Lichtbruin (zelfde als leem)
    "onbekend": (180, 180, 180),   # Grijs
}

# CPT sondering kleur (uniform, geen grondlagen)
CPT_KLEUR = (70, 130, 180)        # Staalblauw

# CPT qc kleurenschema voor geotechnische beoordeling
QC_KLEUREN = {
    "zeer_slap":    (80, 220, 60),    # Felgroen - Veen/slappe klei
    "slap":         (160, 210, 50),   # Geelgroen - Klei
    "matig_stevig": (240, 220, 50),   # Geel - Zandige klei/silt
    "stevig":       (240, 165, 40),   # Oranje - Fijn zand
    "vast":         (220, 120, 30),   # Donkeroranje - Grof zand
    "hard":         (210, 50, 40),    # Rood - Draagkrachtig zand/grind
}


def classificeer_qc(qc_mpa):
    """Classificeer conusweerstand naar grondtype.

    Args:
        qc_mpa: Conusweerstand in MPa

    Returns:
        Classificatie string (zeer_slap, slap, matig_stevig, stevig, vast, hard)
    """
    if qc_mpa < 1:
        return "zeer_slap"
    elif qc_mpa < 2:
        return "slap"
    elif qc_mpa < 5:
        return "matig_stevig"
    elif qc_mpa < 10:
        return "stevig"
    elif qc_mpa < 15:
        return "vast"
    else:
        return "hard"


def get_qc_kleur(qc_mpa):
    """Haal RGB kleur op voor een conusweerstand waarde.

    Args:
        qc_mpa: Conusweerstand in MPa

    Returns:
        Tuple (R, G, B) met kleurwaarden 0-255
    """
    return QC_KLEUREN[classificeer_qc(qc_mpa)]


class CPTMeting(object):
    """Container voor een geaggregeerd CPT meetsegment.

    Attributes:
        bovenkant: Bovenkant in meters onder maaiveld
        onderkant: Onderkant in meters onder maaiveld
        qc: Gemiddelde conusweerstand (MPa)
        fs: Gemiddelde wrijvingsweerstand (MPa)
        rf: Gemiddelde wrijvingsgetal (%)
        classificatie: Auto-berekend uit qc
    """

    def __init__(self, bovenkant, onderkant, qc=0.0, fs=0.0, rf=0.0):
        self.bovenkant = bovenkant
        self.onderkant = onderkant
        self.qc = qc
        self.fs = fs
        self.rf = rf
        self.classificatie = classificeer_qc(qc)

    @property
    def dikte(self):
        """Dikte van het segment in meters."""
        return self.onderkant - self.bovenkant

    def __repr__(self):
        return "CPTMeting({0:.1f}-{1:.1f}m, qc={2:.1f} MPa, {3})".format(
            self.bovenkant, self.onderkant, self.qc, self.classificatie)


# Zoekwoorden voor classificatie
_GRONDSOORT_KEYWORDS = [
    ("klei", ["klei", "clay"]),
    ("zand", ["zand", "sand"]),
    ("grind", ["grind", "gravel"]),
    ("veen", ["veen", "peat", "turf"]),
    ("leem", ["leem", "loam"]),
    ("silt", ["silt"]),
]

# NEN5104 samengestelde namen naar hoofdgrondsoort
_NEN5104_MAPPING = {
    "klei": "klei",
    "zand": "zand",
    "grind": "grind",
    "veen": "veen",
    "leem": "leem",
    "silt": "silt",
    # Samengesteld: hoofdgrondsoort staat achteraan (NEN5104 conventie)
    "siltigzand": "zand",
    "kleiigzand": "zand",
    "grindigzand": "zand",
    "zandigeklei": "klei",
    "siltigeklei": "klei",
    "venigeklei": "klei",
    "zandigsilt": "silt",
    "kleiigveen": "veen",
    "zandigveen": "veen",
    "siltigleem": "leem",
    # Met bijvoeglijke sterkte
    "zwaksiltigzand": "zand",
    "sterksiltigzand": "zand",
    "zwakzandigeklei": "klei",
    "sterkzandigeklei": "klei",
    "zwakzandigsilt": "silt",
    "sterkzandigsilt": "silt",
    "zwakkleiigzand": "zand",
    "sterkkleiigzand": "zand",
    "zwakvenigeklei": "klei",
    "sterkvenigeklei": "klei",
    "zwakkleiigveen": "veen",
    "sterkkleiigveen": "veen",
    "zwakzandigveen": "veen",
    "sterkzandigveen": "veen",
    "humeusleem": "leem",
    "humeuszand": "zand",
    "humeusklei": "klei",
}


def classificeer_grondsoort(beschrijving):
    """Classificeer grondsoort op basis van beschrijving.

    Ondersteunt zowel vrije tekst als NEN5104 samengestelde namen
    (bijv. 'zwakSiltigZand' -> 'zand').

    Args:
        beschrijving: Tekst beschrijving of NEN5104 naam

    Returns:
        Geclassificeerde grondsoort string (klei, zand, etc.)
    """
    if not beschrijving:
        return "onbekend"

    # Probeer eerst NEN5104 mapping (case-insensitive, zonder spaties)
    key = beschrijving.lower().replace(" ", "")
    if key in _NEN5104_MAPPING:
        return _NEN5104_MAPPING[key]

    # Fallback: zoek keywords in de tekst
    tekst = beschrijving.lower()
    for soort, keywords in _GRONDSOORT_KEYWORDS:
        for kw in keywords:
            if kw in tekst:
                return soort

    return "onbekend"


def get_grondsoort_kleur(grondsoort):
    """Haal RGB kleur op voor een grondsoort.

    Args:
        grondsoort: Geclassificeerde grondsoort string

    Returns:
        Tuple (R, G, B) met kleurwaarden 0-255
    """
    return GRONDSOORT_KLEUREN.get(grondsoort, GRONDSOORT_KLEUREN["onbekend"])


# =============================================================================
# XML Helpers
# =============================================================================

def _strip_ns(tag):
    """Strip XML namespace van een tag."""
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def _find_recursive(element, local_name):
    """Zoek element recursief op local name (zonder namespace)."""
    for el in element.iter():
        if _strip_ns(el.tag) == local_name:
            return el
    return None


def _findall_recursive(element, local_name):
    """Zoek alle elementen recursief op local name."""
    return [el for el in element.iter() if _strip_ns(el.tag) == local_name]


def _find_text(element, local_name, default=""):
    """Zoek element en retourneer text content."""
    el = _find_recursive(element, local_name)
    if el is not None and el.text:
        return el.text.strip()
    return default


def _find_float(element, local_name, default=0.0):
    """Zoek element en retourneer float waarde."""
    text = _find_text(element, local_name)
    if text:
        try:
            return float(text)
        except (ValueError, TypeError):
            pass
    return default


# =============================================================================
# BRO API Client
# =============================================================================

# API endpoints (productie)
BRO_CPT_SEARCH_URL = (
    "https://publiek.broservices.nl/sr/cpt/v1/characteristics/searches"
)
BRO_BHR_SEARCH_URL = (
    "https://publiek.broservices.nl/sr/bhrgt/v2/characteristics/searches"
)
BRO_BHR_DETAIL_URL = (
    "https://publiek.broservices.nl/sr/bhrgt/v2/objects"
)
BRO_CPT_DETAIL_URL = (
    "https://publiek.broservices.nl/sr/cpt/v1/objects"
)
BRO_VIEWER_URL = (
    "https://www.broloket.nl/ondergrondgegevens/"
    "registratieobject?registratie={bro_id}"
)


def _rd_to_wgs84_bbox(rd_x, rd_y, radius):
    """Converteer RD centrum + straal naar WGS84 bounding box.

    Args:
        rd_x: RD X coordinaat van centrum
        rd_y: RD Y coordinaat van centrum
        radius: Straal in meters

    Returns:
        Dict met lowerCorner en upperCorner in WGS84
    """
    lat_sw, lon_sw = rd_to_wgs84(rd_x - radius, rd_y - radius)
    lat_ne, lon_ne = rd_to_wgs84(rd_x + radius, rd_y + radius)

    return {
        "lowerCorner": {"lat": lat_sw, "lon": lon_sw},
        "upperCorner": {"lat": lat_ne, "lon": lon_ne}
    }


def _build_search_json(rd_x, rd_y, radius):
    """Bouw JSON body voor BRO zoek-request.

    Args:
        rd_x: RD X coordinaat van centrum
        rd_y: RD Y coordinaat van centrum
        radius: Straal in meters

    Returns:
        JSON string als bytes
    """
    bbox = _rd_to_wgs84_bbox(rd_x, rd_y, radius)

    today = datetime.date.today().isoformat()

    body = {
        "registrationPeriod": {
            "beginDate": "2000-01-01",
            "endDate": today
        },
        "area": {
            "boundingBox": bbox
        }
    }

    return json.dumps(body).encode("utf-8")


class BROClient(object):
    """Client voor de BRO (Basisregistratie Ondergrond) API.

    Haalt CPT (sondering) en BHR-GT (boring) data op via de
    publieke BRO REST API.

    CPT data wordt volledig uit het zoekresultaat gehaald (1 API call).
    BHR-GT grondlagen vereisen een extra detail-call per boring.

    Gebruik:
        client = BROClient()
        cpt_list = client.search_cpt(155000, 463000, 500)
        bhr_ids = client.search_bhr(155000, 463000, 500)
        bhr = client.get_bhr(bhr_ids[0])
    """

    def __init__(self, timeout=30):
        self.timeout = timeout

    # ----- CPT -----

    def search_cpt(self, rd_x, rd_y, radius):
        """Zoek CPT registraties en retourneer volledige data.

        De zoekresultaten bevatten alle benodigde velden (locatie,
        NAP hoogte, einddiepte), dus geen extra detail-calls nodig.

        Args:
            rd_x: RD X coordinaat van centrum
            rd_y: RD Y coordinaat van centrum
            radius: Zoekstraal in meters

        Returns:
            Lijst van CPTData objecten
        """
        root = self._post_search(BRO_CPT_SEARCH_URL, rd_x, rd_y, radius)
        if root is None:
            return []

        # Controleer op rejection
        response_type = _find_text(root, "responseType")
        if response_type == "rejection":
            reason = _find_text(root, "rejectionReason")
            print("BRO CPT search rejected: {0}".format(reason))
            return []

        results = []
        for doc in _findall_recursive(root, "dispatchDocument"):
            cpt = self._parse_cpt_from_search(doc)
            if cpt:
                results.append(cpt)

        return results

    # ----- BHR-GT -----

    def search_bhr(self, rd_x, rd_y, radius):
        """Zoek BHR-GT registraties binnen een bounding box.

        Args:
            rd_x: RD X coordinaat van centrum
            rd_y: RD Y coordinaat van centrum
            radius: Zoekstraal in meters

        Returns:
            Lijst van BRO-ID strings
        """
        root = self._post_search(BRO_BHR_SEARCH_URL, rd_x, rd_y, radius)
        if root is None:
            return []

        # Controleer op rejection
        response_type = _find_text(root, "responseType")
        if response_type == "rejection":
            reason = _find_text(root, "rejectionReason")
            print("BRO BHR search rejected: {0}".format(reason))
            return []

        bro_ids = []
        for el in _findall_recursive(root, "broId"):
            if el.text and el.text.strip():
                bro_id = el.text.strip()
                if bro_id not in bro_ids:
                    bro_ids.append(bro_id)

        return bro_ids

    def get_bhr(self, bro_id):
        """Haal BHR-GT detail data op inclusief grondlagen.

        Args:
            bro_id: BRO identificatie string

        Returns:
            BHRData object of None bij fout
        """
        url = "{0}/{1}".format(BRO_BHR_DETAIL_URL, bro_id)
        xml_data = self._get_xml(url)
        if xml_data is None:
            return None

        return self._parse_bhr(xml_data, bro_id)

    # ----- CPT detail -----

    def get_cpt(self, bro_id):
        """Haal CPT detail data op inclusief meetdata.

        Args:
            bro_id: BRO identificatie string

        Returns:
            CPTData object met metingen, of None bij fout
        """
        url = "{0}/{1}".format(BRO_CPT_DETAIL_URL, bro_id)
        print("CPT detail ophalen: {0}".format(url))
        xml_data = self._get_xml(url)
        if xml_data is None:
            print("CPT detail: geen XML ontvangen voor {0}".format(bro_id))
            return None

        print("CPT detail: XML ontvangen voor {0}".format(bro_id))
        result = self._parse_cpt_detail(xml_data, bro_id)
        if result:
            print("CPT detail: {0} metingen voor {1}".format(
                len(result.metingen), bro_id))
        return result

    def _parse_cpt_detail(self, root, bro_id):
        """Parse CPT XML detail response naar CPTData met metingen."""
        try:
            rd_x, rd_y = self._parse_rd_location(root)
            nap_hoogte = self._parse_nap_hoogte(root)

            einddiepte = _find_float(root, "finalDepth")
            if einddiepte <= 0:
                einddiepte = _find_float(root, "penetrationLength")

            metingen = self._parse_cpt_metingen(root)

            return CPTData(
                bro_id=bro_id,
                rd_x=rd_x,
                rd_y=rd_y,
                nap_hoogte=nap_hoogte,
                einddiepte=einddiepte,
                metingen=metingen
            )

        except Exception as e:
            print("CPT detail parse error ({0}): {1}".format(bro_id, e))
            return None

    def _parse_cpt_metingen(self, root, segment_size=1.0):
        """Parse CPT meetdata uit XML en aggregeer per segment.

        BRO CPT XML bevat meetdata als CSV in <swe:values>.
        Rij-separator: ;  Kolom-separator: ,  No-data: -999999
        Kolommen (0-indexed):
            [0] penetrationLength (m)
            [1] depth (m)
            [3] coneResistance/qc (MPa)
            [18] localFriction/fs (MPa)
            [24] frictionRatio/Rf (%)

        Args:
            root: XML root element
            segment_size: Grootte van aggregatie-interval in meters

        Returns:
            Lijst van CPTMeting objecten
        """
        # Zoek de values element (CSV data)
        values_el = _find_recursive(root, "values")
        if values_el is None:
            print("CPT metingen: 'values' element niet gevonden")
            return []
        if not values_el.text:
            print("CPT metingen: 'values' element heeft geen tekst")
            return []

        csv_text = values_el.text.strip()
        if not csv_text:
            print("CPT metingen: lege CSV tekst")
            return []

        print("CPT metingen: CSV lengte={0} chars".format(len(csv_text)))

        # Parse CSV naar meetpunten
        meetpunten = []  # (depth, qc, fs, rf)
        skipped = 0
        for row in csv_text.split(";"):
            row = row.strip()
            if not row:
                continue
            cols = row.split(",")
            if len(cols) < 4:
                skipped += 1
                continue

            try:
                depth = float(cols[1])
                pen_length = float(cols[0])
                qc = float(cols[3])
                fs = float(cols[18]) if len(cols) > 18 else -999999.0
                rf = float(cols[24]) if len(cols) > 24 else -999999.0
            except (ValueError, IndexError):
                skipped += 1
                continue

            # Fallback: als depth no-data, gebruik penetrationLength
            if depth == -999999.0:
                depth = pen_length

            # Skip no-data waarden
            if depth == -999999.0 or qc == -999999.0:
                continue
            if fs == -999999.0:
                fs = 0.0
            if rf == -999999.0:
                rf = 0.0

            meetpunten.append((depth, qc, fs, rf))

        print("CPT metingen: {0} meetpunten geparsed, {1} overgeslagen".format(
            len(meetpunten), skipped))

        if not meetpunten:
            return []

        # Sorteer op diepte
        meetpunten.sort(key=lambda m: m[0])

        # Aggregeer per segment_size interval
        metingen = []
        max_depth = meetpunten[-1][0]
        seg_start = 0.0

        while seg_start < max_depth:
            seg_end = seg_start + segment_size

            # Verzamel meetpunten in dit segment
            seg_qc = []
            seg_fs = []
            seg_rf = []
            for depth, qc, fs, rf in meetpunten:
                if seg_start <= depth < seg_end:
                    seg_qc.append(qc)
                    seg_fs.append(fs)
                    seg_rf.append(rf)

            if seg_qc:
                avg_qc = sum(seg_qc) / len(seg_qc)
                avg_fs = sum(seg_fs) / len(seg_fs)
                avg_rf = sum(seg_rf) / len(seg_rf)

                metingen.append(CPTMeting(
                    bovenkant=seg_start,
                    onderkant=seg_end,
                    qc=avg_qc,
                    fs=avg_fs,
                    rf=avg_rf
                ))

            seg_start = seg_end

        print("CPT metingen: {0} segmenten geaggregeerd".format(len(metingen)))
        return metingen

    # ----- Interne methoden -----

    def _post_search(self, url, rd_x, rd_y, radius):
        """Voer JSON zoek-request uit, retourneer XML root element."""
        body = _build_search_json(rd_x, rd_y, radius)

        try:
            request = urllib2.Request(url, data=body)
            request.add_header("Content-Type", "application/json")
            request.add_header("Accept", "application/xml")
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

            try:
                response = urllib2.urlopen(request, timeout=self.timeout)
            except Exception as e:
                # BRO API retourneert soms HTTP 400 met valide XML body
                if hasattr(e, 'read'):
                    xml_text = e.read().decode("utf-8")
                    try:
                        return ET.fromstring(xml_text)
                    except ET.ParseError:
                        pass
                raise

            xml_text = response.read().decode("utf-8")
            return ET.fromstring(xml_text)

        except Exception as e:
            print("BRO search error ({0}): {1}".format(url, e))
            return None

    def _get_xml(self, url):
        """Haal XML data op via GET request."""
        try:
            request = urllib2.Request(url)
            request.add_header("Accept", "application/xml")
            request.add_header("User-Agent", "GIS2BIM-pyRevit/1.0")

            response = urllib2.urlopen(request, timeout=self.timeout)
            xml_text = response.read().decode("utf-8")

            return ET.fromstring(xml_text)

        except Exception as e:
            print("BRO GET error ({0}): {1}".format(url, e))
            return None

    def _parse_cpt_from_search(self, doc_element):
        """Parse CPT data uit een zoek-response dispatchDocument.

        Het zoekresultaat bevat: broId, deliveredLocation (RD),
        offset (NAP hoogte), en finalDepth.
        """
        try:
            bro_id = _find_text(doc_element, "broId")
            if not bro_id:
                return None

            rd_x, rd_y = self._parse_rd_location(doc_element)
            nap_hoogte = _find_float(doc_element, "offset")

            einddiepte = _find_float(doc_element, "finalDepth")
            if einddiepte <= 0:
                einddiepte = _find_float(doc_element, "penetrationLength")

            if einddiepte <= 0:
                return None

            return CPTData(
                bro_id=bro_id,
                rd_x=rd_x,
                rd_y=rd_y,
                nap_hoogte=nap_hoogte,
                einddiepte=einddiepte
            )

        except Exception as e:
            print("CPT parse error: {0}".format(e))
            return None

    def _parse_bhr(self, root, bro_id):
        """Parse BHR-GT XML detail response naar BHRData object."""
        try:
            rd_x, rd_y = self._parse_rd_location(root)
            nap_hoogte = self._parse_nap_hoogte(root)
            grondlagen = self._parse_grondlagen(root)

            return BHRData(
                bro_id=bro_id,
                rd_x=rd_x,
                rd_y=rd_y,
                nap_hoogte=nap_hoogte,
                grondlagen=grondlagen
            )

        except Exception as e:
            print("BHR parse error ({0}): {1}".format(bro_id, e))
            return None

    def _parse_rd_location(self, element):
        """Parse RD locatie uit BRO XML element.

        Zoekt pos elementen en retourneert de eerste die RD coords bevat
        (waarden > 1000 om WGS84 lat/lon uit te sluiten).
        """
        for pos_el in _findall_recursive(element, "pos"):
            if pos_el.text:
                parts = pos_el.text.strip().split()
                if len(parts) >= 2:
                    try:
                        x = float(parts[0])
                        y = float(parts[1])
                        if x > 1000 and y > 1000:
                            return (x, y)
                    except (ValueError, TypeError):
                        continue
        return (0.0, 0.0)

    def _parse_nap_hoogte(self, root):
        """Parse NAP hoogte (maaiveld) uit BRO XML response."""
        # Zoek offset in deliveredVerticalPosition context
        vp_el = _find_recursive(root, "deliveredVerticalPosition")
        if vp_el is not None:
            offset = _find_float(vp_el, "offset")
            if offset != 0.0:
                return offset

        # Fallback: zoek offset op root niveau
        return _find_float(root, "offset")

    def _parse_grondlagen(self, root):
        """Parse grondlagen uit BHR-GT XML detail response.

        Structuur: boreholeSampleDescription > descriptiveBoreholeLog >
        layer > upperBoundary, lowerBoundary, soil > geotechnicalSoilName
        """
        grondlagen = []

        for layer_el in _findall_recursive(root, "layer"):
            bovenkant = _find_float(layer_el, "upperBoundary")
            onderkant = _find_float(layer_el, "lowerBoundary")

            # Skip ongeldige lagen
            if onderkant <= bovenkant:
                continue

            # Grondsoort: zoek in soil sub-element
            # soilNameNEN5104 is het meest betrouwbare veld (bijv.
            # 'zwakSiltigZand'), geotechnicalSoilName is vaak nil.
            beschrijving = ""
            for name in ["soilNameNEN5104", "geotechnicalSoilName",
                         "geotechnicalSoilCode", "lithology", "soilName"]:
                beschrijving = _find_text(layer_el, name)
                if beschrijving:
                    break

            grondsoort = classificeer_grondsoort(beschrijving)

            grondlagen.append(Grondlaag(
                bovenkant=bovenkant,
                onderkant=onderkant,
                grondsoort=grondsoort,
                beschrijving=beschrijving
            ))

        # Sorteer op diepte
        grondlagen.sort(key=lambda l: l.bovenkant)

        return grondlagen
