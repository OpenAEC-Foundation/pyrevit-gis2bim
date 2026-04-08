# -*- coding: utf-8 -*-
"""
KLIC GML/XML Parser
====================

Parseer IMKL GML/XML bestanden uit KLIC leveringen (Kabels en Leidingen
Informatie Centrum). Ondersteunt zowel .zip als map input.

Coordinaten zijn in EPSG:28992 (Rijksdriehoekstelsel).

Gebruik:
    from gis2bim.parsers.klic import parse_klic_delivery

    delivery = parse_klic_delivery(r"Z:\pad\naar\KLIC\Levering_23O0024983_1")
    for feature in delivery.features:
        print(feature.feature_type, len(feature.geometry), "punten")
"""

import os
import zipfile
import shutil
import tempfile

try:
    import xml.etree.ElementTree as ET
except ImportError:
    ET = None


# =============================================================================
# Namespaces
# =============================================================================

NS = {
    'gml': 'http://www.opengis.net/gml/3.2',
    'imkl': 'http://www.geostandaarden.nl/imkl/wibon',
    'net': 'http://inspire.ec.europa.eu/schemas/net/4.0',
    'us-net-common': 'http://inspire.ec.europa.eu/schemas/us-net-common/4.0',
    'us-net-el': 'http://inspire.ec.europa.eu/schemas/us-net-el/4.0',
    'us-net-wa': 'http://inspire.ec.europa.eu/schemas/us-net-wa/4.0',
    'us-net-sw': 'http://inspire.ec.europa.eu/schemas/us-net-sw/4.0',
    'us-net-tc': 'http://inspire.ec.europa.eu/schemas/us-net-tc/4.0',
    'us-net-ogc': 'http://inspire.ec.europa.eu/schemas/us-net-ogc/4.0',
    'us-net-th': 'http://inspire.ec.europa.eu/schemas/us-net-th/4.0',
    'base': 'http://inspire.ec.europa.eu/schemas/base/3.3',
    'base2': 'http://inspire.ec.europa.eu/schemas/base2/2.0',
    'xlink': 'http://www.w3.org/1999/xlink',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
}


# =============================================================================
# Feature type classificatie
# =============================================================================

# Map van IMKL element tag (zonder namespace) naar netwerk type
FEATURE_TYPE_MAP = {
    'Elektriciteitskabel': 'electricity',
    'Telecommunicatiekabel': 'telecom',
    'OlieGasChemicalienPijpleiding': 'gas',
    'Waterleiding': 'water',
    'Rioolleiding': 'sewer',
    'Mantelbuis': 'duct',
    'Kabelbed': 'duct',
    'Duct': 'duct',
    'Overig': 'other',
    'TechnischGebouw': 'other',
    'Toren': 'other',
    'Mast': 'other',
}

# Nederlandse weergavenamen
FEATURE_TYPE_DISPLAY = {
    'Elektriciteitskabel': 'Elektriciteitskabel',
    'Telecommunicatiekabel': 'Telecommunicatiekabel',
    'OlieGasChemicalienPijpleiding': 'OlieGasChemicalienPijpleiding',
    'Waterleiding': 'Waterleiding',
    'Rioolleiding': 'Rioolleiding',
    'Mantelbuis': 'Mantelbuis',
    'Kabelbed': 'Kabelbed',
    'Duct': 'Duct',
    'Overig': 'Overig',
}


# =============================================================================
# Errors
# =============================================================================

class KLICError(Exception):
    """Fout bij parsen van KLIC data."""
    pass


# =============================================================================
# Dataklassen
# =============================================================================

class KLICFeature(object):
    """Eeen kabel, leiding, of annotatie uit KLIC data."""

    def __init__(self):
        self.feature_type = ""       # "Elektriciteitskabel", "Rioolleiding", etc.
        self.network_type = ""       # "electricity", "telecom", "gas", "water", "sewer", "duct", "other"
        self.geometry = []           # list[(x,y)]: polyline punten in RD
        self.geometry_type = "line"  # "line", "point", "polygon"
        self.label = ""              # optioneel label tekst
        self.label_position = None   # tuple(x,y): positie voor annotatie
        self.label_rotation = 0.0    # rotatie in graden
        self.properties = {}         # alle attributen (voltage, diameter, materiaal, etc.)

    def __repr__(self):
        pts = len(self.geometry) if self.geometry else 0
        return "KLICFeature({0}, {1}, {2} pts)".format(
            self.feature_type, self.network_type, pts)


class KLICDelivery(object):
    """Representeert een complete KLIC levering."""

    def __init__(self):
        self.klic_number = ""        # "23O0024983"
        self.delivery_date = ""      # "2023-02-21"
        self.request_info = {}       # aanvrager, werkzaamheden, locatie
        self.features = []           # list[KLICFeature] (kabels/leidingen)
        self.annotations = []        # list[KLICFeature] (Annotatie + Maatvoering)
        self.bbox = None             # tuple(xmin, ymin, xmax, ymax)
        self.stakeholders = []       # list[dict]: betrokken netbeheerders

    def feature_summary(self):
        """Geeft dict met aantallen per feature type."""
        counts = {}
        for f in self.features:
            key = f.feature_type
            counts[key] = counts.get(key, 0) + 1
        return counts

    def __repr__(self):
        return "KLICDelivery({0}, {1} features, {2} annotations)".format(
            self.klic_number, len(self.features), len(self.annotations))


# =============================================================================
# Hoofd parse functies
# =============================================================================

def parse_klic_delivery(path):
    """
    Parse KLIC levering van zip of map.

    Args:
        path: Pad naar KLIC levering map of .zip bestand

    Returns:
        KLICDelivery object

    Raises:
        KLICError: Als het pad ongeldig is of parsing mislukt
    """
    if not os.path.exists(path):
        raise KLICError("Pad bestaat niet: {0}".format(path))

    temp_dir = None

    try:
        # ZIP support
        if zipfile.is_zipfile(path):
            temp_dir = tempfile.mkdtemp(prefix="klic_")
            with zipfile.ZipFile(path, 'r') as zf:
                zf.extractall(temp_dir)
            folder = temp_dir
        elif os.path.isdir(path):
            folder = path
        else:
            raise KLICError("Pad is geen map of ZIP: {0}".format(path))

        # Zoek het GML bestand
        xml_path = find_klic_xml(folder)
        if xml_path is None:
            raise KLICError(
                "Geen GI_gebiedsinformatielevering_*.xml gevonden in: {0}".format(folder))

        # Parse
        return parse_klic_gml(xml_path)

    finally:
        # Cleanup temp dir
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass


def find_klic_xml(folder):
    """
    Vind het GI_gebiedsinformatielevering_*.xml bestand.

    Zoekt recursief in de map en submappen.

    Args:
        folder: Map om in te zoeken

    Returns:
        Volledig pad naar het XML bestand, of None
    """
    # Zoek recursief via os.walk (IronPython/Python 2 compatibel)
    for root, dirs, files in os.walk(folder):
        for f in files:
            if f.startswith("GI_gebiedsinformatielevering_") and f.endswith(".xml"):
                return os.path.join(root, f)

    return None


def parse_klic_gml(xml_path):
    """
    Parse het KLIC GML bestand.

    Args:
        xml_path: Pad naar het GI_gebiedsinformatielevering_*.xml bestand

    Returns:
        KLICDelivery object

    Raises:
        KLICError: Bij parsing fouten
    """
    if ET is None:
        raise KLICError("xml.etree.ElementTree is niet beschikbaar")

    try:
        tree = ET.parse(xml_path)
    except ET.ParseError as e:
        raise KLICError("XML parse fout: {0}".format(e))

    root = tree.getroot()
    delivery = KLICDelivery()

    # Stap 1: Parse bounding box
    delivery.bbox = _parse_bbox(root)

    # Stap 2: Parse metadata (aanvraag info)
    _parse_metadata(root, delivery)

    # Stap 3: Index alle UtilityLink elementen (gml:id -> geometrie)
    utility_links = _index_utility_links(root)

    # Stap 4: Index netwerk types (gml:id -> network_type)
    network_types = _index_network_types(root)

    # Stap 5: Parse infrastructure features (kabels/leidingen)
    _parse_infrastructure_features(root, delivery, utility_links, network_types)

    # Stap 6: Parse ExtraGeometrie (mantelbuis polygonen etc.)
    _parse_extra_geometrie(root, delivery, network_types)

    # Stap 7: Parse annotaties en maatvoering
    _parse_annotations(root, delivery)

    # Stap 8: Parse stakeholders
    _parse_stakeholders(root, delivery)

    return delivery


# =============================================================================
# Interne parse functies
# =============================================================================

def _parse_bbox(root):
    """Parse bounding box uit het root element."""
    envelope = root.find('.//gml:boundedBy/gml:Envelope', NS)
    if envelope is None:
        return None

    lower = envelope.find('gml:lowerCorner', NS)
    upper = envelope.find('gml:upperCorner', NS)

    if lower is None or upper is None:
        return None

    try:
        lc = lower.text.strip().split()
        uc = upper.text.strip().split()
        return (float(lc[0]), float(lc[1]), float(uc[0]), float(uc[1]))
    except (ValueError, IndexError):
        return None


def _parse_metadata(root, delivery):
    """Parse metadata uit GebiedsinformatieAanvraag."""
    aanvraag = root.find('.//imkl:GebiedsinformatieAanvraag', NS)
    if aanvraag is not None:
        # KLIC meldnummer
        meld = aanvraag.find('imkl:klicMeldnummer', NS)
        if meld is not None and meld.text:
            delivery.klic_number = meld.text.strip()

        # Aanvraagdatum
        datum = aanvraag.find('imkl:aanvraagDatum', NS)
        if datum is not None and datum.text:
            # Parse ISO datetime, neem alleen datum deel
            dt_text = datum.text.strip()
            if 'T' in dt_text:
                delivery.delivery_date = dt_text.split('T')[0]
            else:
                delivery.delivery_date = dt_text

        # Werkzaamheden
        soort = aanvraag.find('imkl:soortWerkzaamheden', NS)
        if soort is not None:
            href = soort.get('{http://www.w3.org/1999/xlink}href', '')
            if '/' in href:
                delivery.request_info['soortWerkzaamheden'] = href.rsplit('/', 1)[-1]

        omschrijving = aanvraag.find('imkl:omschrijvingWerkzaamheden', NS)
        if omschrijving is not None and omschrijving.text:
            delivery.request_info['omschrijving'] = omschrijving.text.strip()

        # Ordernummer
        order = aanvraag.find('imkl:ordernummer', NS)
        if order is not None and order.text:
            delivery.request_info['ordernummer'] = order.text.strip()

    # Leveringsdatum uit GebiedsinformatieLevering
    levering = root.find('.//imkl:GebiedsinformatieLevering', NS)
    if levering is not None:
        datum_samengesteld = levering.find('imkl:datumLeveringSamengesteld', NS)
        if datum_samengesteld is not None and datum_samengesteld.text:
            dt_text = datum_samengesteld.text.strip()
            if not delivery.delivery_date:
                if 'T' in dt_text:
                    delivery.delivery_date = dt_text.split('T')[0]
                else:
                    delivery.delivery_date = dt_text


def _index_utility_links(root):
    """
    Indexeer alle UtilityLink elementen op gml:id.

    UtilityLinks bevatten de werkelijke geometrie (centrelineGeometry)
    van kabels en leidingen. Features verwijzen hier naar via xlink:href.

    Returns:
        dict: gml:id -> list[(x,y)] polyline punten
    """
    links = {}

    for member in root.iter('{http://inspire.ec.europa.eu/schemas/us-net-common/4.0}UtilityLink'):
        gml_id = member.get('{http://www.opengis.net/gml/3.2}id', '')
        if not gml_id:
            continue

        # Haal geometrie uit centrelineGeometry
        geometry = _parse_geometry_elements(member, 'net:centrelineGeometry')
        if geometry:
            links[gml_id] = geometry

    return links


def _index_network_types(root):
    """
    Indexeer netwerk types uit Utiliteitsnet elementen.

    Returns:
        dict: gml:id -> network_type string
    """
    networks = {}

    for net_elem in root.iter('{http://www.geostandaarden.nl/imkl/wibon}Utiliteitsnet'):
        gml_id = net_elem.get('{http://www.opengis.net/gml/3.2}id', '')
        if not gml_id:
            continue

        # Bepaal type uit utilityNetworkType href
        type_elem = net_elem.find('us-net-common:utilityNetworkType', NS)
        if type_elem is not None:
            href = type_elem.get('{http://www.w3.org/1999/xlink}href', '')
            if 'electricity' in href:
                networks[gml_id] = 'electricity'
            elif 'telecommunications' in href:
                networks[gml_id] = 'telecom'
            elif 'oilGasChemical' in href:
                networks[gml_id] = 'gas'
            elif 'water' in href:
                networks[gml_id] = 'water'
            elif 'sewer' in href:
                networks[gml_id] = 'sewer'
            elif 'thermal' in href:
                networks[gml_id] = 'other'
            else:
                networks[gml_id] = 'other'

        # Haal ook thema op voor extra context
        thema = net_elem.find('imkl:thema', NS)
        if thema is not None:
            href = thema.get('{http://www.w3.org/1999/xlink}href', '')
            if gml_id not in networks:
                if 'riool' in href.lower():
                    networks[gml_id] = 'sewer'
                elif 'water' in href.lower():
                    networks[gml_id] = 'water'

    return networks


def _parse_infrastructure_features(root, delivery, utility_links, network_types):
    """Parse kabels en leidingen features."""
    for member in root.findall('gml:featureMember', NS):
        for child in member:
            tag = _local_name(child.tag)

            if tag not in FEATURE_TYPE_MAP:
                continue

            feature = KLICFeature()
            feature.feature_type = tag
            feature.network_type = FEATURE_TYPE_MAP[tag]

            # Probeer network type te verfijnen via inNetwork referentie
            network_ref = child.find('net:inNetwork', NS)
            if network_ref is not None:
                href = network_ref.get('{http://www.w3.org/1999/xlink}href', '')
                if href in network_types:
                    feature.network_type = network_types[href]

            # Haal geometrie via net:link -> UtilityLink
            all_points = []
            for link_elem in child.findall('net:link', NS):
                href = link_elem.get('{http://www.w3.org/1999/xlink}href', '')
                if href in utility_links:
                    all_points.extend(utility_links[href])
                # Probeer ook met nl.imkl- prefix
                elif href and not href.startswith('nl.imkl-'):
                    alt_href = 'nl.imkl-' + href
                    if alt_href in utility_links:
                        all_points.extend(utility_links[alt_href])

            if all_points:
                feature.geometry = all_points
                feature.geometry_type = "line"
            else:
                # Probeer directe geometrie (sommige features hebben inline geometrie)
                direct_geom = _parse_geometry_elements(child, 'net:geometry')
                if direct_geom:
                    feature.geometry = direct_geom
                    feature.geometry_type = "line"

            # Extraheer properties
            _extract_feature_properties(child, feature)

            # Label
            label_elem = child.find('imkl:label', NS)
            if label_elem is not None and label_elem.text:
                feature.label = label_elem.text.strip()

            # Voeg alleen toe als er geometrie is
            if feature.geometry:
                delivery.features.append(feature)


def _parse_extra_geometrie(root, delivery, network_types):
    """Parse ExtraGeometrie elementen (polygonen voor mantelbuizen etc.)."""
    for member in root.findall('gml:featureMember', NS):
        for child in member:
            tag = _local_name(child.tag)
            if tag != 'ExtraGeometrie':
                continue

            # Parse vlakgeometrie
            vlak = child.find('imkl:vlakgeometrie2D', NS)
            if vlak is None:
                continue

            polygon_points = _parse_polygon(vlak)
            if not polygon_points:
                continue

            feature = KLICFeature()
            feature.feature_type = 'ExtraGeometrie'
            feature.geometry = polygon_points
            feature.geometry_type = "polygon"

            # Bepaal netwerk type
            network_ref = child.find('imkl:inNetwork', NS)
            if network_ref is not None:
                href = network_ref.get('{http://www.w3.org/1999/xlink}href', '')
                if href in network_types:
                    feature.network_type = network_types[href]
                else:
                    feature.network_type = _guess_network_from_id(href)
            else:
                feature.network_type = 'other'

            # Probeer feature type af te leiden uit gml:id
            gml_id = child.get('{http://www.opengis.net/gml/3.2}id', '')
            for ft_name in FEATURE_TYPE_MAP:
                if ft_name in gml_id:
                    feature.feature_type = ft_name
                    feature.network_type = FEATURE_TYPE_MAP[ft_name]
                    break

            delivery.features.append(feature)


def _parse_annotations(root, delivery):
    """Parse Annotatie en Maatvoering elementen."""
    for member in root.findall('gml:featureMember', NS):
        for child in member:
            tag = _local_name(child.tag)

            if tag not in ('Annotatie', 'Maatvoering'):
                continue

            feature = KLICFeature()
            feature.feature_type = tag

            # Label tekst
            label_elem = child.find('imkl:label', NS)
            if label_elem is not None and label_elem.text:
                feature.label = label_elem.text.strip()

            # Rotatiehoek
            rotatie = child.find('imkl:rotatiehoek', NS)
            if rotatie is not None and rotatie.text:
                try:
                    feature.label_rotation = float(rotatie.text.strip())
                except ValueError:
                    pass

            # Geometrie uit imkl:ligging
            ligging = child.find('imkl:ligging', NS)
            if ligging is not None:
                # Check voor Point
                point = ligging.find('gml:Point', NS)
                if point is not None:
                    pos = point.find('gml:pos', NS)
                    if pos is not None and pos.text:
                        coords = _parse_pos(pos.text)
                        if coords:
                            feature.geometry = [coords]
                            feature.geometry_type = "point"
                            feature.label_position = coords

                # Check voor LineString
                linestring = ligging.find('gml:LineString', NS)
                if linestring is not None:
                    poslist = linestring.find('gml:posList', NS)
                    if poslist is not None and poslist.text:
                        points = _parse_pos_list(poslist.text)
                        if points:
                            feature.geometry = points
                            feature.geometry_type = "line"
                            # Label positie = midden van lijn
                            if feature.label:
                                mid_idx = len(points) // 2
                                feature.label_position = points[mid_idx]

            # Netwerk referentie
            network_ref = child.find('imkl:inNetwork', NS)
            if network_ref is not None:
                href = network_ref.get('{http://www.w3.org/1999/xlink}href', '')
                feature.properties['inNetwork'] = href

            # Annotatie/maatvoering type
            if tag == 'Annotatie':
                atype = child.find('imkl:annotatieType', NS)
                if atype is not None:
                    href = atype.get('{http://www.w3.org/1999/xlink}href', '')
                    feature.properties['annotatieType'] = href.rsplit('/', 1)[-1] if '/' in href else href
            elif tag == 'Maatvoering':
                mtype = child.find('imkl:maatvoeringsType', NS)
                if mtype is not None:
                    href = mtype.get('{http://www.w3.org/1999/xlink}href', '')
                    feature.properties['maatvoeringsType'] = href.rsplit('/', 1)[-1] if '/' in href else href

            if feature.geometry:
                delivery.annotations.append(feature)


def _parse_stakeholders(root, delivery):
    """Parse Belanghebbende elementen."""
    for member in root.findall('gml:featureMember', NS):
        for child in member:
            tag = _local_name(child.tag)
            if tag != 'Belanghebbende':
                continue

            stakeholder = {}

            # Naam
            naam = child.find('imkl:naam', NS)
            if naam is not None and naam.text:
                stakeholder['naam'] = naam.text.strip()

            # Netbeheerder ID uit gml:id
            gml_id = child.get('{http://www.opengis.net/gml/3.2}id', '')
            stakeholder['id'] = gml_id

            # Contactgegevens
            telefoon = child.find('.//imkl:telefoon', NS)
            if telefoon is not None and telefoon.text:
                stakeholder['telefoon'] = telefoon.text.strip()

            email = child.find('.//imkl:email', NS)
            if email is not None and email.text:
                stakeholder['email'] = email.text.strip()

            if stakeholder.get('naam'):
                delivery.stakeholders.append(stakeholder)


# =============================================================================
# Geometrie helpers
# =============================================================================

def _parse_geometry_elements(parent, geometry_path):
    """
    Extraheer LineString geometrie uit een element.

    Args:
        parent: Parent XML element
        geometry_path: Namespace-prefixed pad naar geometry container
                       (bijv. 'net:centrelineGeometry')

    Returns:
        list[(x,y)] of lege lijst
    """
    all_points = []

    # Zoek geometry container
    ns_parts = geometry_path.split(':')
    if len(ns_parts) == 2 and ns_parts[0] in NS:
        full_path = '{' + NS[ns_parts[0]] + '}' + ns_parts[1]
    else:
        full_path = geometry_path

    for geom_container in parent.iter(full_path):
        # LineString met posList
        for ls in geom_container.iter('{http://www.opengis.net/gml/3.2}LineString'):
            poslist = ls.find('{http://www.opengis.net/gml/3.2}posList')
            if poslist is not None and poslist.text:
                points = _parse_pos_list(poslist.text)
                all_points.extend(points)

        # Point met pos
        for pt in geom_container.iter('{http://www.opengis.net/gml/3.2}Point'):
            pos = pt.find('{http://www.opengis.net/gml/3.2}pos')
            if pos is not None and pos.text:
                coord = _parse_pos(pos.text)
                if coord:
                    all_points.append(coord)

    # Als geen container gevonden, zoek direct in parent
    if not all_points:
        for ls in parent.iter('{http://www.opengis.net/gml/3.2}LineString'):
            poslist = ls.find('{http://www.opengis.net/gml/3.2}posList')
            if poslist is not None and poslist.text:
                points = _parse_pos_list(poslist.text)
                all_points.extend(points)

    return all_points


def _parse_polygon(parent):
    """Parse gml:Polygon uit een element. Returns list[(x,y)]."""
    polygon = parent.find('.//gml:Polygon', NS)
    if polygon is None:
        return []

    # Exterior ring
    exterior = polygon.find('.//gml:exterior//gml:posList', NS)
    if exterior is None:
        exterior = polygon.find('.//gml:LinearRing/gml:posList', NS)

    if exterior is not None and exterior.text:
        return _parse_pos_list(exterior.text)

    return []


def _parse_pos_list(text):
    """
    Parse 'x1 y1 x2 y2 ...' tekst naar [(x1,y1), (x2,y2), ...].

    Args:
        text: String met spatie-gescheiden coordinaten

    Returns:
        list[(float, float)] coordinaat paren
    """
    if not text:
        return []

    values = text.strip().split()
    points = []

    for i in range(0, len(values) - 1, 2):
        try:
            x = float(values[i])
            y = float(values[i + 1])
            points.append((x, y))
        except (ValueError, IndexError):
            continue

    return points


def _parse_pos(text):
    """
    Parse 'x y' tekst naar (x, y) tuple.

    Args:
        text: String met twee spatie-gescheiden coordinaten

    Returns:
        tuple(float, float) of None
    """
    if not text:
        return None

    parts = text.strip().split()
    if len(parts) >= 2:
        try:
            return (float(parts[0]), float(parts[1]))
        except ValueError:
            return None

    return None


# =============================================================================
# Property extraction
# =============================================================================

def _extract_feature_properties(element, feature):
    """Extraheer relevante properties uit een feature element."""
    # Voltage (elektriciteit)
    for tag in ('us-net-el:operatingVoltage', 'us-net-el:nominalVoltage'):
        ns_parts = tag.split(':')
        el = element.find('{' + NS.get(ns_parts[0], '') + '}' + ns_parts[1])
        if el is not None and el.text:
            try:
                feature.properties['voltage'] = float(el.text)
                uom = el.get('uom', '')
                if uom:
                    feature.properties['voltage_uom'] = uom.rsplit('::', 1)[-1] if '::' in uom else uom
            except ValueError:
                pass
            break

    # Pipe diameter
    diameter_el = element.find('us-net-common:pipeDiameter', NS)
    if diameter_el is not None and diameter_el.text and diameter_el.text.strip():
        try:
            feature.properties['diameter'] = float(diameter_el.text)
            uom = diameter_el.get('uom', '')
            if uom:
                feature.properties['diameter_uom'] = uom.rsplit('::', 1)[-1] if '::' in uom else uom
        except ValueError:
            pass

    # Druk (gas)
    pressure_el = element.find('us-net-common:pressure', NS)
    if pressure_el is not None and pressure_el.text:
        try:
            feature.properties['pressure'] = float(pressure_el.text)
            uom = pressure_el.get('uom', '')
            if uom:
                feature.properties['pressure_uom'] = uom.rsplit('::', 1)[-1] if '::' in uom else uom
        except ValueError:
            pass

    # Buismateriaal
    materiaal = element.find('imkl:buismateriaalType', NS)
    if materiaal is not None:
        href = materiaal.get('{http://www.w3.org/1999/xlink}href', '')
        if '/' in href:
            feature.properties['materiaal'] = href.rsplit('/', 1)[-1]

    # Status
    status = element.find('us-net-common:currentStatus', NS)
    if status is not None:
        href = status.get('{http://www.w3.org/1999/xlink}href', '')
        if '/' in href:
            feature.properties['status'] = href.rsplit('/', 1)[-1]

    # Verticale positie
    vpos = element.find('us-net-common:verticalPosition', NS)
    if vpos is not None and vpos.text:
        feature.properties['verticalPosition'] = vpos.text.strip()

    # Water type
    water_type = element.find('us-net-wa:waterType', NS)
    if water_type is not None:
        href = water_type.get('{http://www.w3.org/1999/xlink}href', '')
        if '/' in href:
            feature.properties['waterType'] = href.rsplit('/', 1)[-1]

    # Rioolwater type
    sewer_type = element.find('us-net-sw:sewerWaterType', NS)
    if sewer_type is not None:
        href = sewer_type.get('{http://www.w3.org/1999/xlink}href', '')
        if '/' in href:
            feature.properties['sewerWaterType'] = href.rsplit('/', 1)[-1]

    # Gas product type
    gas_type = element.find('us-net-ogc:oilGasChemicalsProductType', NS)
    if gas_type is not None:
        href = gas_type.get('{http://www.w3.org/1999/xlink}href', '')
        if '/' in href:
            feature.properties['productType'] = href.rsplit('/', 1)[-1]


# =============================================================================
# Hulpfuncties
# =============================================================================

def _local_name(tag):
    """Haal local name uit een {namespace}name tag."""
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def _guess_network_from_id(href):
    """Raad netwerk type op basis van gml:id/href patronen."""
    href_lower = href.lower()
    if 'electric' in href_lower:
        return 'electricity'
    elif 'telecom' in href_lower or 'datatransport' in href_lower:
        return 'telecom'
    elif 'gas' in href_lower or 'oilgas' in href_lower or 'oilGasChemicals' in href:
        return 'gas'
    elif 'water' in href_lower and 'riool' not in href_lower:
        return 'water'
    elif 'riool' in href_lower or 'sewer' in href_lower:
        return 'sewer'
    else:
        return 'other'
