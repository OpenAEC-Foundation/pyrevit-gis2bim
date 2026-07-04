# -*- coding: utf-8 -*-
"""
DSO Ozon Presenteren API Client
================================

Client voor de DSO-LV "Omgevingsdocument Presenteren" API (Ozon, v8).
Dit is de databron achter "Regels op de kaart" in het Omgevingsloket:
omgevingsplannen, omgevingsverordeningen, waterschapsverordeningen etc.
met de juridische regels per locatie.

API-documentatie: https://developer.omgevingswet.overheid.nl/api-register/api/omgevingsdocument-presenteren/
Authenticatie: API-key in de 'x-api-key' header (aan te vragen via het
Ontwikkelaarsportaal). Geo-bevragingen in RD (EPSG:28992) vereisen de
'Content-Crs' header; coordinaten maximaal 3 decimalen.

Gebruik:
    from gis2bim.api.dso import DSOClient

    client = DSOClient(api_key="...")  # of key uit config.json
    regelingen = client.zoek_regelingen(155000.0, 463000.0)
    for reg in regelingen:
        annotaties = client.zoek_regelteksten(reg["identificatie"], 155000.0, 463000.0)
"""

import json

# Forceer TLS 1.2 (vereist op IronPython/.NET)
try:
    import clr
    clr.AddReference("System")
    from System.Net import ServicePointManager, SecurityProtocolType
    ServicePointManager.SecurityProtocol = (
        SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11
    )
except Exception:
    pass

try:
    import urllib2
    from urllib import urlencode
except ImportError:
    # Python 3
    import urllib.request as urllib2
    from urllib.parse import urlencode

try:
    from urllib2 import HTTPError
except ImportError:
    from urllib.error import HTTPError


# Config key-namen (opgeslagen in %APPDATA%\GIS2BIM\config.json)
CONFIG_KEY_API_KEY = "dso_api_key"
CONFIG_KEY_ENVIRONMENT = "dso_environment"

# Beschikbare DSO-omgevingen
DSO_ENVIRONMENTS = {
    "prod": "https://service.omgevingswet.overheid.nl/publiek/omgevingsdocumenten/api/presenteren/v8",
    "pre": "https://service.pre.omgevingswet.overheid.nl/publiek/omgevingsdocumenten/api/presenteren/v8",
}
DEFAULT_ENVIRONMENT = "prod"

# RD-coordinaten mogen maximaal 3 decimalen bevatten
RD_DECIMALS = 3

# Maximaal aantal pagina's dat gevolgd wordt bij paginering (veiligheidsklep)
MAX_PAGES = 20


class DSOAPIError(Exception):
    """Fout bij een DSO API request.

    Attributes:
        status_code: HTTP status code (0 als onbekend/netwerkfout).
        is_auth_error: True als de fout door een ongeldige/ontbrekende
            API-key komt (de API geeft hiervoor ook HTTP 400 met
            fault-type 'Ongeautoriseerd').
        message: Foutomschrijving.
    """

    def __init__(self, message, status_code=0, is_auth_error=False):
        super(DSOAPIError, self).__init__(message)
        self.status_code = status_code
        self.is_auth_error = is_auth_error


def normalize_identificatie(akn):
    """Normaliseer een AKN-identificatie voor gebruik in een URL-pad.

    Conform de API-documentatie worden alle niet-alfanumerieke tekens
    vervangen door een underscore.
    Bijv: "/akn/nl/act/gm0037/2019/3520-example" -> "_akn_nl_act_gm0037_2019_3520_example"

    Args:
        akn: AKN-identificatie zoals uitgeleverd door de API.

    Returns:
        str: Genormaliseerde identificatie.
    """
    result = []
    for ch in akn:
        if ch.isalnum():
            result.append(ch)
        else:
            result.append("_")
    return "".join(result)


class DSOClient(object):
    """
    Client voor de Ozon Presenteren API v8.

    Ondersteunt:
    - Regelingen zoeken op RD-punt of -polygoon (POST /regelingen/_zoek)
    - Regeltekstannotaties zoeken binnen een regeling (artikelstructuur)
    - Divisieannotaties zoeken binnen een regeling (vrijetekstmodel)
    - Documentstructuur ophalen
    - HAL-paginering (volgt automatisch 'next' links)
    """

    def __init__(self, api_key=None, environment=None, timeout=30):
        """
        Initialiseer de client.

        Args:
            api_key: DSO API-key. Als None wordt de key uit config.json gelezen.
            environment: "prod" of "pre". Als None uit config.json (default prod).
            timeout: Timeout in seconden per request.
        """
        if api_key is None or environment is None:
            from gis2bim import config
            if api_key is None:
                api_key = config.get_api_key(CONFIG_KEY_API_KEY)
            if environment is None:
                environment = config.load_config().get(
                    CONFIG_KEY_ENVIRONMENT, DEFAULT_ENVIRONMENT
                )
        if environment not in DSO_ENVIRONMENTS:
            environment = DEFAULT_ENVIRONMENT
        self.api_key = api_key
        self.environment = environment
        self.base_url = DSO_ENVIRONMENTS[environment]
        self.timeout = timeout

    # ------------------------------------------------------------------
    # Publieke API
    # ------------------------------------------------------------------

    def test_connection(self):
        """Test de verbinding en API-key via het /health endpoint.

        Returns:
            tuple: (ok, message) - ok is True bij succes.
        """
        try:
            data = self._request("/health")
            status = data.get("status", "onbekend") if isinstance(data, dict) else "ok"
            return True, "DSO API bereikbaar (status: {0})".format(status)
        except DSOAPIError as e:
            return False, str(e)

    def zoek_regelingen(self, rd_x, rd_y, geldig_op=None):
        """Zoek vastgestelde regelingen die gelden op een RD-punt.

        Args:
            rd_x: RD X-coordinaat (EPSG:28992).
            rd_y: RD Y-coordinaat.
            geldig_op: Optionele datum "YYYY-MM-DD" voor tijdreizen.

        Returns:
            list: Regeling-dicts zoals uitgeleverd door de API, o.a.
                  identificatie, officieleTitel, citeerTitel, type,
                  aangeleverdDoorEen (bevoegd gezag).
        """
        return self._post_zoek("/regelingen/_zoek", rd_x, rd_y, geldig_op=geldig_op)

    def zoek_regelteksten(self, regeling_identificatie, rd_x, rd_y, geldig_op=None):
        """Zoek regeltekstannotaties (artikelstructuur) op een RD-punt.

        Voor regelingen met een artikelsgewijze opbouw (omgevingsplan,
        omgevingsverordening). Levert de juridische regels (artikelen)
        waarvan het werkingsgebied het zoekpunt raakt.

        Args:
            regeling_identificatie: AKN-identificatie van de regeling
                (rauw of al genormaliseerd).
            rd_x: RD X-coordinaat.
            rd_y: RD Y-coordinaat.
            geldig_op: Optionele datum "YYYY-MM-DD".

        Returns:
            list: Regeltekstannotatie-dicts.
        """
        path = "/regelingen/{0}/regeltekstannotaties/_zoek".format(
            normalize_identificatie(regeling_identificatie)
        )
        return self._post_zoek(path, rd_x, rd_y, geldig_op=geldig_op)

    def zoek_divisies(self, regeling_identificatie, rd_x, rd_y, geldig_op=None):
        """Zoek divisieannotaties (vrijetekstmodel) op een RD-punt.

        Voor regelingen met vrijetekst-opbouw (bijv. omgevingsvisies).

        Args:
            regeling_identificatie: AKN-identificatie van de regeling.
            rd_x: RD X-coordinaat.
            rd_y: RD Y-coordinaat.
            geldig_op: Optionele datum "YYYY-MM-DD".

        Returns:
            list: Divisieannotatie-dicts.
        """
        path = "/regelingen/{0}/divisieannotaties/_zoek".format(
            normalize_identificatie(regeling_identificatie)
        )
        return self._post_zoek(path, rd_x, rd_y, geldig_op=geldig_op)

    def get_documentstructuur(self, regeling_identificatie, document_component=None):
        """Haal de documentstructuur (inhoudsopgave + tekst) van een regeling op.

        Args:
            regeling_identificatie: AKN-identificatie van de regeling.
            document_component: Optioneel wId van een documentdeel om alleen
                dat deel op te vragen (bijv. "gm0297_1__chp_2__art_2.4").

        Returns:
            dict: Documentstructuur zoals uitgeleverd door de API.
        """
        path = "/regelingen/{0}/documentstructuur".format(
            normalize_identificatie(regeling_identificatie)
        )
        params = {}
        if document_component:
            params["documentComponent"] = document_component
        return self._request(path, params=params)

    # ------------------------------------------------------------------
    # Intern
    # ------------------------------------------------------------------

    def _post_zoek(self, path, rd_x, rd_y, geldig_op=None):
        """Voer een geo-zoekbevraging uit en volg paginering.

        Args:
            path: API-pad van het _zoek endpoint.
            rd_x: RD X-coordinaat.
            rd_y: RD Y-coordinaat.
            geldig_op: Optionele datum "YYYY-MM-DD".

        Returns:
            list: Alle items uit '_embedded' over alle pagina's heen.
        """
        body = {
            "geo": {
                "geometrie": {
                    "type": "Point",
                    "coordinates": [
                        round(float(rd_x), RD_DECIMALS),
                        round(float(rd_y), RD_DECIMALS),
                    ],
                },
                "spatialOperator": "intersects",
            }
        }
        params = {}
        if geldig_op:
            params["geldigOp"] = geldig_op
            params["inWerkingOp"] = geldig_op

        items = []
        url = self._build_url(path, params)
        pages = 0
        while url and pages < MAX_PAGES:
            data = self._request_url(url, body=body)
            items.extend(self._extract_embedded(data))
            url = self._next_link(data)
            pages += 1
        return items

    def _build_url(self, path, params=None):
        """Bouw een volledige URL uit pad en query-parameters."""
        url = self.base_url + "/" + path.lstrip("/")
        if params:
            url = url + "?" + urlencode(params)
        return url

    def _request(self, path, params=None):
        """GET-request naar een API-pad. Returns geparste JSON."""
        return self._request_url(self._build_url(path, params))

    def _request_url(self, url, body=None):
        """Voer een request uit (GET, of POST als body meegegeven).

        Args:
            url: Volledige URL.
            body: Optionele dict; wordt als JSON verstuurd (POST).

        Returns:
            dict: Geparste JSON-response.

        Raises:
            DSOAPIError: Bij HTTP-fouten of ongeldig antwoord.
        """
        if not self.api_key:
            raise DSOAPIError(
                "Geen DSO API-key geconfigureerd. Vraag een key aan via "
                "https://developer.omgevingswet.overheid.nl en sla deze op."
            )
        data = None
        if body is not None:
            data = json.dumps(body).encode("utf-8")
        request = urllib2.Request(url, data)
        request.add_header("x-api-key", self.api_key)
        request.add_header("Accept", "application/hal+json, application/json")
        if body is not None:
            request.add_header("Content-Type", "application/json")
            # Verplicht bij geo-bevragingen in RD
            request.add_header("Content-Crs", "epsg:28992")
        try:
            response = urllib2.urlopen(request, timeout=self.timeout)
            raw = response.read()
        except HTTPError as e:
            detail = ""
            try:
                detail = e.read().decode("utf-8", "replace")[:500]
            except Exception:
                pass
            auth_fout = (e.code in (401, 403)
                         or "Ongeautoriseerd" in detail
                         or "API key" in detail)
            if auth_fout:
                msg = ("DSO API weigert de key (HTTP {0}). Controleer of de key "
                       "geldig is voor de '{1}'-omgeving. {2}").format(
                           e.code, self.environment, detail)
            else:
                msg = "DSO API fout (HTTP {0}) op {1}: {2}".format(
                    e.code, url, detail)
            raise DSOAPIError(msg, status_code=e.code, is_auth_error=auth_fout)
        except Exception as e:
            raise DSOAPIError("DSO API niet bereikbaar: {0}".format(e))
        try:
            if isinstance(raw, bytes):
                raw = raw.decode("utf-8")
            return json.loads(raw)
        except ValueError as e:
            raise DSOAPIError("Ongeldige JSON van DSO API: {0}".format(e))

    def _extract_embedded(self, data):
        """Haal de resultaatlijst uit een HAL-response.

        De API levert resultaten in '_embedded' onder een resource-naam
        (bijv. 'regelingen' of 'regelteksten'). We pakken defensief alle
        lijsten daaronder.
        """
        if not isinstance(data, dict):
            return []
        embedded = data.get("_embedded")
        if not isinstance(embedded, dict):
            # Sommige endpoints leveren direct een object
            return [data] if data else []
        items = []
        for value in embedded.values():
            if isinstance(value, list):
                items.extend(value)
        return items

    def _next_link(self, data):
        """Geef de 'next' paginerings-link uit een HAL-response, of None."""
        if not isinstance(data, dict):
            return None
        links = data.get("_links")
        if not isinstance(links, dict):
            return None
        next_link = links.get("next")
        if isinstance(next_link, dict):
            href = next_link.get("href")
            if href and href != links.get("self", {}).get("href"):
                return href
        return None
