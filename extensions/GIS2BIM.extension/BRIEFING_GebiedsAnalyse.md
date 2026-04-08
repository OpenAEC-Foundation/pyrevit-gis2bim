# Briefing: GebiedsAnalyse Tool

*Instructie voor Claude Code — maart 2026*

---

## Doel

Nieuwe pyRevit tool: gebiedswaarde-analyse op basis van OSM voorzieningen.
Heatmap visualisatie in Revit via SpatialFieldManager (AnalysisVisualizationFramework).

---

## Wat al bestaat (HERGEBRUIKEN, niet opnieuw bouwen)

### Overpass client → `lib/gis2bim/api/overpass.py`
- `OverpassClient` met `get_features()`, HTTP POST, JSON parsing
- **LET OP**: `_parse_response_geom()` parst alleen `way` en `relation`, NIET `node`
- **ACTIE**: Voeg node-parsing toe, of maak een `get_points()` methode die `out center;` gebruikt en nodes+way-centers retourneert als `(lat, lon, tags)` tuples

### Coördinaten → `lib/gis2bim/coordinates.py`
- `rd_to_wgs84(x, y)` → `(lat, lon)`
- `wgs84_to_rd(lat, lon)` → `(x, y)`
- `distance_rd(x1, y1, x2, y2)` → meters
- `create_bbox_rd(cx, cy, width, height)` → tuple

### Bounding box → `lib/gis2bim/bbox.py`
- `BoundingBox.from_center(x, y, w, h)`
- `.expand(margin)` — voor ringmarge buiten analysegebied
- `.contains(x, y)`

### Locatie → `lib/gis2bim/revit/location.py`
- `get_rd_from_project_params(doc)` → `{"rd_x": float, "rd_y": float}`
- `get_project_location_rd(doc)` → idem via Survey Point
- `get_site_location(doc)` → WGS84 lat/lon

### UI → `lib/gis2bim/ui/`
- `progress_panel.py` — voortgangsbalk
- `location_setup.py` — locatie invoer formulier

### Bestaande tab structuur
```
GIS2BIM.tab/
├── Setup.panel/     (Locatie button)
├── Data.panel/      (AHN, BAG3D, BGT, OSM, WFS, etc.)
└── Kaarten.panel/   (WMS kaarten)
```

---

## Wat gebouwd moet worden

### 1. Nieuwe mapstructuur

```
GIS2BIM.tab/
└── Analyse.panel/
    └── GebiedsAnalyse.pushbutton/
        ├── script.py
        ├── bundle.yaml
        └── icon.svg

lib/gis2bim/
└── analysis/
    ├── __init__.py
    ├── categories.py    # Categorie definities + presets + OSM tags
    ├── grid.py          # Grid generatie + score berekening
    └── heatmap.py       # SpatialFieldManager visualisatie in Revit
```

### 2. Aanpassing bestaande code

**`lib/gis2bim/api/overpass.py`** — Voeg methode toe:

```python
def get_pois(self, bbox_wgs84, query_tags):
    """
    Haal POI punten op (nodes + way-centers).
    
    Returns:
        Lijst van dicts: [{"lat": float, "lon": float, "tags": dict, "osm_id": int}, ...]
    """
```

Deze methode moet:
- `out center;` gebruiken (niet `out geom;`)
- Zowel `node` als `way` elementen parsen
- Voor ways: het `center` veld gebruiken (lat/lon)
- Voor nodes: direct `lat`/`lon`
- Tags meegeven voor naam-weergave

### 3. `categories.py` — Categorie definities

```python
# Elke categorie bevat:
CATEGORIES = {
    "school": {
        "label": "School / onderwijs",
        "osm_tags": ['node["amenity"="school"]', 'way["amenity"="school"]',
                     'node["amenity"="kindergarten"]', 'way["amenity"="kindergarten"]'],
        "max_score": 1.0,
        "rings": [500, 1000, 2000, 5000],
        "color": (69, 182, 168),  # Teal
        "enabled": True,
    },
    "bus": {
        "label": "Bushalte",
        "osm_tags": ['node["highway"="bus_stop"]',
                     'node["public_transport"="platform"]["bus"="yes"]'],
        "max_score": 0.5,
        "rings": [50, 100, 200, 500],
        "color": (239, 189, 117),  # Yellow
        "enabled": True,
    },
    "trein": {
        "label": "Treinstation",
        "osm_tags": ['node["railway"="station"]', 'node["railway"="halt"]',
                     'way["railway"="station"]'],
        "max_score": 1.0,
        "rings": [100, 200, 500, 1000, 2000],
        "color": (160, 28, 72),  # Magenta
        "enabled": True,
    },
    "park": {
        "label": "Park / groen",
        "osm_tags": ['way["leisure"="park"]', 'relation["leisure"="park"]',
                     'way["landuse"="forest"]'],
        "max_score": 0.5,
        "rings": [100, 200, 500, 1000],
        "color": (29, 158, 117),  # Green
        "enabled": True,
    },
    "supermarkt": {
        "label": "Supermarkt / winkels",
        "osm_tags": ['node["shop"="supermarket"]', 'way["shop"="supermarket"]',
                     'node["shop"="convenience"]'],
        "max_score": 0.75,
        "rings": [200, 500, 1000, 2000],
        "color": (55, 138, 221),  # Blue
        "enabled": True,
    },
    "ziekenhuis": {
        "label": "Ziekenhuis",
        "osm_tags": ['node["amenity"="hospital"]', 'way["amenity"="hospital"]'],
        "max_score": 0.75,
        "rings": [500, 1000, 2000, 5000],
        "color": (219, 76, 64),  # Red
        "enabled": True,
    },
    "huisarts": {
        "label": "Huisarts",
        "osm_tags": ['node["amenity"="doctors"]', 'node["healthcare"="doctor"]',
                     'way["amenity"="doctors"]'],
        "max_score": 0.5,
        "rings": [200, 500, 1000, 2000],
        "color": (127, 77, 157),  # Purple
        "enabled": True,
    },
}

PRESETS = {
    "Woningbouw": { ... alle 7 categorieën met bovenstaande waarden },
    "Kantoor": { ... aangepaste gewichten voor kantoorlocaties },
}
```

### 4. `grid.py` — Grid + scoring

Kernfuncties:
- `generate_grid(center_lat, center_lon, size_m, resolution_m)` → lijst gridpunten
- `calculate_scores(grid_points, pois_per_category, categories)` → grid met scores
- Score per gridpunt: per categorie alleen de **dichtstbijzijnde** POI telt
- Ring score = `max_score × (1 - ring_index / aantal_ringen)`
- Buiten alle ringen = 0 punten
- Optioneel: Gaussian smoothing (simpele 3×3 kernel)
- **Gebruik haversine** voor WGS84 afstanden (niet `distance_rd`, want grid is in WGS84 voor Overpass)

```python
import math

def haversine(lat1, lon1, lat2, lon2):
    """Afstand in meters tussen twee WGS84 punten."""
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlam/2)**2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
```

### 5. `heatmap.py` — Revit SpatialFieldManager

Dit is het Revit-specifieke stuk. Workflow:
1. Maak een Floor element als analyse-oppervlak (of gebruik bestaand)
2. Haal SpatialFieldManager voor de actieve view
3. Registreer AnalysisResultSchema
4. Haal bovenvlak (PlanarFace met normal.Z ≈ 1) van de floor
5. Zet UV-punten + scores op het face
6. Maak AnalysisDisplayStyle aan met oranje→groen kleurschema

**Belangrijk**: Floor coördinaten in Revit internal units (feet). 
Conversie: `meters × (1000 / 304.8)` = feet.

Kleurstops voor het heatmap schema:
```python
# Van laag (0) naar hoog (max):
(232, 89, 60)     # Rood
(239, 139, 44)    # Oranje
(239, 189, 117)   # Geel
(181, 201, 74)    # Lichtgroen
(29, 117, 24)     # Groen
```

### 6. `script.py` — Hoofd UI

Gebruik de **bouwkunde** `ui_template.py` NIET — die zit in de andere extension.
GIS2BIM heeft eigen UI patterns (XAML-based, zie `ui/xaml_helper.py` en bestaande scripts).

Kijk naar hoe bestaande tools in `Data.panel/` hun UI doen en volg dat patroon.

Workflow:
1. Haal locatie op via `get_rd_from_project_params(doc)`, fallback naar `get_project_location_rd(doc)`
2. Toon instellingen formulier (categorieën, gewichten, grid config)
3. Haal OSM data op (met progress indicator)
4. Bereken grid scores
5. Visualiseer in Revit
6. Toon samenvatting (gem. score, max, min, score op projectlocatie)

---

## Bounding box strategie

Het analysegebied is bijv. 2000×2000m. Maar de buitenste ring (bijv. trein: 5000m) reikt verder.
Dus de Overpass query bbox moet zijn: analysegebied + maximale ring als marge.

```python
max_ring = max(max(cat["rings"]) for cat in enabled_categories)
bbox_rd = BoundingBox.from_center(rd_x, rd_y, grid_size + 2*max_ring, grid_size + 2*max_ring)
# Converteer hoeken naar WGS84 voor Overpass
sw_lat, sw_lon = rd_to_wgs84(bbox_rd.xmin, bbox_rd.ymin)
ne_lat, ne_lon = rd_to_wgs84(bbox_rd.xmax, bbox_rd.ymax)
bbox_wgs84 = (sw_lat, sw_lon, ne_lat, ne_lon)
```

---

## Caching

Cache OSM responses in `%APPDATA%\3BM_Bouwkunde\osm_cache\` als JSON.
Key: `{lat}_{lon}_{radius}_{category}.json`
Verlooptijd: 7 dagen.

---

## Niet vergeten

- [ ] `bundle.yaml` met `context: zero-doc` (tool werkt zonder selectie)
- [ ] `icon.svg` in 3BM huisstijl (violet achtergrond, teal tekst)
- [ ] IronPython 2.7 compatibel (geen f-strings, geen type hints, geen walrus operator)
- [ ] `__init__.py` in `analysis/` map
- [ ] `distance_rd()` uit coordinates.py NIET gebruiken voor WGS84 punten — gebruik haversine
- [ ] Overpass API rate limiting: max 1 request per seconde, retry bij 429/503
- [ ] Grid pre-filter: skip POIs die ver buiten de maximale ring vallen
