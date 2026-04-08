# GIS2BIM Migration Status Tracker

**Laatste update**: 8 februari 2026 (BAG3D tool DONE)
**Locatie**: `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\GIS2BIM.extension`

---

## Overall Progress

| Fase | Onderdeel | Voortgang | Status |
|------|-----------|-----------|--------|
| 1 | Core Library | 12/14 modules | In Progress |
| 2 | pyRevit Buttons | 5/12 buttons | In Progress |
| 3 | Testing | 30% | Started |

---

## BGT Button - DONE

Alle kernfunctionaliteit werkt:
- Filled regions met correcte holes/donuts (alle CurveLoops CCW, Revit bepaalt containment)
- Boundary lines via `bgt_bounderyline` lijnstijl (static `SetLineStyleId` call)
- Winding order via Shoelace formula (`_is_ring_ccw()`)
- 19 polygonen met holes correct getest (tot 4 holes per polygon)

### Opgelost op 8 feb 2026
1. `CurveLoop.IsCounterClockwise` bug -> Shoelace formula
2. Holes/donuts -> alle loops CCW, Revit containment
3. Boundary lines -> naam `bgt_bounderyline`, static call conform Dynamo V7

### Minor / toekomstig
- Detail lines (scheiding_lijn) worden nog getekend maar niet gewenst
- FilledRegionType per layer (bv. andere arcering voor water vs terrein)
- Extra BGT layers activeren (kunstwerken, spoor, gebieden)

---

## Module Status

### Coordinates & Projection
| Module | Status | Notes |
|--------|--------|-------|
| `coordinates.py` | DONE | RD - WGS84, IronPython compatible |
| `bbox.py` | DONE | BoundingBox class |
| `morton.py` | TODO | Tile indexing |

### API Clients (`api/`)
| Module | Status | Notes |
|--------|--------|-------|
| `pdok.py` | DONE | PDOKLocatie, PDOKBGT, PDOKWMTS |
| `wfs.py` | DONE | Generic WFS 2.0 client |
| `wfs_layers.py` | DONE | Kadaster, BAG layer configs |
| `ogc_api.py` | DONE | OGC API Features client (nieuwer dan WFS) |
| `bgt_layers.py` | DONE | BGT layer configuraties (19 layers) |
| `ahn.py` | DONE | AHN WCS + LAZ client (DTM/DSM, LAStools integratie) |
| `osm.py` | TODO | Overpass API |
| `bag3d.py` | DONE | 3DBAG WFS tile query + OBJ ZIP download |

### Parsers (`parsers/`)
| Module | Status | Notes |
|--------|--------|-------|
| `geotiff.py` | DONE | Ongecomprimeerde float32 GeoTIFF parser |
| `las.py` | DONE | LAS binary parser + XYZ tekst parser (chunked, thinning, classificatie) |
| `gml.py` | N/A | Niet nodig - OGC API geeft GeoJSON |
| `obj.py` | DONE | Wavefront OBJ parser (vertices, faces, groups) |
| `cityjson.py` | TODO | 3D gebouwen (CityJSON formaat) |

### Revit Interface (`revit/`)
| Module | Status | Notes |
|--------|--------|-------|
| `location.py` | DONE | Survey Point, Site Location, Project Info params |
| `geometry.py` | DONE | Model Lines, Detail Lines, Filled Regions, Text Notes |
| `images.py` | TODO | Raster import |

---

## pyRevit Button Status

### Setup Panel
| Button | Status | Features |
|--------|--------|----------|
| **Locatie** | DONE | Adres/postcode invoer, Revit locatie gebruiken, PDOK geocoding |
| Opschonen | TODO | - |

### Data Panel
| Button | Status | Features |
|--------|--------|----------|
| **WFS** | DONE | Kadaster percelen, BAG huisnummers, panden |
| **BGT** | DONE | Filled regions + holes + boundary lines |
| **AHN** | DONE | Hoogtedata DTM/DSM als TopographySurface/Toposolid |
| **BAG3D** | DONE | 3D gebouwmodellen als DirectShape (GenericModel) |

### Kaarten Panel
| Button | Status |
|--------|--------|
| Luchtfoto | TODO |
| Tijdreis | TODO |

### Overig
| Button | Status |
|--------|--------|
| OSM | TODO |

---

## BAG3D Button - DONE

3D gebouwmodellen laden van 3DBAG:
- WFS query naar `https://data.3dbag.nl/api/BAG3D/wfs` voor tile IDs + OBJ download URLs
- OBJ ZIP download met streaming (64KB chunks)
- Wavefront OBJ parser (vertices, faces, object/group support)
- DirectShape aanmaak via TessellatedShapeBuilder (Target=Mesh, Fallback=Salvage)
- Fan triangulatie voor polygonen > 3 vertices
- Degeneratie-check op driehoeken (kruisproduct)
- LoD selectie: 2.2 (dakvormen), 1.3 (blokken+dak), 1.2 (blokken)
- Per tile een DirectShape (GenericModel) element

---

## AHN Button - DONE

Hoogtedata laden via twee methoden:

### Methode 1: WCS (GeoTIFF) - Default
- Download GeoTIFF via WCS GetCoverage (geen externe tools nodig)
- Pure Python GeoTIFF parser (ongecomprimeerd float32)
- WCS URL: `https://service.pdok.nl/rws/ahn/wcs/v1_0`
- Coverages: `dtm_05m`, `dsm_05m`
- Format: GEOTIFF_FLOAT32 (~160KB voor 100x100m)

### Methode 2: LAZ (Puntenwolk) - Vereist LAStools
- Download AHN5 COPC LAZ tiles (1km x 1km)
- Verwerking via LAStools (las2txt / las2las / laszip)
- LAS binary parser met chunked reading voor grote bestanden
- Classification filter: class 2 (ground) voor DTM
- Grid-based thinning met keep-highest optie voor DSM
- LAZ URL: `https://fsn1.your-objectstorage.com/hwh-ahn/AHN5_KM/01_LAZ/`

### Beide methoden
- DTM (maaiveld) en DSM (oppervlakte) ondersteuning
- Configureerbare resolutie (0.5m, 1.0m, 2.0m)
- TopographySurface (Revit <= 2023) met Toposolid fallback (>= 2024)
- Nodata filtering + puntenschatting in UI

---

## Changelog

### 2026-02-08 (BAG3D)
- **BAG3D.pushbutton** - Nieuwe tool voor 3D gebouwmodellen laden
- `api/bag3d.py` - 3DBAG WFS client (tile query + OBJ ZIP download)
- `parsers/obj.py` - Wavefront OBJ parser (vertices, faces, groups)
- WFS query: `https://data.3dbag.nl/api/BAG3D/wfs` met bbox filter
- OBJ ZIP download met streaming, LOD selectie (lod12/lod13/lod22)
- DirectShape aanmaak via TessellatedShapeBuilder (Mesh target, Salvage fallback)
- Fan triangulatie voor n-gons, degeneratie-check op driehoeken
- WPF UI met LoD keuze, bbox grootte (100-1000m), progress indicator

### 2026-02-08 (AHN)
- **AHN.pushbutton** - Nieuwe tool voor AHN hoogtedata laden
- `api/ahn.py` - PDOK WCS + LAZ client (DTM/DSM, LAStools integratie)
- `parsers/geotiff.py` - Pure Python GeoTIFF parser (float32, uncompressed)
- `parsers/las.py` - LAS binary parser + XYZ tekst parser (chunked, classification filter, thinning)
- Twee methoden: WCS (GeoTIFF, geen tools) en LAZ (puntenwolk, LAStools)
- LAZ: AHN5 COPC tile download, las2txt/las2las/laszip integratie
- LAZ: Classification filter (class 2 = ground voor DTM), keep-highest thinning voor DSM
- WPF UI met methode/DTM-DSM keuze, LAStools status, resolutie, bbox grootte
- TopographySurface aanmaak met versiedetectie (Toposolid >= 2024)

### 2026-02-08 (BGT)
- IsCounterClockwise bug gefixt met Shoelace formula (`_is_ring_ccw()`)
- Filled regions worden weer correct aangemaakt
- Boundary line style naam gecorrigeerd: `bgt_bounderyline` (was `bgt_boundarylines`)
- `SetLineStyleId` aanroep gecorrigeerd naar static call (conform Dynamo V7 script)
- Holes/donuts werkend: alle CurveLoops CCW, Revit bepaalt containment
- 19 polygonen met holes correct aangemaakt in testgebied (tot 4 holes per polygon)
- **BGT button DONE** - alle kernfunctionaliteit werkt

### 2026-02-07
- **BGT.pushbutton** - Nieuwe tool voor BGT data laden
- `ogc_api.py` - OGC API Features client (vervangt oude WFS voor BGT)
- `bgt_layers.py` - 19 BGT layer configuraties
- API gebruikt nu https://api.pdok.nl/lv/bgt/ogc/v1/ (OGC API Features)
- Ondersteunt: wegdeel, terrein, water, pand, scheidingen, kruinlijnen
- WPF UI met layer selectie, bbox grootte, view keuze

### 2026-01-29
- **WFS.pushbutton** - Kadaster/BAG data laden
- `wfs.py` - Generic WFS 2.0 client
- `wfs_layers.py` - PDOK layer configuraties
- `geometry.py` - Model Lines, Filled Regions, Text Notes

### 2026-01-28
- Migratie naar `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\`
- Locatie button: optie toegevoegd om Revit Site Location te gebruiken
- `get_site_location()` functie toegevoegd

### 2026-01-27 (Sprint 1)
- `coordinates.py` - RD/WGS84 transformatie (IronPython 2.7 compatible)
- `bbox.py` - Bounding box utilities
- `api/pdok.py` - PDOK API clients
- `revit/location.py` - Survey Point manipulatie
- **Locatie.pushbutton** - Eerste werkende button!

---

## Extension Structuur

```
GIS2BIM.extension/
├── extension.json
├── GIS2BIM.tab/
│   ├── Setup.panel/
│   │   └── Locatie.pushbutton/     DONE
│   ├── Data.panel/
│   │   ├── WFS.pushbutton/         DONE
│   │   ├── BGT.pushbutton/         DONE
│   │   ├── AHN.pushbutton/         DONE
│   │   └── BAG3D.pushbutton/      DONE
│   ├── Kaarten.panel/              TODO
│   └── BGT.panel/                  -> verplaatst naar Data.panel
└── lib/gis2bim/
    ├── __init__.py
    ├── coordinates.py              DONE
    ├── bbox.py                     DONE
    ├── api/
    │   ├── pdok.py                 DONE
    │   ├── wfs.py                  DONE
    │   ├── wfs_layers.py           DONE
    │   ├── ogc_api.py              DONE
    │   ├── bgt_layers.py           DONE
    │   ├── ahn.py                  DONE
    │   └── bag3d.py                DONE
    ├── parsers/
    │   ├── geotiff.py              DONE
    │   ├── las.py                  DONE
    │   └── obj.py                  DONE
    └── revit/
        ├── location.py             DONE
        └── geometry.py             DONE
```

---

## BGT Layers Beschikbaar

| Categorie | Layers | Geometry | Actief in UI |
|-----------|--------|----------|--------------|
| Verharding | wegdeel, ondersteunendwegdeel, overbruggingsdeel | Polygon | wegdeel, ondersteunendwegdeel |
| Terrein | begroeidterreindeel, onbegroeidterreindeel | Polygon | beide |
| Water | waterdeel, ondersteunendwaterdeel | Polygon | waterdeel |
| Gebouwen | pand, overigbouwwerk | Polygon | pand |
| Scheidingen | scheiding_lijn, scheiding_vlak | Line/Polygon | scheiding_lijn (niet gewenst) |
| Kunstwerken | kunstwerkdeel_lijn, kunstwerkdeel_vlak | Line/Polygon | uitgeschakeld |
| Infrastructuur | spoor | Line | uitgeschakeld |
| Gebieden | functioneelgebied, openbareruimte | Polygon | uitgeschakeld |
| Kruinlijnen | diverse _kruinlijn layers | Line | uitgeschakeld |

---

## Technische Context

### Constraints
- **IronPython 2.7** - Geen type hints, geen f-strings, geen dataclasses
- **Geen externe packages** - Alleen stdlib + Revit API
- **Revit API** - Sommige .NET methoden niet beschikbaar via IronPython (bv. `CurveLoop.IsCounterClockwise`)

### APIs
| Service | URL | Gebruikt door |
|---------|-----|---------------|
| PDOK Locatieserver | `https://api.pdok.nl/bzk/locatieserver/search/v3_1/` | Locatie button |
| PDOK OGC API Features | `https://api.pdok.nl/lv/bgt/ogc/v1/` | BGT button |
| PDOK WFS 2.0 | `https://service.pdok.nl/lv/bag/wfs/v2_0` | WFS button |
| PDOK AHN WCS | `https://service.pdok.nl/rws/ahn/wcs/v1_0` | AHN button |
| 3DBAG WFS | `https://data.3dbag.nl/api/BAG3D/wfs` | BAG3D button |
| PDOK Luchtfoto WMTS | `https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0` | TODO |

### Coordinatensystemen
- **EPSG:28992** - Rijksdriehoekstelsel (RD New) - meters
- **EPSG:4326** - WGS84 (lat/lon) - graden
- **Revit internal** - feet, relatief aan Survey Point

### Referentie Dynamo Scripts
| Script | Locatie | Relevant voor |
|--------|---------|---------------|
| `GIS2BIM_BGT_bounderylines_V7.dyn` | `Z:\50_projecten\...\scripts_GIS\` | Boundary line style toepassen |
| `GIS2BIM_basis_v7.dyn` | `Z:\50_projecten\...\scripts_GIS\` | 3DBAG tiles + OBJ mesh als DirectShape |

---

## Deployment

```batch
REM Sync naar runtime:
X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\scripts\sync_pyrevit.bat
```

Of handmatig:
```powershell
$src = "X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\GIS2BIM.extension"
$dst = "$env:APPDATA\pyRevit\Extensions\GIS2BIM.extension"
Remove-Item $dst -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item $src -Destination $dst -Recurse -Force
```

---

## Volgende Stappen

1. **Luchtfoto button** - WMTS aerial imagery import
2. **Tijdreis button** - Historische kaarten
3. **OSM button** - OpenStreetMap data

---

*Status tracker - Update bij elke sprint*
