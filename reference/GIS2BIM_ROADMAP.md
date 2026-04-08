# GIS2BIM pyRevit Migration Roadmap

**Project**: Migratie Dynamo GIS2BIM scripts naar pure Python pyRevit Extension  
**Start**: Januari 2026  
**Locatie**: `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\GIS2BIM`

---

## 📊 Project Overzicht

### Bronnen Geanalyseerd
- **39 Dynamo scripts** (.dyn) - Workflows
- **89 Dynamo nodes** (.dyf) - Herbruikbare componenten
- **Python modules** - Bestaande BGT.py etc.

### Architectuur Beslissingen
1. **Pure Python** - Geen Dynamo afhankelijkheid
2. **IronPython 2.7** - pyRevit constraint (geen type hints, geen f-strings!)
3. **Modulair** - Herbruikbare lib modules
4. **pyRevit** - Native toolbar integratie

---

## 🏗️ Fase 1: Core Library (lib/)

### 1.1 Coördinaten & Projectie
| Module | Functionaliteit | Status |
|--------|-----------------|--------|
| `coordinates.py` | RD ↔ WGS84 transformatie | ✅ DONE |
| `bbox.py` | Bounding box berekeningen | ✅ DONE |
| `morton.py` | Morton code berekening | ⬜ TODO |

### 1.2 API Clients
| Module | API | Status |
|--------|-----|--------|
| `api/pdok.py` | PDOK Locatieserver, BGT, WMTS | ✅ DONE |
| `api/ahn.py` | AHN Hoogtedata | ⬜ TODO |
| `api/osm.py` | OpenStreetMap Overpass | ⬜ TODO |
| `api/bag3d.py` | 3D BAG CityJSON | ⬜ TODO |

### 1.3 Parsers & Geometrie
| Module | Functionaliteit | Status |
|--------|-----------------|--------|
| `parsers/gml.py` | GML naar geometrie | ⬜ TODO |
| `parsers/cityjson.py` | CityJSON parser | ⬜ TODO |

### 1.4 Revit Interface
| Module | Functionaliteit | Status |
|--------|-----------------|--------|
| `revit/location.py` | Survey Point, Site Location | ✅ DONE |
| `revit/geometry.py` | Curves, points naar Revit | ⬜ TODO |
| `revit/images.py` | Rasterafbeeldingen plaatsen | ⬜ TODO |

---

## 🖥️ Fase 2: pyRevit Buttons

### Toolbar Structuur
```
GIS2BIM.tab/
├── Setup.panel/
│   ├── Locatie.pushbutton     ✅ DONE
│   └── Opschonen.pushbutton   ⬜ TODO
│
├── Kaarten.panel/
│   ├── Luchtfoto.pushbutton   ⬜ TODO
│   └── Tijdreis.pushbutton    ⬜ TODO
│
├── BGT.panel/
│   ├── BGTCompleet.pushbutton ⬜ TODO
│   ├── Bomen.pushbutton       ⬜ TODO
│   └── Grenzen.pushbutton     ⬜ TODO
│
└── Data.panel/
    ├── AHN.pushbutton         ⬜ TODO
    ├── OSM.pushbutton         ⬜ TODO
    └── BAG3D.pushbutton       ⬜ TODO
```

---

## 📋 Sprint Planning

### Sprint 1: Fundament ✅ DONE
- [x] `coordinates.py` - CRS transformatie
- [x] `bbox.py` - Bounding box utilities
- [x] `api/pdok.py` - PDOK clients
- [x] `revit/location.py` - Survey Point
- [x] **Locatie.pushbutton**

### Sprint 2: Luchtfoto's (Volgende)
- [ ] `revit/images.py` - Raster import
- [ ] **Luchtfoto.pushbutton**
- [ ] **Tijdreis.pushbutton**

### Sprint 3: BGT
- [ ] `parsers/gml.py` - GML parsing
- [ ] `revit/geometry.py` - Curves naar Revit
- [ ] **BGTCompleet.pushbutton**

### Sprint 4: AHN & OSM
- [ ] `api/ahn.py` - AHN download
- [ ] `api/osm.py` - OSM Overpass
- [ ] **AHN.pushbutton**
- [ ] **OSM.pushbutton**

### Sprint 5: 3D
- [ ] `api/bag3d.py` - 3D BAG
- [ ] `parsers/cityjson.py` - CityJSON
- [ ] **BAG3D.pushbutton**

---

## 🔧 Technische Specificaties

### IronPython 2.7 Beperkingen
```python
# NIET TOEGESTAAN:
def func(x: int) -> str:  # Type hints
f"Waarde: {x}"            # f-strings
@dataclass                # Dataclasses

# WEL TOEGESTAAN:
def func(x):
"Waarde: {0}".format(x)
class MyClass(object):
    def __init__(self):
        pass
```

### API Endpoints (PDOK)
| Service | URL |
|---------|-----|
| Locatieserver | `https://api.pdok.nl/bzk/locatieserver/search/v3_1/` |
| BGT Download | `https://api.pdok.nl/lv/bgt/download/v1_0/` |
| Luchtfoto WMTS | `https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0` |
| AHN | `https://api.pdok.nl/rws/ahn/` |

### Coördinatensystemen
- **EPSG:28992** - Rijksdriehoekstelsel (RD New)
- **EPSG:4326** - WGS84 (lat/lon)

---

## 📁 Werkdirectory

```
X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\
├── pyrevit/
│   ├── 3BM_Bouwkunde.extension/   ← Bestaande tools
│   ├── GIS2BIM.extension/          ← GIS2BIM pyRevit
│   └── sync_pyrevit.py             ← Sync script (beide extensies)
├── GIS2BIM/
│   ├── 01_dynamo_reference/        ← Bronbestanden
│   ├── ROADMAP.md                  ← Dit document
│   ├── STATUS.md                   ← Voortgang
│   └── README.md
└── pyrevit_logs/
```

---

*Laatste update: 28 januari 2026*
