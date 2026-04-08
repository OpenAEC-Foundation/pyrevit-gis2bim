# GIS2BIM pyRevit Extension

Migratie van Dynamo GIS2BIM scripts naar pure Python pyRevit Extension.

## 🎯 Doel

Nederlandse GIS-data integreren in Revit zonder Dynamo afhankelijkheid:
- PDOK Locatieserver (adres → RD coördinaten)
- BGT (Basisregistratie Grootschalige Topografie)
- Luchtfoto's (WMTS)
- AHN (Actueel Hoogtebestand Nederland)
- BAG 3D (3D gebouwgegevens)

## 📁 Mapstructuur

```
GIS2BIM/
├── 01_dynamo_reference/     ← Bron: originele Dynamo scripts ter referentie
│   ├── scripts/             ← .dyn workflow scripts
│   └── nodes/               ← .dyf node definities
├── pyrevit/
│   └── GIS2BIM.extension/   ← pyRevit extensie
│       ├── GIS2BIM.tab/
│       └── lib/gis2bim/
├── ROADMAP.md               ← Ontwikkelplan per sprint
├── STATUS.md                ← Voortgang tracker
└── README.md                ← Dit bestand
```

## 🚀 Installatie

### Automatisch (aanbevolen)
```
X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\pyrevit\sync_pyrevit.bat
```

### Handmatig
1. Kopieer `pyrevit/GIS2BIM.extension/` naar:
   ```
   %APPDATA%\pyRevit\Extensions\
   ```
2. Herlaad pyRevit in Revit (Alt+Click → Reload)

## ✅ Status

| Module | Status |
|--------|--------|
| coordinates.py | ✅ RD ↔ WGS84 |
| bbox.py | ✅ Bounding box |
| api/pdok.py | ✅ Locatie, BGT, WMTS |
| revit/location.py | ✅ Survey Point |

| Button | Status |
|--------|--------|
| Locatie | ✅ Werkend |
| Luchtfoto | ⬜ TODO |
| BGT | ⬜ TODO |

## 🔧 Technisch

- **Python**: IronPython 2.7 (pyRevit constraint)
- **Geen externe dependencies** - Pure stdlib + Revit API
- **Coördinaten**: EPSG:28992 (RD) ↔ EPSG:4326 (WGS84)

## 📚 Documentatie

- [ROADMAP.md](ROADMAP.md) - Ontwikkelplan
- [STATUS.md](STATUS.md) - Voortgang per module

---

*3BM Bouwkunde - Januari 2026*
