# 3BM pyRevit Project - Status

*Laatste update: 20 april 2026*

---

## Extensies Overzicht

### 3BM_Bouwkunde.extension (24 tools)

| Tool | Panel | UI Framework | Status | Beschrijving |
|------|-------|--------------|--------|--------------|
| RcBerekening | Bouwbesluit | WinForms | Stabiel | Rc/U-waarde + Glaser condensatie-analyse, PDF rapport |
| VentilatieBalans | Bouwbesluit | WinForms | Stabiel | Ventilatiebalans per ruimtezone |
| HellingbaanGenerator | Bouwbesluit | WinForms | Stabiel | NEN 2443 hellingbanen met DirectShape |
| AutoDim | Maatvoering | WinForms | Stabiel | Automatische maatvoering via Detail Lines |
| CrossDim | Maatvoering | WinForms | Stabiel | Kruislings dimensioneren |
| WandVloerAfwerking | Afwerking | WinForms | Stabiel | Wand/vloer afwerkingslagen per ruimte |
| SheetParameters | Document | WinForms | Stabiel | Bulk update titleblock parameters |
| FilterCreator | Filter | WinForms | Stabiel | Dynamisch filters aanmaken |
| ScheduleExport | Data Exchange | WinForms | Stabiel | Schedules exporteren naar Excel |
| ScheduleImport | Data Exchange | WinForms | Stabiel | Excel data importeren in schedules |
| PalenNummeren | Fundering | WinForms | Stabiel | Automatisch nummeren funderingspalen |
| IFCKozijnAnalyzer | IFC | WinForms | Stabiel | IFC kozijn/deur analyse |
| DetailOverzicht | Bibliotheek | WinForms | Stabiel | Overzicht detailbibliotheek als drafting view |
| NAAKTGenerator | Materialen | WinForms | Stabiel | NAA.K.T. materiaalbenaming generator |
| DbExp | Materialen | WinForms | Stabiel | Database export |
| MatExp | Materialen | WinForms | Stabiel | Materiaal export |
| MatImp | Materialen | WinForms | Stabiel | Materiaal import |
| MCPStatus | Test | **WPF** | Stabiel | MCP Server status (WPF referentie) |
| KozijnstaatWizard | Kozijnstaat | WinForms | Nieuw | Wizard voor complete kozijnstaat-workflow (4 stappen) |
| KozijnstaatCreate | Kozijnstaat | WinForms | Nieuw | Plaats unieke kozijntypes op grid (wand of auto-canvas) |
| KozijnstaatMaatvoering | Kozijnstaat | WinForms | Nieuw | Hoofd- en detailmaatvoering per kozijn in view |
| KozijnstaatGlasTag | Kozijnstaat | WinForms | Nieuw | Plaats glas-tags op glas-elementen in view |
| KozijnstaatAantallen | Kozijnstaat | WinForms | Nieuw | Tel kozijnen per type + spiegeling, schrijft params |
| KozijnstaatConfig | Kozijnstaat | WinForms | Nieuw | Instellingen voor alle Kozijnstaat tools |

### GIS2BIM.extension (10 tools)

| Tool | Panel | Status | Beschrijving |
|------|-------|--------|--------------|
| Locatie | Setup | Voltooid | Adres/postcode invoer, PDOK geocoding, Revit site locatie |
| WFS | Data | Voltooid | Kadaster percelen, BAG huisnummers, gebouwen via PDOK WFS 2.0 |
| BGT | Data | Voltooid | Basisregistratie Grootschalige Topografie (19 lagen, holes/donuts) |
| AHN | Data | Voltooid | Hoogte data als TopographySurface (WCS GeoTIFF + LAZ pointcloud) |
| BAG3D | Data | Voltooid | 3D gebouwen uit 3DBAG als DirectShape (OBJ mesh import) |
| NAPPeilmerken | Data | Voltooid | NAP peilmerken |
| WMS | Kaarten | Actief | Web Map Service kaarten op sheet |
| LuchtfotoTijdreis | Kaarten | Actief | PDOK luchtfoto's tijdreeks op sheet (3x2 grid) |
| KaartTijdreis | Kaarten | Todo | Historische kaarten tijdreeks |
| StreetView | Kaarten | Actief | Street View toegang |

---

## Lib Modules

### 3BM_Bouwkunde

| Module | Type | Beschrijving |
|--------|------|--------------|
| `wpf_template.py` | UI | WPF base classes + huisstijl (aanbevolen voor nieuwe tools) |
| `ui_template.py` | UI | Windows Forms: BaseForm, UIFactory, DPIScaler (legacy) |
| `bm_logger.py` | Logging | Centrale logging met auto-cleanup (max 10/tool, max 5MB) |
| `schedule_config.py` | Config | Persistente configuratie voor schedule export/import |
| `xlsx_helper.py` | IO | Excel bestandshelpers |
| `materialen_database.json` | Data | 1700+ materialen (lambda, Rd, mu-waarde) |
| `materialen_database_v2.0.csv` | Data | Materialen database (CSV formaat) |
| `naakt_data/` | Data | NAA.K.T. standaard datasets (namen, kenmerken, toepassingen) |
| `kozijnstaat/` | Logic | Shared helpers voor Kozijnstaat tools (family_collector, handedness, grid_layout, canvas_builder, reference_resolver, dimension_creator, tag_placer, config) |

### GIS2BIM

| Module | Pad | Beschrijving |
|--------|-----|--------------|
| `coordinates.py` | `lib/gis2bim/` | RD (EPSG:28992) ↔ WGS84 conversie |
| `bbox.py` | `lib/gis2bim/` | Bounding box utilities |
| `pdok.py` | `lib/gis2bim/api/` | PDOK Locatie, BGT, WMTS clients |
| `wfs.py` | `lib/gis2bim/api/` | WFS 2.0 client |
| `ogc_api.py` | `lib/gis2bim/api/` | OGC Features API client |
| `wfs_layers.py` | `lib/gis2bim/api/` | Kadaster/BAG configuraties |
| `bgt_layers.py` | `lib/gis2bim/api/` | BGT laag configuraties |
| `ahn.py` | `lib/gis2bim/api/` | AHN WCS + LAZ client |
| `bag3d.py` | `lib/gis2bim/api/` | 3DBAG WFS + OBJ download |
| `geotiff.py` | `lib/gis2bim/parsers/` | GeoTIFF float32 parser |
| `las.py` | `lib/gis2bim/parsers/` | LAS pointcloud parser |
| `obj.py` | `lib/gis2bim/parsers/` | Wavefront OBJ parser |

---

## Panel Structuur

```
3BM_Bouwkunde.extension/
└── 3BM Bouwkunde.tab/
    ├── Afwerking.panel/        WandVloerAfwerking
    ├── Bibliotheek.panel/      DetailOverzicht
    ├── Bouwbesluit.panel/      RcBerekening, VentilatieBalans, HellingbaanGenerator
    ├── Data Exchange.panel/    ScheduleExport, ScheduleImport
    ├── Document.panel/         SheetParameters
    ├── Filter.panel/           FilterCreator
    ├── Fundering.panel/        PalenNummeren
    ├── IFC.panel/              IFCKozijnAnalyzer
    ├── Maatvoering.panel/      AutoDim, CrossDim
    ├── Materialen.panel/       DbExp, MatExp, MatImp, NAAKTGenerator
    └── Test.panel/             MCPStatus

GIS2BIM.extension/
└── GIS2BIM.tab/
    ├── Setup.panel/            Locatie
    ├── Data.panel/             WFS, BGT, AHN, BAG3D, NAPPeilmerken
    └── Kaarten.panel/          WMS, LuchtfotoTijdreis, KaartTijdreis, StreetView
```

---

## Technische Stack

| Onderdeel | Technologie |
|-----------|-------------|
| Runtime | IronPython 2.7 (pyRevit) |
| UI (nieuw) | WPF + XAML |
| UI (legacy) | Windows Forms |
| Logging | bm_logger.py → logs/ |
| GIS APIs | PDOK, WFS 2.0, OGC Features, 3DBAG |
| Parsers | GeoTIFF, LAS, OBJ (pure Python) |
| Rapportage | PDF via 3BM Report API |
| Deploy | sync_pyrevit.bat → %APPDATA%\pyRevit\Extensions\ |

---

## Workflow

1. **Ontwikkel** in `extensions/` (deze map)
2. **Sync** via `scripts/sync_pyrevit.bat`
3. **Reload** in Revit: Alt+Click pyRevit → Reload
4. **Logs** controleren in `logs/`
