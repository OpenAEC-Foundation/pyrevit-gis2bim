# IFC Kozijn Analyzer v2.0

## Beschrijving
Analyseert K-type kozijnen (Kunststof & Schuifkozijnen) in gelinkte IFC modellen, vergelijkt met geladen Gealan families, en exporteert placement data naar CSV voor Revit plaatsing.

## Functionaliteit
1. **IFC Scanning**
   - Scant alle IFC links op Window elementen
   - Extraheert K-type codes uit IfcName (K01-01, SK-K01-hsb, etc.)
   - Leest locatie, orientatie en dimensies

2. **Family Vergelijking**
   - Inventariseert geladen Gealan/Kozijn families
   - Match/mismatch analyse per K-code
   - Identificeert ontbrekende families

3. **CSV Export**
   - Exporteert placement data in standaard formaat
   - Compatibel met MCP placement workflow

## CSV Formaat
```csv
ifc_id,type_name,mark,level,x_mm,y_mm,z_mm,normal_x,normal_y,normal_z,width_mm,height_mm,rotate_180
1236582,K13-02,100,,4344.2,13536.8,-16.0,-0.906308,0.422618,0.0,868.0,2937.0,0
```

| Kolom | Beschrijving |
|-------|--------------|
| ifc_id | Element ID uit IFC |
| type_name | K-type code (K01-01, SK-K02, etc.) |
| mark | Uniek nummer per kozijn |
| level | Verdieping (optioneel) |
| x_mm, y_mm, z_mm | Coördinaten in mm |
| normal_x/y/z | Normal vector voor oriëntatie |
| width_mm, height_mm | Afmetingen in mm |
| rotate_180 | 0 of 1 voor extra rotatie |

## Output Rapport
- Overzicht gevonden IFC links
- Tabel geladen Revit families per K-code
- Tabel IFC K-types met aantallen
- Vergelijkingstabel met status:
  - ✅ Match: K-type in beide aanwezig
  - ❌ Mist in Revit: Geen family geladen
  - ⚪ Alleen Revit: Family wel, niet in IFC

## Integratie in pyRevit Toolbar
1. Kopieer `IFC_Kozijn_Analyzer` map naar extension folder
2. Hernoem naar `IFC Kozijn Analyzer.pushbutton`
3. Plaats in panel structuur
4. Reload pyRevit

```
bouwkunde.extension/
└── Bouwkunde.tab/
    └── IFC Tools.panel/
        └── IFC Kozijn Analyzer.pushbutton/
            ├── script.py
            ├── bundle.yaml
            └── icon.png (optioneel)
```

## Workflow
1. Open Revit model met IFC links
2. Laad benodigde Gealan families
3. Run IFC Kozijn Analyzer
4. Bekijk vergelijkingsrapport
5. Exporteer CSV voor placement
6. Gebruik CSV met MCP tools voor plaatsing

## Versie
2.0 - Januari 2025
- CSV export toegevoegd
- Placement data extractie
- Rotate_180 berekening

## Auteur
3BM Bouwkunde
