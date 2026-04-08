# 3BM Bouwkunde - Huidige Staat

*Laatste update: 30 januari 2026*

---

## 1. pyRevit Tools Overzicht

| Tool | Panel | Status | UI Template | Beschrijving |
|------|-------|--------|-------------|--------------|
| WandVloerAfwerking | Afwerking | ✅ Actief | ❓ Onbekend | Wand/vloer afwerkingslagen beheer |
| RcBerekening | Bouwbesluit | ✅ Actief | ✅ Ja | Rc-waarde berekening met Glaser methode |
| VentilatieBalans | Bouwbesluit | ✅ Actief | ❓ Onbekend | Ventilatie-eisen volgens BBL per ruimtezone |
| ScheduleExport | Data Exchange | ✅ Actief | ❓ Onbekend | Schedule data exporteren naar Excel |
| ScheduleImport | Data Exchange | ✅ Actief | ❓ Onbekend | Schedule data importeren vanuit Excel |
| SheetParameters | Document | ✅ Actief | ⚠️ Oud | Bulk update titleblock parameters |
| FilterCreator | Filter | ✅ Actief | ✅ Ja | Filter aanmaken obv selectie (SfB-code + typename) |
| PalenNummeren | Fundering | ✅ Actief | ❓ Onbekend | Automatisch nummeren funderingspalen |
| AutoDim | Maatvoering | ✅ Actief | ⚠️ Oud | Automatische wanddimensies via Detail Lines |
| CrossDim | Maatvoering | ✅ Actief | ❓ Onbekend | Kruislings dimensioneren |
| DbExp | Materialen | ✅ Actief | ❓ Onbekend | Database export |
| MatExp | Materialen | ✅ Actief | ❓ Onbekend | Materiaal export |
| MatImp | Materialen | ✅ Actief | ❓ Onbekend | Materiaal import |
| NAAKTGenerator | Materialen | ✅ Actief | ✅ Ja | NAA.K.T. materiaalbenaming generator |
| MCPStatus | Test | ✅ Actief | ❓ Onbekend | MCP server connectiviteit testen |

### Legenda Status
- ✅ **Actief**: Werkt, in gebruik
- ⚠️ **Oud**: Werkt, maar gebruikt oude UI code (niet ui_template)
- 🔧 **WIP**: In ontwikkeling
- ❌ **Broken**: Werkt niet
- ❓ **Onbekend**: Niet recent gecontroleerd

---

## 2. Lib Modules

| Module | Status | Beschrijving |
|--------|--------|--------------|
| `ui_template.py` | ✅ Actief | BaseForm, UIFactory, DPIScaler, Huisstijl |
| `bm_logger.py` | ✅ Actief | Centrale logging naar sync folder |
| `schedule_config.py` | ✅ Actief | Configuratie voor Schedule import/export |
| `xlsx_helper.py` | ✅ Actief | Excel bestandsoperaties helper |
| `materialen_database.json` | ✅ Actief | Materialen database (JSON formaat) |
| `materialen_database_v2.0.csv` | ✅ Actief | Materialen database (CSV formaat) |
| `naakt_data/` | ✅ Actief | NAA.K.T. standaard data (namen, kenmerken, toepassingen) |

---

## 3. Panel Structuur

```
Bouwkunde.tab/
├── Afwerking.panel/
│   └── WandVloerAfwerking.pushbutton
├── Bibliotheek.panel/
│   └── DetailOverzicht.pushbutton
├── Bouwbesluit.panel/
│   ├── HellingbaanGenerator.pushbutton
│   ├── RcBerekening.pushbutton
│   └── VentilatieBalans.pushbutton
├── Data Exchange.panel/
│   ├── ScheduleExport.pushbutton
│   └── ScheduleImport.pushbutton
├── Document.panel/
│   └── SheetParameters.pushbutton
├── Filter.panel/
│   └── FilterCreator.pushbutton
├── Fundering.panel/
│   └── PalenNummeren.pushbutton
├── IFC.panel/
│   └── IFCKozijnAnalyzer.pushbutton
├── Maatvoering.panel/
│   ├── AutoDim.pushbutton
│   └── CrossDim.pushbutton
├── Materialen.panel/
│   ├── DbExp.pushbutton
│   ├── MatExp.pushbutton
│   ├── MatImp.pushbutton
│   └── NAAKTGenerator.pushbutton
└── Test.panel/
    └── MCPStatus.pushbutton
```

---

## 4. TODO / Backlog

### Hoge prioriteit
- [ ] SheetParameters migreren naar ui_template
- [ ] AutoDim migreren naar ui_template
- [ ] Alle tools checken op ui_template compliance

### Gepland
- [ ] Rc-tool uitbreiden: dynamische vochtbalans (Glaser → tijdsafhankelijk)
- [ ] GIS2BIM port naar pyRevit

### Nice-to-have
- [ ] Screenshot MCP voor visuele debugging

---

## 5. Bekende Issues

| Tool | Issue | Workaround |
|------|-------|------------|
| AutoDim | Reference detection soms foutief bij complexe wanden | Handmatig edge selectie |
| SheetParameters | Tekst wordt afgesneden op HD schermen | Grotere venstermaat |

---

## 6. Recente Wijzigingen

### Januari 2026
- **FilterCreator** tool toegevoegd aan nieuw Filter.panel
  - Filter aanmaken op basis van geselecteerd element
  - Automatische SfB-code suggestie
  - Ondersteunt categorie-, type- en selectiefilters
  - Filternaam: `{SfB}_{omschrijving}` (underscores)
- Documentatie gesynchroniseerd met actuele toolset
- Data Exchange panel toegevoegd (ScheduleExport, ScheduleImport)
- MCPStatus tool toegevoegd aan Test panel
- NAA.K.T. Generator toegevoegd aan Materialen.panel
- UI template framework voltooid (BaseForm, UIFactory, DPIScaler)
- Centrale logging systeem (bm_logger.py)

### December 2025
- AutoDim tool toegevoegd
- RcBerekening met visuele condensatie-indicatoren
- 4K DPI scaling issues opgelost

---

## 7. Bestandslocaties

### Source (primair)
```
X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\
├── pyrevit\
│   └── bouwkunde.extension\
├── pyrevit_logs\
└── ...
```

### Runtime
```
%APPDATA%\pyRevit\Extensions\bouwkunde.extension\
```

### Legacy locaties
```
Z:\50_projecten\7_3BM_bouwkunde\_AI\pyrevit\
C:\DATA\3BM_projecten\50_projecten\7_3BM_bouwkunde\_AI\pyrevit\
```

---

## 8. Sync Commando

Na wijzigingen in source, sync naar runtime via MCP:

```python
# Via 3BM_Bouwkunde MCP server
sync_to_runtime()  # One-click sync
```

Of handmatig via PowerShell:
```powershell
$source = "X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension"
$runtime = "$env:APPDATA\pyRevit\Extensions\bouwkunde.extension"
Remove-Item $runtime -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item $source -Destination $runtime -Recurse -Force
Write-Host "Sync voltooid - reload pyRevit"
```
