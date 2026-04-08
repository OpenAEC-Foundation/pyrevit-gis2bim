# 3BM Bouwkunde - Architectuur & Structuur

## 1. pyRevit Extension Structuur

```
bouwkunde.extension/
├── extension.json              # Extensie metadata
├── lib/                        # Gedeelde modules
│   ├── __init__.py
│   ├── ui_template.py          # UI framework (BaseForm, UIFactory, etc.)
│   ├── bm_logger.py            # Centrale logging
│   ├── schedule_config.py      # Schedule import/export configuratie
│   ├── xlsx_helper.py          # Excel bestandsoperaties
│   ├── materialen_database.json    # Materialen database
│   ├── materialen_database_v2.0.csv
│   └── naakt_data/             # NAA.K.T. standaard data
│       ├── naakt_namen.json
│       ├── naakt_kenmerken.json
│       └── naakt_toepassingen.json
│
└── Bouwkunde.tab/              # Hoofdtab in Revit ribbon
    │
    ├── Afwerking.panel/        # Afwerkingslagen
    │   └── WandVloerAfwerking.pushbutton/
    │
    ├── Bouwbesluit.panel/      # Bouwfysica & regelgeving
    │   ├── HellingbaanGenerator.pushbutton/
    │   ├── RcBerekening.pushbutton/
    │   └── VentilatieBalans.pushbutton/
    │
    ├── Data Exchange.panel/    # Import/export workflows
    │   ├── ScheduleExport.pushbutton/
    │   └── ScheduleImport.pushbutton/
    │
    ├── Document.panel/         # Document/sheet management
    │   └── SheetParameters.pushbutton/
    │
    ├── Filter.panel/           # View filters
    │   └── FilterCreator.pushbutton/
    │
    ├── Fundering.panel/        # Constructief - fundering
    │   └── PalenNummeren.pushbutton/
    │
    ├── IFC.panel/              # IFC analyse
    │   └── IFCKozijnAnalyzer.pushbutton/
    │
    ├── Maatvoering.panel/      # Automatische dimensionering
    │   ├── AutoDim.pushbutton/
    │   └── CrossDim.pushbutton/
    │
    ├── Materialen.panel/       # Materiaal beheer
    │   ├── DbExp.pushbutton/
    │   ├── MatExp.pushbutton/
    │   ├── MatImp.pushbutton/
    │   └── NAAKTGenerator.pushbutton/
    │
    └── Test.panel/             # Development/test tools
        └── MCPStatus.pushbutton/
```

---

## 2. Nieuwe Tool Maken - Stappenplan

### Stap 1: Mapstructuur aanmaken
```powershell
$base = "X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension\Bouwkunde.tab"
New-Item -Path "$base\NieuwPanel.panel\NieuweTool.pushbutton" -ItemType Directory -Force
```

### Stap 2: script.py maken
```python
# -*- coding: utf-8 -*-
"""Tool beschrijving - wat doet deze tool?"""
__title__ = "Tool\nNaam"
__author__ = "3BM Bouwkunde"

# Imports
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

# UI imports - ALTIJD via ui_template!
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl
from bm_logger import get_logger

# Setup
doc = revit.doc
log = get_logger("NieuweTool")


class NieuweToolForm(BaseForm):
    """Hoofd UI formulier"""
    
    def __init__(self, data):
        super(NieuweToolForm, self).__init__("Tool Titel", 900, 700)
        self.data = data
        self.set_subtitle("{} elementen".format(len(data)))
        self._setup_ui()
    
    def _setup_ui(self):
        """Bouw de UI op"""
        y = 10
        
        # Voorbeeld: Label
        lbl = UIFactory.create_label("Selectie:", bold=True)
        lbl.Location = DPIScaler.scale_point(10, y)
        self.pnl_content.Controls.Add(lbl)
        y += 30
        
        # Voorbeeld: DataGridView
        columns = [
            ("id", "Element ID", 100),
            ("naam", "Naam", 200),
            ("waarde", "Waarde", 100),
        ]
        self.grid = UIFactory.create_datagridview(columns, 500, 300)
        self.grid.Location = DPIScaler.scale_point(10, y)
        self.pnl_content.Controls.Add(self.grid)
        
        # Footer button
        self.add_footer_button("Uitvoeren", 'primary', self._execute)
    
    def _execute(self, sender, args):
        """Hoofdactie uitvoeren"""
        log.info("Uitvoeren gestart")
        try:
            with revit.Transaction("Tool Actie"):
                # Doe iets...
                pass
            self.show_info("Klaar!")
            self.Close()
        except Exception as e:
            log.error("Fout: {}".format(e), exc_info=True)
            self.show_error("Fout: {}".format(e))


def main():
    """Entry point"""
    log.info("Tool gestart")
    
    # Data verzamelen
    data = []  # Vul met elementen
    
    if not data:
        forms.alert("Geen data gevonden.", exitscript=True)
    
    # UI tonen
    form = NieuweToolForm(data)
    form.ShowDialog()


if __name__ == "__main__":
    main()
```

### Stap 3: bundle.yaml maken
```yaml
title: Tool Naam
tooltip: Beschrijving van wat de tool doet
author: 3BM Bouwkunde
highlight: new
context: selection  # of: zero-doc, active-doc
```

### Stap 4: Sync naar runtime
Via MCP server (aanbevolen):
```python
sync_to_runtime()  # One-click sync
```

Of handmatig via PowerShell:
```powershell
$source = "X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension"
$runtime = "$env:APPDATA\pyRevit\Extensions\bouwkunde.extension"
Remove-Item $runtime -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item $source -Destination $runtime -Recurse -Force
```

### Stap 5: pyRevit herladen
- Alt+Click op pyRevit tab → Reload
- Of: pyRevit CLI `pyrevit reload`

---

## 3. Lib Modules

### ui_template.py

**Classes:**
| Klasse | Beschrijving |
|--------|--------------|
| `Huisstijl` | Kleurconstanten (VIOLET, TEAL, YELLOW, etc.), `get_material_color()` |
| `DPIScaler` | `scale()`, `scale_point()`, `scale_size()` - automatische 4K scaling |
| `UIFactory` | `create_label()`, `create_button()`, `create_datagridview()`, `create_combobox()`, `create_textbox()`, `create_checkbox()`, `create_groupbox()`, `create_panel()` |
| `BaseForm` | Basis Form met header, content panel, footer. Methods: `set_subtitle()`, `add_footer_button()`, `show_info()`, `show_error()`, `show_warning()`, `ask_confirm()`, `save_file_dialog()` |
| `LayoutHelper` | `stack_vertical()`, `stack_horizontal()`, `create_form_row()` |

### bm_logger.py

**Gebruik:**
```python
from bm_logger import get_logger
log = get_logger("ToolNaam")

log.debug("Gedetailleerde info")
log.info("Normale info")
log.warning("Waarschuwing")
log.error("Fout", exc_info=True)  # Met stack trace
```

### schedule_config.py

Configuratie voor Schedule import/export tools:
- Schedule field mappings
- Export/import settings
- Column definitions

### xlsx_helper.py

Excel bestandsoperaties helper:
- Lezen/schrijven van .xlsx bestanden
- Cell formatting
- Sheet management

### naakt_data/

NAA.K.T. (Naam-Attribuut-Kenmerk-Toepassing) standaard voor materiaalbenaming:
- `naakt_namen.json` - Hoofdgroepen (beton, hout, isolatie, etc.)
- `naakt_kenmerken.json` - Kenmerken per naam
- `naakt_toepassingen.json` - Toepassingen per naam

---

## 4. Multi-Device Workflow

### Bestanden synchroniseren

| Device | Source pad | Sync methode |
|--------|------------|--------------|
| PC (X: drive) | `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\` | Primair |
| PC (Z: drive) | `Z:\50_projecten\7_3BM_bouwkunde\_AI\pyrevit\` | Legacy |
| Laptop | `C:\DATA\3BM_projecten\...\` | OneDrive files-on-demand (legacy) |
| Telefoon | N/A | Geen file access |

### Claude workflow per device

**PC sessie:**
1. MCP server leest/schrijft direct naar source
2. `sync_to_runtime()` na wijzigingen
3. Logs beschikbaar in `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\pyrevit_logs\`

**Laptop sessie:**
1. OneDrive synct automatisch
2. Zelfde workflow als PC

**Telefoon sessie:**
1. Alleen planning/discussie
2. Geen code wijzigingen mogelijk

---

## 5. Extension.json

```json
{
    "type": "extension",
    "name": "Bouwkunde",
    "description": "Bouwkundige tools voor Revit - ventilatie, bouwfysica, controles",
    "author": "3BM Bouwkunde",
    "version": "1.0.0"
}
```

---

## 6. Troubleshooting

### Tool verschijnt niet in Revit
1. Check naamconventies: `.extension`, `.tab`, `.panel`, `.pushbutton`
2. Controleer `extension.json` aanwezig
3. Reload pyRevit (Alt+Click → Reload)

### Import errors
1. Check `lib/__init__.py` aanwezig
2. Controleer sys.path append in script
3. Bekijk pyRevit output window voor details

### UI scaling issues
1. Gebruik ALTIJD `DPIScaler.scale_point()` voor posities
2. BaseForm heeft automatische DPI handling
3. Test op zowel HD als 4K scherm
