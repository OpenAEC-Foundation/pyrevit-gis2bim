# PyRevit 3BM Toolbar

Bouwkundige toolbalk extensie voor 3BM Bouwkunde workflows in Autodesk Revit.

---

## Doel

Custom PyRevit extensie met tools voor dagelijkse 3BM workflows: data-extractie, berekeningen triggeren, rapporten genereren — allemaal vanuit de Revit UI.

---

## Technologie

- **Taal:** IronPython 2.7 (Revit context)
- **Framework:** PyRevit
- **Revit API:** `Autodesk.Revit.DB`
- **Runtime pad:** `D:\Github\pyrevit\extensions` (direct geladen via pyRevit config)

---

## Extensie structuur

```
extensions/
  bouwkunde.extension/
    3BMtab.tab/
      [PanelNaam].panel/
        [ToolNaam].pushbutton/
          script.py       ← hoofdscript (IronPython 2.7)
          bundle.yaml     ← metadata (optioneel)
          icon.png        ← 32x32 of 16x16 (optioneel)
```

---

## Script conventies

```python
# Standaard imports
from pyrevit import forms, script, revit, DB
from Autodesk.Revit.DB import *

# Actief document
doc = revit.doc
uidoc = revit.uidoc

# Output
output = script.get_output()
```

- **Geen f-strings** — IronPython 2.7, gebruik `.format()` of `%`
- **Geen pip** — alleen standaard library + PyRevit/Revit API
- **Altijd try/except** — crashes in Revit process zijn impactvol
- **forms.alert()** voor foutmeldingen naar gebruiker

---

## Sync naar runtime

pyRevit laadt extensies direct vanuit `D:\Github\pyrevit\extensions` (geen sync nodig).
Na wijzigingen in Revit: Alt+Click op PyRevit logo → Reload

---

## Data uitwisseling met andere projecten

- Schrijf data weg als JSON voor consumptie door andere tools
- Lees JSON van andere tools in (warmteverlies input, report trigger)
- Tijdelijk uitwisselingpad: `%TEMP%\3bm_exchange\`

```python
import json, os

# Wegschrijven voor warmteverlies
pad = os.path.join(os.environ['TEMP'], '3bm_exchange', 'ruimtes_input.json')
with open(pad, 'w') as f:
    json.dump(data, f, indent=2)
```

---

## Agent Broker
- **project_id:** `pyrevit`
- **display_name:** `PyRevit 3BM Toolbar`
- **capabilities:** `["revit-api", "ifc-export", "data-extraction"]`
- **subscriptions:** `["warmteverlies/*", "bim/*", "shared/*"]`

---

## Orchestrator — Sessie afsluiting

**ALTIJD uitvoeren aan het einde van elke sessie** (of na een significante mijlpaal):

Schrijf een update naar:
`C:\Users\JochemK\.claude\orchestrator\sessions\pyrevit_latest.md`

Gebruik dit formaat:
```markdown
# PyRevit — Sessie update
**Datum:** YYYY-MM-DD HH:MM

## Wat is gedaan
- (bullet per afgeronde taak)

## Huidige staat
(1-3 zinnen over de staat van het project)

## Gewijzigde bestanden
- (relevante paden binnen extensions/)

## Openstaande issues / next steps
- (wat nog moet gebeuren)

## Cross-project notities
(iets wat relevant is voor warmteverlies of report integratie)
```

**Orchestrator context:** `C:\Users\JochemK\.claude\orchestrator\context\pyrevit.md`
**Project registry:** `C:\Users\JochemK\.claude\orchestrator\project-registry.json`
