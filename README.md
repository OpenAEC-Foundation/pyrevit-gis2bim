# 3BM pyRevit Extensions

pyRevit extensies voor bouwkunde en GIS-integratie in Revit.

## 📁 Structuur

```
pyrevit/
├── extensions/              ← SOURCE CODE (hier werken!)
│   ├── bouwkunde.extension/        ← Bouwkunde tools
│   └── GIS2BIM.extension/         ← GIS data integratie
├── scripts/                 ← Helper scripts
│   ├── sync_pyrevit.bat     ← SYNC NAAR REVIT
│   └── ...
├── logs/                    ← Tool logs
├── reference/               ← Documentatie & bronbestanden
└── *.md                     ← Project documentatie
```

## 🚀 Workflow

### 1. Code aanpassen
Werk in `extensions\*.extension\`

### 2. Sync naar Revit
```batch
scripts\sync_pyrevit.bat
```

### 3. Herlaad pyRevit
In Revit: Alt+Click op pyRevit tab → Reload

## 📦 Extensies

| Extensie | Beschrijving |
|----------|--------------|
| **bouwkunde** | Rc-berekening, Ventilatiebalans, AutoDim, NAA.K.T., etc. |
| **GIS2BIM** | PDOK Locatie, Luchtfoto's, BGT, AHN (WIP) |

## 📚 Documentatie

- [ARCHITECTURE.md](ARCHITECTURE.md) - Technische structuur
- [CONVENTIONS.md](CONVENTIONS.md) - Code conventies & huisstijl
- [CURRENT_STATE.md](CURRENT_STATE.md) - Status per tool

---

*3BM Bouwkunde - 2026*
