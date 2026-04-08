# 3BM Bouwkunde - Conventies & Standaarden

## 1. Huisstijl Kleuren

| Naam | Hex | RGB | Gebruik |
|------|-----|-----|---------|
| Magic Violet | `#350E35` | `(53, 14, 53)` | Headers, accenten, primaire tekst |
| Verdigris (Teal) | `#45B6A8` | `(69, 182, 168)` | Primaire actieknoppen, succes |
| Friendly Yellow | `#EFBD75` | `(239, 189, 117)` | Waarschuwingen |
| Warm Magenta | `#A01C48` | `(160, 28, 72)` | Alerts, belangrijke meldingen |
| Flaming Peach | `#DB4C40` | `(219, 76, 64)` | Errors, destructieve acties |

### Neutrale kleuren
- `WHITE` - Achtergrond, tekst op donker
- `LIGHT_GRAY` - `#F5F5F5` - Subtiele achtergronden
- `TEXT_PRIMARY` - `#323232` - Hoofdtekst
- `TEXT_SECONDARY` - `#808080` - Secundaire tekst

### Materiaal kleuren (bouwfysica)
```python
'isolatie': (255, 230, 150)  # Geel
'beton': (180, 180, 180)     # Grijs
'hout': (205, 170, 125)      # Bruin
'steen': (200, 100, 100)     # Rood
'folie': (100, 150, 255)     # Blauw
'gips': (245, 245, 245)      # Wit
'metaal': (160, 160, 180)    # Zilver
'lucht': (220, 240, 255)     # Lichtblauw
```

---

## 2. UI Framework

### WPF (AANBEVOLEN voor nieuwe tools)

WPF biedt declaratieve XAML layouts met automatische DPI scaling. Referentie: `UITestWPF.pushbutton`

**Bestandsstructuur:**
```
MijnTool.pushbutton/
├── script.py       # Python logic + event handlers
├── UI.xaml         # XAML layout
├── bundle.yaml
└── icon.svg
```

**XAML Template:**
```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Tool Naam" Width="500" SizeToContent="Height">
    
    <Window.Resources>
        <SolidColorBrush x:Key="VioletBrush" Color="#350E35"/>
        <SolidColorBrush x:Key="TealBrush" Color="#45B6A8"/>
        <SolidColorBrush x:Key="LightGrayBrush" Color="#F5F5F5"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#808080"/>
    </Window.Resources>
    
    <DockPanel>
        <!-- Header: Violet background, white title, teal subtitle -->
        <StackPanel DockPanel.Dock="Top" Background="{StaticResource VioletBrush}">
            <TextBlock Text="Tool Titel" FontSize="18" Foreground="White" Margin="20,16,20,4"/>
            <TextBlock Text="Subtitel" FontSize="12" Foreground="{StaticResource TealBrush}" Margin="20,0,20,16"/>
        </StackPanel>
        <Rectangle DockPanel.Dock="Top" Height="3" Fill="{StaticResource TealBrush}"/>
        
        <!-- Footer: Light gray, buttons right-aligned -->
        <Border DockPanel.Dock="Bottom" Background="{StaticResource LightGrayBrush}" Padding="20,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btn_cancel" Content="Annuleren" Padding="20,8" Margin="0,0,8,0"/>
                <Button x:Name="btn_execute" Content="Uitvoeren" Padding="20,8"
                        Background="{StaticResource TealBrush}" Foreground="White"/>
            </StackPanel>
        </Border>
        
        <!-- Content: ScrollViewer met StackPanel -->
        <ScrollViewer Padding="20">
            <StackPanel>
                <!-- Controls hier -->
            </StackPanel>
        </ScrollViewer>
    </DockPanel>
</Window>
```

**Python Script Pattern:**
```python
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('System.Xml')

from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows import Window
from System.Windows.Markup import XamlReader

class MijnToolWindow(Window):
    def __init__(self):
        Window.__init__(self)
        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        with open(xaml_path, 'r') as f:
            xaml = f.read()
        loaded = XamlReader.Load(SysXmlReader.Create(StringReader(xaml)))
        self.Content = loaded.Content
        self.Title = loaded.Title
        
        # Bind elements
        self.btn_execute = loaded.FindName('btn_execute')
        self.btn_execute.Click += self._on_execute
```

---

### Windows Forms (Legacy)

> Alleen voor onderhoud bestaande tools. Nieuwe tools: gebruik WPF.

```python
from ui_template import BaseForm, UIFactory, DPIScaler, Huisstijl

class MijnToolForm(BaseForm):
    def __init__(self):
        super(MijnToolForm, self).__init__("Tool Naam", 800, 600)
        self._setup_ui()
    
    def _setup_ui(self):
        y = 10
        lbl = UIFactory.create_label("Tekst", bold=True)
        lbl.Location = DPIScaler.scale_point(10, y)
        self.pnl_content.Controls.Add(lbl)
```

### Windows Forms Componenten (Legacy)

| Klasse | Doel |
|--------|------|
| `BaseForm` | Basis formulier met header, content area, footer |
| `UIFactory` | Factory voor gestylde controls (labels, buttons, grids) |
| `DPIScaler` | Automatische schaling voor 4K displays |
| `Huisstijl` | Kleurconstanten en materiaal lookup |
| `LayoutHelper` | Hulpmethodes voor control positionering |

---

### Button stijlen (WPF & WinForms)

| Stijl | Achtergrond | Tekst | Gebruik |
|-------|-------------|-------|---------|
| `primary` | Teal | Wit | Hoofdactie (Opslaan, Berekenen) |
| `secondary` | Wit | Violet | Secundaire actie (Annuleren) |
| `warning` | Geel | Violet | Waarschuwing |
| `danger` | Rood | Wit | Destructieve actie |
| `icon` | Wit | Violet | Kleine icon knop (+, -, etc.) |

### Font sizes

| Constante | Grootte | Gebruik |
|-----------|---------|---------|
| `FONT_TITLE` | 18 | Form titels |
| `FONT_SUBTITLE` | 14 | Subtitels |
| `FONT_HEADING` | 12 | Sectie headers |
| `FONT_NORMAL` | 10 | Normale tekst |
| `FONT_SMALL` | 9 | Labels, hints |
| `FONT_TINY` | 8 | Kleine annotaties |

---

## 3. pyRevit Naamconventies

### Mapstructuur
```
bouwkunde.extension/
├── extension.json
├── lib/
│   ├── __init__.py
│   ├── wpf_template.py     # WPF framework (AANBEVOLEN)
│   ├── ui_template.py      # Windows Forms framework (legacy)
│   ├── bm_logger.py
│   └── naakt_data/
└── Bouwkunde.tab/
    ├── Bouwbesluit.panel/
    │   └── RcBerekening.pushbutton/
    │       ├── script.py
    │       ├── UI.xaml         # WPF layout (indien WPF)
    │       ├── bundle.yaml
    │       └── icon.svg
    └── Test.panel/
        ├── UITest.pushbutton/      # Windows Forms demo
        └── UITestWPF.pushbutton/   # WPF demo (REFERENTIE)
```

### Bestanden per pushbutton

| Bestand | Verplicht | Inhoud |
|---------|-----------|--------|
| `script.py` | ✅ | Hoofdscript met `__title__` en `__author__` |
| `UI.xaml` | ⚠️ WPF | XAML layout (alleen voor WPF tools) |
| `bundle.yaml` | ⚠️ | Tooltip, auteur, context |
| `icon.png` of `icon.svg` | ⚠️ | 32x32 of 96x96 pixels (SVG preferred) |

### Script header
```python
# -*- coding: utf-8 -*-
"""Tool beschrijving"""
__title__ = "Tool\nNaam"  # \n voor twee regels op knop
__author__ = "3BM Bouwkunde"
__doc__ = "Tooltip tekst voor de knop"
```

### Icon SVG template (3BM huisstijl)
```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
  <rect width="32" height="32" rx="3" fill="#350E35"/>
  <text x="16" y="21" font-family="Segoe UI" font-size="12" 
        font-weight="bold" fill="#45B6A8" text-anchor="middle">TXT</text>
</svg>
```

---

## 4. Logging

### Gebruik `bm_logger.py`
```python
from bm_logger import get_logger
log = get_logger("ToolNaam")

log.info("Tool gestart")
log.debug("Detail: {}".format(waarde))
log.warning("Let op...")
log.error("Fout!", exc_info=True)
```

### Log locaties (volgorde)
1. `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\pyrevit_logs\` (primair)
2. `Z:\50_projecten\7_3BM_bouwkunde\_AI\pyrevit_logs\` (legacy)
3. `C:\DATA\3BM_projecten\50_projecten\7_3BM_bouwkunde\_AI\pyrevit_logs\` (legacy)
4. `%APPDATA%\3BM_Bouwkunde\logs\` (Fallback)

---

## 5. Revit API Conventies

### Unit conversie
```python
# Thermische geleidbaarheid (λ) van Revit naar SI
THERMAL_CONDUCTIVITY_FACTOR = 6.93347
lambda_si = revit_waarde / THERMAL_CONDUCTIVITY_FACTOR  # W/(m·K)

# Lengtes
from Autodesk.Revit.DB import UnitUtils, UnitTypeId
mm = UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Millimeters)
```

### Transacties
```python
with revit.Transaction("Actie beschrijving"):
    # Wijzigingen aan model
    element.LookupParameter("Param").Set(waarde)
```

---

## 6. Bestandslocaties

| Wat | Locatie |
|-----|---------|
| pyRevit source (primair) | `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\pyrevit\` |
| pyRevit runtime | `%APPDATA%\pyRevit\Extensions\` |
| Logs | `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\pyrevit_logs\` |
| MCP servers source | `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\_MCP_servers\` |
| MCP servers runtime | `C:\MCP_servers\` |

### Sync commando
```powershell
$source = "X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension"
$runtime = "$env:APPDATA\pyRevit\Extensions\bouwkunde.extension"
Remove-Item $runtime -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item $source -Destination $runtime -Recurse -Force
```

---

## 7. NAA.K.T. Materiaalbenaming

### Format
```
naam_kenmerk_toepassing_[eigen-invulling]
```

### Voorbeelden
- `beton_gewapend_prefab-element_C30/37`
- `isolatie_pir_plaat_100mm`
- `hout_eiken_profiel_50x100`

### Data locatie
```
lib/naakt_data/
├── naakt_namen.json        # Hoofdgroepen
├── naakt_kenmerken.json    # Kenmerken per naam
└── naakt_toepassingen.json # Toepassingen per naam
```
