# 3BM Bouwkunde - Architectuur & Structuur

## 1. pyRevit Extension Structuur

```
bouwkunde.extension/
├── extension.json              # Extensie metadata
├── lib/                        # Gedeelde modules
│   ├── __init__.py
│   ├── ui_template.py          # Windows Forms UI framework (legacy)
│   ├── wpf_template.py         # WPF UI framework (AANBEVOLEN voor nieuwe tools)
│   ├── bm_logger.py            # Centrale logging
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
    │   ├── RcBerekening.pushbutton/
    │   └── VentilatieBalans.pushbutton/
    │
    ├── Document.panel/         # Document/sheet management
    │   └── SheetParameters.pushbutton/
    │
    ├── Fundering.panel/        # Constructief - fundering
    │   └── PalenNummeren.pushbutton/
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
        ├── UITest.pushbutton/      # Windows Forms test
        └── UITestWPF.pushbutton/   # WPF test (REFERENTIE)
```

---

## 2. Nieuwe Tool Maken - WPF (AANBEVOLEN)

WPF biedt declaratieve XAML layouts, automatische DPI scaling, en betere data binding.
Gebruik WPF voor alle nieuwe tools.

### Stap 1: Mapstructuur aanmaken
```powershell
$base = "X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension\Bouwkunde.tab"
New-Item -Path "$base\NieuwPanel.panel\NieuweTool.pushbutton" -ItemType Directory -Force
```

### Stap 2: UI.xaml maken (layout)
```xml
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Tool Titel"
    Width="500"
    SizeToContent="Height"
    WindowStartupLocation="CenterScreen"
    Background="White">
    
    <Window.Resources>
        <!-- 3BM Huisstijl kleuren -->
        <SolidColorBrush x:Key="VioletBrush" Color="#350E35"/>
        <SolidColorBrush x:Key="TealBrush" Color="#45B6A8"/>
        <SolidColorBrush x:Key="YellowBrush" Color="#EFBD75"/>
        <SolidColorBrush x:Key="PeachBrush" Color="#DB4C40"/>
        <SolidColorBrush x:Key="LightGrayBrush" Color="#F5F5F5"/>
        <SolidColorBrush x:Key="MediumGrayBrush" Color="#E0E0E0"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#808080"/>
    </Window.Resources>
    
    <DockPanel LastChildFill="True">
        <!-- Header -->
        <StackPanel DockPanel.Dock="Top" Background="{StaticResource VioletBrush}">
            <TextBlock Text="Tool Titel" FontSize="18" FontWeight="SemiBold" 
                       Foreground="White" Margin="20,16,20,4"/>
            <TextBlock x:Name="txt_subtitle" Text="Subtitel" FontSize="12" 
                       Foreground="{StaticResource TealBrush}" Margin="20,0,20,16"/>
        </StackPanel>
        
        <!-- Accent Line -->
        <Rectangle DockPanel.Dock="Top" Height="3" Fill="{StaticResource TealBrush}"/>
        
        <!-- Footer -->
        <Border DockPanel.Dock="Bottom" BorderBrush="{StaticResource MediumGrayBrush}" 
                BorderThickness="0,1,0,0" Background="{StaticResource LightGrayBrush}" Padding="20,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btn_cancel" Content="Annuleren" Padding="20,8" Margin="0,0,8,0"/>
                <Button x:Name="btn_execute" Content="Uitvoeren" Padding="20,8"
                        Background="{StaticResource TealBrush}" Foreground="White" 
                        FontWeight="SemiBold" BorderThickness="0"/>
            </StackPanel>
        </Border>
        
        <!-- Content -->
        <ScrollViewer Padding="20">
            <StackPanel>
                <!-- Jouw controls hier -->
                <TextBlock Text="SECTIE TITEL" FontSize="11" FontWeight="SemiBold" 
                           Foreground="{StaticResource VioletBrush}" Margin="0,0,0,12"/>
                
                <Grid Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2*"/>
                        <ColumnDefinition Width="16"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <ComboBox x:Name="cmb_selectie" Grid.Column="0" Padding="8,6"/>
                    <TextBox x:Name="txt_waarde" Grid.Column="2" Padding="8,6"/>
                </Grid>
            </StackPanel>
        </ScrollViewer>
    </DockPanel>
</Window>
```

### Stap 3: script.py maken
```python
# -*- coding: utf-8 -*-
"""Tool beschrijving"""
__title__ = "Tool\nNaam"
__author__ = "3BM Bouwkunde"

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows import Window
from System.Windows.Markup import XamlReader

from pyrevit import revit, forms, script
import os
import sys

# Lib path
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from bm_logger import get_logger

log = get_logger("NieuweTool")
doc = revit.doc


class NieuweToolWindow(Window):
    def __init__(self):
        Window.__init__(self)
        self._load_xaml()
        self._bind_events()
    
    def _load_xaml(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        with open(xaml_path, 'r') as f:
            xaml = f.read()
        
        reader = StringReader(xaml)
        loaded = XamlReader.Load(SysXmlReader.Create(reader))
        
        # Copy window properties
        self.Title = loaded.Title
        self.Width = loaded.Width
        self.SizeToContent = loaded.SizeToContent
        self.WindowStartupLocation = loaded.WindowStartupLocation
        self.Content = loaded.Content
        
        # Bind named elements
        self.btn_cancel = loaded.FindName('btn_cancel')
        self.btn_execute = loaded.FindName('btn_execute')
        self.cmb_selectie = loaded.FindName('cmb_selectie')
        self.txt_waarde = loaded.FindName('txt_waarde')
    
    def _bind_events(self):
        self.btn_cancel.Click += self._on_cancel
        self.btn_execute.Click += self._on_execute
    
    def _on_cancel(self, sender, args):
        self.Close()
    
    def _on_execute(self, sender, args):
        log.info("Uitvoeren geklikt")
        # Doe iets met self.cmb_selectie.SelectedItem, self.txt_waarde.Text
        self.Close()


def main():
    log.info("Tool gestart")
    window = NieuweToolWindow()
    window.ShowDialog()


if __name__ == "__main__":
    main()
```

### Stap 4: bundle.yaml + icon.svg
```yaml
title: Tool Naam
tooltip: Beschrijving
author: 3BM Bouwkunde
highlight: new
context: zero-doc
```

### Stap 5: Sync en reload
Zie sectie 4 hieronder.

---

## 3. Nieuwe Tool Maken - Windows Forms (Legacy)

> **Let op:** Gebruik WPF voor nieuwe tools. Windows Forms alleen voor onderhoud bestaande tools.

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

## 4. Lib Modules

### wpf_template.py (AANBEVOLEN)

WPF UI framework met 3BM huisstijl. Zie `UITestWPF.pushbutton` voor werkend voorbeeld.

**Belangrijke patterns:**
- XAML voor layout (UI.xaml)
- Python voor logic (script.py)
- `FindName()` voor element binding
- Resources voor herbruikbare kleuren

### ui_template.py (Legacy)

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

### naakt_data/

NAA.K.T. (Naam-Attribuut-Kenmerk-Toepassing) standaard voor materiaalbenaming:
- `naakt_namen.json` - Hoofdgroepen (beton, hout, isolatie, etc.)
- `naakt_kenmerken.json` - Kenmerken per naam
- `naakt_toepassingen.json` - Toepassingen per naam

---

## 5. Multi-Device Workflow

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
2. Sync naar runtime na wijzigingen
3. Logs beschikbaar in `X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\pyrevit_logs\`

**Laptop sessie:**
1. OneDrive synct automatisch
2. Zelfde workflow als PC

**Telefoon sessie:**
1. Alleen planning/discussie
2. Geen code wijzigingen mogelijk

---

## 6. Extension.json

```json
{
    "type": "extension",
    "name": "3BM Bouwkunde",
    "description": "pyRevit tools voor constructief ontwerp en bouwfysica",
    "author": "3BM Bouwkunde",
    "version": "1.0.0",
    "rocket_mode_compatible": true
}
```

---

## 7. Troubleshooting

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
