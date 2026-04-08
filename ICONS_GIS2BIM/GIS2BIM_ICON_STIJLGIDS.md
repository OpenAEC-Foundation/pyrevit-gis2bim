# GIS2BIM Icon Stijlgids
## 3BM Bouwkunde — Stijl A (Violet base, Teal accent)

---

## 1. Kleurenpalet

| Rol | Naam | Hex | RGB | Gebruik |
|-----|------|-----|-----|---------|
| **Achtergrond** | Magic Violet | `#350E35` | `(53, 14, 53)` | Altijd de icon achtergrond |
| **Primair** | Verdigris (Teal) | `#45B6A8` | `(69, 182, 168)` | Hoofdvormen, lijnen, symbolen |
| **Accent** | Friendly Yellow | `#EFBD75` | `(239, 189, 117)` | Badges, markers, highlights |
| **Tekst** | White | `#FFFFFF` | `(255, 255, 255)` | Labels, tekst op violet |
| **Alert** | Warm Magenta | `#A01C48` | `(160, 28, 72)` | Alleen bij waarschuwings-iconen |
| **Error** | Flaming Peach | `#DB4C40` | `(219, 76, 64)` | Alleen bij error-iconen |

### Kleurregels
- **Achtergrond**: Altijd Violet (`#350E35`), nooit transparant
- **Icoon-elementen**: Teal als hoofdkleur, max 2 extra kleuren per icon
- **Yellow**: Spaarzaam — alleen voor markers, badges, of punten van aandacht
- **White**: Alleen voor tekst of subtiele lijnen (met opacity)
- **Geen gradients** in de basis-iconen (wel toegestaan voor marketing-varianten)

---

## 2. Afmetingen & Grid

### Standaard formaten

| Formaat | Gebruik | Hoekradius |
|---------|---------|------------|
| **32×32 px** | pyRevit ribbon buttons | `rx="3"` |
| **96×96 px** | Grote knoppen, panel headers | `rx="8"` |
| **192×192 px** | High-DPI / Retina versie van 96px | `rx="16"` |

### Veilige zone (padding)
- **32×32**: Minimaal 4px padding aan alle zijden → tekenvlak 24×24
- **96×96**: Minimaal 12px padding aan alle zijden → tekenvlak 72×72
- Tekst/labels mogen buiten het tekenvlak, maar binnen 2px van de rand

### Lijndikte
- **32×32**: Outlines `1.5px`, details `1px`
- **96×96**: Outlines `2px`, details `1.5px`

---

## 3. Bestandsnaamconventie

```
{ToolNaam}_{formaat}.png
```

### Voorbeelden
```
GIS_Locatie_32.png      → 32×32 ribbon icon
GIS_Locatie_96.png      → 96×96 groot icon
Kadaster_Data_32.png    → 32×32 ribbon icon
```

### Voor pyRevit
Plaats het icon als `icon.png` (32×32) in de `.pushbutton` map:
```
GIS2BIM.panel/
├── GISLocatie.pushbutton/
│   ├── script.py
│   ├── icon.png          ← 32×32 versie
│   └── bundle.yaml
```

---

## 4. SVG Templates

### Template: 32×32 basis
```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <!-- Achtergrond: altijd Violet met afgeronde hoeken -->
  <rect width="32" height="32" rx="3" fill="#350E35"/>
  
  <!-- === ICON CONTENT HIER === -->
  <!-- Gebruik Teal (#45B6A8) voor hoofdvormen -->
  <!-- Gebruik Yellow (#EFBD75) voor accenten -->
  <!-- Gebruik White (#FFFFFF) voor tekst/labels -->
  
</svg>
```

### Template: 96×96 basis
```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96" width="96" height="96">
  <!-- Achtergrond -->
  <rect width="96" height="96" rx="8" fill="#350E35"/>
  
  <!-- === ICON CONTENT HIER === -->
  
</svg>
```

### Template: Map Pin (GIS2BIM kenmerk)
```svg
<!-- Map pin met "i" — het GIS2BIM signatuur-element -->
<!-- cx, cy = center positie van de pin -->
<g transform="translate(cx, cy)">
  <!-- Pin body -->
  <path d="M0 0 C-6 0 -9 4 -9 9 C-9 16 0 21 0 21 
           C0 21 9 16 9 9 C9 4 6 0 0 0 Z" fill="#45B6A8"/>
  <!-- Binnenste cirkel -->
  <circle cx="0" cy="8.5" r="3.5" fill="#350E35"/>
  <!-- Optioneel: "i" in de cirkel -->
  <text x="0" y="10.5" font-family="Segoe UI" font-size="5" 
        font-weight="bold" fill="#45B6A8" text-anchor="middle">i</text>
</g>
```

### Template: "2" Badge
```svg
<!-- Gele badge met "2" — verwijzing naar GIS→BIM -->
<rect x="X" y="Y" width="22" height="20" rx="4" fill="#EFBD75"/>
<text x="X+11" y="Y+15" font-family="Segoe UI" font-size="16" 
      font-weight="900" fill="#350E35" text-anchor="middle">2</text>
```

### Template: Tekst label
```svg
<!-- Tekst onderaan het icon (bijv. "AHN", "WFS", "BAG") -->
<text x="16" y="27" font-family="Segoe UI" font-size="6" 
      fill="white" text-anchor="middle" opacity="0.8">LABEL</text>
```

---

## 5. Design Principes

### Herkenbare elementen
Elk GIS2BIM icon deelt visuele DNA:

| Element | Hoe | Wanneer |
|---------|-----|---------|
| **Map pin** | Teal pin met violet dot | Locatie-gerelateerde tools |
| **"2" badge** | Geel rondje/rect met "2" | Hoofd-icon, conversie tools |
| **Kadaster grid** | Teal lijnen, opgedeeld vlak | Perceeldata, bestemmingsplan |
| **Terreinlijnen** | Golvende teal lijnen | AHN, hoogte, topografie |
| **Signaal-ringen** | Concentrische cirkels | API's, services, connecties |
| **Gebouw silhouet** | Eenvoudig huis met dak | BIM-kant van de tools |

### Stijlregels

1. **Simpliciteit** — Elk icon heeft max 3-4 visuele elementen
2. **Leesbaarheid op 32px** — Als het niet leesbaar is op 32×32, vereenvoudig
3. **Consistent gewicht** — Vergelijkbare lijndikte en vulling over alle iconen
4. **Geen emoji/unicode** — Gebruik alleen SVG paden en tekst (IronPython compatibiliteit)
5. **Altijd violet achtergrond** — Nooit transparant, nooit een andere kleur

### Opacity richtlijnen
- Achtergrond-elementen (kaartlijnen, rasters): `opacity="0.2"` tot `0.4`
- Secundaire vormen: `opacity="0.5"` tot `0.7`  
- Primaire vormen: Volle kleur (geen opacity)
- Tekst: `opacity="0.8"` of volle kleur

---

## 6. Checklist Nieuw Icon

Bij het maken van een nieuw GIS2BIM icon:

- [ ] Violet achtergrond (`#350E35`) met correcte hoekradius
- [ ] Hoofdvorm in Teal (`#45B6A8`)
- [ ] Max 1 accent kleur (Yellow `#EFBD75` of White)
- [ ] Leesbaar op 32×32 pixels
- [ ] Consistent met bestaande icon set
- [ ] Bestandsnaam volgt conventie: `{ToolNaam}_{formaat}.png`
- [ ] 32px versie als `icon.png` in pushbutton map

---

## 7. Bestaande Icon Set

| Icon | Tool | Hoofdelement | Accent |
|------|------|-------------|--------|
| GIS2BIM_main | Panel header | Map pin + "i" | Yellow "2" badge |
| GIS_Locatie | Locatie ophalen | Map pin | — |
| Kadaster_Data | Kadaster/percelen | Grid met parcelen | Yellow marker |
| GIS_naar_BIM | GIS→BIM conversie | Gebouw + pin | Yellow verbindingslijn |
| AHN_Hoogte | AHN hoogtedata | Terreinlijnen | Yellow meetpunten |
| WFS_Service | WFS API calls | Signaal-ringen | Yellow arcs |
| BAG_Data | BAG gebouwen | Gebouw footprints | Yellow label |
| Bestemmingsplan | Bestemmingsplan | Gekleurde zones | Teal/Yellow/Magenta |

---

## 8. Prompt voor Claude

Gebruik deze prompt om Claude nieuwe icons te laten genereren in dezelfde stijl:

```
Maak een nieuw GIS2BIM pyRevit icon (32×32 en 96×96 PNG) voor de tool "{TOOL_NAAM}".

Stijl: Violet achtergrond (#350E35), Teal hoofdvormen (#45B6A8), 
Yellow accenten (#EFBD75). Hoekradius 3px (32px) / 8px (96px).

Het icon moet {BESCHRIJVING_VAN_WAT_HET_MOET_VOORSTELLEN} uitbeelden.

Gebruik de GIS2BIM stijlgids en genereer met Pillow (8x supersampling + LANCZOS downscale).
```
