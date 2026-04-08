# GIS2BIM Icon Stijlgids v3
## Variant B — Teal + Yellow lijnen, transparante achtergrond

---

## 1. Kernprincipes

- **Alleen lijnen** — geen gevulde vlakken (fill), alles is `stroke` / `outline`
- **Transparante achtergrond** — geen achtergrondvlak, werkt op elke kleur
- **Twee kleuren** — Teal als hoofdkleur, Yellow als accent
- **Minimale fills** — alleen kleine dots/stippen mogen gevuld zijn

---

## 2. Kleurenpalet

| Rol | Naam | Hex | RGB | Gebruik |
|-----|------|-----|-----|---------|
| **Hoofd** | Verdigris (Teal) | `#45B6A8` | `(69, 182, 168)` | Alle hoofdvormen en lijnen |
| **Accent** | Friendly Yellow | `#EFBD75` | `(239, 189, 117)` | Markers, badges, meetlijnen, labels |
| **Optioneel** | Magic Violet | `#350E35` | `(53, 14, 53)` | Alleen als achtergrondvariant nodig is |

### Opacity levels
- **100%** — Hoofdvormen, primaire lijnen
- **40-60%** — Secundaire lijnen (ramen, rasters)
- **20-30%** — Achtergrond-echo's (terreinlijnen, arceringen)

---

## 3. Lijndikte

| Formaat | Hoofdlijn | Secundair | Details |
|---------|-----------|-----------|---------|
| **32×32** | 1.5–1.8 px | 1.0–1.2 px | 0.7–0.9 px |
| **96×96** | 2.0–2.5 px | 1.2–1.5 px | 0.8–1.0 px |

### Lijnstijl
- `stroke-linecap: round` — altijd afgeronde uiteinden
- `stroke-linejoin: round` — afgeronde hoeken
- Geen dashes behalve voor meetlijnen (AHN hoogteverschil)

---

## 4. Afmetingen

| Formaat | Gebruik | Canvas |
|---------|---------|--------|
| **32×32 px** | pyRevit ribbon buttons | Volledig canvas, geen padding nodig |
| **96×96 px** | Grote knoppen, previews | ~8px padding rondom |
| **192×192 px** | High-DPI / Retina | Opgeschaald van 96px |

---

## 5. SVG Templates

### Basis 32×32
```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">
  <!-- Geen achtergrond rect! Transparant canvas -->
  
  <!-- Hoofdvormen in Teal -->
  <!-- stroke="#45B6A8" stroke-width="1.5" fill="none" -->
  
  <!-- Accenten in Yellow -->
  <!-- stroke="#EFBD75" stroke-width="1.2" fill="none" -->
</svg>
```

### Map Pin (lijn-versie)
```svg
<path d="M16 3 C10 3 6 8 6 13 C6 21 16 29 16 29 C16 29 26 21 26 13 C26 8 22 3 16 3 Z" 
      fill="none" stroke="#45B6A8" stroke-width="1.8" stroke-linejoin="round"/>
<circle cx="16" cy="12.5" r="4" fill="none" stroke="#45B6A8" stroke-width="1.5"/>
```

### "2" Badge (lijn-versie)
```svg
<rect x="X" y="Y" width="24" height="20" rx="5" fill="none" stroke="#EFBD75" stroke-width="2"/>
<text x="X+12" y="Y+15" font-family="Segoe UI" font-size="14" 
      font-weight="700" fill="#EFBD75" text-anchor="middle">2</text>
```

### Meetpunt (circle outline + dot)
```svg
<circle cx="X" cy="Y" r="1.8" fill="none" stroke="#EFBD75" stroke-width="1.2"/>
<!-- Optioneel center dot: -->
<circle cx="X" cy="Y" r="0.7" fill="#EFBD75"/>
```

### Data lijnen (WFS-stijl)
```svg
<line x1="7" y1="27" x2="11.5" y2="27" stroke="#EFBD75" stroke-width="1.3" stroke-linecap="round"/>
<line x1="13.5" y1="27" x2="20" y2="27" stroke="#EFBD75" stroke-width="1.3" stroke-linecap="round"/>
<line x1="22" y1="27" x2="25" y2="27" stroke="#EFBD75" stroke-width="1.3" stroke-linecap="round"/>
```

---

## 6. Visueel Vocabulaire

| Element | Vorm | Kleur | Wanneer |
|---------|------|-------|---------|
| **Map pin** | Druppel-outline + cirkel | Teal | Locatie tools |
| **"2" badge** | Rounded rect outline + tekst | Yellow | Hoofd-icon, conversie |
| **Grid** | Rechthoek + binnenlijnen | Teal | Kadaster, percelen |
| **Terreinprofiel** | Polyline met echo | Teal + Yellow dots | AHN, hoogte |
| **Concentrische ringen** | 3 cirkels, afnemende opacity | Teal | API's, services |
| **Gebouw footprints** | Overlappende rechthoeken | Teal | BAG, bebouwing |
| **Zoning grid** | 4 vakken, 2 kleuren + arcering | Teal + Yellow | Bestemmingsplan |
| **NL silhouet** | Polygon outline | Teal + Yellow dots | PDOK, landelijke data |
| **Gebouw + dak** | Rechthoek + driehoek polyline | Teal | BIM/gebouw-gerelateerd |
| **Pijl** | Lijn + arrowhead | Yellow | Richting, conversie |

---

## 7. Bestaande Icon Set

| Bestand | Tool | Beschrijving |
|---------|------|--------------|
| `GIS2BIM_main` | Panel header | Pin met "i" + yellow "2" badge |
| `GIS_Locatie` | Locatie ophalen | Map pin outline |
| `Kadaster_Data` | Kadaster percelen | Grid met perceellijnen + marker |
| `GIS_naar_BIM` | GIS→BIM conversie | Gebouw + pijl vanuit pin |
| `AHN_Hoogte` | AHN hoogtedata | Terreinprofiel + meetpunten |
| `WFS_Service` | WFS API calls | Concentrische ringen + data lijnen |
| `BAG_Data` | BAG gebouwen | Overlappende footprints |
| `Bestemmingsplan` | Bestemmingsplan | Gekleurde zone grid + arcering |
| `PDOK_Data` | PDOK services | NL silhouet + datapunten |

---

## 8. Checklist Nieuw Icon

- [ ] Transparante achtergrond (geen rect fill)
- [ ] Alle vormen als `stroke` / `outline`, geen `fill` (behalve kleine dots)
- [ ] Hoofdvorm in Teal (`#45B6A8`)
- [ ] Max 1 accent in Yellow (`#EFBD75`)
- [ ] Lijndikte 1.5–1.8px voor 32px icons
- [ ] `stroke-linecap: round` en `stroke-linejoin: round`
- [ ] Leesbaar op 32×32 pixels
- [ ] Consistent met bestaande set

---

## 9. PNG Generatie

Icons worden gegenereerd met Pillow (Python) op 10× supersampling + LANCZOS downscale:

```python
SCALE = 10
img = Image.new("RGBA", (32*SCALE, 32*SCALE), (0,0,0,0))  # Transparant!
# ... draw at high res ...
result = img.resize((32, 32), Image.LANCZOS)
result.save("icon.png")
```

### Prompt voor Claude
```
Maak een nieuw GIS2BIM v3 icon (32×32 en 96×96 PNG, transparant) voor "{TOOL_NAAM}".

Stijl: Alleen lijnen, geen fills. Teal (#45B6A8) hoofdlijnen, 
Yellow (#EFBD75) accenten. Transparante achtergrond.
Lijndikte 1.5-1.8px op 32px. stroke-linecap: round.

Het icon moet {BESCHRIJVING} uitbeelden.

Genereer met Pillow (10x supersampling + LANCZOS downscale).
```
