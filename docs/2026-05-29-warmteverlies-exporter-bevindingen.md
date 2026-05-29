# Warmteverlies-exporter — bevindingen & werkpakket (2026-05-29)

**Datum:** 2026-05-29
**Model:** `2786_Bouwkundige_model` (host) + gelinkt `2381_model R25` (constructie)
**Methode:** empirische sessie via Revit MCP — gemeten op echte SEGC-geometrie, niet op code-trace.
**Vervangt/actualiseert:** delen van `docs/2026-05-22-warmteverlies-scanner-leefruimte-werkpakket.md`
(Bug #3 root-cause + curtain-wall + area zijn hier herzien op basis van runtime-bewijs).

> **Kernconclusie:** de exporter levert voor dit model fundamenteel onjuiste **volumes** en zwaar
> **gefragmenteerde** constructies. Beide oorzaken zijn hard gediagnosticeerd met exacte
> `bestand:regel`-locaties. De host/link laag-verdubbeling is deze sessie al in code opgelost.

---

## 1. Wat is er deze sessie aan het MODEL gewijzigd

Door de gebruiker / via MCP, blijvend in het Revit-model:

| Wijziging | Detail | Waarom |
|-----------|--------|--------|
| **Room Limit Offset → 5,0 m** (alle geplaatste rooms) | was 2,44 m / 3,0 m | rooms reikten niet tot het echte dak (kap loopt tot ~4,12 m) → SEGC kapte af op de room-top |
| Workset `massa_rooms_model` aangemaakt | door gebruiker | filteren van de verificatie-massa's |
| DirectShapes geplaatst (Generic Models, op die workset) | `WV_MASS` ×18 (room-massa's), `WV_GRP` ×12 (host-groep-prototype Leefruimte), `WV_CMP` ×2 (contour-vergelijk) | visuele verificatie |

> **Let op bij opruimen:** de offset-wijziging raakte óók de twee "Buiten"-rooms (50, 70) → die
> hebben nu enorme volumes (2336 / 3744 m³). Niet-verwarmd, vallen buiten de export. De `WV_*`
> DirectShapes zijn wegwerp-verificatie; te verwijderen via filter op Comments `WV_`.

---

## 2. De zes bevindingen (geprioriteerd)

### 🔴 #1 — Geëxporteerd volume = `area × nominale hoogte` (60-80% te groot)

- **Locatie:** `lib/warmteverlies/thermal_json_builder.py:192-194`
  ```python
  area_m2 = round(room_data.get("floor_area_m2", 0.0), 2)
  height_m = round(room_data.get("height_m", 2.6), 2)   # = LimitOffset (unit_utils.py:130)
  volume_m3 = round(area_m2 * height_m, 2)               # area × NOMINALE hoogte
  ```
- **Oorzaak:** `get_room_height` (`unit_utils.py:109-132`) geeft de **LimitOffset** terug, niet het
  werkelijke volume. De exporter raakt `Room.Volume` nooit aan.
- **Bewijs (na offset→5 m):**

  | Room | geëxporteerd | echte `Room.Volume` | fout |
  |------|-------------:|--------------------:|-----:|
  | 17 Leefruimte | 174,05 | 108,58 | +60% |
  | 2. Gang | 53,45 | 29,93 | +79% |
  | 12. Toilet | 8,50 | 4,99 | +70% |

- **Contra-intuïtief:** de offset naar 5 m zetten **verbeterde** `Room.Volume` (klipt correct tegen
  het dak) maar **verslechterde** de export (area × 5,0 i.p.v. de echte 108,58).
- **Fix:**
  1. `room_collector.py:49-60` — voeg toe: `"volume_m3": internal_to_cubicm(room.Volume)`
     (`Room.Volume`, ft³→m³ ×0,0283168).
  2. `thermal_json_builder.py:192-194` — gebruik `Room.Volume`:
     ```python
     volume_m3 = round(room_data.get("volume_m3", 0.0), 2)
     if volume_m3 <= 0:                                   # volumeberekening uit
         volume_m3 = round(area_m2 * height_m, 2)          # fallback
     height_m = round(volume_m3 / area_m2, 2) if area_m2 > 0 else height_m  # effectieve hoogte
     ```
  3. Vereist dat Revit's **Areas and Volumes → Volumes** aan staat (in dit model: AAN). Bij UIT
     valt de fallback in → markeer als benadering.

### 🟢 #2 — Host/link laag-verdubbeling — OPGELOST in code (uncommitted)

- **Oorzaak (runtime-bewijs, weerlegt de 22-05-hypothese over enter/exit-references):** de
  constructie staat zowel in het **host-model** áls in de gelinkte `2381_model R25`, op identieke
  positie. De raycast (`FindReferencesInRevitLinks=True`) raakt elk pakket twee keer →
  `2× hout_vuren 160mm`, `2× Holz 120mm`.
- **SEGC-bewijs:** 18 van 19 room-boundaries van de Leefruimte zijn **host**-elementen (host is
  room-bounding/compleet). Slechts 1 vlak (0,56 m²) is link-only. Dus links negeren in de raycast
  verwijdert de duplicaat én behoudt de volledige constructie.
- **Fix (gedaan):** `constants.py` → `SCAN_LINKED_MODELS = False`; `raycast_scanner.py:566` gebruikt
  die flag i.p.v. hardcoded `True`. Geverifieerd: `2× Holz 120` → enkel.
- **⚠️ Default-besluit nog open:** voor dit model is `False` correct, maar de scanner is ontworpen
  voor cross-model (constructie in link, architectuur in host) — daar is `True` nodig. **Committen
  als default `False` zou andere projecten breken.** Opties: (a) default `True` houden + per-project
  override, (b) auto-dedup van samenvallende host/link-hits in `_hits_to_layer_stack`
  (robuuster, werkt in beide gevallen). Aanbeveling: (b) als duurzame fix, (a) als interim.

### 🟡 #3 — 147 constructies voor 18 rooms (zware fragmentatie)

- **Bewijs:** volledige export gaf 147 constructies; Leefruimte alleen al fragmenteert in tot
  21 stukken. SEGC hakt één plafond-met-koof in 47 vlakken (waarvan 21 < 0,25 m²; 23 verticale
  sliver-vlakjes zonder host = koof-randen).
- **Gevalideerde oplossing — host-element-groepering (2-traps):**
  1. **Geometrisch:** groepeer vlakken op `(room, oriëntatie, host_element_id)`, som de **echte
     SEGC-area** (= meteen Fix #4 uit het 22-05-werkpakket), scan de laagopbouw **1× per element**.
     Het veld `host_element_id` wordt al per face geëxtraheerd in de WIP-code.
  2. **Thermisch:** laat de bestaande builder-fingerprint `(room, oriëntatie, compass, lagen)`
     daarna elementen met identieke opbouw samenvoegen.
  - **Resultaat (gemeten, Leefruimte):** 47 vlakken → 12 groepen → na thermische merge ~6 nette
    constructies. De 23 koof-slivers verdwijnen (verticaal + geen host = al overgeslagen sinds
    commit `39c6768`).
- **Locatie:** `_get_faces_from_segc` (groepering) + `_build_constructions`
  (`thermal_json_builder.py:~356-428`, 2e trap bestaat al).

### 🟡 #4 — `boundary_polygon` wordt nooit geëxporteerd

- **Bewijs:** 0 van 19 rooms in de JSON hebben `boundary_polygon` (schema-veld bestaat, optioneel —
  gebruikt voor plattegronden in open-heatloss-studio).
- **Beschikbaar:** `Room.GetBoundarySegments` (Finish) geeft de exacte contour; geverifieerd dat
  shoelace-area = `Room.Area` (Toilet 1,66 vs 1,65 m²). Alle segmenten recht (0 bogen) in dit model.
- **Fix:** extract in `collect_rooms`, schrijf naar room-JSON in builder.

### 🟡 #5 — Vliesgevel (curtain wall) → onzin-laagstacks (Bug #2 22-05, nog open)

- **Bewijs:** Leefruimte-export gaf constructies met `kunststof 1033mm`, `1728mm` — de raycast door
  een curtain wall (géén CompoundStructure) produceert rommel. Openingen voor de vliesgevel = 0.
- **Onderscheid:** `28_CWA_glas15mm` = gewone Wall mét glas-CompoundStructure → blijft wand-
  constructie. `28_CW_reynaers...` = échte curtain wall (`Wall.CurtainGrid is not None`) → moet
  `curtain_wall`-opening + host-constructie worden.
- **Synergie:** de host-element-groepering (#3) zet dit klaar — de vliesgevel valt uiteen in 3
  discrete curtain-wall-elementen (samen 55 m²) → 3 `curtain_wall`-openingen.
- **Detectie/emit-ontwerp:** zie 22-05-werkpakket §4 Fix #2 (ongewijzigd geldig).

### ⚪ #6 — "Rooms te laag getekend" — modelvereiste, geen silent fallback

- **Bewijs:** bij offset 3,0 m kapte de Leefruimte-massa af op een vlak top-vlak (27,98 m² @ z=3,0)
  terwijl de echte kap tot 3,96 m loopt → ~4,8 m³ volume gemist.
- **Besluit:** géén auto-raise heuristiek (maskeert modelfouten). Wél: (a) modelvereiste "rooms
  reiken tot de bovenliggende constructie", (b) **validatie-waarschuwing** in de export die rooms
  flagt waarvan de SEGC-top geen room-bounding element heeft of `Room.Volume / Area`
  verdacht laag is t.o.v. de verdiepingshoogte.

---

## 3. Contour-versimpeling (los van bovenstaande, optioneel)

Gebruikerswens: kleine in/uitspringingen (inhammetjes) rechttrekken zoals bij handmatige
berekeningen. **Een area-filter is hiervoor het verkeerde gereedschap** (gooit area wég i.p.v.
recht te trekken). Juiste aanpak: **boundary-polygon simplificatie met tolerantie**.

- **Gemeten (Leefruimte, 11 hoekpunten):**

  | Tolerantie | hoekpunten | area | delta |
  |-----------|:----------:|-----:|------:|
  | origineel | 11 | 34,81 | — |
  | **100 mm** (iteratief DP) | 8 | 34,61 | −0,20 m² (0,6%) |
  | 200 mm (naïef) | 4 | 23,27 | −11,54 m² ❌ sloopt L-hoek |

- **Besluit:** tolerantie ~**100 mm**, met een **robuust iteratief Douglas-Peucker** (verwijder
  telkens de minst-significante hoek, herbereken) — NIET de naïeve simultane variant (slaat door
  bij ≥200 mm). **Log** wat rechtgetrokken wordt ("N jogs < 100 mm, area −0,20 m²") — aanname
  zichtbaar houden.
- **Scope:** drijft vloer/plafond-area + `boundary_polygon`. De inham-wandjes worden al door de
  host-groepering (#3) samengevoegd. Optioneel/configureerbaar; geen blocker.

---

## 4. Implementatie-volgorde (voor delegatie)

Alles in `lib/warmteverlies/`, IronPython 2.7 (geen f-strings/type hints, feet ×304,8 = mm,
Transaction-wrapping). Delegeren via `general-purpose` agent (specialized write-agents
backgrounden/falen — zie orchestrator-memory).

1. **#1 Volume** — geïsoleerd, hoogste impact, klein. `collect_rooms` + `thermal_json_builder`.
2. **#3 Host-element-groepering** (incl. echte SEGC-area = oude Fix #4) — grootste kwaliteitswinst,
   raakt `_get_faces_from_segc` + `_build_constructions`. Sequentieel ná #1.
3. **#4 boundary_polygon** — `collect_rooms` + builder + 100 mm DP-helper.
4. **#5 Curtain wall** — bouwt voort op #3.
5. **#2 default-besluit** — kies (a) of (b); committen.
6. **#6 validatie-waarschuwing** — laagste prioriteit.

**Verificatie per stap:** pyRevit reload → export Room 66 (Leefruimte) → vergelijk met deze waarden
(volume 108,58 · ~6 constructies · vliesgevel als curtain_wall-openingen · contour 8 hoekpunten).

---

## 5. Revit MCP — herbruikbare valkuilen (deze sessie ontdekt)

- `from Autodesk.Revit.DB import X` **faalt** in `execute_revit_code` (error "Name") → gebruik de
  geïnjecteerde **`DB.*`** namespace.
- **`def` binnen MCP-code** ziet de geïnjecteerde globals (`DB`, `doc`) niet → NameError. Inline
  alle logica, geen functie-definities.
- `IList[BoundarySegment]` heeft geen `.Size` → gebruik `len()`.
- IronPython is streng met `"{:.2f}".format(int)` ("Precision not allowed in integer format
  specifier") → cast expliciet naar `float()`.
- Module-reload in MCP: `for m in list(sys.modules): if m.startswith("warmteverlies"): del
  sys.modules[m]` vóór her-import, anders draait gecachte code.
- Lib op het pad zetten: `sys.path.append(r"...\bouwkunde.extension\lib")` (MCP-context heeft de
  extensie-lib niet standaard).

---

## 6. Status van de code (einde sessie)

**Uncommitted in `pyrevit-gis2bim` (working tree):**
- 639 r WIP uit eerdere sessie: open_connections (Fix #1) werkt, debug-infra, subface-dedup.
- `DEBUG_SCANNER = True` (productie-export krijgt `_debug_scanner`-blok — uitzetten vóór release).
- `SCAN_LINKED_MODELS = False` (deze sessie; zie #2 — default-besluit open).

Deze code is **bewust niet gecommit** (Fix #2/#4 onafgemaakt, toggle-default nog te besluiten).
Alleen dit document is gecommit.
