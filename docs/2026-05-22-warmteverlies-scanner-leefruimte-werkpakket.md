# Werkpakket — Warmteverlies-scanner reparatie (leefruimte-testcase)

**Datum:** 2026-05-22
**Status:** Diagnose compleet — klaar voor implementatie
**Scope:** `extensions/bouwkunde.extension/lib/warmteverlies/raycast_scanner.py` + `thermal_json_builder.py`
**Aanleiding:** De "Warmteverlies Export" pushbutton produceert voor de moeilijkste ruimte (Room 17 "Leefruimte" in model `2786_Bouwkundige_model`) een sterk onvolledige Thermal Import JSON.
**Gerelateerd:** `TODO.md` → sectie "Warmteverlies-exporter — code-audit 2026-05-22" (dit werkpakket detailleert en vervangt de audit-items B3, B4 en voegt curtain-wall + subface-dedupe toe).

---

## 1. Samenvatting

De scanner exporteert voor de leefruimte slechts 4 constructies, 0 openingen en 0 open verbindingen, terwijl de ruimte in werkelijkheid 16 SEGC-vlakken heeft met o.a. ~50 m² vliesgevel en 2 open verbindingen. Vier bugs zijn gediagnosticeerd en op echte geometrie gereproduceerd. Dit werkpakket bevat per bug de root-cause (met `bestand:regel`) en het afgestemde fix-ontwerp.

**Geen schuine-vlakken-probleem:** het dak komt binnen met n.Z = 0,99–1,00 (~8° helling) en de vliesgevel staat perfect verticaal (n.Z = 0,00). De wand/dak-classificatie struikelt hier niet — sloped-surface-ondersteuning is buiten scope.

---

## 2. Testcase — Room 17 "Leefruimte"

| Eigenschap | Waarde |
|---|---|
| Model | `2786_Bouwkundige_model` |
| Room ElementId | `2502243` |
| Naam in model | `17. Ieeftuimte` (typefout van "Leefruimte") · nummer 50→68-reeks: nr. 66 |
| Area | 34,81 m² |
| Volume | 104,20 m³ |
| Upper Limit / Offset | `00_beganegrond = Peil` + 3,00 m (door gebruiker opgehoogd van 2,44 m) |
| Bijzonderheden | 2 open verbindingen (Room Separation Lines), veel vliesgevel (curtain wall), nagenoeg-vlak dak (~8°) |

**Test-exports (op bureaublad gebruiker, 22-05):**
- `2786_Bouwkundige_model_thermal_import.json` — volledig model, 16 ruimten
- `2786_Bouwkundige_model_thermal_import_leefruimte.json` — alleen leefruimte, Upper Limit 3,0 m

### 2.1 SEGC-vlakken van de leefruimte (gemeten via Revit MCP, 22-05, Upper Limit 3,0 m)

```
[ 0] n.Z=+1.00 HORIZ  area= 28.98 m2  subfaces=0  host=-                          (vlakke doos-top, geen room-bounding element)
[ 1] n.Z=-1.00 HORIZ  area= 34.81 m2  subfaces=2  host=120-03s-1-TL               (vloer — 2 subfaces, ZELFDE element)
[ 2] n.Z= 0.00 VERT   area=  0.73 m2  subfaces=1  host=060-03s-1-TL               (dichte wand)
[ 3] n.Z= 0.00 VERT   area= 15.19 m2  subfaces=1  host=Model Lines                (Room Separation Line -> open verbinding 1)
[ 4] n.Z= 0.00 VERT   area=  3.52 m2  subfaces=1  host=Model Lines                (Room Separation Line -> open verbinding 2)
[ 5] n.Z= 0.00 VERT   area=  0.55 m2  subfaces=1  host=28_CWA_glas15mm            (glaswand — gewone Wall met 15mm-glaslaag)
[ 6] n.Z= 0.00 VERT   area=  9.55 m2  subfaces=1  host=28_CWA_glas15mm            (glaswand)
[ 7] n.Z= 0.00 VERT   area= 13.96 m2  subfaces=1  host=28_CWA_Patio glass...      (CURTAIN WALL)
[ 8] n.Z= 0.00 VERT   area= 26.70 m2  subfaces=1  host=...stijlen_10000 NM        (CURTAIN WALL)
[ 9] n.Z= 0.00 VERT   area=  9.34 m2  subfaces=1  host=28_CWA_Patio glass...      (CURTAIN WALL)
[10] n.Z= 0.00 VERT   area=  0.37 m2  subfaces=1  host=28_CWA_Patio glass...      (CURTAIN WALL)
[11] n.Z= 0.00 VERT   area=  0.37 m2  subfaces=0  host=-                          (geen grens)
[12] n.Z= 0.00 VERT   area=  0.56 m2  subfaces=1  host=? (gelinkt model)
[13] n.Z=+1.00 HORIZ  area=  3.27 m2  subfaces=1  host=160-05s-1-TL               (dak)
[14] n.Z=+0.99 HORIZ  area=  0.61 m2  subfaces=1  host=160-05s-1-TL               (dak, ~8°)
[15] n.Z=+0.99 HORIZ  area=  1.97 m2  subfaces=1  host=160-05s-1-TL               (dak, ~8°)
```

> Curtain walls = `OST_Walls`-elementen mét `CurtainGrid` en zónder `CompoundStructure`. De `28_CWA_glas15mm`-wanden zijn gewone Walls met een glas-`CompoundStructure` — die exporteren wél correct en blijven wand-constructies.

### 2.2 Wat de export oplevert vs. wat het zou moeten zijn

| | Geëxporteerd (leefruimte-JSON) | Verwacht |
|---|---|---|
| Constructies | 4 | dak + vloer + dichte wand(en) + glaswanden |
| Openingen | **0** | ~4 `curtain_wall`-openingen (vlakken 7–10) |
| Open verbindingen | **0** | **2** (vlakken 3 + 4) |
| Vliesgevel ~50 m² | ontbreekt grotendeels | aanwezig als openingen |
| Lagen dak/vloer | **verdubbeld** (`2× hout_vuren 160`, `2× Holz 120`) | enkelvoudig |

Geëxporteerde constructies:
- `constr-0` ceiling 34,83 m² — lagen `2× hout_vuren_generiek 160mm` ← **verdubbeld**
- `constr-1` floor 32,77 m² — lagen `2× Holz 120mm | n7_isolatie_resol 120 | 110mm air | f2_dekvloer 10` ← **verdubbeld**
- `constr-2` wall compass=S 25,58 m² — laag `o1_glas_helder 15mm`
- `constr-3` wall compass=W 0,42 m² — lagen `2× hout_vuren_generiek 160mm` ← **verdubbeld**

---

## 3. De 4 bugs — root-causes

> Regelnummers volgen de code-trace van 22-05-2026. De implementer moet de exacte regels verifiëren — de bestanden kunnen sindsdien wijzigen.

### Bug #1 — Open verbindingen worden nooit geëxporteerd
- SEGC-faces met een `<Room Separation>` Model Line als sub-boundary worden door `_get_faces_from_segc` (`raycast_scanner.py:295-313`) puur op geometrie als gewone `wall`-face geclassificeerd — **geen categorie-check**.
- Ze gaan naar `_scan_wall_face`; de raycast vindt niets (een separation line is geen fysiek element) → `stacks_by_height` blijft leeg → vroege `return` op `raycast_scanner.py:548-549`. Geen construction, geen open connection.
- `_detect_open_connections` (functie op `raycast_scanner.py:1378-1426`) zou hier de open verbinding maken, maar de aanroep staat **uitgecommentarieerd** op `raycast_scanner.py:543-546` — bewust uitgezet wegens dubbeltelling met `FindInserts()`.
- **Risico:** als een ray dóór de separation line de echte buitengevel raakt binnen `RAY_MAX_DIST_M = 3.0 m`, krijgt vlak 3/4 onterecht een construction die de gevel dubbeltelt.

### Bug #2 — Curtain walls verdwijnen / worden geen opening
- Een curtain wall heeft géén `CompoundStructure`. In `_hits_to_layer_stack` valt de dikte terug op `max(1, 0) = 1 mm` (`raycast_scanner.py:996-1004`), `_get_element_thickness_mm` (`:1044`) geeft `0`, `_get_hit_lambda` (`:1079`) geeft `None` → de stack wordt een triviale 1 mm-laag of leeg.
- De curtain-panelen (`OST_CurtainWallPanels`) gaan in `_cast_ray_at_height` (`:743-746`) naar `opening_hits`, maar `_collect_openings_from_boundary_walls` (`:1477`) detecteert openingen via `wall.FindInserts()` (`:1545`) — een standalone curtain-wall-boundary levert daar geen inserts → `openings: []`.
- In de builder `_build_constructions` (`thermal_json_builder.py:407-428`) collapsen faces met identieke fingerprint `(room_a, room_b, orientation, compass, layers)`; faces met lege/triviale stack of < `MIN_CONSTRUCTION_AREA_M2` (0,25 m², filter `thermal_json_builder.py:451`) vallen weg. Eindresultaat: het grootste deel van de vliesgevel verdwijnt.

### Bug #3 — Laag-verdubbeling dak + vloer
- Vloer-vlak [1] heeft 2 SEGC-subfaces die naar **hetzelfde** vloer-element wijzen. `_get_faces_from_segc` (`:295-313`) dedupliceert niet op onderliggend element → 2 floor-face-dicts → `_scan_horizontal_face` draait 2× door hetzelfde vloerpakket.
- **⚠️ Belangrijke correctie op de code-trace:** de trace concludeerde dat het *plafond* aan de verdubbeling ontkomt (face [0], subfaces=0, één face-dict). De geëxporteerde JSON bewijst echter dat **óók `constr-0` (plafond) én `constr-3` (wand) verdubbeld zijn** (`2× hout_vuren_generiek 160mm`). De subface-telling verklaart de verdubbeling dus **niet volledig**. De implementer moet aanvullend onderzoeken hoe `_hits_to_layer_stack` (`:961-1004`) één enkele raycast door één element afhandelt — vermoedelijk worden enter- én exit-referenties, of overlappende geometrie-hits, beide als laag toegevoegd. Dit is de eigenlijke root-cause en geldt voor álle constructies, niet alleen de vloer.

### Bug #4 — `gross_area_m2` is een raycast-schatting, niet de echte vlak-area
- `_scan_wall_face` schat de breedte: `face_width_m = face_area_m2 / face_height_m` (`:555`), daarna per zone `zone_area = face_width_m × zone_height` (`:558-559`).
- `_build_constructions` sommeert die zone-areas (`thermal_json_builder.py:428`). Door lege ray-hoogtes en `RAY_HEIGHT_STEP_M/2`-marges wijkt dit af van de echte SEGC-area (vandaar 25,58 m² i.p.v. de SEGC-som).

**Samenhang:** #1 en #3 zijn geïsoleerd. **#2 en #4 zijn gekoppeld** — zolang curtain walls een dummy-stack/raycast-schatting krijgen blijft de area onbetrouwbaar; samen oppakken.

---

## 4. Fix-ontwerp

### Fix #1 — Separation lines → `open_connections`
1. In `_get_faces_from_segc`: lees per sub-face de categorie van het grenselement (`subface.SpatialBoundaryElement` → host-element → `Category.Id`). Is dat `BuiltInCategory.OST_RoomSeparationLines` → markeer de face als type `open_connection` i.p.v. `wall`.
2. Emit per zo'n face een `OpenConnection`-record. **Schema** (`thermal-import.schema.json` → `OpenConnection`): required = `room_a`, `room_b`, `area_m2`.
   - `room_a` = de gescande ruimte.
   - `room_b` = de aangrenzende ruimte. Bepaal die via de bestaande adjacent-room-logica (`all_rooms` wordt al doorgegeven aan `scan_room_boundaries`) — geometrisch welke room aan de andere kant van de separation line ligt. Als de buur niet te bepalen valt: overweeg een pseudo-room of sla over met log-melding.
   - `area_m2` = de **echte SEGC-vlak-area** (zie Fix #4).
3. Heractiveer **niet** de oude `_detect_open_connections` empty-height-heuristiek (`:1378-1426`, `:543-546`) — die was bewust uit wegens dubbeltelling. De SEGC-categorie-detectie is de schone vervanger.
4. Zorg dat zulke faces **niet** alsnog als wand-construction worden gescand (geen dubbeltelling van de gevel erachter).

### Fix #2 — Curtain wall → `curtain_wall`-opening
1. **Detectie:** een boundary-wand waar `wall.CurtainGrid is not None`. (Let op: `28_CWA_glas15mm` is een gewone Wall mét glas-`CompoundStructure` → géén curtain wall, blijft wand-construction.)
2. **Emit een host-`Construction`** voor het curtain-wall-vlak: `orientation = "wall"`, `compass` uit de face-normal, `gross_area_m2` = echte SEGC-vlak-area (Fix #4), `layers = []` (de Construction-schema vereist `layers` niet — leeg is geldig; er is geen opaak pakket). Vul `revit_element_id` + `revit_type_name` van de curtain wall.
3. **Emit een `Opening`** die naar die construction verwijst. **Schema** (`Opening`): required = `id`, `construction_id`, `type`, `width_mm`, `height_mm`.
   - `type = "curtain_wall"` — al in de schema-enum aanwezig.
   - `construction_id` → de host-construction uit stap 2.
   - `width_mm` / `height_mm` → uit de bounding box van het SEGC-vlak (verticaal vlak: horizontale extent × verticale extent).
   - `u_value` → uit het curtain-wall-**type** in Revit (thermische/analytische U-parameter) indien aanwezig; anders een default glas-U (stem af met bestaande opening-defaults in `constants.py`, bv. ~1,6 W/m²K). Markeer default vs. Revit-waarde indien mogelijk (vgl. audit-item D6).
   - `revit_element_id` + `revit_type_name` van de curtain wall.
4. **Eén opening per curtain-wall-vlak** (de hele vliesgevel als één glasvlak). Dichte (spandrel) panelen apart classificeren is een latere verfijning — buiten scope.
5. Resultaat: faces 7–10 worden 4 host-constructies + 4 `curtain_wall`-openingen i.p.v. te verdwijnen.

### Fix #3 — Laag-verdubbeling
1. **Primair:** in `_get_faces_from_segc` SEGC-subfaces dedupliceren op `SpatialBoundaryElement` host-element-id — meerdere subfaces van hetzelfde element samenvoegen tot één face-dict (areas optellen) vóór het scannen. Lost vloer-vlak [1] op.
2. **Eigenlijke root-cause (zie ⚠️ in §3):** onderzoek `_hits_to_layer_stack` (`:961-1004`) — waarom levert één raycast door één element (plafond, géén subface-verdubbeling) tóch een dubbele laag op? Vermoedelijk worden enter+exit of overlappende hits beide als laag toegevoegd. Fix daar de hit→laag-conversie zodat één doorboord element één laag oplevert.
3. **Vangnet:** dedupliceer binnen één layer-stack op `(material, thickness_mm)` opeenvolgende identieke lagen — maar dit mag de échte fix (stap 2) niet maskeren (twee identieke fysieke lagen ná elkaar bestaan ook legitiem).

### Fix #4 — Echte SEGC-area
1. Draag de werkelijke SEGC-(sub)face-area (`face.Area` / `subface.GetSubface().Area`, ft²→m²) mee tot in de builder voor wand-constructies, i.p.v. de `face_width × zone_height`-schatting.
2. Gebruik die area voor `Construction.gross_area_m2`, `OpenConnection.area_m2` (Fix #1) en als basis voor `Opening.width_mm`/`height_mm` (Fix #2).

---

## 5. Schema-referentie (`open-heatloss-studio/schemas/v1/thermal-import.schema.json`)

| Definitie | Required | Velden |
|---|---|---|
| `Opening` | `id, construction_id, type, width_mm, height_mm` | + `sill_height_mm, u_value, revit_element_id, revit_type_name`. `type` enum = `window / door / curtain_wall` |
| `Construction` | `id, room_a, room_b, orientation, gross_area_m2` | + `compass, revit_element_id, revit_type_name, layers`. `orientation` enum = `wall / floor / ceiling / roof` |
| `OpenConnection` | `room_a, room_b, area_m2` | — |
| `ConstructionLayer` | `material, thickness_mm` | + `distance_from_interior_mm, type (solid/air_gap), lambda` |

---

## 6. Werkwijze voor de implementatie

1. **Volgorde:** Fix #3 + #1 eerst (geïsoleerd, snel te verifiëren), daarna #2 + #4 samen (gekoppeld). Beide raken `raycast_scanner.py` → sequentieel uitvoeren, niet parallel.
2. **Agent:** delegeren via `general-purpose` agent met IronPython 2.7-context (geen f-strings, geen type hints, feet-eenheden ×304,8 = mm, Transaction-wrapping). Specialized write-agents niet gebruiken (zie orchestrator-memory: backgrounden/falen).
3. **IronPython-conventies:** `~/.claude/rules/ironpython.md` + `extensions/bouwkunde.extension/CONVENTIONS.md`.
4. **Verificatie:** pyRevit reload (Alt+klik logo → Reload) → Warmteverlies Export opnieuw voor Room 17 → vergelijk met deze testcase. Verwacht na fix:
   - `open_connections` = 2 (vlakken 3 + 4)
   - `openings` = ~4 met `type: "curtain_wall"` (vlakken 7–10)
   - geen verdubbelde lagen in dak/vloer/wand
   - `gross_area_m2` ≈ de SEGC-vlak-areas uit §2.1
5. **Niet-doelen:** schuine-vlakken-ondersteuning (dak is ~8°, niet nodig) · spandrel/glas-splitsing binnen één curtain wall · de overige audit-items uit `TODO.md` (B1, B5–B8, D-reeks, U-reeks, T-reeks).

---

## 7. Bestanden

| Bestand | Rol |
|---|---|
| `extensions/bouwkunde.extension/lib/warmteverlies/raycast_scanner.py` | Hoofdbestand — alle 4 fixes |
| `extensions/bouwkunde.extension/lib/warmteverlies/thermal_json_builder.py` | Builder — #2 (openingen + host-constructie), #4 (area) |
| `extensions/bouwkunde.extension/lib/warmteverlies/constants.py` | Default U-waarden, categorie-constanten |
| `extensions/bouwkunde.extension/Bouwkunde.tab/Bouwbesluit.panel/RaycastExport.pushbutton/script.py` | Pushbutton — naar verwachting ongewijzigd |
| `open-heatloss-studio/schemas/v1/thermal-import.schema.json` | Contract (alleen lezen — niet wijzigen) |
