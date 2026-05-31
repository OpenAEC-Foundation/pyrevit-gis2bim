# 3BM pyRevit Project - TODO

*Laatste update: 31 mei 2026*

---

## TrapTekenen / Trap 2D (24 mei 2026 — wachten op hertest in Revit)

- [ ] Hertest in Revit met L-sparing na concave-fix (interior_in_dir/interior_out_dir uit polygon-CCW)
- [ ] Eigen icoon voor `TrapTekenen.pushbutton` (huidig: default pyRevit)
- [ ] Vervang komma-gescheiden parameter-prompt door WPF-form met live preview
- [ ] 2D-preview in pyRevit output window (HTML/SVG via `lib/trap/plot_svg.py`)
- [ ] Spilpaal-circle visualiseren (cirkel met radius = inner_radius)
- [ ] Boog-segmenten in sparing-curves correct ondersteunen (nu chord-approximatie)
- [ ] Uitbreiden naar bovenkwart, U-trap (halve slag), spiltrap
- [ ] Stap naar 3D: DirectShape (Generic Model) per trede met extrusion
- [ ] IFC export-pad (`IfcStair` met `IfcStairFlight` + winders)
- [ ] Native Revit Stairs API plaatsing (`StairsEditScope` + `CreateSketchedRun`)
- [ ] Looplijn-curve verbeteren: gebruik tangent-boog tussen rechte loops i.p.v. boog rond spil
- [ ] Echte Franse balanceermethode (méthode du balancement) i.p.v. simpele overlap-uitbreiding
- [ ] Treden-export naar JSON in `%TEMP%\3bm_exchange\` voor downstream tools

---

## Kozijnstaat (mei 2026 sessie — pyRevit reload + test pending)

- [ ] `KozijnstaatSync` pushbutton (optie B) — herbruikbare canvas→model param-sync zonder delete+recreate
- [ ] Muursparing-instances die per ongeluk in workset `kozijnstaat` belanden — onderzoeken: komen uit `KozijnstaatWizard`?
- [ ] `KozijnstaatGlasTag` / `KozijnstaatWindowTag` workset-aware maken (idem als Create) — annotaties moeten ook in workset `kozijnstaat`
- [ ] Bij Create-output een waarschuwing tonen wanneer 1 type meerdere `sparing_type`-varianten heeft in het model (first-wins-flag)
- [ ] Pyrevit Routes API stabiliteit — valt soms uit na heavy `execute_revit_code` calls; herstartrecept documenteren

---

## Warmteverlies — pre-export visuele check (pushbutton) — idee 2026-05-30

> Nieuw feature-idee (sessie 30-05, geprototyped via Revit MCP op `2786_Bouwkundige_model`). Doel: de gebruiker krijgt vóór JSON-export een **visuele controle** van wat de exporter "ziet" — gekleurde SEGC-grensvlakken per ruimte, direct in een 3D-view.

> **GEÏMPLEMENTEERD 30-05** (commit volgt): `WarmteverliesGrensvlakCheck.pushbutton` + `WarmteverliesGrensvlakWis.pushbutton` + `lib/warmteverlies/boundary_preview.py`. Kern live gevalideerd via Revit MCP op `2786_Bouwkundige_model` (21 ruimten, 0 failures, vliesgevel/deuren/ramen/vloer-counts matchen het prototype). **UI getest in Revit 31-05: dialog/help/Tonen/Wissen OK.** De 4 parameters zijn geïmplementeerd (min vlakgrootte, openingen tonen, host-loze slivers verbergen, alleen verwarmde ruimten).

- [x] ~~`WarmteverliesGrensvlakCheck.pushbutton`~~ naast de bestaande Export-knop (`Bouwbesluit.panel`) — rendert per ruimte de SEGC-grensvlakken als gekleurde DirectShapes (Generic Models, Comments-prefix `WV_BND`).
  - Kleur: **rood** = dak/plafond (n.Z > 0,7), **geel** = wand, **groen** = vloer (n.Z < −0,7), **blauw** = openingen (vliesgevel via eigen schuine SEGC-face + deuren/ramen als rechthoek).
  - **Host-loze wand-slivers verbergen** (`GetBoundaryFaceInfo Count == 0` + verticaal) — spiegelt de exporter-skip (commit `39c6768`); horizontale host-loze vlakken wél tonen (raycast vindt pakket).
  - [x] ~~**Openingen detecteren en apart kleuren**~~ — deuren/ramen via `Wall.FindInserts` (extent-rechthoek), vliesgevels via het SEGC-grensvlak met host = curtain wall (volgt de schuine kap, blijft binnen de ruimtecontour). Beide blauw.
  - [x] ~~**Clear-knop**~~ — aparte `WarmteverliesGrensvlakWis.pushbutton` verwijdert alle DirectShapes met Comments-prefix `WV_BND`.
  - [ ] **Pushbutton-UI testen in Revit** (forms-prompts: multiselect-opties + min-vlakgrootte-invoer) — alleen de lib-kern is via MCP getest.
  - [ ] **Iconen** voor de twee pushbuttons (nu pyRevit default-icoon).
- [x] ~~**Shared parameters per grensvlak**~~ (commits `d78a3fe` + `85a2ec3`, 30-05 deel 3) — elk WV_BND-vlak draagt 7 instance-params (groep "Berekeningen", prefix `warmteverlies_`): `ruimte`, `naar_ruimte`, `grenstype`, `orientatie`, `oppervlak_m2`, `host_type`, `type_stapel`. Plus type-param `warmteverlies_afwerklaag` (Yes/No, gebootstrapt uit Type Comments "afwerk").
  - Adjacency **geometrisch** via `GetRoomAtPoint(punt, room-phase)` + naar-buiten-normaal-probe (phase verplicht!). Type-stapel via hergebruik scanner-raycast, afgekapt op min(buurruimte-afstand, eerste luchtspleet), **element-bewust** (voor/achtervlak van 1 element ≠ spleet). Afwerklaag-Types uit de stapel gefilterd. Resultaat op 2786: 147 → **19 distinct per-vlak type-stapels**, isolatie behouden.
- [x] ~~**UI-polish + bugfixes (31-05)**~~ — openingen krijgen kozijn/deur-Type als type_stapel (was leeg), glaswanden→opening-classificatie, multi-punt wand-stapel-sampling, defaults 0.5/verwarmd-uit.
- [x] ~~**View-behoud-fix (31-05)**~~ — WV-Grensvlak-check 3D view wordt niet meer verwijderd/gerecreëerd bij Tonen, blijft op sheet met template staan.
- [ ] **Consolidatie 19 → ~5 echte constructies** (volgende stap, hoort in `thermal_json_builder.py` = exporter #3): de per-vlak stapels samenvouwen via een **canonieke fingerprint** — lagen sorteren in vaste richting (binnen→buiten) zodat `060-TL>PIR` en `PIR>060-TL` één worden — plus merge van deel-vangsten (subset/superset met identieke kern). Doel-telling (user, model 2786): 2 daken · 1 buitenwand · 2 binnenwanden · openingen apart. Optioneel tussenstapje: volgorde-canonicalisatie al in `boundary_preview.py::_type_stack_for_face`.
- **Ontwerpregel (KRITISCH):** preview en export MOETEN dezelfde face-extractie/-filter/-groepering delen (`_get_faces_from_segc` + #3-host-groepering). Aparte implementatie = divergentie = onbetrouwbare check.
- **Volgorde:** eerst #3 (host-element-groepering) + #5 (vliesgevel → `curtain_wall`) perfectioneren, dán de preview bovenop die gedeelde code — dan toont de preview meteen de gegroepeerde, schone constructies.
- **Blauwdruk:** de MCP-prototype-code uit sessie 30-05 (oriëntatie-classificatie + `TessellatedShapeBuilder` per face + 3 materials `WV_BND_TOP/WALL/BOT` + host-loos-skip). Zie `docs/2026-05-29-warmteverlies-exporter-bevindingen.md`.
- **Modelvereiste die deze sessie bevestigd is:** afwerk-/afschotvloeren **niet-room-bounding**, dragende constructievloer wél → halveert de fragmentatie aan de bron (vloer overal 1 vlak). Hoort als modelvereiste bij bevinding #6.
- Relatie: vervangt/verrijkt audit-item **U1** (was: tekst-dialoog preview) met een echte 3D-visuele check; levert ook visuele validatie voor **D2** (boundary_polygon).

---

## Warmteverlies-exporter — code-audit 2026-05-22

> Read-only audit van `raycast_scanner.py` + `thermal_json_builder.py` + `RaycastExport.pushbutton` (consument-keten: open-heatloss-studio thermal-import). 25 bevindingen; U4 deze sessie gefixt. Schema-afhankelijke items (D3/D4) staan in de open-heatloss-studio `TODO.md`.

### Bugs (correctheid)
- [x] ~~U4 — samenvatting telt op niet-bestaande sleutel `room_type` → "Export geslaagd" toont altijd "0 ruimten"~~ → `RaycastExport.pushbutton/script.py:376` `room_type`→`type` (2026-05-22)
- [ ] B1 — `revit_type_name`/`revit_element_id` ontbreken op construction-dicts (`raycast_scanner.py:561-569,675-683`) → alle catalog-entries krijgen `revit_type_name=None`
- [ ] B2 — geen phase-filter op rooms (`room_collector.py:37-61`) → gesloopte/bestaande-toestand ruimten lekken in de export
- [ ] B3 — wand-`area_m2` is rechthoek-schatting width×height (`raycast_scanner.py:555-559`) → fout bij schuine/L-vormige faces
- [ ] B4 — Room Separation Lines niet gefilterd op SEGC-faces (`raycast_scanner.py:295-313`) → separation-line-grens wordt als volwaardige wand gescand
- [ ] B5 — laagdikte/spouw-gap fout bij `enter==exit` raycast-hit (`raycast_scanner.py:996-1032`)
- [ ] B6 — link-detectie inconsistent (Title-vergelijk vs LinkedElementId) (`raycast_scanner.py:1786-1791`)
- [ ] B7 — fragiele sentinel 1000/1500 in opening-afmetingen → `None`-sentinels gebruiken (`raycast_scanner.py:2353-2391`)
- [ ] B8 — `exported_at` zonder timezone, schema verwacht RFC3339 (`thermal_json_builder.py:107`)

### Datakwaliteit
- [ ] D1 — `sill_height_mm` (kozijnhoogte) nooit geëxporteerd, z-range is al berekend (`raycast_scanner.py:1660-1671,1721-1732`)
- [ ] D2 — geen room `boundary_polygon` → 3D-viewer in open-heatloss-studio blijft leeg; schema ondersteunt het al (`thermal_json_builder.py:176-191`)
- [ ] D5 — multi-layer host-wand verliest laagopbouw, exporteert één materiaalnaam (`raycast_scanner.py:893-928`)
- [ ] D6 — opening default-U (1.60/1.70) niet te onderscheiden van Revit-waarde → `u_value_source` toevoegen (`raycast_scanner.py:2437-2439`)
- [ ] Verify — rooms in export-JSON hebben `function`/`floor_area_m2` = `None`; bepalen of dat hoort (worden in de import-wizard gezet) of een datagat is

### UX pushbutton
- [ ] U1 — geen preview/validatie vóór opslaan; samenvatting-dialoog + lokale schema-check toevoegen (`script.py:357-441`)
- [ ] U2 — afgeleide functie + heated-status niet zichtbaar in ruimteselectie → misclassificatie onzichtbaar (`script.py:35-44`)
- [ ] U3 — silent `return` bij scan-exception, reden niet gelogd (`raycast_scanner.py:69-72`)

### Tech-debt
- [ ] T1 — dode pushbuttons `_ThermalExport.pushbutton` + `_WarmteverliesExport.pushbutton` verwijderen
- [ ] T2 — verifiëren welke lib-modules het raycast-pad nog gebruikt; `boundary_analyzer/adjacent_detector/wall_assembly_resolver/uvalue_extractor/opening_extractor` mogelijk legacy
- [ ] T3 — dode `_collect_openings_from_hits` verwijderen (`raycast_scanner.py:1967-2082`, ~110 regels legacy)
- [ ] T4 — DEBUG_OPENINGS-prints scheiden van kernlogica (`raycast_scanner.py:1510-1732`)
- [ ] T5 — bare `except:` → `except Exception:` (`raycast_scanner.py:1531,1641,1698`)
- [ ] T6 — `json_builder.py` — bepalen of nog een actieve pushbutton dit gebruikt, anders verwijderen
- [ ] T7 — dubbele fingerprint-implementatie consolideren (`raycast_scanner.py:605-626` + `thermal_json_builder.py:241-266`)

---

## Hoge Prioriteit

### AHN Texture validatie
- [ ] Scale factor `* 100` (cm) valideren — 100m moet 100.000mm tonen in Revit Material Editor
- [ ] Texture positionering controleren (offset 0,0 correct?)
- [ ] Texture alleen zichtbaar in Realistic/Raytraced visual style

### WPF Migratie
Alle bestaande tools gebruiken Windows Forms. Nieuwe tools worden in WPF gebouwd (zie `MCPStatus` als referentie).

- [x] ~~SheetParameters migreren naar WPF~~ → voltooid
- [x] ~~AutoDim migreren naar WPF~~ → voltooid
- [x] ~~HellingbaanGenerator migreren naar WPF~~ → voltooid
- [ ] RcBerekening migreren naar WPF (complex: custom paint panels, diagrammen)

### GIS2BIM - Ontbrekende Tools
- [ ] KaartTijdreis tool bouwen (historische kaarten tijdreeks)
- [ ] OSM data import tool (OpenStreetMap gebouwen/wegen naar Revit)

### 3D Mesh Import (Mesh3D) - Testen
- [ ] Testen in Revit 2025 met OBJ bestand
- [ ] Testen met GLB bestand (ECEF coordinaten)
- [ ] EEA-waarschuwing toevoegen in Google 3D UI panel

---

## Normaal

### GIS2BIM Verbetering
- [x] ~~Natura2000 gebieden tool~~ → voltooid
- [ ] Gedeelde `_setup_styles()` extraheren (lijnstijl/filled region dropdowns in WFS/BGT)
- [ ] Alle 7 GIS2BIM tools testen na refactoring met gedeelde modules

### 3BM_Bouwkunde Verbetering
- [ ] Rc-tool uitbreiden: dynamische vochtbalans (Glaser → tijdsafhankelijk)
- [ ] AutoDim: reference detection verbeteren bij complexe wanden
- [ ] SheetParameters: `V Peil Zichtbaar` + `Kenmerknummer` — params bestaan niet in titleblock-family `A4_A0_grootformaat`. WIP: family aanpassen óf UI-velden verwijderen
- [x] ~~SheetParameters: tekst-afsnijding op HD schermen oplossen~~ → opgelost door WPF migratie

### Documentatie
- [ ] ARCHITECTURE.md bijwerken (GIS2BIM structuur toevoegen)
- [ ] CONVENTIONS.md bijwerken (GIS2BIM conventies)

---

## Housekeeping (uit Lessons Learned audit 2026-02-24)

- [ ] `lessons_learned.md` aanmaken op basis van template (zie `../lessons_learned_template.md`)
- [ ] Vastleggen: IronPython 2.7 beperkingen (geen f-strings, geen type hints, geen moderne syntax) — nieuwe ontwikkelaars struikelen hier altijd over
- [ ] Vastleggen: WPF migratie-ervaring documenteren (wat werkt, wat niet, tijdschatting per tool-complexiteit)
- [ ] Vastleggen: PDOK API's hebben timeout + retry nodig — standaard wrapper bouwen in `lib/`
- [ ] Vastleggen: Thermische geleidbaarheid Revit → SI conversiefactor (6.93347) ergens centraal documenteren
- [ ] Vastleggen: DPI scaling problemen op HD schermen — WPF lost dit automatisch op, WinForms vereist DPIScaler
- [ ] Overweeg: gedeelde `_setup_styles()` extraheren als lib-module i.p.v. per-tool duplicatie

---

## Laag Prioriteit / Nice-to-have

### WPF Migratie (overige tools)
- [ ] VentilatieBalans → WPF
- [ ] WandVloerAfwerking → WPF
- [ ] FilterCreator → WPF
- [ ] CrossDim → WPF
- [ ] NAAKTGenerator → WPF
- [ ] PalenNummeren → WPF
- [ ] ScheduleExport/Import → WPF

### Overig
- [ ] Test.panel opruimen (MCPStatus verplaatsen of verwijderen)
- [ ] Materialen database uitbreiden / actualiseren
- [ ] Installer script testen en updaten

---

## Voltooid

### Mei 2026
- [x] **Kozijnstaat tooling end-to-end werkend in Revit 2025**:
  - Create: per-kozijn variabele wand-fill layout (eigen breedte + 500 mm h, 2000 mm v tussen rijen), start linksboven, view-aligned u-direction via `wall.Orientation`
  - Maatvoeren: 4 dim-lines per kozijn (detail/totaal × H/V) met view-aware placement op 150/250 mm offset, gebruikt werkelijke kozijn-afmetingen i.p.v. bbox
  - GlasTag: anchor op bottom-left hoek glas-bbox + view-aligned h/v offsets (50/500 mm default)
  - WindowTag: nieuwe pushbutton, tagt kozijnen met `31_TAG_wi_kozijnstaat_window` op 500 mm onder sill
  - Aantallen: filter tag-families uit (TAG in naam), defensieve `.Name` via .NET reflection, param `getekend` + `aantal_gespiegeld`
  - Wizard: 5 stappen (Create → Maatvoeren → GlasTag → WindowTag → Aantallen)
- [x] File-logger module (`lib/kozijnstaat/logger.py`) voor in-Revit debug output naar `%TEMP%\3bm_exchange\kozijnstaat_debug.log`
- [x] `_safe_name()` helper in family_collector + scripts: omzeilt IronPython 2.7 / Revit 2025 `.Name` AttributeError via `GetType().GetProperty("Name").GetValue()`
- [x] Parameter readout met sanity-check: detecteert Length-vs-Number-storage, valt terug op raw mm bij waardes buiten 100..6000 mm range
- [x] Icon voor WindowTag-pushbutton (kozijn + leader + tag in 3BM huisstijl)

### Maart 2026
- [x] GIS2BIM: Mesh3D tool (OBJ/GLB import, Google 3D Tiles API, MTL kleuren, ECEF conversie)
- [x] Nieuwe parsers: GLB (binary glTF 2.0), MTL (Wavefront materialen)
- [x] OBJ parser uitgebreid met mtllib/usemtl materiaal-tracking
- [x] Google 3D Tiles client (tileset traversal, bounding volume filtering)
- [x] ECEF ↔ WGS84 coordinaat conversie
- [x] Icon in 3BM huisstijl gegenereerd

### Februari 2026
- [x] WPF migratie: SheetParameters, AutoDim, HellingbaanGenerator (WinForms → WPF + XAML)
- [x] GIS2BIM: Natura2000 tool (WFS query, afstandsberekening, filled regions, parameters)
- [x] DetailOverzicht tool (detailbibliotheek overzicht)
- [x] GIS2BIM: LuchtfotoTijdreis tool (PDOK luchtfoto's op sheet, 3x2 grid)
- [x] GIS2BIM: Grote refactoring gedeelde modules (7 tools bijgewerkt)
- [x] GIS2BIM: NAPPeilmerken tool
- [x] GIS2BIM Icons Stijl A (alle tool-iconen)
- [x] Projectmap opgeruimd (logs, prototypes, verouderde docs)

### Januari 2026
- [x] GIS2BIM: BAG3D tool (3D gebouwen OBJ mesh → DirectShape)
- [x] GIS2BIM: AHN tool (hoogte data WCS/LAZ → TopographySurface)
- [x] GIS2BIM: BGT tool (19 lagen, holes/donuts, boundary lines)
- [x] GIS2BIM: WFS tool (kadaster, BAG, gebouwen)
- [x] GIS2BIM: Locatie tool (PDOK geocoding)
- [x] FilterCreator tool
- [x] IFCKozijnAnalyzer tool
- [x] MCPStatus WPF referentie-implementatie
- [x] WPF template (`lib/wpf_template.py`)
- [x] HellingbaanGenerator (NEN 2443)
- [x] NAA.K.T. Generator
- [x] 3BM Bouwkunde Icons v3

### December 2025
- [x] AutoDim tool
- [x] RcBerekening met Glaser condensatie
- [x] 4K DPI scaling opgelost
- [x] UI template framework (BaseForm, UIFactory, DPIScaler)
- [x] Centrale logging (bm_logger.py)
