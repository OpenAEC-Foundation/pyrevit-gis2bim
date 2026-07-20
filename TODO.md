# 3BM pyRevit Project - TODO

*Laatste update: 20 juli 2026*

---

## CrossDim — linked models + kolommen (20 juli 2026, nieuw, Revit-test pending)

> CrossDim v1.3.0: herkent nu wanden, kolommen (OST_Columns + OST_StructuralColumns) én rooms uit host en gelinkte modellen. Wanden/kolommen uit alle geladen (niet-verborgen) links met transform naar host-coördinaten; maatlijn-referenties via `Reference.CreateLinkReference()`; room-detectie in links via inverse transform (met Z-fallback op room-level bij verticale link-offset). Kolommen via face-based ray-hit (geen LocationCurve): ray-plane snijpunt + containment-check via `face.Project()`, geometrie uit `GeometryInstance.GetInstanceGeometry()`.

- [ ] **Live-test in Revit** — plattegrond met gelinkt bouwkundig model: klik in room uit link → H+V maatlijnen naar link-wanden. Check ook mixed geval (host-rooms + link-wanden).
- [ ] **Kolommen-test** — room met vrijstaande kolom: maatlijn moet op kolom-face stoppen als die dichterbij is dan de wand; test zowel host- als link-kolommen, en ronde kolommen (worden bewust genegeerd: geen PlanarFace).
- [x] **Badkamer/toilet root cause GEVONDEN + gefixt (20 juli, v1.5.0, live gediagnosticeerd op model ALP 5006)** — kliks in badkamer/toilet deden niets omdat `IsPointInRoom` op klik-Z faalde: de room-bounding afwerkvloer (tegelvloer ~15mm) tilt de onderkant van het roomvolume boven level-hoogte, en PickPoint klikt op level-hoogte. Fix: `point_in_room()` hertest op klik-Z + 600mm en room-level + 600mm. Kliks zonder room worden nu geteld/gelogd en in de eindmelding getoond i.p.v. stil genegeerd. Wandzoek in toilet live geverifieerd via MCP-simulatie: 4 richtingen vinden gibo70/kalkzandsteen100/metalstud70, tegelwanden `42_inbo_WA_wandtegels_300x300_15` correct geskipt.
- [x] **Afwerkingsfilter versimpeld (20 juli, v1.5.1, user-besluit)** — gipsplaat/voorzetwanden moeten wél gemaatvoerd worden; filter is nu uitsluitend `'tegel' in typenaam`. Dikte-regel (<30mm), overige keywords (stuc/afwerk/finish) en de `3BM_Afwerking_*`-marker zijn vervallen.
- [ ] **Default maatlijntype 'verkoop' (v1.5.0)** — combobox selecteert nu eerst type met 'verkoop' in de naam (match in ALP-model: `inbo_maatvoering_2.5_verkoop`), fallback '1.8'. Even visueel bevestigen.
- [ ] **View-refresh na elke klik verifiëren** — `uidoc.RefreshActiveView()` na elke transaction; maatlijnen moeten direct zichtbaar zijn vóór de volgende klik. Echte hover-preview (kruislijnen die met de muis meebewegen vóór de klik) kan niet met `PickPoint` — zou DirectContext3D vergen; alleen oppakken als de refresh onvoldoende blijkt.
- [ ] **NewDimension met link-references verifiëren** — werkt in de regel vanaf Revit 2019+, maar kan per view/werkvlak weigeren; fouten staan op debug-level in de bm_logger-log.
- [ ] Bij traag gedrag in grote links: wanden/kolommen voorfilteren op bounding box van de room i.p.v. alles meenemen.

---

## BblToets / Bbl Toets (14 juli 2026 — nieuw, Revit-test pending)

> Nieuwe tool `Bouwkunde.tab/Bouwbesluit.panel/BblToets.pushbutton` (eerder Ventilatievoud, zelfde sessie uitgebreid): Bbl-toets voor de bouwaanvraag (nieuwbouw woonfunctie) met twee tabbladen. **Ventilatie** (§4.3.5): eis per ruimte + Ducoton 10 'ZR' roosterlengtes (10,2 dm³/s per m¹ bij 1 Pa, geverifieerd tegen Duco-datasheet). **Daglicht** (§4.3.7): equivalente daglichtopp. 10% vloeropp. min. 0,5 m², aanwezig via windows per ruimte (FromRoom/ToRoom, kozijnmaten cf. Kozijnstaat-conventie `kozijn_breedte`/`kozijn_hoogte`) × glasfactor × belemmeringsfactor, per rij handmatig te overrulen. Type-detectie op ruimtenaam met override, gecombineerde tekst-tabel via klikpunt op actieve view.

- [x] **Daglicht-koppeling geverifieerd (14 juli, live via Revit MCP op model 5001)** — glas-families hebben nu een Room Calculation Point; koppeling loopt via `GetSpatialElementCalculationPoint` + `doc.GetRoomAtPoint(punt, fase)` (get_ToRoom geeft None op geneste families, óók met RCP). Fasen-fix: zoeken in fasen mét rooms i.p.v. laatste fase. 12/12 panelen rond testruimte correct
- [x] **Kolomuitlijning via tabs (15 juli, live geverifieerd)** — tabel gebruikt nu `\t`-uitlijning op de Tab Size van het gekozen teksttype i.p.v. spatie-padding; celbreedtes geschat via Arial-em-tabel × teksthoogte × Width Factor, beide secties op één gedeeld kolomraster. Monospace-teksttype niet meer nodig. User bevestigt: werkt goed in Revit
- [x] **Ruimte-selectie (15 juli, live geverifieerd)** — vinkkolom in beide grids (uitgevinkt = grijs, buiten tabel/totalen/JSON); rooms die bij het starten in Revit geselecteerd zijn worden voorgeselecteerd
- [x] **UI-test in Revit (15 juli)** — form, tabs, vinkkolom en plaatsing bevestigd werkend; bij verlopen kolommen in toekomstige projecten: Tab Size van het teksttype vergroten of `_KOLOM_MARGE_MM` bijstellen
- [ ] **Glas-maten exact maken** — glasopp. komt nu uit bbox-fallback (23,35 m² voor 12 panelen, mogelijk overschat bij draaiende delen); `glas_breedte`/`glas_hoogte` params aan glas-family toevoegen
- [ ] **Rooms plaatsen in model 5001** — nu 1 room (fase Nieuw); oude family `31 dubbel_glas` (2x) heeft geen RCP/maat
- [ ] Ventilatievoud-kolom verifiëren in project met volumeberekening aan (Areas & Volumes)
- [ ] Evt. uitbreiden: spuiventilatie (§4.3.5), geluidwering, daglicht per verblijfsgebied i.p.v. per ruimte

---

## Regels op de kaart / DSO (4 juli 2026 — nieuw, live-test pending)

> Nieuwe tool `GIS2BIM.tab/Data.panel/RegelsOpDeKaart.pushbutton` + client `lib/gis2bim/api/dso.py` (Ozon Presenteren API v8). Smoke-test tegen productie-API gedaan met dummy-key (endpoint/TLS/headers/foutafhandeling OK). API-key is aangevraagd maar nog niet binnen.

- [ ] **Live-test zodra DSO API-key binnen is** — key invoeren via de tool-prompt (opslag in `%APPDATA%\GIS2BIM\config.json` onder `dso_api_key`; omgeving via `dso_environment`: `prod`/`pre`)
- [ ] **Responseveld-namen verifiëren tegen echte data** — parsing is defensief (`_first`/`_als_tekst`), maar `citeerTitel`/`type`/`aangeleverdDoorEen` en de `_embedded`-resourcenamen zijn nog niet tegen een echte response gecheckt
- [ ] **Artikeltekst ophalen** — nu alleen annotaties (kruimelpad/wId); volledige tekst via `get_documentstructuur(regeling, documentComponent=wId)` tonen of exporteren
- [ ] **Rapportkoppeling** — JSON-export (`%TEMP%\3bm_exchange\dso_regels_op_locatie.json`) doorzetten naar BM Reports ("geldende regels voor deze locatie" PDF)
- [ ] **Werkingsgebieden tekenen** — geometrieën van werkingsgebieden als filled regions (via Omgevingsdocumenten geometrie opvragen API), zelfde patroon als WFS-tool
- [ ] Icoon voor de pushbutton

---

## Warmteverlies JSON-export — conformiteit (3 juni 2026, Reddingspost Kijkduin validatie)

> Gevonden bij PM-verificatie van een echte export (`Reddingspost_kijkduin.heatloss.json`) tegen de Revit-grondwaarheid op model `2786_Bouwkundige_model`. Producent = Catalogus-/RaycastExport (`lib/warmteverlies/raycast_scanner.py`). De gebruiker moest de ontbrekende vliesgevels handmatig aanvullen vóór de berekening klopte.

- [ ] **Vliesgevel valt weg waar geen verwarmde ruimte direct achter de grens zit** — `raycast_scanner.py:617` skipt vertikale `GetBoundaryFaceInfo Count==0`-faces (bewust, spiegelt commit `39c6768`). Bij een dubbele schil / spouw / glasgevel die aan een geplaatste **"Buiten"**-ruimte grenst, zit de curtain wall niet als room-bounding sub-face op de verwarmde ruimte → **~87 van 165 m² (53%) vliesgevel ontbrak** in de export. Fix: voor vertikale Count==0-faces raycasten (ReferenceIntersector, zoals al voor horizontale faces gebeurt) i.p.v. skippen — minstens wanneer er een curtain wall / glasgevel binnen X m achter ligt. Testcase: model 2786, ruimtes Politiepost/Instructie/Ieeftuimte/piket; verwacht totaal vliesgevel ≈ 165 m².
- [ ] **uitkijkpost / niveau-01 dubbelhoge beglazing** — ~32 m² curtain wall op niveau 01 (`uitkijkpost` + `uitkijkpost-uit`) grenst aan `18. piket` + `70 buiten`; raycast miste 'm (piket kreeg 1,8 m² extWand, 0 vliesgevel). Verifiëren dat dubbelhoge/schuine beglazing op hogere niveaus meekomt.
- [ ] **`openverbinding / glas`-constructie exporteert met U=0 (lege `layers`)** — cat-7 ("22") elementen (Instructie 0,32 m² + Politiepost 1,00 m²) komen met `layers: []` → U=0 → 0 W verlies. Of een glas-U toekennen, of als opening behandelen. Klein (1,3 m²) maar stil onder-tellend.

---

## Vlakaanzicht — vliesgevel-support (17 juni 2026, WIP/ONGETEST)

> Tool `Aanzichten.panel/Vlakaanzicht.pushbutton`: section loodrecht op gepickt vlak. Vliesgevel-pick (curtain-panel) cropte op één paneel i.p.v. de gevel. Crop-fix gecommit maar ongetest (actieve doc was family). Oriëntatie-bug nog open.

- [ ] **Crop-fix testen** — open project met hellende vliesgevel → Reload → pick glaspaneel → check of hele gevel in crop zit. `_resolve_curtain_host` klimt via `.Host` naar de curtain wall; `_compute_section_box(use_face_extent=False)` cropt op host-bbox.
- [x] **Kijkrichting-bug (view vanaf verkeerde kant)** — GEFIXT 16 juli 2026. Root cause via debug-log + live test: `ViewSection.CreateSection` zet de kijker aan de **-BasisZ**-kant (resulterende `ViewDirection = -BasisZ`); script gaf `BasisZ = +normaal` → view keek altijd vanaf de achterkant, gespiegeld. Fix: `BasisZ = -normaal`, `BasisX = BasisY × BasisZ`, Z-diepte in `_compute_section_box` gespiegeld (materiaal nu op +Z, far clip = `max(zs)`). Live geverifieerd op dakvlak 8405484: `viewdir == normaal`, updir hellingopwaarts, rightdir niet gespiegeld.
- [ ] **Oriëntatie-bug (sloped section blijft recht) hertesten** — user meldde eerder: "tekent hem recht, niet onder de hoek van de slanted wall". Mogelijk zelfde root cause (frame-flip); na kijkrichting-fix opnieuw testen op hellende wand/vliesgevel. Zo niet: verdachte A (normaal uit paneel-face i.p.v. host-wand) vs B (Revit snapt tilt naar verticaal).
- [ ] **Open vragen user**: over-crop bij sterk hellende gevel (host-bbox is wereld-AABB) acceptabel of strak op gevel-omtrek? Geknikte/gebogen gevel = echte ontvouwing/uitslag — apart traject, nu nodig?
- [ ] Diagnostische `_log`-regels weer verwijderen/dempen zodra de bug gefixt is.

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

## Deurstaat (8 juni 2026 — gedeelde core met Kozijnstaat, getest in Revit: Create/Config/Aantallen OK, Maatvoeren gefixt)

> Nieuw `Deurstaat.panel` als deuren-variant (OST_Doors) op de gedeelde Kozijnstaat-core. Architectuur: `lib/kozijnstaat/config.py` profiel-aware (`load_config("kozijn"|"deur")`, eigen `user_config_deur.json`), `family_collector` collect-functies met `category`-param, deur-knoppen als dunne shims via `lib/kozijnstaat/shim.py` (`run("deur")`). GlasTag bewust niet gedupliceerd (raam-specifiek).

- [ ] **Verticale deur-maatvoering** — family `32_DO_binnenkozijn_woning` heeft GEEN named Sill/Head reference planes (`GetReferenceByName` → None). `detail_v_refs`/`main_v_refs` staan daarom leeg. Fix: in de deur-family head- + sill-planes een **naam + Is Reference** geven, dan namen in `DEUR_OVERRIDES` zetten.
- [ ] **Deur-defaults verifiëren per project** — `param_merk="deurmerk"` en `kozijn_tag_family="31_TAG_de_deurstaat_door"` zijn placeholders; echte 3BM-deur-tag-family + merk-parameter bevestigen.
- [ ] **Spellingsinconsistentie in family** — `sponing_links` (1 n) vs `sponning_rechts` (2 n) in `32_DO_binnenkozijn_woning`; bij family-cleanup gelijktrekken + `detail_h_refs` mee-updaten.
- [ ] Deur-Legend (POC) + Rename in Revit testen met deur-profiel (alleen Create/Config/Aantallen/Maatvoeren bevestigd).
- [ ] Eigen iconen voor `DeurstaatRename` / `DeurstaatLegend` (nu geen icoon, net als kozijn-equivalent).
- [ ] Latente bug opgemerkt: 4x-`dirname` sys.path in oude Kozijnstaat-knoppen wijst naar `extensions/` i.p.v. de extension-root; werkt alleen omdat pyRevit `lib/` auto-toevoegt. Bij refactor rechttrekken (deur-shims gebruiken al correcte 3x).

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
- [x] ~~**UI-polish + bugfixes (31-05)**~~ — openingen krijgen kozijn/deur-Type als type_stapel (was leeg), glaswanden→opening-classificatie, multi-punt wand-stapel-sampling, defaults 0.5/verwarmd-uit, air-gap-break conditioneel (volledige dak/wand-opbouw), `oppervlak_m2` 2 decimalen, vliesgevel achter room-separation-line→opening (commits `1a96610`, `7f8c083`, `c0e51f2`, `ddc1892`).
- [x] ~~**View-behoud-fix (31-05)**~~ — WV-Grensvlak-check 3D view wordt niet meer verwijderd/gerecreëerd bij Tonen, blijft op sheet met template staan.
- [x] ~~**Schedule "WV - Grensvlakken per ruimte" + WV-3D-view geplaatst op sheet `UO-901` (31-05)**~~ — volledig operationeel.

### Constructie-type-catalogus (Fase 1 afgerond, Fase 2 open)

**Doel:** model-brede tabel met genummerde named typen (`buitenwand1`, `dak1`…) + laag-opbouw-schedule (materiaal·dikte·λ) "selecteerbaar voor de berekening".

- [x] ~~**Fase 1 — constructie-type-catalogus (named typen + merge)**~~ — **AFGEROND (31-05)**: `assign_construction_catalog()` in `boundary_preview.py` implementeert deel-vangst-merge algoritme via orientatie+grenstype bucketing, reverse-merge, wig-allowlist, dominante ankers, fold van onder-vangsten met `vangst=onvolledig`-flag. 2 nieuwe shared params `warmteverlies_constructie` + `warmteverlies_vangst`. Live geverifieerd op model 2786: 7 construction-typen + 9 opening-typen, 185 faces benoemd.
- [x] ~~**Blokker #1 — vliesgevel achter room-separation-line → opening**~~ — opgelost in commit `ddc1892` (31-05).
- [x] ~~**Vliesgevel-detectie verbetering**~~ — **AFGEROND (31-05)**: `_curtain_glass_behind_sampled()` multi-sample herkent curtain panels/mullions/grids + CurtainGrid-walls via meerderheidsregel. Vliesgevel achter room-separation-line wordt correct als opening geclassificeerd (id 5679438, 29.5m²).
- [x] ~~**Fase 2 — laag-opbouw-schedule uit Revit**~~ — **AFGEROND (31-05)**: `CompoundStructure` van component-types geëxtraheerd via `_extract_layers_from_type()` + `_format_layers_string()`. Nieuwe shared param `warmteverlies_lagen` bevat per named-constructie de laagopbouw (materiaal + dikte_mm + functie, GEEN λ — rekentool levert λ zelf). Geverifieerd op model 2786: alle 7 constructies leveren correcte lagen.
- [x] ~~**Fase 3a — azimuth + buurruimte-id per grensvlak**~~ — **AFGEROND (31-05)**: `warmteverlies_azimut` (kompasrichting 0-360°, -1 voor horizontaal) + `warmteverlies_naar_id` (buurruimte element-id, 0 voor exterior/ground) als shared parameters. `_calculate_azimuth(normal)` (atan2-gebaseerd, 0=Noord) + `_resolve_adjacency` geeft buurruimte-id als 4e returnwaarde. Prerequisite voor ThermalImport JSON-export (compass + room_b).
- [x] ~~**Fase 3b — JSON-export builder**~~ — **AFGEROND (31-05)**: `catalog_export.py` → ThermalImport JSON met rooms + constructies-met-lagen + grensvlakken (grenstype/naar_ruimte/azimuth/netto-opp) + openingen (nog leeg). Tool levert λ + glas-U. `WarmteverliesExport.pushbutton` implementeert de export-dialog met save-file.
- [x] ~~**Fase 3c — openingen in JSON-export**~~ — **AFGEROND (31-05)**: `catalog_export.py` vult `openings` per WV_BND opening-shape (breedte/hoogte/sill uit bbox, type-detectie, koppeling aan exterieure host-constructie met room_b=outside + net>0 guard). Fallback-constructies voor volglas-rooms. Tool geeft nu volledig conforme ThermalImport JSON. Geverifieerd op model 2786: 31 openingen, alle hosts exterieur + netto>0.
	- [ ] **Minor follow-ups**: (a) 2 openingen vallen af op room_a-labelmismatch, (b) veel openingen op volglas-fallback i.p.v. echte buitenwand (compass-matching kan strakker), (c) oude exporter (raycast_scanner/uvalue_extractor) heeft nog λ-conversiebug `/6.93347` (nu omzeild want nieuwe export geeft geen λ).

**Architectuur-noot:** schedule komt uit `boundary_preview` (WV_BND-shapes), NIET uit `thermal_json_builder`. Exporter-#3 (JSON) maakt de schedule dus niet compacter — apart codepad.  
**Valkuil:** Revit ViewSchedule max 4 sort/group-velden.

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
- [ ] **D7 — vliesgevels 0-vertex fallback → 3D-viewer kan ze niet tekenen. FIX GEÏMPLEMENTEERD 16-06, ONGETEST.**
  - **Diagnose (model 2786):** con-147..151 op room-14/16/19/20, `revit_type_name` eindigt op " fallback"; gekoppelde curtain_wall-openings (open-23/24/27/28/29) óók 0 vertices. SEGC `GetBoundaryFaceInfo(face).Count==0` voor curtain-walls (óók als wand Room Bounding is, live geverifieerd: alle 12 zijn rb=1) → face viel in `not hosts`-tak → gele wand/sliver/orphan-fallback.
  - **PIJPLIJN-CORRECTIE:** live export = `WarmteverliesExport.pushbutton` → `catalog_export.build_catalog_thermal_import` → `boundary_preview.py` (WV_BND). **NIET** `raycast_scanner.py` (= `_archived/RaycastExport`, dood pad). Eerdere diagnose op raycast_scanner was misdirected.
  - **Fix (ongecommit→commit deze sessie):** `boundary_preview.py:~3161` (not-hosts-tak: `_curtain_glass_behind_sampled` → bij glas `orient="vliesgevel"` met echte `_face_to_vertices_json`-vertices) + `catalog_export.py:~836` (orphan-vertices behouden op opening).
  - **VERIFICATIE PENDING (volgende Revit-sessie):** pyRevit **Reload** → `WarmteverliesExport` op model 2786 (ruimtes 14/16/19/20) → check in JSON: con-147..151 hebben `vertices` + géén " fallback"-suffix; curtain_wall-openings `vertices` niet-leeg; `fallback_constructions`-teller omlaag; negatieve test: gewone ruimtes krijgen geen valse blauwe vlakken. Werkt → D7 afvinken. Werkt niet → agent bijsturen met echte export-output.
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
