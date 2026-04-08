# 3BM pyRevit Project - TODO

*Laatste update: 7 maart 2026*

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
