# BGT 3D Tool - Status

## Laatst bijgewerkt: 2026-02-28

## Huidige status: Werkend met DirectShape fallback

### Wat is gedaan
1. **Hybrid Family Loading** geimplementeerd (`families.py`)
   - PlacementStrategy pattern: family / .rfa laden / DirectShape fallback
   - Auto-optie `[Automatisch laden / generiek]` in dropdowns
   - DirectShape boom: cilinder stam + cilinder kruin (met materialen)
   - DirectShape lamp: cilinder paal + box armatuur (met materialen)

2. **Nieuwe geometry functies** in `geometry.py`
   - `create_box_solid()` — bewezen extrusie-aanpak
   - `create_sphere_solid()` — TessellatedShapeBuilder (werkt NIET in alle Revit contexten)

3. **Bugfixes**
   - TLS 1.2 geforceerd in `ogc_api.py` voor PDOK compatibiliteit
   - Diagnostische logging toegevoegd (API URL, feature counts, RD validatie)
   - Deduplicatie in `_extract_points_from_features()` (PDOK retourneert historische versies)
   - Boomkruin gewijzigd van bol (TessellatedShapeBuilder → crash) naar cilinder

4. **UI update** (`UI.xaml`)
   - Tip-tekst over auto-optie toegevoegd

### Bekende issues
- `create_sphere_solid()` geeft "An internal error" in Revit — niet gebruikt voor bomen, staat nog in geometry.py
- Voorstraat 172 Dordrecht heeft weinig BGT puntdata (3 bomen in 400m, 0 lichtmasten)
- Geen .rfa bestanden meegeleverd in `families/` map (DirectShape fallback werkt)

### Gewijzigde bestanden
- `lib/gis2bim/revit/families.py` — NIEUW
- `lib/gis2bim/revit/geometry.py` — create_box_solid, create_sphere_solid
- `lib/gis2bim/revit/__init__.py` — families exports
- `lib/gis2bim/api/ogc_api.py` — TLS 1.2, logging
- `GIS2BIM.tab/Data.panel/BGT3D.pushbutton/script.py` — strategy pattern, dedup
- `GIS2BIM.tab/Data.panel/BGT3D.pushbutton/UI.xaml` — tip-tekst
- `GIS2BIM.tab/Data.panel/BGT3D.pushbutton/families/` — NIEUW (lege map)
