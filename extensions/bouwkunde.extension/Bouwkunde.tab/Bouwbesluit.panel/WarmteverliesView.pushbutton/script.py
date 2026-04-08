# -*- coding: utf-8 -*-
"""Warmteverlies 3D View — diagnostische visualisatie.

Maakt een 3D isometrische view aan met kleur-overrides per boundary type.
Hergebruikt dezelfde analyse-pipeline als WarmteverliesExport.

IronPython 2.7 — geen f-strings, geen type hints.
"""

__title__ = "Warmteverlies\nView"
__author__ = "3BM Bouwkunde"
__doc__ = "Maak een 3D controleview met kleurcodes per grenstype"

import os
import sys

from pyrevit import revit, DB, forms, script

from Autodesk.Revit.DB import (
    View3D,
    ViewFamilyType,
    ViewFamily,
    ViewDetailLevel,
    Transaction,
    ElementId,
    OverrideGraphicSettings,
    FillPatternElement,
    Color,
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    Wall,
)

# Lib pad toevoegen
sys.path.append(os.path.join(
    os.path.dirname(__file__), "..", "..", "..", "lib"
))

from warmteverlies.room_collector import collect_rooms
from warmteverlies.room_function_mapper import map_all_rooms
from warmteverlies.boundary_analyzer import analyze_boundaries
from warmteverlies.adjacent_detector import (
    classify_boundaries,
    build_element_room_lookup,
)
from warmteverlies.opening_extractor import extract_openings
from warmteverlies.wall_assembly_resolver import (
    collect_all_walls,
    resolve_wall_assembly,
)

# =============================================================================
# Constanten
# =============================================================================
VIEW_NAME = "WV - Controle"

# Kleuren per boundary type — RGB tuples
COLOR_EXTERIOR_WALL = (220, 50, 50)       # Rood
COLOR_EXTERIOR_ROOF = (150, 50, 200)      # Paars
COLOR_GROUND = (140, 100, 50)             # Bruin
COLOR_ADJACENT_HEATED = (50, 100, 200)    # Blauw
COLOR_UNHEATED_SPACE = (230, 150, 30)     # Oranje
COLOR_WINDOW = (0, 180, 200)              # Cyaan
COLOR_DOOR = (200, 180, 50)              # Goud
COLOR_ASSEMBLY_NON_BOUNDARY = (255, 150, 150)  # Roze

# Prioriteiten voor deduplicatie — hogere waarde wint
BOUNDARY_TYPE_PRIORITY = {
    "opening": 50,
    "exterior": 40,
    "unheated_space": 30,
    "ground": 20,
    "adjacent_room": 10,
    "assembly": 5,
}

# Categorieën die zichtbaar moeten blijven in de diagnostische view
VISIBLE_CATEGORIES = (
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_CurtainWallMullions,
)

# Labels voor de legenda
LEGEND_ITEMS = [
    ("Buitenwand", COLOR_EXTERIOR_WALL),
    ("Dak / buitenplafond", COLOR_EXTERIOR_ROOF),
    ("Grondvloer", COLOR_GROUND),
    ("Aangrenzend verwarmd", COLOR_ADJACENT_HEATED),
    ("Onverwarmde ruimte", COLOR_UNHEATED_SPACE),
    ("Ramen / curtain walls", COLOR_WINDOW),
    ("Deuren", COLOR_DOOR),
    ("Assembly (niet-grens)", COLOR_ASSEMBLY_NON_BOUNDARY),
]


# =============================================================================
# Helperfuncties — View management
# =============================================================================
def _delete_existing_view(doc, view_name):
    """Verwijder bestaande view met deze naam (geen templates)."""
    collector = (
        FilteredElementCollector(doc)
        .OfClass(View3D)
        .WhereElementIsNotElementType()
    )
    for view in collector:
        if view.IsTemplate:
            continue
        if view.Name == view_name:
            doc.Delete(view.Id)
            return


def _create_3d_view(doc, view_name):
    """Maak een nieuwe isometrische 3D view aan.

    Args:
        doc: Revit Document
        view_name: Naam voor de view

    Returns:
        View3D: De aangemaakte view
    """
    # Zoek een 3D ViewFamilyType
    vft_collector = (
        FilteredElementCollector(doc)
        .OfClass(ViewFamilyType)
    )
    view_family_type_id = None
    for vft in vft_collector:
        if vft.ViewFamily == ViewFamily.ThreeDimensional:
            view_family_type_id = vft.Id
            break

    if view_family_type_id is None:
        raise Exception("Geen 3D ViewFamilyType gevonden in het model.")

    view_3d = View3D.CreateIsometric(doc, view_family_type_id)
    view_3d.Name = view_name
    view_3d.DetailLevel = ViewDetailLevel.Fine
    return view_3d


def _get_solid_fill_pattern_id(doc):
    """Haal het solid fill pattern op voor surface overrides.

    Returns:
        ElementId of None: ID van het solid fill pattern
    """
    collector = FilteredElementCollector(doc).OfClass(FillPatternElement)
    for fpe in collector:
        fp = fpe.GetFillPattern()
        if fp and fp.IsSolidFill:
            return fpe.Id
    return None


def _hide_irrelevant_categories(doc, view):
    """Verberg alle categorieën die niet relevant zijn voor warmteverlies.

    Itereert alle categorieën in het document en verbergt alles
    dat niet in VISIBLE_CATEGORIES staat.
    """
    visible_ids = set()
    for bic in VISIBLE_CATEGORIES:
        visible_ids.add(ElementId(bic).IntegerValue)

    categories = doc.Settings.Categories
    for cat in categories:
        try:
            if cat.Id.IntegerValue not in visible_ids:
                if view.CanCategoryBeHidden(cat.Id):
                    view.SetCategoryHidden(cat.Id, True)
        except Exception:
            continue


# =============================================================================
# Helperfuncties — Curtain wall detectie
# =============================================================================
def _is_curtain_wall(doc, wall):
    """Controleer of een wand een curtain wall is."""
    try:
        wall_type = doc.GetElement(wall.GetTypeId())
        if wall_type and wall_type.Kind.ToString() == "Curtain":
            return True
    except Exception:
        pass
    return False


def _get_curtain_wall_sub_ids(wall):
    """Haal alle panel en mullion element IDs van een curtain wall.

    Returns:
        list[int]: Element IDs van panels en mullions
    """
    ids = []
    try:
        grid = wall.CurtainGrid
        if grid:
            for pid in grid.GetPanelIds():
                ids.append(pid.IntegerValue)
            for mid in grid.GetMullionIds():
                ids.append(mid.IntegerValue)
    except Exception:
        pass
    return ids


# =============================================================================
# Helperfuncties — Supplementaire element scan
# =============================================================================
FEET_TO_M = 0.3048
GROUND_LEVEL_THRESHOLD_M = 0.5
ROOF_CEILING_TOLERANCE_M = 0.1


def _is_exterior_wall_type(wall):
    """Controleer of een wand een buitenwand is via WallType.Function.

    WallFunction.Exterior = 0 in de Revit API.
    """
    try:
        wall_type = wall.WallType
        if wall_type is None:
            return False
        func_param = wall_type.get_Parameter(
            BuiltInParameter.FUNCTION_PARAM
        )
        if func_param and func_param.HasValue:
            return func_param.AsInteger() == 0
    except Exception:
        pass
    return False


def _get_level_elevation_m(doc, element):
    """Haal de level elevation op van een element in meters."""
    try:
        level_id = element.LevelId
        if level_id and level_id.IntegerValue > 0:
            level = doc.GetElement(level_id)
            if level:
                elev_param = level.get_Parameter(
                    BuiltInParameter.LEVEL_ELEV
                )
                if elev_param and elev_param.HasValue:
                    return elev_param.AsDouble() * FEET_TO_M
    except Exception:
        pass
    return None


def _compute_room_level_thresholds(rooms):
    """Bepaal laagste en hoogste level elevatie uit room data.

    Gebruikt alleen levels waarop daadwerkelijk rooms staan,
    niet alle levels in het model. Dit is betrouwbaarder voor
    het onderscheiden van grondvloer/tussenvloer/dakniveau.

    Args:
        rooms: list[dict] — room data dicts met level_elevation_m

    Returns:
        tuple: (lowest_m, highest_m) — room-based level drempels
    """
    elevations = []
    for r in rooms:
        elev = r.get("level_elevation_m")
        if elev is not None:
            elevations.append(elev)

    if not elevations:
        return (0.0, 3.0)

    return (min(elevations), max(elevations))


def _color_wall_openings(doc, wall, element_color_map):
    """Kleur ramen en deuren in een wand.

    Gedeelde helper om code-duplicatie te voorkomen.

    Args:
        doc: Revit Document
        wall: Wall element
        element_color_map: dict om bij te werken
    """
    wall_openings = extract_openings(doc, wall)
    for opn in wall_openings:
        opn_id = opn.get("element_id")
        opn_cat = opn.get("category")
        if opn_id is None:
            continue
        if opn_cat == "window":
            _register_element_color(
                element_color_map, opn_id,
                "opening", COLOR_WINDOW, "window",
            )
        elif opn_cat == "door":
            _register_element_color(
                element_color_map, opn_id,
                "opening", COLOR_DOOR, "door",
            )


def _supplement_uncolored_elements(
    doc, element_color_map, rooms, boundary_wall_ids
):
    """Kleur elementen die niet door de boundary-analyse gevonden zijn.

    Gebruikt room-adjacency data om wanden correct te classificeren:
    - Wand grenst aan 0 rooms -> exterior (rood)
    - Wand grenst aan 1 room + WallType.Function Exterior -> exterior
    - Wand grenst aan 1 room + niet Exterior -> onverwarmd (oranje)
    - Wand grenst aan 2+ rooms -> adjacent (blauw)
    - Curtain wall -> window (cyaan)

    Horizontale elementen gebruiken room-level drempels:
    - Floor op/onder laagste room level -> grond (bruin)
    - Floor boven hoogste room level -> dak (paars)
    - Floor daartussen -> tussenvloer (blauw)
    - Ceiling op/boven hoogste room level -> dak (paars)
    - Ceiling lager -> tussenplafond (blauw)
    - Roof -> altijd dak (paars)

    Args:
        doc: Revit Document
        element_color_map: dict {element_id: (priority, color_rgb, category)}
        rooms: list[dict] — alle room data
        boundary_wall_ids: set[int] — wall IDs uit boundary-analyse
    """
    colored_ids = set(element_color_map.keys())
    supplement_count = 0

    # Room-adjacency lookup voor wanden
    wall_room_lookup = build_element_room_lookup(doc, rooms)

    # Room-based level drempels
    lowest_m, highest_m = _compute_room_level_thresholds(rooms)

    # --- Alle wanden ---
    wall_collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
    )
    for wall in wall_collector:
        wid = wall.Id.IntegerValue
        if wid in colored_ids:
            continue

        if not isinstance(wall, Wall):
            continue

        # Curtain wall -> raam
        if _is_curtain_wall(doc, wall):
            _register_element_color(
                element_color_map, wid,
                "opening", COLOR_WINDOW, "window",
            )
            for sub_id in _get_curtain_wall_sub_ids(wall):
                _register_element_color(
                    element_color_map, sub_id,
                    "opening", COLOR_WINDOW, "window",
                )
            supplement_count += 1
            continue

        # Room-adjacency classificatie
        adjacent_rooms = wall_room_lookup.get(wid)
        room_count = len(adjacent_rooms) if adjacent_rooms else 0

        if room_count == 0:
            # Geen rooms -> exterior
            _register_element_color(
                element_color_map, wid,
                "exterior", COLOR_EXTERIOR_WALL, "exterior",
            )
            _color_wall_openings(doc, wall, element_color_map)
        elif room_count == 1:
            # 1 room: WallType.Function als tiebreaker
            if _is_exterior_wall_type(wall):
                _register_element_color(
                    element_color_map, wid,
                    "exterior", COLOR_EXTERIOR_WALL, "exterior",
                )
                _color_wall_openings(doc, wall, element_color_map)
            else:
                # Conservatief: onverwarmd (oranje)
                _register_element_color(
                    element_color_map, wid,
                    "unheated_space", COLOR_UNHEATED_SPACE,
                    "unheated_space",
                )
        else:
            # 2+ rooms -> adjacent (blauw)
            _register_element_color(
                element_color_map, wid,
                "adjacent_room", COLOR_ADJACENT_HEATED,
                "adjacent_room",
            )
        supplement_count += 1

    # --- Alle vloeren ---
    floor_collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Floors)
        .WhereElementIsNotElementType()
    )
    for floor in floor_collector:
        fid = floor.Id.IntegerValue
        if fid in colored_ids:
            continue

        elev = _get_level_elevation_m(doc, floor)

        if elev is not None and elev <= lowest_m + GROUND_LEVEL_THRESHOLD_M:
            # Op of onder laagste room level -> grondvloer (bruin)
            _register_element_color(
                element_color_map, fid,
                "ground", COLOR_GROUND, "ground",
            )
        elif elev is not None and elev > highest_m:
            # Boven hoogste room level -> daklaag (paars)
            _register_element_color(
                element_color_map, fid,
                "exterior", COLOR_EXTERIOR_ROOF, "exterior",
            )
        else:
            # Tussenvloer -> aangrenzend (blauw)
            _register_element_color(
                element_color_map, fid,
                "adjacent_room", COLOR_ADJACENT_HEATED,
                "adjacent_room",
            )
        supplement_count += 1

    # --- Alle daken ---
    roof_collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Roofs)
        .WhereElementIsNotElementType()
    )
    for roof in roof_collector:
        rid = roof.Id.IntegerValue
        if rid in colored_ids:
            continue

        _register_element_color(
            element_color_map, rid,
            "exterior", COLOR_EXTERIOR_ROOF, "exterior",
        )
        supplement_count += 1

    # --- Alle plafonds ---
    ceiling_collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Ceilings)
        .WhereElementIsNotElementType()
    )
    for ceiling in ceiling_collector:
        cid = ceiling.Id.IntegerValue
        if cid in colored_ids:
            continue

        elev = _get_level_elevation_m(doc, ceiling)

        if (elev is not None
                and elev >= highest_m - ROOF_CEILING_TOLERANCE_M):
            # Op of boven hoogste room level -> dak (paars)
            _register_element_color(
                element_color_map, cid,
                "exterior", COLOR_EXTERIOR_ROOF, "exterior",
            )
        else:
            # Tussenplafond -> aangrenzend (blauw)
            _register_element_color(
                element_color_map, cid,
                "adjacent_room", COLOR_ADJACENT_HEATED,
                "adjacent_room",
            )
        supplement_count += 1

    return supplement_count


# =============================================================================
# Helperfuncties — Kleur logica
# =============================================================================
def _get_color_for_boundary(boundary_type, position_type, host_category):
    """Bepaal de kleur op basis van boundary type en positie.

    Args:
        boundary_type: "exterior", "ground", "unheated_space", "adjacent_room"
        position_type: "wall", "floor", "ceiling"
        host_category: "Wall", "Floor", "Roof", "Ceiling"

    Returns:
        tuple: (r, g, b) kleur, of None als niet te kleuren
    """
    if boundary_type == "ground":
        return COLOR_GROUND

    if boundary_type == "adjacent_room":
        return COLOR_ADJACENT_HEATED

    if boundary_type == "unheated_space":
        return COLOR_UNHEATED_SPACE

    if boundary_type == "exterior":
        # Dak of buitenplafond
        if host_category == "Roof":
            return COLOR_EXTERIOR_ROOF
        if position_type == "ceiling":
            return COLOR_EXTERIOR_ROOF
        # Buitenwand of buitenvloer
        return COLOR_EXTERIOR_WALL

    return None


def _register_element_color(
    color_map, element_id, boundary_type, color_rgb, category
):
    """Registreer een element met kleur, deduplicatie via prioriteit.

    Args:
        color_map: dict {element_id: (priority, color_rgb, category)}
        element_id: int — Revit element ID
        boundary_type: str — voor prioriteitsbepaling
        color_rgb: tuple (r, g, b)
        category: str — beschrijvend label
    """
    priority = BOUNDARY_TYPE_PRIORITY.get(boundary_type, 0)
    existing = color_map.get(element_id)
    if existing is None or priority > existing[0]:
        color_map[element_id] = (priority, color_rgb, category)


def _apply_color_override(view, element_id_int, color_rgb, solid_pattern_id):
    """Pas kleur-override toe op een element in de view.

    Args:
        view: View3D
        element_id_int: int — Revit element ID
        color_rgb: tuple (r, g, b)
        solid_pattern_id: ElementId of None
    """
    try:
        eid = ElementId(element_id_int)
        color = Color(color_rgb[0], color_rgb[1], color_rgb[2])

        override = OverrideGraphicSettings()

        # Surface patterns (zichtbaar vlak)
        if solid_pattern_id is not None:
            override.SetSurfaceForegroundPatternId(solid_pattern_id)
            override.SetSurfaceForegroundPatternColor(color)
            override.SetCutForegroundPatternId(solid_pattern_id)
            override.SetCutForegroundPatternColor(color)

        # Lijnkleur
        override.SetProjectionLineColor(color)

        view.SetElementOverrides(eid, override)
    except Exception:
        # Element niet zichtbaar in deze view — overslaan
        pass


# =============================================================================
# Helperfuncties — Legenda
# =============================================================================
_LEGEND_CSS = """
<style>
.wv-legend {
    font-family: 'Segoe UI', Tahoma, sans-serif;
    max-width: 700px;
    margin: 16px 0;
}
.wv-legend h3 {
    margin: 0 0 10px 0;
    color: #1a5276;
}
.wv-legend table {
    border-collapse: collapse;
    width: 100%;
    font-size: 13px;
}
.wv-legend th {
    background: #f0f0f0;
    text-align: left;
    padding: 6px 12px;
    border-bottom: 2px solid #ccc;
    font-weight: 600;
}
.wv-legend td {
    padding: 5px 12px;
    border-bottom: 1px solid #eee;
}
.wv-legend tr:hover { background: #fafafa; }
.wv-swatch {
    display: inline-block;
    width: 20px;
    height: 14px;
    border: 1px solid #999;
    border-radius: 2px;
    vertical-align: middle;
    margin-right: 6px;
}
td.num {
    text-align: right;
    font-variant-numeric: tabular-nums;
}
.wv-summary {
    font-size: 12px;
    color: #555;
    margin-top: 8px;
}
</style>
"""


def _print_legend(output, element_color_map, heated_count, total_count):
    """Print een HTML legenda met kleur-swatches en tellingen.

    Args:
        output: pyRevit output object
        element_color_map: dict {element_id: (priority, color_rgb, category)}
        heated_count: int — aantal verwarmde ruimten
        total_count: int — totaal aantal ruimten
    """
    # Tel elementen per categorie
    category_counts = {}
    for _, (_, color_rgb, category) in element_color_map.items():
        if category not in category_counts:
            category_counts[category] = {"count": 0, "color": color_rgb}
        category_counts[category]["count"] += 1

    html = [_LEGEND_CSS, '<div class="wv-legend">']
    html.append('<h3>Warmteverlies View — Legenda</h3>')
    html.append(
        '<div class="wv-summary">'
        'Verwarmde ruimten: <b>{0}</b> van {1} totaal &middot; '
        'Gekleurde elementen: <b>{2}</b>'
        '</div>'.format(
            heated_count, total_count, len(element_color_map)
        )
    )

    html.append("<table>")
    html.append(
        "<tr><th>Type</th><th>Kleur</th>"
        '<th style="text-align:right">Elementen</th></tr>'
    )

    for label, color_rgb in LEGEND_ITEMS:
        # Zoek count voor deze kleur
        count = 0
        for cat_data in category_counts.values():
            if cat_data["color"] == color_rgb:
                count += cat_data["count"]

        swatch = (
            '<span class="wv-swatch" '
            'style="background:rgb({0},{1},{2})"></span>'.format(
                color_rgb[0], color_rgb[1], color_rgb[2]
            )
        )
        html.append(
            "<tr>"
            "<td>{0}</td>"
            "<td>{1}{2}</td>"
            '<td class="num">{3}</td>'
            "</tr>".format(
                label, swatch,
                "rgb({0},{1},{2})".format(
                    color_rgb[0], color_rgb[1], color_rgb[2]
                ),
                count,
            )
        )

    html.append("</table>")
    html.append("</div>")

    output.print_html("".join(html))


# =============================================================================
# Hoofdfunctie
# =============================================================================
def run_warmteverlies_view(doc):
    """Maak een diagnostische 3D view met kleurcodes per grenstype."""
    output = script.get_output()
    output.print_md("## Warmteverlies View — Diagnostische controle")

    # --- Stap 1: Rooms ophalen en functies mappen ---
    output.print_md("**Stap 1:** Rooms verzamelen...")
    rooms = collect_rooms(doc)
    if not rooms:
        forms.alert(
            "Geen rooms gevonden in het model.\n"
            "Plaats rooms via Architecture > Room.",
            title="Geen Rooms",
        )
        return

    rooms = map_all_rooms(rooms)
    output.print_md("Gevonden: **{0}** rooms".format(len(rooms)))

    # --- Stap 2: Heated/unheated splitsen ---
    heated_rooms = [r for r in rooms if r.get("is_heated")]
    if not heated_rooms:
        forms.alert(
            "Geen verwarmde ruimten gevonden.\n"
            "Controleer de ruimtenamen in het model.",
            title="Geen verwarmde ruimten",
        )
        return

    heated_room_ids = set(r["element_id"] for r in heated_rooms)
    output.print_md(
        "Verwarmde ruimten: **{0}** van {1}".format(
            len(heated_rooms), len(rooms)
        )
    )

    # --- Stap 3: Wanden pre-collecten ---
    output.print_md("**Stap 2:** Wanden verzamelen voor assembly detectie...")
    all_walls_data = collect_all_walls(doc)
    all_wall_ids = set(wd["element_id"] for wd in all_walls_data)
    output.print_md(
        "Wanden: **{0}**".format(len(all_walls_data))
    )

    # --- Stap 4: Analyse per verwarmde ruimte ---
    output.print_md("**Stap 3:** Grensvlakken analyseren...")
    element_color_map = {}  # {element_id: (priority, color_rgb, category)}
    boundary_wall_ids = set()
    assembly_outer_ids = set()

    for room_data in heated_rooms:
        room_element = room_data["element"]

        # Boundary analyse
        boundaries = analyze_boundaries(doc, room_element)
        boundaries = classify_boundaries(
            doc, room_data, boundaries, rooms, heated_room_ids
        )

        for boundary in boundaries:
            host = boundary.get("host_element")
            host_id = boundary.get("host_element_id")
            host_category = boundary.get("host_category")
            position_type = boundary.get("position_type")
            boundary_type = boundary.get("boundary_type", "exterior")

            if host_id is None:
                continue

            # Track boundary wand IDs
            if host_category == "Wall":
                boundary_wall_ids.add(host_id)

            # Curtain wall → behandel als raam (cyaan)
            if (host_category == "Wall"
                    and host is not None
                    and _is_curtain_wall(doc, host)):
                _register_element_color(
                    element_color_map, host_id,
                    "opening", COLOR_WINDOW, "window",
                )
                # Panels en mullions ook als raam kleuren
                for sub_id in _get_curtain_wall_sub_ids(host):
                    _register_element_color(
                        element_color_map, sub_id,
                        "opening", COLOR_WINDOW, "window",
                    )
                continue  # geen assembly detectie voor curtain walls

            # Kleur voor dit grensvlak
            color = _get_color_for_boundary(
                boundary_type, position_type, host_category
            )
            if color is not None:
                _register_element_color(
                    element_color_map, host_id,
                    boundary_type, color, boundary_type,
                )

            # Assembly detectie bij exterior/unheated wanden
            if (host_category == "Wall"
                    and host is not None
                    and boundary_type in ("exterior", "unheated_space")):
                face_normal = boundary.get(
                    "face_normal", (0.0, 0.0, 0.0)
                )
                assembly = resolve_wall_assembly(
                    doc, host, face_normal, all_walls_data
                )
                # Buitenste lagen van assembly registreren
                if len(assembly) > 1:
                    for wall in assembly[1:]:
                        outer_id = wall.Id.IntegerValue
                        assembly_outer_ids.add(outer_id)

            # Openings extraheren bij buitenwanden
            if (host_category == "Wall"
                    and host is not None
                    and boundary_type == "exterior"):
                _color_wall_openings(doc, host, element_color_map)

    # --- Stap 5: Assembly wanden (niet-grens) toevoegen ---
    for outer_id in assembly_outer_ids:
        if outer_id not in boundary_wall_ids:
            _register_element_color(
                element_color_map, outer_id,
                "assembly", COLOR_ASSEMBLY_NON_BOUNDARY, "assembly",
            )

    boundary_count = len(element_color_map)

    # --- Stap 6: Supplementaire scan — vang gemiste elementen ---
    output.print_md(
        "**Stap 4:** Supplementaire element scan..."
    )
    supplement_count = _supplement_uncolored_elements(
        doc, element_color_map, rooms, boundary_wall_ids
    )
    output.print_md(
        "Boundary-analyse: **{0}** elementen, "
        "supplement: **{1}** extra".format(boundary_count, supplement_count)
    )

    # --- Stap 7: View aanmaken en kleuren toepassen ---
    output.print_md("**Stap 5:** 3D view aanmaken en kleuren toepassen...")

    t = Transaction(doc, "WV - Controle view aanmaken")
    t.Start()
    try:
        _delete_existing_view(doc, VIEW_NAME)
        view_3d = _create_3d_view(doc, VIEW_NAME)
        _hide_irrelevant_categories(doc, view_3d)
        solid_pattern_id = _get_solid_fill_pattern_id(doc)

        for elem_id, (_, color_rgb, _) in element_color_map.items():
            _apply_color_override(
                view_3d, elem_id, color_rgb, solid_pattern_id
            )

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(
            "Fout bij aanmaken view:\n{0}".format(str(ex)),
            title="Fout",
        )
        return

    # --- Stap 7: View activeren ---
    try:
        revit.uidoc.ActiveView = view_3d
    except Exception:
        output.print_md(
            "*View kon niet automatisch geactiveerd worden. "
            "Open '{0}' handmatig.*".format(VIEW_NAME)
        )

    # --- Stap 8: Legenda ---
    output.print_md("---")
    _print_legend(
        output, element_color_map,
        len(heated_rooms), len(rooms),
    )
    output.print_md(
        "View **{0}** is aangemaakt en geactiveerd.".format(VIEW_NAME)
    )


# =============================================================================
# Entry point
# =============================================================================
if __name__ == "__main__":
    doc = revit.doc
    if doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run_warmteverlies_view(doc)
