# -*- coding: utf-8 -*-
"""Bouwt Thermal Import JSON uit raycast scan resultaten.

Output format is 'thermal-import' v1, conform thermal-import.schema.json.
Dit format wordt geimporteerd via warmteverlies.open-aec.com/import/thermal.
"""
import datetime

from warmteverlies.room_collector import generate_room_id


# =============================================================================
# Pseudo-room definities (schema-conform)
# =============================================================================
PSEUDO_ROOMS = {
    "outside": {
        "id": "room-outside",
        "name": "Buiten",
        "type": "outside",
    },
    "water": {
        "id": "room-water",
        "name": "Water",
        "type": "water",
    },
    "ground": {
        "id": "room-ground",
        "name": "Grond",
        "type": "ground",
    },
}

# Terminal types die naar pseudo-rooms verwijzen
TERMINAL_PSEUDO_TYPES = {"outside", "water", "ground"}

# Minimale constructie-oppervlakte na consolidatie (m2)
MIN_CONSTRUCTION_AREA_M2 = 0.25

# Kompasrichtingen: azimuth ranges (graden, N=0, clockwise)
_COMPASS_RANGES = [
    (337.5, 360.0, "N"),
    (0.0, 22.5, "N"),
    (22.5, 67.5, "NE"),
    (67.5, 112.5, "E"),
    (112.5, 157.5, "SE"),
    (157.5, 202.5, "S"),
    (202.5, 247.5, "SW"),
    (247.5, 292.5, "W"),
    (292.5, 337.5, "NW"),
]


def _azimuth_to_compass(azimuth_deg):
    """Vertaal azimuth (graden, 0=N, clockwise) naar kompasrichting.

    Args:
        azimuth_deg: Azimuth in graden (0-360)

    Returns:
        str of None: Kompasrichting (N, NE, E, SE, S, SW, W, NW) of None
    """
    if azimuth_deg is None:
        return None
    az = azimuth_deg % 360.0
    for low, high, label in _COMPASS_RANGES:
        if low <= az < high:
            return label
    return "N"


def build_thermal_import(project_name, rooms_data, scan_results):
    """Bouw het complete thermal import JSON object.

    Args:
        project_name: Projectnaam uit Revit doc.Title
        rooms_data: Lijst van room dicts uit room_collector.py
        scan_results: Dict keyed by room element_id, waarde is dict met
            constructions, openings, open_connections

    Returns:
        dict: Thermal import JSON conform v1 schema
    """
    # Stap 1: bouw room ID mapping (intern ID -> schema ID)
    room_id_map = _build_room_id_map(rooms_data)

    # Stap 2: bouw rooms (schema-conform)
    rooms = _build_rooms(rooms_data, scan_results, room_id_map)

    # Stap 3: bouw constructies met consolidatie. Geeft ook een
    # fingerprint_index terug voor de opening-to-zone lookup in stap 4.
    constructions, fingerprint_index = _build_constructions(
        rooms_data, scan_results, room_id_map
    )

    # Stap 4: bouw openings (verwijst naar geconsolideerde construction IDs).
    # Gebruikt fingerprint_index voor correcte sub-zone toewijzing (Bug E.3).
    openings = _build_openings(
        rooms_data, scan_results, constructions, fingerprint_index,
        room_id_map
    )

    # Stap 5: bouw open connections
    open_connections = _build_open_connections(
        rooms_data, scan_results, room_id_map
    )

    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

    return {
        "version": "1.0",
        "source": "revit-raycast",
        "exported_at": now,
        "project_name": project_name,
        "rooms": rooms,
        "constructions": constructions,
        "openings": openings,
        "open_connections": open_connections,
    }


def _build_room_id_map(rooms_data):
    """Bouw een mapping van interne room ID naar schema-conform room ID.

    Schema vereist: room-0, room-1, room-2, ... (oplopend)
    Pseudo-rooms: room-outside, room-ground, room-water

    Args:
        rooms_data: Lijst van room dicts

    Returns:
        dict: {intern_room_id: schema_room_id}
    """
    room_id_map = {}
    for idx, room_data in enumerate(rooms_data):
        intern_id = generate_room_id(room_data, rooms_data)
        schema_id = "room-{0}".format(idx)
        room_id_map[intern_id] = schema_id
        # Ook element_id mappen voor lookup vanuit terminal_type
        elem_id = room_data.get("element_id")
        if elem_id is not None:
            room_id_map[elem_id] = schema_id

    # Pseudo-rooms
    for pseudo_key in TERMINAL_PSEUDO_TYPES:
        room_id_map[pseudo_key] = "room-{0}".format(pseudo_key)

    return room_id_map


def _build_rooms(rooms_data, scan_results, room_id_map):
    """Bouw room array inclusief pseudo-rooms.

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id
        room_id_map: Mapping intern ID -> schema ID

    Returns:
        list: Room dicts conform schema
    """
    rooms = []
    needed_pseudo = {"outside"}  # outside altijd aanwezig

    for idx, room_data in enumerate(rooms_data):
        schema_id = "room-{0}".format(idx)

        if room_data.get("is_heated", True):
            room_type = "heated"
        else:
            room_type = "unheated"

        area_m2 = round(room_data.get("floor_area_m2", 0.0), 2)
        height_m = round(room_data.get("height_m", 2.6), 2)
        volume_m3 = round(area_m2 * height_m, 2)

        room = {
            "id": schema_id,
            "name": room_data.get("name", "Onbekend"),
            "type": room_type,
            "level": room_data.get("level_name", room_data.get("level", "")),
            "area_m2": area_m2,
            "height_m": height_m,
            "volume_m3": volume_m3,
        }

        # Revit element ID toevoegen als beschikbaar
        elem_id = room_data.get("element_id")
        if elem_id is not None:
            room["revit_id"] = elem_id

        rooms.append(room)

    # Scan terminal conditions voor benodigde pseudo-rooms
    for room_scan in scan_results.values():
        for constr in room_scan.get("constructions", []):
            terminal = constr.get("terminal_type", "")
            if terminal in TERMINAL_PSEUDO_TYPES:
                needed_pseudo.add(terminal)

    # Pseudo-rooms toevoegen
    for pseudo_key in sorted(needed_pseudo):
        if pseudo_key in PSEUDO_ROOMS:
            rooms.append(dict(PSEUDO_ROOMS[pseudo_key]))

    return rooms


def _resolve_room_b(terminal_type, room_id_map):
    """Vertaal een terminal_type naar een schema-conform room_b ID.

    Args:
        terminal_type: "outside", "water", "ground", of een element_id
        room_id_map: Mapping intern ID -> schema ID

    Returns:
        str: Schema-conform room_b ID
    """
    # Pseudo-room types
    if terminal_type in TERMINAL_PSEUDO_TYPES:
        return "room-{0}".format(terminal_type)

    # Integer element_id: zoek in de map
    if isinstance(terminal_type, int):
        mapped = room_id_map.get(terminal_type)
        if mapped:
            return mapped
        return "room-outside"

    # String die geen pseudo-type is: probeer als element_id
    try:
        elem_id = int(terminal_type)
        mapped = room_id_map.get(elem_id)
        if mapped:
            return mapped
    except (ValueError, TypeError):
        pass

    return "room-outside"


def _make_layer_fingerprint(layers):
    """Maak een hashbare fingerprint van een laagopbouw.

    Canonicaliseert air_gap lagen met `__air_gap__` marker zodat deze
    fingerprint matcht met `_make_zone_fingerprint` uit raycast_scanner.
    Dit is kritisch voor de opening-to-zone lookup in `_build_openings`
    (Bug E.3): zonder deze canonicalisatie hebben builder en scanner
    verschillende fingerprints voor dezelfde zone met air-gap lagen,
    waardoor openings op de verkeerde sub-zone worden geplaatst.

    Args:
        layers: Lijst van layer dicts (van binnen naar buiten)

    Returns:
        tuple: ((material, thickness_mm), ...) gesorteerd van binnen
            naar buiten. Voor air_gap lagen is material `"__air_gap__"`.
    """
    parts = []
    for layer in layers:
        if layer.get("type") == "air_gap":
            material = "__air_gap__"
        else:
            material = layer.get("material", "")
        thickness = layer.get("thickness_mm", 0)
        parts.append((material, thickness))
    return tuple(parts)


def _normalize_fingerprint(fp):
    """Normaliseer een fingerprint tot hashable tuple of tuples.

    De scanner zet `zone_layer_fingerprint` op openings als een
    `list of [material, thickness_mm]` (JSON-vriendelijk), terwijl
    de builder `tuple of (material, thickness_mm)` gebruikt (hashable
    voor dict keys). Deze helper normaliseert beide formats tot een
    tuple-of-tuples zodat ze als dict key bruikbaar zijn.

    Args:
        fp: fingerprint in list-of-lists of tuple-of-tuples format,
            of None

    Returns:
        tuple of tuples, of None als input None is
    """
    if fp is None:
        return None
    return tuple(tuple(pair) for pair in fp)


def _convert_layers(raw_layers):
    """Converteer laagopbouw naar schema-conform formaat.

    Voegt 'type' (solid/air_gap) en 'distance_from_interior_mm' toe.
    Behoudt 'lambda' (NIET hernoemen naar lambda_value).

    Args:
        raw_layers: Lijst van layer dicts uit scan

    Returns:
        list: Geconverteerde layers conform schema
    """
    converted = []
    cumulative_distance = 0.0

    for layer in raw_layers:
        new_layer = {
            "material": layer.get("material", ""),
            "thickness_mm": layer.get("thickness_mm", 0),
            "distance_from_interior_mm": round(cumulative_distance, 1),
        }

        # Type bepalen: air_gap als expliciet aangegeven, anders solid
        if layer.get("type") == "air_gap" or layer.get("is_air_gap", False):
            new_layer["type"] = "air_gap"
        else:
            new_layer["type"] = "solid"

        # Lambda behouden als aanwezig (schema veldnaam is "lambda")
        if "lambda" in layer and layer["lambda"] is not None:
            new_layer["lambda"] = layer["lambda"]
        elif "lambda_value" in layer and layer["lambda_value"] is not None:
            # Fallback: als scanner toevallig lambda_value gebruikt
            new_layer["lambda"] = layer["lambda_value"]

        cumulative_distance += layer.get("thickness_mm", 0)
        converted.append(new_layer)

    return converted


def _build_constructions(rooms_data, scan_results, room_id_map):
    """Bouw construction array met consolidatie.

    Consolidatie-logica:
    1. Verzamel alle ruwe constructies per room
    2. Maak fingerprint: (room_a, room_b, orientation, compass, layer_hash)
    3. Groepeer en merge (som van gross_area_m2)
    4. Filter constructies < 0.25 m2 na merge

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id
        room_id_map: Mapping intern ID -> schema ID

    Returns:
        tuple: (constructions, fingerprint_index)
            - constructions: lijst van schema-conforme construction dicts
            - fingerprint_index: dict met key
              `(room_a, compass, layer_fingerprint)` -> construction_id.
              Gebruikt door `_build_openings` om openings aan de juiste
              verticale sub-zone te koppelen (Bug E.3 fix).
    """
    # Fase 1: Verzamel alle ruwe constructies
    raw_constructions = []

    for room_data in rooms_data:
        if not room_data.get("is_heated", True):
            continue

        intern_id = generate_room_id(room_data, rooms_data)
        room_a = room_id_map.get(intern_id, intern_id)
        elem_id = room_data.get("element_id")

        room_scan = scan_results.get(elem_id, {})
        for constr in room_scan.get("constructions", []):
            terminal_type = constr.get("terminal_type", "outside")
            room_b = _resolve_room_b(terminal_type, room_id_map)

            orientation = constr.get("position_type", "wall")

            # Kompasrichting bepalen (alleen relevant voor walls)
            compass = None
            if orientation == "wall":
                # Probeer azimuth eerst, dan direction als string
                azimuth = constr.get("azimuth_deg")
                if azimuth is not None:
                    compass = _azimuth_to_compass(azimuth)
                else:
                    direction = constr.get("direction", "")
                    if direction and direction.upper() in (
                        "N", "NE", "E", "SE", "S", "SW", "W", "NW"
                    ):
                        compass = direction.upper()

            area = constr.get("area_m2", 0.0)
            raw_layers = constr.get("layers", [])
            converted_layers = _convert_layers(raw_layers)
            layer_fp = _make_layer_fingerprint(raw_layers)

            # Revit metadata bewaren (eerste hit wint bij merge)
            revit_element_id = constr.get("revit_element_id")
            revit_type_name = constr.get("revit_type_name", "")

            raw_constructions.append({
                "room_a": room_a,
                "room_b": room_b,
                "orientation": orientation,
                "compass": compass,
                "area": area,
                "layers": converted_layers,
                "layer_fingerprint": layer_fp,
                "revit_element_id": revit_element_id,
                "revit_type_name": revit_type_name,
            })

    # Fase 2: Consolidatie — groepeer per fingerprint
    groups = {}
    for raw in raw_constructions:
        fp = (
            raw["room_a"],
            raw["room_b"],
            raw["orientation"],
            raw["compass"],
            raw["layer_fingerprint"],
        )
        if fp not in groups:
            groups[fp] = {
                "room_a": raw["room_a"],
                "room_b": raw["room_b"],
                "orientation": raw["orientation"],
                "compass": raw["compass"],
                "total_area": 0.0,
                "layers": raw["layers"],
                "layer_fingerprint": raw["layer_fingerprint"],
                "revit_element_id": raw["revit_element_id"],
                "revit_type_name": raw["revit_type_name"],
            }
        groups[fp]["total_area"] += raw["area"]
        # Bewaar eerste niet-None revit_element_id
        if (groups[fp]["revit_element_id"] is None
                and raw["revit_element_id"] is not None):
            groups[fp]["revit_element_id"] = raw["revit_element_id"]

    # Fase 3: Filter en bouw schema-conforme output
    constructions = []
    # Secundaire index: (room_a, compass, layer_fingerprint) -> construction_id
    # Gebruikt door `_build_openings` om openings aan de juiste verticale
    # sub-zone te koppelen (Bug E.3 fix). De fingerprint matcht met
    # `zone_layer_fingerprint` op opening dicts uit raycast_scanner.
    fingerprint_index = {}
    counter = 0

    # Sorteer voor deterministische output
    sorted_keys = sorted(groups.keys())

    for key in sorted_keys:
        group = groups[key]
        total_area = round(group["total_area"], 2)

        # Filter kleine constructies na merge
        if total_area < MIN_CONSTRUCTION_AREA_M2:
            continue

        c_id = "constr-{0}".format(counter)

        construction = {
            "id": c_id,
            "room_a": group["room_a"],
            "room_b": group["room_b"],
            "orientation": group["orientation"],
            "gross_area_m2": total_area,
        }

        # Compass alleen toevoegen voor walls (en alleen als bekend)
        compass = group.get("compass")
        if compass is not None:
            construction["compass"] = compass

        # Revit metadata
        if group["revit_element_id"] is not None:
            construction["revit_element_id"] = group["revit_element_id"]
        if group["revit_type_name"]:
            construction["revit_type_name"] = group["revit_type_name"]

        # Layers
        if group["layers"]:
            construction["layers"] = group["layers"]

        constructions.append(construction)

        # Registreer in fingerprint index (voor opening lookup)
        layer_fp = group.get("layer_fingerprint")
        if layer_fp is not None:
            fp_key = (group["room_a"], compass or "", layer_fp)
            # Eerste match wint (deterministisch door sorted_keys)
            if fp_key not in fingerprint_index:
                fingerprint_index[fp_key] = c_id

        counter += 1

    return constructions, fingerprint_index


def _build_heated_room_lookup(rooms_data, room_id_map):
    """Bouw een mapping van schema_room_id -> is_heated (bool).

    Gebruikt voor Bug E.5c: interne deuren tussen twee verwarmde
    ruimtes moeten uit de opening-export worden gefilterd.

    Fallback-logica voor onbekende/ontbrekende `is_heated`:
    - Pseudo-rooms (`room-outside`, `room-water`, `room-ground`) zijn
      per definitie unheated (exterieur).
    - Ruimtes met naam `"Buiten"` of schema id `"room-outside"` zijn
      unheated.
    - Default voor 'echte' rooms zonder `is_heated`-veld: heated
      (consistent met bestaand patroon `room_data.get("is_heated", True)`
      elders in dit bestand).

    Args:
        rooms_data: Lijst van room dicts
        room_id_map: Mapping intern ID -> schema ID

    Returns:
        dict: {schema_room_id: bool}
    """
    heated_map = {}

    for room_data in rooms_data:
        intern_id = generate_room_id(room_data, rooms_data)
        schema_id = room_id_map.get(intern_id)
        if schema_id is None:
            continue

        raw = room_data.get("is_heated")
        if raw is None:
            # Fallback: naam "Buiten" of schema id room-outside = unheated
            name = (room_data.get("name") or "").strip().lower()
            if name == "buiten" or schema_id == "room-outside":
                heated_map[schema_id] = False
            else:
                heated_map[schema_id] = True
        else:
            heated_map[schema_id] = bool(raw)

    # Pseudo-rooms zijn altijd unheated (outside, water, ground)
    for pseudo_key in TERMINAL_PSEUDO_TYPES:
        heated_map["room-{0}".format(pseudo_key)] = False

    return heated_map


def _build_openings(rooms_data, scan_results, constructions_list,
                    fingerprint_index, room_id_map):
    """Bouw opening array conform schema.

    Koppelt elke opening aan een parent constructie via een 3-staps
    lookup:

    1. **Primair** — `(room_a, compass, zone_layer_fingerprint)` match
       via `fingerprint_index` uit `_build_constructions`. Dit plaatst
       een opening correct op de verticale sub-zone (bijv. gevel ipv
       betonstrip) wanneer een muur meerdere layer-stacks heeft (Bug
       E.3 fix).
    2. **Fallback** — `(room_a, compass)` eerste match. Legacy gedrag
       voor openings zonder `zone_layer_fingerprint` metadata.
    3. **Laatste redmiddel** — `(room_a, orientation="wall")` match.

    Filter (Bug E.5c): openings waarbij zowel `room_a` als `room_b`
    heated zijn worden gedropt. Voor NEN-EN 12831 / ISSO 51 zijn enkel
    openings naar buiten of naar koudere ruimtes relevant. Binnendeuren
    tussen verwarmde ruimtes dragen niet bij aan het warmteverlies.

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id
        constructions_list: Eerder gebouwde construction dicts
        fingerprint_index: dict `(room_a, compass, fp) -> construction_id`
            uit `_build_constructions`
        room_id_map: Mapping intern ID -> schema ID

    Returns:
        list: Opening dicts conform schema
    """
    openings = []
    counter = 0
    skipped_internal = 0

    # Index: (room_a, compass) -> construction_id voor snelle lookup
    constr_by_room_compass = {}
    for c in constructions_list:
        key = (c["room_a"], c.get("compass", ""))
        if key not in constr_by_room_compass:
            constr_by_room_compass[key] = c["id"]

    # Fallback index: (room_a, orientation) -> construction_id
    constr_by_room_orient = {}
    for c in constructions_list:
        key = (c["room_a"], c.get("orientation", ""))
        if key not in constr_by_room_orient:
            constr_by_room_orient[key] = c["id"]

    # Index: construction_id -> (room_a, room_b) voor E.5c filter
    constr_rooms_by_id = {}
    for c in constructions_list:
        constr_rooms_by_id[c["id"]] = (
            c.get("room_a", ""),
            c.get("room_b", ""),
        )

    # Heated lookup (E.5c): schema_room_id -> bool
    heated_by_room = _build_heated_room_lookup(rooms_data, room_id_map)

    for room_data in rooms_data:
        if not room_data.get("is_heated", True):
            continue

        intern_id = generate_room_id(room_data, rooms_data)
        room_a = room_id_map.get(intern_id, intern_id)
        elem_id = room_data.get("element_id")

        room_scan = scan_results.get(elem_id, {})
        for opening in room_scan.get("openings", []):
            o_id = "opening-{0}".format(counter)

            # Construction ID lookup via compass richting
            wall_direction = opening.get("wall_direction", "")
            compass = wall_direction.upper() if wall_direction else ""

            construction_id = ""

            # Primair: fingerprint-based lookup (Bug E.3 fix).
            # Plaatst de opening op de juiste verticale sub-zone.
            zone_fp = opening.get("zone_layer_fingerprint")
            if zone_fp is not None:
                normalized_fp = _normalize_fingerprint(zone_fp)
                fp_key = (room_a, compass, normalized_fp)
                construction_id = fingerprint_index.get(fp_key, "")

            # Fallback: eerste match op (room_a, compass)
            if not construction_id:
                lookup_key = (room_a, compass)
                construction_id = constr_by_room_compass.get(
                    lookup_key, ""
                )

            # Laatste redmiddel: zoek op orientation "wall"
            if not construction_id:
                fallback_key = (room_a, "wall")
                construction_id = constr_by_room_orient.get(
                    fallback_key, ""
                )

            # E.5c filter: skip als BEIDE kanten van de parent
            # construction heated zijn (interne deur tussen twee
            # verwarmde ruimtes draagt niet bij aan warmteverlies).
            if construction_id:
                rooms_ab = constr_rooms_by_id.get(construction_id)
                if rooms_ab is not None:
                    ra_id, rb_id = rooms_ab
                    # Default True voor onbekende rooms (veilig:
                    # niet-skippen bij twijfel zou dubbeltelling
                    # geven, maar heated_by_room bevat alle bekende
                    # rooms; onbekende = niet in map = default False
                    # zodat we alleen met ZEKERHEID heated rooms
                    # filteren).
                    ra_heated = heated_by_room.get(ra_id, False)
                    rb_heated = heated_by_room.get(rb_id, False)
                    if ra_heated and rb_heated:
                        skipped_internal += 1
                        continue

            opening_dict = {
                "id": o_id,
                "construction_id": construction_id,
                "type": opening.get("type", "window"),
                "width_mm": opening.get("width_mm", 0),
                "height_mm": opening.get("height_mm", 0),
            }

            # Optionele velden
            sill = opening.get("sill_height_mm")
            if sill is not None:
                opening_dict["sill_height_mm"] = sill

            u_val = opening.get("u_value")
            if u_val is not None:
                opening_dict["u_value"] = round(u_val, 2)

            revit_elem = opening.get("revit_element_id")
            if revit_elem is not None:
                opening_dict["revit_element_id"] = revit_elem

            revit_type = opening.get("revit_type_name")
            if revit_type:
                opening_dict["revit_type_name"] = revit_type

            openings.append(opening_dict)
            counter += 1

    if skipped_internal > 0:
        # Builder is een lib-module zonder eigen logger; print wordt
        # door pyRevit script.output opgevangen en ook in journal.
        print(
            "[thermal_json_builder] E.5c: {0} interne opening(en) "
            "geskipt (beide zijden verwarmd)".format(skipped_internal)
        )

    return openings


def _build_open_connections(rooms_data, scan_results, room_id_map):
    """Bouw open_connections array conform schema.

    Schema vereist alleen: room_a, room_b, area_m2

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id
        room_id_map: Mapping intern ID -> schema ID

    Returns:
        list: Open connection dicts conform schema
    """
    connections = []

    for room_data in rooms_data:
        if not room_data.get("is_heated", True):
            continue

        intern_id = generate_room_id(room_data, rooms_data)
        room_a = room_id_map.get(intern_id, intern_id)
        elem_id = room_data.get("element_id")

        room_scan = scan_results.get(elem_id, {})
        for conn in room_scan.get("open_connections", []):
            terminal_type = conn.get("terminal_type", "outside")
            room_b = _resolve_room_b(terminal_type, room_id_map)

            # Bereken area uit width * height
            width_m = conn.get("width_m", 0.0)
            height_m = conn.get("height_m", 0.0)
            area_m2 = conn.get("area_m2")
            if area_m2 is None:
                area_m2 = width_m * height_m

            connections.append({
                "room_a": room_a,
                "room_b": room_b,
                "area_m2": round(area_m2, 2),
            })

    return connections
