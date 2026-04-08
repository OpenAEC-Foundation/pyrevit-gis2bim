# -*- coding: utf-8 -*-
"""Bouwt Thermal Import JSON uit raycast scan resultaten.

Output format is 'thermal-import' v1, niet het ISSO 51 project format.
Dit format wordt geimporteerd via warmteverlies.open-aec.com/import/thermal.
"""
import datetime

from warmteverlies.constants import DEFAULT_THETA_E
from warmteverlies.room_collector import generate_room_id


# =============================================================================
# Pseudo-room defaults
# =============================================================================
PSEUDO_ROOMS = {
    "outside": {
        "id": "outside",
        "name": "Buiten",
        "room_type": "outside",
        "temperature": DEFAULT_THETA_E,
    },
    "water": {
        "id": "water",
        "name": "Water",
        "room_type": "water",
        "temperature": 10.0,
    },
    "ground": {
        "id": "ground",
        "name": "Grond",
        "room_type": "ground",
        "temperature": 10.0,
    },
}

# Terminal types die naar pseudo-rooms verwijzen
TERMINAL_PSEUDO_TYPES = {"outside", "water", "ground"}


def build_thermal_import(project_name, rooms_data, scan_results):
    """Bouw het complete thermal import JSON object.

    Args:
        project_name: Projectnaam uit Revit doc.Title
        rooms_data: Lijst van room dicts uit room_collector.py
        scan_results: Dict keyed by room element_id, waarde is dict met
            constructions, openings, open_connections

    Returns:
        dict: Thermal import JSON conform v1 format
    """
    rooms = _build_rooms(rooms_data, scan_results)
    constructions = _build_constructions(rooms_data, scan_results)
    openings = _build_openings(rooms_data, scan_results, constructions)
    open_connections = _build_open_connections(rooms_data, scan_results)

    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

    return {
        "version": "1.0",
        "format": "thermal-import",
        "source": "revit-raycast",
        "exported_at": now,
        "project_name": project_name,
        "rooms": rooms,
        "constructions": constructions,
        "openings": openings,
        "open_connections": open_connections,
    }


def _build_rooms(rooms_data, scan_results):
    """Bouw room array inclusief pseudo-rooms.

    Genereert room dicts voor alle echte rooms en voegt pseudo-rooms
    toe op basis van welke terminal conditions voorkomen in de scan.

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id

    Returns:
        list[dict]: Room dicts inclusief pseudo-rooms
    """
    rooms = []
    needed_pseudo = {"outside"}  # outside altijd aanwezig

    for room_data in rooms_data:
        room_id = generate_room_id(room_data, rooms_data)

        if room_data.get("is_heated", True):
            room_type = "heated"
        else:
            room_type = "unheated"

        rooms.append({
            "id": room_id,
            "name": room_data.get("name", "Onbekend"),
            "room_type": room_type,
            "temperature": room_data.get("temperature", 20.0),
            "floor_area_m2": round(
                room_data.get("floor_area_m2", 0.0), 2
            ),
            "height_m": round(room_data.get("height_m", 2.6), 2),
            "level": room_data.get("level", ""),
        })

    # Scan terminal conditions om te bepalen welke pseudo-rooms nodig zijn
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


def _resolve_room_b(terminal_type, room_data, rooms_data):
    """Vertaal een terminal_type naar een room_b ID.

    Args:
        terminal_type: "outside", "water", "ground", of een element_id (int)
        room_data: Huidige room data
        rooms_data: Alle room dicts

    Returns:
        str: Room ID voor room_b
    """
    if terminal_type in TERMINAL_PSEUDO_TYPES:
        return terminal_type

    # Integer element_id: zoek bijbehorende room
    if isinstance(terminal_type, int):
        for r in rooms_data:
            if r.get("element_id") == terminal_type:
                return generate_room_id(r, rooms_data)
        # Niet gevonden: fallback naar outside
        return "outside"

    # String die geen pseudo-type is: probeer als element_id
    try:
        elem_id = int(terminal_type)
        return _resolve_room_b(elem_id, room_data, rooms_data)
    except (ValueError, TypeError):
        return "outside"


def _build_constructions(rooms_data, scan_results):
    """Bouw construction array uit zone-data.

    Genereert unieke C-IDs (C001, C002, ...) voor elke constructie
    uit de scan resultaten.

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id

    Returns:
        list[dict]: Construction dicts met unieke IDs
    """
    constructions = []
    counter = 1

    for room_data in rooms_data:
        if not room_data.get("is_heated", True):
            continue

        room_id = generate_room_id(room_data, rooms_data)
        elem_id = room_data.get("element_id")

        room_scan = scan_results.get(elem_id, {})
        for constr in room_scan.get("constructions", []):
            c_id = "C{0:03d}".format(counter)

            direction = constr.get("direction", "")
            position_type = constr.get("position_type", "wall")
            z_min = constr.get("z_min", 0.0)
            z_max = constr.get("z_max", 0.0)

            name = "{dir} {pos} zone ({z_min:.2f}-{z_max:.2f}m)".format(
                dir=direction,
                pos=position_type,
                z_min=z_min,
                z_max=z_max,
            )

            terminal_type = constr.get("terminal_type", "outside")
            room_b = _resolve_room_b(terminal_type, room_data, rooms_data)

            # Laagopbouw kopieren, "lambda" hernoemen naar "lambda_value"
            layers = _convert_layers(constr.get("layers", []))

            construction = {
                "id": c_id,
                "name": name,
                "room_a": room_id,
                "room_b": room_b,
                "direction": direction,
                "position_type": position_type,
                "area_m2": round(constr.get("area_m2", 0.0), 2),
            }

            if layers:
                construction["layers"] = layers

            constructions.append(construction)
            counter += 1

    return constructions


def _convert_layers(raw_layers):
    """Converteer laagopbouw, hernoem 'lambda' naar 'lambda_value'.

    Args:
        raw_layers: Lijst van layer dicts uit scan

    Returns:
        list[dict]: Geconverteerde layers
    """
    converted = []
    for layer in raw_layers:
        new_layer = {}
        for key, value in layer.items():
            if key == "lambda":
                new_layer["lambda_value"] = value
            else:
                new_layer[key] = value
        converted.append(new_layer)
    return converted


def _build_openings(rooms_data, scan_results, constructions_list):
    """Bouw opening array.

    Koppelt elke opening aan een parent constructie op basis van
    direction match.

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id
        constructions_list: Eerder gebouwde construction dicts

    Returns:
        list[dict]: Opening dicts met unieke IDs
    """
    openings = []
    counter = 1

    # Index: (room_a, direction) -> construction_id voor snelle lookup
    constr_by_room_dir = {}
    for c in constructions_list:
        key = (c["room_a"], c.get("direction", ""))
        # Eerste match per room+direction bewaren
        if key not in constr_by_room_dir:
            constr_by_room_dir[key] = c["id"]

    for room_data in rooms_data:
        if not room_data.get("is_heated", True):
            continue

        room_id = generate_room_id(room_data, rooms_data)
        elem_id = room_data.get("element_id")

        room_scan = scan_results.get(elem_id, {})
        for opening in room_scan.get("openings", []):
            o_id = "O{0:03d}".format(counter)

            wall_direction = opening.get("wall_direction", "")
            lookup_key = (room_id, wall_direction)
            construction_id = constr_by_room_dir.get(lookup_key, "")

            openings.append({
                "id": o_id,
                "type": opening.get("type", "window"),
                "construction_id": construction_id,
                "width_mm": opening.get("width_mm", 0),
                "height_mm": opening.get("height_mm", 0),
                "u_value": round(opening.get("u_value", 1.6), 2),
                "quantity": 1,
            })
            counter += 1

    return openings


def _build_open_connections(rooms_data, scan_results):
    """Bouw open_connections array.

    Args:
        rooms_data: Lijst van room dicts
        scan_results: Scan resultaten per room element_id

    Returns:
        list[dict]: Open connection dicts met unieke IDs
    """
    connections = []
    counter = 1

    for room_data in rooms_data:
        if not room_data.get("is_heated", True):
            continue

        room_id = generate_room_id(room_data, rooms_data)
        elem_id = room_data.get("element_id")

        room_scan = scan_results.get(elem_id, {})
        for conn in room_scan.get("open_connections", []):
            oc_id = "OC{0:03d}".format(counter)

            terminal_type = conn.get("terminal_type", "outside")
            room_b = _resolve_room_b(
                terminal_type, room_data, rooms_data
            )

            connections.append({
                "id": oc_id,
                "room_a": room_id,
                "room_b": room_b,
                "direction": conn.get("direction", ""),
                "width_m": round(conn.get("width_m", 0.0), 3),
                "height_m": round(conn.get("height_m", 0.0), 3),
            })
            counter += 1

    return connections
