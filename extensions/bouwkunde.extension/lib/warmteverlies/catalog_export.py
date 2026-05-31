# -*- coding: utf-8 -*-
"""Export WV_BND catalogus naar ThermalImport JSON voor open-heatloss-studio.

IronPython 2.7 — geen f-strings, geen type hints.
"""
import json

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    DirectShape,
)


def _azimuth_to_compass(az):
    """Converteer azimuth (0-360 graden) naar 8-wind kompasrichting.

    Args:
        az: float in [0, 360)

    Returns:
        str: "N", "NE", "E", "SE", "S", "SW", "W", "NW"
    """
    # 8-wind: elke richting is 45 graden breed
    # N: 337.5-360 of 0-22.5, NE: 22.5-67.5, E: 67.5-112.5, etc.
    if az >= 337.5 or az < 22.5:
        return "N"
    elif az < 67.5:
        return "NE"
    elif az < 112.5:
        return "E"
    elif az < 157.5:
        return "SE"
    elif az < 202.5:
        return "S"
    elif az < 247.5:
        return "SW"
    elif az < 292.5:
        return "W"
    else:  # 292.5 <= az < 337.5
        return "NW"


def _parse_layers_string(layers_string):
    """Parse warmteverlies_lagen naar ThermalConstruction layers lijst.

    Args:
        layers_string: str "mat1 dikte | mat2 dikte | ..."

    Returns:
        list: [{"material": str, "thickness_mm": float, "type": str}, ...]
    """
    if not layers_string:
        return []

    layers = []
    parts = layers_string.split(" | ")

    for part in parts:
        part = part.strip()
        if not part:
            continue

        try:
            # Splits op spaties, laatste token is dikte, rest is material
            tokens = part.split()
            if len(tokens) < 2:
                continue

            thickness_str = tokens[-1]
            material_tokens = tokens[:-1]
            material = " ".join(material_tokens)

            thickness_mm = float(thickness_str)

            # Bepaal layer type
            layer_type = "solid"
            if material.lower():
                for kw in ["lucht", "spouw", "air gap", "air space", "cavity"]:
                    if kw in material.lower():
                        layer_type = "air_gap"
                        break

            layers.append({
                "material": material,
                "thickness_mm": thickness_mm,
                "type": layer_type
            })

        except (ValueError, IndexError):
            # Skip malformed layer entries
            continue

    return layers


def build_catalog_thermal_import(doc, rooms_data, exported_at=None):
    """Bouw ThermalImport JSON uit rooms_data + WV_BND DirectShapes.

    Args:
        doc: Revit Document
        rooms_data: list van room dicts (uit collect_rooms + map_all_rooms)
        exported_at: str timestamp of None (dan auto-generate)

    Returns:
        dict: ThermalImport JSON structure
    """
    if exported_at is None:
        try:
            from System import DateTime
            exported_at = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss")
        except Exception:
            exported_at = "unknown"

    # Project naam
    project_name = "Unknown"
    try:
        project_name = doc.Title or "Untitled"
    except Exception:
        pass

    # ========================================
    # 1. ROOMS verwerken
    # ========================================
    thermal_rooms = []
    room_eid_map = {}  # element_id -> "room-{idx}"
    room_label_map = {}  # "number name" -> "room-{idx}"

    # Pseudo-rooms
    thermal_rooms.append({
        "id": "outside",
        "name": "Buiten",
        "type": "outside"
    })

    has_ground = False  # Check later of ground nodig is

    for idx, room_data in enumerate(rooms_data):
        try:
            room_id = "room-{0}".format(idx + 1)

            # Room properties
            element_id = room_data.get("element_id", 0)
            name = room_data.get("name", "") or "Unnamed"
            number = room_data.get("number", "") or str(idx + 1)
            is_heated = room_data.get("is_heated", False)
            level_name = room_data.get("level_name", "") or "Unknown"
            area_m2 = room_data.get("floor_area_m2", 0.0)
            height_m = room_data.get("height_m", 0.0)
            volume_m3 = room_data.get("volume_m3", 0.0)

            # Fallback volume
            if volume_m3 <= 0.0:
                volume_m3 = area_m2 * height_m

            thermal_room = {
                "id": room_id,
                "revit_id": element_id,
                "name": name,
                "type": "heated" if is_heated else "unheated",
                "level": level_name,
                "area_m2": area_m2,
                "height_m": height_m,
                "volume_m3": volume_m3
            }
            thermal_rooms.append(thermal_room)

            # Maps
            room_eid_map[element_id] = room_id
            room_label = "{0} {1}".format(number, name).strip()
            room_label_map[room_label] = room_id

        except Exception:
            # Skip malformed room data
            continue

    # ========================================
    # 2. CONSTRUCTIONS uit WV_BND DirectShapes
    # ========================================
    constructions = []
    warnings = []
    construction_count = 0

    # Collect DirectShapes
    collector = (
        FilteredElementCollector(doc)
        .OfClass(DirectShape)
        .WhereElementIsNotElementType()
    )

    for ds in collector:
        try:
            # Check WV_BND prefix
            comment_param = ds.LookupParameter("Comments")
            if comment_param is None or not comment_param.HasValue:
                continue
            comment_value = comment_param.AsString()
            if not comment_value or not comment_value.startswith("WV_BND"):
                continue

            # Check orientation (alleen dichte constructies)
            orient_param = ds.LookupParameter("warmteverlies_orientatie")
            if orient_param is None or not orient_param.HasValue:
                continue
            orient = orient_param.AsString()
            if orient not in ("wand", "dak", "vloer"):
                continue

            # Check constructie naam
            const_param = ds.LookupParameter("warmteverlies_constructie")
            if const_param is None or not const_param.HasValue:
                continue
            constructie_naam = const_param.AsString()
            if not constructie_naam:
                continue

            # room_a (bron ruimte)
            ruimte_param = ds.LookupParameter("warmteverlies_ruimte")
            if ruimte_param is None or not ruimte_param.HasValue:
                continue
            ruimte_label = ruimte_param.AsString()

            room_a = room_label_map.get(ruimte_label)
            if room_a is None:
                warnings.append("Shape zonder geldige room_a: {0}".format(ruimte_label))
                continue

            # room_b (buurruimte)
            room_b = "outside"  # default
            try:
                naar_id_param = ds.LookupParameter("warmteverlies_naar_id")
                grens_param = ds.LookupParameter("warmteverlies_grenstype")

                if naar_id_param and naar_id_param.HasValue:
                    naar_id = int(naar_id_param.AsDouble())
                    if naar_id > 0:
                        room_b = room_eid_map.get(naar_id, "outside")
                    else:
                        # Bepaal uit grenstype
                        if grens_param and grens_param.HasValue:
                            grenstype = grens_param.AsString()
                            if grenstype == "ground":
                                room_b = "ground"
                                has_ground = True
                            else:
                                room_b = "outside"
            except Exception:
                pass

            # orientation mapping
            orientation_map = {"wand": "wall", "dak": "roof", "vloer": "floor"}
            orientation = orientation_map.get(orient, "wall")

            # compass
            compass = None
            try:
                azimut_param = ds.LookupParameter("warmteverlies_azimut")
                if azimut_param and azimut_param.HasValue:
                    azimut = azimut_param.AsDouble()
                    if azimut >= 0:
                        compass = _azimuth_to_compass(azimut % 360.0)
            except Exception:
                pass

            # area
            area_m2 = 0.0
            try:
                area_param = ds.LookupParameter("warmteverlies_oppervlak_m2")
                if area_param and area_param.HasValue:
                    area_m2 = area_param.AsDouble()
            except Exception:
                pass

            # layers
            layers = []
            try:
                lagen_param = ds.LookupParameter("warmteverlies_lagen")
                if lagen_param and lagen_param.HasValue:
                    lagen_str = lagen_param.AsString()
                    layers = _parse_layers_string(lagen_str)
            except Exception:
                pass

            # Construction object
            construction_count += 1
            construction = {
                "id": "con-{0}".format(construction_count),
                "room_a": room_a,
                "room_b": room_b,
                "orientation": orientation,
                "compass": compass,
                "gross_area_m2": area_m2,
                "revit_type_name": constructie_naam,
                "layers": layers
            }
            constructions.append(construction)

        except Exception:
            # Skip malformed shapes
            continue

    # Voeg ground pseudo-room toe indien nodig
    if has_ground:
        thermal_rooms.append({
            "id": "ground",
            "name": "Grond",
            "type": "ground"
        })

    # ========================================
    # 3. OPENINGS uit WV_BND DirectShapes
    # ========================================
    openings = []
    opening_count = 0

    # Collect orphan openings per room (voor volglas-fallback)
    orphan_openings_per_room = {}

    # Tweede pass over DirectShapes voor openingen
    for ds in collector:
        try:
            # Check WV_BND prefix
            comment_param = ds.LookupParameter("Comments")
            if comment_param is None or not comment_param.HasValue:
                continue
            comment_value = comment_param.AsString()
            if not comment_value or not comment_value.startswith("WV_BND"):
                continue

            # Check orientation (alleen openingen)
            orient_param = ds.LookupParameter("warmteverlies_orientatie")
            if orient_param is None or not orient_param.HasValue:
                continue
            orient = orient_param.AsString()
            if orient != "opening":
                continue

            # Check constructie naam/type
            const_param = ds.LookupParameter("warmteverlies_constructie")
            type_param = ds.LookupParameter("warmteverlies_type_stapel")

            constructie_naam = ""
            if const_param and const_param.HasValue:
                constructie_naam = const_param.AsString() or ""
            if not constructie_naam and type_param and type_param.HasValue:
                constructie_naam = type_param.AsString() or ""
            if not constructie_naam:
                continue

            # room_a
            ruimte_param = ds.LookupParameter("warmteverlies_ruimte")
            if ruimte_param is None or not ruimte_param.HasValue:
                continue
            ruimte_label = ruimte_param.AsString()

            room_a = room_label_map.get(ruimte_label)
            if room_a is None:
                warnings.append("Opening zonder geldige room_a: {0}".format(ruimte_label))
                continue

            # Geometrie uit bounding box
            bb = ds.get_BoundingBox(None)
            if bb is None:
                continue

            height_mm = (bb.Max.Z - bb.Min.Z) * 304.8
            width_mm = ((bb.Max.X - bb.Min.X)**2 + (bb.Max.Y - bb.Min.Y)**2)**0.5 * 304.8

            # Sill height (best-effort)
            sill_height_mm = None
            try:
                # Room level elevation
                room_level_elevation_m = None
                for room_data in rooms_data:
                    room_id = room_eid_map.get(room_data.get("element_id", 0))
                    if room_id == room_a:
                        room_level_elevation_m = room_data.get("level_elevation_m")
                        break

                if room_level_elevation_m is not None:
                    sill_height_mm = max(0.0, (bb.Min.Z * 0.3048 - room_level_elevation_m) * 1000.0)
            except Exception:
                pass

            # Azimuth en compass
            compass = None
            try:
                # Bepaal azimuth uit bounding box normaal
                dx = bb.Max.X - bb.Min.X
                dy = bb.Max.Y - bb.Min.Y
                dz = bb.Max.Z - bb.Min.Z

                # De dunne richting is waarschijnlijk de normaal
                if abs(dx) < abs(dy) and abs(dx) < abs(dz):
                    # Dun in X, normaal wijst in X richting
                    azimuth = 90.0 if dx >= 0 else 270.0
                elif abs(dy) < abs(dx) and abs(dy) < abs(dz):
                    # Dun in Y, normaal wijst in Y richting
                    azimuth = 0.0 if dy >= 0 else 180.0
                else:
                    # Fallback azimuth
                    azimuth = 0.0

                compass = _azimuth_to_compass(azimuth % 360.0)
            except Exception:
                pass

            # Type bepalen
            opening_type = "window"  # default
            constructie_lower = constructie_naam.lower()
            if "vlies" in constructie_lower or "curtain" in constructie_lower:
                opening_type = "curtain_wall"
            elif "deur" in constructie_lower or "door" in constructie_lower:
                opening_type = "door"

            opening_area_m2 = width_mm * height_mm / 1000000.0

            # Host-construction zoeken
            host_construction_id = None

            # Kandidaten: echte buitenwanden van zelfde room_a
            candidates = []
            for con in constructions:
                if (con["orientation"] == "wall" and
                    con["room_b"] == "outside" and
                    con["room_a"] == room_a and
                    len(con["layers"]) > 0):
                    candidates.append(con)

            # Kies beste kandidaat
            if candidates:
                # Zoek zelfde compass
                for con in candidates:
                    if con["compass"] == compass:
                        host_construction_id = con["id"]
                        # Corrigeer bruto area
                        con["gross_area_m2"] += opening_area_m2
                        break

                # Geen compass match -> grootste area
                if host_construction_id is None:
                    best_con = max(candidates, key=lambda c: c["gross_area_m2"])
                    host_construction_id = best_con["id"]
                    best_con["gross_area_m2"] += opening_area_m2

            # Geen host gevonden -> orphan opening
            if host_construction_id is None:
                if room_a not in orphan_openings_per_room:
                    orphan_openings_per_room[room_a] = []
                orphan_openings_per_room[room_a].append({
                    "constructie_naam": constructie_naam,
                    "opening_type": opening_type,
                    "width_mm": width_mm,
                    "height_mm": height_mm,
                    "sill_height_mm": sill_height_mm,
                    "compass": compass,
                    "area_m2": opening_area_m2
                })
                continue

            # Bouw opening
            opening_count += 1
            opening = {
                "id": "open-{0}".format(opening_count),
                "construction_id": host_construction_id,
                "type": opening_type,
                "width_mm": width_mm,
                "height_mm": height_mm,
                "revit_type_name": constructie_naam
            }

            if sill_height_mm is not None:
                opening["sill_height_mm"] = sill_height_mm

            openings.append(opening)

        except Exception:
            # Skip malformed opening shapes
            continue

    # Maak volglas-fallback constructions voor orphan openings
    fallback_constructions = 0
    for room_a, orphan_list in orphan_openings_per_room.items():
        if not orphan_list:
            continue

        # Bepaal fallback compass en totale area
        total_area_m2 = sum(o["area_m2"] for o in orphan_list)
        largest_opening = max(orphan_list, key=lambda o: o["area_m2"])
        fallback_compass = largest_opening["compass"]

        # Maak fallback construction
        construction_count += 1
        fallback_constructions += 1
        fallback_con = {
            "id": "con-{0}".format(construction_count),
            "room_a": room_a,
            "room_b": "outside",
            "orientation": "wall",
            "compass": fallback_compass,
            "gross_area_m2": total_area_m2 + 1.0,  # +1m² marge voor net > 0
            "revit_type_name": "Opening fallback",
            "layers": []
        }
        constructions.append(fallback_con)

        # Koppel alle orphan openings aan deze fallback
        for orphan in orphan_list:
            opening_count += 1
            opening = {
                "id": "open-{0}".format(opening_count),
                "construction_id": fallback_con["id"],
                "type": orphan["opening_type"],
                "width_mm": orphan["width_mm"],
                "height_mm": orphan["height_mm"],
                "revit_type_name": orphan["constructie_naam"]
            }

            if orphan["sill_height_mm"] is not None:
                opening["sill_height_mm"] = orphan["sill_height_mm"]

            openings.append(opening)

    # ========================================
    # 4. RESULT
    # ========================================
    result = {
        "version": "1.0",
        "source": "revit-catalog",
        "exported_at": exported_at,
        "project_name": project_name,
        "rooms": thermal_rooms,
        "constructions": constructions,
        "openings": openings,
        "open_connections": []
    }

    # Voeg debug info toe (wordt door pushbutton gestript)
    result["_debug"] = {
        "warnings": warnings,
        "fallback_constructions": fallback_constructions
    }

    return result


# =============================================================================
# Self-test (uitvoeren met: python catalog_export.py)
# =============================================================================
if __name__ == "__main__":
    # Test layer parsing
    print("=== LAYER PARSING TEST ===")
    test_layers = "hout_vuren_generiek 60 | n7_isolatie_PIR 120 | dampremmer 0"
    parsed = _parse_layers_string(test_layers)
    for layer in parsed:
        print("  {0}: {1}mm ({2})".format(layer["material"], layer["thickness_mm"], layer["type"]))

    # Test compass mapping
    print("\n=== COMPASS TEST ===")
    test_azimuths = [0, 45, 90, 135, 180, 225, 270, 315, 360]
    for az in test_azimuths:
        compass = _azimuth_to_compass(az % 360)
        print("  {0}° -> {1}".format(az, compass))

    print("\n=== TESTS COMPLETE ===")