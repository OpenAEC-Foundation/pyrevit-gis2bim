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
                "thickness_mm": round(thickness_mm, 2),
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
            # Skip buitenruimten (terras, balkon, etc.)
            if room_data.get("is_outside"):
                continue

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
                "area_m2": round(area_m2, 2),
                "height_m": round(height_m, 2),
                "volume_m3": round(volume_m3, 2)
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
    # FASE 1: Maak per-face constructions (ruwe lijst)
    raw_constructions = []
    warnings = []
    raw_construction_count = 0

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

            # host_type voor wand-constructies
            host_type = None
            try:
                host_type_param = ds.LookupParameter("warmteverlies_host_type")
                if host_type_param and host_type_param.HasValue:
                    host_type = host_type_param.AsString()
            except Exception:
                pass

            # Raw construction object (per face)
            raw_construction_count += 1
            construction = {
                "room_a": room_a,
                "room_b": room_b,
                "orientation": orientation,
                "compass": compass,
                "gross_area_m2": round(area_m2, 2),
                "revit_type_name": constructie_naam,
                "layers": layers,
                "_host_type": host_type  # interne hint
            }
            raw_constructions.append(construction)

        except Exception:
            # Skip malformed shapes
            continue

    # FASE 2: Consolideer constructions per groep
    # Groepeer op (room_a, room_b, orientation, compass, revit_type_name)
    construction_groups = {}

    for raw_con in raw_constructions:
        group_key = (
            raw_con["room_a"],
            raw_con["room_b"],
            raw_con["orientation"],
            raw_con["compass"],
            raw_con["revit_type_name"]
        )

        if group_key not in construction_groups:
            construction_groups[group_key] = []
        construction_groups[group_key].append(raw_con)

    # Bouw geconsolideerde constructions lijst
    constructions = []
    construction_count = 0

    for group_key, group_constructions in construction_groups.items():
        construction_count += 1

        # Sommeer areas
        total_area = sum(con["gross_area_m2"] for con in group_constructions)

        # Neem eerste element voor gemeenschappelijke velden
        first_con = group_constructions[0]

        consolidated_con = {
            "id": "con-{0}".format(construction_count),
            "room_a": first_con["room_a"],
            "room_b": first_con["room_b"],
            "orientation": first_con["orientation"],
            "compass": first_con["compass"],
            "gross_area_m2": round(total_area, 2),
            "revit_type_name": first_con["revit_type_name"],
            "layers": first_con["layers"],  # Binnen zelfde type identiek
            "_host_type": first_con["_host_type"]  # interne hint
        }
        constructions.append(consolidated_con)

    # Voeg ground pseudo-room toe indien nodig
    if has_ground:
        thermal_rooms.append({
            "id": "ground",
            "name": "Grond",
            "type": "ground"
        })

    # Bouw lookup voor host_type -> exterior wall voor orphan fallback
    host_type_to_ext_wall = {}
    for con in constructions:
        if (con["room_b"] == "outside" and
            con["orientation"] == "wall" and
            con.get("_host_type")):
            host_type = con["_host_type"]
            if host_type not in host_type_to_ext_wall:
                host_type_to_ext_wall[host_type] = {
                    "revit_type_name": con["revit_type_name"],
                    "layers": con["layers"]
                }

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

            height_mm = round((bb.Max.Z - bb.Min.Z) * 304.8, 2)
            width_mm = round(((bb.Max.X - bb.Min.X)**2 + (bb.Max.Y - bb.Min.Y)**2)**0.5 * 304.8, 2)

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
                    sill_height_mm = round(max(0.0, (bb.Min.Z * 0.3048 - room_level_elevation_m) * 1000.0), 2)
            except Exception:
                pass

            # Lees echte params: azimut, area, host_type
            opening_azimut = None
            opening_area_m2 = 0.0
            host_type = None
            compass = None

            try:
                # warmteverlies_azimut
                azimut_param = ds.LookupParameter("warmteverlies_azimut")
                if azimut_param and azimut_param.HasValue:
                    opening_azimut = azimut_param.AsDouble()

                # warmteverlies_oppervlak_m2
                area_param = ds.LookupParameter("warmteverlies_oppervlak_m2")
                if area_param and area_param.HasValue:
                    opening_area_m2 = area_param.AsDouble()

                # warmteverlies_host_type
                host_param = ds.LookupParameter("warmteverlies_host_type")
                if host_param and host_param.HasValue:
                    host_type = host_param.AsString()
            except Exception:
                pass

            # Area fallback naar bbox als param ontbreekt/0
            if opening_area_m2 <= 0.0:
                opening_area_m2 = width_mm * height_mm / 1000000.0

            # Rond af
            opening_area_m2 = round(opening_area_m2, 2)

            # Compass bepalen
            if opening_azimut is not None and opening_azimut >= 0:
                compass = _azimuth_to_compass(opening_azimut % 360.0)

            # Type bepalen
            opening_type = "window"  # default
            constructie_lower = constructie_naam.lower()
            if "vlies" in constructie_lower or "curtain" in constructie_lower:
                opening_type = "curtain_wall"
            elif "deur" in constructie_lower or "door" in constructie_lower:
                opening_type = "door"

            # Host-construction zoeken met herziene matching
            host_construction_id = None

            # Kandidaten: echte buitenwanden van zelfde room_a
            candidates = []
            for con in constructions:
                if (con["orientation"] == "wall" and
                    con["room_b"] == "outside" and
                    con["room_a"] == room_a and
                    len(con["layers"]) > 0):
                    candidates.append(con)

            # Stap 1 — exacte koppeling
            if candidates:
                if compass is not None:
                    # Opening heeft geldige compass → match op compass
                    for con in candidates:
                        if con["compass"] == compass:
                            host_construction_id = con["id"]
                            con["gross_area_m2"] = round(con["gross_area_m2"] + opening_area_m2, 2)
                            break
                elif host_type:
                    # Opening azimut=-1 (deur) → match op host_type
                    matching_hosts = []
                    for con in candidates:
                        if con.get("_host_type") == host_type:
                            matching_hosts.append(con)

                    if matching_hosts:
                        # Meerdere host_type matches → kies grootste area
                        best_con = max(matching_hosts, key=lambda c: c["gross_area_m2"])
                        host_construction_id = best_con["id"]
                        best_con["gross_area_m2"] = round(best_con["gross_area_m2"] + opening_area_m2, 2)

                # Stap 2 — fallback binnen kandidaten
                if host_construction_id is None and candidates:
                    best_con = max(candidates, key=lambda c: c["gross_area_m2"])
                    host_construction_id = best_con["id"]
                    best_con["gross_area_m2"] = round(best_con["gross_area_m2"] + opening_area_m2, 2)

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
                    "compass": compass,  # dit is nu de echte azimut-compass
                    "area_m2": opening_area_m2,
                    "host_type": host_type
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

    # Maak fallback constructions voor orphan openings
    # Split per room tussen exterior doors (echte buitenwand) en overige (glas fallback)
    fallback_constructions = 0
    for room_a, orphan_list in orphan_openings_per_room.items():
        if not orphan_list:
            continue

        # Groepeer orphans per host_type matchbaarheid
        exterior_groups = {}  # host_type -> [orphans]
        other_orphans = []

        for orphan in orphan_list:
            host_type = orphan.get("host_type")
            if host_type and host_type in host_type_to_ext_wall:
                # Exterior door met bekende host_type
                if host_type not in exterior_groups:
                    exterior_groups[host_type] = []
                exterior_groups[host_type].append(orphan)
            else:
                # Overige orphans (vliesgevels, etc.)
                other_orphans.append(orphan)

        # Maak synthetische buitenwand-constructions voor exterior door groepen
        for host_type, ext_orphans in exterior_groups.items():
            total_area_m2 = sum(o["area_m2"] for o in ext_orphans)
            wall_info = host_type_to_ext_wall[host_type]

            construction_count += 1
            fallback_constructions += 1
            ext_wall_con = {
                "id": "con-{0}".format(construction_count),
                "room_a": room_a,
                "room_b": "outside",
                "orientation": "wall",
                "compass": None,  # deuren hebben azimut=-1
                "gross_area_m2": round(total_area_m2 + 1.0, 2),  # +1m² marge
                "revit_type_name": wall_info["revit_type_name"],
                "layers": wall_info["layers"]
            }
            constructions.append(ext_wall_con)

            # Koppel exterior door orphans aan deze buitenwand
            for orphan in ext_orphans:
                opening_count += 1
                opening = {
                    "id": "open-{0}".format(opening_count),
                    "construction_id": ext_wall_con["id"],
                    "type": orphan["opening_type"],
                    "width_mm": orphan["width_mm"],
                    "height_mm": orphan["height_mm"],
                    "revit_type_name": orphan["constructie_naam"]
                }

                if orphan["sill_height_mm"] is not None:
                    opening["sill_height_mm"] = orphan["sill_height_mm"]

                openings.append(opening)

        # Maak lege glas-fallback voor overige orphans
        if other_orphans:
            total_area_m2 = sum(o["area_m2"] for o in other_orphans)
            largest_opening = max(other_orphans, key=lambda o: o["area_m2"])
            fallback_compass = largest_opening["compass"]

            construction_count += 1
            fallback_constructions += 1
            fallback_con = {
                "id": "con-{0}".format(construction_count),
                "room_a": room_a,
                "room_b": "outside",
                "orientation": "wall",
                "compass": fallback_compass,
                "gross_area_m2": round(total_area_m2 + 1.0, 2),  # +1m² marge voor net > 0
                "revit_type_name": "Opening fallback",
                "layers": []
            }
            constructions.append(fallback_con)

            # Koppel overige orphan openings aan deze glas fallback
            for orphan in other_orphans:
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
    # 4. RESULT - strip _host_type hints
    # ========================================
    # Strip interne _host_type hints uit constructions voor output
    clean_constructions = []
    for con in constructions:
        clean_con = dict(con)  # shallow copy
        if "_host_type" in clean_con:
            del clean_con["_host_type"]
        clean_constructions.append(clean_con)

    result = {
        "version": "1.0",
        "source": "revit-raycast",
        "exported_at": exported_at,
        "project_name": project_name,
        "rooms": thermal_rooms,
        "constructions": clean_constructions,
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