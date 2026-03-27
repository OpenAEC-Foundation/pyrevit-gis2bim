# -*- coding: utf-8 -*-
"""
Grid Generatie en Score Berekening - GebiedsAnalyse
====================================================

Genereert een grid van analysepunten en berekent per punt
een voorzieningenscore op basis van afstand tot POIs.

Score logica:
    - Per categorie telt alleen de dichtstbijzijnde POI
    - Ring score = max_score * (1 - ring_index / aantal_ringen)
    - Buiten alle ringen = 0 punten
    - Totaalscore = som van alle categorie-scores

Gebruik:
    from gis2bim.analysis.grid import generate_grid, calculate_scores
"""

import math


def haversine(lat1, lon1, lat2, lon2):
    """Afstand in meters tussen twee WGS84 punten.

    Args:
        lat1, lon1: Eerste punt in decimale graden
        lat2, lon2: Tweede punt in decimale graden

    Returns:
        Afstand in meters
    """
    R = 6371000
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = (math.sin(dphi / 2) ** 2 +
         math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2) ** 2)
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def generate_grid(center_lat, center_lon, size_m, resolution_m):
    """Genereer een grid van analysepunten in WGS84.

    Maakt een vierkant grid gecentreerd op het opgegeven punt.
    Punten worden als WGS84 (lat, lon) teruggegeven.

    Args:
        center_lat: Breedtegraad van het centrum
        center_lon: Lengtegraad van het centrum
        size_m: Grootte van het grid in meters (breedte en hoogte)
        resolution_m: Afstand tussen gridpunten in meters

    Returns:
        Lijst van dicts: [{"lat": float, "lon": float, "row": int, "col": int}, ...]
        Tuple (n_rows, n_cols): Aantal rijen en kolommen
    """
    half = size_m / 2.0
    n_steps = int(size_m / resolution_m)
    if n_steps < 1:
        n_steps = 1

    # Benaderde omrekening meters -> graden
    # 1 graad breedtegraad ~ 111320 m
    # 1 graad lengtegraad ~ 111320 * cos(lat) m
    lat_per_m = 1.0 / 111320.0
    lon_per_m = 1.0 / (111320.0 * math.cos(math.radians(center_lat)))

    points = []
    n_rows = n_steps + 1
    n_cols = n_steps + 1

    for row in range(n_rows):
        for col in range(n_cols):
            dy = -half + row * resolution_m
            dx = -half + col * resolution_m
            lat = center_lat + dy * lat_per_m
            lon = center_lon + dx * lon_per_m
            points.append({
                "lat": lat,
                "lon": lon,
                "row": row,
                "col": col,
            })

    return points, (n_rows, n_cols)


def _score_for_category(grid_lat, grid_lon, pois, category):
    """Bereken de score van een gridpunt voor een categorie.

    Alleen de dichtstbijzijnde POI telt.

    Args:
        grid_lat, grid_lon: Gridpunt coordinaten
        pois: Lijst van POI dicts met 'lat' en 'lon'
        category: Categorie dict met 'rings' en 'max_score'

    Returns:
        Score (float, 0.0 tot max_score)
    """
    rings = category["rings"]
    max_score = category["max_score"]
    max_ring = rings[-1]

    # Vind dichtstbijzijnde POI
    min_dist = None
    for poi in pois:
        dist = haversine(grid_lat, grid_lon, poi["lat"], poi["lon"])
        if min_dist is None or dist < min_dist:
            min_dist = dist

    if min_dist is None or min_dist > max_ring:
        return 0.0

    # Bepaal in welke ring de dichtstbijzijnde POI valt
    n_rings = len(rings)
    for ring_index, ring_dist in enumerate(rings):
        if min_dist <= ring_dist:
            # Score neemt af met ring_index
            score = max_score * (1.0 - float(ring_index) / float(n_rings))
            return score

    return 0.0


def calculate_scores(grid_points, pois_per_category, categories):
    """Bereken scores voor alle gridpunten.

    Args:
        grid_points: Lijst van grid dicts met 'lat', 'lon', 'row', 'col'
        pois_per_category: Dict van categorie-id -> lijst van POI dicts
        categories: Dict van categorie-id -> categorie definitie

    Returns:
        Lijst van floats (totaalscore per gridpunt, zelfde volgorde als grid_points)
    """
    scores = []

    # Pre-filter: bepaal max ring per categorie voor snelle afwijzing
    enabled_cats = {}
    for cat_id, cat in categories.items():
        if cat.get("enabled", True) and cat_id in pois_per_category:
            pois = pois_per_category[cat_id]
            if pois:
                enabled_cats[cat_id] = {
                    "category": cat,
                    "pois": pois,
                    "max_ring": max(cat["rings"]),
                }

    for point in grid_points:
        total = 0.0
        plat = point["lat"]
        plon = point["lon"]

        for cat_id, cat_data in enabled_cats.items():
            # Pre-filter: skip POIs die ver weg zijn
            nearby_pois = []
            max_ring = cat_data["max_ring"]
            for poi in cat_data["pois"]:
                # Snelle breedtegraad-check (~111km per graad)
                dlat = abs(poi["lat"] - plat)
                if dlat > max_ring / 111000.0:
                    continue
                nearby_pois.append(poi)

            if nearby_pois:
                total += _score_for_category(
                    plat, plon, nearby_pois, cat_data["category"]
                )

        scores.append(total)

    return scores


def smooth_scores(scores, n_rows, n_cols):
    """Simpele 3x3 Gaussian smoothing op het scoregrid.

    Args:
        scores: Lijst van floats (lengte = n_rows * n_cols)
        n_rows: Aantal rijen
        n_cols: Aantal kolommen

    Returns:
        Nieuwe lijst van smoothed scores
    """
    # 3x3 Gaussian kernel (genormaliseerd)
    kernel = [
        [1, 2, 1],
        [2, 4, 2],
        [1, 2, 1],
    ]
    k_sum = 16.0

    result = list(scores)

    for row in range(n_rows):
        for col in range(n_cols):
            weighted_sum = 0.0
            weight_total = 0.0
            for kr in range(-1, 2):
                for kc in range(-1, 2):
                    r = row + kr
                    c = col + kc
                    if 0 <= r < n_rows and 0 <= c < n_cols:
                        w = kernel[kr + 1][kc + 1]
                        idx = r * n_cols + c
                        weighted_sum += scores[idx] * w
                        weight_total += w

            idx = row * n_cols + col
            if weight_total > 0:
                result[idx] = weighted_sum / weight_total

    return result
