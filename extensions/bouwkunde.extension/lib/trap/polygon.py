# -*- coding: utf-8 -*-
"""Polygon-utilities voor sparing-input.

Pure-Python (geen Revit imports). Werkt op lijsten van (x, y)-tuples
en op lijsten van line-segments [(p1, p2), ...].

Verwacht segmenten in EEN gesloten ring, mogen in willekeurige volgorde
en orientatie zijn — chain() puzzelt ze aan elkaar.
"""
from __future__ import division
import math


TOL = 1e-4   # 0.1 mm tolerantie voor 'gelijke punten' (input in meters)


def _almost_eq(p, q, tol=TOL):
    return abs(p[0] - q[0]) < tol and abs(p[1] - q[1]) < tol


def chain_segments(segments, tol=TOL):
    """Maak van losse line-segments een gesloten polygon.

    Input: list van (p1, p2) tuples met p = (x, y).
    Output: list van vertices [(x, y), ...] in volgorde, GESLOTEN
            betekent dat verts[0] == verts[-1] NIET dubbel staat
            (laatste vertex is uniek; sluiting is impliciet).

    Raises ValueError als segmenten niet vormbaar tot gesloten ring.
    """
    if not segments:
        raise ValueError("Geen segmenten meegegeven")
    remaining = [list(s) for s in segments]
    chain = list(remaining.pop(0))   # [p_start, p_end]

    while remaining:
        end = chain[-1]
        found = False
        for i, seg in enumerate(remaining):
            if _almost_eq(seg[0], end, tol):
                chain.append(seg[1])
                remaining.pop(i)
                found = True
                break
            if _almost_eq(seg[1], end, tol):
                chain.append(seg[0])
                remaining.pop(i)
                found = True
                break
        if not found:
            raise ValueError(
                "Kan segmenten niet aan elkaar koppelen na vertex {0}; "
                "{1} segmenten over".format(end, len(remaining)))

    # Sluit-check: laatste vertex moet samenvallen met eerste
    if _almost_eq(chain[0], chain[-1], tol):
        chain.pop()  # remove duplicate
    else:
        raise ValueError("Polygon is niet gesloten: start={0}, end={1}".format(
            chain[0], chain[-1]))

    if len(chain) < 3:
        raise ValueError("Polygon heeft minder dan 3 vertices")

    return chain


def signed_area(verts):
    """Shoelace; positief = CCW."""
    a = 0.0
    n = len(verts)
    for i in range(n):
        x0, y0 = verts[i]
        x1, y1 = verts[(i + 1) % n]
        a += x0 * y1 - x1 * y0
    return 0.5 * a


def ensure_ccw(verts):
    """Geef vertices terug in CCW-volgorde."""
    if signed_area(verts) < 0:
        return list(reversed(verts))
    return list(verts)


def nearest_vertex_index(verts, point):
    """Index van de vertex die het dichtst bij point ligt + afstand."""
    best_i = 0
    best_d = float("inf")
    for i, v in enumerate(verts):
        d = math.hypot(v[0] - point[0], v[1] - point[1])
        if d < best_d:
            best_d = d
            best_i = i
    return best_i, best_d


def adjacent_vertices(verts, idx):
    """(prev, next) vertex tuples voor gegeven index in CCW-polygon."""
    n = len(verts)
    return verts[(idx - 1) % n], verts[(idx + 1) % n]


def edge_direction(p_from, p_to):
    """Genormaliseerde richting p_from -> p_to."""
    dx = p_to[0] - p_from[0]
    dy = p_to[1] - p_from[1]
    L = math.hypot(dx, dy)
    if L < 1e-12:
        return (0.0, 0.0), 0.0
    return (dx / L, dy / L), L


def pick_axis_edge(verts, spil_idx, ref_point):
    """Bepaal welke van de twee aangrenzende polygon-edges het 'inkomende' loopas is.

    De edge wiens midpoint het dichtst bij ref_point ligt wordt gekozen.

    Returns:
        axis_in: eenheidsvector NAAR de spil (van buurvertex naar spil)
        width:   lengte van de ANDERE aangrenzende edge (loodrecht op axis_in
                 in een rechthoek-aanname)
        other_vertex: vertex aan het andere uiteinde van de axis_in-edge
    """
    spil = verts[spil_idx]
    prev_v, next_v = adjacent_vertices(verts, spil_idx)

    mid_prev = ((spil[0] + prev_v[0]) / 2.0, (spil[1] + prev_v[1]) / 2.0)
    mid_next = ((spil[0] + next_v[0]) / 2.0, (spil[1] + next_v[1]) / 2.0)

    d_prev = math.hypot(ref_point[0] - mid_prev[0], ref_point[1] - mid_prev[1])
    d_next = math.hypot(ref_point[0] - mid_next[0], ref_point[1] - mid_next[1])

    if d_prev <= d_next:
        axis_other = prev_v
        width_other = next_v
    else:
        axis_other = next_v
        width_other = prev_v

    dir_to_spil, _len_axis = edge_direction(axis_other, spil)
    _, width = edge_direction(spil, width_other)
    return dir_to_spil, width, axis_other


def edge_length_along_axis(verts, spil_idx, axis_in):
    """Lengte van de polygon-edge bij spilpunt in de richting -axis_in.

    Wordt gebruikt om max aantal rechte treden VOOR de kwartdraai te schatten.
    """
    spil = verts[spil_idx]
    prev_v, next_v = adjacent_vertices(verts, spil_idx)
    for v in (prev_v, next_v):
        d, L = edge_direction(spil, v)
        # axis_in wijst NAAR spil, dus -axis_in wijst weg van spil langs inkomende
        if d[0] * (-axis_in[0]) + d[1] * (-axis_in[1]) > 0.9:
            return L
    return 0.0


# ---------- Concave (inspringende) hoeken — kandidaat-spilpunten ----------

def _cross_z(ax, ay, bx, by):
    return ax * by - ay * bx


def is_concave_vertex(verts, idx):
    """True als de hoek bij verts[idx] een inspringende hoek is in CCW-polygon.

    Voor een CCW-polygon: convex = links draaien (cross > 0),
    concave = rechts draaien (cross < 0).
    """
    n = len(verts)
    if n < 3:
        return False
    prev_v = verts[(idx - 1) % n]
    cur = verts[idx]
    next_v = verts[(idx + 1) % n]
    ax = cur[0] - prev_v[0]
    ay = cur[1] - prev_v[1]
    bx = next_v[0] - cur[0]
    by = next_v[1] - cur[1]
    return _cross_z(ax, ay, bx, by) < -1e-9


def find_concave_vertices(verts):
    """Indices van alle concave vertices in een CCW-polygon.

    Voor een L-sparing geeft dit meestal 1 hoek (de inspringende) terug.
    Voor U/Z-vormige sparingen kunnen er meerdere zijn.
    """
    return [i for i in range(len(verts)) if is_concave_vertex(verts, i)]


# ---------- Ray-cast tegen polygon ----------

def _segment_ray_intersection(p1, p2, origin, direction, tol=1e-9):
    """Snijpunt van halflijn (origin + t*direction, t>=0) met segment p1-p2.

    Returns t (afstand langs ray) of None.
    """
    sx = p2[0] - p1[0]
    sy = p2[1] - p1[1]
    rx, ry = direction
    denom = (-rx) * sy + sx * ry
    if abs(denom) < tol:
        return None
    qx = origin[0] - p1[0]
    qy = origin[1] - p1[1]
    s = ((-sy) * qx + sx * qy) / denom   # nee — herafleiden
    # Eenvoudiger: los stelsel op
    #  origin + t*direction = p1 + u*(p2 - p1),  t>=0,  0<=u<=1
    #  t*rx - u*sx = p1.x - origin.x
    #  t*ry - u*sy = p1.y - origin.y
    a = rx
    b = -sx
    c = ry
    d = -sy
    e = p1[0] - origin[0]
    f = p1[1] - origin[1]
    det = a * d - b * c
    if abs(det) < tol:
        return None
    t = (e * d - b * f) / det
    u = (a * f - e * c) / det
    if t < -tol:
        return None
    if u < -tol or u > 1.0 + tol:
        return None
    return t


def ray_to_polygon(verts, origin, direction, exclude_vertex_idx=None, tol=1e-6):
    """Kortste afstand vanaf origin in richting direction tot een polygon-edge.

    exclude_vertex_idx: optioneel — sla edges over die aan deze vertex hangen
                       (om triviale 0-hit bij eigen vertex te voorkomen).
    Returns: (afstand, hit_edge_index)  of  (None, None) als geen hit.
    """
    n = len(verts)
    best_t = None
    best_e = None
    for i in range(n):
        if exclude_vertex_idx is not None:
            if i == exclude_vertex_idx or (i + 1) % n == exclude_vertex_idx:
                continue
        p1 = verts[i]
        p2 = verts[(i + 1) % n]
        t = _segment_ray_intersection(p1, p2, origin, direction)
        if t is None or t < tol:
            continue
        if best_t is None or t < best_t:
            best_t = t
            best_e = i
    return best_t, best_e


def width_from_polygon(verts, spil_idx, axis_in, direction_ccw):
    """Bepaal trapbreedte als klemafstand vanaf spilpunt loodrecht op axis_in.

    Voor CCW direction_ccw=True: trap loopt na bocht in richting _perp(axis_in, ccw).
    De BREEDTE-richting (van trapbreedte-zijde) is dan loodrecht op axis_in
    aan de kant waar de inkomende rechte loop ligt.

    Voor rechthoekige sparing: beide loodrechte richtingen hitten dezelfde
    afstand (= trapbreedte) door eigen geometrie.
    Voor L-sparing (concave spil): kortste loodrechte klemafstand = trapbreedte.

    Returns:
        width_short, width_long: kleinste/grootste loodrechte klemafstand [m]
    """
    spil = verts[spil_idx]
    # Twee loodrechte richtingen op axis_in:
    perp_pos = (-axis_in[1], axis_in[0])   # CCW 90 graden
    perp_neg = (axis_in[1], -axis_in[0])   # CW 90 graden
    t_pos, _ = ray_to_polygon(verts, spil, perp_pos, exclude_vertex_idx=spil_idx)
    t_neg, _ = ray_to_polygon(verts, spil, perp_neg, exclude_vertex_idx=spil_idx)
    candidates = [t for t in (t_pos, t_neg) if t is not None]
    if not candidates:
        return 0.0, 0.0
    return min(candidates), max(candidates)


def pick_in_out_walls(verts, spil_idx, ref_point):
    """Bepaal in/out-wand richtingen + INTERIOR-zijde van elke wand.

    De edge waarvan midpoint het dichtst bij ref_point ligt = INKOMENDE wand.
    De andere aangrenzende edge = UITGAANDE wand.

    Interior-zijde wordt berekend uit polygon-CCW-loop-conventie: voor een
    CCW polygon ligt het interior LINKS van elke edge in de loop-richting.

    Voor de "next" wand (van spil naar volgende vertex in CCW): CCW-edge
    loopt vanaf spil → next, dus interior = perp_CCW(wand_dir).
    Voor de "prev" wand: CCW-edge loopt vanaf prev → spil, dus interior =
    perp_CCW(-wand_dir) = -perp_CCW(wand_dir) = perp_CW(wand_dir).

    Returns:
        in_other_vertex, out_other_vertex,
        in_wall_dir, out_wall_dir,
        interior_in_dir, interior_out_dir
    """
    spil = verts[spil_idx]
    prev_v, next_v = adjacent_vertices(verts, spil_idx)
    mid_prev = ((spil[0] + prev_v[0]) / 2.0, (spil[1] + prev_v[1]) / 2.0)
    mid_next = ((spil[0] + next_v[0]) / 2.0, (spil[1] + next_v[1]) / 2.0)
    d_prev = math.hypot(ref_point[0] - mid_prev[0], ref_point[1] - mid_prev[1])
    d_next = math.hypot(ref_point[0] - mid_next[0], ref_point[1] - mid_next[1])
    if d_prev <= d_next:
        in_other = prev_v
        in_is_prev = True
        out_other = next_v
        out_is_prev = False
    else:
        in_other = next_v
        in_is_prev = False
        out_other = prev_v
        out_is_prev = True

    def _udir(a, b):
        dx = b[0] - a[0]; dy = b[1] - a[1]
        L = math.hypot(dx, dy)
        if L < 1e-12:
            return (1.0, 0.0)
        return (dx / L, dy / L)

    in_dir = _udir(spil, in_other)
    out_dir = _udir(spil, out_other)

    # Interior-side van een wand:
    #   prev-wand  → CCW edge gaat van prev_v naar spil = -wand_dir;
    #              perp_CCW(-wand_dir) = (wand.y, -wand.x)
    #   next-wand  → CCW edge gaat van spil naar next_v = +wand_dir;
    #              perp_CCW(+wand_dir) = (-wand.y, wand.x)
    if in_is_prev:
        interior_in_dir = (in_dir[1], -in_dir[0])
    else:
        interior_in_dir = (-in_dir[1], in_dir[0])
    if out_is_prev:
        interior_out_dir = (out_dir[1], -out_dir[0])
    else:
        interior_out_dir = (-out_dir[1], out_dir[0])

    return (in_other, out_other, in_dir, out_dir,
            interior_in_dir, interior_out_dir)


def suggest_pivot(verts, click_xy):
    """Geef de meest waarschijnlijke spilpunt-vertex terug.

    Voorkeur:
      1. Concave vertex dichtst bij de klik (als die binnen 1 m ligt)
      2. Anders: dichtstbijzijnde vertex
    """
    concave = find_concave_vertices(verts)
    if concave:
        best_i = concave[0]
        best_d = float("inf")
        for i in concave:
            v = verts[i]
            d = math.hypot(v[0] - click_xy[0], v[1] - click_xy[1])
            if d < best_d:
                best_d = d
                best_i = i
        if best_d < 1.0:
            return best_i, best_d, True   # was_concave
    idx, d = nearest_vertex_index(verts, click_xy)
    return idx, d, False
