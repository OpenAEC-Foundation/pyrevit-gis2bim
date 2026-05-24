# -*- coding: utf-8 -*-
"""Vier verdrijvingsmethoden voor L-trap onderkwart.

Alle methoden nemen een LStairSpec (nieuwe conventie: in_wall_dir +
out_wall_dir) en geven een StairResult terug. Coordinaten in lokaal
stelsel met spilpunt in (0, 0).
"""
from __future__ import division
import math

from geometry import (
    XY, add, sub, scale, length, normalize, polar, angle_of,
    signed_angular_delta, cross_z,
    Tread, StairResult, LStairSpec, validate_tread,
)


# =============================================================================
# Gedeelde helpers
# =============================================================================

def _straight_before_treads(spec):
    """Rechte treden VOOR de bocht (onderaan).

    Aantrede langs in_wall_dir, breedte langs interior_in_dir (loodrecht,
    naar polygon-interior).
    """
    treads = []
    n = spec.n_straight_before
    if n <= 0:
        return treads
    t = spec.tread_straight
    in_d = spec.in_wall_dir
    width_d = spec.interior_in_dir
    r_in = spec.inner_radius
    r_out = spec.inner_radius + spec.width
    for k in range(n):
        nr = k + 1
        back_d = (n - k) * t
        front_d = back_d - t
        back_anchor = scale(in_d, back_d)
        front_anchor = scale(in_d, front_d)
        inner_off = scale(width_d, r_in)
        outer_off = scale(width_d, r_out)
        tr = Tread(nr,
                   (add(front_anchor, inner_off), add(front_anchor, outer_off)),
                   (add(back_anchor, inner_off), add(back_anchor, outer_off)),
                   is_winder=False)
        tr.tread_at_walkline = t
        tr.tread_at_narrow = t
        treads.append(tr)
    return treads


def _straight_after_treads(spec, start_nummer):
    """Rechte treden NA de bocht (bovenste loop).

    Aantrede langs out_wall_dir, breedte langs interior_out_dir.
    """
    treads = []
    n = spec.n_straight_after
    if n <= 0:
        return treads
    t = spec.tread_straight
    out_d = spec.out_wall_dir
    width_d = spec.interior_out_dir
    r_in = spec.inner_radius
    r_out = spec.inner_radius + spec.width
    for j in range(n):
        nr = start_nummer + j + 1
        back_d = j * t
        front_d = back_d + t
        back_anchor = scale(out_d, back_d)
        front_anchor = scale(out_d, front_d)
        inner_off = scale(width_d, r_in)
        outer_off = scale(width_d, r_out)
        tr = Tread(nr,
                   (add(front_anchor, inner_off), add(front_anchor, outer_off)),
                   (add(back_anchor, inner_off), add(back_anchor, outer_off)),
                   is_winder=False)
        tr.tread_at_walkline = t
        tr.tread_at_narrow = t
        treads.append(tr)
    return treads


def _winder_thetas(spec):
    """Hoekgrenzen verdreven zone in lokaal stelsel.

    Verdreven zone vult de wedge tussen de interior-zijdes van de twee wanden.
    theta_start ligt op de interior_in_dir-straal (aansluiting met tread-voor-bocht),
    theta_end op de interior_out_dir-straal (aansluiting met tread-na-bocht).
    """
    th_start = angle_of(spec.interior_in_dir)
    th_end = angle_of(spec.interior_out_dir)
    delta = signed_angular_delta(th_start, th_end)
    return th_start, th_end, delta


def _walkline_curve(spec, th_start, th_end):
    """Bouw looplijn-curve in lokaal stelsel.

    Rechte deel vóór bocht: ligt op afstand r_walk LOODRECHT van inkomende
    wand (langs interior_in_dir), parallel aan in_wall_dir.
    Rechte deel na bocht: idem maar langs out_wall_dir met interior_out_dir.
    """
    items = []
    in_d = spec.in_wall_dir
    out_d = spec.out_wall_dir
    int_in = spec.interior_in_dir
    int_out = spec.interior_out_dir
    t = spec.tread_straight
    r_walk = spec.inner_radius + spec.walkline_offset

    if spec.n_straight_before > 0:
        p_far = add(scale(in_d, spec.n_straight_before * t),
                    scale(int_in, r_walk))
        p_arc_start = polar(r_walk, th_start)
        items.append(("line", p_far, p_arc_start))

    items.append(("arc", spec.pivot, r_walk, th_start, th_end))

    if spec.n_straight_after > 0:
        p_arc_end = polar(r_walk, th_end)
        p_far = add(scale(out_d, spec.n_straight_after * t),
                    scale(int_out, r_walk))
        items.append(("line", p_arc_end, p_far))
    return items


def _echo_params(spec):
    return {
        "width": spec.width,
        "n_winders": spec.n_winders,
        "n_straight_before": spec.n_straight_before,
        "n_straight_after": spec.n_straight_after,
        "tread_straight": spec.tread_straight,
        "walkline_offset": spec.walkline_offset,
        "inner_radius": spec.inner_radius,
        "in_wall_dir": spec.in_wall_dir,
        "out_wall_dir": spec.out_wall_dir,
        "is_ccw": spec.is_ccw(),
    }


# =============================================================================
# Methode 1: looplijn-methode
# =============================================================================

def method_walkline(spec):
    """Looplijn-methode: gelijke aantrede op looplijn.

    Risers radiaal vanaf spilpunt. Hoekverdeling gelijk over alle verdreven
    treden. Aantrede op looplijn = boog op r_walk / n_winders.
    """
    res = StairResult("looplijn")
    res.params = _echo_params(spec)

    res.treads.extend(_straight_before_treads(spec))

    th_start, th_end, delta = _winder_thetas(spec)
    n = spec.n_winders
    r_in = spec.inner_radius
    r_out = spec.inner_radius + spec.width
    r_walk = spec.inner_radius + spec.walkline_offset

    if n > 0:
        dtheta = delta / n
        arc_walk = abs(delta) * r_walk
        tread_walk = arc_walk / n
    else:
        dtheta = 0
        tread_walk = 0

    nr0 = spec.n_straight_before
    for k in range(n):
        nr = nr0 + k + 1
        th_back = th_start + k * dtheta
        th_front = th_start + (k + 1) * dtheta
        back_in = polar(r_in, th_back)
        back_out = polar(r_out, th_back)
        front_in = polar(r_in, th_front)
        front_out = polar(r_out, th_front)
        tr = Tread(nr, (front_in, front_out), (back_in, back_out), is_winder=True)
        tr.tread_at_walkline = tread_walk
        tr.tread_at_narrow = r_in * abs(dtheta) if r_in > 1e-6 else 0.0
        res.treads.append(tr)

    res.treads.extend(_straight_after_treads(spec, nr0 + n))
    res.walkline = _walkline_curve(spec, th_start, th_end)

    for tr in res.treads:
        validate_tread(tr, spec, res.warnings)
    return res


# =============================================================================
# Methode 2: proportioneel verdrijven
# =============================================================================

def method_proportional(spec, narrow_min=0.060):
    """Minimum aantrede aan smalle kant = narrow_min (default 60 mm).

    Effectieve binnen-radius wordt aangepast zodat boog/n_winders = narrow_min.
    Risers blijven radiaal vanaf spil (vanaf r_in_eff i.p.v. inner_radius).
    """
    res = StairResult("proportioneel")
    res.params = _echo_params(spec)
    res.params["narrow_min"] = narrow_min

    res.treads.extend(_straight_before_treads(spec))

    th_start, th_end, delta = _winder_thetas(spec)
    n = spec.n_winders
    if n > 0:
        # boog_inner = r_eff * |delta| = narrow_min * n
        r_in_eff = narrow_min * n / abs(delta) if abs(delta) > 1e-9 else spec.inner_radius
        r_in_eff = max(r_in_eff, spec.inner_radius)
    else:
        r_in_eff = spec.inner_radius
    r_out = spec.inner_radius + spec.width
    r_walk = spec.inner_radius + spec.walkline_offset
    dtheta = delta / n if n > 0 else 0

    nr0 = spec.n_straight_before
    for k in range(n):
        nr = nr0 + k + 1
        th_back = th_start + k * dtheta
        th_front = th_start + (k + 1) * dtheta
        back_in = polar(r_in_eff, th_back)
        back_out = polar(r_out, th_back)
        front_in = polar(r_in_eff, th_front)
        front_out = polar(r_out, th_front)
        tr = Tread(nr, (front_in, front_out), (back_in, back_out), is_winder=True)
        # Aantrede op looplijn: lineair interpoleren tussen narrow_min en r_out*|dtheta|
        if r_out > r_in_eff:
            frac = (r_walk - r_in_eff) / (r_out - r_in_eff)
            outer_tread = r_out * abs(dtheta)
            tr.tread_at_walkline = narrow_min + frac * (outer_tread - narrow_min)
        else:
            tr.tread_at_walkline = narrow_min
        tr.tread_at_narrow = narrow_min
        res.treads.append(tr)

    res.treads.extend(_straight_after_treads(spec, nr0 + n))
    res.walkline = _walkline_curve(spec, th_start, th_end)

    for tr in res.treads:
        validate_tread(tr, spec, res.warnings)
    return res


# =============================================================================
# Methode 3: Franse balanceermethode
# =============================================================================

def method_french(spec, overlap_frac=0.30):
    """Frans: verdrijving loopt door in de rechte zones (vloeiende overgang).

    De verdreven hoek wordt aan beide zijden iets verbreed (overlap_frac van
    één segment per zijde), waardoor aantrede op looplijn iets ruimer wordt
    en de overgang vloeiender oogt.
    """
    res = StairResult("frans")
    res.params = _echo_params(spec)
    res.params["overlap_frac"] = overlap_frac

    th_start, th_end, delta = _winder_thetas(spec)
    n = spec.n_winders
    if n > 0:
        seg = delta / n
        th_start_ext = th_start - seg * overlap_frac
        th_end_ext = th_end + seg * overlap_frac
        delta_ext = th_end_ext - th_start_ext
        dtheta = delta_ext / n
    else:
        th_start_ext = th_start
        th_end_ext = th_end
        dtheta = 0

    res.treads.extend(_straight_before_treads(spec))

    r_in = spec.inner_radius
    r_out = spec.inner_radius + spec.width
    r_walk = spec.inner_radius + spec.walkline_offset

    nr0 = spec.n_straight_before
    for k in range(n):
        nr = nr0 + k + 1
        th_back = th_start_ext + k * dtheta
        th_front = th_start_ext + (k + 1) * dtheta
        back_in = polar(r_in, th_back)
        back_out = polar(r_out, th_back)
        front_in = polar(r_in, th_front)
        front_out = polar(r_out, th_front)
        tr = Tread(nr, (front_in, front_out), (back_in, back_out), is_winder=True)
        tr.tread_at_walkline = r_walk * abs(dtheta)
        tr.tread_at_narrow = r_in * abs(dtheta) if r_in > 1e-6 else 0.0
        res.treads.append(tr)

    res.treads.extend(_straight_after_treads(spec, nr0 + n))
    res.walkline = _walkline_curve(spec, th_start_ext, th_end_ext)

    for tr in res.treads:
        validate_tread(tr, spec, res.warnings)
    return res


# =============================================================================
# Methode 4: harmonisch (cirkelboog-edges)
# =============================================================================

def method_harmonic(spec):
    """Looplijn met expliciete cirkelboog-binnen/buitenedges per verdreven trede.

    Risers blijven radiaal, maar inner_edge en outer_edge worden als arc
    opgeslagen voor mooie tekening (in plaats van rechte tred-zijkanten).
    """
    res = method_walkline(spec)
    res.methode = "harmonisch"
    th_start, th_end, delta = _winder_thetas(spec)
    n = spec.n_winders
    r_in = spec.inner_radius
    r_out = spec.inner_radius + spec.width
    dtheta = delta / n if n > 0 else 0
    nr0 = spec.n_straight_before
    for k in range(n):
        idx = nr0 + k
        if idx >= len(res.treads):
            break
        tr = res.treads[idx]
        if not tr.is_winder:
            continue
        th_back = th_start + k * dtheta
        th_front = th_start + (k + 1) * dtheta
        tr.inner_edge = ("arc", spec.pivot, r_in, th_back, th_front)
        tr.outer_edge = ("arc", spec.pivot, r_out, th_back, th_front)
    return res


# =============================================================================
# Registry
# =============================================================================

METHODS = {
    "looplijn": method_walkline,
    "proportioneel": method_proportional,
    "frans": method_french,
    "harmonisch": method_harmonic,
}


def run(method_name, spec):
    fn = METHODS.get(method_name)
    if fn is None:
        raise ValueError("Onbekende methode: {0}".format(method_name))
    return fn(spec)
