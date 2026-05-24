# -*- coding: utf-8 -*-
"""Pure-Python geometriekernel voor verdreven trappen (IronPython 2.7 compatibel).

CONVENTIE
=========
Lokaal stelsel: spilpunt (wand-binnenhoek) in oorsprong (0, 0).
De spec definieert TWEE wand-richtingen:

  in_wall_dir:   eenheidsvector langs INKOMENDE wand, vanaf spil weg
  out_wall_dir:  eenheidsvector langs UITGAANDE wand, vanaf spil weg

Deze twee staan loodrecht. Het polygon-interior ligt in het kwadrant
dat door beide vectoren positief gedefinieerd is. Voorbeeld:
  in_wall_dir = (+1, 0), out_wall_dir = (0, +1)  -> interior = eerste kwadrant.

WANDELAAR (in plan-zicht)
-------------------------
* Vóór bocht: loopt PARALLEL aan inkomende wand, op afstand
  inner_radius + walkline_offset (loodrecht naar interior = out_wall_dir),
  beweegt in -in_wall_dir richting (van wand-eind naar spil-hoek toe).
* Na bocht: loopt PARALLEL aan uitgaande wand, op afstand
  inner_radius + walkline_offset loodrecht naar interior (= in_wall_dir),
  beweegt in +out_wall_dir richting (van spil weg).

TREDEN
------
* Rechte loop VÓÓR bocht: aantrede langs in_wall_dir, breedte loodrecht
  (in out_wall_dir richting), van inner_radius tot inner_radius + width.
  Trede k=1 ligt het VERST van spil, trede k=n_straight_before direct
  bij de bocht (front-edge op spil-hoek-as).
* Verdreven zone: radiale risers vanaf spilpunt, theta loopt van
  atan2(in_wall_dir) naar atan2(out_wall_dir). Aantrede op LOOPLIJN
  = boog op r_walk / n_winders.
* Rechte loop NA bocht: aantrede langs out_wall_dir, breedte loodrecht
  (in in_wall_dir richting), idem.

TRED-NUMMERING (vanaf onderkant trap, wandelaar OMHOOG)
-------------------------------------------------------
  1 .. n_straight_before                   -> rechte VOOR bocht (onderaan)
  n_straight_before+1 .. +n_winders        -> verdreven
  +n_winders+1 .. +n_straight_after        -> rechte NA bocht (boven)
"""
from __future__ import division
import math


# ---------- 2D-vector helpers ----------

def XY(x, y):
    return (float(x), float(y))


def add(p, q):
    return (p[0] + q[0], p[1] + q[1])


def sub(p, q):
    return (p[0] - q[0], p[1] - q[1])


def scale(p, s):
    return (p[0] * s, p[1] * s)


def length(p):
    return math.sqrt(p[0] * p[0] + p[1] * p[1])


def normalize(v):
    L = length(v)
    if L < 1e-12:
        return (1.0, 0.0)
    return (v[0] / L, v[1] / L)


def cross_z(a, b):
    return a[0] * b[1] - a[1] * b[0]


def polar(r, theta):
    return (r * math.cos(theta), r * math.sin(theta))


def angle_of(v):
    return math.atan2(v[1], v[0])


def signed_angular_delta(a, b):
    """Shortest signed angle from a to b (in radialen, range (-pi, pi])."""
    d = b - a
    while d > math.pi:
        d -= 2 * math.pi
    while d <= -math.pi:
        d += 2 * math.pi
    return d


# ---------- Trede & resultaat-records ----------

class Tread(object):
    """Een trede in 2D-plattegrond.

    Velden:
      nummer:           1-based volgnummer (1 = laagst)
      front:            (p_in, p_out)  voorkant-riser (kant waar voet komt)
      back:             (p_in, p_out)  achterkant-riser
      is_winder:        True voor verdreven, False voor rechte trede
      tread_at_walkline: aantrede op looplijn [m]
      tread_at_narrow:  aantrede aan smalle kant [m]
      inner_edge:       optioneel: ("arc", center, r, theta_a, theta_b)
      outer_edge:       idem
    """
    def __init__(self, nummer, front, back, is_winder=False):
        self.nummer = nummer
        self.front = front
        self.back = back
        self.is_winder = is_winder
        self.tread_at_walkline = 0.0
        self.tread_at_narrow = 0.0
        self.inner_edge = None
        self.outer_edge = None

    def polygon(self):
        """4 hoekpunten CCW: front_in, front_out, back_out, back_in."""
        return [self.front[0], self.front[1], self.back[1], self.back[0]]


class StairResult(object):
    """Bundel van treden + looplijn-segments + diagnostiek per methode."""
    def __init__(self, methode):
        self.methode = methode
        self.treads = []
        self.walkline = []     # list van ("line", p1, p2) en/of ("arc", c, r, th_a, th_b)
        self.warnings = []
        self.params = {}


# ---------- Specificatie L-trap onderkwart ----------

class LStairSpec(object):
    """Input voor L-trap met onderkwart (of bovenkwart via n_straight_before).

    pivot:            (x, y) wand-binnenhoek (default (0, 0))
    in_wall_dir:      eenheidsvector langs inkomende wand vanaf spil
                      (NIET de wandelaar-richting; de WAND zelf)
    out_wall_dir:     eenheidsvector langs uitgaande wand vanaf spil
                      (loodrecht op in_wall_dir; cross_z != 0)
    width:            trapbreedte loodrecht op loop [m]
    n_winders:        aantal verdreven treden (typisch 3)
    n_straight_before: rechte treden VOOR bocht (onderaan voor onderkwart = 0)
    n_straight_after:  rechte treden NA bocht (bovenste loop)
    tread_straight:   rechte aantrede [m] (default 0.220)
    walkline_offset:  afstand looplijn tot binnenkant trap [m] (default 0.45)
    inner_radius:     spilpaal-straal [m] (0.0 = punt-spil)
    """
    def __init__(self,
                 pivot=(0.0, 0.0),
                 in_wall_dir=(1.0, 0.0),
                 out_wall_dir=(0.0, 1.0),
                 interior_in_dir=None,
                 interior_out_dir=None,
                 width=0.90,
                 n_winders=3,
                 n_straight_before=0,
                 n_straight_after=10,
                 tread_straight=0.220,
                 walkline_offset=0.45,
                 inner_radius=0.0):
        self.pivot = pivot
        self.in_wall_dir = normalize(in_wall_dir)
        self.out_wall_dir = normalize(out_wall_dir)
        self.width = float(width)
        self.n_winders = int(n_winders)
        self.n_straight_before = int(n_straight_before)
        self.n_straight_after = int(n_straight_after)
        self.tread_straight = float(tread_straight)
        self.walkline_offset = float(walkline_offset)
        self.inner_radius = float(inner_radius)

        # Sanity: wanden loodrecht?
        dot = self.in_wall_dir[0] * self.out_wall_dir[0] + \
              self.in_wall_dir[1] * self.out_wall_dir[1]
        if abs(dot) > 1e-3:
            raise ValueError(
                "in_wall_dir en out_wall_dir moeten loodrecht zijn "
                "(dot product = {0:.3f})".format(dot))

        # Interior-richtingen: default = aanname dat de andere wand-dir
        # de interior-kant aangeeft (geldt voor CONVEXE spil van CCW polygon).
        # Voor CONCAVE spil moeten ze expliciet doorgegeven worden.
        if interior_in_dir is None:
            self.interior_in_dir = self.out_wall_dir
        else:
            self.interior_in_dir = normalize(interior_in_dir)
        if interior_out_dir is None:
            self.interior_out_dir = self.in_wall_dir
        else:
            self.interior_out_dir = normalize(interior_out_dir)

    def is_ccw(self):
        """Trap-draairichting (math-CCW vanuit boven gezien).

        cross_z(in, out) > 0  =>  out_dir ligt 90 graden CCW van in_dir
                                  => wandelaar OMHOOG draait math-CCW
                                  => ego-perspectief: LINKSAF
        """
        return cross_z(self.in_wall_dir, self.out_wall_dir) > 0


# ---------- Bouwbesluit-validatie ----------

MIN_TREAD_NARROW = 0.050       # 50 mm (Bb art 2.30 woning)
MIN_TREAD_WALKLINE = 0.220     # 220 mm op looplijn


def validate_tread(t, spec, warnings):
    if t.is_winder and t.tread_at_narrow < MIN_TREAD_NARROW - 1e-6:
        warnings.append(
            "Trede {0}: aantrede smalle kant = {1:.0f} mm < {2:.0f} mm (Bb)".format(
                t.nummer, t.tread_at_narrow * 1000, MIN_TREAD_NARROW * 1000))
    if t.tread_at_walkline > 0 and t.tread_at_walkline < MIN_TREAD_WALKLINE - 1e-6:
        warnings.append(
            "Trede {0}: aantrede looplijn = {1:.0f} mm < {2:.0f} mm".format(
                t.nummer, t.tread_at_walkline * 1000, MIN_TREAD_WALKLINE * 1000))
