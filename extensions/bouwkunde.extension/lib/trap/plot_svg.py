# -*- coding: utf-8 -*-
"""SVG-dump van StairResult voor visuele debug zonder Revit."""
from __future__ import division, print_function
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from geometry import LStairSpec
from methods import run


SCALE = 90
PAD = 50


def _bbox(result):
    xs, ys = [0.0], [0.0]   # spil meenemen
    for t in result.treads:
        for p in (t.front[0], t.front[1], t.back[0], t.back[1]):
            xs.append(p[0])
            ys.append(p[1])
    return min(xs), min(ys), max(xs), max(ys)


def to_svg(result, path):
    xmin, ymin, xmax, ymax = _bbox(result)
    pad_m = 0.2
    xmin -= pad_m; ymin -= pad_m; xmax += pad_m; ymax += pad_m
    w = (xmax - xmin) * SCALE + 2 * PAD
    h = (ymax - ymin) * SCALE + 2 * PAD

    def x2(x):
        return PAD + (x - xmin) * SCALE

    def y2(y):
        # Flip y zodat noord boven ligt
        return h - PAD - (y - ymin) * SCALE

    out = []
    out.append('<?xml version="1.0"?>')
    out.append(
        '<svg xmlns="http://www.w3.org/2000/svg" width="{0}" height="{1}" '
        'viewBox="0 0 {0} {1}">'.format(int(w), int(h)))
    out.append('<rect width="100%" height="100%" fill="white"/>')

    # Hulplijntjes: x-as en y-as (in lokaal stelsel) door spil
    out.append('<line x1="{0}" y1="{1}" x2="{2}" y2="{1}" '
               'stroke="#ddd" stroke-width="1"/>'.format(
                   x2(xmin), y2(0), x2(xmax)))
    out.append('<line x1="{0}" y1="{1}" x2="{0}" y2="{2}" '
               'stroke="#ddd" stroke-width="1"/>'.format(
                   x2(0), y2(ymin), y2(ymax)))

    # Spil
    cx = x2(0)
    cy = y2(0)
    out.append('<circle cx="{0}" cy="{1}" r="5" fill="red"/>'.format(cx, cy))
    out.append('<text x="{0}" y="{1}" font-size="11" fill="red">spil</text>'.format(
        cx + 8, cy - 8))

    # Treden
    for t in result.treads:
        color = "#0066cc" if t.is_winder else "#444"
        fill = "rgba(0,102,204,0.18)" if t.is_winder else "rgba(80,80,80,0.08)"
        poly_d = "M{0},{1} L{2},{3} L{4},{5} L{6},{7} Z".format(
            x2(t.front[0][0]), y2(t.front[0][1]),
            x2(t.front[1][0]), y2(t.front[1][1]),
            x2(t.back[1][0]), y2(t.back[1][1]),
            x2(t.back[0][0]), y2(t.back[0][1]))
        out.append('<path d="{0}" fill="{1}" stroke="{2}" stroke-width="1"/>'.format(
            poly_d, fill, color))
        # Front-riser dikker
        out.append(
            '<line x1="{0}" y1="{1}" x2="{2}" y2="{3}" stroke="{4}" '
            'stroke-width="2.5"/>'.format(
                x2(t.front[0][0]), y2(t.front[0][1]),
                x2(t.front[1][0]), y2(t.front[1][1]),
                color))
        # Nummer in midden
        cx_t = (t.front[0][0] + t.front[1][0] + t.back[0][0] + t.back[1][0]) / 4.0
        cy_t = (t.front[0][1] + t.front[1][1] + t.back[0][1] + t.back[1][1]) / 4.0
        out.append(
            '<text x="{0}" y="{1}" font-size="13" text-anchor="middle" '
            'font-weight="bold" fill="black">{2}</text>'.format(
                x2(cx_t), y2(cy_t) + 4, t.nummer))

    # Looplijn (groen, gestreept)
    for item in result.walkline:
        if item[0] == "line":
            _, p1, p2 = item
            out.append(
                '<line x1="{0}" y1="{1}" x2="{2}" y2="{3}" stroke="green" '
                'stroke-width="1.6" stroke-dasharray="6,3"/>'.format(
                    x2(p1[0]), y2(p1[1]), x2(p2[0]), y2(p2[1])))
        elif item[0] == "arc":
            _, center, radius, th_a, th_b = item
            sp = (center[0] + radius * math.cos(th_a),
                  center[1] + radius * math.sin(th_a))
            ep = (center[0] + radius * math.cos(th_b),
                  center[1] + radius * math.sin(th_b))
            delta = th_b - th_a
            large = 1 if abs(delta) > math.pi else 0
            # SVG y-flip dus sweep omgekeerd
            sweep = 0 if delta > 0 else 1
            d = "M{0},{1} A{2},{3} 0 {4},{5} {6},{7}".format(
                x2(sp[0]), y2(sp[1]),
                radius * SCALE, radius * SCALE,
                large, sweep,
                x2(ep[0]), y2(ep[1]))
            out.append('<path d="{0}" stroke="green" stroke-width="1.6" '
                       'fill="none" stroke-dasharray="6,3"/>'.format(d))

    # Titel
    out.append('<text x="{0}" y="22" font-size="15" font-weight="bold">'
               '{1}  (spec: in={2}, out={3})</text>'.format(
                   PAD, result.methode,
                   result.params.get("in_wall_dir"),
                   result.params.get("out_wall_dir")))

    out.append('</svg>')
    with open(path, "w") as f:
        f.write("\n".join(out))


def main():
    # Test-case: in_wall = +x, out_wall = +y, dus interior in eerste kwadrant.
    # CCW (linksaf voor wandelaar omhoog).
    spec = LStairSpec(
        in_wall_dir=(1.0, 0.0),
        out_wall_dir=(0.0, 1.0),
        width=0.90,
        n_winders=3,
        n_straight_before=2,
        n_straight_after=10,
        tread_straight=0.220,
        walkline_offset=0.45,
        inner_radius=0.05,
    )
    for name in ("looplijn", "proportioneel", "frans", "harmonisch"):
        res = run(name, spec)
        path = os.path.join(HERE, "plot_{0}.svg".format(name))
        to_svg(res, path)
        print("Geschreven:", path,
              "({0} treden, {1} warns)".format(len(res.treads), len(res.warnings)))


if __name__ == "__main__":
    main()
