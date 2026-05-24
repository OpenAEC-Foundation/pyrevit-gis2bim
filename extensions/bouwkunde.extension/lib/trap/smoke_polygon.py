# -*- coding: utf-8 -*-
"""Smoke test voor polygon.py."""
from __future__ import division, print_function
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from polygon import (
    chain_segments, ensure_ccw, nearest_vertex_index,
    pick_axis_edge, edge_length_along_axis,
)


def test_rectangle():
    # Sparing 3m x 4m, hoeken op (0,0) (3,0) (3,4) (0,4)
    # Geef segmenten in willekeurige volgorde en orientatie:
    segs = [
        ((3, 0), (0, 0)),    # bottom, reversed
        ((0, 4), (0, 0)),    # left, reversed
        ((3, 4), (3, 0)),    # right, reversed
        ((0, 4), (3, 4)),    # top
    ]
    verts = chain_segments(segs)
    print("Verts:", verts)
    print("Signed area:", )
    verts = ensure_ccw(verts)
    print("CCW verts:", verts)

    # User klikt bij hoek (0,0) — vind spilpunt
    idx, d = nearest_vertex_index(verts, (0.02, 0.01))
    print("Spil-idx:", idx, "afstand:", d)
    spil = verts[idx]
    print("Spil:", spil)

    # User klikt langs inkomende as bij (0, 2) — middelpunt linker edge
    axis_in, width, axis_other = pick_axis_edge(verts, idx, (0.05, 2.0))
    print("axis_in (naar spil):", axis_in)
    print("width:", width)
    print("axis_other:", axis_other)

    # Lengte langs inkomende:
    L = edge_length_along_axis(verts, idx, axis_in)
    print("Lengte rechte deel beschikbaar:", L)


def test_l_shape():
    """L-vormige sparing voor onderkwart-trap.

         (0,5) +-----+ (2,5)
               |     |
               |     |
         (0,2) +-----+ (2,2)
                     |
                     |
         (0,0) ......+ (2,0)  ← maar de polygon loopt om:

    Eigenlijk: L-vorm met outer corner (0,0)(5,0)(5,2)(2,2)(2,5)(0,5)
    Concave hoek = (2,2).
    """
    from polygon import find_concave_vertices, suggest_pivot

    segs = [
        ((0, 0), (5, 0)),
        ((5, 0), (5, 2)),
        ((5, 2), (2, 2)),
        ((2, 2), (2, 5)),
        ((2, 5), (0, 5)),
        ((0, 5), (0, 0)),
    ]
    verts = chain_segments(segs)
    verts = ensure_ccw(verts)
    print("L-verts (CCW):", verts)

    concave = find_concave_vertices(verts)
    print("Concave indices:", concave, "->", [verts[i] for i in concave])

    # User klikt vaag bij de inspringende hoek
    idx, d, was_concave = suggest_pivot(verts, (2.1, 2.1))
    print("Suggested pivot idx={0} ({1}), afstand={2:.3f}, concave={3}".format(
        idx, verts[idx], d, was_concave))

    # Klik langs één van de aangrenzende edges (langs de korte hals naar boven)
    from polygon import pick_axis_edge, edge_length_along_axis, width_from_polygon
    axis_in, width_edge, other = pick_axis_edge(verts, idx, (2.05, 3.5))
    print("axis_in:", axis_in, "edge-width (oud):", width_edge, "other:", other)
    L = edge_length_along_axis(verts, idx, axis_in)
    print("Lengte rechte zone voor bocht:", L)
    w_short, w_long = width_from_polygon(verts, idx, axis_in, True)
    print("Ray-cast klemafstanden loodrecht: kort={0}  lang={1}".format(w_short, w_long))


if __name__ == "__main__":
    test_rectangle()
    print()
    print("=" * 50)
    test_l_shape()
