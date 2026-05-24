# -*- coding: utf-8 -*-
"""Standalone smoke test (CPython 3 of IronPython 2.7 compatible).

Run vanuit deze directory:
  python smoke_test.py

Dumpt vier methoden naar JSON en print een ASCII-overzicht.
"""
from __future__ import division, print_function
import json
import math
import os
import sys

# Zorg dat lokale imports werken vanuit deze map
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from geometry import LStairSpec
from methods import run, METHODS


def tread_to_dict(t):
    d = {
        "nr": t.nummer,
        "winder": t.is_winder,
        "front_in": list(t.front[0]),
        "front_out": list(t.front[1]),
        "back_in": list(t.back[0]),
        "back_out": list(t.back[1]),
        "tread_walkline_mm": round(t.tread_at_walkline * 1000.0, 1),
        "tread_narrow_mm": round(t.tread_at_narrow * 1000.0, 1),
    }
    if t.inner_edge:
        d["inner_edge"] = list(t.inner_edge)
    if t.outer_edge:
        d["outer_edge"] = list(t.outer_edge)
    return d


def result_to_dict(res):
    return {
        "methode": res.methode,
        "params": res.params,
        "treads": [tread_to_dict(t) for t in res.treads],
        "warnings": res.warnings,
    }


def print_summary(res):
    print("=" * 64)
    print("Methode: {0}".format(res.methode))
    print("-" * 64)
    print("{0:>4}  {1:>6}  {2:>10}  {3:>10}".format("nr", "winder", "loop[mm]", "smal[mm]"))
    for t in res.treads:
        print("{0:>4}  {1:>6}  {2:>10.1f}  {3:>10.1f}".format(
            t.nummer, "ja" if t.is_winder else "nee",
            t.tread_at_walkline * 1000.0, t.tread_at_narrow * 1000.0))
    if res.warnings:
        print("WAARSCHUWINGEN:")
        for w in res.warnings:
            print("  - " + w)
    else:
        print("(geen waarschuwingen)")


def main():
    spec = LStairSpec(
        pivot=(0.0, 0.0),
        axis_in=(0.0, -1.0),
        width=0.90,
        n_winders=3,
        n_straight_before=0,
        n_straight_after=10,
        tread_straight=0.220,
        walkline_offset=0.45,
        inner_radius=0.05,
        direction_ccw=True,
    )
    all_results = {}
    for name in ["looplijn", "proportioneel", "frans", "harmonisch"]:
        res = run(name, spec)
        print_summary(res)
        all_results[name] = result_to_dict(res)

    out_path = os.path.join(HERE, "smoke_test_output.json")
    with open(out_path, "w") as f:
        json.dump(all_results, f, indent=2)
    print()
    print("JSON dump: {0}".format(out_path))


if __name__ == "__main__":
    main()
