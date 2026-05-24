# -*- coding: utf-8 -*-
"""Trap 2D — teken verdreven L-trap als detail lines in actieve plan view.

Nieuwe spec-conventie: spilpunt = wand-binnenhoek, in_wall_dir en
out_wall_dir worden afgeleid uit de twee aangrenzende polygon-edges.

Workflow:
  1. Pick methode.
  2. Pick sparing-curves (DetailLines / ModelCurves).
  3. Pick punt nabij spilpunt-hoek (concave hoek wordt voorgetrokken).
  4. Pick punt LANGS de inkomende wand (kiest welke edge = inkomend).
  5. Dialog: trapbreedte (polygon-suggestie) + treden-config.
  6. Plaatsing in actieve view + rapport.
"""
__title__ = "Trap 2D"
__author__ = "3BM Bouwkunde"
__doc__ = "Teken trap in plattegrond met diverse verdrijvingsmethoden"

import os
import sys
import math

from Autodesk.Revit.DB import XYZ, CurveElement, Arc
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms, script

HERE = os.path.dirname(__file__)
EXT_ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
TRAP_LIB = os.path.join(EXT_ROOT, "bouwkunde.extension", "lib", "trap")
if TRAP_LIB not in sys.path:
    sys.path.insert(0, TRAP_LIB)

from geometry import LStairSpec
from methods import run, METHODS
from draw import draw_stair, M_TO_FT
from polygon import (
    chain_segments, ensure_ccw, suggest_pivot,
    find_concave_vertices, pick_in_out_walls,
    width_from_polygon, edge_length_along_axis,
)

FT_TO_M = 0.3048

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ---------- Curve filter ----------

class CurveOnlyFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, CurveElement)

    def AllowReference(self, ref, point):
        return True


def _curve_to_segment_xy(curve):
    geom = curve.GeometryCurve if hasattr(curve, "GeometryCurve") else curve
    p1 = geom.GetEndPoint(0)
    p2 = geom.GetEndPoint(1)
    seg = (
        (p1.X * FT_TO_M, p1.Y * FT_TO_M),
        (p2.X * FT_TO_M, p2.Y * FT_TO_M),
    )
    is_arc = isinstance(geom, Arc)
    return seg, is_arc


def pick_sparing():
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            CurveOnlyFilter(),
            "Selecteer curves die de sparing vormen (gesloten polygon)")
    except OperationCanceledException:
        return None, None
    if not refs:
        return None, None
    segments = []
    had_arc = False
    for r in refs:
        c = doc.GetElement(r)
        seg, arc_flag = _curve_to_segment_xy(c)
        segments.append(seg)
        if arc_flag:
            had_arc = True
    return segments, had_arc


def pick_world_xy(prompt):
    try:
        p = uidoc.Selection.PickPoint(prompt)
    except OperationCanceledException:
        return None
    return (p.X * FT_TO_M, p.Y * FT_TO_M), p


def ask_method():
    return forms.SelectFromList.show(
        sorted(METHODS.keys()),
        title="Verdrijvingsmethode",
        multiselect=False,
    )


def ask_parameters(suggest_width_m, available_before_m, available_after_m):
    width_mm = int(round(suggest_width_m * 1000))
    avail_b = int(round(available_before_m * 1000))
    avail_a = int(round(available_after_m * 1000))
    defaults = "{0}, 3, 0, 13, 220, 450, 50".format(width_mm)
    prompt = (
        "Polygon-info:\n"
        "  - Trapbreedte (klemafstand loodrecht op inkomende wand): {0} mm\n"
        "  - Beschikbare lengte INKOMENDE wand: {1} mm\n"
        "  - Beschikbare lengte UITGAANDE wand: {2} mm\n\n"
        "Komma-gescheiden parameters:\n"
        "  trapbreedte[mm], n_winders, n_recht_voor, n_recht_na, "
        "aantrede[mm], looplijn-offset[mm], binnenstraal[mm]"
    ).format(width_mm, avail_b, avail_a)
    txt = forms.ask_for_string(default=defaults, prompt=prompt,
                                title="Trap parameters")
    if not txt:
        return None
    parts = [p.strip() for p in txt.split(",")]
    if len(parts) != 7:
        forms.alert("Verwacht 7 waarden, kreeg {0}".format(len(parts)),
                    title="Foute input")
        return None
    try:
        return {
            "width": float(parts[0]) / 1000.0,
            "n_winders": int(parts[1]),
            "n_straight_before": int(parts[2]),
            "n_straight_after": int(parts[3]),
            "tread_straight": float(parts[4]) / 1000.0,
            "walkline_offset": float(parts[5]) / 1000.0,
            "inner_radius": float(parts[6]) / 1000.0,
        }
    except ValueError as e:
        forms.alert("Kon parameters niet parsen: {0}".format(e),
                    title="Foute input")
        return None


def main():
    view = doc.ActiveView
    vt = view.ViewType.ToString()
    if vt not in ("FloorPlan", "CeilingPlan", "EngineeringPlan", "AreaPlan"):
        forms.alert("Open eerst een plan-view.", exitscript=True)

    method_name = ask_method()
    if not method_name:
        return

    # Sparing
    segments, had_arc = pick_sparing()
    if not segments:
        return
    if had_arc:
        forms.alert("Let op: gebogen segmenten worden als chord behandeld.",
                    title="MVP-beperking")
    try:
        verts = chain_segments(segments)
    except ValueError as e:
        forms.alert("Kon polygon niet vormen:\n{0}".format(e),
                    title="Sparing-fout", exitscript=True)
    verts = ensure_ccw(verts)
    output.print_md("**Polygon:** {0} vertices, {1} segmenten".format(
        len(verts), len(segments)))

    concave = find_concave_vertices(verts)
    if concave:
        output.print_md("**Concave hoeken:** " + ", ".join(
            ["v{0}=({1:.2f},{2:.2f})".format(i, verts[i][0], verts[i][1])
             for i in concave]))

    # Spilpunt
    pick = pick_world_xy("Klik nabij de SPILPUNT-hoek (concave hoek voorkeur)")
    if pick is None:
        return
    spil_xy, spil_xyz = pick
    spil_idx, d_spil, was_concave = suggest_pivot(verts, spil_xy)
    snapped_spil = verts[spil_idx]
    output.print_md("Spilpunt -> vertex {0} {1} ({2:.3f} m van klik)".format(
        spil_idx,
        "(concave, auto)" if was_concave else "(dichtstbijzijnde)",
        d_spil))

    # Inkomende wand kiezen
    pick = pick_world_xy(
        "Klik LANGS de INKOMENDE wand (de wand waarlangs de wandelaar "
        "VOOR de bocht loopt)")
    if pick is None:
        return
    in_click_xy, _ = pick
    (in_other, out_other, in_wall_dir, out_wall_dir,
     interior_in_dir, interior_out_dir) = pick_in_out_walls(
        verts, spil_idx, in_click_xy)
    output.print_md(
        "in_wall_dir = ({0:.2f}, {1:.2f}) naar v=({2:.2f},{3:.2f}) | "
        "interior naar ({4:.2f}, {5:.2f})".format(
            in_wall_dir[0], in_wall_dir[1], in_other[0], in_other[1],
            interior_in_dir[0], interior_in_dir[1]))
    output.print_md(
        "out_wall_dir = ({0:.2f}, {1:.2f}) naar v=({2:.2f},{3:.2f}) | "
        "interior naar ({4:.2f}, {5:.2f})".format(
            out_wall_dir[0], out_wall_dir[1], out_other[0], out_other[1],
            interior_out_dir[0], interior_out_dir[1]))

    # Beschikbare lengtes per wand
    Li = math.hypot(in_other[0] - snapped_spil[0],
                    in_other[1] - snapped_spil[1])
    Lo = math.hypot(out_other[0] - snapped_spil[0],
                    out_other[1] - snapped_spil[1])
    # Trapbreedte: klemafstand vanaf spil in out_wall_dir richting
    # (= loodrecht op inkomende wand, naar interior)
    w_short, w_long = width_from_polygon(verts, spil_idx, in_wall_dir, True)
    suggest_width = w_short if w_short and w_short > 0.05 else 0.9

    params = ask_parameters(suggest_width, Li, Lo)
    if params is None:
        return

    spec = LStairSpec(
        pivot=(0.0, 0.0),
        in_wall_dir=in_wall_dir,
        out_wall_dir=out_wall_dir,
        interior_in_dir=interior_in_dir,
        interior_out_dir=interior_out_dir,
        width=params["width"],
        n_winders=params["n_winders"],
        n_straight_before=params["n_straight_before"],
        n_straight_after=params["n_straight_after"],
        tread_straight=params["tread_straight"],
        walkline_offset=params["walkline_offset"],
        inner_radius=params["inner_radius"],
    )
    result = run(method_name, spec)

    # Spilpunt in feet voor plaatsing — alle treden zijn in wereld-XY-meters
    # relatief tot spil. Drawing-laag voegt origin toe (rotation_rad=0).
    origin_ft = XYZ(snapped_spil[0] / FT_TO_M,
                    snapped_spil[1] / FT_TO_M,
                    spil_xyz.Z)
    placed = draw_stair(
        doc, view, result,
        origin_ft=origin_ft,
        rotation_rad=0.0,
        draw_walkline=True,
        draw_numbers=True,
    )

    output.print_md("## Trap geplaatst — methode: **{0}**".format(result.methode))
    output.print_md(
        "- Risers: {0} | Boundaries: {1} | Looplijn: {2} | Labels: {3}".format(
            len(placed["risers"]), len(placed["boundaries"]),
            len(placed["walkline"]), len(placed["labels"])))
    if result.warnings:
        output.print_md("### Waarschuwingen")
        for w in result.warnings:
            output.print_md("- " + w)
    else:
        output.print_md("Geen normwaarschuwingen.")

    output.print_md("### Trede-overzicht")
    rows = []
    for tr in result.treads:
        rows.append([
            tr.nummer,
            "verdreven" if tr.is_winder else "recht",
            "{0:.0f}".format(tr.tread_at_walkline * 1000),
            "{0:.0f}".format(tr.tread_at_narrow * 1000),
        ])
    output.print_table(
        table_data=rows,
        columns=["Nr", "Type", "Aantrede looplijn [mm]",
                 "Aantrede smalle kant [mm]"])


if __name__ == "__main__":
    main()
