# -*- coding: utf-8 -*-
"""Kozijnstaat - Create.

Plaatst alle unieke kozijn-FamilyTypes op een grid (rows x cols) op een
bestaande wand of een automatisch gegenereerde canvas-wall. Plaatst
tevens een tag onder elk kozijn.

Vervangt 31_3BM_Kozijnstaat_create.dyn.
IronPython 2.7.
"""

__title__ = "Create\nKozijnstaat"
__author__ = "3BM Bouwkunde"
__doc__ = "Plaats unieke kozijntypes op grid op een wand (of auto-canvas)"

import os
import sys

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    Transaction,
    XYZ,
    Wall,
    ElementId,
    BuiltInParameter,
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from kozijnstaat.config import load_config
from kozijnstaat.family_collector import (
    collect_window_symbols,
    get_symbol_width_mm,
    get_symbol_height_mm,
)
from kozijnstaat.grid_layout import (
    compute_grid_points,
    compute_tag_points,
    estimate_canvas_size_mm,
)
from kozijnstaat.canvas_builder import (
    find_level,
    find_wall_type,
    find_canvas_wall,
    create_canvas_wall,
)
from kozijnstaat.tag_placer import place_tag

MM_TO_FT = 1.0 / 304.8


def _pick_wall(uidoc):
    """Laat user een wand kiezen; None = geannuleerd."""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, "Selecteer host-wand voor kozijnstaat"
        )
        return uidoc.Document.GetElement(ref.ElementId)
    except OperationCanceledException:
        return None


def _get_wall_geometry(wall):
    """Haal origin (XYZ) + horizontale richtingsvector + lengte/hoogte.

    Returns:
        tuple (origin_ft, u_dir, u_length_ft, v_length_ft)
    """
    loc = wall.Location
    curve = loc.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    length = p0.DistanceTo(p1)

    # Horizontale richting langs de wand
    u_dir = XYZ(
        (p1.X - p0.X) / length,
        (p1.Y - p0.Y) / length,
        0.0,
    )

    # Hoogte uit wall type / parameter
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        height_ft = p.AsDouble() if p and p.HasValue else 10.0
    except Exception:
        height_ft = 10.0

    # Origin = linker-onder-hoek
    origin = XYZ(p0.X, p0.Y, p0.Z)

    return origin, u_dir, length, height_ft


def run():
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()
    output.print_md("## Kozijnstaat - Create")

    cfg = load_config()
    kozijn_family = cfg.get("kozijn_family", "3BM_kozijn")
    rows = int(cfg.get("grid_rows", 6))
    cols = int(cfg.get("grid_cols", 8))
    tag_offset = float(cfg.get("tag_offset_mm", -1000.0)) * MM_TO_FT

    # 1. Verzamel unieke types
    symbols = collect_window_symbols(
        doc, name_contains=kozijn_family
    )
    if not symbols:
        forms.alert(
            "Geen kozijn FamilyTypes gevonden met filter '{0}'."
            .format(kozijn_family),
            title="Geen types",
        )
        return

    output.print_md(
        "Gevonden: **{0}** unieke kozijntypes".format(len(symbols))
    )

    # 2. Host-wand kiezen of genereren
    choice = forms.alert(
        "Host-wand voor kozijnstaat?\n\n"
        "Ja = bestaande wand selecteren\n"
        "Nee = auto-canvas wall genereren",
        yes=True, no=True, cancel=True,
    )
    if choice is None:
        output.print_md("*Geannuleerd.*")
        return

    host_wall = None
    created_canvas = False

    if choice:
        host_wall = _pick_wall(uidoc)
        if host_wall is None:
            output.print_md("*Geen wand geselecteerd.*")
            return
    else:
        # Auto-canvas wall
        widths = [get_symbol_width_mm(s) or 1000.0 for s in symbols]
        heights = [get_symbol_height_mm(s) or 2000.0 for s in symbols]
        w_mm, h_mm = estimate_canvas_size_mm(
            widths, heights, cols, rows, padding_mm=500.0
        )

        level = find_level(doc, cfg.get("canvas_wall_level"))
        wt = find_wall_type(doc, cfg.get("canvas_wall_type"))
        if level is None or wt is None:
            forms.alert(
                "Kan geen Level of WallType vinden voor auto-canvas.",
                title="Fout",
            )
            return

        mark = cfg.get("canvas_wall_name", "3BM_Kozijnstaat_Canvas")
        existing = find_canvas_wall(doc, mark)
        if existing:
            host_wall = existing
            output.print_md(
                "Bestaande canvas-wall hergebruikt (Mark='{0}').".format(mark)
            )
        else:
            origin = XYZ(0.0, 0.0, 0.0)
            tx = Transaction(doc, "Create Kozijnstaat Canvas Wall")
            tx.Start()
            try:
                host_wall = create_canvas_wall(
                    doc, w_mm, h_mm, origin, wt, level, mark
                )
                tx.Commit()
                created_canvas = True
                output.print_md(
                    "Canvas-wall aangemaakt: **{0} x {1} mm**"
                    .format(int(w_mm), int(h_mm))
                )
            except Exception as ex:
                if tx.HasStarted() and not tx.HasEnded():
                    tx.RollBack()
                forms.alert(
                    "Canvas-wall aanmaken mislukt:\n{0}".format(ex),
                    title="Fout",
                )
                return

    # 3. Bereken grid
    origin, u_dir, u_length, v_length = _get_wall_geometry(host_wall)
    v_dir = XYZ(0.0, 0.0, 1.0)  # verticaal langs wall

    points = compute_grid_points(
        origin, u_dir, v_dir, u_length, v_length,
        cols, rows, len(symbols),
    )
    tag_points = compute_tag_points(points, tag_offset)

    if not points:
        forms.alert(
            "Grid berekening faalde (0 punten).",
            title="Fout",
        )
        return

    # 4. Plaats kozijnen + tags
    tx = Transaction(doc, "Kozijnstaat - Plaats kozijnen")
    tx.Start()
    try:
        placed = 0
        failed = 0

        for i, symbol in enumerate(symbols):
            if i >= len(points):
                break

            if not symbol.IsActive:
                symbol.Activate()
                doc.Regenerate()

            try:
                doc.Create.NewFamilyInstance(
                    points[i], symbol, host_wall,
                    StructuralType.NonStructural,
                )
                placed += 1
            except Exception as ex:
                output.print_md(
                    "  - *FOUT type '{0}': {1}*"
                    .format(symbol.Name, ex)
                )
                failed += 1

        tx.Commit()
        output.print_md(
            "Geplaatst: **{0}**, Mislukt: **{1}**"
            .format(placed, failed)
        )

        # Samenvatting
        msg_canvas = ""
        if created_canvas:
            msg_canvas = "\nCanvas-wall: '{0}'".format(
                cfg.get("canvas_wall_name", "3BM_Kozijnstaat_Canvas")
            )
        forms.alert(
            "Klaar.\n{0} kozijnen geplaatst, {1} mislukt.{2}"
            .format(placed, failed, msg_canvas),
            title="Create Kozijnstaat",
        )
    except Exception as ex:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        forms.alert("Fout:\n{0}".format(ex), title="Fout")


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run()
