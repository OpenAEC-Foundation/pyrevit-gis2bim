# -*- coding: utf-8 -*-
"""Revit drawing layer: StairResult -> DetailLines + TextNotes in actieve view.

IronPython 2.7 / Revit API. Tekent in de actieve plan view in 2D.
Pivot van de StairResult ligt in (0,0) lokaal stelsel; deze module zet
treden om naar wereldcoordinaten via de meegegeven `origin` (XYZ in feet)
en `rotation` (radialen, CCW rond Z).
"""
from __future__ import division
import math

from Autodesk.Revit.DB import (
    XYZ, Line, Arc, Transaction, ViewType,
    DetailLine, DetailArc, ElementId,
)
try:
    from Autodesk.Revit.DB import TextNote, TextNoteOptions
    HAS_TEXTNOTE = True
except Exception:
    HAS_TEXTNOTE = False


M_TO_FT = 1.0 / 0.3048   # meters naar Revit-feet


def _to_world_xyz(p_local_m, origin_ft, rot_rad, z_ft=0.0):
    """Lokaal 2D-punt (in meters) -> wereld-XYZ in feet.

    Roteert eerst over rot_rad rond (0,0), schaalt naar feet, telt origin op.
    """
    c = math.cos(rot_rad)
    s = math.sin(rot_rad)
    x_local = p_local_m[0]
    y_local = p_local_m[1]
    x_rot = c * x_local - s * y_local
    y_rot = s * x_local + c * y_local
    return XYZ(
        origin_ft.X + x_rot * M_TO_FT,
        origin_ft.Y + y_rot * M_TO_FT,
        origin_ft.Z + z_ft,
    )


def _make_detail_line(doc, view, p1, p2):
    if p1.DistanceTo(p2) < 1e-6:
        return None
    geom_line = Line.CreateBound(p1, p2)
    return doc.Create.NewDetailCurve(view, geom_line)


def _make_detail_arc(doc, view, center, radius_ft, theta_start, theta_end):
    """Maakt een DetailArc in plan view (rond center, in XY-vlak)."""
    if radius_ft < 1e-6:
        return None
    arc = Arc.Create(
        center,
        radius_ft,
        theta_start,
        theta_end,
        XYZ.BasisX, XYZ.BasisY,
    )
    return doc.Create.NewDetailCurve(view, arc)


def draw_stair(doc, view, result, origin_ft, rotation_rad=0.0,
               draw_walkline=True, draw_numbers=True):
    """Plaats StairResult als detail lines in de gegeven view.

    Args:
        doc: Revit Document
        view: actieve plan/detail view (ViewPlan of detail)
        result: StairResult (uit methods.run)
        origin_ft: XYZ waar het lokale spilpunt belandt (in feet)
        rotation_rad: rotatie van het lokale stelsel rond Z
        draw_walkline: ook looplijn tekenen
        draw_numbers: trede-nummers als TextNotes plaatsen

    Returns:
        dict met ElementId-lijsten per categorie (info/debug)
    """
    placed = {"risers": [], "boundaries": [], "walkline": [], "labels": []}

    tx = Transaction(doc, "Trap tekenen ({0})".format(result.methode))
    tx.Start()
    try:
        for tr in result.treads:
            # Front-riser
            p1 = _to_world_xyz(tr.front[0], origin_ft, rotation_rad)
            p2 = _to_world_xyz(tr.front[1], origin_ft, rotation_rad)
            dc = _make_detail_line(doc, view, p1, p2)
            if dc is not None:
                placed["risers"].append(dc.Id)

            # Boundary lines (zij-randen): inner-edge en outer-edge als lijn
            b_in_1 = _to_world_xyz(tr.front[0], origin_ft, rotation_rad)
            b_in_2 = _to_world_xyz(tr.back[0], origin_ft, rotation_rad)
            b_out_1 = _to_world_xyz(tr.front[1], origin_ft, rotation_rad)
            b_out_2 = _to_world_xyz(tr.back[1], origin_ft, rotation_rad)
            for a, b in ((b_in_1, b_in_2), (b_out_1, b_out_2)):
                dc = _make_detail_line(doc, view, a, b)
                if dc is not None:
                    placed["boundaries"].append(dc.Id)

            # Harmonische arc-edges hebben voorrang voor verdreven treden
            if tr.inner_edge and tr.inner_edge[0] == "arc":
                _, center_local, radius_m, th_a, th_b = tr.inner_edge
                center_world = _to_world_xyz(center_local, origin_ft, rotation_rad)
                radius_ft = radius_m * M_TO_FT
                dc = _make_detail_arc(doc, view, center_world, radius_ft,
                                      th_a + rotation_rad, th_b + rotation_rad)
                if dc is not None:
                    placed["boundaries"].append(dc.Id)
            if tr.outer_edge and tr.outer_edge[0] == "arc":
                _, center_local, radius_m, th_a, th_b = tr.outer_edge
                center_world = _to_world_xyz(center_local, origin_ft, rotation_rad)
                radius_ft = radius_m * M_TO_FT
                dc = _make_detail_arc(doc, view, center_world, radius_ft,
                                      th_a + rotation_rad, th_b + rotation_rad)
                if dc is not None:
                    placed["boundaries"].append(dc.Id)

        # Looplijn
        if draw_walkline:
            for item in result.walkline:
                if item[0] == "line":
                    _, p1, p2 = item
                    w1 = _to_world_xyz(p1, origin_ft, rotation_rad)
                    w2 = _to_world_xyz(p2, origin_ft, rotation_rad)
                    dc = _make_detail_line(doc, view, w1, w2)
                    if dc is not None:
                        placed["walkline"].append(dc.Id)
                elif item[0] == "arc":
                    _, center_local, radius_m, th_a, th_b = item
                    center_world = _to_world_xyz(center_local, origin_ft, rotation_rad)
                    radius_ft = radius_m * M_TO_FT
                    dc = _make_detail_arc(doc, view, center_world, radius_ft,
                                          th_a + rotation_rad, th_b + rotation_rad)
                    if dc is not None:
                        placed["walkline"].append(dc.Id)

        # Tread-nummers
        if draw_numbers and HAS_TEXTNOTE:
            text_type_id = doc.GetDefaultElementTypeId(
                __import__("Autodesk.Revit.DB", fromlist=["ElementTypeGroup"]).ElementTypeGroup.TextNoteType)
            if text_type_id and text_type_id != ElementId.InvalidElementId:
                for tr in result.treads:
                    # Plaats label op centroide van trede
                    cx = (tr.front[0][0] + tr.front[1][0] + tr.back[0][0] + tr.back[1][0]) / 4.0
                    cy = (tr.front[0][1] + tr.front[1][1] + tr.back[0][1] + tr.back[1][1]) / 4.0
                    pos = _to_world_xyz((cx, cy), origin_ft, rotation_rad)
                    try:
                        tn = TextNote.Create(doc, view.Id, pos, str(tr.nummer), text_type_id)
                        placed["labels"].append(tn.Id)
                    except Exception:
                        pass

        tx.Commit()
    except Exception:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        raise
    return placed
