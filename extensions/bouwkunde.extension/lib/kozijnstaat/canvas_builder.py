# -*- coding: utf-8 -*-
"""Canvas builder - auto-genereer een wall om kozijnen op te plaatsen.

Kozijnen zijn hosted families; ze moeten op een wall staan. Deze module
maakt optioneel een tijdelijke 'canvas' wall waarop de kozijnstaat wordt
opgebouwd, zodat de gebruiker niet zelf een wand hoeft voor te bereiden.
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Wall,
    WallType,
    Level,
    Line,
    XYZ,
    BuiltInParameter,
    ElementId,
    Transaction,
)

MM_TO_FT = 1.0 / 304.8


def find_level(doc, level_name=None):
    """Zoek een Level op naam, of retourneer het laagste level."""
    levels = list(
        FilteredElementCollector(doc)
        .OfClass(Level)
        .ToElements()
    )
    if not levels:
        return None

    if level_name:
        for lvl in levels:
            if lvl.Name == level_name:
                return lvl

    levels.sort(key=lambda l: l.Elevation)
    return levels[0]


def find_wall_type(doc, wall_type_name=None):
    """Zoek een WallType op naam, of retourneer de eerste gevonden type."""
    types = list(
        FilteredElementCollector(doc)
        .OfClass(WallType)
        .ToElements()
    )
    if not types:
        return None

    if wall_type_name:
        for wt in types:
            if wt.Name == wall_type_name:
                return wt
    return types[0]


def find_canvas_wall(doc, mark_value):
    """Zoek een bestaande canvas-wall op de 'Mark' parameter.

    Returns:
        Wall of None
    """
    walls = (
        FilteredElementCollector(doc)
        .OfClass(Wall)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    for w in walls:
        try:
            p = w.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if p and p.AsString() == mark_value:
                return w
        except Exception:
            continue
    return None


def create_canvas_wall(doc, width_mm, height_mm, origin_xyz,
                      wall_type, level, mark_value,
                      transaction_name="Create Kozijnstaat Canvas"):
    """Maak een nieuwe wall die dient als kozijnstaat-canvas.

    De wall loopt vanaf origin_xyz langs de X-as met opgegeven breedte
    en hoogte. Moet binnen een open Transaction aangeroepen worden, of
    wrapt een eigen Transaction als geen actieve.

    Args:
        doc: Revit Document
        width_mm: float, lengte van de wall in mm
        height_mm: float, hoogte in mm
        origin_xyz: XYZ start in feet
        wall_type: WallType
        level: Level waarop de wall staat
        mark_value: str voor Mark parameter (voor latere lookup)
        transaction_name: naam van de (eventueel eigen) Transaction

    Returns:
        Wall (de aangemaakte wall)
    """
    width_ft = width_mm * MM_TO_FT
    height_ft = height_mm * MM_TO_FT

    end = XYZ(origin_xyz.X + width_ft, origin_xyz.Y, origin_xyz.Z)
    curve = Line.CreateBound(origin_xyz, end)

    need_own_tx = not _transaction_active(doc)
    tx = None
    if need_own_tx:
        tx = Transaction(doc, transaction_name)
        tx.Start()

    try:
        wall = Wall.Create(
            doc,
            curve,
            wall_type.Id,
            level.Id,
            height_ft,
            0.0,      # offset
            False,    # flip
            False,    # structural
        )
        # Mark instellen voor latere lookup
        try:
            p = wall.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if p and not p.IsReadOnly:
                p.Set(mark_value)
        except Exception:
            pass

        if tx is not None:
            tx.Commit()
        return wall
    except Exception:
        if tx is not None and tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        raise


def _transaction_active(doc):
    """Check of er een actieve transactie is op het document."""
    try:
        return doc.IsModifiable
    except Exception:
        return False
