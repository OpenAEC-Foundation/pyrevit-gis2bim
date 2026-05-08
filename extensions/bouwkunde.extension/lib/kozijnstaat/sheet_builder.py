# -*- coding: utf-8 -*-
"""Sheet & view builder voor Kozijnstaat output.

Maakt een ViewSection (frontale elevatie) van een canvas-wand,
een ViewSheet met een gekozen titleblock, en plaatst de view als
Viewport op die sheet.

IronPython 2.7 compatible.
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilySymbol,
    BuiltInCategory,
    BuiltInParameter,
    ViewFamilyType,
    ViewFamily,
    ViewSection,
    ViewSheet,
    Viewport,
    BoundingBoxXYZ,
    Transform,
    XYZ,
)


def find_titleblocks(doc):
    """Alle TitleBlock FamilySymbols in het project.

    Returns:
        list[FamilySymbol]
    """
    return list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .OfClass(FamilySymbol)
        .WhereElementIsElementType()
        .ToElements()
    )


def find_section_view_type(doc):
    """Eerste ViewFamilyType met family = Section.

    Returns:
        ViewFamilyType of None
    """
    for vft in (
        FilteredElementCollector(doc)
        .OfClass(ViewFamilyType)
        .ToElements()
    ):
        try:
            if vft.ViewFamily == ViewFamily.Section:
                return vft
        except Exception:
            continue
    return None


def _wall_basis(wall):
    """Bereken (p_start, u_dir, normal_dir, length_ft, height_ft).

    u_dir is de horizontale richting langs de wand. normal_dir is
    de horizontale wand-normaal (loodrecht op u, in XY-vlak).
    """
    loc = wall.Location
    curve = loc.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    length = p0.DistanceTo(p1)
    if length <= 0.0:
        length = 1.0
    u = XYZ((p1.X - p0.X) / length, (p1.Y - p0.Y) / length, 0.0)
    # Normaal in XY-vlak: rechts-draai van u
    n = XYZ(u.Y, -u.X, 0.0)

    height = 10.0
    try:
        p = wall.get_Parameter(
            BuiltInParameter.WALL_USER_HEIGHT_PARAM
        )
        if p and p.HasValue:
            height = p.AsDouble()
    except Exception:
        pass

    return (p0, u, n, length, height)


def create_wall_elevation(doc, wall, view_type,
                          depth_ft=2.0, margin_ft=1.0):
    """Maak een ViewSection die de wand frontaal toont.

    Args:
        doc: Document
        wall: de canvas-wand
        view_type: ViewFamilyType (Section)
        depth_ft: section-diepte in feet (back-clip)
        margin_ft: extra marge rondom de wand in feet

    Returns:
        ViewSection
    """
    p0, u, n, length, height = _wall_basis(wall)

    mid = XYZ(
        p0.X + u.X * length / 2.0,
        p0.Y + u.Y * length / 2.0,
        p0.Z + height / 2.0,
    )

    t = Transform.Identity
    t.Origin = mid
    t.BasisX = u                        # rechts in view = langs wand
    t.BasisY = XYZ(0.0, 0.0, 1.0)       # boven in view = wereld-Z
    t.BasisZ = n                        # naar kijker = wandnormaal

    bbox = BoundingBoxXYZ()
    bbox.Transform = t
    half_w = length / 2.0 + margin_ft
    half_h = height / 2.0 + margin_ft
    bbox.Min = XYZ(-half_w, -half_h, -depth_ft)
    bbox.Max = XYZ(half_w, half_h, depth_ft)

    return ViewSection.CreateSection(doc, view_type.Id, bbox)


def set_view_scale(view, scale):
    """Stel View.Scale in (1:scale)."""
    try:
        view.Scale = int(scale)
    except Exception:
        pass


def set_view_name(view, name):
    """View hernoemen; faalt stil bij dubbele naam."""
    try:
        view.Name = name
    except Exception:
        pass


def create_sheet(doc, titleblock_symbol,
                 sheet_number=None, sheet_name=None):
    """Maak ViewSheet met opgegeven titleblock-symbol."""
    if not titleblock_symbol.IsActive:
        titleblock_symbol.Activate()
        doc.Regenerate()
    sheet = ViewSheet.Create(doc, titleblock_symbol.Id)
    if sheet_number:
        try:
            sheet.SheetNumber = sheet_number
        except Exception:
            pass
    if sheet_name:
        try:
            sheet.Name = sheet_name
        except Exception:
            pass
    return sheet


def get_titleblock_center(doc, sheet):
    """Midden (sheet-feet) van de titleblock-instance op de sheet."""
    titles = list(
        FilteredElementCollector(doc, sheet.Id)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .ToElements()
    )
    if not titles:
        return XYZ(0.0, 0.0, 0.0)
    bb = titles[0].get_BoundingBox(sheet)
    if bb is None:
        return XYZ(0.0, 0.0, 0.0)
    return XYZ(
        (bb.Min.X + bb.Max.X) / 2.0,
        (bb.Min.Y + bb.Max.Y) / 2.0,
        0.0,
    )


def place_view_on_sheet(doc, sheet, view, location_xyz=None):
    """Plaats view als Viewport op de sheet.

    Args:
        location_xyz: sheet-feet XYZ; None = midden van de titleblock
    """
    if location_xyz is None:
        location_xyz = get_titleblock_center(doc, sheet)
    if not Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
        return None
    return Viewport.Create(doc, sheet.Id, view.Id, location_xyz)
