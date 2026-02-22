# -*- coding: utf-8 -*-
"""
Revit Sheet Utilities - GIS2BIM
===============================

Gedeelde functies voor het werken met Revit sheets:
- Sheet bounds detectie (titelblok of sheet outline)
- Grid layout berekening (images in rijen/kolommen)
- Image plaatsing op sheets
- Label plaatsing bij images

Gebruik:
    from gis2bim.revit.sheets import (
        get_sheet_bounds,
        calculate_grid_layout,
        calculate_grid_position,
        place_image_on_sheet,
        place_label_on_sheet,
    )
"""

from pyrevit import DB

MM_TO_FEET = 1.0 / 304.8


def get_sheet_bounds(doc, sheet, log=None):
    """Bepaal de beschikbare ruimte op een sheet.

    Zoekt het titelblok op de sheet en gebruikt diens bounding box.
    Fallback: sheet.Outline.
    Laatste fallback: A3 standaard (420x297mm).

    Args:
        doc: Revit Document
        sheet: ViewSheet element
        log: Optionele log functie

    Returns:
        Tuple (x_min, y_min, x_max, y_max) in feet
    """
    if log is None:
        log = lambda msg: None

    # Probeer titelblok bounding box
    try:
        collector = DB.FilteredElementCollector(doc, sheet.Id)
        titleblocks = collector.OfCategory(
            DB.BuiltInCategory.OST_TitleBlocks
        ).WhereElementIsNotElementType().ToElements()

        if titleblocks:
            tb = titleblocks[0]
            bb = tb.get_BoundingBox(sheet)
            if bb:
                bounds = (bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y)
                w_mm = (bounds[2] - bounds[0]) / MM_TO_FEET
                h_mm = (bounds[3] - bounds[1]) / MM_TO_FEET
                log("Titelblok bounds: {0:.0f} x {1:.0f} mm".format(w_mm, h_mm))
                return bounds
    except Exception as e:
        log("Titelblok bounds fout: {0}".format(e))

    # Fallback: sheet outline
    try:
        outline = sheet.Outline
        bounds = (outline.Min.U, outline.Min.V, outline.Max.U, outline.Max.V)
        w_mm = (bounds[2] - bounds[0]) / MM_TO_FEET
        h_mm = (bounds[3] - bounds[1]) / MM_TO_FEET
        log("Sheet outline: {0:.0f} x {1:.0f} mm".format(w_mm, h_mm))
        return bounds
    except Exception as e:
        log("Sheet outline fout: {0}".format(e))

    # Laatste fallback: A3 standaard
    log("Fallback naar A3 standaard afmetingen")
    return (0, 0, 420.0 * MM_TO_FEET, 297.0 * MM_TO_FEET)


def calculate_grid_layout(sheet_bounds, cols=3, rows=2,
                          gap_h_mm=12.0, gap_v_mm=20.0,
                          inset_mm=20.0, log=None):
    """Bereken optimale grid layout voor images op een sheet.

    Berekent de maximale image grootte (vierkant) die past in het
    beschikbare gebied, en centreert het grid.

    Args:
        sheet_bounds: Tuple (x_min, y_min, x_max, y_max) in feet
        cols: Aantal kolommen (default 3)
        rows: Aantal rijen (default 2)
        gap_h_mm: Horizontale ruimte tussen images in mm
        gap_v_mm: Verticale ruimte tussen rijen in mm (incl. label)
        inset_mm: Marge binnen de sheet rand in mm
        log: Optionele log functie

    Returns:
        Dict met:
            x_start: X start positie (linkerrand grid) in feet
            y_start: Y start positie (bovenrand grid) in feet
            img_size: Image grootte (breedte=hoogte) in feet
            cols: Aantal kolommen
            rows: Aantal rijen
            gap_h: Horizontale gap in feet
            gap_v: Verticale gap in feet
    """
    if log is None:
        log = lambda msg: None

    x_min, y_min, x_max, y_max = sheet_bounds

    inset_ft = inset_mm * MM_TO_FEET
    gap_h_ft = gap_h_mm * MM_TO_FEET
    gap_v_ft = gap_v_mm * MM_TO_FEET

    # Beschikbare ruimte met inset
    avail_x_min = x_min + inset_ft
    avail_y_min = y_min + inset_ft
    avail_x_max = x_max - inset_ft
    avail_y_max = y_max - inset_ft

    avail_w = avail_x_max - avail_x_min
    avail_h = avail_y_max - avail_y_min

    # Bereken maximale image grootte (vierkant) die in grid past
    max_img_w = (avail_w - (cols - 1) * gap_h_ft) / cols
    max_img_h = (avail_h - (rows - 1) * gap_v_ft) / rows
    img_size = min(max_img_w, max_img_h)

    # Centreer het grid in de beschikbare ruimte
    content_w = cols * img_size + (cols - 1) * gap_h_ft
    content_h = rows * img_size + (rows - 1) * gap_v_ft
    margin_x = (avail_w - content_w) / 2.0
    margin_y = (avail_h - content_h) / 2.0

    img_mm = img_size / MM_TO_FEET
    log("Grid layout: image={0:.0f}mm, beschikbaar={1:.0f}x{2:.0f}mm".format(
        img_mm, avail_w / MM_TO_FEET, avail_h / MM_TO_FEET))

    return {
        "x_start": avail_x_min + margin_x,
        "y_start": avail_y_max - margin_y,
        "img_size": img_size,
        "cols": cols,
        "rows": rows,
        "gap_h": gap_h_ft,
        "gap_v": gap_v_ft,
    }


def calculate_grid_position(slot_idx, layout):
    """Bereken de sheet-coordinaten voor een grid positie.

    Grid layout (bijv. 3 kolommen x 2 rijen):
      slot 0  slot 1  slot 2
      slot 3  slot 4  slot 5

    Args:
        slot_idx: Positie index (0-gebaseerd)
        layout: Dict van calculate_grid_layout()

    Returns:
        Tuple (center_x, center_y) in feet op de sheet
    """
    cols = layout["cols"]
    col = slot_idx % cols
    row = slot_idx // cols
    img = layout["img_size"]
    gap_h = layout["gap_h"]
    gap_v = layout["gap_v"]

    # X: start + halve image + (image + gap) * kolom
    center_x = layout["x_start"] + (img / 2.0) + col * (img + gap_h)

    # Y: start (boven) - halve image - (image + gap) * rij
    center_y = layout["y_start"] - (img / 2.0) - row * (img + gap_v)

    return (center_x, center_y)


def place_image_on_sheet(doc, sheet, image_path, center_x, center_y,
                         img_size, log=None):
    """Plaats een image op een sheet op de opgegeven positie.

    Args:
        doc: Revit Document
        sheet: ViewSheet element
        image_path: Pad naar het afbeeldingsbestand
        center_x: X center positie op sheet in feet
        center_y: Y center positie op sheet in feet
        img_size: Gewenste breedte/hoogte in feet
        log: Optionele log functie

    Returns:
        Het aangemaakte ImageInstance element
    """
    if log is None:
        log = lambda msg: None

    # 1. Maak ImageType aan
    options = DB.ImageTypeOptions(image_path, False, DB.ImageTypeSource.Import)
    image_type = DB.ImageType.Create(doc, options)

    # 2. Placement opties
    placement = DB.ImagePlacementOptions()
    placement.PlacementPoint = DB.BoxPlacement.Center
    placement.Location = DB.XYZ(center_x, center_y, 0)

    # 3. Maak ImageInstance aan
    image_instance = DB.ImageInstance.Create(
        doc, sheet, image_type.Id, placement
    )

    # 4. Schaal instellen
    current_width = image_instance.Width
    if current_width > 0:
        image_instance.Width = img_size
        log("Image geschaald: {0:.4f}ft (was {1:.4f}ft)".format(
            img_size, current_width))

    return image_instance


def place_label_on_sheet(doc, sheet, label_text, center_x, center_y,
                         img_size, label_offset_mm=3.0,
                         text_type_id=None, log=None):
    """Plaats een tekstlabel onder een image op een sheet.

    Args:
        doc: Revit Document
        sheet: ViewSheet element
        label_text: Tekst voor het label
        center_x: X center van de image in feet
        center_y: Y center van de image in feet
        img_size: Image grootte in feet (voor positie berekening)
        label_offset_mm: Afstand onder de image in mm (default 3)
        text_type_id: Optioneel TextNoteType Id.
                      Als None wordt een geschikt type gezocht.
        log: Optionele log functie

    Returns:
        Het aangemaakte TextNote element, of None bij fout
    """
    if log is None:
        log = lambda msg: None

    try:
        if text_type_id is None:
            text_type_id = _find_small_text_note_type(doc)

        if text_type_id is None:
            log("Geen TextNoteType gevonden, label overgeslagen")
            return None

        label_offset_ft = label_offset_mm * MM_TO_FEET
        label_y = center_y - (img_size / 2.0) - label_offset_ft
        point = DB.XYZ(center_x, label_y, 0)

        text_note_options = DB.TextNoteOptions()
        text_note_options.TypeId = text_type_id
        text_note_options.HorizontalAlignment = DB.HorizontalTextAlignment.Center

        text_note = DB.TextNote.Create(
            doc, sheet.Id, point, label_text, text_note_options
        )

        log("Label geplaatst: {0}".format(label_text))
        return text_note

    except Exception as e:
        log("Label fout: {0}".format(e))
        return None


def _find_small_text_note_type(doc):
    """Zoek een geschikt klein TextNoteType (<=4mm).

    Returns:
        ElementId van het TextNoteType, of None
    """
    collector = DB.FilteredElementCollector(doc)
    text_types = collector.OfClass(DB.TextNoteType).ToElements()

    for text_type in text_types:
        try:
            height_param = text_type.get_Parameter(
                DB.BuiltInParameter.TEXT_SIZE
            )
            if height_param:
                height_ft = height_param.AsDouble()
                height_mm = height_ft / MM_TO_FEET
                if height_mm <= 4.0:
                    return text_type.Id
        except Exception:
            pass

    # Fallback: eerste beschikbare
    if text_types:
        return text_types[0].Id

    return None


def find_a3_titleblock(doc, log=None):
    """Zoek een A3 titelblok FamilySymbol.

    Args:
        doc: Revit Document
        log: Optionele log functie

    Returns:
        ElementId van het titelblok, of None
    """
    if log is None:
        log = lambda msg: None

    collector = DB.FilteredElementCollector(doc)
    titleblocks = collector.OfCategory(
        DB.BuiltInCategory.OST_TitleBlocks
    ).WhereElementIsElementType().ToElements()

    for tb in titleblocks:
        name = tb.FamilyName.lower() if tb.FamilyName else ""
        type_name = ""
        try:
            type_name = DB.Element.Name.GetValue(tb).lower()
        except Exception:
            pass

        combined = name + " " + type_name
        if "a3" in combined or "420" in combined:
            log("A3 titelblok gevonden: {0}".format(tb.FamilyName))
            return tb.Id

    return None


def find_any_titleblock(doc, log=None):
    """Zoek het eerste beschikbare titelblok.

    Args:
        doc: Revit Document
        log: Optionele log functie

    Returns:
        ElementId van het titelblok, of None
    """
    if log is None:
        log = lambda msg: None

    collector = DB.FilteredElementCollector(doc)
    titleblocks = collector.OfCategory(
        DB.BuiltInCategory.OST_TitleBlocks
    ).WhereElementIsElementType().ToElements()

    for tb in titleblocks:
        log("Fallback titelblok: {0}".format(tb.FamilyName))
        return tb.Id

    return None


def populate_sheets_dropdown(combo, doc, default_sheet_number=None, log=None):
    """Vul een ComboBox met beschikbare sheets.

    Args:
        combo: WPF ComboBox element
        doc: Revit Document
        default_sheet_number: Sheetnummer om te pre-selecteren (bijv. "091")
        log: Optionele log functie
    """
    if log is None:
        log = lambda msg: None

    try:
        from System.Windows.Controls import ComboBoxItem

        collector = DB.FilteredElementCollector(doc)
        sheets = collector.OfClass(DB.ViewSheet).ToElements()
        default_item = None

        for sheet in sorted(sheets, key=lambda s: s.SheetNumber):
            item = ComboBoxItem()
            item.Content = "{0} - {1}".format(sheet.SheetNumber, sheet.Name)
            item.Tag = sheet.Id
            combo.Items.Add(item)

            if default_sheet_number and sheet.SheetNumber == default_sheet_number:
                default_item = item

        if default_item is not None:
            combo.SelectedItem = default_item
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0

    except Exception as e:
        log("Error populating sheets: {0}".format(e))
