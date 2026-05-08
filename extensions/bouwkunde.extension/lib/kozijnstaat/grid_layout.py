# -*- coding: utf-8 -*-
"""Grid layout berekening voor kozijn-plaatsing op een surface.

Vervangt de Surface.PointAtParameter + urev/vrev logica uit
31_3BM_Kozijnstaat_create.dyn.
"""

from Autodesk.Revit.DB import XYZ

MM_TO_FT = 1.0 / 304.8

# Standaard ISO papierformaten in landscape (mm)
PAPER_SIZES_MM = {
    "A0": (1189.0, 841.0),
    "A1": (841.0, 594.0),
    "A2": (594.0, 420.0),
    "A3": (420.0, 297.0),
}

DEFAULT_SCALES = [20, 50, 100, 200]


def compute_drawable_mm(paper_size_mm, scale,
                        margin_mm=30.0,
                        title_reserved_mm=100.0):
    """Drawable model-area in mm = (paper - margin - titleblock) * scale.

    Args:
        paper_size_mm: tuple (w_mm, h_mm) van het papier (landscape)
        scale: int — 1:scale (bv. 50)
        margin_mm: marge rondom op het papier
        title_reserved_mm: gereserveerde rechterstrook voor titleblock

    Returns:
        tuple (drawable_w_model_mm, drawable_h_model_mm)
    """
    pw, ph = paper_size_mm
    drawable_paper_w = pw - 2 * margin_mm - title_reserved_mm
    drawable_paper_h = ph - 2 * margin_mm
    if drawable_paper_w < 0:
        drawable_paper_w = 0.0
    if drawable_paper_h < 0:
        drawable_paper_h = 0.0
    return (drawable_paper_w * scale, drawable_paper_h * scale)


def compute_optimal_grid(item_count, cell_w_mm, cell_h_mm,
                         max_w_mm, max_h_mm):
    """Bepaal een uitgebalanceerde cols x rows voor item_count items.

    Optimaliseert op twee criteria:
      1. Aspect-ratio van de grid moet de drawable-aspect benaderen
         (zodat het canvas het blad vult i.p.v. een dunne strip)
      2. Aantal lege cellen minimaliseren

    Returns:
        tuple (cols, rows, fits_bool)
    """
    if item_count <= 0 or cell_w_mm <= 0 or cell_h_mm <= 0:
        return (0, 0, False)

    cols_max = int(max_w_mm // cell_w_mm)
    if cols_max < 1:
        cols_max = 1
    rows_max = int(max_h_mm // cell_h_mm)
    if rows_max < 1:
        rows_max = 1

    # Doel-aspect: grid_w / grid_h ≈ drawable_w / drawable_h
    # → cols/rows = (max_w/max_h) * (cell_h/cell_w)
    target_ratio = (
        (float(max_w_mm) / float(max_h_mm))
        * (float(cell_h_mm) / float(cell_w_mm))
    )

    cols_upper = min(cols_max, item_count)
    best_score = None
    best_cols = cols_upper
    best_rows = (item_count + cols_upper - 1) // cols_upper

    for cols in range(1, cols_upper + 1):
        rows = (item_count + cols - 1) // cols  # ceil
        fits_h = rows <= rows_max

        ratio = float(cols) / float(rows)
        ratio_dist = abs((ratio / target_ratio) - 1.0)
        empty = cols * rows - item_count
        empty_frac = float(empty) / float(cols * rows)

        # Penalty voor niet-passende hoogte
        penalty = 0.0 if fits_h else 5.0

        score = ratio_dist + empty_frac + penalty
        if best_score is None or score < best_score:
            best_score = score
            best_cols = cols
            best_rows = rows

    fits = (
        best_cols * cell_w_mm <= max_w_mm
        and best_rows * cell_h_mm <= max_h_mm
    )
    return (best_cols, best_rows, fits)


def compute_variable_layout(widths_mm, heights_mm,
                            spacing_mm, tag_clearance_mm,
                            drawable_w_mm, drawable_h_mm):
    """Layout met per-kozijn breedte + vaste tussenruimte.

    Verdeelt items even over een aantal rijen (bepaald via
    aspect-ratio-optimalisatie t.o.v. drawable area), plaatst
    sequentieel per rij vanaf x=0 met spacing_mm tussen items.

    Args:
        widths_mm: list[float] werkelijke breedte per kozijn
        heights_mm: list[float] werkelijke hoogte per kozijn
        spacing_mm: float horizontale tussenruimte tussen kozijnen
                    (en gelijktijdig vertikale marge tussen rijen)
        tag_clearance_mm: float ruimte onder kozijn voor tag
        drawable_w_mm, drawable_h_mm: drawable area op blad
                    (model-mm op gekozen schaal)

    Returns:
        dict met:
          'rows'         : list of list of (item_idx, x_offset_mm, w_mm)
          'canvas_w_mm'  : breedste rij + spacing aan rechterzijde
          'canvas_h_mm'  : n_rows * row_h
          'row_h_mm'     : hoogte per rij in mm
          'n_rows', 'cols_max'
          'fits'         : True als canvas binnen drawable past
        of None bij lege input.
    """
    n = len(widths_mm)
    if n == 0:
        return None

    max_h_mm = max(heights_mm) if heights_mm else 2000.0
    row_h_mm = max_h_mm + tag_clearance_mm + spacing_mm

    # Ratio-zoekstrategie: probeer 1..n rijen, kies degene waarbij
    # canvas-aspect het dichtst bij drawable-aspect ligt en zo min
    # mogelijk lege slots overhoudt.
    if drawable_h_mm <= 0 or drawable_w_mm <= 0:
        target_ratio = 1.0
    else:
        target_ratio = float(drawable_w_mm) / float(drawable_h_mm)

    avg_cell_w_mm = sum(widths_mm) / float(n) + spacing_mm

    best = None
    for n_rows_try in range(1, n + 1):
        items_per_row = (n + n_rows_try - 1) // n_rows_try
        est_canvas_w = items_per_row * avg_cell_w_mm
        est_canvas_h = n_rows_try * row_h_mm
        if est_canvas_h <= 0:
            continue
        ratio = est_canvas_w / est_canvas_h
        ratio_dist = abs((ratio / target_ratio) - 1.0) \
            if target_ratio > 0 else 0.0
        empty = n_rows_try * items_per_row - n
        empty_frac = float(empty) / float(n_rows_try * items_per_row)
        score = ratio_dist + empty_frac
        if best is None or score < best[0]:
            best = (score, n_rows_try, items_per_row)

    _, n_rows, items_per_row = best

    rows = []
    for r in range(n_rows):
        start = r * items_per_row
        end = min(start + items_per_row, n)
        if start >= n:
            break
        row = []
        x = 0.0
        for i in range(start, end):
            w = widths_mm[i]
            row.append((i, x, w))
            x += w + spacing_mm
        rows.append(row)

    if rows:
        canvas_w_mm = max(
            (row[-1][1] + row[-1][2] + spacing_mm) for row in rows
        )
    else:
        canvas_w_mm = 0.0
    canvas_h_mm = len(rows) * row_h_mm

    fits = (
        canvas_w_mm <= drawable_w_mm
        and canvas_h_mm <= drawable_h_mm
    )

    cols_max = max(len(r) for r in rows) if rows else 0

    return {
        "rows": rows,
        "canvas_w_mm": canvas_w_mm,
        "canvas_h_mm": canvas_h_mm,
        "row_h_mm": row_h_mm,
        "n_rows": len(rows),
        "cols_max": cols_max,
        "fits": fits,
    }


def compute_points_from_layout(layout, origin, u_dir, v_dir,
                               tag_clearance_mm):
    """Bereken plaatsingspunten uit een variable_layout-dict.

    Items worden geplaatst op (row_bottom + tag_clearance) verticaal en
    in het midden van hun eigen breedte horizontaal. Origin = linker-
    onder-hoek van het canvas.

    Returns:
        list[XYZ] geïndexeerd op item-index (gaten = None)
    """
    rows = layout.get("rows") or []
    if not rows:
        return []

    max_idx = -1
    for r in rows:
        for item_idx, _x, _w in r:
            if item_idx > max_idx:
                max_idx = item_idx
    points = [None] * (max_idx + 1)

    n_rows = len(rows)
    row_h_ft = layout["row_h_mm"] * MM_TO_FT
    tag_clearance_ft = tag_clearance_mm * MM_TO_FT

    for r_idx, row in enumerate(rows):
        row_bottom_v = (n_rows - r_idx - 1) * row_h_ft
        kozijn_v = row_bottom_v + tag_clearance_ft
        for item_idx, x_offset_mm, w_mm in row:
            x_center_ft = (x_offset_mm + w_mm / 2.0) * MM_TO_FT
            pt = XYZ(
                origin.X + u_dir.X * x_center_ft + v_dir.X * kozijn_v,
                origin.Y + u_dir.Y * x_center_ft + v_dir.Y * kozijn_v,
                origin.Z + u_dir.Z * x_center_ft + v_dir.Z * kozijn_v,
            )
            points[item_idx] = pt
    return points


def compute_wall_fill_layout(widths_mm, heights_mm,
                             wall_length_mm, wall_height_mm,
                             horizontal_spacing_mm=500.0,
                             row_spacing_mm=2000.0):
    """Pack kozijnen sequentieel op een wand, variabele rij-hoogte.

    Regels:
    - Items op één rij staan in volgorde naast elkaar; de afstand
      tussen het einde van item i en het begin van item i+1 is
      horizontal_spacing_mm (default 500 mm).
    - Een item dat horizontaal niet meer past op de huidige rij
      wrapt naar een nieuwe rij. De nieuwe rij begint
      row_spacing_mm (default 2000 mm) boven het hoogste kozijn
      van de huidige rij.
    - Items die ook vertikaal niet meer passen worden NIET geplaatst
      en geteld in 'overflow_count'.

    Args:
        widths_mm, heights_mm: parallelle lijsten van werkelijke
            kozijn-afmetingen (mm)
        wall_length_mm, wall_height_mm: beschikbare wand-afmetingen
        horizontal_spacing_mm: lege ruimte tussen kozijnen op een rij
        row_spacing_mm: lege ruimte tussen rijen verticaal

    Returns:
        dict of None bij lege input. Velden:
          'rows'           : list of list of
                             (item_idx, x_offset_mm, row_bottom_mm, w, h)
          'placed_count'   : aantal succesvol gepakte items
          'overflow_count' : aantal items die niet meer pasten
          'used_w_mm'      : breedste rij (rechterkant laatste kozijn)
          'used_h_mm'      : top van het hoogste kozijn op laatste rij
    """
    n = len(widths_mm)
    if n == 0 or wall_length_mm <= 0 or wall_height_mm <= 0:
        return None

    # Fase 1: pack items in rijen (rij 0 = eerste/bovenste rij),
    # bepaal max-hoogte per rij. Verticale Z wordt later toegekend.
    raw_rows = []   # list[(items, row_max_h)] waar items =
                    # list[(item_idx, x_offset, w, h)]
    current_row = []
    current_x = 0.0
    current_row_max_h = 0.0

    for i in range(n):
        w = widths_mm[i]
        h = heights_mm[i]

        if current_row and (current_x + w) > wall_length_mm:
            raw_rows.append((current_row, current_row_max_h))
            current_row = []
            current_x = 0.0
            current_row_max_h = 0.0

        current_row.append((i, current_x, w, h))
        current_x += w + horizontal_spacing_mm
        if h > current_row_max_h:
            current_row_max_h = h

    if current_row:
        raw_rows.append((current_row, current_row_max_h))

    # Fase 2: ken row_bottom toe vanaf de TOP van de wand naar beneden.
    # Eerste rij raakt de top, volgende rijen erónder met row_spacing.
    rows = []
    placed = 0
    overflow = 0
    next_row_top = wall_height_mm  # bovenkant van de eerstvolgende rij

    for r_idx, (row_items, row_max_h) in enumerate(raw_rows):
        row_bottom = next_row_top - row_max_h
        if row_bottom < 0:
            # Past niet meer, deze + alle volgende items zijn overflow
            for rr in range(r_idx, len(raw_rows)):
                overflow += len(raw_rows[rr][0])
            break

        new_row = []
        for item_idx, x_off, w, h in row_items:
            new_row.append((item_idx, x_off, row_bottom, w, h))
            placed += 1
        rows.append(new_row)

        next_row_top = row_bottom - row_spacing_mm

    used_w = 0.0
    for row in rows:
        for _idx, x_off, _zb, w, _h in row:
            right = x_off + w
            if right > used_w:
                used_w = right
    if rows:
        last_row_bot = rows[-1][0][2]
        used_h = wall_height_mm - last_row_bot
    else:
        used_h = 0.0

    return {
        "rows": rows,
        "placed_count": placed,
        "overflow_count": overflow,
        "used_w_mm": used_w,
        "used_h_mm": used_h,
    }


def compute_points_from_wall_fill(layout, origin, u_dir, v_dir):
    """Bereken plaatsingspunten uit een wall_fill layout.

    Plaatsingspunt = (x_offset + w/2) horizontaal,
                     row_bottom verticaal (kozijn-vloer).
    Origin = linker-onder-hoek van het canvas (= wall start point).

    Returns:
        list[XYZ] geïndexeerd op item-index (None waar overflow)
    """
    rows = layout.get("rows") or []
    if not rows:
        return []

    max_idx = -1
    for r in rows:
        for item_idx, _x, _zb, _w, _h in r:
            if item_idx > max_idx:
                max_idx = item_idx
    points = [None] * (max_idx + 1)

    for row in rows:
        for item_idx, x_offset_mm, row_bottom_mm, w_mm, _h_mm in row:
            x_center_ft = (x_offset_mm + w_mm / 2.0) * MM_TO_FT
            z_ft = row_bottom_mm * MM_TO_FT
            pt = XYZ(
                origin.X + u_dir.X * x_center_ft + v_dir.X * z_ft,
                origin.Y + u_dir.Y * x_center_ft + v_dir.Y * z_ft,
                origin.Z + u_dir.Z * x_center_ft + v_dir.Z * z_ft,
            )
            points[item_idx] = pt
    return points


def compute_grid_points_tight(origin, u_dir, v_dir,
                              cell_w_ft, cell_h_ft,
                              cols, rows, item_count,
                              kozijn_floor_offset_ft=0.0):
    """Plaats items in een grid met vaste celgrootte.

    Origin = linker-onder-hoek van het canvas. Rij 0 = bovenste rij.
    Het plaatsingspunt zit horizontaal in het midden van de cel en
    verticaal op cell_bottom + kozijn_floor_offset_ft (zodat er ruimte
    onder het kozijn over blijft voor een tag).

    Args:
        origin: XYZ linker-onder-hoek (feet)
        u_dir, v_dir: orthogonale richtingsvectoren
        cell_w_ft, cell_h_ft: celafmetingen in feet
        cols, rows: grid dimensies
        item_count: aantal te plaatsen items
        kozijn_floor_offset_ft: verticale offset boven cell_bottom
            voor het kozijn-plaatsingspunt (om plek te laten voor tag)

    Returns:
        list[XYZ] van lengte min(item_count, cols*rows)
    """
    if cols <= 0 or rows <= 0:
        return []
    points = []
    for row in range(rows):
        cell_bottom_v = (rows - row - 1) * cell_h_ft
        kozijn_v = cell_bottom_v + kozijn_floor_offset_ft
        for col in range(cols):
            if len(points) >= item_count:
                return points
            u_center = (col + 0.5) * cell_w_ft
            pt = XYZ(
                origin.X + u_dir.X * u_center + v_dir.X * kozijn_v,
                origin.Y + u_dir.Y * u_center + v_dir.Y * kozijn_v,
                origin.Z + u_dir.Z * u_center + v_dir.Z * kozijn_v,
            )
            points.append(pt)
    return points


def compute_grid_points(origin, u_dir, v_dir, u_length, v_length,
                        cols, rows, item_count):
    """Bereken plaatsingspunten voor een grid op een vlak.

    Het vlak wordt beschreven door een origin + twee orthogonale
    richtingsvectoren (u = horizontaal, v = verticaal). Punten
    worden per rij van boven naar onder uitgedeeld en per rij van
    links naar rechts gevuld.

    Args:
        origin: XYZ linker-onder-hoek van het canvas (feet)
        u_dir: XYZ horizontaal normalized richtingsvector
        v_dir: XYZ verticaal normalized richtingsvector
        u_length: float, breedte canvas in feet
        v_length: float, hoogte canvas in feet
        cols: int, aantal kolommen
        rows: int, aantal rijen
        item_count: int, aantal items om te plaatsen (kan < cols*rows)

    Returns:
        list[XYZ] van lengte min(item_count, cols*rows)
    """
    if cols <= 0 or rows <= 0:
        return []

    u_step = u_length / float(cols)
    v_step = v_length / float(rows)

    points = []
    # Rijen van boven (row=0) naar onder (row=rows-1)
    for row in range(rows):
        v_param = v_length - (row + 0.5) * v_step  # midden van de rij
        for col in range(cols):
            if len(points) >= item_count:
                return points
            u_param = (col + 0.5) * u_step  # midden van de kolom
            pt = XYZ(
                origin.X + u_dir.X * u_param + v_dir.X * v_param,
                origin.Y + u_dir.Y * u_param + v_dir.Y * v_param,
                origin.Z + u_dir.Z * u_param + v_dir.Z * v_param,
            )
            points.append(pt)
    return points


def compute_tag_points(placement_points, tag_offset_ft):
    """Bepaal tag-posities als Z-offset t.o.v. plaatsingspunten.

    Args:
        placement_points: list[XYZ]
        tag_offset_ft: float offset in feet (negatief = onder)

    Returns:
        list[XYZ]
    """
    return [
        XYZ(p.X, p.Y, p.Z + tag_offset_ft)
        for p in placement_points
    ]


def estimate_canvas_size_mm(widths_mm, heights_mm, cols, rows,
                            padding_mm=500.0):
    """Schat minimum canvas-afmetingen voor een grid van kozijnen.

    Args:
        widths_mm: list[float] breedtes van te plaatsen kozijnen
        heights_mm: list[float] hoogtes
        cols, rows: grid-dimensies
        padding_mm: marge tussen cellen

    Returns:
        tuple (breedte_mm, hoogte_mm)
    """
    if not widths_mm:
        widths_mm = [1000.0]
    if not heights_mm:
        heights_mm = [2000.0]

    max_w = max(widths_mm) + padding_mm
    max_h = max(heights_mm) + padding_mm
    return (max_w * cols, max_h * rows)
