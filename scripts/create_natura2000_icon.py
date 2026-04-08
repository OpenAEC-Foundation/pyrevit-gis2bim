# -*- coding: utf-8 -*-
"""
Genereer Natura 2000 icon voor GIS2BIM pyRevit toolbar.
32x32 pixels met 3BM huisstijl: Teal blad/natuur silhouet + Yellow accent.
Gebruikt 10x supersampling voor anti-aliasing.
"""
from PIL import Image, ImageDraw
import os
import math

# Kleuren (3BM huisstijl)
VIOLET = (53, 14, 53)
TEAL = (69, 182, 168)
YELLOW = (239, 189, 117)

# Supersampling factor
SS = 10
SIZE = 32
SS_SIZE = SIZE * SS


def create_natura2000_icon(path):
    """Teal blad/natuur silhouet met yellow accent marker."""
    img = Image.new('RGBA', (SS_SIZE, SS_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Blad silhouet (gestileerd) - teal
    # Bladsteel (stengel)
    stem_x = 16 * SS
    stem_top = 7 * SS
    stem_bottom = 26 * SS
    draw.line(
        [(stem_x, stem_top), (stem_x, stem_bottom)],
        fill=TEAL, width=int(1.5 * SS)
    )

    # Blad vorm (ellips-achtig, iets schuin)
    # Linker blad
    leaf_points_left = [
        (16 * SS, 8 * SS),    # top
        (10 * SS, 12 * SS),   # linker curve
        (8 * SS, 16 * SS),    # linker breedte
        (10 * SS, 20 * SS),   # linker onder
        (16 * SS, 22 * SS),   # onderkant midden
    ]
    # Rechter blad
    leaf_points_right = [
        (16 * SS, 22 * SS),   # onderkant midden
        (22 * SS, 20 * SS),   # rechter onder
        (24 * SS, 16 * SS),   # rechter breedte
        (22 * SS, 12 * SS),   # rechter curve
        (16 * SS, 8 * SS),    # top
    ]

    # Teken gevuld blad
    leaf_points = leaf_points_left + leaf_points_right
    draw.polygon(leaf_points, fill=TEAL + (200,))

    # Bladnerven (dunne lijnen in violet)
    nerve_color = VIOLET + (180,)
    # Middennerf
    draw.line(
        [(16 * SS, 9 * SS), (16 * SS, 21 * SS)],
        fill=nerve_color, width=int(0.8 * SS)
    )
    # Zijnerven links
    draw.line([(16 * SS, 13 * SS), (11 * SS, 11 * SS)],
              fill=nerve_color, width=int(0.6 * SS))
    draw.line([(16 * SS, 16 * SS), (10 * SS, 15 * SS)],
              fill=nerve_color, width=int(0.6 * SS))
    draw.line([(16 * SS, 19 * SS), (11 * SS, 19 * SS)],
              fill=nerve_color, width=int(0.6 * SS))
    # Zijnerven rechts
    draw.line([(16 * SS, 13 * SS), (21 * SS, 11 * SS)],
              fill=nerve_color, width=int(0.6 * SS))
    draw.line([(16 * SS, 16 * SS), (22 * SS, 15 * SS)],
              fill=nerve_color, width=int(0.6 * SS))
    draw.line([(16 * SS, 19 * SS), (21 * SS, 19 * SS)],
              fill=nerve_color, width=int(0.6 * SS))

    # Yellow accent: locatie marker (klein, rechtsonder)
    marker_cx = 24 * SS
    marker_cy = 25 * SS
    marker_r = 3 * SS
    draw.ellipse(
        [marker_cx - marker_r, marker_cy - marker_r,
         marker_cx + marker_r, marker_cy + marker_r],
        fill=YELLOW
    )
    # Stip in marker
    dot_r = int(1.2 * SS)
    draw.ellipse(
        [marker_cx - dot_r, marker_cy - dot_r,
         marker_cx + dot_r, marker_cy + dot_r],
        fill=VIOLET
    )

    # Downsample naar 32x32 met anti-aliasing
    img = img.resize((SIZE, SIZE), Image.LANCZOS)
    img.save(path)
    print("Created: {}".format(path))


if __name__ == "__main__":
    icon_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "extensions", "GIS2BIM.extension", "GIS2BIM.tab",
        "Data.panel", "Natura2000.pushbutton", "icon.png"
    )

    # Fallback: als het pad niet bestaat, sla op in huidige dir
    icon_dir = os.path.dirname(icon_path)
    if not os.path.isdir(icon_dir):
        icon_path = os.path.join(os.path.dirname(__file__), "natura2000_icon.png")

    create_natura2000_icon(icon_path)
