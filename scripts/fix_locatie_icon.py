# Fix GIS2BIM Locatie icon - pin correct centeren
# 8x supersampling + LANCZOS downscale voor scherpe resultaten

from PIL import Image, ImageDraw
import math
import os

def draw_map_pin(draw, cx, cy, radius, pin_height, color, inner_color, stroke_width):
    """Teken een map pin met cirkel bovenaan en punt onderaan."""
    # De pin bestaat uit:
    # 1. Een cirkel (de "kop") op positie (cx, cy) met straal radius
    # 2. Twee lijnen die van de cirkel naar een punt onderaan lopen

    # Punt onderaan de pin
    tip_y = cy + pin_height

    # Hoek waar de tangentlijnen de cirkel raken
    # sin(angle) = radius / pin_height (benadering)
    angle = math.asin(min(radius / (pin_height * 0.7), 0.99))

    # Tangent punten op de cirkel
    left_x = cx - radius * math.cos(angle) * 0.85
    left_y = cy + radius * math.sin(angle) * 0.85
    right_x = cx + radius * math.cos(angle) * 0.85
    right_y = cy + radius * math.sin(angle) * 0.85

    # Teken de druppelvorm als polygon path
    # We tekenen een cirkel + driehoek naar de punt
    steps = 60
    points = []

    # Bovenste deel: cirkel van rechts-onder naar links-onder (over de bovenkant)
    start_angle = math.atan2(left_y - cy, left_x - cx)
    end_angle = math.atan2(right_y - cy, right_x - cx)

    # Van rechts-onder rond de bovenkant naar links-onder
    for i in range(steps + 1):
        t = i / steps
        a = end_angle + t * (2 * math.pi - (end_angle - start_angle))
        px = cx + radius * math.cos(a)
        py = cy + radius * math.sin(a)
        points.append((px, py))

    # Punt onderaan
    points.append((cx, tip_y))

    # Teken gevulde vorm
    draw.polygon(points, fill=color)

    # Teken outline voor meer definitie
    draw.polygon(points, outline=color, width=stroke_width)

    # Binnenste cirkel (het "gat" in de pin)
    inner_radius = radius * 0.45
    draw.ellipse(
        [cx - inner_radius, cy - inner_radius,
         cx + inner_radius, cy + inner_radius],
        fill=inner_color
    )


def create_locatie_icon(output_path, size, padding):
    """Maak een locatie pin icoon."""
    scale = 8  # supersampling factor
    s = size * scale
    p = padding * scale

    img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Kleuren
    teal = (69, 182, 168, 255)
    white = (255, 255, 255, 0)  # transparant voor het binnenste

    # Pin dimensies - gecentreerd in het tekenvlak
    draw_area = s - 2 * p

    # De pin is hoger dan breed, dus we centeren verticaal
    pin_total_width = draw_area * 0.75
    pin_radius = pin_total_width / 2

    # Cirkel centrum - iets boven het midden zodat de punt onderaan past
    cx = s / 2
    cy = p + pin_radius + draw_area * 0.02

    # Pin hoogte (van cirkel centrum tot punt)
    pin_height = draw_area - pin_radius - draw_area * 0.02

    # Stroke width
    stroke_w = max(2 * scale, int(s * 0.03))

    # Teken de pin
    draw_map_pin(draw, cx, cy, pin_radius, pin_height, teal, (0, 0, 0, 0), stroke_w)

    # Outline ring in de pin (teal ring, transparant midden)
    inner_r = pin_radius * 0.45
    ring_width = max(scale * 1.5, pin_radius * 0.12)
    draw.ellipse(
        [cx - inner_r, cy - inner_r,
         cx + inner_r, cy + inner_r],
        fill=(0, 0, 0, 0),
        outline=teal,
        width=int(ring_width)
    )

    # Downscale met LANCZOS
    img = img.resize((size, size), Image.LANCZOS)
    img.save(output_path)
    print("Saved: {}".format(output_path))


# Output paden
base_icons = r"X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\GIS2BIM_Icons_StijlA"
pushbutton = r"X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\GIS2BIM.extension\GIS2BIM.tab\Setup.panel\Locatie.pushbutton"

# Genereer 96x96 en 32x32
create_locatie_icon(os.path.join(base_icons, "GIS_Locatie_96.png"), 96, 12)
create_locatie_icon(os.path.join(base_icons, "GIS_Locatie_32.png"), 32, 4)

# Kopieer 96x96 als icon.png voor de pushbutton (pyRevit gebruikt icon.png)
create_locatie_icon(os.path.join(pushbutton, "icon.png"), 96, 12)

print("Done! Locatie icons zijn opnieuw gegenereerd.")
