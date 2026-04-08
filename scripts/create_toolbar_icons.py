# -*- coding: utf-8 -*-
"""
Genereer PNG icons voor Schedule Export/Import, Sheet Parameters en NAA.K.T.
32x32 pixels met 3BM huisstijl
"""
from PIL import Image, ImageDraw
import os

# Kleuren (3BM huisstijl)
VIOLET = (53, 14, 53)
TEAL = (69, 182, 168)

def create_schedule_export_icon(path):
    """Spreadsheet met export pijl naar rechts"""
    img = Image.new('RGBA', (32, 32), VIOLET + (255,))
    draw = ImageDraw.Draw(img)
    
    # Spreadsheet
    draw.rectangle([6, 6, 20, 22], outline=TEAL, width=2)
    
    # Grid lijnen horizontaal
    draw.line([(6, 10), (20, 10)], fill=TEAL, width=1)
    draw.line([(6, 14), (20, 14)], fill=TEAL, width=1)
    draw.line([(6, 18), (20, 18)], fill=TEAL, width=1)
    
    # Grid lijnen verticaal
    draw.line([(11, 6), (11, 22)], fill=TEAL, width=1)
    draw.line([(15, 6), (15, 22)], fill=TEAL, width=1)
    
    # Export pijl naar rechts
    draw.line([(21, 14), (27, 14)], fill=TEAL, width=2)
    draw.polygon([(27, 14), (24, 11), (24, 17)], fill=TEAL)
    
    img.save(path)
    print("Created: {}".format(path))

def create_schedule_import_icon(path):
    """Spreadsheet met import pijl naar links"""
    img = Image.new('RGBA', (32, 32), VIOLET + (255,))
    draw = ImageDraw.Draw(img)
    
    # Spreadsheet
    draw.rectangle([12, 6, 26, 22], outline=TEAL, width=2)
    
    # Grid lijnen horizontaal
    draw.line([(12, 10), (26, 10)], fill=TEAL, width=1)
    draw.line([(12, 14), (26, 14)], fill=TEAL, width=1)
    draw.line([(12, 18), (26, 18)], fill=TEAL, width=1)
    
    # Grid lijnen verticaal
    draw.line([(17, 6), (17, 22)], fill=TEAL, width=1)
    draw.line([(21, 6), (21, 22)], fill=TEAL, width=1)
    
    # Import pijl naar links
    draw.line([(5, 14), (11, 14)], fill=TEAL, width=2)
    draw.polygon([(5, 14), (8, 11), (8, 17)], fill=TEAL)
    
    img.save(path)
    print("Created: {}".format(path))

def create_sheet_parameters_icon(path):
    """A0 sheet met titleblock"""
    img = Image.new('RGBA', (32, 32), VIOLET + (255,))
    draw = ImageDraw.Draw(img)
    
    # A0 Sheet outline (landscape)
    draw.rectangle([4, 7, 28, 25], outline=TEAL, width=2)
    
    # Titleblock (gevuld)
    draw.rectangle([4, 22, 28, 25], fill=TEAL + (80,), outline=TEAL)
    
    # Sheet lines (schematisch tekening)
    teal_light = TEAL + (128,)
    draw.line([(8, 11), (18, 11)], fill=teal_light, width=1)
    draw.line([(8, 14), (15, 14)], fill=teal_light, width=1)
    draw.line([(8, 17), (20, 17)], fill=teal_light, width=1)
    draw.line([(20, 11), (24, 11)], fill=teal_light, width=1)
    draw.line([(20, 14), (24, 14)], fill=teal_light, width=1)
    
    # Parameter symbool (settings gear) in titleblock
    draw.ellipse([24, 22, 27, 25], outline=TEAL, width=1)
    draw.line([(25.5, 21), (25.5, 22)], fill=TEAL, width=1)
    draw.line([(25.5, 25), (25.5, 26)], fill=TEAL, width=1)
    draw.line([(23, 23.5), (24, 23.5)], fill=TEAL, width=1)
    draw.line([(27, 23.5), (28, 23.5)], fill=TEAL, width=1)
    
    img.save(path)
    print("Created: {}".format(path))

def create_naakt_icon(path):
    """Materiaal lagen opbouw"""
    img = Image.new('RGBA', (32, 32), VIOLET + (255,))
    draw = ImageDraw.Draw(img)
    
    # Materiaal lagen (verschillende opacity)
    # Laag 1 (beton/draagconstructie)
    draw.rectangle([6, 8, 26, 11], fill=TEAL + (204,))
    
    # Laag 2 (isolatie)
    draw.rectangle([6, 11, 26, 14], fill=TEAL + (128,))
    
    # Laag 3 (spouw/lucht) - dashed
    for x in range(6, 26, 3):
        draw.line([(x, 14), (min(x+1, 26), 14)], fill=TEAL, width=2)
    
    # Laag 4 (afwerking/gevel)
    draw.rectangle([6, 16, 26, 19], fill=TEAL + (153,))
    
    # Dimensie pijltjes (schematisch)
    draw.line([(4, 8), (4, 19)], fill=TEAL, width=1)
    # Pijlpunten
    draw.line([(3, 9), (4, 8)], fill=TEAL, width=1)
    draw.line([(5, 9), (4, 8)], fill=TEAL, width=1)
    draw.line([(3, 18), (4, 19)], fill=TEAL, width=1)
    draw.line([(5, 18), (4, 19)], fill=TEAL, width=1)
    
    # NAA.K.T. tekst
    try:
        from PIL import ImageFont
        font = ImageFont.truetype("segoeui.ttf", 7)
        draw.text((16, 23), "NAA.K.T.", fill=TEAL, anchor="mm", font=font)
    except:
        # Fallback zonder font
        pass
    
    img.save(path)
    print("Created: {}".format(path))

# Main
if __name__ == "__main__":
    # Test PIL import
    try:
        from PIL import Image
        print("PIL is geïnstalleerd")
    except ImportError:
        print("PIL niet gevonden. Installeer met: pip install Pillow")
        import sys
        sys.exit(1)
    base = r"X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension\Bouwkunde.tab"
    
    # Schedule Export
    exp_path = os.path.join(base, "Data Exchange.panel", "ScheduleExport.pushbutton", "icon.png")
    create_schedule_export_icon(exp_path)
    
    # Schedule Import
    imp_path = os.path.join(base, "Data Exchange.panel", "ScheduleImport.pushbutton", "icon.png")
    create_schedule_import_icon(imp_path)
    
    # Sheet Parameters
    sheet_path = os.path.join(base, "Document.panel", "SheetParameters.pushbutton", "icon.png")
    create_sheet_parameters_icon(sheet_path)
    
    # NAA.K.T.
    naakt_path = os.path.join(base, "Materialen.panel", "NAAKTGenerator.pushbutton", "icon.png")
    create_naakt_icon(naakt_path)
    
    print("\nAll icons created! Sync to runtime and reload pyRevit.")
