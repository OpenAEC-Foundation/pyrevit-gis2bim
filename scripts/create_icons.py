# Icon generator voor Mat Exp en Mat Imp
# Run dit eenmalig in Revit Python Shell of via pyRevit

import clr
clr.AddReference('System.Drawing')
from System.Drawing import Bitmap, Graphics, Color, Pen, SolidBrush, Rectangle, Point
from System.Drawing.Drawing2D import SmoothingMode
import os

def create_icon(output_path, is_export=True):
    """Maak een 32x32 icoon met materiaal bol en pijl"""
    size = 32
    bmp = Bitmap(size, size)
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.AntiAlias
    
    # Achtergrond transparant (wit voor nu)
    g.Clear(Color.Transparent)
    
    # Materiaal bol - gradient effect met 2 cirkels
    # Hoofdcirkel
    brush_main = SolidBrush(Color.FromArgb(255, 69, 182, 168))  # Teal
    g.FillEllipse(brush_main, 4, 6, 18, 18)
    
    # Highlight
    brush_light = SolidBrush(Color.FromArgb(150, 255, 255, 255))
    g.FillEllipse(brush_light, 7, 8, 6, 6)
    
    # Rand
    pen_border = Pen(Color.FromArgb(255, 53, 14, 53), 1.5)  # Violet
    g.DrawEllipse(pen_border, 4, 6, 18, 18)
    
    # Pijl
    pen_arrow = Pen(Color.FromArgb(255, 53, 14, 53), 2.5)  # Violet
    brush_arrow = SolidBrush(Color.FromArgb(255, 53, 14, 53))
    
    if is_export:
        # Pijl naar rechts (export)
        # Lijn
        g.DrawLine(pen_arrow, 20, 15, 29, 15)
        # Pijlpunt
        points = [Point(29, 15), Point(24, 10), Point(24, 20)]
        g.FillPolygon(brush_arrow, points)
    else:
        # Pijl naar links (import)
        # Lijn
        g.DrawLine(pen_arrow, 22, 15, 31, 15)
        # Pijlpunt
        points = [Point(22, 15), Point(27, 10), Point(27, 20)]
        g.FillPolygon(brush_arrow, points)
    
    # Opslaan
    bmp.Save(output_path)
    g.Dispose()
    bmp.Dispose()
    print("Icon saved: " + output_path)

# Paden
base = r"X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension\Bouwkunde.tab\Template.panel"

create_icon(os.path.join(base, "MatExp.pushbutton", "icon.png"), is_export=True)
create_icon(os.path.join(base, "MatImp.pushbutton", "icon.png"), is_export=False)

print("Done!")
