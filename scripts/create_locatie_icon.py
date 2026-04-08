# Icon generator voor GIS2BIM Locatie button
# Run in Revit Python Shell of via pyRevit

import clr
clr.AddReference('System.Drawing')
from System.Drawing import Bitmap, Graphics, Color, Pen, SolidBrush, Rectangle, Point
from System.Drawing.Drawing2D import SmoothingMode
import os

def create_locatie_icon(output_path):
    """Maak 32x32 icoon: Nederland silhouet met crosshair"""
    size = 32
    bmp = Bitmap(size, size)
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.AntiAlias
    
    # Achtergrond - Violet
    brush_bg = SolidBrush(Color.FromArgb(255, 53, 14, 53))
    g.FillRectangle(brush_bg, 0, 0, size, size)
    
    # Nederland vorm - Teal (simpele rechthoek/pentagon)
    brush_nl = SolidBrush(Color.FromArgb(255, 69, 182, 168))
    points_nl = [
        Point(10, 6),
        Point(22, 6),
        Point(24, 10),
        Point(22, 20),
        Point(16, 26),
        Point(10, 20),
        Point(8, 10)
    ]
    g.FillPolygon(brush_nl, points_nl)
    
    # Crosshair - Geel
    pen_cross = Pen(Color.FromArgb(255, 239, 189, 117), 2)
    
    # Cirkel
    g.DrawEllipse(pen_cross, 12, 11, 8, 8)
    
    # Lijnen
    g.DrawLine(pen_cross, 16, 6, 16, 11)   # boven
    g.DrawLine(pen_cross, 16, 19, 16, 24)  # onder
    g.DrawLine(pen_cross, 7, 15, 12, 15)   # links
    g.DrawLine(pen_cross, 20, 15, 25, 15)  # rechts
    
    # Opslaan
    bmp.Save(output_path)
    g.Dispose()
    bmp.Dispose()
    print("Icon saved: " + output_path)

# Pad
output = r"X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\GIS2BIM.extension\GIS2BIM.tab\Setup.panel\Locatie.pushbutton\icon.png"
create_locatie_icon(output)
print("Done!")
