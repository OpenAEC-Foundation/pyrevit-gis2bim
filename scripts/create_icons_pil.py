# Icon generator voor Mat Exp en Mat Imp
from PIL import Image, ImageDraw
import os

def create_icon(output_path, is_export=True):
    """Maak een 32x32 icoon met materiaal bol en pijl"""
    size = 32
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Kleuren
    teal = (69, 182, 168, 255)
    violet = (53, 14, 53, 255)
    white_highlight = (255, 255, 255, 150)
    
    # Materiaal bol
    draw.ellipse([4, 6, 22, 24], fill=teal, outline=violet, width=2)
    
    # Highlight
    draw.ellipse([7, 8, 13, 14], fill=white_highlight)
    
    # Pijl
    if is_export:
        # Pijl naar rechts
        draw.line([(20, 15), (29, 15)], fill=violet, width=3)
        draw.polygon([(29, 15), (24, 10), (24, 20)], fill=violet)
    else:
        # Pijl naar links
        draw.line([(22, 15), (31, 15)], fill=violet, width=3)
        draw.polygon([(22, 15), (27, 10), (27, 20)], fill=violet)
    
    img.save(output_path)
    print(f"Saved: {output_path}")

base = r"X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension\Bouwkunde.tab\Template.panel"

create_icon(os.path.join(base, "MatExp.pushbutton", "icon.png"), is_export=True)
create_icon(os.path.join(base, "MatImp.pushbutton", "icon.png"), is_export=False)
print("Done!")
