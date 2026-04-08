# -*- coding: utf-8 -*-
"""
Test script voor Kadaster WFS API
Runt BUITEN Revit om de WFS calls te verifiëren.
"""
import sys
import os

# Add lib to path
lib_path = os.path.join(os.path.dirname(__file__), 'lib')
sys.path.insert(0, lib_path)

from gis2bim.api.pdok import PDOKKadaster

def main():
    kad = PDOKKadaster()
    
    # Test bbox: gebied rond Gorinchem centrum
    xmin, ymin = 126000, 426000
    xmax, ymax = 126500, 426500
    
    print("=" * 60)
    print("TEST KADASTER WFS - JSON FORMAT")
    print("Bbox: {}, {}, {}, {}".format(xmin, ymin, xmax, ymax))
    print("=" * 60)
    
    # Test 1: Percelen
    print("\n[1] Percelen...")
    percelen = kad.get_percelen(xmin, ymin, xmax, ymax)
    print("   Gevonden: {} percelen".format(len(percelen)))
    if percelen:
        p = percelen[0]
        print("   Eerste: nr={}, sectie={}, vertices={}".format(
            p.perceelnummer, p.sectie, len(p.geometry)))
    
    # Test 2: Annotaties
    print("\n[2] Perceelnummer annotaties...")
    annotaties = kad.get_perceel_annotaties(xmin, ymin, xmax, ymax)
    print("   Gevonden: {} annotaties".format(len(annotaties)))
    if annotaties:
        a = annotaties[0]
        print("   Eerste: '{}' @ ({:.1f}, {:.1f}) rot={:.1f}".format(
            a.tekst, a.x, a.y, a.rotatie))
    
    # Test 3: Straatnamen
    print("\n[3] Straatnamen...")
    straatnamen = kad.get_straatnamen(xmin, ymin, xmax, ymax)
    print("   Gevonden: {} straatnamen".format(len(straatnamen)))
    if straatnamen:
        s = straatnamen[0]
        print("   Eerste: '{}' @ ({:.1f}, {:.1f})".format(s.tekst, s.x, s.y))
    
    # Test 4: Huisnummers
    print("\n[4] Huisnummers...")
    huisnummers = kad.get_huisnummers(xmin, ymin, xmax, ymax)
    print("   Gevonden: {} huisnummers".format(len(huisnummers)))
    if huisnummers:
        h = huisnummers[0]
        print("   Eerste: '{}' @ ({:.1f}, {:.1f})".format(h.volledig, h.x, h.y))
    
    print("\n" + "=" * 60)
    print("TEST COMPLEET")
    print("=" * 60)
    
    return len(percelen) > 0

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
