# -*- coding: utf-8 -*-
"""
Detail Items Overzicht - TEST
=============================
Test: laad 00_DI_hout met alle catalog types en plaats in drafting view.
"""
__title__ = "Detail\nOverzicht"
__author__ = "3BM Bouwkunde"

import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from bm_logger import get_logger

log = get_logger("DetailOverzicht")

from pyrevit import revit, forms, script

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    FilteredElementCollector, Family, FamilySymbol,
    ViewDrafting, ViewFamilyType, Transaction,
    XYZ, ElementTransformUtils
)

output = script.get_output()
doc = revit.doc

RFA_PATH = r"Z:\50_projecten\7_3BM_bouwkunde\000_revit\00_bibliotheek\00_detail_components\00_algemeen\00_DI_hout.rfa"
TXT_PATH = r"Z:\50_projecten\7_3BM_bouwkunde\000_revit\00_bibliotheek\00_detail_components\00_algemeen\00_DI_hout.txt"
VIEW_NAME = "00_TEST_hout"
MM = 1.0 / 304.8
PADDING = 5 * MM


def parse_catalog(txt_path):
    """Lees type-namen uit catalog file."""
    names = []
    with open(txt_path, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith(";"):
                continue
            name = line.split(";")[0].strip()
            if name:
                names.append(name)
    return names


def main():
    output.print_md("# TEST: 00_DI_hout laden en plaatsen")
    output.print_md("")

    # 1. Parse catalog
    type_names = parse_catalog(TXT_PATH)
    output.print_md("## 1. Catalog: {} types in .txt".format(len(type_names)))
    for n in type_names[:5]:
        output.print_md("  - {}".format(n))
    output.print_md("  - ...")
    output.print_md("")

    # 2. Laad family + types
    output.print_md("## 2. Laden")

    # Probeer eerst LoadFamilySymbol per type
    ok = 0
    skip = 0
    fail = 0
    for type_name in type_names:
        try:
            result = doc.LoadFamilySymbol(RFA_PATH, type_name)
            if result:
                ok += 1
            else:
                skip += 1
        except Exception as e:
            fail += 1
            if fail <= 3:
                output.print_md("  - FOUT {}: {}".format(type_name, e))

    output.print_md("- LoadFamilySymbol: ok={}, skip={}, fail={}".format(
        ok, skip, fail))

    # 3. Check of family in document zit
    family = None
    collector = FilteredElementCollector(doc).OfClass(Family)
    for fam in collector:
        if fam.Name == "00_DI_hout":
            family = fam
            break

    if not family:
        output.print_md("- **Family NIET gevonden in document!**")
        # Probeer LoadFamily als fallback
        output.print_md("- Probeer doc.LoadFamily()...")
        try:
            result = doc.LoadFamily(RFA_PATH)
            output.print_md("  - LoadFamily result: {}".format(result))
        except Exception as e:
            output.print_md("  - LoadFamily FOUT: {}".format(e))

        # Opnieuw zoeken
        collector = FilteredElementCollector(doc).OfClass(Family)
        for fam in collector:
            if fam.Name == "00_DI_hout":
                family = fam
                break

    if not family:
        forms.alert("Family 00_DI_hout niet geladen. Stop.")
        return

    # 4. Tel types
    sym_ids = family.GetFamilySymbolIds()
    output.print_md("- Family gevonden, {} FamilySymbolIds".format(
        sym_ids.Count))
    output.print_md("")

    # 5. Verzamel symbols
    symbols = []
    for sym_id in sym_ids:
        sym = doc.GetElement(sym_id)
        if sym:
            symbols.append(sym)

    output.print_md("## 3. {} symbols verzameld".format(len(symbols)))
    if not symbols:
        forms.alert("Geen symbols gevonden. Stop.")
        return

    # Sorteer op naam
    def get_sym_name(sym):
        try:
            return sym.Family.Name + ":" + sym.LookupParameter("Type Name").AsString()
        except:
            return ""
    symbols.sort(key=get_sym_name)

    # 6. Maak drafting view + plaats items
    output.print_md("")
    output.print_md("## 4. Plaatsen in drafting view")

    t = Transaction(doc, "TEST hout plaatsen")
    t.Start()

    try:
        # Verwijder bestaande view
        vc = FilteredElementCollector(doc).OfClass(ViewDrafting)
        for v in vc:
            try:
                if v.Name == VIEW_NAME:
                    doc.Delete(v.Id)
                    break
            except:
                pass

        # Maak view
        vft_collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
        type_id = None
        for vft in vft_collector:
            if vft.ViewFamily.ToString() == "Drafting":
                type_id = vft.Id
                break

        view = ViewDrafting.Create(doc, type_id)
        view.Name = VIEW_NAME
        view.Scale = 1

        # Activeer alle symbols
        for sym in symbols:
            if not sym.IsActive:
                sym.Activate()
        doc.Regenerate()

        # Plaats per stuk in een rij
        x = 0.0
        placed = 0
        errors = []

        for sym in symbols:
            try:
                inst = doc.Create.NewFamilyInstance(
                    XYZ(x, 0, 0), sym, view)
            except Exception as e:
                errors.append(str(e))
                continue

            bb = inst.get_BoundingBox(view)
            if bb:
                # Schuif linkerkant naar x
                shift = x - bb.Min.X
                ElementTransformUtils.MoveElement(
                    doc, inst.Id, XYZ(shift, 0, 0))
                width = bb.Max.X - bb.Min.X
                x += width + PADDING
            else:
                x += 20 * MM

            placed += 1

        t.Commit()

    except Exception as e:
        t.RollBack()
        output.print_md("**FOUT:** {}".format(e))
        import traceback
        output.print_md("```")
        output.print_md(traceback.format_exc())
        output.print_md("```")
        return

    output.print_md("- **{} items geplaatst**".format(placed))
    if errors:
        output.print_md("- {} fouten:".format(len(errors)))
        for e in errors[:5]:
            output.print_md("  - {}".format(e))

    try:
        revit.uidoc.ActiveView = view
    except:
        pass

    output.print_md("")
    output.print_md("**Klaar!**")


if __name__ == "__main__":
    main()
