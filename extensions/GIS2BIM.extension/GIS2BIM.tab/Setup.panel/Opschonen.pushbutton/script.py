# -*- coding: utf-8 -*-
"""
Opschonen Button - GIS2BIM (3BM WPF-template)
=============================================

Verwijdert geimporteerde GIS2BIM-data zodat je het project opnieuw kunt
opbouwen. Port van GIS2BIM_clean_v7.dyn (alleen ImageTypes), uitgebreid
met filled regions, detail lijnen, tekstnotities, detail groepen en detail
componenten. UI gebruikt het 3BM WPF-template (wpf_template.WPFWindow).

Scope: hele document (default) of alleen de actieve view.
Rasterafbeeldingen met de prefix 'legend' worden altijd behouden
(zoals in het oorspronkelijke Dynamo-script). ImageTypes zijn
project-breed en negeren de scope-keuze.
"""

__title__ = "Opschonen"
__author__ = "OpenAEC Foundation"
__doc__ = ("Verwijder geimporteerde GIS2BIM-data (afbeeldingen, filled "
           "regions, detail lijnen, tekst, detail groepen/componenten).")

# Standaard library
import sys
import os
import traceback

# Voeg lib folder toe aan path (gedeelde modules)
extension_path = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(extension_path, "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# pyRevit
from pyrevit import revit, DB, forms

# 3BM WPF-template
from wpf_template import WPFWindow

try:
    from bm_logger import get_logger
    log = get_logger("Opschonen")
except Exception:
    def log(msg):
        print(msg)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ImageType,
    FilledRegion,
    CurveElement,
    CurveElementType,
    TextNote,
    BuiltInCategory,
    BuiltInParameter,
)

# Behoud-prefix conform GIS2BIM_clean_v7.dyn
KEEP_PREFIX = "legend"


# =============================================================================
# COLLECTORS / DELETE
# =============================================================================
def _collector(doc, view):
    """FilteredElementCollector over hele doc (view=None) of een view."""
    if view is None:
        return FilteredElementCollector(doc)
    return FilteredElementCollector(doc, view.Id)


def get_image_typename(img):
    """Echte type-naam van een ImageType (zoals in het Dynamo-script)."""
    try:
        p = img.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString() or ""
    except Exception:
        pass
    return ""


def collect_images_to_delete(doc, keep_prefix):
    """ImageTypes die niet met de behoud-prefix beginnen (project-breed)."""
    images = FilteredElementCollector(doc).OfClass(ImageType).ToElements()
    to_delete = []
    kept = []
    for img in images:
        name = get_image_typename(img)
        if name.lower().startswith(keep_prefix.lower()):
            kept.append(name)
        else:
            to_delete.append(img)
    return to_delete, kept


def collect_filled_regions(doc, view):
    return list(_collector(doc, view).OfClass(FilledRegion).ToElements())


def collect_detail_lines(doc, view):
    """Detail curves (NewDetailCurve-resultaat)."""
    result = []
    for c in _collector(doc, view).OfClass(CurveElement).ToElements():
        try:
            if c.CurveElementType == CurveElementType.DetailCurve:
                result.append(c)
        except Exception:
            pass
    return result


def collect_text_notes(doc, view):
    return list(_collector(doc, view).OfClass(TextNote).ToElements())


def collect_by_category(doc, view, bic):
    """Element-instanties (geen types) in een categorie."""
    return list(
        _collector(doc, view)
        .OfCategory(bic)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def delete_elements(doc, elements):
    """Verwijder elementen, retourneer (aantal_ok, aantal_fout).

    Reeds-verwijderde id's (bv. group-members) tellen niet als fout.
    """
    deleted = 0
    failed = 0
    for el in elements:
        try:
            eid = el.Id
            if doc.GetElement(eid) is None:
                continue
            doc.Delete(eid)
            deleted += 1
        except Exception:
            failed += 1
    return deleted, failed


# =============================================================================
# CATEGORIE-DEFINITIE
# key, label, collect(doc, view), respecteert_scope
# =============================================================================
def _categories():
    return [
        ("images",
         "Rasterafbeeldingen (behoud '{0}')".format(KEEP_PREFIX),
         None, False),
        ("regions", "Filled regions", collect_filled_regions, True),
        ("lines", "Detail lijnen", collect_detail_lines, True),
        ("text", "Tekstnotities", collect_text_notes, True),
        ("detailgroups", "Detail groepen",
         lambda d, v: collect_by_category(
             d, v, BuiltInCategory.OST_IOSDetailGroups), True),
        ("detailcomp", "Detail componenten",
         lambda d, v: collect_by_category(
             d, v, BuiltInCategory.OST_DetailComponents), True),
    ]


# =============================================================================
# UI WINDOW
# =============================================================================
class OpschonenWindow(WPFWindow):
    """Opschoon-UI op basis van het 3BM WPF-template."""

    def __init__(self, doc):
        xaml_file = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        super(OpschonenWindow, self).__init__(
            xaml_file, "GIS2BIM Opschonen", width=460, height=560
        )
        self.doc = doc
        # checkbox per categorie-key
        self._chk = {
            "images": self.chk_images,
            "regions": self.chk_regions,
            "lines": self.chk_lines,
            "text": self.chk_text,
            "detailgroups": self.chk_detailgroups,
            "detailcomp": self.chk_detailcomp,
        }
        self.btn_cancel.Click += self.on_cancel
        self.btn_execute.Click += self.on_execute

    def on_cancel(self, sender, args):
        self.close_cancel()

    def _selected_keys(self):
        return [k for k, chk in self._chk.items() if chk.IsChecked]

    def _scope_view(self):
        """None = hele document, anders de actieve view."""
        if self.rb_view.IsChecked:
            return self.doc.ActiveView
        return None

    def on_execute(self, sender, args):
        keys = self._selected_keys()
        if not keys:
            self.show_warning("Selecteer minstens een categorie.",
                              "GIS2BIM Opschonen")
            return

        view = self._scope_view()
        cats = [c for c in _categories() if c[0] in keys]

        # Inventariseren
        found = {}
        kept_images = []
        for key, label, collect, respects_scope in cats:
            if key == "images":
                els, kept_images = collect_images_to_delete(
                    self.doc, KEEP_PREFIX)
            else:
                scope_view = view if respects_scope else None
                els = collect(self.doc, scope_view)
            found[key] = els

        total = sum(len(v) for v in found.values())
        if total == 0:
            self.show_info("Niets gevonden om te verwijderen.",
                           "GIS2BIM Opschonen")
            return

        scope_txt = "actieve view '{0}'".format(self.doc.ActiveView.Name) \
            if view is not None else "hele document"

        msg_lines = []
        for key, label, collect, respects_scope in cats:
            n = len(found.get(key, []))
            if key == "images":
                msg_lines.append("- {0} rasterafbeelding(en) "
                                 "({1} behouden via '{2}')".format(
                                     n, len(kept_images), KEEP_PREFIX))
            else:
                msg_lines.append("- {0} {1}".format(n, label.lower()))

        if not self.ask_confirm(
                "Scope: {0}\n\nHet volgende wordt VERWIJDERD:\n\n{1}\n\n"
                "Doorgaan?".format(scope_txt, "\n".join(msg_lines)),
                "GIS2BIM Opschonen - bevestigen"):
            return

        # Verwijderen
        results = {}
        try:
            with revit.Transaction("GIS2BIM Opschonen"):
                for key, label, collect, respects_scope in cats:
                    results[key] = delete_elements(self.doc, found.get(key, []))
        except Exception as e:
            log("Fout tijdens opschonen: {0}".format(e))
            log(traceback.format_exc())
            self.show_error("Fout tijdens opschonen:\n{0}".format(e),
                            "GIS2BIM Opschonen")
            return

        summary = []
        for key, label, collect, respects_scope in cats:
            ok, fail = results.get(key, (0, 0))
            summary.append("{0}: {1} verwijderd, {2} mislukt".format(
                label, ok, fail))

        log("Opschonen klaar ({0}): {1}".format(
            scope_txt, " | ".join(summary)))
        self.show_info("Opschonen voltooid ({0}).\n\n{1}".format(
            scope_txt, "\n".join(summary)), "GIS2BIM Opschonen")
        self.close_ok()


def main():
    doc = revit.doc
    try:
        window = OpschonenWindow(doc)
        window.show_dialog()
    except Exception as e:
        log("Opschonen kon niet starten: {0}".format(e))
        log(traceback.format_exc())
        forms.alert("Opschonen kon niet starten:\n{0}".format(e),
                    title="GIS2BIM Opschonen", warn_icon=True)


if __name__ == "__main__":
    main()
