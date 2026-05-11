# -*- coding: utf-8 -*-
"""Kozijnstaat - Rename kozijn types op basis van breedte/hoogte params.

Updatet het `_b<width>mm_h<height>mm` deel van de FamilySymbol-naam
zodat het exact de huidige `kozijn_breedte` / `kozijn_hoogte` waardes
weerspiegelt. Prefix (bv. H01_) en suffix (bv. _WBDBO30, ' 2')
blijven behouden.

IronPython 2.7.
"""

__title__ = "Rename\nKozijnen"
__author__ = "3BM Bouwkunde"
__doc__ = "Hernoem kozijn-types obv kozijn_breedte/kozijn_hoogte parameters"

import os
import sys
import re

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import Transaction, BuiltInParameter

from kozijnstaat.config import load_config
from kozijnstaat.family_collector import (
    collect_window_symbols,
    get_symbol_width_mm,
    get_symbol_height_mm,
)


# `^(.*?)(_b\d+mm_h\d+mm)(.*)$` — non-greedy prefix zodat eerste match
# wint bij meerdere occurrences (zou niet moeten voorkomen).
DIM_RE = re.compile(r"^(.*?)(_b\d+mm_h\d+mm)(.*)$")


def _sym_name(symbol):
    """Robuust de FamilySymbol-naam via SYMBOL_NAME_PARAM (Revit 2025
    quirk waarbij .Name reflection-access faalt)."""
    if symbol is None:
        return u""
    try:
        p = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString() or u""
    except Exception:
        pass
    try:
        return symbol.Name or u""
    except Exception:
        return u""


def _build_new_name(old_name, w_mm, h_mm):
    """Return new name with updated dimensions.

    Returns:
        (new_name, reason) — new_name=None bij skip; reason beschrijft
        waarom (voor de preview-tabel).
    """
    m = DIM_RE.match(old_name)
    if not m:
        return (None, u"geen `_b<n>mm_h<n>mm` pattern")
    prefix = m.group(1)
    suffix = m.group(3)
    new_dim = u"_b{0}mm_h{1}mm".format(
        int(round(w_mm)), int(round(h_mm)),
    )
    new_name = prefix + new_dim + suffix
    if new_name == old_name:
        return (None, u"naam is al correct")
    return (new_name, u"")


def run():
    doc = revit.doc
    output = script.get_output()
    output.print_md("## Kozijnstaat - Rename Types")

    cfg = load_config()
    kozijn_family = cfg.get("kozijn_family", "31_kozijn")
    output.print_md(
        "Filter: family-naam bevat **'{0}'**".format(kozijn_family)
    )

    # Geen sort-param: pure collect, geen volgorde-eisen voor rename.
    symbols = collect_window_symbols(
        doc, name_contains=kozijn_family, merk_param=u"",
    )
    if not symbols:
        forms.alert(
            "Geen kozijn types gevonden met filter '{0}'."
            .format(kozijn_family),
            title="Geen types",
        )
        return

    # 1. Bouw rename-plan
    plan = []  # list[(symbol, old, new, w, h, reason)]
    for s in symbols:
        old_name = _sym_name(s)
        w = get_symbol_width_mm(s)
        h = get_symbol_height_mm(s)
        if w <= 0.0 or h <= 0.0:
            plan.append((s, old_name, None, w, h,
                         u"geen breedte/hoogte param"))
            continue
        new_name, reason = _build_new_name(old_name, w, h)
        plan.append((s, old_name, new_name, w, h, reason))

    # 2. Preview-tabel
    output.print_md("### Preview")
    output.print_md(
        "| # | Oude naam | Nieuwe naam | breedte | hoogte | Status |"
    )
    output.print_md("|---|---|---|---:|---:|---|")
    to_rename = []
    for idx, (s, old, new, w, h, reason) in enumerate(plan, start=1):
        if new:
            status = u"rename"
            to_rename.append((s, old, new))
        else:
            status = u"skip: {0}".format(reason)
        output.print_md(
            u"| {0} | `{1}` | `{2}` | {3:.0f} | {4:.0f} | {5} |".format(
                idx, old, new or u"—", w, h, status,
            )
        )

    if not to_rename:
        forms.alert(
            "Geen types om te hernoemen (alle namen al correct of "
            "geen pattern).",
            title="Niets te doen",
        )
        return

    # 3. Conflict-check tegen ALLE bestaande type-namen (ook andere
    # categorien) — Revit eist unieke type-namen binnen een family.
    # Voor onze use-case is family-scope genoeg (alleen kozijn-types).
    existing = {}
    for s in symbols:
        nm = _sym_name(s)
        if nm:
            existing[nm] = s

    conflicts = []
    for s, old, new in to_rename:
        if new in existing and existing[new] is not s:
            conflicts.append((old, new))

    if conflicts:
        output.print_md("### Conflicts (rename geblokkeerd)")
        for old, new in conflicts:
            output.print_md(u"- `{0}` -> `{1}`".format(old, new))
        forms.alert(
            "{0} doel-namen bestaan al — kan niet hernoemen.\n"
            "Zie preview voor details.".format(len(conflicts)),
            title="Conflicts",
        )
        return

    # 4. Bevestiging
    ok = forms.alert(
        "{0} kozijn-types worden hernoemd. Doorgaan?\n\n"
        "Tip: undo werkt na de operatie (Ctrl+Z).".format(len(to_rename)),
        yes=True, no=True,
    )
    if not ok:
        return

    # 5. Rename in transaction. Volgorde: eerst types waarvan de doel-
    # naam niet conflicteert met een huidige naam. Iteratief: per pas
    # alleen types renamen die nu vrij zijn — voorkomt collision tijdens
    # de rename zelf (bv. swap A->B B->A).
    tx = Transaction(doc, "Rename kozijn types obv breedte/hoogte")
    tx.Start()
    renamed = 0
    failed = []
    try:
        pending = list(to_rename)
        guard = 0
        while pending and guard < 100:
            guard += 1
            current_names = set(_sym_name(s) for s in symbols)
            progress = False
            still = []
            for s, old, new in pending:
                # doelnaam mag niet bezet zijn door een ANDER symbol
                target_busy = (
                    new in current_names and _sym_name(s) != new
                )
                if target_busy:
                    still.append((s, old, new))
                    continue
                try:
                    s.Name = new
                    renamed += 1
                    progress = True
                except Exception as ex:
                    failed.append((old, new, u"{0}".format(ex)))
            if not progress:
                # Resterend zijn echte conflicts
                for s, old, new in still:
                    failed.append((old, new, u"conflict tijdens rename"))
                break
            pending = still
        tx.Commit()
    except Exception as ex:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        forms.alert(
            "Fout tijdens rename:\n{0}".format(ex),
            title="Fout",
        )
        return

    output.print_md("---")
    output.print_md(
        "**Hernoemd:** {0} / **Mislukt:** {1}".format(renamed, len(failed))
    )
    for old, new, reason in failed:
        output.print_md(
            u"- *`{0}` -> `{1}`: {2}*".format(old, new, reason)
        )
    forms.alert(
        "Klaar.\nHernoemd: {0}\nMislukt: {1}".format(renamed, len(failed)),
        title="Rename Kozijnen",
    )


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        run()
