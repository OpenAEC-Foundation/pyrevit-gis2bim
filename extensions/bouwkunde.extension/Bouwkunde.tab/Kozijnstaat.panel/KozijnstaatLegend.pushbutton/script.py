# -*- coding: utf-8 -*-
"""Kozijnstaat - Legend (POC).

Maakt voor het eerste kozijn-type uit de kozijnstaat (na merk-sort)
een Legend-view met Plan + Front + Right op schaal 1:20.

Prerequisite: het document bevat een handmatig aangemaakte template
Legend-view (naam via Config: `template_legend_name`) met daarin
minstens 1 LegendComponent. Reden: Revit API kan Legend-views en
LegendComponents niet from-scratch aanmaken.

IronPython 2.7.
"""

__title__ = "Legend\n(POC)"
__author__ = "3BM Bouwkunde"
__doc__ = "POC: Legend-view met Plan/Front/Right voor 1 kozijn (1:20)"

import os
import sys
import datetime
import traceback as _tb_mod

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)


# --- Pre-import logger (schrijft VOOR de imports lukken) ------------------
#
# Reden: als een import faalt komen we nooit bij klog.make_logger() aan toe.
# Deze early logger schrijft direct naar dezelfde file zodat de user altijd
# kan zien wat er ging gebeuren — zelfs als de pyRevit-bundle al crasht op
# parse/import niveau.

_EARLY_LOG_DIR = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")),
    "3bm_exchange",
)
_EARLY_LOG_FILE = os.path.join(_EARLY_LOG_DIR, "kozijnstaat_legend.log")


def _early_log(msg):
    try:
        if not os.path.isdir(_EARLY_LOG_DIR):
            os.makedirs(_EARLY_LOG_DIR)
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f = open(_EARLY_LOG_FILE, "a")
        try:
            try:
                f.write(u"[{0}] EARLY: {1}\n".format(ts, msg).encode("utf-8"))
            except Exception:
                f.write("[{0}] EARLY: {1}\n".format(ts, msg))
            try:
                f.flush()
            except Exception:
                pass
        finally:
            f.close()
    except Exception:
        pass


_early_log("=" * 60)
_early_log("LEGEND script-load start")
_early_log("SCRIPT_DIR={0}".format(SCRIPT_DIR))
_early_log("EXTENSION_DIR={0}".format(EXTENSION_DIR))
_early_log("LIB_DIR={0} exists={1}".format(LIB_DIR, os.path.isdir(LIB_DIR)))
_early_log("python={0}".format(sys.version.replace("\n", " ")))

try:
    _early_log("import pyrevit...")
    from pyrevit import revit, forms, script
    _early_log("import pyrevit OK")

    _early_log("import Autodesk.Revit.DB.Transaction...")
    from Autodesk.Revit.DB import Transaction
    _early_log("import Autodesk OK")

    _early_log("import kozijnstaat.config...")
    from kozijnstaat.config import load_config
    _early_log("import config OK")

    _early_log("import kozijnstaat.family_collector...")
    from kozijnstaat.family_collector import (
        collect_window_symbols,
        _read_string_param,
    )
    _early_log("import family_collector OK")

    _early_log("import kozijnstaat.legend_builder...")
    from kozijnstaat.legend_builder import (
        find_template_legend,
        find_first_legend_component,
        build_kozijn_legend,
        list_legend_views,
    )
    _early_log("import legend_builder OK")

    _early_log("import kozijnstaat.logger...")
    from kozijnstaat import logger as klog
    _early_log("import logger OK")
except Exception:
    try:
        _early_log("IMPORT FAILED\n{0}".format(_tb_mod.format_exc()))
    except Exception:
        pass
    raise


# Dedicated logger voor deze tool — append-only, eigen file
log = klog.make_logger("kozijnstaat_legend.log")


def _safe_name(element):
    """Robuuste Name-getter via .NET reflection."""
    if element is None:
        return u""
    try:
        prop = element.GetType().GetProperty("Name")
        if prop is not None:
            v = prop.GetValue(element, None)
            if v is not None:
                return v
    except Exception:
        pass
    try:
        return element.Name or u""
    except Exception:
        return u""


def _id_int(eid):
    if eid is None:
        return -1
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(eid, attr)
            if callable(v):
                v = v()
            return int(v)
        except Exception:
            continue
    return -1


def run(profile="kozijn"):
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()

    log.session("KozijnstaatLegend run")
    log.info(u"step 1: entry — doc.Title='{0}' active_view='{1}'".format(
        getattr(doc, "Title", "?"),
        _safe_name(getattr(doc, "ActiveView", None)),
    ))
    output.print_md(
        "*Debug-log: `{0}`*".format(log.get_log_path())
    )

    log.info(u"step 2: load_config()")
    try:
        cfg = load_config(profile)
        log.info(u"step 2 OK — config keys={0}".format(sorted(cfg.keys())))
    except Exception:
        log.exc(u"step 2 FAILED — load_config")
        forms.alert("Config laden mislukt — zie log.", title="Fout")
        return

    tool_label = cfg.get("tool_label", "Kozijnstaat")
    category = cfg.get("element_category")
    output.print_md("## {0} - Legend (POC)".format(tool_label))

    kozijn_family = cfg.get("kozijn_family", "3BM_kozijn")
    merk_param = cfg.get("param_merk", "merk")
    template_name = cfg.get("template_legend_name", "TEMPLATE_kozijn_legend")
    scale = int(cfg.get("legend_scale", 20))
    spacing_ft = float(cfg.get("legend_spacing_ft", 3.0))

    log.info(
        u"step 3: resolved config — kozijn_family='{0}' merk_param='{1}' "
        u"template='{2}' scale=1:{3} spacing={4}ft".format(
            kozijn_family, merk_param, template_name, scale, spacing_ft,
        )
    )

    log.info(u"step 4: collect_window_symbols(name_contains='{0}')".format(
        kozijn_family,
    ))
    try:
        symbols = collect_window_symbols(
            doc, name_contains=kozijn_family, merk_param=merk_param,
            category=category,
        )
        log.info(u"step 4 OK — found {0} symbols".format(len(symbols)))
        for i, s in enumerate(symbols[:10]):
            log.info(u"  symbol[{0}] id={1} name='{2}'".format(
                i, _id_int(s.Id), _safe_name(s),
            ))
    except Exception:
        log.exc(u"step 4 FAILED — collect_window_symbols")
        forms.alert(
            "Window-symbols verzamelen mislukt — zie log:\n{0}".format(
                log.get_log_path(),
            ),
            title="Fout",
        )
        return

    if not symbols:
        log.warn(u"step 4: geen symbols — abort")
        forms.alert(
            "Geen kozijn-types gevonden met filter '{0}'."
            .format(kozijn_family),
            title="Geen types",
        )
        return

    symbol = symbols[0]
    sym_name = _safe_name(symbol)
    if not sym_name:
        # FamilySymbol.Name kan leeg returnen via Element.Name in IP 2.7;
        # val terug op Family-name + symbol ElementId voor herkenbaarheid.
        try:
            fam_name = _safe_name(getattr(symbol, "Family", None))
        except Exception:
            fam_name = u""
        sym_name = fam_name or u"kozijn-{0}".format(_id_int(symbol.Id))
    try:
        merk_val = _read_string_param(symbol, merk_param) if merk_param else u""
    except Exception:
        log.exc(u"step 5: _read_string_param failed — fallback to empty")
        merk_val = u""
    legend_name_label = merk_val if merk_val else sym_name

    output.print_md(
        "Geselecteerd type: **{0}** (merk: {1})".format(
            sym_name, merk_val if merk_val else "*(geen)*",
        )
    )
    log.info(
        u"step 5: selected symbol id={0} name='{1}' merk='{2}'".format(
            _id_int(symbol.Id), sym_name, merk_val,
        )
    )

    log.info(u"step 6: find_template_legend('{0}')".format(template_name))
    try:
        template = find_template_legend(doc, template_name)
    except Exception:
        log.exc(u"step 6 FAILED — find_template_legend")
        forms.alert(
            "Template zoeken mislukt — zie log:\n{0}".format(
                log.get_log_path(),
            ),
            title="Fout",
        )
        return

    if template is None:
        log.warn(u"step 6: template '{0}' niet gevonden".format(template_name))
        try:
            available = [_safe_name(v) for v in list_legend_views(doc)]
        except Exception:
            log.exc(u"step 6: list_legend_views failed")
            available = []
        log.info(u"step 6: available legend views={0}".format(available))
        avail_md = (
            "\n- " + "\n- ".join(available)
            if available else " *(geen Legend-views in document)*"
        )
        msg = (
            "Template Legend-view '{0}' niet gevonden.\n\n"
            "Aanwezige Legend-views:{1}\n\n"
            "Aanmaken:\n"
            " 1. Maak in Revit handmatig een Legend-view "
            "(View tab > Legends > Legend).\n"
            " 2. Plaats er minstens 1 LegendComponent in "
            "(welk type dan ook).\n"
            " 3. Hernoem die view naar '{0}' "
            "(of pas Config aan).".format(template_name, avail_md)
        )
        forms.alert(msg, title="Template ontbreekt")
        return

    log.info(u"step 6 OK — template id={0} name='{1}'".format(
        _id_int(template.Id), _safe_name(template),
    ))

    log.info(u"step 7: find_first_legend_component(template)")
    try:
        source_comp = find_first_legend_component(doc, template)
    except Exception:
        log.exc(u"step 7 FAILED — find_first_legend_component")
        forms.alert(
            "LegendComponent zoeken mislukt — zie log:\n{0}".format(
                log.get_log_path(),
            ),
            title="Fout",
        )
        return

    if source_comp is None:
        log.warn(u"step 7: geen LegendComponent in template")
        forms.alert(
            "Template-legend '{0}' bevat geen LegendComponent.\n\n"
            "Plaats er handmatig minstens 1 component in (welk type "
            "dan ook), de tool kopieert hem dan en zet hem om naar het "
            "juiste kozijn-type.".format(template_name),
            title="Template leeg",
        )
        return

    log.info(u"step 7 OK — source_comp id={0} category='{1}'".format(
        _id_int(source_comp.Id),
        _safe_name(getattr(source_comp, "Category", None)),
    ))

    new_name = u"Kozijn legend - {0}".format(legend_name_label)
    log.info(u"step 8: call build_kozijn_legend — new_name='{0}'".format(
        new_name,
    ))

    # GEEN outer Transaction — de builder maakt zelf separate
    # transactions per stap zodat een "internal error" tijdens één
    # view-placement niet alle andere placements meeneemt.
    try:
        result = build_kozijn_legend(
            doc, symbol, template, source_comp,
            new_name=new_name,
            scale=scale,
            spacing_ft=spacing_ft,
            uidoc=uidoc,
        )
        log.info(u"step 8 OK — builder returned")
    except Exception as ex:
        log.exc(u"step 8 FAILED — build_kozijn_legend raised")
        forms.alert(
            "Legend bouwen mislukt:\n{0}: {1}\n\nLog: {2}".format(
                type(ex).__name__, ex, log.get_log_path(),
            ),
            title="Fout",
        )
        return

    placed = result.get("placed", [])
    skipped = result.get("skipped", [])
    new_legend = result.get("new_legend")
    log.info(u"step 9: result — placed={0} skipped={1} new_legend_id={2}".format(
        len(placed), skipped, _id_int(getattr(new_legend, "Id", None)),
    ))

    output.print_md(
        "Nieuwe Legend-view: **{0}**".format(_safe_name(new_legend))
    )
    output.print_md(
        "Geplaatst: **{0}/3** views ({1})".format(
            len(placed),
            ", ".join(["Plan", "Front", "Right"]),
        )
    )
    if skipped:
        output.print_md(
            "*Skipped: {0}*".format(", ".join(skipped))
        )

    try:
        uidoc.ActiveView = new_legend
        log.info(u"step 10: activated new view")
    except Exception as ex:
        log.warn(u"step 10: kon nieuwe view niet activeren: {0}".format(ex))

    forms.alert(
        "Legend POC klaar.\n"
        "View: {0}\n"
        "Geplaatst: {1}/3 (Plan/Front/Right)\n"
        "Schaal: 1:{2}\n\n"
        "LET OP: als de view-modes niet kloppen "
        "(bv. Plan staat als Front), pas LEGEND_VIEW_* "
        "integers aan in lib/kozijnstaat/legend_builder.py.\n\n"
        "Log: {3}".format(
            new_name, len(placed), scale, log.get_log_path(),
        ),
        title="Legend POC",
    )


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        try:
            run()
        except Exception as _ex:
            try:
                _early_log("run() raised\n{0}".format(_tb_mod.format_exc()))
            except Exception:
                pass
            try:
                log.exc(u"run() crashed")
            except Exception:
                pass
            try:
                _tb_text = _tb_mod.format_exc()
            except Exception:
                _tb_text = "<no traceback>"
            forms.alert(
                "Onverwachte fout in KozijnstaatLegend:\n"
                "{0}: {1}\n\n{2}\n\nLog: {3}".format(
                    type(_ex).__name__, _ex, _tb_text, _EARLY_LOG_FILE,
                ),
                title="Fout",
            )
