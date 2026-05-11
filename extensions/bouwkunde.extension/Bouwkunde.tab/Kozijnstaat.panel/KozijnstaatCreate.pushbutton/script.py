# -*- coding: utf-8 -*-
"""Kozijnstaat - Create.

User selecteert een wand (in het project staan al voorbereide A0/A1/A2/A3
wanden voor schaal 1:20). Tool pakt unieke kozijntypes en plaatst ze op
de wand:
  - 500 mm horizontale tussenruimte tussen kozijnen
  - 2000 mm verticale ruimte boven het hoogste kozijn van een rij
    voordat de volgende rij begint
  - wraps automatisch naar nieuwe rij als horizontaal niet meer past

IronPython 2.7.
"""

__title__ = "Create\nKozijnstaat"
__author__ = "3BM Bouwkunde"
__doc__ = "Plaats unieke kozijntypes op een geselecteerde wand"

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


# Early-log: schrijft direct naar logbestand zonder afhankelijkheid
# van logger.py. Werkt zelfs als imports verderop falen.
_EARLY_LOG_DIR = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")),
    "3bm_exchange",
)
_EARLY_LOG_FILE = os.path.join(_EARLY_LOG_DIR, "kozijnstaat_debug.log")


def _early_log(msg):
    try:
        if not os.path.isdir(_EARLY_LOG_DIR):
            os.makedirs(_EARLY_LOG_DIR)
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = "[{0}] EARLY: {1}\n".format(ts, msg)
        f = open(_EARLY_LOG_FILE, "a")
        try:
            f.write(line)
        finally:
            f.close()
    except Exception:
        pass


def _early_log_exc(msg):
    try:
        tb = _tb_mod.format_exc()
    except Exception:
        tb = "<no traceback>"
    _early_log("{0}\n{1}".format(msg, tb))


_early_log("===== SCRIPT MODULE LOAD START =====")

try:
    from pyrevit import revit, forms, script
    _early_log("pyrevit imported")

    from Autodesk.Revit.DB import (
        Transaction,
        XYZ,
        BuiltInParameter,
        StorageType,
    )
    from Autodesk.Revit.DB.Structure import StructuralType
    from Autodesk.Revit.UI.Selection import ObjectType
    from Autodesk.Revit.Exceptions import OperationCanceledException
    _early_log("Revit API imported")

    from kozijnstaat.config import load_config
    _early_log("config imported")
    from kozijnstaat.family_collector import (
        collect_window_symbols,
        get_symbol_width_mm,
        get_symbol_height_mm,
        _read_string_param,
        find_first_instance_per_symbol,
        find_workset_id_by_name,
        _id_int,
    )
    _early_log("family_collector imported")
    from kozijnstaat.grid_layout import (
        compute_wall_fill_layout,
        compute_points_from_wall_fill,
    )
    _early_log("grid_layout imported")
    from kozijnstaat import logger as klog
    _early_log("kozijnstaat.logger imported")
except Exception:
    _early_log_exc("IMPORT FAILED")
    raise


HORIZONTAL_SPACING_MM = 500.0
ROW_SPACING_MM = 2000.0


def _set_instance_workset(inst, workset_id_int, label):
    """Zet de workset van een instance via ELEM_PARTITION_PARAM.

    No-op als de param read-only is (bv. niet-workshared doc) of als
    workset_id_int None is. Logt resultaat naar klog.
    """
    if workset_id_int is None:
        return False
    try:
        p = inst.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    except Exception as ex:
        try:
            klog.warn(u"workset param lookup {0} failed: {1}".format(label, ex))
        except Exception:
            pass
        return False
    if p is None:
        return False
    if p.IsReadOnly:
        try:
            klog.warn(u"workset param read-only voor {0}".format(label))
        except Exception:
            pass
        return False
    try:
        p.Set(int(workset_id_int))
        try:
            klog.info(
                u"workset op {0} gezet -> id={1}".format(label, workset_id_int)
            )
        except Exception:
            pass
        return True
    except Exception as ex:
        try:
            klog.warn(u"workset set {0} failed: {1}".format(label, ex))
        except Exception:
            pass
        return False


def _copy_instance_params(src_inst, dst_inst, param_names, label):
    """Kopieer instance-param-waardes van src naar dst.

    Slaat read-only / ontbrekende params stil over. Logt elke
    poging+resultaat naar klog zodat afwijkingen traceerbaar zijn.

    Returns:
        list[(name, ok, src_val_str, msg)] — per param een tuple voor
        output-tabel in run().
    """
    rows = []
    for name in param_names:
        try:
            p_src = src_inst.LookupParameter(name)
        except Exception as ex:
            rows.append((name, False, u"", u"src lookup: {0}".format(ex)))
            continue
        try:
            p_dst = dst_inst.LookupParameter(name)
        except Exception as ex:
            rows.append((name, False, u"", u"dst lookup: {0}".format(ex)))
            continue

        if p_src is None:
            rows.append((name, False, u"", u"src param niet aanwezig"))
            continue
        if p_dst is None:
            rows.append((name, False, u"", u"dst param niet aanwezig"))
            continue
        if not p_src.HasValue:
            rows.append((name, False, u"", u"src leeg"))
            continue
        if p_dst.IsReadOnly:
            rows.append((name, False, u"", u"dst read-only"))
            continue

        try:
            st = p_src.StorageType
        except Exception:
            st = None

        src_val_str = u""
        try:
            v = p_src.AsValueString()
            if v:
                src_val_str = v
        except Exception:
            pass

        try:
            if st == StorageType.String:
                v = p_src.AsString()
                if v is None:
                    v = u""
                p_dst.Set(v)
                if not src_val_str:
                    src_val_str = v
            elif st == StorageType.Double:
                p_dst.Set(p_src.AsDouble())
            elif st == StorageType.Integer:
                p_dst.Set(p_src.AsInteger())
            elif st == StorageType.ElementId:
                p_dst.Set(p_src.AsElementId())
            else:
                rows.append((name, False, src_val_str, u"unsupported storage"))
                continue
            rows.append((name, True, src_val_str, u""))
            try:
                klog.info(
                    u"copy-param '{0}' on {1}: '{2}' OK".format(
                        name, label, src_val_str,
                    )
                )
            except Exception:
                pass
        except Exception as ex:
            rows.append((name, False, src_val_str, u"set: {0}".format(ex)))
            try:
                klog.warn(
                    u"copy-param '{0}' on {1} FAILED: {2}".format(
                        name, label, ex,
                    )
                )
            except Exception:
                pass
    return rows


def _get_name(element):
    """Robuuste Name-getter — werkt rond IronPython 2.7 / Revit quirks.

    Probeert in volgorde:
      1. .NET reflection via GetType().GetProperty("Name")
      2. Direct .Name attribute access
      3. BuiltInParameter SYMBOL_NAME_PARAM / ALL_MODEL_TYPE_NAME

    Returns:
        str (kan leeg zijn als alle methodes falen)
    """
    if element is None:
        return ""
    try:
        clr_type = element.GetType()
        prop = clr_type.GetProperty("Name")
        if prop is not None:
            v = prop.GetValue(element, None)
            if v is not None:
                return v
    except Exception:
        pass
    try:
        v = element.Name
        if v is not None:
            return v
    except Exception:
        pass
    try:
        for bip in (
            BuiltInParameter.SYMBOL_NAME_PARAM,
            BuiltInParameter.ALL_MODEL_TYPE_NAME,
        ):
            try:
                p = element.get_Parameter(bip)
                if p and p.HasValue:
                    v = p.AsString()
                    if v:
                        return v
            except Exception:
                continue
    except Exception:
        pass
    return ""


def _id_str(element):
    """Return ElementId als string, robuust over Revit-versies."""
    if element is None:
        return "None"
    try:
        eid = element.Id
    except Exception:
        return "?"
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(eid, attr)
            if callable(v):
                v = v()
            return str(v)
        except Exception:
            continue
    try:
        return str(eid)
    except Exception:
        return "?"


def _pick_wall(uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, "Selecteer host-wand voor kozijnstaat"
        )
        return uidoc.Document.GetElement(ref.ElementId)
    except OperationCanceledException:
        return None


def _wall_geometry(wall):
    """origin (XYZ), u_dir, length_ft, height_ft.

    u_dir wijst altijd 'naar rechts' als je vanaf de exterior-zijde naar
    de wand kijkt — zodat origin de linker-onder-hoek is, ongeacht hoe
    de wand getekend is. Detectie via wall.Orientation (exterior normal)
    en het kruisproduct u × Z.
    """
    curve = wall.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    length = p0.DistanceTo(p1)
    u_dir = XYZ(
        (p1.X - p0.X) / length,
        (p1.Y - p0.Y) / length,
        0.0,
    )

    # u × Z geeft de horizontale vector loodrecht op u (in XY-vlak).
    # Als die vector dezelfde kant op wijst als de exterior-normaal,
    # is u 'rechts' bij kijk van buiten — anders moeten we omdraaien.
    try:
        n = wall.Orientation
        u_cross_z_x = u_dir.Y
        u_cross_z_y = -u_dir.X
        dot = u_cross_z_x * n.X + u_cross_z_y * n.Y
        try:
            klog.info(
                u"wall orientation n=({0:.2f},{1:.2f},{2:.2f}) "
                u"u=({3:.2f},{4:.2f}) u_cross_z=({5:.2f},{6:.2f}) "
                u"dot={7:.2f}".format(
                    n.X, n.Y, n.Z,
                    u_dir.X, u_dir.Y,
                    u_cross_z_x, u_cross_z_y, dot,
                )
            )
        except Exception:
            pass
        if dot < 0:
            u_dir = XYZ(-u_dir.X, -u_dir.Y, 0.0)
            origin = p1
            try:
                klog.info(u"  -> flipped u_dir, origin = p1")
            except Exception:
                pass
        else:
            origin = p0
    except Exception:
        origin = p0

    height_ft = 10.0
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        if p and p.HasValue:
            height_ft = p.AsDouble()
    except Exception:
        pass
    return (origin, u_dir, length, height_ft)


def run():
    _early_log("run() entered")
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()
    output.print_md("## Kozijnstaat - Create")
    _early_log("got doc/uidoc/output")

    klog.reset(header="KozijnstaatCreate run")
    klog.info(u"=== START ===")
    _early_log("klog.reset done")
    output.print_md(
        "*Debug-log: `{0}`*".format(klog.get_log_path())
    )

    cfg = load_config()
    kozijn_family = cfg.get("kozijn_family", "3BM_kozijn")
    merk_param = cfg.get("param_merk", "merk")
    instance_params_to_copy = cfg.get("instance_params_to_copy", []) or []
    only_placed = bool(cfg.get("only_placed_types", True))
    workset_name = cfg.get("kozijnstaat_workset_name", "") or ""
    target_workset_id = None
    if workset_name:
        target_workset_id = find_workset_id_by_name(doc, workset_name)
        if target_workset_id is None:
            klog.warn(
                u"workset '{0}' niet gevonden — canvas-instances "
                u"blijven op actieve workset".format(workset_name)
            )
        else:
            klog.info(
                u"target workset '{0}' id={1}".format(
                    workset_name, target_workset_id,
                )
            )
    klog.info(
        u"config kozijn_family='{0}' merk_param='{1}' "
        u"instance_params_to_copy={2} workset='{3}'".format(
            kozijn_family, merk_param, instance_params_to_copy,
            workset_name,
        )
    )

    # 1. Verzamel unieke types (gesorteerd op merk-param, dan family+type)
    symbols = collect_window_symbols(
        doc, name_contains=kozijn_family, merk_param=merk_param,
    )
    if not symbols:
        forms.alert(
            "Geen kozijn FamilyTypes gevonden met filter '{0}'."
            .format(kozijn_family),
            title="Geen types",
        )
        return

    # 1b. Filter op "alleen types met >=1 model-instance" (excl. canvas-
    # workset). Tegelijk verzamelen we de eerste model-instance per
    # symbol — die hergebruiken we later voor instance-param copy.
    first_inst_per_symbol = find_first_instance_per_symbol(
        doc, symbols, exclude_workset_id=target_workset_id,
    )
    if only_placed:
        before = len(symbols)
        symbols = [
            s for s in symbols if _id_int(s) in first_inst_per_symbol
        ]
        skipped = before - len(symbols)
        if skipped > 0:
            klog.info(
                u"only_placed_types: {0} types overgeslagen "
                u"(geen model-instance buiten kozijnstaat-workset)"
                .format(skipped)
            )
            output.print_md(
                "Overgeslagen: **{0}** types zonder model-instance "
                "(`only_placed_types=True`)".format(skipped)
            )
        if not symbols:
            forms.alert(
                "Geen kozijn-types met model-instances gevonden.\n\n"
                "Zet 'only_placed_types' op False in user_config.json "
                "om alle geladen types te plaatsen.",
                title="Geen geplaatste types",
            )
            return

    n = len(symbols)
    output.print_md("Gevonden: **{0}** unieke kozijntypes".format(n))
    klog.info(u"collected {0} kozijn-types".format(n))

    # Activeer alle symbols vóór dimensies uitlezen, zodat
    # computed/derived parameters correct gerapporteerd worden.
    tx = Transaction(doc, "Kozijnstaat - Activate symbols")
    tx.Start()
    try:
        for s in symbols:
            if not s.IsActive:
                s.Activate()
        doc.Regenerate()
        tx.Commit()
    except Exception:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()

    widths_mm = []
    heights_mm = []
    fallback_used = False
    output.print_md("### Gedetecteerde afmetingen (gesorteerd op merk)")
    output.print_md("| # | Merk | Family | Type | Breedte | Hoogte |")
    output.print_md("|---|---|---|---|---|---|")
    for idx, s in enumerate(symbols):
        try:
            fam = s.Family
        except Exception:
            fam = None
        fam_name = _get_name(fam) or "?"
        type_name = _get_name(s) or "?"
        merk_val = _read_string_param(s, merk_param) if merk_param else u""
        merk_disp = merk_val if merk_val else u"*(geen)*"
        w_raw = get_symbol_width_mm(s)
        h_raw = get_symbol_height_mm(s)
        w_eff = w_raw if w_raw > 0 else 1000.0
        h_eff = h_raw if h_raw > 0 else 2000.0
        widths_mm.append(w_eff)
        heights_mm.append(h_eff)
        flag_w = "" if w_raw > 0 else " *(fallback)*"
        flag_h = "" if h_raw > 0 else " *(fallback)*"
        if w_raw <= 0 or h_raw <= 0:
            fallback_used = True
        output.print_md(
            "| {0} | {1} | {2} | {3} | {4:.0f}{5} | {6:.0f}{7} |".format(
                idx, merk_disp, fam_name, type_name,
                w_eff, flag_w, h_eff, flag_h,
            )
        )

    if fallback_used:
        output.print_md(
            "**LET OP:** een of meer breedtes/hoogtes konden niet uit "
            "family-parameters gelezen worden — fallback gebruikt. "
            "Voeg de juiste param-naam toe in "
            "`lib/kozijnstaat/family_collector.py` "
            "(zoek naar `get_symbol_width_mm`)."
        )

    # 2. Wand selecteren
    host_wall = _pick_wall(uidoc)
    if host_wall is None:
        output.print_md("*Geen wand geselecteerd.*")
        return

    origin, u_dir, wall_len_ft, wall_h_ft = _wall_geometry(host_wall)
    wall_len_mm = wall_len_ft * 304.8
    wall_h_mm = wall_h_ft * 304.8
    v_dir = XYZ(0.0, 0.0, 1.0)

    output.print_md(
        "Wand: **{0:.0f} x {1:.0f} mm**".format(wall_len_mm, wall_h_mm)
    )
    klog.info(
        u"selected wall id={0} length={1:.0f}mm height={2:.0f}mm "
        u"origin=({3:.2f},{4:.2f},{5:.2f})ft u_dir=({6:.2f},{7:.2f},{8:.2f})"
        .format(
            _id_str(host_wall),
            wall_len_mm, wall_h_mm,
            origin.X, origin.Y, origin.Z,
            u_dir.X, u_dir.Y, u_dir.Z,
        )
    )

    # 3. Pack kozijnen sequentieel
    layout = compute_wall_fill_layout(
        widths_mm, heights_mm,
        wall_length_mm=wall_len_mm,
        wall_height_mm=wall_h_mm,
        horizontal_spacing_mm=HORIZONTAL_SPACING_MM,
        row_spacing_mm=ROW_SPACING_MM,
    )
    if layout is None:
        forms.alert("Layout-berekening faalde.", title="Fout")
        return

    placed_count = layout["placed_count"]
    overflow = layout["overflow_count"]
    n_rows = len(layout["rows"])

    output.print_md(
        "Layout: **{0} rijen**, gebruikt **{1:.0f} x {2:.0f} mm** "
        "(passend: {3} van {4})".format(
            n_rows,
            layout["used_w_mm"], layout["used_h_mm"],
            placed_count, n,
        )
    )
    klog.info(
        u"layout: rows={0} used={1:.0f}x{2:.0f}mm placed={3} overflow={4}"
        .format(n_rows, layout["used_w_mm"], layout["used_h_mm"],
                placed_count, overflow)
    )
    for r_idx, row in enumerate(layout["rows"]):
        for item_idx, x_off, row_bot, w, h in row:
            klog.info(
                u"  row{0}/idx{1}: x_off={2:.0f} row_bot={3:.0f} "
                u"w={4:.0f} h={5:.0f} center_x={6:.0f}".format(
                    r_idx, item_idx, x_off, row_bot, w, h,
                    x_off + w / 2.0,
                )
            )

    if overflow > 0:
        ok = forms.alert(
            "{0} van de {1} kozijnen passen niet op deze wand.\n\n"
            "Wand: {2:.0f} x {3:.0f} mm\n"
            "Doorgaan met alleen de eerste {4}?".format(
                overflow, n,
                wall_len_mm, wall_h_mm,
                placed_count,
            ),
            yes=True, no=True,
        )
        if not ok:
            return

    if placed_count == 0:
        forms.alert(
            "Geen enkel kozijn past op deze wand.",
            title="Wand te klein",
        )
        return

    # 4. Bereken plaatsingspunten — origin = wall start
    points = compute_points_from_wall_fill(
        layout, origin, u_dir, v_dir,
    )

    # 4b. first_inst_per_symbol is al opgehaald in stap 1b voor de
    # only_placed_types filter — hergebruiken voor instance-param copy.
    # Als de copy-feature uit staat, leeg maken.
    if not instance_params_to_copy:
        first_inst_per_symbol = {}
    else:
        klog.info(
            u"first_inst_per_symbol available for copy: {0} types "
            u"(excl. workset id={1})".format(
                len(first_inst_per_symbol), target_workset_id,
            )
        )

    # 5. Plaats kozijnen
    tx = Transaction(doc, "Kozijnstaat - Plaats kozijnen")
    tx.Start()
    try:
        placed = 0
        failed = 0
        for i, symbol in enumerate(symbols):
            if i >= len(points):
                break
            pt = points[i]
            if pt is None:
                continue
            # Naam VOORAF vastleggen — symbol-reference kan onbruikbaar
            # worden na een mislukte NewFamilyInstance call
            sym_name = _get_name(symbol) or "<no-name>"
            if not symbol.IsActive:
                symbol.Activate()
                doc.Regenerate()
            try:
                inst = doc.Create.NewFamilyInstance(
                    pt, symbol, host_wall,
                    StructuralType.NonStructural,
                )
                placed += 1
                klog.info(
                    u"placed[{0}] {1} at xyz=({2:.2f},{3:.2f},{4:.2f})ft "
                    u"= ({5:.0f},{6:.0f},{7:.0f})mm  inst_id={8}".format(
                        i, sym_name,
                        pt.X, pt.Y, pt.Z,
                        pt.X * 304.8, pt.Y * 304.8, pt.Z * 304.8,
                        _id_str(inst),
                    )
                )

                # Optie A: kopieer instance-params van eerste model-instance
                if instance_params_to_copy:
                    sid = _id_int(symbol)
                    src_inst = first_inst_per_symbol.get(sid)
                    if src_inst is None:
                        klog.info(
                            u"  geen model-instance voor type {0}, "
                            u"instance-params blijven default".format(sym_name)
                        )
                    else:
                        _copy_instance_params(
                            src_inst, inst,
                            instance_params_to_copy, sym_name,
                        )

                # Zet workset op canvas-instance (na placement, voor regen)
                if target_workset_id is not None:
                    _set_instance_workset(inst, target_workset_id, sym_name)
            except Exception as ex:
                try:
                    ex_msg = str(ex)
                except Exception:
                    ex_msg = "<unprintable>"
                try:
                    ex_type = type(ex).__name__
                except Exception:
                    ex_type = "Exception"
                output.print_md(
                    "  - *FOUT type '{0}': {1}: {2}*"
                    .format(sym_name, ex_type, ex_msg)
                )
                klog.exc(
                    u"placement failed for [{0}] {1}: {2}: {3}".format(
                        i, sym_name, ex_type, ex_msg,
                    )
                )
                failed += 1
        tx.Commit()
    except Exception as ex:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        import traceback as _tb
        tb_text = _tb.format_exc()
        klog.exc(u"placement transaction failed")
        forms.alert(
            "Kozijnen plaatsen mislukt:\n{0}: {1}\n\n{2}".format(
                type(ex).__name__, ex, tb_text,
            ),
            title="Fout",
        )
        return

    forms.alert(
        "Klaar.\n"
        "Geplaatst: {0} (mislukt: {1}, overflow: {2})\n"
        "Layout: {3} rijen".format(
            placed, failed, overflow, n_rows,
        ),
        title="Create Kozijnstaat",
    )


if __name__ == "__main__":
    _early_log("__main__ block entered")
    if revit.doc is None:
        _early_log("revit.doc is None")
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        try:
            run()
            _early_log("run() returned normally")
        except Exception as _ex:
            _early_log_exc("run() raised exception")
            try:
                _tb_text = _tb_mod.format_exc()
            except Exception:
                _tb_text = "<no traceback>"
            try:
                klog.exc(u"run() crashed")
            except Exception:
                pass
            try:
                _name = type(_ex).__name__
            except Exception:
                _name = "Exception"
            try:
                _msg = str(_ex)
            except Exception:
                _msg = "<unprintable>"
            forms.alert(
                "Onverwachte fout in KozijnstaatCreate:\n"
                "{0}: {1}\n\n{2}\n\n"
                "Log: {3}".format(
                    _name, _msg, _tb_text, _EARLY_LOG_FILE,
                ),
                title="Fout",
            )
