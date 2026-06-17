# -*- coding: utf-8 -*-
"""Vlakaanzicht - section view loodrecht op gepickt vlak.

Workflow:
  1. Gebruiker pickt een face (op dak, wand, vloer, plafond, etc.).
  2. Script bepaalt de face-normaal en bouwt een orthonormaal frame:
     - BasisZ = face_normal (kijkrichting uit het vlak)
     - BasisY = "upslope": projectie van wereld-Z op het vlak
     - BasisX = BasisY x BasisZ
     Voor horizontale faces (vloer/plafond) wordt wereld-Y als
     up-referentie gebruikt.
  3. Bounding box van het host-element wordt naar de view-locale
     ruimte getransformeerd; min/max geeft de section box.
  4. ViewSection.CreateSection genereert de view; schaal + naam +
     view template (optioneel) worden gezet.

IronPython 2.7 / Revit API.
"""

import os
import sys
import math
import datetime
import traceback as _tb_mod

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)


# Debug-log
_EARLY_LOG_DIR = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")),
    "3bm_exchange",
)
_EARLY_LOG_FILE = os.path.join(_EARLY_LOG_DIR, "vlakaanzicht_debug.log")


def _log(msg):
    try:
        if not os.path.isdir(_EARLY_LOG_DIR):
            os.makedirs(_EARLY_LOG_DIR)
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f = open(_EARLY_LOG_FILE, "a")
        try:
            f.write("[{0}] {1}\n".format(ts, msg))
        finally:
            f.close()
    except Exception:
        pass


def _log_exc(msg):
    try:
        tb = _tb_mod.format_exc()
    except Exception:
        tb = "<no traceback>"
    _log("{0}\n{1}".format(msg, tb))


_log("===== SCRIPT LOAD =====")

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    Transaction,
    XYZ,
    UV,
    Transform,
    BoundingBoxXYZ,
    ElementId,
    FilteredElementCollector,
    BuiltInCategory,
    ViewSection,
    ViewFamilyType,
    ViewFamily,
    View,
    ViewType,
    PlanarFace,
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException


SCALE_DEFAULT = 20
MARGIN_FT = 3.28084          # ~1.0 m marge rond vlak
DEPTH_BEHIND_FT = 0.0164      # ~5 mm achter het vlak (dunne plak)
DEPTH_FRONT_FT = 1.64042      # ~500 mm voor het vlak (richting kijker)
HORIZONTAL_NORMAL_TOL = 0.001  # tolerantie voor "vrijwel horizontaal vlak"

# Curtain-wall sub-categorieen: bij het picken van een vliesgevel kies je
# een glaspaneel of stijl (los element). Voor de crop willen we de hele
# gevel (host-wand), niet een enkel paneel.
CURTAIN_SUB_BICS = (
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_CurtainWallMullions,
)


# ---------------------------------------------------------------- helpers

def _name(elem):
    if elem is None:
        return ""
    try:
        t = elem.GetType().GetProperty("Name")
        if t is not None:
            v = t.GetValue(elem, None)
            if v is not None:
                return v
    except Exception:
        pass
    try:
        return elem.Name or ""
    except Exception:
        return ""


def _id_int(elem):
    if elem is None:
        return None
    try:
        eid = elem.Id
    except Exception:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(eid, attr)
            if callable(v):
                v = v()
            return int(v)
        except Exception:
            continue
    return None


def _category_label(elem):
    try:
        cat = elem.Category
        if cat is not None:
            return _name(cat) or ""
    except Exception:
        pass
    return ""


def _category_bic_matches(elem, bics):
    """True als de categorie van elem matcht met een van de BuiltInCategory's.

    Robuust tegen Revit 2024+ (ElementId.Value) en ouder (IntegerValue).
    """
    try:
        cat = elem.Category
    except Exception:
        cat = None
    if cat is None:
        return False
    try:
        cid = cat.Id
    except Exception:
        return False
    cid_int = None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(cid, attr)
            if callable(v):
                v = v()
            cid_int = int(v)
            break
        except Exception:
            continue
    if cid_int is None:
        return False
    for bic in bics:
        try:
            if cid_int == int(bic):
                return True
        except Exception:
            continue
    return False


def _resolve_curtain_host(elem):
    """Geef de host-vliesgevel als elem een curtain-panel/stijl is, anders None.

    Bij het picken van een vliesgevel selecteer je een glaspaneel of een
    stijl/regel (mullion). Voor een correcte uitslag-crop hebben we de
    host-wand nodig die het hele vlak omvat.
    """
    if not _category_bic_matches(elem, CURTAIN_SUB_BICS):
        return None
    host = None
    try:
        host = elem.Host
    except Exception:
        host = None
    if host is None:
        try:
            sc = elem.SuperComponent
            if sc is not None:
                host = sc
        except Exception:
            host = None
    return host


def _unique_view_name(doc, base):
    existing = set()
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            existing.add(_name(v))
        except Exception:
            pass
    if base not in existing:
        return base
    i = 2
    while True:
        cand = u"{0} ({1})".format(base, i)
        if cand not in existing:
            return cand
        i += 1


def _get_section_view_family_type(doc):
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        try:
            if vft.ViewFamily == ViewFamily.Section:
                return vft
        except Exception:
            continue
    return None


def _collect_section_templates(doc):
    out = []
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if not v.IsTemplate:
                continue
            if v.ViewType == ViewType.Section:
                out.append(v)
        except Exception:
            continue
    out.sort(key=lambda x: _name(x).lower())
    return out


def _select_optional_template(label, items):
    if not items:
        return None
    name_to_elem = {}
    display = ["(geen template)"]
    for it in items:
        nm = _name(it)
        if nm in name_to_elem:
            nm = u"{0} [{1}]".format(nm, _id_int(it))
        name_to_elem[nm] = it
        display.append(nm)
    pick = forms.SelectFromList.show(
        display, title=label, button_name="Kies", multiselect=False,
    )
    if pick is None or pick == "(geen template)":
        return None
    return name_to_elem.get(pick)


# ---------------------------------------------------------------- pick

def _pick_face(uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Face,
            "Selecteer een vlak voor het loodrechte aanzicht",
        )
    except OperationCanceledException:
        return None, None, None
    if ref is None:
        return None, None, None
    doc = uidoc.Document
    elem = doc.GetElement(ref.ElementId)
    if elem is None:
        return None, None, None
    try:
        face = elem.GetGeometryObjectFromReference(ref)
    except Exception as ex:
        _log_exc("GetGeometryObjectFromReference failed: {0}".format(ex))
        return None, None, None
    return ref, elem, face


# ---------------------------------------------------------------- frame

def _face_center_and_normal(face):
    """Geef (center_world_point, normal) op het midden van de face.

    Werkt voor PlanarFace en curved faces. Bij curved faces wordt het
    UV-midden gebruikt.
    """
    try:
        bb = face.GetBoundingBox()
        u_mid = (bb.Min.U + bb.Max.U) / 2.0
        v_mid = (bb.Min.V + bb.Max.V) / 2.0
        uv = UV(u_mid, v_mid)
        normal = face.ComputeNormal(uv)
        center = face.Evaluate(uv)
        return center, normal
    except Exception:
        pass
    # planar fallback
    try:
        normal = face.FaceNormal
        origin = face.Origin
        return origin, normal
    except Exception:
        return XYZ.Zero, XYZ.BasisZ


def _build_view_frame(normal):
    """Bouw orthonormaal frame voor een section view.

    - BasisZ = normal (uit het vlak, naar de kijker)
    - BasisY = upslope (projectie wereld-Z op vlak)
    - BasisX = BasisY x BasisZ

    Voor (vrijwel) horizontale vlakken (vloer/plafond) wordt wereld-Y
    als up-referentie gebruikt om degenerate cross product te vermijden.
    """
    n = normal.Normalize()
    world_z = XYZ.BasisZ
    if abs(n.DotProduct(world_z)) > 1.0 - HORIZONTAL_NORMAL_TOL:
        # vrijwel horizontaal vlak: wereld-Y als up
        up_ref = XYZ.BasisY
    else:
        up_ref = world_z
    # Projecteer up_ref op het vlak loodrecht op n
    up_in_plane = up_ref.Subtract(n.Multiply(up_ref.DotProduct(n)))
    basis_y = up_in_plane.Normalize()
    basis_x = basis_y.CrossProduct(n).Normalize()
    basis_z = n
    return basis_x, basis_y, basis_z


def _element_bbox_corners(elem):
    """Geef 8 wereld-corners van element-bbox, of None."""
    try:
        bb = elem.get_BoundingBox(None)
    except Exception:
        return None
    if bb is None:
        return None
    mn, mx = bb.Min, bb.Max
    return [
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    ]


def _face_bbox_corners_world(face):
    """Geef wereld-corners van face UV-bbox (4 hoekpunten), of None."""
    try:
        bb = face.GetBoundingBox()
        u0, v0 = bb.Min.U, bb.Min.V
        u1, v1 = bb.Max.U, bb.Max.V
        return [
            face.Evaluate(UV(u0, v0)),
            face.Evaluate(UV(u1, v0)),
            face.Evaluate(UV(u0, v1)),
            face.Evaluate(UV(u1, v1)),
        ]
    except Exception:
        return None


def _project_to_local(point, origin, bx, by, bz):
    d = point.Subtract(origin)
    return (d.DotProduct(bx), d.DotProduct(by), d.DotProduct(bz))


def _compute_section_box(face, elem, center, bx, by, bz,
                         use_face_extent=True):
    """Bouw een BoundingBoxXYZ in view-coordinaten.

    Combineert face-UV-corners en element-bbox-corners zodat het
    aanzicht het hele element pakt achter het vlak. Origin van de
    Transform = `center` (face-midden).

    use_face_extent: True = strakke X/Y-crop op de gepickte face (default,
    gewone wand/vloer). False = X/Y-crop op de element-bbox (vliesgevel:
    de gepickte face is maar een paneel, de host-wand omvat de hele gevel).
    """
    pts = []
    fc = _face_bbox_corners_world(face)
    if fc:
        pts.extend(fc)
    ec = _element_bbox_corners(elem)
    if ec:
        pts.extend(ec)
    if not pts:
        # fallback: kleine box rond center
        return _build_box(center, bx, by, bz, 5.0, 5.0, DEPTH_FRONT_FT)

    xs, ys, zs = [], [], []
    for p in pts:
        lx, ly, lz = _project_to_local(p, center, bx, by, bz)
        xs.append(lx)
        ys.append(ly)
        zs.append(lz)

    # X/Y-extent bepalen. Gewone wand/vloer: strakke crop op de gepickte
    # face. Vliesgevel: face is maar een paneel -> gebruik de host-element-
    # bbox zodat de hele gevel in beeld komt.
    if use_face_extent and fc:
        extent_pts = fc
    elif ec:
        extent_pts = ec
    elif fc:
        extent_pts = fc
    else:
        extent_pts = pts
    ex_xs, ex_ys = [], []
    for p in extent_pts:
        lx, ly, lz = _project_to_local(p, center, bx, by, bz)
        ex_xs.append(lx)
        ex_ys.append(ly)
    x_min = min(ex_xs) - MARGIN_FT
    x_max = max(ex_xs) + MARGIN_FT
    y_min = min(ex_ys) - MARGIN_FT
    y_max = max(ex_ys) + MARGIN_FT

    # Z-diepte: section box Z-as wijst naar kijker (positief +Z),
    # alles achter het vlak heeft negatieve Z. Element-bbox bepaalt
    # hoe diep we moeten gaan om het achterliggende materiaal te
    # pakken.
    z_far_back = min(zs) if zs else -1.0
    if z_far_back > -DEPTH_BEHIND_FT:
        z_far_back = -DEPTH_BEHIND_FT
    z_min = z_far_back - DEPTH_BEHIND_FT
    z_max = DEPTH_FRONT_FT

    transform = Transform.Identity
    transform.Origin = center
    transform.BasisX = bx
    transform.BasisY = by
    transform.BasisZ = bz

    box = BoundingBoxXYZ()
    box.Transform = transform
    box.Min = XYZ(x_min, y_min, z_min)
    box.Max = XYZ(x_max, y_max, z_max)
    return box


def _build_box(center, bx, by, bz, half_w, half_h, depth):
    transform = Transform.Identity
    transform.Origin = center
    transform.BasisX = bx
    transform.BasisY = by
    transform.BasisZ = bz
    box = BoundingBoxXYZ()
    box.Transform = transform
    box.Min = XYZ(-half_w, -half_h, -depth)
    box.Max = XYZ(half_w, half_h, depth)
    return box


# ---------------------------------------------------------------- run

def run():
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()
    output.print_md("## Vlakaanzicht")
    output.print_md("*Debug-log: `{0}`*".format(_EARLY_LOG_FILE))

    # 1. Section ViewFamilyType ophalen
    vft = _get_section_view_family_type(doc)
    if vft is None:
        forms.alert(
            "Geen ViewFamilyType voor Section gevonden in dit model.",
            title="Fout",
        )
        return

    # 2. Vlak picken
    ref, elem, face = _pick_face(uidoc)
    if ref is None or face is None:
        output.print_md("*Geen vlak geselecteerd.*")
        return

    elem_cat = _category_label(elem)
    elem_name = _name(elem) or "Element"
    is_planar = isinstance(face, PlanarFace)
    _log("picked face: elem_id={0} cat={1} planar={2}".format(
        _id_int(elem), elem_cat, is_planar,
    ))

    # 2b. Vliesgevel: gepickt glaspaneel/stijl retargeten naar host-wand.
    host = _resolve_curtain_host(elem)
    is_curtain = host is not None and _id_int(host) != _id_int(elem)
    extent_elem = host if is_curtain else elem
    name_elem = host if is_curtain else elem
    if is_curtain:
        _log("curtain retarget: panel id={0} cat={1} -> host id={2}".format(
            _id_int(elem), elem_cat, _id_int(host),
        ))

    # 3. Center + normaal
    center, normal = _face_center_and_normal(face)
    if normal is None or normal.GetLength() < 1e-9:
        forms.alert(
            "Kon geen geldige face-normaal bepalen.",
            title="Fout",
        )
        return

    _log("normal=({0:.4f},{1:.4f},{2:.4f}) center=({3:.3f},{4:.3f},{5:.3f})".format(
        normal.X, normal.Y, normal.Z, center.X, center.Y, center.Z,
    ))

    # 4. View frame
    try:
        bx, by, bz = _build_view_frame(normal)
        _log("frame bx=({0:.4f},{1:.4f},{2:.4f}) by=({3:.4f},{4:.4f},{5:.4f}) "
             "bz=({6:.4f},{7:.4f},{8:.4f})".format(
                 bx.X, bx.Y, bx.Z, by.X, by.Y, by.Z, bz.X, bz.Y, bz.Z,
             ))
    except Exception as ex:
        _log_exc("build_view_frame failed: {0}".format(ex))
        forms.alert(
            "Kon view-orientatie niet berekenen: {0}".format(ex),
            title="Fout",
        )
        return

    # 5. Section box. Bij een vliesgevel cropt de host-bbox de hele gevel;
    # bij gewone elementen blijft de strakke face-crop behouden.
    section_box = _compute_section_box(
        face, extent_elem, center, bx, by, bz,
        use_face_extent=(not is_curtain),
    )

    # 6. Template (optioneel)
    templates = _collect_section_templates(doc)
    template = None
    if templates:
        template = _select_optional_template(
            "Kies template voor section view", templates,
        )
    template_id = template.Id if template is not None else None

    # 7. Schaal vragen
    scale_str = forms.ask_for_string(
        default=str(SCALE_DEFAULT),
        prompt="Schaal 1:X",
        title="Vlakaanzicht - schaal",
    )
    if scale_str is None:
        return
    try:
        scale = int(scale_str)
        if scale <= 0:
            raise ValueError("scale must be > 0")
    except Exception:
        forms.alert(
            "Ongeldige schaal: {0}".format(scale_str),
            title="Fout",
        )
        return

    # 8. Section view aanmaken
    tx = Transaction(doc, "Vlakaanzicht")
    tx.Start()
    new_view = None
    try:
        new_view = ViewSection.CreateSection(doc, vft.Id, section_box)
        if new_view is None:
            raise Exception("ViewSection.CreateSection gaf None terug")

        try:
            vd = new_view.ViewDirection
            ud = new_view.UpDirection
            rd = new_view.RightDirection
            _log("view RESULT viewdir=({0:.4f},{1:.4f},{2:.4f}) "
                 "updir=({3:.4f},{4:.4f},{5:.4f}) "
                 "rightdir=({6:.4f},{7:.4f},{8:.4f})".format(
                     vd.X, vd.Y, vd.Z, ud.X, ud.Y, ud.Z, rd.X, rd.Y, rd.Z,
                 ))
        except Exception as _ex:
            _log("view RESULT dirs lezen mislukt: {0}".format(_ex))

        name_cat = _category_label(name_elem)
        base_name = u"Vlakaanzicht {0} - id{1}".format(
            name_cat or _name(name_elem) or elem_name, _id_int(name_elem),
        )
        try:
            new_view.Name = _unique_view_name(doc, base_name)
        except Exception as ex:
            _log("rename view failed: {0}".format(ex))

        try:
            new_view.Scale = scale
        except Exception as ex:
            _log("set scale failed: {0}".format(ex))

        if template_id is not None:
            try:
                new_view.ViewTemplateId = template_id
            except Exception as ex:
                _log("apply template failed: {0}".format(ex))

        tx.Commit()
    except Exception as ex:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
        _log_exc("CreateSection failed")
        forms.alert(
            "Section view aanmaken mislukt:\n{0}: {1}".format(
                type(ex).__name__, ex,
            ),
            title="Fout",
        )
        return

    # 9. Activeer view + rapport
    try:
        uidoc.ActiveView = new_view
    except Exception as ex:
        _log("set ActiveView failed: {0}".format(ex))

    output.print_md("### Resultaat")
    output.print_md("- Element: **{0}** (id {1})".format(
        elem_cat or elem_name, _id_int(elem),
    ))
    if is_curtain:
        output.print_md(
            "- Vliesgevel: gepickt paneel geretarget naar host-wand "
            "**{0}** (id {1})".format(
                _category_label(host) or _name(host), _id_int(host),
            )
        )
    output.print_md("- Vlak-type: **{0}**".format(
        "PlanarFace" if is_planar else type(face).__name__,
    ))
    output.print_md(
        "- View: **{0}** (schaal 1:{1})".format(_name(new_view), scale)
    )
    output.print_md(
        "- Normaal: ({0:.3f}, {1:.3f}, {2:.3f})".format(
            normal.X, normal.Y, normal.Z,
        )
    )
    try:
        output.print_md(
            "[Open view: {0}]({1})".format(
                _name(new_view), output.linkify(new_view.Id),
            )
        )
    except Exception:
        pass


if __name__ == "__main__":
    if revit.doc is None:
        forms.alert("Geen Revit document geopend.", title="Fout")
    else:
        try:
            run()
        except Exception as ex:
            _log_exc("run() crashed")
            try:
                tb_text = _tb_mod.format_exc()
            except Exception:
                tb_text = "<no traceback>"
            forms.alert(
                "Onverwachte fout:\n{0}: {1}\n\n{2}\n\nLog: {3}".format(
                    type(ex).__name__, ex, tb_text, _EARLY_LOG_FILE,
                ),
                title="Fout",
            )
