# -*- coding: utf-8 -*-
"""
Family Loading & DirectShape Fallback
======================================

Hybride aanpak voor het plaatsen van BGT 3D elementen:
1. Zoek bestaande family in document
2. Laad .rfa bestand uit families/ map
3. Genereer DirectShape als fallback

Gebruik:
    from gis2bim.revit.families import resolve_family_or_fallback, place_element

    strategy = resolve_family_or_fallback(doc, "loofboom", families_dir, "tree")
    place_element(doc, strategy, xyz, "Boom 1")
"""

import os

from .geometry import (
    meters_to_feet,
    create_cylinder_solid,
    create_box_solid,
    create_directshape,
    get_or_create_material,
    set_element_material,
    IN_REVIT,
)


# =============================================================================
# PlacementStrategy
# =============================================================================

class PlacementStrategy(object):
    """Encapsuleert hoe een element geplaatst wordt.

    Attributes:
        mode: "family" of "directshape"
        family_symbol: FamilySymbol als mode == "family", anders None
        fallback_type: "tree" of "lamp" als mode == "directshape"
        family_name: Oorspronkelijke family naam (voor logging)
    """

    def __init__(self, mode, family_symbol=None, fallback_type=None,
                 family_name=""):
        """Initialiseer PlacementStrategy.

        Args:
            mode: "family" of "directshape"
            family_symbol: FamilySymbol object (bij mode "family")
            fallback_type: "tree" of "lamp" (bij mode "directshape")
            family_name: Naam van de oorspronkelijke family
        """
        self.mode = mode
        self.family_symbol = family_symbol
        self.fallback_type = fallback_type
        self.family_name = family_name

    def __repr__(self):
        """String representatie voor logging."""
        if self.mode == "family":
            return "PlacementStrategy(family={0})".format(self.family_name)
        return "PlacementStrategy(directshape={0})".format(self.fallback_type)


# =============================================================================
# Family Resolution
# =============================================================================

def find_family_symbol(doc, family_name):
    """Zoek FamilySymbol op naam in het document.

    Args:
        doc: Revit document
        family_name: Naam van de family

    Returns:
        FamilySymbol of None
    """
    if not IN_REVIT:
        return None

    from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol

    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for fs in collector:
        if fs.Family.Name == family_name:
            return fs
    return None


def load_family_from_file(doc, rfa_path):
    """Laad een .rfa bestand in het document.

    Args:
        doc: Revit document
        rfa_path: Volledig pad naar het .rfa bestand

    Returns:
        FamilySymbol van de geladen family, of None bij fout
    """
    if not IN_REVIT:
        return None

    if not os.path.isfile(rfa_path):
        return None

    try:
        from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol
        import clr

        # doc.LoadFamily(path) retourneert bool in IronPython
        # We gebruiken de overload met out-parameter via clr
        family_ref = clr.Reference[type(None)]()
        try:
            # Probeer overload met out-parameter
            loaded = doc.LoadFamily(rfa_path, family_ref)
            family = family_ref.Value
        except TypeError:
            # Fallback: simpele overload (retourneert alleen bool)
            loaded = doc.LoadFamily(rfa_path)
            family = None

        if not loaded and family is None:
            # Family bestaat mogelijk al — zoek op bestandsnaam
            family_name = os.path.splitext(os.path.basename(rfa_path))[0]
            return find_family_symbol(doc, family_name)

        # Zoek eerste FamilySymbol van de geladen family
        if family is not None:
            symbol_ids = family.GetFamilySymbolIds()
            if symbol_ids.Count > 0:
                for sid in symbol_ids:
                    return doc.GetElement(sid)

        # Fallback: zoek op bestandsnaam
        family_name = os.path.splitext(os.path.basename(rfa_path))[0]
        return find_family_symbol(doc, family_name)

    except Exception:
        return None


def get_bundled_family_path(family_name, families_dir):
    """Zoek een .rfa bestand in de families/ map.

    Args:
        family_name: Naam van de family (zonder extensie)
        families_dir: Pad naar de families/ map

    Returns:
        Volledig pad naar het .rfa bestand, of None
    """
    if not families_dir or not os.path.isdir(families_dir):
        return None

    # Exacte match op bestandsnaam
    rfa_path = os.path.join(families_dir, "{0}.rfa".format(family_name))
    if os.path.isfile(rfa_path):
        return rfa_path

    # Case-insensitive zoeken
    target = family_name.lower() + ".rfa"
    for filename in os.listdir(families_dir):
        if filename.lower() == target:
            return os.path.join(families_dir, filename)

    return None


def resolve_family_or_fallback(doc, family_name, families_dir, fallback_type):
    """Bepaal de plaatsingsstrategie via resolutie-keten.

    Volgorde:
    1. Zoek family in document
    2. Laad .rfa uit families/ map
    3. Gebruik DirectShape fallback

    Args:
        doc: Revit document
        family_name: Naam van de gewenste family
        families_dir: Pad naar de families/ map (mag None zijn)
        fallback_type: "tree" of "lamp" voor DirectShape

    Returns:
        PlacementStrategy object
    """
    # Stap 1: zoek in document
    symbol = find_family_symbol(doc, family_name)
    if symbol is not None:
        return PlacementStrategy(
            mode="family",
            family_symbol=symbol,
            family_name=family_name,
        )

    # Stap 2: laad uit .rfa bestand
    rfa_path = get_bundled_family_path(family_name, families_dir)
    if rfa_path is not None:
        symbol = load_family_from_file(doc, rfa_path)
        if symbol is not None:
            return PlacementStrategy(
                mode="family",
                family_symbol=symbol,
                family_name=family_name,
            )

    # Stap 3: DirectShape fallback
    return PlacementStrategy(
        mode="directshape",
        fallback_type=fallback_type,
        family_name=family_name,
    )


# =============================================================================
# Element Placement
# =============================================================================

def place_element(doc, strategy, xyz, name=None):
    """Plaats een element op het opgegeven punt.

    Plaatst een FamilyInstance (strategy.mode == "family") of
    DirectShape (strategy.mode == "directshape") op de xyz positie.

    Args:
        doc: Revit document
        strategy: PlacementStrategy object
        xyz: XYZ punt in Revit internal units (feet)
        name: Optionele naam voor het element

    Returns:
        Het geplaatste Revit element, of None bij fout
    """
    if strategy.mode == "family":
        return _place_family_instance(doc, strategy.family_symbol, xyz)
    elif strategy.mode == "directshape":
        if strategy.fallback_type == "tree":
            return create_tree_directshape(doc, xyz, name=name)
        elif strategy.fallback_type == "lamp":
            return create_lamp_directshape(doc, xyz, name=name)
    return None


def _place_family_instance(doc, family_symbol, xyz):
    """Plaats een FamilyInstance op het opgegeven punt.

    Args:
        doc: Revit document
        family_symbol: FamilySymbol om te plaatsen
        xyz: XYZ punt in Revit internal units (feet)

    Returns:
        FamilyInstance element
    """
    if not IN_REVIT:
        return None

    from Autodesk.Revit.DB import Structure

    if not family_symbol.IsActive:
        family_symbol.Activate()
        doc.Regenerate()

    return doc.Create.NewFamilyInstance(
        xyz, family_symbol,
        Structure.StructuralType.NonStructural
    )


# =============================================================================
# DirectShape Fallback Geometry
# =============================================================================

# Materiaal definities (naam, RGB)
_MAT_STAM = ("GIS2BIM - Boomstam", (139, 90, 43))
_MAT_KRUIN = ("GIS2BIM - Boomkruin", (60, 130, 60))
_MAT_PAAL = ("GIS2BIM - Lichtmastpaal", (130, 130, 140))
_MAT_ARMATUUR = ("GIS2BIM - Armatuur", (220, 200, 100))


def create_tree_directshape(doc, xyz, name=None):
    """Maak een boom als DirectShape: cilinder stam + cilinder kruin.

    Afmetingen:
    - Stam: r=0.15m, h=3.0m, basis z=0
    - Kruin: r=2.0m, h=4.0m, basis z=3.0m (bovenop stam)

    Gebruikt cilinders i.p.v. bol voor maximale Revit-compatibiliteit.

    Args:
        doc: Revit document
        xyz: XYZ punt (basis van de boom) in feet
        name: Optionele naam

    Returns:
        DirectShape element
    """
    cx = xyz.X
    cy = xyz.Y
    z_base = xyz.Z

    # Stam: cilinder r=0.15m, h=3.0m
    stam_r = meters_to_feet(0.15)
    stam_h = meters_to_feet(3.0)
    stam_solid = create_cylinder_solid(cx, cy, z_base, stam_r, stam_h)

    # Kruin: brede cilinder r=2.0m, h=4.0m, start op z=3.0m
    kruin_r = meters_to_feet(2.0)
    kruin_h = meters_to_feet(4.0)
    kruin_z = z_base + meters_to_feet(3.0)
    kruin_solid = create_cylinder_solid(cx, cy, kruin_z, kruin_r, kruin_h,
                                        segments=32)

    ds = create_directshape(doc, [stam_solid, kruin_solid],
                            name=name or "GIS2BIM Boom")

    # Materialen toewijzen
    _apply_tree_materials(doc, ds)

    return ds


def create_lamp_directshape(doc, xyz, name=None):
    """Maak een lantaarnpaal als DirectShape: cilinder paal + box armatuur.

    Afmetingen:
    - Paal: r=0.075m, h=5.0m, basis z=0
    - Armatuur: 0.4x0.4x0.15m, top z=5.0m

    Args:
        doc: Revit document
        xyz: XYZ punt (basis van de paal) in feet
        name: Optionele naam

    Returns:
        DirectShape element
    """
    cx = xyz.X
    cy = xyz.Y
    z_base = xyz.Z

    # Paal: cilinder r=0.075m, h=5.0m
    paal_r = meters_to_feet(0.075)
    paal_h = meters_to_feet(5.0)
    paal_solid = create_cylinder_solid(cx, cy, z_base, paal_r, paal_h)

    # Armatuur: box 0.4x0.4x0.15m, top op z=5.0m
    arm_w = meters_to_feet(0.4)
    arm_d = meters_to_feet(0.4)
    arm_h = meters_to_feet(0.15)
    arm_z = z_base + meters_to_feet(5.0) - arm_h  # top = 5.0m
    arm_solid = create_box_solid(cx, cy, arm_z, arm_w, arm_d, arm_h)

    ds = create_directshape(doc, [paal_solid, arm_solid],
                            name=name or "GIS2BIM Lichtmast")

    # Materialen toewijzen
    _apply_lamp_materials(doc, ds)

    return ds


def _apply_tree_materials(doc, ds):
    """Wijs boom-materialen toe aan een DirectShape.

    Args:
        doc: Revit document
        ds: DirectShape element met stam- en kruin-solids
    """
    try:
        stam_mat_id = get_or_create_material(doc, _MAT_STAM[0], _MAT_STAM[1])
        kruin_mat_id = get_or_create_material(doc, _MAT_KRUIN[0], _MAT_KRUIN[1])

        # Paint alle faces — stam eerst (cylinder), dan kruin (bol)
        # Omdat we 2 solids hebben, paint met het kruin-materiaal (dominant)
        # en het stam-materiaal komt via de eerste solid
        _paint_solids_separately(doc, ds, [stam_mat_id, kruin_mat_id])
    except Exception:
        pass


def _apply_lamp_materials(doc, ds):
    """Wijs lamp-materialen toe aan een DirectShape.

    Args:
        doc: Revit document
        ds: DirectShape element met paal- en armatuur-solids
    """
    try:
        paal_mat_id = get_or_create_material(doc, _MAT_PAAL[0], _MAT_PAAL[1])
        arm_mat_id = get_or_create_material(doc, _MAT_ARMATUUR[0], _MAT_ARMATUUR[1])

        _paint_solids_separately(doc, ds, [paal_mat_id, arm_mat_id])
    except Exception:
        pass


def _paint_solids_separately(doc, element, material_ids):
    """Paint individuele solids van een DirectShape met verschillende materialen.

    Itereert door de geometry van het element en wijst per solid
    het corresponderende materiaal toe.

    Args:
        doc: Revit document
        element: DirectShape element
        material_ids: Lijst van ElementId's, één per solid
    """
    if not IN_REVIT:
        return

    from Autodesk.Revit.DB import Solid, Options

    opts = Options()
    opts.ComputeReferences = True
    geom = element.get_Geometry(opts)
    if geom is None:
        return

    solid_index = 0
    for geom_obj in geom:
        if not isinstance(geom_obj, Solid):
            continue
        if geom_obj.Faces.Size == 0:
            continue

        mat_id = material_ids[min(solid_index, len(material_ids) - 1)]
        for face in geom_obj.Faces:
            try:
                doc.Paint(element.Id, face, mat_id)
            except Exception:
                pass
        solid_index += 1
