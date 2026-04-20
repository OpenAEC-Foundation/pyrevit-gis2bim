# -*- coding: utf-8 -*-
"""Resolver voor named family references (voor maatvoering).

Kozijn families bevatten reference planes met namen als 'Left', 'Right',
'Sill', 'Head', 'vakvulling_a1_l' etc. Deze module vindt die References
op een FamilyInstance zodat ze als input voor Dimension.Create kunnen
dienen.
"""

from Autodesk.Revit.DB import (
    FamilyInstance,
    Options,
    ViewDetailLevel,
)


def get_named_references(instance, view=None):
    """Haal dict van reference-naam -> Reference voor een FamilyInstance.

    Gebruikt FamilyInstance.GetReferences(FamilyInstanceReferenceType)
    voor de standaard types (Left/Right/Top/Bottom/Front/Back) en
    FamilyInstance.GetReferenceByName voor de custom namen zoals
    'vakvulling_a1_l'.

    Args:
        instance: FamilyInstance
        view: View waarin de references zichtbaar zijn (optioneel,
            sommige reference types vereisen een view context)

    Returns:
        dict[str, Reference]
    """
    refs = {}

    # 1. Custom named references via GetReferenceByName
    #    We kennen de namen niet vooraf volledig; via GetRefences +
    #    Document.GetDefaultFamilyTypeId zouden we de lijst kunnen
    #    ontdekken, maar simpeler is: de caller geeft een lijst met
    #    te zoeken namen door aan resolve_references().
    #    Hier dus alleen helpers - de resolve-flow zit in resolve_references.

    return refs


def resolve_references(instance, names):
    """Los een lijst reference-namen op naar Reference objects.

    Args:
        instance: FamilyInstance
        names: list[str] van reference namen
            (bv. ['Left', 'Right', 'vakvulling_a1_l'])

    Returns:
        tuple (found_dict, missing_list) waarbij
            found_dict: dict[str, Reference]
            missing_list: list[str] van niet-gevonden namen
    """
    found = {}
    missing = []

    for name in names:
        ref = None
        try:
            ref = instance.GetReferenceByName(name)
        except Exception:
            ref = None

        if ref is None:
            missing.append(name)
        else:
            found[name] = ref

    return found, missing


def collect_all_named_references(instance):
    """Probeer alle named references van een instance te ontdekken.

    Dit werkt door de family-document te openen en alle reference planes
    met een niet-leeg 'Name' attribuut op te halen.

    Args:
        instance: FamilyInstance

    Returns:
        list[str] van gevonden reference namen (alleen namen, niet de
        Reference objects zelf - die zijn alleen via GetReferenceByName
        op te halen vanuit de project context)
    """
    doc = instance.Document
    try:
        family = instance.Symbol.Family
        fam_doc = doc.EditFamily(family)
    except Exception:
        return []

    names = []
    try:
        from Autodesk.Revit.DB import (
            FilteredElementCollector,
            ReferencePlane,
        )
        ref_planes = (
            FilteredElementCollector(fam_doc)
            .OfClass(ReferencePlane)
            .ToElements()
        )
        for rp in ref_planes:
            try:
                n = rp.Name
                if n and n.strip():
                    names.append(n)
            except Exception:
                continue
    finally:
        try:
            fam_doc.Close(False)
        except Exception:
            pass

    return names
