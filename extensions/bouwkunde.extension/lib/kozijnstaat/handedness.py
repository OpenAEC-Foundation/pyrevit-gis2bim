# -*- coding: utf-8 -*-
"""Handedness-detectie voor kozijn FamilyInstances.

Vervangt Konrad Sobon's archi-lab handedness node. Bepaalt per kozijn:
  - 'right'   : rechts-draaiend, intern
  - 'left'    : links-draaiend, intern
  - 'rhr'     : rechts-draaiend, gereverseerd (gevel)
  - 'lhr'     : links-draaiend, gereverseerd (gevel)

De logica combineert FacingFlipped x HandFlipped x ToRoom, met een
fallback voor instances waar ToRoom None is (interne kozijnen zonder
room, b.v. voordeur zonder 'to room' klas.).
"""


def classify_instance(instance):
    """Classificeer een kozijn FamilyInstance naar handedness-bucket.

    Args:
        instance: Revit FamilyInstance in Windows-categorie

    Returns:
        str: 'right', 'left', 'rhr', 'lhr' of 'unknown'
    """
    try:
        facing_flipped = bool(instance.FacingFlipped)
        hand_flipped = bool(instance.HandFlipped)
    except Exception:
        return "unknown"

    # ToRoom is None voor kozijnen zonder 'to-room' relatie.
    # In dat geval behandelen we ze als gereverseerd (gevel/outbound).
    try:
        to_room = instance.ToRoom
    except Exception:
        to_room = None

    has_to_room = to_room is not None

    if has_to_room:
        # Intern kozijn - standaard swing richting
        if not facing_flipped:
            return "right" if hand_flipped else "left"
        return "left" if hand_flipped else "right"

    # Gevel / gereverseerd kozijn
    if facing_flipped:
        return "rhr" if hand_flipped else "lhr"
    return "lhr" if hand_flipped else "rhr"


def classify_many(instances):
    """Classificeer een lijst instances. Returns dict met 4 buckets.

    Returns:
        dict[str, list[FamilyInstance]] met keys right/left/rhr/lhr/unknown
    """
    buckets = {"right": [], "left": [], "rhr": [], "lhr": [], "unknown": []}
    for inst in instances:
        bucket = classify_instance(inst)
        buckets[bucket].append(inst)
    return buckets


def is_mirrored(bucket_name):
    """Of een handedness-bucket een 'gespiegelde' variant is.

    Convention (3BM):
      - 'left'  en 'lhr' = niet gespiegeld (basis)
      - 'right' en 'rhr' = gespiegeld
    """
    return bucket_name in ("right", "rhr")
