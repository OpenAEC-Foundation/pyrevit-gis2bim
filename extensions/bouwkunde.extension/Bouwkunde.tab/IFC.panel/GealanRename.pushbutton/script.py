# -*- coding: utf-8 -*-
"""Hernoemt Gealan kozijn families naar K-merken op basis van mapping.txt"""
__title__ = "Gealan\nRename"
__author__ = "3BM Bouwkunde"
__doc__ = "Hernoemt Gealan kozijn family types naar K-merken.\n\n" \
          "Vereist: mapping.txt uit GealanSchemaReader.\n" \
          "Draai eerst 'Read K-merken' in het Gealan Tools panel."

import clr
import io
import os

clr.AddReference('RevitAPI')
clr.AddReference('PresentationFramework')

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# === CONFIG ===
MAPPING_FILE = os.path.join(
    os.environ.get('USERPROFILE', ''),
    'Documents', 'GealanReader', 'mapping.txt'
)

# === LOAD MAPPING ===
if not os.path.exists(MAPPING_FILE):
    forms.alert(
        "Mapping bestand niet gevonden!\n\n"
        "Pad: {}\n\n"
        "Draai eerst:\n"
        "  Gealan Tools > Read K-merken".format(MAPPING_FILE),
        title="Gealan Rename",
        exitscript=True
    )

mapping = {}
with io.open(MAPPING_FILE, 'r', encoding='utf-8-sig') as f:
    for line in f:
        line = line.strip()
        if not line or '|' not in line:
            continue
        parts = line.split('|')
        if len(parts) >= 2:
            family_name = parts[0].strip()
            k_code = parts[1].strip()
            brandwerend = parts[2].strip() == 'True' if len(parts) > 2 else False
            mapping[family_name] = (k_code, brandwerend)

if not mapping:
    forms.alert(
        "Mapping bestand is leeg!\n\n"
        "Draai eerst 'Read K-merken' in het Gealan Tools panel.",
        title="Gealan Rename",
        exitscript=True
    )

# === ANALYZE ===
collector = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Windows) \
    .WhereElementIsElementType()

to_rename = []
already_correct = []
no_mapping = []

for typ in list(collector):
    current_name = Element.Name.__get__(typ)
    
    try:
        family_name = typ.FamilyName
    except:
        continue
    
    if "Gealan" not in family_name:
        continue
    
    if family_name in mapping:
        target_k, brandwerend = mapping[family_name]
        if current_name == target_k:
            already_correct.append((family_name, current_name, brandwerend))
        else:
            to_rename.append((typ, family_name, current_name, target_k, brandwerend))
    else:
        no_mapping.append((family_name, current_name))

# === SHOW PREVIEW ===
total_gealan = len(to_rename) + len(already_correct) + len(no_mapping)

if not to_rename and not no_mapping:
    forms.alert(
        "Alle {} Gealan types zijn al correct hernoemd!".format(len(already_correct)),
        title="Gealan Rename"
    )
    script.exit()

# Build summary
summary_lines = []
summary_lines.append("GEALAN KOZIJNMERK HERNOEMER")
summary_lines.append("=" * 50)
summary_lines.append("")

if to_rename:
    summary_lines.append("TE HERNOEMEN ({} types):".format(len(to_rename)))
    for typ, fn, cur, tgt, bw in to_rename:
        bw_tag = " [BRANDWEREND]" if bw else ""
        summary_lines.append("  {} -> {}{}".format(cur, tgt, bw_tag))
    summary_lines.append("")

if already_correct:
    summary_lines.append("AL CORRECT ({} types)".format(len(already_correct)))
    summary_lines.append("")

if no_mapping:
    summary_lines.append("GEEN MAPPING ({} types):".format(len(no_mapping)))
    for fn, cur in no_mapping:
        summary_lines.append("  {} ({})".format(fn[-50:], cur))
    summary_lines.append("  -> Draai 'Read K-merken' om mapping bij te werken")
    summary_lines.append("")

summary_lines.append("Totaal Gealan types: {}".format(total_gealan))

summary = "\n".join(summary_lines)

if not to_rename:
    forms.alert(
        "Geen types om te hernoemen.\n\n"
        "{} types zonder mapping:\n{}".format(
            len(no_mapping),
            "\n".join(["  {} ({})".format(fn[-40:], cur) for fn, cur in no_mapping])
        ),
        title="Gealan Rename"
    )
    script.exit()

# Confirm
if not forms.alert(
    "{} types worden hernoemd:\n\n{}".format(
        len(to_rename),
        "\n".join(["  {} -> {}{}".format(cur, tgt, " [BW]" if bw else "") 
                   for _, _, cur, tgt, bw in to_rename])
    ),
    title="Gealan Rename - Bevestiging",
    yes=True, no=True
):
    script.exit()

# === RENAME ===
renamed = 0
errors = []

with revit.Transaction("Gealan kozijnmerken hernoemen"):
    for typ, fn, cur, tgt, bw in to_rename:
        try:
            typ.Name = tgt
            renamed += 1
        except Exception as ex:
            errors.append((fn, cur, tgt, str(ex)))

# === REPORT ===
output.print_md("# Gealan Kozijnmerk Hernoemer")
output.print_md("---")
output.print_md("**Hernoemd: {}** | Al correct: {} | Fouten: {}".format(
    renamed, len(already_correct), len(errors)))
output.print_md("")

if renamed > 0:
    output.print_md("## Hernoemd")
    for typ, fn, cur, tgt, bw in to_rename:
        if (fn, cur, tgt) not in [(e[0], e[1], e[2]) for e in errors]:
            bw_tag = " `BRANDWEREND`" if bw else ""
            output.print_md("- **{}** -> **{}**{}".format(cur, tgt, bw_tag))

if errors:
    output.print_md("## Fouten")
    for fn, cur, tgt, err in errors:
        output.print_md("- {} -> {}: `{}`".format(cur, tgt, err[:60]))

if no_mapping:
    output.print_md("## Geen mapping")
    output.print_md("*Draai 'Read K-merken' om deze toe te voegen:*")
    for fn, cur in no_mapping:
        output.print_md("- {} (`{}`)".format(fn[-50:], cur))

output.print_md("---")
output.print_md("*Mapping: {}*".format(MAPPING_FILE))
