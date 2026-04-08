# -*- coding: utf-8 -*-
"""Export Revit materialen naar CSV met database matching"""

__title__ = "Mat\nExp"
__author__ = "3BM Bouwkunde"
__doc__ = "Exporteer Revit materialen naar CSV met database matching suggesties"

from pyrevit import revit, DB, forms, script
import os
import sys
import json

sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from bm_logger import get_logger

log = get_logger("MatExp")

# GEEN doc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen


def load_database():
    """Laad materialen uit JSON database"""
    lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
    json_path = os.path.join(lib_path, 'materialen_database.json')
    
    if os.path.exists(json_path):
        with open(json_path, 'r') as f:
            data = json.load(f)
            return data.get('materialen', [])
    return []


def get_thermal_conductivity(material, doc):
    """Lees lambda uit Revit Thermal Asset"""
    if not material:
        return None
    
    thermal_asset_id = material.ThermalAssetId
    if not thermal_asset_id or thermal_asset_id == DB.ElementId.InvalidElementId:
        return None
    
    thermal_asset = doc.GetElement(thermal_asset_id)
    if not thermal_asset:
        return None
    
    try:
        tc_param = thermal_asset.get_Parameter(DB.BuiltInParameter.PHY_MATERIAL_PARAM_THERMAL_CONDUCTIVITY)
        if tc_param and tc_param.HasValue:
            raw = tc_param.AsDouble()
            try:
                return DB.UnitUtils.ConvertFromInternalUnits(raw, DB.UnitTypeId.WattsPerMeterKelvin)
            except:
                return raw
    except:
        pass
    return None


def get_mu_from_material(material):
    """Lees Mu custom parameter uit Revit materiaal (Number - dimensieloos)"""
    if not material:
        return None
    try:
        mu_param = material.LookupParameter("Mu")
        if mu_param and mu_param.HasValue:
            if mu_param.StorageType == DB.StorageType.Double:
                # Number parameter: direct uitlezen, geen conversie
                return int(round(mu_param.AsDouble()))
            elif mu_param.StorageType == DB.StorageType.Integer:
                return mu_param.AsInteger()
    except:
        pass
    return None


def get_categorie_from_material(material):
    """Lees Categorie custom parameter uit Revit materiaal"""
    if not material:
        return None
    try:
        cat_param = material.LookupParameter("Categorie")
        if cat_param and cat_param.HasValue:
            return cat_param.AsString()
    except:
        pass
    return None


def match_material(name, database):
    """Zoek beste match in database op basis van keywords"""
    if not name or not database:
        return None, None, None, None, ""
    
    name_lower = name.lower()
    best_match = None
    best_score = 0
    
    for mat in database:
        keywords = mat.get('keywords', [])
        score = 0
        
        for kw in keywords:
            kw_lower = kw.lower()
            if kw_lower in name_lower:
                score += len(kw) * 2
                if kw_lower == name_lower:
                    score += 50
        
        if score > best_score:
            best_score = score
            best_match = mat
    
    if best_match and best_score >= 3:
        lam = best_match.get('lambda')
        mu = best_match.get('mu')
        cat = best_match.get('categorie')
        match_type = "exact" if best_score >= 50 else ("goed" if best_score >= 10 else "zwak")
        return best_match.get('naam'), lam, mu, cat, match_type
    
    return None, None, None, None, ""


def get_all_revit_materials(doc):
    """Haal alle materialen uit Revit model"""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.Material)
    materials = []
    
    for mat in collector:
        mat_data = {
            'id': mat.Id.IntegerValue,
            'name': mat.Name,
            'lambda_revit': get_thermal_conductivity(mat, doc),
            'mu_revit': get_mu_from_material(mat),
            'categorie_revit': get_categorie_from_material(mat),
            'has_thermal_asset': mat.ThermalAssetId != DB.ElementId.InvalidElementId
        }
        materials.append(mat_data)
    
    return sorted(materials, key=lambda x: x['name'])


def export_to_csv(revit_materials, database, output_path):
    """Exporteer Revit materialen naar CSV met matching"""
    try:
        stats = {'total': 0, 'matched': 0, 'has_lambda': 0, 'has_mu': 0, 'has_cat': 0}
        
        with open(output_path, 'w') as f:
            f.write('\xef\xbb\xbf')  # UTF-8 BOM
            # Uitgebreide header met categorie
            f.write("Revit ID;Materiaal Naam;Huidige Lambda;Huidige Mu;Huidige Categorie;Match Type;Voorgestelde Match;Voorgestelde Lambda;Voorgestelde Mu;Voorgestelde Categorie;Importeren (J/N)\n")
            
            for mat in revit_materials:
                stats['total'] += 1
                
                # Matching
                match_naam, match_lam, match_mu, match_cat, match_type = match_material(mat['name'], database)
                
                if match_naam:
                    stats['matched'] += 1
                if mat['lambda_revit'] is not None:
                    stats['has_lambda'] += 1
                if mat['mu_revit'] is not None:
                    stats['has_mu'] += 1
                if mat['categorie_revit']:
                    stats['has_cat'] += 1
                
                # Default importeren: J als er ontbrekende waarden zijn
                needs_lambda = mat['lambda_revit'] is None and match_lam is not None
                needs_mu = mat['mu_revit'] is None and match_mu is not None
                needs_cat = not mat['categorie_revit'] and match_cat
                importeren = "J" if (needs_lambda or needs_mu or needs_cat) else "N"
                
                # Formatteren met komma als decimaal
                lam_str = "{:.4f}".format(mat['lambda_revit']).replace('.', ',') if mat['lambda_revit'] is not None else ''
                match_lam_str = "{:.4f}".format(match_lam).replace('.', ',') if match_lam is not None else ''
                
                f.write(u"{};{};{};{};{};{};{};{};{};{};{}\n".format(
                    mat['id'],
                    mat['name'].replace(';', ','),
                    lam_str,
                    mat['mu_revit'] if mat['mu_revit'] is not None else '',
                    mat['categorie_revit'] or '',
                    match_type or "geen",
                    (match_naam or '').replace(';', ','),
                    match_lam_str,
                    match_mu if match_mu is not None else '',
                    match_cat or '',
                    importeren
                ))
        
        return True, stats
    except Exception as e:
        return False, str(e)


def main():
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Materialen Export")
        return
    
    # Laad database
    database = load_database()
    if not database:
        forms.alert("Materialen database niet gevonden!", warn_icon=True)
        return
    
    # Haal Revit materialen
    revit_materials = get_all_revit_materials(doc)
    
    if not revit_materials:
        forms.alert("Geen materialen gevonden in Revit model.", title="Materialen Export")
        return
    
    # Output bestand kiezen - default naar lib folder (naast json)
    lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
    project_name = doc.Title.replace('.rvt', '') if doc.Title else 'project'
    default_path = os.path.join(lib_path, '{}_materialen.csv'.format(project_name))
    
    output_path = forms.save_file(
        file_ext='csv',
        default_name='{}_materialen.csv'.format(project_name),
        title='Exporteer Revit Materialen'
    )
    
    if not output_path:
        return
    
    # Export
    success, result = export_to_csv(revit_materials, database, output_path)
    
    if success:
        stats = result
        forms.alert(
            "Export voltooid!\n\n"
            "Totaal: {} materialen\n"
            "Matched: {} ({:.0f}%)\n"
            "Met lambda: {}\n"
            "Met Mu: {}\n"
            "Met categorie: {}\n\n"
            "Review 'Importeren' kolom en gebruik\n"
            "Mat Imp om waarden terug te schrijven.".format(
                stats['total'],
                stats['matched'],
                100.0 * stats['matched'] / stats['total'] if stats['total'] > 0 else 0,
                stats['has_lambda'],
                stats['has_mu'],
                stats['has_cat']
            ),
            title="Materialen Export"
        )
        os.startfile(output_path)
    else:
        forms.alert("Export mislukt:\n\n{}".format(result), title="Materialen Export", warn_icon=True)


if __name__ == '__main__':
    main()
