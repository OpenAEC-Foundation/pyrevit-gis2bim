# -*- coding: utf-8 -*-
"""Import materiaal eigenschappen uit CSV naar Revit"""

__title__ = "Mat\nImp"
__author__ = "3BM Bouwkunde"
__doc__ = "Importeer lambda, Mu en Categorie uit CSV naar Revit materialen"

from pyrevit import revit, DB, forms, script
import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from bm_logger import get_logger

log = get_logger("MatImp")

# GEEN doc = revit.doc hier! Wordt in main() gedaan om startup-vertraging te voorkomen
# Globale referenties voor functies die doc/app nodig hebben
doc = None
app = None


def ensure_shared_param_file():
    """Zorg dat er een shared parameter file is"""
    spf_path = app.SharedParametersFilename
    
    if not spf_path or not os.path.exists(spf_path):
        # Maak shared parameter file aan
        temp_spf = os.path.join(os.environ['TEMP'], '3BM_SharedParams.txt')
        with open(temp_spf, 'w') as f:
            f.write("# 3BM Shared Parameters\n")
            f.write("*META\tVERSION\tMINVERSION\n")
            f.write("META\t2\t1\n")
            f.write("*GROUP\tID\tNAME\n")
            f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\n")
        app.SharedParametersFilename = temp_spf
        return temp_spf
    return spf_path


def create_shared_parameter(param_name, param_type, description):
    """Maak een shared parameter aan op Materials categorie"""
    try:
        # Check of parameter al bestaat
        test_mat = DB.FilteredElementCollector(doc).OfClass(DB.Material).FirstElement()
        if test_mat and test_mat.LookupParameter(param_name):
            return True, "{} bestaat al".format(param_name)
        
        # Zorg voor shared parameter file
        ensure_shared_param_file()
        
        # Open shared parameter file
        spf = app.OpenSharedParameterFile()
        if not spf:
            return False, "Kan shared parameter file niet openen"
        
        # Zoek of maak groep
        group = None
        for g in spf.Groups:
            if g.Name == "3BM Bouwfysica":
                group = g
                break
        
        if not group:
            group = spf.Groups.Create("3BM Bouwfysica")
        
        # Check of definitie al bestaat
        param_def = None
        for d in group.Definitions:
            if d.Name == param_name:
                param_def = d
                break
        
        # Maak definitie aan
        if not param_def:
            # Gebruik Number voor Mu (dimensieloos), Text voor Categorie
            if param_type == "number":
                opt = DB.ExternalDefinitionCreationOptions(param_name, DB.SpecTypeId.Number)
            else:  # text
                opt = DB.ExternalDefinitionCreationOptions(param_name, DB.SpecTypeId.String.Text)
            
            opt.Description = description
            opt.UserModifiable = True
            param_def = group.Definitions.Create(opt)
        
        # Bind aan Materials categorie
        cat_set = app.Create.NewCategorySet()
        cat_set.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Materials))
        
        binding = app.Create.NewInstanceBinding(cat_set)
        
        # Voeg toe aan project
        if doc.ParameterBindings.Insert(param_def, binding, DB.GroupTypeId.Data):
            return True, "{} aangemaakt".format(param_name)
        else:
            # Probeer ReInsert als het al gebonden was
            doc.ParameterBindings.ReInsert(param_def, binding, DB.GroupTypeId.Data)
            return True, "{} was al aanwezig".format(param_name)
        
    except Exception as e:
        return False, "Fout bij aanmaken {}: {}".format(param_name, str(e))


def read_csv(file_path):
    """Lees CSV en return lijst van te importeren materialen"""
    import_list = []
    
    with open(file_path, 'r') as f:
        lines = f.readlines()
    
    # Skip BOM en header
    start = 0
    for i, line in enumerate(lines):
        if line.startswith('\xef\xbb\xbf'):
            line = line[3:]
            lines[i] = line
        if 'Revit ID' in line:
            start = i + 1
            break
    
    for line in lines[start:]:
        line = line.strip()
        if not line:
            continue
        
        parts = line.split(';')
        if len(parts) < 11:
            continue
        
        try:
            revit_id = int(parts[0])
            importeren = parts[10].strip().upper()
            
            if importeren != 'J':
                continue
            
            # Parse voorgestelde waarden (NL formaat met komma)
            # Kolommen: ID;Naam;Lambda;Mu;Cat;MatchType;Match;VoorstLam;VoorstMu;VoorstCat;Import
            lam_str = parts[7].strip().replace(',', '.')
            mu_str = parts[8].strip()
            cat_str = parts[9].strip()
            
            lam = float(lam_str) if lam_str else None
            mu = int(float(mu_str)) if mu_str else None
            cat = cat_str if cat_str else None
            
            if lam is not None or mu is not None or cat is not None:
                import_list.append({
                    'id': revit_id,
                    'name': parts[1],
                    'lambda': lam,
                    'mu': mu,
                    'categorie': cat
                })
        except:
            continue
    
    return import_list


def set_thermal_conductivity(material, lambda_val):
    """Schrijf lambda naar Thermal Asset"""
    try:
        thermal_asset_id = material.ThermalAssetId
        
        if not thermal_asset_id or thermal_asset_id == DB.ElementId.InvalidElementId:
            # Maak thermal asset aan
            thermal_asset = DB.ThermalAsset("ThermalAsset_" + material.Name, DB.ThermalMaterialType.Solid)
            thermal_asset.ThermalConductivity = DB.UnitUtils.ConvertToInternalUnits(lambda_val, DB.UnitTypeId.WattsPerMeterKelvin)
            
            prop_set = DB.PropertySetElement.Create(doc, thermal_asset)
            material.SetMaterialAspectByPropertySet(DB.MaterialAspect.Thermal, prop_set.Id)
        else:
            # Update bestaande
            thermal_asset_elem = doc.GetElement(thermal_asset_id)
            if thermal_asset_elem:
                tc_param = thermal_asset_elem.get_Parameter(DB.BuiltInParameter.PHY_MATERIAL_PARAM_THERMAL_CONDUCTIVITY)
                if tc_param:
                    internal_val = DB.UnitUtils.ConvertToInternalUnits(lambda_val, DB.UnitTypeId.WattsPerMeterKelvin)
                    tc_param.Set(internal_val)
        
        return True
    except Exception as e:
        return False


def set_mu_value(material, mu_val):
    """Schrijf Mu naar custom parameter (Number - dimensieloos)"""
    try:
        mu_param = material.LookupParameter("Mu")
        if mu_param:
            if mu_param.StorageType == DB.StorageType.Double:
                # Number parameter: direct zetten zonder conversie
                mu_param.Set(float(mu_val))
                return True
            elif mu_param.StorageType == DB.StorageType.Integer:
                mu_param.Set(int(mu_val))
                return True
        return False
    except:
        return False


def set_categorie_value(material, cat_val):
    """Schrijf Categorie naar custom parameter"""
    try:
        cat_param = material.LookupParameter("Categorie")
        if cat_param and cat_param.StorageType == DB.StorageType.String:
            cat_param.Set(cat_val)
            return True
        return False
    except:
        return False


def main():
    global doc, app
    
    # Document check - hier, niet op module-niveau!
    doc = revit.doc
    if not doc:
        forms.alert("Open eerst een Revit project.", title="Materialen Import")
        return
    
    app = doc.Application
    
    # Vraag CSV bestand
    csv_path = forms.pick_file(
        file_ext='csv',
        title='Selecteer materialen CSV'
    )
    
    if not csv_path:
        return
    
    # Lees CSV
    import_list = read_csv(csv_path)
    
    if not import_list:
        forms.alert("Geen materialen met 'J' in Importeren kolom gevonden.", title="Materialen Import")
        return
    
    # Toon preview
    preview = "Te importeren: {} materialen\n\n".format(len(import_list))
    for item in import_list[:10]:
        preview += u"- {} (L={}, M={}, C={})\n".format(
            item['name'][:25],
            "{:.3f}".format(item['lambda']) if item['lambda'] else '-',
            item['mu'] if item['mu'] else '-',
            item['categorie'][:15] if item['categorie'] else '-'
        )
    if len(import_list) > 10:
        preview += "... en {} meer\n".format(len(import_list) - 10)
    
    preview += "\nDoorgaan?"
    
    if not forms.alert(preview, title="Materialen Import", yes=True, no=True):
        return
    
    # Check welke parameters nodig zijn
    needs_mu = any(item['mu'] is not None for item in import_list)
    needs_cat = any(item['categorie'] is not None for item in import_list)
    
    # Check of parameters bestaan
    test_mat = DB.FilteredElementCollector(doc).OfClass(DB.Material).FirstElement()
    has_mu = test_mat and test_mat.LookupParameter("Mu") is not None
    has_cat = test_mat and test_mat.LookupParameter("Categorie") is not None
    
    # Maak parameters aan indien nodig
    param_msgs = []
    
    if needs_mu and not has_mu:
        with revit.Transaction("Maak Mu parameter"):
            success, msg = create_shared_parameter("Mu", "number", "Dampdiffusieweerstandsfactor (dimensieloos)")
            param_msgs.append("Mu: " + msg)
            if not success:
                forms.alert("Kon Mu parameter niet aanmaken:\n{}\n\nMaak handmatig aan.".format(msg), warn_icon=True)
    
    if needs_cat and not has_cat:
        with revit.Transaction("Maak Categorie parameter"):
            success, msg = create_shared_parameter("Categorie", "text", "Materiaalcategorie voor filtering")
            param_msgs.append("Categorie: " + msg)
            if not success:
                forms.alert("Kon Categorie parameter niet aanmaken:\n{}\n\nMaak handmatig aan.".format(msg), warn_icon=True)
    
    # Import uitvoeren
    stats = {'lambda_ok': 0, 'lambda_fail': 0, 'mu_ok': 0, 'mu_fail': 0, 'cat_ok': 0, 'cat_fail': 0, 'not_found': 0}
    
    with revit.Transaction("Import materiaal eigenschappen"):
        for item in import_list:
            mat = doc.GetElement(DB.ElementId(item['id']))
            if not mat or not isinstance(mat, DB.Material):
                stats['not_found'] += 1
                continue
            
            if item['lambda'] is not None:
                if set_thermal_conductivity(mat, item['lambda']):
                    stats['lambda_ok'] += 1
                else:
                    stats['lambda_fail'] += 1
            
            if item['mu'] is not None:
                if set_mu_value(mat, item['mu']):
                    stats['mu_ok'] += 1
                else:
                    stats['mu_fail'] += 1
            
            if item['categorie'] is not None:
                if set_categorie_value(mat, item['categorie']):
                    stats['cat_ok'] += 1
                else:
                    stats['cat_fail'] += 1
    
    # Resultaat
    result_msg = "Import voltooid!\n\n"
    if param_msgs:
        result_msg += "Parameters:\n" + "\n".join(param_msgs) + "\n\n"
    
    result_msg += "Lambda: {} OK, {} mislukt\n".format(stats['lambda_ok'], stats['lambda_fail'])
    result_msg += "Mu: {} OK, {} mislukt\n".format(stats['mu_ok'], stats['mu_fail'])
    result_msg += "Categorie: {} OK, {} mislukt\n".format(stats['cat_ok'], stats['cat_fail'])
    result_msg += "Niet gevonden: {}".format(stats['not_found'])
    
    forms.alert(result_msg, title="Materialen Import")


if __name__ == '__main__':
    main()
