# -*- coding: utf-8 -*-
"""
Revit Location Interface
========================

Beheer van Survey Point en Project Base Point voor GIS2BIM.
Inclusief automatisch aanmaken van Project Parameters.
"""

import math
import os

# Debug: print import status
print("GIS2BIM location.py: Starting imports...")

# pyRevit imports - alleen beschikbaar in Revit context
IN_REVIT = False
HAS_SPEC_TYPE_ID = False
HAS_FORGEYPEID = False

try:
    from Autodesk.Revit.DB import (
        Transaction,
        FilteredElementCollector,
        XYZ,
        ProjectPosition,
        CategorySet,
        InstanceBinding,
        BuiltInCategory,
        ExternalDefinitionCreationOptions
    )
    IN_REVIT = True
    print("GIS2BIM location.py: Basic Revit imports OK")
except ImportError as e:
    print("GIS2BIM location.py: Basic import failed: {0}".format(e))
except Exception as e:
    print("GIS2BIM location.py: Basic import error: {0}".format(e))

# Revit 2022+: SpecTypeId en ForgeTypeId
if IN_REVIT:
    try:
        from Autodesk.Revit.DB import SpecTypeId, ForgeTypeId
        HAS_SPEC_TYPE_ID = True
        HAS_FORGEYPEID = True
        print("GIS2BIM location.py: SpecTypeId/ForgeTypeId OK (Revit 2022+)")
    except ImportError:
        print("GIS2BIM location.py: SpecTypeId not available")

# Revit 2021 en eerder: ParameterType
if IN_REVIT:
    try:
        from Autodesk.Revit.DB import ParameterType
        HAS_PARAMETER_TYPE = True
        print("GIS2BIM location.py: ParameterType OK (legacy)")
    except ImportError:
        HAS_PARAMETER_TYPE = False
        print("GIS2BIM location.py: ParameterType not available")

# BuiltInParameterGroup (deprecated in 2024+, gebruik GroupTypeId)
if IN_REVIT:
    try:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        HAS_BIPG = True
        print("GIS2BIM location.py: BuiltInParameterGroup OK")
    except ImportError:
        HAS_BIPG = False
        print("GIS2BIM location.py: BuiltInParameterGroup not available (Revit 2024+)")
    
    try:
        from Autodesk.Revit.DB import GroupTypeId
        HAS_GROUP_TYPE_ID = True
        print("GIS2BIM location.py: GroupTypeId OK (Revit 2024+)")
    except ImportError:
        HAS_GROUP_TYPE_ID = False

print("GIS2BIM location.py: IN_REVIT = {0}".format(IN_REVIT))

# Constanten
FEET_TO_METER = 0.3048
METER_TO_FEET = 1 / FEET_TO_METER

# Shared parameter file voor GIS2BIM
SHARED_PARAM_GROUP = "GIS2BIM"

# Mortoncode parameters voor Nederlandse tiling (GIS2BIM standaard)
MORTON_X_OFFSET = -100000     # RD oorsprong X offset
MORTON_Y_OFFSET = 200000      # RD oorsprong Y offset  
MORTON_TILE_SIZE = 2000       # Tile grootte in meters


def calculate_mortoncode(rd_x, rd_y, x_offset=None, y_offset=None, tile_size=None):
    """
    Bereken Mortoncode (Z-order curve) voor RD coördinaten.
    
    Gebaseerd op de Dynamo implementatie:
    - Relatieve positie tov offset
    - Gedeeld door tile grootte
    - Bit interleaving van Y en X
    
    Args:
        rd_x: RD X coördinaat
        rd_y: RD Y coördinaat
        x_offset: X offset van tile oorsprong (default: MORTON_X_OFFSET)
        y_offset: Y offset van tile oorsprong (default: MORTON_Y_OFFSET)
        tile_size: Tile grootte in meters (default: MORTON_TILE_SIZE)
    
    Returns:
        Mortoncode als integer
    """
    import math
    
    # Gebruik defaults als niet opgegeven
    if x_offset is None:
        x_offset = MORTON_X_OFFSET
    if y_offset is None:
        y_offset = MORTON_Y_OFFSET
    if tile_size is None:
        tile_size = MORTON_TILE_SIZE
    
    # Bereken tile indices
    x_tile = int(math.floor((rd_x - x_offset) / tile_size))
    y_tile = int(math.floor((rd_y - y_offset) / tile_size))
    
    # Converteer naar binaire strings (zonder '0b' prefix)
    x_bin = bin(x_tile)[2:] if x_tile >= 0 else bin(x_tile)[3:]
    y_bin = bin(y_tile)[2:] if y_tile >= 0 else bin(y_tile)[3:]
    
    # Maak beide strings even lang (pad met nullen aan de linkerkant)
    max_len = max(len(x_bin), len(y_bin))
    x_bin = x_bin.zfill(max_len)
    y_bin = y_bin.zfill(max_len)
    
    # Interleave: y bits eerst, dan x bits (zoals in je Dynamo code)
    interleaved = "".join(y + x for y, x in zip(y_bin, x_bin))
    
    # Converteer terug naar integer
    morton = int(interleaved, 2) if interleaved else 0
    
    return morton


def _get_or_create_shared_param_file(app):
    """
    Haal bestaande shared parameter file op of maak nieuwe aan.
    
    Returns:
        DefinitionFile object
    """
    current_file = app.OpenSharedParameterFile()
    
    if current_file is not None:
        return current_file
    
    # Maak tijdelijke shared parameter file aan
    temp_folder = os.environ.get("TEMP", os.environ.get("TMP", "C:\\Temp"))
    shared_param_path = os.path.join(temp_folder, "GIS2BIM_SharedParams.txt")
    
    # Maak bestand aan als het niet bestaat
    if not os.path.exists(shared_param_path):
        with open(shared_param_path, "w") as f:
            f.write("# GIS2BIM Shared Parameters\n")
    
    app.SharedParametersFilename = shared_param_path
    return app.OpenSharedParameterFile()


def _get_or_create_definition_group(def_file, group_name):
    """Haal definitiegroep op of maak aan."""
    groups = def_file.Groups
    group = groups.get_Item(group_name)
    
    if group is None:
        group = groups.Create(group_name)
    
    return group


def _create_project_parameter(doc, param_name):
    """
    Maak een nieuwe Project Parameter aan voor ProjectInformation.
    
    Args:
        doc: Revit document
        param_name: Naam van de parameter
    
    Returns:
        True bij succes, False bij fout
    """
    if not IN_REVIT:
        return False
    
    try:
        app = doc.Application
        
        # Haal of maak shared parameter file
        def_file = _get_or_create_shared_param_file(app)
        if def_file is None:
            print("GIS2BIM: Kon shared parameter file niet openen/aanmaken")
            return False
        
        # Haal of maak definitiegroep
        group = _get_or_create_definition_group(def_file, SHARED_PARAM_GROUP)
        
        # Check of definitie al bestaat
        definition = group.Definitions.get_Item(param_name)
        
        if definition is None:
            # Maak nieuwe definitie aan
            try:
                if HAS_SPEC_TYPE_ID:
                    # Revit 2022+
                    options = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text)
                else:
                    # Revit 2021 en eerder
                    options = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)
                options.Visible = True
                definition = group.Definitions.Create(options)
                print("GIS2BIM: Parameter definitie aangemaakt: {0}".format(param_name))
            except Exception as e:
                print("GIS2BIM: Kon parameter definitie niet aanmaken: {0}".format(e))
                return False
        
        # Check of parameter al gebonden is aan ProjectInformation
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        
        while iterator.MoveNext():
            if iterator.Key.Name == param_name:
                # Parameter bestaat al
                print("GIS2BIM: Parameter binding bestaat al: {0}".format(param_name))
                return True
        
        # Bind parameter aan ProjectInformation category
        cat_set = CategorySet()
        pi_category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ProjectInformation)
        cat_set.Insert(pi_category)
        
        instance_binding = InstanceBinding(cat_set)
        
        # Voeg binding toe - verschillende API voor verschillende Revit versies
        success = False
        if HAS_GROUP_TYPE_ID:
            # Revit 2024+
            try:
                success = binding_map.Insert(definition, instance_binding, GroupTypeId.IdentityData)
                print("GIS2BIM: Binding via GroupTypeId: {0}".format(success))
            except Exception as e:
                print("GIS2BIM: GroupTypeId binding failed: {0}".format(e))
        
        if not success and HAS_BIPG:
            # Revit 2021-2023
            try:
                success = binding_map.Insert(definition, instance_binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
                print("GIS2BIM: Binding via BuiltInParameterGroup: {0}".format(success))
            except Exception as e:
                print("GIS2BIM: BIPG binding failed: {0}".format(e))
        
        if not success:
            # Fallback: probeer zonder group
            try:
                success = binding_map.Insert(definition, instance_binding)
                print("GIS2BIM: Binding zonder group: {0}".format(success))
            except Exception as e:
                print("GIS2BIM: Fallback binding failed: {0}".format(e))
        
        return success
        
    except Exception as e:
        print("GIS2BIM: Fout bij aanmaken parameter '{0}': {1}".format(param_name, e))
        return False


def get_rd_from_project_params(doc):
    """
    Lees RD coordinaten uit GIS2BIM project parameters.
    
    Zoekt naar GIS2BIM_RD_X en GIS2BIM_RD_Y in Project Information.
    Dit is betrouwbaarder dan het Survey Point omdat EW/NS omgedraaid kan zijn.
    
    Args:
        doc: Revit document
    
    Returns:
        Dict met {"rd_x": float, "rd_y": float} of None
    """
    if not IN_REVIT:
        return None
    
    try:
        project_info = doc.ProjectInformation
        
        param_x = project_info.LookupParameter("GIS2BIM_RD_X")
        param_y = project_info.LookupParameter("GIS2BIM_RD_Y")
        
        if param_x is None or param_y is None:
            print("GIS2BIM: Project parameters GIS2BIM_RD_X/Y niet gevonden")
            return None
        
        val_x = param_x.AsString()
        val_y = param_y.AsString()
        
        if not val_x or not val_y:
            print("GIS2BIM: Project parameters GIS2BIM_RD_X/Y zijn leeg")
            return None
        
        rd_x = float(val_x)
        rd_y = float(val_y)
        
        if rd_x == 0 and rd_y == 0:
            return None
        
        print("GIS2BIM: RD uit project params: X={0}, Y={1}".format(rd_x, rd_y))
        return {"rd_x": rd_x, "rd_y": rd_y}
    
    except Exception as e:
        print("GIS2BIM: Fout bij lezen project params: {0}".format(e))
        return None


def get_survey_point(doc):
    """Haal Survey Point coördinaten op via ProjectPosition."""
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit (IN_REVIT={0})".format(IN_REVIT))
    
    project_location = doc.ActiveProjectLocation
    position = project_location.GetProjectPosition(XYZ(0, 0, 0))
    
    return {
        "east": position.EastWest * FEET_TO_METER,
        "north": position.NorthSouth * FEET_TO_METER,
        "elevation": position.Elevation * FEET_TO_METER,
        "angle": math.degrees(position.Angle)
    }


def get_site_location(doc):
    """Haal de Revit Site Location op (Manage > Location)."""
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit (IN_REVIT={0})".format(IN_REVIT))
    
    site = doc.SiteLocation
    
    return {
        "latitude": math.degrees(site.Latitude),
        "longitude": math.degrees(site.Longitude),
        "place_name": site.PlaceName,
        "time_zone": site.TimeZone
    }


def set_survey_point(doc, east_m, north_m, elevation_m=0.0, angle_deg=0.0):
    """Stel Survey Point in op RD coördinaten via SetProjectPosition."""
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")
    
    east_ft = east_m * METER_TO_FEET
    north_ft = north_m * METER_TO_FEET
    elev_ft = elevation_m * METER_TO_FEET
    angle_rad = math.radians(angle_deg)
    
    t = Transaction(doc, "GIS2BIM - Set Survey Point")
    t.Start()
    
    try:
        project_location = doc.ActiveProjectLocation
        
        # ProjectPosition(northSouth, eastWest, elevation, angle)
        new_position = ProjectPosition(
            north_ft,      # NorthSouth = RD Y
            east_ft,       # EastWest = RD X
            elev_ft,
            angle_rad
        )
        
        project_location.SetProjectPosition(XYZ(0, 0, 0), new_position)
        t.Commit()
        return True
        
    except Exception as e:
        t.RollBack()
        raise e


def set_project_location_rd(doc, rd_x, rd_y, elevation=0.0, angle=0.0):
    """Stel projectlocatie in met RD coördinaten."""
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")
    
    try:
        success = set_survey_point(doc, rd_x, rd_y, elevation, angle)
        
        if success:
            return {
                "success": True,
                "rd_x": rd_x,
                "rd_y": rd_y,
                "elevation": elevation,
                "angle": angle,
                "message": "Survey Point ingesteld op RD %d, %d" % (int(rd_x), int(rd_y))
            }
        else:
            return {
                "success": False,
                "message": "Kon Survey Point niet instellen"
            }
            
    except Exception as e:
        return {
            "success": False,
            "message": str(e)
        }


def get_project_location_rd(doc):
    """Haal huidige projectlocatie op in RD coördinaten."""
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")
    
    survey = get_survey_point(doc)
    
    if survey:
        return {
            "rd_x": survey["east"],
            "rd_y": survey["north"],
            "elevation": survey["elevation"],
            "angle": survey["angle"]
        }
    
    return None


def set_site_location_wgs84(doc, latitude, longitude, place_name="Project"):
    """Stel de Site Location in voor zonnestand berekeningen."""
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")
    
    try:
        t = Transaction(doc, "GIS2BIM - Set Site Location")
        t.Start()
        
        site = doc.SiteLocation
        site.Latitude = math.radians(latitude)
        site.Longitude = math.radians(longitude)
        site.PlaceName = place_name
        
        t.Commit()
        return True
        
    except Exception as e:
        print("Error setting site location: %s" % e)
        return False


def set_project_info_from_location(doc, location_data, street="", housenumber=""):
    """
    Vul Project Info parameters in met locatiegegevens van PDOK.
    Maakt ontbrekende parameters automatisch aan.
    
    Args:
        doc: Revit document
        location_data: LocationData object of dict met locatiegegevens
        street: Straatnaam
        housenumber: Huisnummer
    
    Returns:
        dict met {"filled": {param: value}, "not_found": [params], "errors": [errors], "created": [params]}
    """
    if not IN_REVIT:
        raise RuntimeError("Deze functie werkt alleen in Revit")
    
    # Converteer LocationData naar dict indien nodig
    if hasattr(location_data, 'to_dict'):
        data = location_data.to_dict()
        # Voeg extra velden toe die niet in to_dict zitten
        if hasattr(location_data, 'straatnaam'):
            data['straatnaam'] = location_data.straatnaam
        if hasattr(location_data, 'huisnummer'):
            data['huisnummer_pdok'] = location_data.huisnummer
    elif hasattr(location_data, 'rd_x'):
        # LocationData object - haal attributen direct op
        data = {
            "rd_x": location_data.rd_x,
            "rd_y": location_data.rd_y,
            "postcode": location_data.postcode,
            "gemeente": location_data.gemeente,
            "provincie": location_data.provincie,
            "kadaster_gemeente": location_data.kadaster_gemeente,
            "kadaster_sectie": location_data.kadaster_sectie,
            "kadaster_perceel": location_data.kadaster_perceel,
            "windgebied": getattr(location_data, 'windgebied', ''),
            "straatnaam": getattr(location_data, 'straatnaam', ''),
            "huisnummer_pdok": getattr(location_data, 'huisnummer', ''),
        }
    else:
        data = dict(location_data)
    
    # Gebruik straatnaam van PDOK als niet handmatig opgegeven
    if not street and data.get("straatnaam"):
        street = data.get("straatnaam")
    
    # Gebruik huisnummer van PDOK als niet handmatig opgegeven
    if not housenumber and data.get("huisnummer_pdok"):
        housenumber = data.get("huisnummer_pdok")
    
    # Bouw volledig adres
    address_parts = []
    if street:
        address_parts.append(street)
    if housenumber:
        address_parts.append(str(housenumber))
    
    if address_parts:
        full_address = " ".join(address_parts) + ", " + data.get("postcode", "") + " " + data.get("gemeente", "")
    else:
        full_address = data.get("postcode", "") + " " + data.get("gemeente", "")
    
    # Mortoncode berekenen (echte Z-order curve, niet X,Y string)
    rd_x = data.get("rd_x", 0)
    rd_y = data.get("rd_y", 0)
    mortoncode = str(calculate_mortoncode(rd_x, rd_y))
    
    # Mapping: Revit parameter naam -> waarde
    # Standaard Revit parameters eerst, dan custom
    param_mapping = {
        # Standaard Revit parameters (bestaan altijd)
        "Project Address": full_address.strip(),
        "Project Name": data.get("gemeente", ""),
        # Custom GIS2BIM parameters (worden aangemaakt indien nodig)
        "GIS2BIM_Plaats": data.get("gemeente", ""),
        "GIS2BIM_Straat": street,
        "GIS2BIM_Huisnummer": str(housenumber) if housenumber else "",
        "GIS2BIM_Postcode": data.get("postcode", ""),
        "GIS2BIM_Provincie": data.get("provincie", ""),
        "GIS2BIM_Windgebied": str(data.get("windgebied", "")),
        "GIS2BIM_Kadaster_Gemeente": str(data.get("kadaster_gemeente", "")),
        "GIS2BIM_Kadaster_Sectie": str(data.get("kadaster_sectie", "")),
        "GIS2BIM_Kadaster_Perceel": str(data.get("kadaster_perceel", "")),
        "GIS2BIM_RD_X": str(int(rd_x)),
        "GIS2BIM_RD_Y": str(int(rd_y)),
        "GIS2BIM_Mortoncode": mortoncode,
    }
    
    # Standaard Revit parameters die NIET aangemaakt hoeven worden
    builtin_params = ["Project Address", "Project Name"]
    
    project_info = doc.ProjectInformation
    
    filled = {}
    not_found = []
    errors = []
    created = []
    
    # === FASE 1: Maak ontbrekende custom parameters aan ===
    params_to_create = []
    for param_name in param_mapping.keys():
        if param_name in builtin_params:
            continue  # Skip standaard parameters
        
        param = project_info.LookupParameter(param_name)
        if param is None:
            params_to_create.append(param_name)
    
    if params_to_create:
        print("GIS2BIM: Parameters om aan te maken: {0}".format(params_to_create))
        # Start transactie voor parameter aanmaak
        t_create = Transaction(doc, "GIS2BIM - Create Parameters")
        t_create.Start()
        
        try:
            for param_name in params_to_create:
                success = _create_project_parameter(doc, param_name)
                if success:
                    created.append(param_name)
                else:
                    not_found.append(param_name)
            
            t_create.Commit()
            
            # Refresh project info reference na parameter aanmaak
            project_info = doc.ProjectInformation
            
        except Exception as e:
            t_create.RollBack()
            errors.append("Fout bij aanmaken parameters: {0}".format(str(e)))
    
    # === FASE 2: Vul parameters in ===
    t = Transaction(doc, "GIS2BIM - Set Project Parameters")
    t.Start()
    
    try:
        for param_name, value in param_mapping.items():
            if not value:
                continue
            
            param = project_info.LookupParameter(param_name)
            
            if param is None:
                if param_name not in not_found and param_name not in builtin_params:
                    not_found.append(param_name)
                continue
            
            if param.IsReadOnly:
                errors.append("{0}: read-only".format(param_name))
                continue
            
            try:
                param.Set(str(value))
                filled[param_name] = value
            except Exception as ex:
                errors.append("{0}: {1}".format(param_name, str(ex)))
        
        t.Commit()
        
    except Exception as e:
        t.RollBack()
        return {"filled": {}, "not_found": [], "errors": [str(e)], "created": []}
    
    return {"filled": filled, "not_found": not_found, "errors": errors, "created": created}
