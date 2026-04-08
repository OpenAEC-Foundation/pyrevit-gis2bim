# -*- coding: utf-8 -*-
"""U-waarde bepaling uit Revit elementen.

Drie pogingen in volgorde:
1. CompoundStructure: laagopbouw -> R per laag -> U = 1/Rtot
2. Revit parameter: ANALYTICAL_HEAT_TRANSFER_COEFFICIENT
3. Default: lookup in constants per constructietype
"""
from Autodesk.Revit.DB import BuiltInParameter, ElementId

from warmteverlies.unit_utils import internal_to_meters, get_param_value
from warmteverlies.constants import (
    DEFAULT_U_VALUES,
    RSI_HORIZONTAL, RSE_HORIZONTAL,
    RSI_UPWARD, RSE_UPWARD,
    RSI_DOWNWARD, RSE_DOWNWARD,
    RSI_GROUND, RSE_GROUND,
    REVIT_LAMBDA_DIVISOR,
)


def get_u_value(doc, host_element, position_type="wall",
                boundary_type="exterior", element_doc=None):
    """Bepaal de U-waarde van een constructie-element.

    Args:
        doc: Revit Document (host document)
        host_element: Revit Element (Wall/Floor/Roof)
        position_type: "wall", "floor", of "ceiling"
        boundary_type: "exterior", "adjacent_room", "unheated_space", "ground"
        element_doc: Document waarin het element leeft (linked model of host).
                     Indien None wordt doc gebruikt.

    Returns:
        tuple: (u_value, source) waar source "compound"/"parameter"/"default"
    """
    if element_doc is None:
        element_doc = doc

    if host_element is None:
        return _get_default_u(position_type, boundary_type), "default"

    u_compound = _try_compound_structure(
        element_doc, host_element, position_type, boundary_type
    )
    if u_compound is not None:
        return u_compound, "compound"

    u_param = _try_parameter(host_element)
    if u_param is not None:
        return u_param, "parameter"

    return _get_default_u(position_type, boundary_type), "default"


def _try_compound_structure(doc, element, position_type, boundary_type):
    """Probeer U-waarde te berekenen uit de CompoundStructure laagopbouw."""
    try:
        elem_type = doc.GetElement(element.GetTypeId())
        if elem_type is None:
            return None

        compound = elem_type.GetCompoundStructure()
        if compound is None:
            return None

        layers = compound.GetLayers()
        if not layers or layers.Count == 0:
            return None

        total_r = 0.0
        valid_layers = 0

        for layer in layers:
            width_ft = layer.Width
            thickness_m = internal_to_meters(width_ft)

            if thickness_m <= 0:
                continue

            mat_id = layer.MaterialId
            if mat_id is None or mat_id == ElementId.InvalidElementId:
                continue

            material = doc.GetElement(mat_id)
            if material is None:
                continue

            conductivity = _get_thermal_conductivity(material)
            if conductivity is None or conductivity <= 0:
                continue

            r_layer = thickness_m / conductivity
            total_r += r_layer
            valid_layers += 1

        if valid_layers == 0 or total_r <= 0:
            return None

        rsi, rse = _get_surface_resistances(position_type, boundary_type)
        r_total = rsi + total_r + rse

        if r_total <= 0:
            return None

        return round(1.0 / r_total, 3)

    except Exception:
        return None


def _get_thermal_conductivity(material):
    """Haal lambda [W/(m*K)] op uit een Revit Material."""
    try:
        thermal_asset_id = material.ThermalAssetId
        if (thermal_asset_id is None
                or thermal_asset_id == ElementId.InvalidElementId):
            return None

        doc = material.Document
        prop_set = doc.GetElement(thermal_asset_id)
        if prop_set is None:
            return None

        thermal_asset = prop_set.GetThermalAsset()
        if thermal_asset is None:
            return None

        # Revit ThermalConductivity is in BTU*in/(hr*ft2*degF)
        # Conversie naar W/(m*K): / 6.93347
        raw = thermal_asset.ThermalConductivity
        if raw is None or raw <= 0:
            return None
        return raw / REVIT_LAMBDA_DIVISOR

    except Exception:
        return None


def _try_parameter(element):
    """Probeer U-waarde uit Revit parameter te halen."""
    for param_name in [
        "Heat Transfer Coefficient (U)",
        "Warmtedoorgangscoefficient (U)",
        "Thermal Transmittance",
    ]:
        u = get_param_value(element, param_name)
        if u is not None and u > 0:
            return u

    try:
        elem_type = element.Document.GetElement(element.GetTypeId())
        if elem_type:
            for param_name in [
                "Heat Transfer Coefficient (U)",
                "Warmtedoorgangscoefficient (U)",
                "Thermal Transmittance",
            ]:
                u = get_param_value(elem_type, param_name)
                if u is not None and u > 0:
                    return u

            param = elem_type.get_Parameter(
                BuiltInParameter.ANALYTICAL_HEAT_TRANSFER_COEFFICIENT
            )
            if param and param.HasValue:
                val = param.AsDouble()
                if val > 0:
                    return round(val * 5.678263, 3)
    except Exception:
        pass

    return None


def _get_surface_resistances(position_type, boundary_type):
    """Bepaal Rsi en Rse op basis van positie en grenstype."""
    if boundary_type == "ground":
        return RSI_GROUND, RSE_GROUND

    if position_type == "ceiling":
        rsi = RSI_UPWARD
    elif position_type == "floor":
        rsi = RSI_DOWNWARD
    else:
        rsi = RSI_HORIZONTAL

    if boundary_type in ("adjacent_room", "unheated_space"):
        rse = rsi  # Binnenzijde aan beide kanten
    else:
        if position_type == "ceiling":
            rse = RSE_UPWARD
        elif position_type == "floor":
            rse = RSE_DOWNWARD
        else:
            rse = RSE_HORIZONTAL

    return rsi, rse


def extract_layers(doc, host_element, element_doc=None):
    """Extraheer de laagopbouw uit een constructie-element.

    Args:
        doc: Revit Document (host document)
        host_element: Revit Element (Wall/Floor/Roof)
        element_doc: Document waarin het element leeft (linked model of host).
                     Indien None wordt doc gebruikt.

    Returns:
        list: Lagen met material, thickness_mm, lambda, type.
              Lege lijst als geen CompoundStructure beschikbaar.
    """
    if element_doc is None:
        element_doc = doc
    if host_element is None:
        return []

    try:
        elem_type = element_doc.GetElement(host_element.GetTypeId())
        if elem_type is None:
            return []

        compound = elem_type.GetCompoundStructure()
        if compound is None:
            return []

        layers_data = compound.GetLayers()
        if not layers_data or layers_data.Count == 0:
            return []

        result = []
        cumulative_mm = 0.0

        for layer in layers_data:
            width_ft = layer.Width
            thickness_m = internal_to_meters(width_ft)
            thickness_mm = round(thickness_m * 1000.0, 1)

            if thickness_mm <= 0:
                continue

            mat_id = layer.MaterialId
            material_name = "Onbekend"
            lambda_val = None

            if mat_id is not None and mat_id != ElementId.InvalidElementId:
                material = element_doc.GetElement(mat_id)
                if material is not None:
                    material_name = material.Name or "Onbekend"
                    conductivity = _get_thermal_conductivity(material)
                    if conductivity is not None and conductivity > 0:
                        lambda_val = round(conductivity, 4)

            # Bepaal layer type op basis van materiaal naam
            layer_type = "solid"
            if material_name and any(
                kw in material_name.lower()
                for kw in [
                    "luchtspouw", "air gap", "air space", "cavity",
                    "lucht", "spouw",
                ]
            ):
                layer_type = "air_gap"

            layer_dict = {
                "material": material_name,
                "thickness_mm": thickness_mm,
                "distance_from_interior_mm": round(cumulative_mm, 1),
                "type": layer_type,
            }
            if lambda_val is not None:
                layer_dict["lambda"] = lambda_val

            result.append(layer_dict)
            cumulative_mm += thickness_mm

        return result

    except Exception:
        return []


def _get_default_u(position_type, boundary_type):
    """Haal default U-waarde op basis van positie en grenstype."""
    if boundary_type == "ground":
        return DEFAULT_U_VALUES["floor_ground"]

    if position_type == "wall":
        if boundary_type == "exterior":
            return DEFAULT_U_VALUES["exterior_wall"]
        return DEFAULT_U_VALUES["interior_wall"]
    elif position_type == "floor":
        if boundary_type == "exterior":
            return DEFAULT_U_VALUES["floor_ground"]
        return DEFAULT_U_VALUES["floor_interior"]
    elif position_type == "ceiling":
        if boundary_type == "exterior":
            return DEFAULT_U_VALUES["roof"]
        return DEFAULT_U_VALUES["ceiling_interior"]

    return DEFAULT_U_VALUES["exterior_wall"]
