# -*- coding: utf-8 -*-
from pyrevit import DB
import traceback
import logging

logger = logging.getLogger(__name__)


def normalize_string(text):
    """Safely normalize string values"""
    if text is None:
        return "Unnamed"
    return str(text).strip()


def get_element_name(element):
    """
    Get the name of a Revit element.
    Useful for both FamilySymbol and other elements.
    """
    try:
        return element.Name
    except AttributeError:
        return DB.Element.Name.__get__(element)


def find_family_symbol_safely(doc, target_family_name, target_type_name=None):
    """
    Safely find a family symbol by name.
    Now supports case-insensitive partial matching.
    """
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
        target_family_lower = target_family_name.lower()
        target_type_lower = target_type_name.lower() if target_type_name else None

        # First try exact match
        for symbol in collector:
            try:
                fam_name = symbol.FamilyName if hasattr(symbol, 'FamilyName') else ""
                sym_name = DB.Element.Name.__get__(symbol) if symbol else ""
                
                if fam_name == target_family_name:
                    if not target_type_name or sym_name == target_type_name:
                        return symbol
            except:
                continue
        
        # Then try case-insensitive match
        collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
        for symbol in collector:
            try:
                fam_name = symbol.FamilyName if hasattr(symbol, 'FamilyName') else ""
                sym_name = DB.Element.Name.__get__(symbol) if symbol else ""
                
                if fam_name.lower() == target_family_lower:
                    if not target_type_lower or sym_name.lower() == target_type_lower:
                        return symbol
            except:
                continue
        
        # Finally try partial match (contains)
        collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
        for symbol in collector:
            try:
                fam_name = symbol.FamilyName if hasattr(symbol, 'FamilyName') else ""
                sym_name = DB.Element.Name.__get__(symbol) if symbol else ""
                
                if target_family_lower in fam_name.lower():
                    if not target_type_lower or target_type_lower in sym_name.lower():
                        return symbol
            except:
                continue
                
        return None
    except Exception as e:
        logger.error("Error finding family symbol: %s", str(e))
        return None
