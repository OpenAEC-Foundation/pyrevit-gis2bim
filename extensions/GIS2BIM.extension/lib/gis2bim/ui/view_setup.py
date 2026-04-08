# -*- coding: utf-8 -*-
"""
View Setup - GIS2BIM
====================

Gedeelde view dropdown populatie voor GIS2BIM pyRevit tools.

Gebruik:
    from gis2bim.ui.view_setup import populate_view_dropdown

    populate_view_dropdown(self.cmb_view, doc)
"""

from pyrevit import DB
from System.Windows.Controls import ComboBoxItem


# Standaard view types voor 2D tools (BGT, WFS)
VIEW_TYPES_2D = [
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.AreaPlan,
    DB.ViewType.DraftingView,
]

# View types inclusief 3D (AHN)
VIEW_TYPES_3D = [
    DB.ViewType.ThreeD,
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.AreaPlan,
]


def populate_view_dropdown(combo, doc, view_types=None, select_active=True,
                           log=None):
    """Vul een ComboBox met beschikbare Revit views.

    Args:
        combo: WPF ComboBox element
        doc: Revit Document
        view_types: Lijst van DB.ViewType waarden om te filteren.
                    Default: VIEW_TYPES_2D
        select_active: Als True, selecteer de actieve view (default True)
        log: Optionele log functie

    Returns:
        De geselecteerde view, of None
    """
    if log is None:
        log = lambda msg: None

    if view_types is None:
        view_types = VIEW_TYPES_2D

    try:
        collector = DB.FilteredElementCollector(doc)
        views = collector.OfClass(DB.View).ToElements()

        suitable_views = []
        for view in views:
            if view.IsTemplate:
                continue
            if view.ViewType in view_types:
                suitable_views.append(view)

        suitable_views.sort(key=lambda v: v.Name)

        combo.Items.Clear()
        active_view_id = doc.ActiveView.Id

        for view in suitable_views:
            item = ComboBoxItem()
            item.Content = view.Name
            item.Tag = view.Id
            combo.Items.Add(item)

            if select_active and view.Id == active_view_id:
                combo.SelectedItem = item

        if combo.SelectedItem is None and combo.Items.Count > 0:
            combo.SelectedIndex = 0

    except Exception as e:
        log("Error loading views: {0}".format(e))


def get_selected_view(combo, doc):
    """Haal de geselecteerde view op uit een ComboBox.

    Args:
        combo: WPF ComboBox element (gevuld door populate_view_dropdown)
        doc: Revit Document

    Returns:
        View element, of None
    """
    item = combo.SelectedItem
    if item and hasattr(item, 'Tag'):
        return doc.GetElement(item.Tag)
    return None
