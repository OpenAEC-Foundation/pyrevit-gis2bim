# -*- coding: utf-8 -*-
"""Revit interface modules voor GIS2BIM."""

from .location import (
    get_survey_point,
    set_survey_point,
    set_project_location_rd,
    get_project_location_rd,
    get_rd_from_project_params,
    set_site_location_wgs84,
    get_site_location,
    set_project_info_from_location
)

from .geometry import (
    rd_to_revit_xyz,
    meters_to_feet,
    feet_to_meters,
    create_model_lines,
    create_model_lines_from_features,
    create_text_notes,
    create_text_notes_from_features,
    create_filled_regions,
    create_filled_regions_from_features,
    get_line_style,
    get_text_type,
    get_filled_region_type
)

from .sheets import (
    get_sheet_bounds,
    calculate_grid_layout,
    calculate_grid_position,
    place_image_on_sheet,
    place_label_on_sheet,
    find_a3_titleblock,
    find_any_titleblock,
    populate_sheets_dropdown
)

from .families import (
    PlacementStrategy,
    find_family_symbol,
    load_family_from_file,
    get_bundled_family_path,
    resolve_family_or_fallback,
    place_element,
    create_tree_directshape,
    create_lamp_directshape,
)
