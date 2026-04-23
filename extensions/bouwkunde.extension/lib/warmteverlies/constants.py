# -*- coding: utf-8 -*-
"""Constanten en defaults voor ISSO 51 warmteverliesberekening export.

Alle waarden conform ISSO 51:2023/2024 en EN ISO 6946.
"""

# =============================================================================
# Eenhedenconversie
# =============================================================================
FEET_TO_M = 0.3048
SQFT_TO_M2 = 0.09290304
CUFT_TO_M3 = 0.02831685

# =============================================================================
# Default U-waarden [W/(m2*K)]
# Conservatieve waarden wanneer Revit geen betrouwbare data levert.
# =============================================================================
DEFAULT_U_VALUES = {
    "exterior_wall": 0.21,
    "interior_wall": 2.78,
    "floor_ground": 0.21,
    "floor_interior": 2.50,
    "roof": 0.15,
    "ceiling_interior": 2.50,
    "window": 1.60,
    "door_exterior": 1.70,
    "door_interior": 2.50,
}

# =============================================================================
# Oppervlakteweerstanden [m2*K/W] conform EN ISO 6946:2017
# =============================================================================
RSI_HORIZONTAL = 0.13  # Binnenzijde, horizontale warmtestroom (wanden)
RSE_HORIZONTAL = 0.04  # Buitenzijde, horizontale warmtestroom
RSI_UPWARD = 0.10  # Binnenzijde, opwaartse warmtestroom (plafond)
RSE_UPWARD = 0.04  # Buitenzijde, opwaarts
RSI_DOWNWARD = 0.17  # Binnenzijde, neerwaartse warmtestroom (vloer)
RSE_DOWNWARD = 0.04  # Buitenzijde, neerwaarts

# Ground heeft alleen Rsi (geen Rse)
RSI_GROUND = 0.17
RSE_GROUND = 0.00

# =============================================================================
# Ontwerptemperaturen per RoomFunction [graden C]
# =============================================================================
DESIGN_TEMPERATURES = {
    "living_room": 20.0,
    "kitchen": 20.0,
    "bedroom": 20.0,
    "bathroom": 22.0,
    "toilet": 15.0,
    "hallway": 15.0,
    "landing": 15.0,
    "storage": 5.0,
    "attic": 20.0,
    "custom": 20.0,
}

# =============================================================================
# Default klimaatgegevens
# =============================================================================
DEFAULT_THETA_E = -10.0
DEFAULT_THETA_B_RESIDENTIAL = 17.0
DEFAULT_THETA_B_NON_RESIDENTIAL = 14.0
DEFAULT_WIND_FACTOR = 1.0

# =============================================================================
# Default grondparameters
# =============================================================================
DEFAULT_FG2 = 1.0
DEFAULT_GROUND_WATER_FACTOR = 1.0
DEFAULT_U_EQUIVALENT = 0.21

# =============================================================================
# Default gebouwinstellingen
# =============================================================================
DEFAULT_QV10 = 150.0
DEFAULT_BUILDING_HEIGHT = 6.0
DEFAULT_ROOM_HEIGHT = 2.6
DEFAULT_NUM_FLOORS = 2

# =============================================================================
# Minimale oppervlaktefilter
# Filtert kolom-sneden, aansluitdetails en andere micro-vlakken
# =============================================================================
MIN_FACE_AREA_M2 = 0.05

# =============================================================================
# Revit eenheden conversie — thermische geleidbaarheid
# =============================================================================
# Revit ThermalConductivity is in BTU*in/(hr*ft2*degF)
# Conversie naar W/(m*K): / 6.93347
REVIT_LAMBDA_DIVISOR = 6.93347

# =============================================================================
# Wandassemblage detectie
# =============================================================================
MAX_ASSEMBLY_GAP_M = 0.30
PARALLEL_COS_TOLERANCE = 0.985
MIN_OVERLAP_FRACTION = 0.50
MAX_ASSEMBLY_DEPTH = 5

# =============================================================================
# Raycast scanner
# =============================================================================
RAY_HEIGHT_STEP_M = 0.1       # 100mm scanstap per hoogte
RAY_MAX_DIST_M = 3.0          # Max ray afstand van room center
MIN_CAVITY_MM = 10             # Min gap voor spouw detectie
MAX_CAVITY_MM = 300            # Max gap voor spouw (>300mm = geen constructie)
CONSTRUCTION_CATEGORIES = {    # Revit categorieen die constructie-elementen zijn
    -2000011: "Wall",
    -2000032: "Floor",
    -2000035: "Roof",
    -2000038: "Ceiling",
}
OPENING_CATEGORIES = {         # Categorieen die als opening behandeld worden
    -2000014: "window",        # OST_Windows
    -2000023: "door",          # OST_Doors
    -2000170: "curtain_wall",  # OST_CurtainWallPanels
}
IGNORE_CATEGORIES = {          # Categorieen die gefilterd worden uit ray hits
    -2001352: "Furniture",
    -2001370: "Casework",
    -2001040: "GenericModel",
    -2001060: "ElectricalFixtures",
}
WATER_MATERIAL_KEYWORDS = ["water", "meer", "rivier", "zee"]
GROUND_MATERIAL_KEYWORDS = ["grond", "earth", "soil", "klei", "zand", "topo"]

# =============================================================================
# Debug flags
# =============================================================================
DEBUG_OPENINGS = False            # Log opening detection per boundary wall
