# -*- coding: utf-8 -*-
"""
Coordinate Transformations
==========================

Transformaties tussen coördinatensystemen:
- EPSG:28992 (Rijksdriehoekstelsel / RD New)
- EPSG:4326 (WGS84 - lat/lon)

Gebaseerd op: GIS2BIM_TransformCRS_epsg.dyf

Example:
    from gis2bim.coordinates import rd_to_wgs84, wgs84_to_rd
    lat, lon = rd_to_wgs84(155000, 463000)
"""

import math

# Constants for RD projection (Bessel 1841 ellipsoid)
X0 = 155000.0  # False Easting
Y0 = 463000.0  # False Northing
PHI0 = 52.15517440  # Latitude origin
LAM0 = 5.38720621   # Longitude origin

# Transformation coefficients RD -> WGS84
KP = [0, 2, 0, 2, 0, 2, 1, 4, 2, 4, 1]
KQ = [1, 0, 2, 1, 3, 2, 0, 0, 3, 1, 1]
KPQ = [
    3235.65389, -32.58297, -0.24750, -0.84978, -0.06550,
    -0.01709, -0.00738, 0.00530, -0.00039, 0.00033, -0.00012
]

LP = [1, 1, 1, 3, 1, 3, 0, 3, 1, 0, 2, 5]
LQ = [0, 1, 2, 0, 3, 1, 1, 2, 4, 2, 0, 0]
LPQ = [
    5260.52916, 105.94684, 2.45656, -0.81885, 0.05594,
    -0.05607, 0.01199, -0.00256, 0.00128, 0.00022, -0.00022, 0.00026
]


def rd_to_wgs84(x, y):
    """
    Converteer RD naar WGS84 coördinaten.
    
    Args:
        x: RD X-coördinaat (Easting) in meters
        y: RD Y-coördinaat (Northing) in meters
        
    Returns:
        Tuple van (latitude, longitude) in decimale graden
    """
    dx = (x - X0) * 1e-5
    dy = (y - Y0) * 1e-5
    
    # Calculate latitude
    phi = PHI0
    for i in range(len(KP)):
        phi += KPQ[i] * (dx ** KP[i]) * (dy ** KQ[i]) / 3600
    
    # Calculate longitude
    lam = LAM0
    for i in range(len(LP)):
        lam += LPQ[i] * (dx ** LP[i]) * (dy ** LQ[i]) / 3600
    
    return phi, lam


# Inverse transformation coefficients WGS84 -> RD
RP = [0, 1, 2, 0, 1, 3, 1, 0, 2]
RQ = [1, 1, 1, 3, 0, 1, 3, 2, 3]
RPQ = [
    190094.945, -11832.228, -114.221, -32.391, -0.705,
    -2.340, -0.608, -0.008, 0.148
]

SP = [1, 0, 2, 1, 3, 0, 2, 1, 0, 1]
SQ = [0, 2, 0, 2, 0, 1, 2, 1, 4, 4]
SPQ = [
    309056.544, 3638.893, 73.077, -157.984, 59.788,
    0.433, -6.439, -0.032, 0.092, -0.054
]


def wgs84_to_rd(lat, lon):
    """
    Converteer WGS84 naar RD coördinaten.
    
    Args:
        lat: Latitude in decimale graden
        lon: Longitude in decimale graden
        
    Returns:
        Tuple van (x, y) RD coördinaten in meters
    """
    dphi = 0.36 * (lat - PHI0)
    dlam = 0.36 * (lon - LAM0)
    
    # Calculate X
    x = X0
    for i in range(len(RP)):
        x += RPQ[i] * (dphi ** RP[i]) * (dlam ** RQ[i])
    
    # Calculate Y
    y = Y0
    for i in range(len(SP)):
        y += SPQ[i] * (dphi ** SP[i]) * (dlam ** SQ[i])
    
    return x, y


def create_bbox_rd(center_x, center_y, width, height=None):
    """
    Maak een bounding box rond een centerpunt in RD coördinaten.
    
    Args:
        center_x: RD X-coördinaat van het centrum
        center_y: RD Y-coördinaat van het centrum
        width: Breedte van de bbox in meters
        height: Hoogte van de bbox in meters (default: gelijk aan width)
        
    Returns:
        Tuple van (xmin, ymin, xmax, ymax)
    """
    if height is None:
        height = width
        
    half_w = width / 2
    half_h = height / 2
    
    return (
        center_x - half_w,
        center_y - half_h,
        center_x + half_w,
        center_y + half_h
    )


def bbox_to_polygon_wkt(xmin, ymin, xmax, ymax):
    """
    Converteer bounding box naar WKT POLYGON string.
    
    Returns:
        WKT POLYGON string
    """
    return (
        "POLYGON (({0} {1}, "
        "{2} {1}, "
        "{2} {3}, "
        "{0} {3}, "
        "{0} {1}))".format(xmin, ymin, xmax, ymax)
    )


def distance_rd(x1, y1, x2, y2):
    """
    Bereken afstand tussen twee RD punten in meters.
    """
    return math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
