# -*- coding: utf-8 -*-
"""
GIS2BIM Library
===============

Modulaire Python library voor GIS data integratie in Revit.
Compatibel met IronPython 2.7 (pyRevit).
"""

__version__ = "1.0.0"
__author__ = "OpenAEC Foundation"

from .coordinates import rd_to_wgs84, wgs84_to_rd
from .bbox import BoundingBox, create_bbox
