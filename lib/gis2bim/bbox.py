# -*- coding: utf-8 -*-
"""
Bounding Box Utilities
======================

Hulpmiddelen voor bounding box berekeningen in RD coördinaten.
"""


class BoundingBox:
    """
    Bounding box in RD coördinaten.
    
    Attributes:
        xmin: Minimum X (west)
        ymin: Minimum Y (zuid)
        xmax: Maximum X (oost)
        ymax: Maximum Y (noord)
    """
    
    def __init__(self, xmin, ymin, xmax, ymax):
        self.xmin = xmin
        self.ymin = ymin
        self.xmax = xmax
        self.ymax = ymax
    
    @classmethod
    def from_center(cls, x, y, width, height):
        """Maak bounding box vanuit middenpunt en afmetingen."""
        half_w = width / 2
        half_h = height / 2
        return cls(
            xmin=x - half_w,
            ymin=y - half_h,
            xmax=x + half_w,
            ymax=y + half_h
        )
    
    @classmethod
    def from_point_radius(cls, x, y, radius):
        """Maak vierkante bounding box vanuit punt en radius."""
        return cls.from_center(x, y, radius * 2, radius * 2)
    
    @property
    def width(self):
        """Breedte van de bounding box."""
        return self.xmax - self.xmin
    
    @property
    def height(self):
        """Hoogte van de bounding box."""
        return self.ymax - self.ymin
    
    @property
    def center(self):
        """Middenpunt (x, y) van de bounding box."""
        return (
            (self.xmin + self.xmax) / 2,
            (self.ymin + self.ymax) / 2
        )
    
    @property
    def area(self):
        """Oppervlakte in vierkante meters."""
        return self.width * self.height
    
    def to_wkt(self):
        """Converteer naar WKT POLYGON string."""
        return (
            "POLYGON(({0} {1}, {2} {1}, "
            "{2} {3}, {0} {3}, {0} {1}))".format(
                self.xmin, self.ymin, self.xmax, self.ymax
            )
        )
    
    def to_tuple(self):
        """Retourneer als tuple (xmin, ymin, xmax, ymax)."""
        return (self.xmin, self.ymin, self.xmax, self.ymax)
    
    def expand(self, margin):
        """Vergroot bounding box met marge aan alle zijden."""
        return BoundingBox(
            xmin=self.xmin - margin,
            ymin=self.ymin - margin,
            xmax=self.xmax + margin,
            ymax=self.ymax + margin
        )
    
    def contains(self, x, y):
        """Check of punt binnen bounding box valt."""
        return (self.xmin <= x <= self.xmax and 
                self.ymin <= y <= self.ymax)
    
    def intersects(self, other):
        """Check of twee bounding boxes overlappen."""
        return not (
            self.xmax < other.xmin or
            self.xmin > other.xmax or
            self.ymax < other.ymin or
            self.ymin > other.ymax
        )
    
    def __repr__(self):
        return "BoundingBox({0:.0f}, {1:.0f}, {2:.0f}, {3:.0f})".format(
            self.xmin, self.ymin, self.xmax, self.ymax
        )


def create_bbox(x, y, size):
    """
    Shortcut voor vierkante bounding box rond punt.
    
    Args:
        x: X coördinaat middenpunt
        y: Y coördinaat middenpunt  
        size: Grootte (breedte en hoogte) in meters
        
    Returns:
        BoundingBox object
    """
    return BoundingBox.from_center(x, y, size, size)
