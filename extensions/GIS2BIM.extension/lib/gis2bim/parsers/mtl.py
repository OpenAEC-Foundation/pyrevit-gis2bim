# -*- coding: utf-8 -*-
"""
Wavefront MTL Parser
=====================

Simpele parser voor Wavefront MTL (Material Library) bestanden.
Ondersteunt diffuse kleur (Kd) en diffuse texture (map_Kd).

Gebruik:
    from gis2bim.parsers.mtl import MTLReader

    reader = MTLReader()
    materials = reader.read("model.mtl")
    # {"brick": {"name": "brick", "Kd": (0.8, 0.4, 0.2), "map_Kd": "brick.jpg"}, ...}
"""

import os


class MTLError(Exception):
    """Fout bij het parsen van een MTL bestand."""
    pass


class MTLReader(object):
    """Parser voor Wavefront MTL bestanden.

    Ondersteunt:
    - newmtl: materiaal naam
    - Kd r g b: diffuse kleur (0.0-1.0)
    - map_Kd: diffuse texture pad
    - Overgeslagen: Ka, Ks, Ns, d, Tr, illum, etc.
    """

    def read(self, filepath):
        """Parse MTL bestand naar dict van materialen.

        Args:
            filepath: Pad naar .mtl bestand

        Returns:
            Dict mapping materiaal_naam -> {
                "name": str,
                "Kd": (r, g, b) of None,  # 0.0-1.0 range
                "map_Kd": str of None,     # absoluut pad naar texture
            }

        Raises:
            MTLError: Bij lees- of parse fouten
        """
        try:
            with open(filepath, "r") as f:
                lines = f.readlines()
        except IOError as e:
            raise MTLError("Kan MTL bestand niet lezen: {0}".format(e))

        mtl_dir = os.path.dirname(os.path.abspath(filepath))
        materials = {}
        current_name = None
        current_mat = None

        for line in lines:
            line = line.strip()

            if not line or line[0] == "#":
                continue

            parts = line.split()
            keyword = parts[0]

            if keyword == "newmtl" and len(parts) >= 2:
                # Sla vorig materiaal op
                if current_name is not None and current_mat is not None:
                    materials[current_name] = current_mat

                current_name = " ".join(parts[1:])
                current_mat = {
                    "name": current_name,
                    "Kd": None,
                    "map_Kd": None,
                }

            elif keyword == "Kd" and len(parts) >= 4 and current_mat is not None:
                try:
                    r = float(parts[1])
                    g = float(parts[2])
                    b = float(parts[3])
                    current_mat["Kd"] = (r, g, b)
                except (ValueError, IndexError):
                    pass

            elif keyword == "map_Kd" and len(parts) >= 2 and current_mat is not None:
                texture_path = " ".join(parts[1:])
                # Maak absoluut pad relatief aan MTL locatie
                if not os.path.isabs(texture_path):
                    texture_path = os.path.join(mtl_dir, texture_path)
                current_mat["map_Kd"] = texture_path

        # Sla laatste materiaal op
        if current_name is not None and current_mat is not None:
            materials[current_name] = current_mat

        return materials

    def get_rgb_255(self, material):
        """Converteer Kd (0.0-1.0) naar RGB (0-255) tuple.

        Args:
            material: Materiaal dict met "Kd" key

        Returns:
            (R, G, B) tuple met waarden 0-255, of (180, 180, 180) als fallback
        """
        kd = material.get("Kd") if material else None
        if kd is None:
            return (180, 180, 180)

        return (
            int(min(255, max(0, kd[0] * 255))),
            int(min(255, max(0, kd[1] * 255))),
            int(min(255, max(0, kd[2] * 255))),
        )
