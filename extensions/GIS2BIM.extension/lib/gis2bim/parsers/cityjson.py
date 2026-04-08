# -*- coding: utf-8 -*-
"""
CityJSON Parser
===============

Parser voor CityJSON bestanden (v1.x en v2.0).
Ondersteunt alle objecttypen: Building, Road, WaterBody, TINRelief, etc.

Workflow:
1. Laad JSON bestand (IronPython json module)
2. Decomprimeer vertices (scale + translate)
3. Selecteer geometry op LoD
4. Converteer Solid/MultiSurface boundaries naar face indices
5. Merge Building + BuildingPart geometry

Gebruik:
    from gis2bim.parsers.cityjson import CityJSONParser, CityJSONObject

    parser = CityJSONParser()
    objects = parser.parse_file("tile.json", target_lod="2.2")
"""

import json
import math
import os


class CityJSONError(Exception):
    """Fout bij het parsen van CityJSON data."""
    pass


def _jtoken_to_python(token):
    """Converteer Newtonsoft.Json JToken naar Python dict/list/waarde.

    Recursieve conversie zodat parse_data() normaal kan werken
    met .get(), .items(), indexing, etc.

    Args:
        token: Newtonsoft.Json.Linq.JToken

    Returns:
        Python dict, list, str, int, float, bool, of None
    """
    from Newtonsoft.Json.Linq import JTokenType

    jtype = token.Type

    if jtype == JTokenType.Object:
        result = {}
        for prop in token.Properties():
            result[str(prop.Name)] = _jtoken_to_python(prop.Value)
        return result

    elif jtype == JTokenType.Array:
        return [_jtoken_to_python(item) for item in token]

    elif jtype == JTokenType.Integer:
        return int(token.Value)

    elif jtype == JTokenType.Float:
        return float(token.Value)

    elif jtype == JTokenType.String:
        return str(token.Value)

    elif jtype == JTokenType.Boolean:
        return bool(token.Value)

    elif jtype == JTokenType.Null or jtype == JTokenType.Undefined:
        return None

    else:
        return str(token)


# Kleurenschema per CityJSON objecttype (RGB 0-255)
CITYJSON_KLEUREN = {
    "Building":           (200, 180, 160),   # Beige
    "BuildingPart":       (200, 180, 160),   # Beige
    "Road":               (100, 100, 100),   # Donkergrijs
    "Railway":            (60, 60, 60),      # Zeer donkergrijs
    "LandUse":            (150, 200, 100),   # Lichtgroen
    "PlantCover":         (80, 160, 60),     # Groen
    "WaterBody":          (100, 150, 220),   # Blauw
    "Waterway":           (100, 150, 220),   # Blauw
    "Bridge":             (180, 180, 180),   # Grijs
    "GenericCityObject":  (200, 200, 200),   # Lichtgrijs
    "TINRelief":          (160, 140, 120),   # Bruin
}

# Mapping van objecttype naar materiaal categorie
MATERIAL_CATEGORIE = {
    "Building":           "Gebouw",
    "BuildingPart":       "Gebouw",
    "Road":               "Weg",
    "Railway":            "Weg",
    "LandUse":            "Terrein",
    "PlantCover":         "Groen",
    "WaterBody":          "Water",
    "Waterway":           "Water",
    "Bridge":             "Overig",
    "GenericCityObject":  "Overig",
    "TINRelief":          "Terrein",
}


class CityJSONObject(object):
    """Container voor een geparsed CityJSON object.

    Attributes:
        obj_id: Unieke ID (str)
        obj_type: CityJSON type ("Building", "Road", etc.)
        attributes: Dict met bouwjaar, status, etc.
        vertices: Lijst van (x, y, z) tuples (RD + NAP, in meters)
        faces: Lijst van tuples met vertex indices (0-based)
        lod: Geselecteerde LoD string ("1.2", "2.2", etc.)
        children_ids: Lijst van child object IDs
    """

    def __init__(self, obj_id, obj_type):
        self.obj_id = obj_id
        self.obj_type = obj_type
        self.attributes = {}
        self.vertices = []
        self.faces = []
        self.lod = ""
        self.children_ids = []

    def get_material_name(self):
        """Geef materiaal naam voor Revit (bijv. 'CityJSON - Gebouw')."""
        cat = MATERIAL_CATEGORIE.get(self.obj_type, "Overig")
        return "CityJSON - {0}".format(cat)

    def get_color(self):
        """Geef RGB kleur tuple voor dit objecttype."""
        return CITYJSON_KLEUREN.get(self.obj_type, (200, 200, 200))

    def is_building(self):
        """Check of dit een gebouw(-deel) is."""
        return self.obj_type in ("Building", "BuildingPart")

    def __repr__(self):
        return "CityJSONObject({0}, {1}, {2} verts, {3} faces)".format(
            self.obj_id, self.obj_type,
            len(self.vertices), len(self.faces))


class CityJSONParser(object):
    """Parser voor CityJSON bestanden.

    Ondersteunt CityJSON v1.x (met "transform") en v2.0.
    """

    # Bestandsgrootte limiet voor IronPython json module (bytes)
    _PYTHON_JSON_MAX_SIZE = 50 * 1024 * 1024  # 50 MB

    def parse_file(self, filepath, target_lod="2.2", bbox=None):
        """Parse een CityJSON bestand.

        Voor grote bestanden (>50MB) wordt automatisch .NET's
        JavaScriptSerializer gebruikt i.p.v. IronPython's json module.

        Args:
            filepath: Pad naar .json of .city.json bestand
            target_lod: Gewenste LoD ("1.2", "1.3", "2.2")
            bbox: Optionele (xmin, ymin, xmax, ymax) filter in RD

        Returns:
            Lijst van CityJSONObject instanties

        Raises:
            CityJSONError: Bij parse fouten
        """
        try:
            file_size = os.path.getsize(filepath)
        except OSError:
            file_size = 0

        # Grote bestanden: gebruik .NET JSON parser (minder geheugen)
        if file_size > self._PYTHON_JSON_MAX_SIZE:
            data = self._load_json_dotnet(filepath, file_size)
        else:
            data = self._load_json_python(filepath)

        return self.parse_data(data, target_lod=target_lod, bbox=bbox)

    def _load_json_python(self, filepath):
        """Laad JSON met Python's json module."""
        try:
            with open(filepath, "r") as f:
                return json.load(f)
        except (IOError, OSError) as e:
            raise CityJSONError(
                "Kan bestand niet openen: {0}".format(e))
        except (ValueError, MemoryError) as e:
            # MemoryError: probeer .NET fallback
            if "MemoryError" in type(e).__name__ or "OutOfMemory" in str(e):
                return self._load_json_dotnet(filepath, 0)
            raise CityJSONError(
                "Ongeldig JSON formaat: {0}".format(e))

    def _load_json_dotnet(self, filepath, file_size):
        """Laad groot JSON bestand met .NET I/O + fallbacks.

        Strategie:
        1. .NET File.ReadAllText + Python json.loads (vermijdt Python
           file I/O overhead die extra geheugen kost)
        2. Newtonsoft.Json (meegeleverd met Revit 2022+)
        3. Duidelijke foutmelding

        Args:
            filepath: Pad naar JSON bestand
            file_size: Bestandsgrootte in bytes

        Returns:
            Parsed JSON data

        Raises:
            CityJSONError: Bij parse of geheugen fouten
        """
        size_mb = file_size / (1024.0 * 1024.0)

        # Strategie 1: .NET I/O + Python json.loads
        # .NET File.ReadAllText is efficienter dan Python's open/read
        try:
            from System.IO import File as DotNetFile
            text = DotNetFile.ReadAllText(filepath)
            return json.loads(text)
        except Exception:
            pass

        # Strategie 2: Newtonsoft.Json (Revit 2022+)
        try:
            import clr
            clr.AddReference('Newtonsoft.Json')
            from Newtonsoft.Json.Linq import JObject as NJObject
            from System.IO import File as DotNetFile

            text = DotNetFile.ReadAllText(filepath)
            jobj = NJObject.Parse(text)
            return _jtoken_to_python(jobj)
        except Exception:
            pass

        raise CityJSONError(
            "CityJSON bestand is te groot ({0:.0f} MB).\n"
            "Onvoldoende geheugen beschikbaar.\n\n"
            "Tip: gebruik 'Alleen gebouwen' i.p.v. "
            "'Gebouwen + Terreinen' voor kleinere bestanden.".format(
                size_mb))

    def parse_data(self, data, target_lod="2.2", bbox=None):
        """Parse een al-geladen CityJSON dict.

        Args:
            data: Dict van json.load (CityJSON root)
            target_lod: Gewenste LoD ("1.2", "1.3", "2.2")
            bbox: Optionele (xmin, ymin, xmax, ymax) filter in RD

        Returns:
            Lijst van CityJSONObject instanties

        Raises:
            CityJSONError: Bij parse fouten
        """
        # Valideer CityJSON
        cj_type = data.get("type", "")
        if cj_type != "CityJSON":
            raise CityJSONError(
                "Geen CityJSON bestand (type={0})".format(cj_type))

        version = data.get("version", "1.0")

        # Decomprimeer vertices
        raw_vertices = data.get("vertices", [])
        transform = data.get("transform", None)

        if transform:
            all_vertices = self._decompress_vertices(raw_vertices, transform)
        else:
            # Geen transform: vertices zijn al floats
            all_vertices = [(v[0], v[1], v[2]) for v in raw_vertices]

        if not all_vertices:
            raise CityJSONError("Geen vertices gevonden in CityJSON")

        # Parse alle CityObjects
        city_objects = data.get("CityObjects", {})
        parsed = {}  # obj_id -> CityJSONObject

        for obj_id, obj_data in city_objects.items():
            obj_type = obj_data.get("type", "GenericCityObject")
            attributes = obj_data.get("attributes", {})
            children = obj_data.get("children", [])
            parents = obj_data.get("parents", [])
            geometries = obj_data.get("geometry", [])

            # Selecteer beste geometry voor target LoD
            geom = self._select_geometry(geometries, target_lod)
            if geom is None:
                # Geen geometry (bijv. Building parent zonder eigen geom)
                obj = CityJSONObject(obj_id, obj_type)
                obj.attributes = attributes
                obj.children_ids = children
                obj.lod = target_lod
                parsed[obj_id] = obj
                continue

            # Converteer geometry naar faces
            faces = self._geometry_to_faces(geom)

            # Verzamel gebruikte vertex indices
            used_indices = set()
            for face in faces:
                for idx in face:
                    used_indices.add(idx)

            # Maak object aan met alle vertices (indices blijven globaal)
            obj = CityJSONObject(obj_id, obj_type)
            obj.attributes = attributes
            obj.children_ids = children
            obj.lod = str(geom.get("lod", target_lod))
            obj.vertices = all_vertices  # Referentie naar globale lijst
            obj.faces = faces

            parsed[obj_id] = obj

        # Merge Building + BuildingPart
        result = self._merge_buildings(parsed, all_vertices)

        # Filter op bbox
        if bbox:
            result = [obj for obj in result
                      if self._check_bbox(obj.vertices, obj.faces, bbox)]

        return result

    def _decompress_vertices(self, raw_vertices, transform):
        """Pas scale + translate toe op integer vertices.

        CityJSON comprimeert vertices als integers:
            real = integer * scale + translate

        Args:
            raw_vertices: Lijst van [int, int, int]
            transform: Dict met "scale" en "translate"

        Returns:
            Lijst van (float, float, float) tuples
        """
        scale = transform.get("scale", [1, 1, 1])
        translate = transform.get("translate", [0, 0, 0])

        sx, sy, sz = scale[0], scale[1], scale[2]
        tx, ty, tz = translate[0], translate[1], translate[2]

        result = []
        for v in raw_vertices:
            x = v[0] * sx + tx
            y = v[1] * sy + ty
            z = v[2] * sz + tz
            result.append((x, y, z))

        return result

    def _select_geometry(self, geometries, target_lod):
        """Kies de beste geometry: hoogste LoD <= target_lod.

        Args:
            geometries: Lijst van CityJSON geometry objecten
            target_lod: Gewenste LoD string ("1.2", "2.2", etc.)

        Returns:
            Geometry dict of None
        """
        if not geometries:
            return None

        target = self._lod_to_float(target_lod)

        # Filter en sorteer op LoD
        candidates = []
        for geom in geometries:
            lod_val = self._lod_to_float(str(geom.get("lod", "0")))
            if lod_val <= target:
                candidates.append((lod_val, geom))

        if not candidates:
            # Geen LoD <= target: neem laagste beschikbare
            for geom in geometries:
                lod_val = self._lod_to_float(str(geom.get("lod", "0")))
                candidates.append((lod_val, geom))

        if not candidates:
            return None

        # Kies hoogste LoD
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]

    def _lod_to_float(self, lod_str):
        """Converteer LoD string naar float voor vergelijking."""
        try:
            return float(lod_str)
        except (ValueError, TypeError):
            return 0.0

    def _geometry_to_faces(self, geometry):
        """Converteer CityJSON geometry boundaries naar face indices.

        Ondersteunt:
        - Solid: boundaries[shell][face][ring] -> gebruik outer shell
        - MultiSurface: boundaries[face][ring]
        - CompositeSurface: boundaries[face][ring]
        - MultiSolid: boundaries[solid][shell][face][ring]
        - CompositeSolid: boundaries[solid][shell][face][ring]

        Fan-triangulatie voor polygonen met >3 vertices.

        Args:
            geometry: CityJSON geometry dict met "type" en "boundaries"

        Returns:
            Lijst van tuples met vertex indices (triangles)
        """
        geom_type = geometry.get("type", "")
        boundaries = geometry.get("boundaries", [])
        faces = []

        if geom_type == "Solid":
            # boundaries = [shell[face[ring[vertex_idx]]]]
            # Gebruik outer shell = boundaries[0]
            if boundaries:
                outer_shell = boundaries[0]
                for face_rings in outer_shell:
                    self._add_face_rings(face_rings, faces)

        elif geom_type in ("MultiSurface", "CompositeSurface"):
            # boundaries = [face[ring[vertex_idx]]]
            for face_rings in boundaries:
                self._add_face_rings(face_rings, faces)

        elif geom_type in ("MultiSolid", "CompositeSolid"):
            # boundaries = [solid[shell[face[ring[vertex_idx]]]]]
            for solid in boundaries:
                if solid:
                    outer_shell = solid[0]
                    for face_rings in outer_shell:
                        self._add_face_rings(face_rings, faces)

        return faces

    def _add_face_rings(self, face_rings, faces):
        """Voeg een face toe vanuit ring-gebaseerde CityJSON boundaries.

        Gebruikt alleen de outer ring (face_rings[0]).
        Voert fan-triangulatie uit voor polygonen met >3 vertices.

        Args:
            face_rings: [[outer_ring_indices], [hole_indices...]]
            faces: Output lijst waar triangles aan worden toegevoegd
        """
        if not face_rings:
            return

        # Outer ring
        ring = face_rings[0]
        if len(ring) < 3:
            return

        if len(ring) == 3:
            faces.append(tuple(ring))
        else:
            # Fan-triangulatie: v0 als pivot
            v0 = ring[0]
            for i in range(1, len(ring) - 1):
                faces.append((v0, ring[i], ring[i + 1]))

    def _merge_buildings(self, parsed, all_vertices):
        """Merge Building parents met hun BuildingPart children.

        Buildings in CityJSON hebben vaak geen eigen geometry.
        De geometry zit in de BuildingPart children.

        Strategie:
        1. Voor Buildings met children: merge BuildingPart faces
        2. Attributen: gebruik Building attributen (bouwjaar, status)
        3. Retourneer als 1 CityJSONObject per Building

        Args:
            parsed: Dict van obj_id -> CityJSONObject
            all_vertices: Globale vertex lijst

        Returns:
            Lijst van CityJSONObject (Buildings samengevoegd, rest individueel)
        """
        result = []
        used_as_child = set()

        for obj_id, obj in parsed.items():
            if obj.obj_type != "Building":
                continue

            if not obj.children_ids:
                # Building zonder children: gebruik eigen geometry
                if obj.faces:
                    result.append(obj)
                continue

            # Building met children: merge
            merged_faces = list(obj.faces)  # Start met eigen faces (mogelijk leeg)

            for child_id in obj.children_ids:
                used_as_child.add(child_id)
                child = parsed.get(child_id)
                if child and child.faces:
                    merged_faces.extend(child.faces)

            if merged_faces:
                merged = CityJSONObject(obj_id, "Building")
                merged.attributes = obj.attributes
                merged.vertices = all_vertices
                merged.faces = merged_faces
                merged.lod = obj.lod
                merged.children_ids = obj.children_ids
                result.append(merged)

        # Voeg niet-gebouw objecten en orphan BuildingParts toe
        for obj_id, obj in parsed.items():
            if obj_id in used_as_child:
                continue
            if obj.obj_type == "Building":
                continue  # Al verwerkt
            if obj.faces:
                result.append(obj)

        return result

    def _check_bbox(self, vertices, faces, bbox):
        """Check of minstens 1 face-vertex binnen bbox valt.

        Args:
            vertices: Globale vertex lijst
            faces: Lijst van face tuples met indices
            bbox: (xmin, ymin, xmax, ymax)

        Returns:
            True als minstens 1 vertex binnen bbox valt
        """
        xmin, ymin, xmax, ymax = bbox

        # Check alleen vertices die in faces worden gebruikt
        checked = set()
        for face in faces:
            for idx in face:
                if idx in checked:
                    continue
                checked.add(idx)
                if 0 <= idx < len(vertices):
                    vx, vy, vz = vertices[idx]
                    if xmin <= vx <= xmax and ymin <= vy <= ymax:
                        return True

        return False

    def count_by_type(self, objects):
        """Tel objecten per type.

        Args:
            objects: Lijst van CityJSONObject

        Returns:
            Dict van type -> count
        """
        counts = {}
        for obj in objects:
            t = obj.obj_type
            counts[t] = counts.get(t, 0) + 1
        return counts
