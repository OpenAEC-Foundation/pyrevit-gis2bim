# -*- coding: utf-8 -*-
"""
Wavefront OBJ Parser
=====================

Simpele parser voor Wavefront OBJ bestanden.
Ondersteunt vertices, faces, en object/group namen.

Gebruik:
    from gis2bim.parsers.obj import OBJReader

    reader = OBJReader()
    meshes = reader.read("gebouwen.obj")
    # [{"name": "building_1", "vertices": [(x,y,z),...], "faces": [(0,1,2),...]}, ...]

    mesh = reader.read_as_single_mesh("gebouwen.obj")
    # {"name": "merged", "vertices": [(x,y,z),...], "faces": [(0,1,2),...]}
"""


class OBJError(Exception):
    """Fout bij het parsen van een OBJ bestand."""
    pass


class OBJReader(object):
    """Parser voor Wavefront OBJ bestanden.

    Ondersteunt:
    - v x y z: vertex posities
    - f v1 v2 v3 ...: face indices (1-based, negatieve indices)
    - f v1/vt1/vn1 ...: face indices met texture/normal (alleen vertex index gebruikt)
    - o/g: object/group namen
    - Overgeslagen: #, mtllib, usemtl, vt, vn, s, l
    """

    def read(self, filepath):
        """Parse OBJ bestand naar lijst van mesh dicts per object/group.

        Args:
            filepath: Pad naar .obj bestand

        Returns:
            Lijst van mesh dicts:
            [{"name": str, "vertices": [(x,y,z),...], "faces": [(i0,i1,i2,...),...]}, ...]

            Vertices zijn float tuples in de originele coordinaten.
            Face indices zijn 0-based.

        Raises:
            OBJError: Bij lees- of parse fouten
        """
        try:
            with open(filepath, "r") as f:
                lines = f.readlines()
        except IOError as e:
            raise OBJError("Kan OBJ bestand niet lezen: {0}".format(e))

        # Globale vertex lijst (faces refereren naar globale indices)
        all_vertices = []
        meshes = []
        current_name = "default"
        current_faces = []

        for line_num, raw_line in enumerate(lines, 1):
            line = raw_line.strip()

            # Skip lege regels en comments
            if not line or line[0] == "#":
                continue

            parts = line.split()
            keyword = parts[0]

            if keyword == "v" and len(parts) >= 4:
                # Vertex: v x y z [w]
                try:
                    x = float(parts[1])
                    y = float(parts[2])
                    z = float(parts[3])
                    all_vertices.append((x, y, z))
                except (ValueError, IndexError):
                    pass  # Skip ongeldige vertices

            elif keyword == "f" and len(parts) >= 4:
                # Face: f v1 v2 v3 ... of f v1/vt1/vn1 ...
                face_indices = []
                valid = True
                for token in parts[1:]:
                    idx = self._parse_face_index(token, len(all_vertices))
                    if idx is None:
                        valid = False
                        break
                    face_indices.append(idx)

                if valid and len(face_indices) >= 3:
                    current_faces.append(tuple(face_indices))

            elif keyword in ("o", "g"):
                # Object/Group: sla huidige mesh op en start nieuwe
                if current_faces:
                    meshes.append({
                        "name": current_name,
                        "faces": current_faces,
                    })
                    current_faces = []

                if len(parts) > 1:
                    current_name = " ".join(parts[1:])
                else:
                    current_name = "unnamed"

            # Skip: mtllib, usemtl, vt, vn, s, l

        # Laatste mesh opslaan
        if current_faces:
            meshes.append({
                "name": current_name,
                "faces": current_faces,
            })

        if not meshes:
            raise OBJError("Geen geldige mesh data gevonden in OBJ bestand")

        if not all_vertices:
            raise OBJError("Geen vertices gevonden in OBJ bestand")

        # Voeg vertices toe aan elke mesh
        # Alle meshes delen dezelfde globale vertex lijst
        for mesh in meshes:
            mesh["vertices"] = list(all_vertices)

        return meshes

    def read_as_single_mesh(self, filepath):
        """Parse OBJ bestand naar een enkele samengevoegde mesh.

        Alle object/group grenzen worden genegeerd; alle faces
        worden samengevoegd tot een enkel mesh object.

        Args:
            filepath: Pad naar .obj bestand

        Returns:
            Mesh dict:
            {"name": "merged", "vertices": [(x,y,z),...], "faces": [(i0,i1,i2,...),...]

        Raises:
            OBJError: Bij lees- of parse fouten
        """
        try:
            with open(filepath, "r") as f:
                lines = f.readlines()
        except IOError as e:
            raise OBJError("Kan OBJ bestand niet lezen: {0}".format(e))

        vertices = []
        faces = []

        for line in lines:
            line = line.strip()

            if not line or line[0] == "#":
                continue

            parts = line.split()
            keyword = parts[0]

            if keyword == "v" and len(parts) >= 4:
                try:
                    x = float(parts[1])
                    y = float(parts[2])
                    z = float(parts[3])
                    vertices.append((x, y, z))
                except (ValueError, IndexError):
                    pass

            elif keyword == "f" and len(parts) >= 4:
                face_indices = []
                valid = True
                for token in parts[1:]:
                    idx = self._parse_face_index(token, len(vertices))
                    if idx is None:
                        valid = False
                        break
                    face_indices.append(idx)

                if valid and len(face_indices) >= 3:
                    faces.append(tuple(face_indices))

        if not vertices:
            raise OBJError("Geen vertices gevonden in OBJ bestand")

        if not faces:
            raise OBJError("Geen faces gevonden in OBJ bestand")

        return {
            "name": "merged",
            "vertices": vertices,
            "faces": faces,
        }

    def _parse_face_index(self, token, vertex_count):
        """Parse een face index token.

        Ondersteunt formaten:
        - "v"           -> vertex index
        - "v/vt"        -> vertex/texture index
        - "v/vt/vn"     -> vertex/texture/normal index
        - "v//vn"       -> vertex//normal index

        Indices zijn 1-based in OBJ, worden geconverteerd naar 0-based.
        Negatieve indices worden relatief aan huidige vertex count geinterpreteerd.

        Returns:
            0-based vertex index, of None bij ongeldige input
        """
        try:
            # Split op '/' en pak eerste getal (vertex index)
            idx_str = token.split("/")[0]
            idx = int(idx_str)

            if idx > 0:
                return idx - 1  # 1-based naar 0-based
            elif idx < 0:
                return vertex_count + idx  # Negatieve index
            else:
                return None  # 0 is ongeldig in OBJ
        except (ValueError, IndexError):
            return None
