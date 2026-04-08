# -*- coding: utf-8 -*-
"""
GLB (Binary glTF 2.0) Parser
==============================

Parser voor GLB bestanden zoals gebruikt door Google 3D Tiles.
Extraheert mesh geometry (vertices en faces) uit binary glTF containers.

Gebruik:
    from gis2bim.parsers.glb import GLBReader

    reader = GLBReader()
    meshes = reader.read("tile.glb")
    # [{"name": "mesh_0", "vertices": [(x,y,z),...], "faces": [(0,1,2),...]}, ...]
"""

import json
import struct


class GLBError(Exception):
    """Fout bij het parsen van een GLB bestand."""
    pass


# glTF component type constants
COMPONENT_BYTE = 5120
COMPONENT_UBYTE = 5121
COMPONENT_SHORT = 5122
COMPONENT_USHORT = 5123
COMPONENT_UINT = 5125
COMPONENT_FLOAT = 5126

# struct format per component type
_COMPONENT_FORMAT = {
    COMPONENT_BYTE: "b",
    COMPONENT_UBYTE: "B",
    COMPONENT_SHORT: "h",
    COMPONENT_USHORT: "H",
    COMPONENT_UINT: "I",
    COMPONENT_FLOAT: "f",
}

_COMPONENT_SIZE = {
    COMPONENT_BYTE: 1,
    COMPONENT_UBYTE: 1,
    COMPONENT_SHORT: 2,
    COMPONENT_USHORT: 2,
    COMPONENT_UINT: 4,
    COMPONENT_FLOAT: 4,
}

# Components per type
_TYPE_COUNT = {
    "SCALAR": 1,
    "VEC2": 2,
    "VEC3": 3,
    "VEC4": 4,
    "MAT2": 4,
    "MAT3": 9,
    "MAT4": 16,
}

# GLB magic number
GLB_MAGIC = 0x46546C67  # "glTF"
CHUNK_JSON = 0x4E4F534A  # "JSON"
CHUNK_BIN = 0x004E4942   # "BIN\0"


class GLBReader(object):
    """Parser voor GLB (binary glTF 2.0) bestanden.

    Extraheert mesh geometry: vertex posities en triangle face indices.
    Ondersteunt alle gangbare accessor types en component formaten.
    """

    def read(self, filepath):
        """Parse GLB bestand naar lijst van mesh dicts.

        Args:
            filepath: Pad naar .glb bestand

        Returns:
            Lijst van mesh dicts:
            [{"name": str, "vertices": [(x,y,z),...], "faces": [(i0,i1,i2),...]}, ...]

            Vertices zijn float tuples in originele coordinaten (vaak ECEF).
            Faces zijn 0-based triangle indices.

        Raises:
            GLBError: Bij lees- of parse fouten
        """
        try:
            with open(filepath, "rb") as f:
                data = f.read()
        except IOError as e:
            raise GLBError("Kan GLB bestand niet lezen: {0}".format(e))

        return self.read_from_bytes(data)

    def read_from_bytes(self, data):
        """Parse GLB data uit bytes.

        Args:
            data: Bytes van het GLB bestand

        Returns:
            Lijst van mesh dicts (zie read())

        Raises:
            GLBError: Bij parse fouten
        """
        if len(data) < 12:
            raise GLBError("GLB bestand te klein (< 12 bytes)")

        # Parse header
        magic, version, total_length = struct.unpack_from("<III", data, 0)
        if magic != GLB_MAGIC:
            raise GLBError(
                "Ongeldig GLB bestand (magic: 0x{0:08X})".format(magic))
        if version not in (1, 2):
            raise GLBError(
                "Niet-ondersteunde glTF versie: {0}".format(version))

        # Parse chunks
        json_data = None
        bin_data = None
        offset = 12

        while offset < len(data) - 8:
            chunk_length, chunk_type = struct.unpack_from("<II", data, offset)
            chunk_start = offset + 8

            if chunk_type == CHUNK_JSON:
                json_bytes = data[chunk_start:chunk_start + chunk_length]
                try:
                    json_data = json.loads(json_bytes.decode("utf-8"))
                except (ValueError, UnicodeDecodeError) as e:
                    raise GLBError("Ongeldige JSON chunk: {0}".format(e))

            elif chunk_type == CHUNK_BIN:
                bin_data = data[chunk_start:chunk_start + chunk_length]

            offset = chunk_start + chunk_length
            # Chunks zijn 4-byte aligned
            if offset % 4 != 0:
                offset += 4 - (offset % 4)

        if json_data is None:
            raise GLBError("Geen JSON chunk gevonden in GLB")

        # Extract meshes
        return self._extract_meshes(json_data, bin_data)

    def _extract_meshes(self, gltf, bin_data):
        """Extract alle meshes uit glTF JSON + binary data.

        Args:
            gltf: Geparsed glTF JSON dict
            bin_data: Binary buffer data (of None)

        Returns:
            Lijst van mesh dicts
        """
        accessors = gltf.get("accessors", [])
        buffer_views = gltf.get("bufferViews", [])
        gltf_meshes = gltf.get("meshes", [])

        if not gltf_meshes:
            raise GLBError("Geen meshes gevonden in GLB")

        meshes = []

        for mesh_idx, mesh in enumerate(gltf_meshes):
            mesh_name = mesh.get("name", "mesh_{0}".format(mesh_idx))
            all_vertices = []
            all_faces = []
            vertex_offset = 0

            for prim in mesh.get("primitives", []):
                attrs = prim.get("attributes", {})

                # Vertex posities (POSITION accessor)
                pos_idx = attrs.get("POSITION")
                if pos_idx is None:
                    continue

                vertices = self._read_vec3_accessor(
                    accessors[pos_idx], buffer_views, bin_data)

                if not vertices:
                    continue

                # Face indices
                indices_idx = prim.get("indices")
                if indices_idx is not None:
                    raw_indices = self._read_scalar_accessor(
                        accessors[indices_idx], buffer_views, bin_data)

                    # Trianguleer (mode 4 = TRIANGLES is default)
                    mode = prim.get("mode", 4)
                    faces = self._indices_to_triangles(
                        raw_indices, mode, vertex_offset)
                else:
                    # Geen indices: genereer sequentiele triangles
                    faces = []
                    for i in range(0, len(vertices) - 2, 3):
                        faces.append((
                            vertex_offset + i,
                            vertex_offset + i + 1,
                            vertex_offset + i + 2,
                        ))

                all_vertices.extend(vertices)
                all_faces.extend(faces)
                vertex_offset += len(vertices)

            if all_vertices and all_faces:
                meshes.append({
                    "name": mesh_name,
                    "vertices": all_vertices,
                    "faces": all_faces,
                    "material": None,
                    "mtllib": None,
                })

        if not meshes:
            raise GLBError("Geen geldige mesh geometry gevonden in GLB")

        return meshes

    def _read_vec3_accessor(self, accessor, buffer_views, bin_data):
        """Lees VEC3 float accessor (vertex posities).

        Args:
            accessor: glTF accessor dict
            buffer_views: Lijst van bufferView dicts
            bin_data: Binary buffer bytes

        Returns:
            Lijst van (x, y, z) float tuples
        """
        if accessor.get("type") != "VEC3":
            return []

        comp_type = accessor.get("componentType", COMPONENT_FLOAT)
        count = accessor.get("count", 0)
        bv_idx = accessor.get("bufferView")

        if bv_idx is None or bin_data is None or count == 0:
            return []

        bv = buffer_views[bv_idx]
        bv_offset = bv.get("byteOffset", 0)
        bv_stride = bv.get("byteStride", 0)
        acc_offset = accessor.get("byteOffset", 0)

        start = bv_offset + acc_offset
        comp_size = _COMPONENT_SIZE.get(comp_type, 4)
        comp_fmt = _COMPONENT_FORMAT.get(comp_type, "f")
        element_size = comp_size * 3

        if bv_stride == 0:
            bv_stride = element_size

        vertices = []
        fmt = "<3{0}".format(comp_fmt)

        for i in range(count):
            pos = start + i * bv_stride
            if pos + element_size > len(bin_data):
                break
            x, y, z = struct.unpack_from(fmt, bin_data, pos)
            vertices.append((float(x), float(y), float(z)))

        return vertices

    def _read_scalar_accessor(self, accessor, buffer_views, bin_data):
        """Lees SCALAR accessor (face indices).

        Args:
            accessor: glTF accessor dict
            buffer_views: Lijst van bufferView dicts
            bin_data: Binary buffer bytes

        Returns:
            Lijst van integer indices
        """
        if accessor.get("type") != "SCALAR":
            return []

        comp_type = accessor.get("componentType", COMPONENT_USHORT)
        count = accessor.get("count", 0)
        bv_idx = accessor.get("bufferView")

        if bv_idx is None or bin_data is None or count == 0:
            return []

        bv = buffer_views[bv_idx]
        bv_offset = bv.get("byteOffset", 0)
        acc_offset = accessor.get("byteOffset", 0)

        start = bv_offset + acc_offset
        comp_size = _COMPONENT_SIZE.get(comp_type, 2)
        comp_fmt = _COMPONENT_FORMAT.get(comp_type, "H")

        indices = []
        fmt = "<{0}".format(comp_fmt)

        for i in range(count):
            pos = start + i * comp_size
            if pos + comp_size > len(bin_data):
                break
            val = struct.unpack_from(fmt, bin_data, pos)[0]
            indices.append(int(val))

        return indices

    def _indices_to_triangles(self, indices, mode, offset):
        """Converteer raw indices naar triangle face tuples.

        Args:
            indices: Lijst van integer indices
            mode: glTF primitive mode (4=TRIANGLES, 5=TRIANGLE_STRIP, 6=TRIANGLE_FAN)
            offset: Vertex offset voor samengevoegde meshes

        Returns:
            Lijst van (i0, i1, i2) tuples (0-based, met offset)
        """
        faces = []

        if mode == 4:  # TRIANGLES
            for i in range(0, len(indices) - 2, 3):
                faces.append((
                    offset + indices[i],
                    offset + indices[i + 1],
                    offset + indices[i + 2],
                ))

        elif mode == 5:  # TRIANGLE_STRIP
            for i in range(len(indices) - 2):
                if i % 2 == 0:
                    faces.append((
                        offset + indices[i],
                        offset + indices[i + 1],
                        offset + indices[i + 2],
                    ))
                else:
                    faces.append((
                        offset + indices[i],
                        offset + indices[i + 2],
                        offset + indices[i + 1],
                    ))

        elif mode == 6:  # TRIANGLE_FAN
            for i in range(1, len(indices) - 1):
                faces.append((
                    offset + indices[0],
                    offset + indices[i],
                    offset + indices[i + 1],
                ))

        return faces
