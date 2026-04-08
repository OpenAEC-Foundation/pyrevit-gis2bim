# -*- coding: utf-8 -*-
"""
LAS File Parser
===============

Parser voor ongecomprimeerde LAS bestanden (output van laszip.exe).
Ondersteunt LAS 1.2, 1.3 en 1.4 met alle gangbare point formats.

Leest in chunks voor geheugen-efficientie bij grote bestanden.
Bevat ook een XYZ tekst parser voor las2txt output.

Gebruik:
    from gis2bim.parsers.las import LASReader

    reader = LASReader()
    # LAS binary bestand (na decompressie met laszip)
    points = reader.read("output.las", bbox=(155000, 463000, 155100, 463100))

    # XYZ tekst bestand (output van las2txt)
    points = reader.read_xyz_text("output.xyz", thin_grid=0.5)
"""

import struct
import os


class LASError(Exception):
    """Fout bij het parsen van een LAS bestand."""
    pass


class LASReader(object):
    """Parse ongecomprimeerde LAS bestanden en XYZ tekstbestanden.

    Ondersteunt:
    - LAS 1.2, 1.3, 1.4
    - Point formats 0-10 (alleen X, Y, Z en classification worden gelezen)
    - Bbox filtering (spatial crop)
    - Grid-based thinning (eerste punt of hoogste punt per cel)
    - Classification filtering (bijv. class 2 = ground voor DTM)
    """

    # Chunk grootte voor het lezen van punten (50k punten per keer)
    CHUNK_SIZE = 50000

    def read(self, filepath, bbox=None, thin_grid=None, classification=None,
             keep_highest=False):
        """Lees punten uit een ongecomprimeerd LAS bestand.

        Args:
            filepath: Pad naar .las bestand
            bbox: Optioneel (xmin, ymin, xmax, ymax) voor spatial filter
            thin_grid: Optioneel grid celgrootte in meters voor thinning
            classification: Optioneel set/lijst van classificatiecodes
                           (bijv. [2] voor ground, [2, 6] voor ground+building)
            keep_highest: Bij thinning: bewaar hoogste punt per cel (True)
                         of eerste punt (False). True is nuttig voor DSM.

        Returns:
            Lijst van (x, y, z) tuples

        Raises:
            LASError: Bij ongeldige bestanden
        """
        if not os.path.exists(filepath):
            raise LASError("LAS bestand niet gevonden: {0}".format(filepath))

        with open(filepath, "rb") as f:
            # Lees header (max 375 bytes voor LAS 1.4)
            header = f.read(375)
            if len(header) < 227:
                raise LASError("Bestand te klein voor LAS header")

            # Valideer signature
            sig = header[0:4]
            if sig != b"LASF":
                raise LASError("Ongeldig LAS bestand (signature: {0})".format(
                    repr(sig)))

            # Versie
            version_minor = struct.unpack_from("B", header, 25)[0]

            # Point data info
            point_data_offset = struct.unpack_from("<I", header, 96)[0]
            point_format = struct.unpack_from("B", header, 104)[0]
            point_record_length = struct.unpack_from("<H", header, 105)[0]

            # Aantal punten (64-bit voor LAS 1.4)
            if version_minor >= 4 and len(header) >= 255:
                num_points = struct.unpack_from("<Q", header, 247)[0]
            else:
                num_points = struct.unpack_from("<I", header, 107)[0]

            # Scale en offset (altijd op dezelfde positie)
            x_scale = struct.unpack_from("<d", header, 131)[0]
            y_scale = struct.unpack_from("<d", header, 139)[0]
            z_scale = struct.unpack_from("<d", header, 147)[0]
            x_offset = struct.unpack_from("<d", header, 155)[0]
            y_offset = struct.unpack_from("<d", header, 163)[0]
            z_offset = struct.unpack_from("<d", header, 171)[0]

            # Classification byte positie verschilt per point format
            # Format 0-5: byte 15, Format 6-10: byte 16
            class_byte_offset = 16 if point_format >= 6 else 15

            # Converteer classification naar set voor snelle lookup
            class_filter = set(classification) if classification else None

            # Grid dict voor thinning, of directe lijst
            thin_cells = {} if thin_grid is not None else None
            points = [] if thin_grid is None else None

            # Seek naar punt data
            f.seek(point_data_offset)

            # Lees in chunks
            read_count = 0

            while read_count < num_points:
                remaining = num_points - read_count
                to_read = min(self.CHUNK_SIZE, remaining)
                data = f.read(to_read * point_record_length)

                if not data:
                    break

                actual = len(data) // point_record_length

                for i in range(actual):
                    off = i * point_record_length

                    # X, Y, Z als int32
                    xi = struct.unpack_from("<i", data, off)[0]
                    yi = struct.unpack_from("<i", data, off + 4)[0]
                    zi = struct.unpack_from("<i", data, off + 8)[0]

                    # Converteer naar werkelijke coordinaten
                    x = xi * x_scale + x_offset
                    y = yi * y_scale + y_offset
                    z = zi * z_scale + z_offset

                    # Classification filter
                    if class_filter:
                        cls_off = off + class_byte_offset
                        if cls_off < len(data):
                            cls = struct.unpack_from("B", data, cls_off)[0]
                            if cls not in class_filter:
                                continue

                    # Bbox filter
                    if bbox:
                        if (x < bbox[0] or x > bbox[2] or
                                y < bbox[1] or y > bbox[3]):
                            continue

                    # Nodata / onzin filter
                    if z > 1000 or z < -100:
                        continue

                    # Thinning of directe opslag
                    if thin_cells is not None:
                        gx = int(x // thin_grid)
                        gy = int(y // thin_grid)
                        key = (gx, gy)

                        if key in thin_cells:
                            if keep_highest and z > thin_cells[key][2]:
                                thin_cells[key] = (x, y, z)
                            continue
                        thin_cells[key] = (x, y, z)
                    else:
                        points.append((x, y, z))

                read_count += actual

            # Resultaat
            if thin_cells is not None:
                return list(thin_cells.values())
            return points

    def read_xyz_text(self, filepath, bbox=None, thin_grid=None,
                      keep_highest=False):
        """Lees XYZ tekstbestand (output van las2txt of PTS).

        Ondersteunt spatie, komma en tab gescheiden formaten.
        Eerste regel met minder dan 3 kolommen wordt overgeslagen
        (bijv. PTS puntentelling header).

        Args:
            filepath: Pad naar .xyz / .txt / .pts bestand
            bbox: Optioneel (xmin, ymin, xmax, ymax) voor spatial filter
            thin_grid: Optioneel grid celgrootte voor thinning
            keep_highest: Bij thinning: bewaar hoogste punt per cel

        Returns:
            Lijst van (x, y, z) tuples
        """
        if not os.path.exists(filepath):
            raise LASError("Bestand niet gevonden: {0}".format(filepath))

        thin_cells = {} if thin_grid else None
        points = []
        sep = None

        with open(filepath, "r") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or line.startswith("/"):
                    continue

                # Auto-detect separator bij eerste data regel
                if sep is None:
                    if "," in line:
                        sep = ","
                    elif "\t" in line:
                        sep = "\t"
                    else:
                        sep = None  # whitespace split

                if sep:
                    parts = line.split(sep)
                else:
                    parts = line.split()

                if len(parts) < 3:
                    # PTS header (enkele getal = puntentelling) of ongeldig
                    continue

                try:
                    x = float(parts[0])
                    y = float(parts[1])
                    z = float(parts[2])
                except ValueError:
                    continue

                # Bbox filter
                if bbox:
                    if (x < bbox[0] or x > bbox[2] or
                            y < bbox[1] or y > bbox[3]):
                        continue

                # Nodata filter
                if z > 1000 or z < -100:
                    continue

                # Thinning
                if thin_grid is not None:
                    gx = int(x // thin_grid)
                    gy = int(y // thin_grid)
                    key = (gx, gy)

                    if key in thin_cells:
                        if keep_highest and z > thin_cells[key][2]:
                            thin_cells[key] = (x, y, z)
                        continue
                    thin_cells[key] = (x, y, z)
                else:
                    points.append((x, y, z))

        if thin_grid is not None:
            return list(thin_cells.values())

        return points
