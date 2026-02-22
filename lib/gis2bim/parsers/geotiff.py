# -*- coding: utf-8 -*-
"""
GeoTIFF Parser
==============

Pure IronPython/.NET parser voor ongecomprimeerde single-band float32 GeoTIFF.
Specifiek voor PDOK WCS AHN data (DTM/DSM).

Gebruikt System.IO.BinaryReader voor bestandstoegang.

Voorbeeld:
    from gis2bim.parsers.geotiff import GeoTiffReader

    reader = GeoTiffReader()
    grid = reader.read("ahn_dtm.tif")
    points = reader.to_xyz_points(grid)
"""

import struct
import os
import zlib


class GeoTiffError(Exception):
    """Fout bij het parsen van een GeoTIFF bestand."""
    pass


class GeoTiffReader(object):
    """Parse ongecomprimeerde single-band float32 GeoTIFF bestanden.

    Ondersteunt enkel ongecomprimeerde TIFF met float32 pixel data,
    zoals PDOK WCS standaard retourneert.
    """

    # TIFF tag IDs
    TAG_IMAGE_WIDTH = 256
    TAG_IMAGE_LENGTH = 257
    TAG_BITS_PER_SAMPLE = 258
    TAG_COMPRESSION = 259
    TAG_PHOTOMETRIC = 262
    TAG_STRIP_OFFSETS = 273
    TAG_SAMPLES_PER_PIXEL = 277
    TAG_ROWS_PER_STRIP = 278
    TAG_STRIP_BYTE_COUNTS = 279
    TAG_PREDICTOR = 317
    TAG_SAMPLE_FORMAT = 339

    # GeoTIFF tag IDs
    TAG_MODEL_PIXEL_SCALE = 33550
    TAG_MODEL_TIEPOINT = 33922
    TAG_GEO_KEY_DIRECTORY = 34735
    TAG_GDAL_NODATA = 42113

    # TIFF type sizes (in bytes)
    TYPE_SIZES = {
        1: 1,   # BYTE
        2: 1,   # ASCII
        3: 2,   # SHORT
        4: 4,   # LONG
        5: 8,   # RATIONAL
        6: 1,   # SBYTE
        7: 1,   # UNDEFINED
        8: 2,   # SSHORT
        9: 4,   # SLONG
        10: 8,  # SRATIONAL
        11: 4,  # FLOAT
        12: 8,  # DOUBLE
    }

    def read(self, filepath):
        """Parse GeoTIFF en return elevation grid.

        Args:
            filepath: Pad naar het GeoTIFF bestand

        Returns:
            dict met:
            - "width", "height": afmetingen in pixels
            - "origin_x", "origin_y": RD coordinaat linksboven
            - "pixel_size_x", "pixel_size_y": pixel grootte in meters
            - "data": lijst van float waarden (row-major, top-to-bottom)
            - "nodata": nodata waarde

        Raises:
            GeoTiffError: Bij ongeldige of niet-ondersteunde TIFF
        """
        if not os.path.exists(filepath):
            raise GeoTiffError("Bestand niet gevonden: {0}".format(filepath))

        with open(filepath, "rb") as f:
            raw = f.read()

        # 1. Lees header: byte order + magic + IFD offset
        if len(raw) < 8:
            raise GeoTiffError("Bestand te klein voor TIFF header")

        byte_order = raw[0:2]
        if byte_order == b"II":
            endian = "<"  # little-endian
        elif byte_order == b"MM":
            endian = ">"  # big-endian
        else:
            raise GeoTiffError("Ongeldig TIFF byte order: {0}".format(byte_order))

        magic = struct.unpack_from(endian + "H", raw, 2)[0]
        if magic != 42:
            raise GeoTiffError("Ongeldig TIFF magic number: {0}".format(magic))

        ifd_offset = struct.unpack_from(endian + "I", raw, 4)[0]

        # 2. Lees IFD entries
        tags = self._read_ifd(raw, ifd_offset, endian)

        # 3. Valideer TIFF structuur
        compression = self._get_tag_value(tags, self.TAG_COMPRESSION, default=1)
        if compression not in (1, 8):
            raise GeoTiffError(
                "GeoTIFF compression type {0} niet ondersteund. "
                "Alleen ongecomprimeerd (1) en Deflate (8) worden ondersteund.".format(
                    compression)
            )

        predictor = self._get_tag_value(tags, self.TAG_PREDICTOR, default=1)

        bits_per_sample = self._get_tag_value(tags, self.TAG_BITS_PER_SAMPLE, default=32)
        sample_format = self._get_tag_value(tags, self.TAG_SAMPLE_FORMAT, default=3)

        # sample_format: 1=uint, 2=int, 3=float
        if bits_per_sample != 32 or sample_format not in (1, 2, 3):
            raise GeoTiffError(
                "Niet-ondersteund pixel formaat: {0} bits, format {1}. "
                "Verwacht: 32-bit float.".format(bits_per_sample, sample_format)
            )

        # 4. Lees afmetingen
        width = self._get_tag_value(tags, self.TAG_IMAGE_WIDTH)
        height = self._get_tag_value(tags, self.TAG_IMAGE_LENGTH)

        if not width or not height:
            raise GeoTiffError("ImageWidth of ImageLength ontbreekt in TIFF")

        # 5. Lees GeoTIFF metadata
        pixel_scale = self._get_tag_doubles(tags, self.TAG_MODEL_PIXEL_SCALE, raw, endian)
        tiepoint = self._get_tag_doubles(tags, self.TAG_MODEL_TIEPOINT, raw, endian)

        if not pixel_scale or len(pixel_scale) < 2:
            raise GeoTiffError("ModelPixelScaleTag ontbreekt of onvolledig")

        if not tiepoint or len(tiepoint) < 6:
            raise GeoTiffError("ModelTiepointTag ontbreekt of onvolledig")

        pixel_size_x = pixel_scale[0]
        pixel_size_y = pixel_scale[1]

        # Tiepoint: (raster_x, raster_y, raster_z, model_x, model_y, model_z)
        origin_x = tiepoint[3] - tiepoint[0] * pixel_size_x
        origin_y = tiepoint[4] + tiepoint[1] * pixel_size_y

        # 6. Lees NODATA waarde
        nodata = self._get_nodata(tags, raw, endian)

        # 7. Lees pixel data
        data = self._read_pixel_data(
            raw, tags, width, height, endian, sample_format,
            compression, predictor
        )

        return {
            "width": width,
            "height": height,
            "origin_x": origin_x,
            "origin_y": origin_y,
            "pixel_size_x": pixel_size_x,
            "pixel_size_y": pixel_size_y,
            "data": data,
            "nodata": nodata,
        }

    def to_xyz_points(self, grid_data, nodata_value=None):
        """Converteer grid naar lijst van (x, y, z) tuples.

        Punten worden berekend vanuit de linksboven origin.
        Y loopt van boven naar beneden in TIFF (rij 0 = noordelijkste punt).

        Args:
            grid_data: dict van read() methode
            nodata_value: Optionele nodata waarde override

        Returns:
            lijst van (rd_x, rd_y, elevation_m) tuples
        """
        width = grid_data["width"]
        height = grid_data["height"]
        origin_x = grid_data["origin_x"]
        origin_y = grid_data["origin_y"]
        pixel_size_x = grid_data["pixel_size_x"]
        pixel_size_y = grid_data["pixel_size_y"]
        data = grid_data["data"]
        nodata = nodata_value if nodata_value is not None else grid_data.get("nodata", -9999)

        points = []
        for row in range(height):
            for col in range(width):
                idx = row * width + col
                z = data[idx]

                # Filter nodata en NaN waarden
                if z != z:  # NaN check (NaN != NaN)
                    continue
                if nodata is not None and z == nodata:
                    continue
                if z > 1000 or z < -100:
                    continue

                # Bereken RD coordinaten
                # TIFF rij 0 = bovenste rij (noordelijkste)
                rd_x = origin_x + (col + 0.5) * pixel_size_x
                rd_y = origin_y - (row + 0.5) * pixel_size_y

                points.append((rd_x, rd_y, z))

        return points

    def _read_ifd(self, raw, offset, endian):
        """Lees alle IFD (Image File Directory) entries.

        Returns:
            dict van tag_id -> {type, count, value_offset, value}
        """
        tags = {}
        num_entries = struct.unpack_from(endian + "H", raw, offset)[0]
        entry_offset = offset + 2

        for i in range(num_entries):
            pos = entry_offset + i * 12
            tag_id = struct.unpack_from(endian + "H", raw, pos)[0]
            type_id = struct.unpack_from(endian + "H", raw, pos + 2)[0]
            count = struct.unpack_from(endian + "I", raw, pos + 4)[0]

            type_size = self.TYPE_SIZES.get(type_id, 1)
            total_size = type_size * count

            if total_size <= 4:
                # Waarde past in de 4 bytes van het offset veld
                value_data_offset = pos + 8
            else:
                # Waarde staat elders in het bestand
                value_data_offset = struct.unpack_from(endian + "I", raw, pos + 8)[0]

            # Lees enkele waarde voor veelgebruikte tags
            value = None
            if count == 1:
                if type_id == 3:  # SHORT
                    value = struct.unpack_from(endian + "H", raw, value_data_offset)[0]
                elif type_id == 4:  # LONG
                    value = struct.unpack_from(endian + "I", raw, value_data_offset)[0]
                elif type_id == 11:  # FLOAT
                    value = struct.unpack_from(endian + "f", raw, value_data_offset)[0]
                elif type_id == 12:  # DOUBLE
                    value = struct.unpack_from(endian + "d", raw, value_data_offset)[0]

            tags[tag_id] = {
                "type": type_id,
                "count": count,
                "value_data_offset": value_data_offset,
                "value": value,
            }

        return tags

    def _get_tag_value(self, tags, tag_id, default=None):
        """Haal enkele waarde op uit een tag."""
        tag = tags.get(tag_id)
        if tag and tag["value"] is not None:
            return tag["value"]
        return default

    def _get_tag_doubles(self, tags, tag_id, raw, endian):
        """Lees een array van DOUBLE waarden uit een tag."""
        tag = tags.get(tag_id)
        if not tag:
            return None

        count = tag["count"]
        offset = tag["value_data_offset"]
        values = []

        for i in range(count):
            val = struct.unpack_from(endian + "d", raw, offset + i * 8)[0]
            values.append(val)

        return values

    def _get_nodata(self, tags, raw, endian):
        """Lees NODATA waarde (GDAL_NODATA tag of standaard -9999)."""
        tag = tags.get(self.TAG_GDAL_NODATA)
        if tag:
            # GDAL_NODATA is een ASCII string
            offset = tag["value_data_offset"]
            count = tag["count"]
            try:
                nodata_str = raw[offset:offset + count].rstrip(b"\x00").decode("ascii").strip()
                return float(nodata_str)
            except (ValueError, UnicodeDecodeError):
                pass

        return -9999.0

    def _read_pixel_data(self, raw, tags, width, height, endian, sample_format,
                         compression=1, predictor=1):
        """Lees pixel data uit strips.

        Ondersteunt ongecomprimeerd (compression=1) en Deflate (compression=8)
        met optioneel floating point predictor (predictor=3).

        Returns:
            Lijst van float waarden (row-major, top-to-bottom)
        """
        strip_offsets_tag = tags.get(self.TAG_STRIP_OFFSETS)
        strip_counts_tag = tags.get(self.TAG_STRIP_BYTE_COUNTS)
        rows_per_strip = self._get_tag_value(tags, self.TAG_ROWS_PER_STRIP, default=height)

        if not strip_offsets_tag:
            raise GeoTiffError("StripOffsets tag ontbreekt")

        # Lees strip offsets
        strip_offsets = self._read_long_array(
            raw, strip_offsets_tag, endian
        )

        # Lees strip byte counts
        if strip_counts_tag:
            strip_byte_counts = self._read_long_array(
                raw, strip_counts_tag, endian
            )
        else:
            bytes_per_row = width * 4  # float32 = 4 bytes
            strip_byte_counts = [rows_per_strip * bytes_per_row] * len(strip_offsets)

        # Kies het juiste struct format
        if sample_format == 3:  # float
            fmt_char = "f"
        elif sample_format == 2:  # signed int
            fmt_char = "i"
        elif sample_format == 1:  # unsigned int
            fmt_char = "I"
        else:
            fmt_char = "f"

        bytes_per_sample = 4  # float32/int32
        total_pixels = width * height
        all_bytes = bytearray()

        for strip_idx in range(len(strip_offsets)):
            offset = strip_offsets[strip_idx]
            byte_count = strip_byte_counts[strip_idx]
            strip_raw = raw[offset:offset + byte_count]

            # Stap 1: Decompressie
            if compression == 8:
                try:
                    strip_data = zlib.decompress(strip_raw)
                except zlib.error as e:
                    raise GeoTiffError(
                        "Deflate decompressie mislukt voor strip {0}: {1}".format(
                            strip_idx, e)
                    )
            else:
                strip_data = strip_raw

            # Stap 2: Reverse predictor
            if predictor == 3:
                # Floating point predictor: byte-shuffle + horizontal differencing
                strip_rows = min(rows_per_strip, height - strip_idx * rows_per_strip)
                strip_data = self._reverse_predictor3(
                    strip_data, width, strip_rows, bytes_per_sample, endian
                )
            elif predictor == 2:
                # Horizontale differencing predictor (voor integers)
                strip_rows = min(rows_per_strip, height - strip_idx * rows_per_strip)
                strip_data = self._reverse_predictor2(
                    strip_data, width, strip_rows, bytes_per_sample
                )

            all_bytes.extend(strip_data)

        # Parse float/int waarden
        data = []
        for i in range(total_pixels):
            pos = i * bytes_per_sample
            if pos + bytes_per_sample > len(all_bytes):
                break
            val = struct.unpack_from(endian + fmt_char, bytes(all_bytes), pos)[0]
            if fmt_char in ("i", "I"):
                val = float(val)
            data.append(val)

        if len(data) < total_pixels:
            raise GeoTiffError(
                "Onvoldoende pixel data: verwacht {0}, gevonden {1}".format(
                    total_pixels, len(data))
            )

        return data[:total_pixels]

    def _reverse_predictor3(self, data, width, rows, bytes_per_sample, endian="<"):
        """Reverse floating point predictor (TIFF Predictor=3).

        De floating point predictor doet per rij:
        1. Byte-shuffle: groepeer bytes per significantie (MSB eerst)
        2. Horizontale differencing op de byte stream

        Wij keren dit om:
        1. Reverse horizontale differencing (accumuleer)
        2. Un-shuffle bytes (hergroepeer per pixel)

        Byte planes zijn altijd MSB-eerst geordend (TIFF Tech Note 3).
        Bij little-endian data moet de volgorde omgekeerd worden.
        """
        data = bytearray(data)
        row_bytes = width * bytes_per_sample
        result = bytearray(len(data))

        for row in range(rows):
            row_start = row * row_bytes

            # Stap 1: Reverse horizontale differencing op bytes
            for i in range(row_start + 1, row_start + row_bytes):
                data[i] = (data[i] + data[i - 1]) & 0xFF

            # Stap 2: Un-shuffle bytes
            # Planes zijn MSB-eerst: plane 0 = MSB, plane N-1 = LSB
            # Voor LE floats: MSB = byte N-1, LSB = byte 0
            # Voor BE floats: MSB = byte 0, LSB = byte N-1
            for px in range(width):
                for b in range(bytes_per_sample):
                    if endian == "<":
                        # LE: plane b (MSB-first) -> byte (N-1-b) in memory
                        dest_byte = bytes_per_sample - 1 - b
                    else:
                        # BE: plane b -> byte b
                        dest_byte = b
                    result[row_start + px * bytes_per_sample + dest_byte] = \
                        data[row_start + b * width + px]

        return bytes(result)

    def _reverse_predictor2(self, data, width, rows, bytes_per_sample):
        """Reverse horizontale differencing predictor (TIFF Predictor=2)."""
        data = bytearray(data)
        row_bytes = width * bytes_per_sample

        for row in range(rows):
            row_start = row * row_bytes
            for px in range(1, width):
                for b in range(bytes_per_sample):
                    pos = row_start + px * bytes_per_sample + b
                    prev = row_start + (px - 1) * bytes_per_sample + b
                    data[pos] = (data[pos] + data[prev]) & 0xFF

        return bytes(data)

    def _read_long_array(self, raw, tag, endian):
        """Lees een array van LONG of SHORT waarden uit een tag."""
        count = tag["count"]
        offset = tag["value_data_offset"]
        type_id = tag["type"]

        values = []
        if type_id == 3:  # SHORT
            for i in range(count):
                val = struct.unpack_from(endian + "H", raw, offset + i * 2)[0]
                values.append(val)
        elif type_id == 4:  # LONG
            for i in range(count):
                val = struct.unpack_from(endian + "I", raw, offset + i * 4)[0]
                values.append(val)
        elif count == 1 and tag["value"] is not None:
            values.append(int(tag["value"]))
        else:
            # Fallback: probeer als LONG
            for i in range(count):
                val = struct.unpack_from(endian + "I", raw, offset + i * 4)[0]
                values.append(val)

        return values
