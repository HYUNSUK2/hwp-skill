#!/usr/bin/env python3
"""HWP binary (v5) file reader - olefile based"""
import os
import sys
import json
import struct
import zlib


class HwpReader:
    """HWP v5 binary file reader (olefile based)"""

    def __init__(self, filepath):
        self.filepath = filepath
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")

    def read_text(self):
        """Extract text from HWP file"""
        try:
            import olefile
        except ImportError:
            return {
                "error": "olefile package is required. Install: pip install olefile",
            }

        try:
            ole = olefile.OleFileIO(self.filepath)
            texts = []

            # Extract text from BodyText streams
            for stream_name in ole.listdir():
                joined = "/".join(stream_name)
                if joined.startswith("BodyText/") or joined.startswith("BodyText\\"):
                    try:
                        data = ole.openstream(stream_name).read()
                        text = self._extract_text_from_bodytext(data)
                        if text.strip():
                            texts.append(text)
                    except Exception:
                        pass

            ole.close()

            if not texts:
                return {
                    "text": "",
                    "message": "Could not extract text. The file may be encrypted or corrupted.",
                }

            return {"text": "\n\n".join(texts), "sections": len(texts)}
        except Exception as e:
            return {"error": str(e)}

    def _extract_text_from_bodytext(self, data):
        """Extract text from BodyText stream binary"""
        # HWP v5 BodyText may be compressed
        try:
            decompressed = zlib.decompress(data, -15)
        except Exception:
            decompressed = data

        texts = []
        # Parse HWP record structure
        offset = 0
        while offset < len(decompressed):
            if offset + 4 > len(decompressed):
                break

            # Record header (4 bytes)
            header = struct.unpack_from("<I", decompressed, offset)[0]
            tag_id = header & 0x3FF
            level = (header >> 10) & 0x3FF
            size = (header >> 20) & 0xFFF

            offset += 4

            if size == 0xFFF:
                # Extended size (additional 4 bytes)
                if offset + 4 > len(decompressed):
                    break
                size = struct.unpack_from("<I", decompressed, offset)[0]
                offset += 4

            if offset + size > len(decompressed):
                break

            record_data = decompressed[offset: offset + size]
            offset += size

            # HWPTAG_PARA_TEXT = 67 (paragraph text)
            if tag_id == 67:
                text = self._parse_para_text(record_data)
                if text.strip():
                    texts.append(text)

        return "\n".join(texts)

    def _parse_para_text(self, data):
        """Parse paragraph text record"""
        chars = []
        i = 0
        while i < len(data) - 1:
            code = struct.unpack_from("<H", data, i)[0]
            i += 2

            if code == 0:
                break
            elif code < 32:
                # Control characters
                if code == 13:  # Line break
                    chars.append("\n")
                elif code == 10:  # Tab
                    chars.append("\t")
                elif code in (1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23):
                    # Inline/extended control characters - skip additional bytes
                    pass
            else:
                try:
                    chars.append(chr(code))
                except (ValueError, OverflowError):
                    pass

        return "".join(chars)

    def read_metadata(self):
        """Read HWP file metadata"""
        try:
            import olefile
        except ImportError:
            return {"error": "olefile package is required. Install: pip install olefile"}

        try:
            ole = olefile.OleFileIO(self.filepath)
            meta = {"filename": os.path.basename(self.filepath)}

            # OLE metadata
            ole_meta = ole.get_metadata()
            if ole_meta:
                for attr in ["title", "subject", "author", "keywords", "comments",
                              "last_saved_by", "creating_application"]:
                    val = getattr(ole_meta, attr, None)
                    if val:
                        if isinstance(val, bytes):
                            try:
                                val = val.decode("utf-8")
                            except UnicodeDecodeError:
                                try:
                                    val = val.decode("euc-kr")
                                except UnicodeDecodeError:
                                    val = val.decode("latin-1")
                        meta[attr] = val

                for attr in ["create_time", "last_saved_time"]:
                    val = getattr(ole_meta, attr, None)
                    if val:
                        meta[attr] = str(val)

            # HWP version info from FileHeader stream
            try:
                header_data = ole.openstream("FileHeader").read()
                if len(header_data) >= 36:
                    signature = header_data[:32]
                    version = struct.unpack_from("<I", header_data, 32)[0]
                    major = (version >> 24) & 0xFF
                    minor = (version >> 16) & 0xFF
                    build = (version >> 8) & 0xFF
                    revision = version & 0xFF
                    meta["hwp_version"] = f"{major}.{minor}.{build}.{revision}"
                    meta["signature"] = signature.decode("ascii", errors="ignore").strip("\x00")
            except Exception:
                pass

            # Stream list
            meta["streams"] = ["/".join(s) for s in ole.listdir()]

            ole.close()
            return meta
        except Exception as e:
            return {"error": str(e)}

    def read_tables(self):
        """Extract tables from HWP file (limited support)"""
        try:
            import olefile
        except ImportError:
            return {"error": "olefile package is required. Install: pip install olefile"}

        try:
            ole = olefile.OleFileIO(self.filepath)
            tables = []
            current_table = []
            in_table = False

            for stream_name in ole.listdir():
                joined = "/".join(stream_name)
                if not joined.startswith("BodyText/"):
                    continue

                data = ole.openstream(stream_name).read()
                try:
                    decompressed = zlib.decompress(data, -15)
                except Exception:
                    decompressed = data

                # Parse records to extract table structure
                offset = 0
                while offset < len(decompressed):
                    if offset + 4 > len(decompressed):
                        break

                    header = struct.unpack_from("<I", decompressed, offset)[0]
                    tag_id = header & 0x3FF
                    size = (header >> 20) & 0xFFF
                    offset += 4

                    if size == 0xFFF:
                        if offset + 4 > len(decompressed):
                            break
                        size = struct.unpack_from("<I", decompressed, offset)[0]
                        offset += 4

                    if offset + size > len(decompressed):
                        break

                    record_data = decompressed[offset: offset + size]
                    offset += size

                    # HWPTAG_TABLE = 78 (table start)
                    if tag_id == 78:
                        if current_table:
                            tables.append(current_table)
                        current_table = []
                        in_table = True

                    # HWPTAG_PARA_TEXT = 67 (paragraph text inside table cell)
                    elif tag_id == 67 and in_table:
                        text = self._parse_para_text(record_data)
                        if text.strip():
                            current_table.append(text.strip())

            if current_table:
                tables.append(current_table)

            ole.close()

            result_tables = []
            for idx, tbl in enumerate(tables):
                result_tables.append({
                    "index": idx,
                    "cells": tbl,
                    "note": "Row/column structure cannot be fully reconstructed from HWP binary. Cell contents are listed in order.",
                })

            return {
                "tables": result_tables,
                "count": len(result_tables),
                "note": "HWP binary table extraction is limited. Use HWPX format for accurate row/column structure.",
            }
        except Exception as e:
            return {"error": str(e)}


def main():
    import argparse

    parser = argparse.ArgumentParser(description="HWP binary file reader")
    parser.add_argument("command", choices=["text", "tables", "meta"],
                        help="Command: text, tables, meta")
    parser.add_argument("filepath", help="HWP file path")
    args = parser.parse_args()

    reader = HwpReader(args.filepath)

    if args.command == "text":
        result = reader.read_text()
    elif args.command == "tables":
        result = reader.read_tables()
    elif args.command == "meta":
        result = reader.read_metadata()

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
