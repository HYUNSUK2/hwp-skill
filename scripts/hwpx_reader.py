#!/usr/bin/env python3
"""HWPX file reader - ZIP + XML parsing"""
import os
import sys
import json
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict


# HWPX XML namespaces
NAMESPACES = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "ht": "http://www.hancom.co.kr/hwpml/2011/table",
    "ho": "http://www.hancom.co.kr/hwpml/2011/objectType",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "urn:oasis:names:tc:opendocument:xmlns:container",
    "odf": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
}

# Register namespaces
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class HwpxReader:
    def __init__(self, filepath):
        self.filepath = filepath
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")

    def _open_zip(self):
        return zipfile.ZipFile(self.filepath, "r")

    def _find_content_files(self, zf):
        """Find content XML files (section0.xml, section1.xml, ...)"""
        content_files = []
        for name in sorted(zf.namelist()):
            lower = name.lower()
            if "contents/" in lower and lower.endswith(".xml") and "section" in lower:
                content_files.append(name)
        # Fallback to content.xml if no section files found
        if not content_files:
            for name in zf.namelist():
                lower = name.lower()
                if lower.endswith("content.xml") or lower.endswith("contents.xml"):
                    content_files.append(name)
        return content_files

    def _extract_text_from_element(self, elem):
        """Recursively extract text from XML element"""
        texts = []
        # Direct text
        if elem.text and elem.text.strip():
            texts.append(elem.text.strip())

        for child in elem:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            # Text element - extract text directly
            if tag in ("t", "T"):
                if child.text:
                    texts.append(child.text)
            # Ignored elements
            elif tag in ("linesegarray", "lineBreak", "LineSeg"):
                pass
            # Paragraph - recurse then add newline
            elif tag in ("p", "P"):
                sub = self._extract_text_from_element(child)
                if sub:
                    texts.append(sub)
                texts.append("\n")
            # Line break
            elif tag == "br":
                texts.append("\n")
            # Other elements - recurse
            else:
                sub = self._extract_text_from_element(child)
                if sub:
                    texts.append(sub)

            if child.tail and child.tail.strip():
                texts.append(child.tail.strip())

        return "".join(texts)

    def read_text(self):
        """Extract all text from the document"""
        try:
            with self._open_zip() as zf:
                content_files = self._find_content_files(zf)
                if not content_files:
                    return {"error": "Content files not found", "files": zf.namelist()}

                all_text = []
                for cf in content_files:
                    xml_data = zf.read(cf)
                    root = ET.fromstring(xml_data)
                    section_text = self._extract_text_from_element(root)
                    if section_text.strip():
                        all_text.append(section_text)

                return {
                    "text": "\n\n".join(all_text),
                    "sections": len(content_files),
                    "content_files": content_files,
                }
        except Exception as e:
            return {"error": str(e)}

    def read_tables(self):
        """Extract table data from the document"""
        try:
            with self._open_zip() as zf:
                content_files = self._find_content_files(zf)
                if not content_files:
                    return {"error": "Content files not found"}

                tables = []
                table_idx = 0

                for cf in content_files:
                    xml_data = zf.read(cf)
                    root = ET.fromstring(xml_data)

                    # Find all table elements
                    for tbl in self._find_tables(root):
                        table_data = self._parse_table(tbl)
                        if table_data:
                            tables.append({
                                "index": table_idx,
                                "source_file": cf,
                                "rows": len(table_data),
                                "cols": max(len(row) for row in table_data) if table_data else 0,
                                "data": table_data,
                            })
                            table_idx += 1

                return {"tables": tables, "count": len(tables)}
        except Exception as e:
            return {"error": str(e)}

    def _find_tables(self, root):
        """Recursively find table elements in XML tree"""
        tables = []
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag.lower() in ("tbl", "table"):
                tables.append(elem)
        return tables

    def _parse_table(self, tbl_elem):
        """Parse table element into 2D array"""
        rows = []
        for child in tbl_elem.iter():
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag.lower() in ("tr", "row"):
                row_data = []
                for cell_elem in child:
                    cell_tag = cell_elem.tag.split("}")[-1] if "}" in cell_elem.tag else cell_elem.tag
                    if cell_tag.lower() in ("tc", "cell", "td"):
                        cell_text = self._extract_text_from_element(cell_elem)
                        row_data.append(cell_text.strip())
                if row_data:
                    rows.append(row_data)
        return rows

    def read_metadata(self):
        """Read document metadata"""
        try:
            with self._open_zip() as zf:
                meta = {"filename": os.path.basename(self.filepath)}

                # Check META-INF/manifest.xml
                for name in zf.namelist():
                    lower = name.lower()
                    if "meta" in lower and lower.endswith(".xml"):
                        try:
                            xml_data = zf.read(name)
                            root = ET.fromstring(xml_data)
                            meta[name] = self._extract_metadata_fields(root)
                        except Exception:
                            pass

                # Check header.xml
                for name in zf.namelist():
                    lower = name.lower()
                    if "header" in lower and lower.endswith(".xml"):
                        try:
                            xml_data = zf.read(name)
                            root = ET.fromstring(xml_data)
                            meta["header_info"] = self._extract_metadata_fields(root)
                        except Exception:
                            pass

                meta["all_files"] = zf.namelist()
                return meta
        except Exception as e:
            return {"error": str(e)}

    def _extract_metadata_fields(self, root):
        """Extract fields from metadata XML"""
        fields = {}
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if elem.text and elem.text.strip():
                key = tag.lower()
                fields[key] = elem.text.strip()
        return fields

    def read_structure(self):
        """Return document structure as a tree"""
        try:
            with self._open_zip() as zf:
                structure = {"files": {}}
                for name in sorted(zf.namelist()):
                    info = zf.getinfo(name)
                    structure["files"][name] = {
                        "size": info.file_size,
                        "compressed": info.compress_size,
                    }
                    # Add root tag info for XML files
                    if name.endswith(".xml"):
                        try:
                            xml_data = zf.read(name)
                            root = ET.fromstring(xml_data)
                            tag = root.tag.split("}")[-1] if "}" in root.tag else root.tag
                            structure["files"][name]["root_tag"] = tag
                            structure["files"][name]["children"] = len(list(root))
                        except Exception:
                            pass
                return structure
        except Exception as e:
            return {"error": str(e)}


def main():
    import argparse

    parser = argparse.ArgumentParser(description="HWPX file reader")
    parser.add_argument("command", choices=["text", "tables", "meta", "structure"],
                        help="Command: text, tables, meta, structure")
    parser.add_argument("filepath", help="HWPX file path")
    args = parser.parse_args()

    reader = HwpxReader(args.filepath)

    if args.command == "text":
        result = reader.read_text()
    elif args.command == "tables":
        result = reader.read_tables()
    elif args.command == "meta":
        result = reader.read_metadata()
    elif args.command == "structure":
        result = reader.read_structure()

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
