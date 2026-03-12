#!/usr/bin/env python3
"""HWPX file editor - unzip, modify XML, repack"""
import os
import sys
import json
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET


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
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class HwpxEditor:
    def __init__(self, filepath):
        self.filepath = filepath
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")

    def _find_content_files(self, zf):
        """Find content XML files"""
        content_files = []
        for name in sorted(zf.namelist()):
            lower = name.lower()
            if "contents/" in lower and lower.endswith(".xml") and "section" in lower:
                content_files.append(name)
        if not content_files:
            for name in zf.namelist():
                lower = name.lower()
                if lower.endswith("content.xml") or lower.endswith("contents.xml"):
                    content_files.append(name)
        return content_files

    def _repack_zip(self, temp_dir, output_path):
        """Repack temp directory back into a ZIP file"""
        # Preserve original file order
        with zipfile.ZipFile(self.filepath, "r") as orig_zf:
            original_names = orig_zf.namelist()

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in original_names:
                file_path = os.path.join(temp_dir, name)
                if os.path.exists(file_path):
                    zf.write(file_path, name)

            # Add any new files not in original
            for root, dirs, files in os.walk(temp_dir):
                for f in files:
                    full_path = os.path.join(root, f)
                    arcname = os.path.relpath(full_path, temp_dir)
                    if arcname not in original_names:
                        zf.write(full_path, arcname)

    def replace_text(self, find_text, replace_text, output_path=None):
        """Replace text in HWPX file"""
        if output_path is None:
            output_path = self.filepath

        try:
            temp_dir = tempfile.mkdtemp()
            replace_count = 0

            # Extract ZIP
            with zipfile.ZipFile(self.filepath, "r") as zf:
                zf.extractall(temp_dir)
                content_files = self._find_content_files(zf)

            # Replace text in each content file
            for cf in content_files:
                cf_path = os.path.join(temp_dir, cf)
                if not os.path.exists(cf_path):
                    continue

                tree = ET.parse(cf_path)
                root = tree.getroot()
                count = self._replace_text_in_element(root, find_text, replace_text)
                replace_count += count

                if count > 0:
                    tree.write(cf_path, encoding="utf-8", xml_declaration=True)

            if replace_count == 0:
                shutil.rmtree(temp_dir)
                return {
                    "success": False,
                    "message": f"'{find_text}' not found",
                    "replacements": 0,
                }

            # Repack
            self._repack_zip(temp_dir, output_path)
            shutil.rmtree(temp_dir)

            return {
                "success": True,
                "message": f"Replaced '{find_text}' with '{replace_text}'",
                "replacements": replace_count,
                "output": output_path,
            }
        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return {"success": False, "error": str(e)}

    def _replace_text_in_element(self, elem, find_text, replace_text):
        """Recursively replace text in XML element"""
        count = 0

        if elem.text and find_text in elem.text:
            elem.text = elem.text.replace(find_text, replace_text)
            count += 1

        if elem.tail and find_text in elem.tail:
            elem.tail = elem.tail.replace(find_text, replace_text)
            count += 1

        for child in elem:
            count += self._replace_text_in_element(child, find_text, replace_text)

        return count

    def edit_table_cell(self, table_idx, row, col, value, output_path=None):
        """Edit a table cell in HWPX file"""
        if output_path is None:
            output_path = self.filepath

        try:
            temp_dir = tempfile.mkdtemp()

            with zipfile.ZipFile(self.filepath, "r") as zf:
                zf.extractall(temp_dir)
                content_files = self._find_content_files(zf)

            # Find target table
            current_table_idx = 0
            modified = False

            for cf in content_files:
                cf_path = os.path.join(temp_dir, cf)
                if not os.path.exists(cf_path):
                    continue

                tree = ET.parse(cf_path)
                root = tree.getroot()
                tables = self._find_tables(root)

                for tbl in tables:
                    if current_table_idx == table_idx:
                        success = self._set_cell_value(tbl, row, col, value)
                        if success:
                            tree.write(cf_path, encoding="utf-8", xml_declaration=True)
                            modified = True
                        else:
                            shutil.rmtree(temp_dir)
                            return {
                                "success": False,
                                "error": f"Cell ({row}, {col}) not found",
                            }
                        break
                    current_table_idx += 1

                if modified:
                    break

            if not modified:
                shutil.rmtree(temp_dir)
                return {
                    "success": False,
                    "error": f"Table index {table_idx} not found (total: {current_table_idx})",
                }

            self._repack_zip(temp_dir, output_path)
            shutil.rmtree(temp_dir)

            return {
                "success": True,
                "message": f"Table[{table_idx}] cell({row},{col}) = '{value}' updated",
                "output": output_path,
            }
        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return {"success": False, "error": str(e)}

    def fill_table(self, table_idx, data, start_row=0, start_col=0, output_path=None):
        """Bulk fill table data in HWPX file"""
        if output_path is None:
            output_path = self.filepath

        try:
            temp_dir = tempfile.mkdtemp()

            with zipfile.ZipFile(self.filepath, "r") as zf:
                zf.extractall(temp_dir)
                content_files = self._find_content_files(zf)

            current_table_idx = 0
            modified = False
            cells_filled = 0

            for cf in content_files:
                cf_path = os.path.join(temp_dir, cf)
                if not os.path.exists(cf_path):
                    continue

                tree = ET.parse(cf_path)
                root = tree.getroot()
                tables = self._find_tables(root)

                for tbl in tables:
                    if current_table_idx == table_idx:
                        for ri, row_data in enumerate(data):
                            for ci, cell_value in enumerate(row_data):
                                target_row = start_row + ri
                                target_col = start_col + ci
                                if self._set_cell_value(tbl, target_row, target_col, str(cell_value)):
                                    cells_filled += 1
                        tree.write(cf_path, encoding="utf-8", xml_declaration=True)
                        modified = True
                        break
                    current_table_idx += 1

                if modified:
                    break

            if not modified:
                shutil.rmtree(temp_dir)
                return {
                    "success": False,
                    "error": f"Table index {table_idx} not found",
                }

            self._repack_zip(temp_dir, output_path)
            shutil.rmtree(temp_dir)

            return {
                "success": True,
                "message": f"Filled {cells_filled} cells in table[{table_idx}]",
                "cells_filled": cells_filled,
                "output": output_path,
            }
        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return {"success": False, "error": str(e)}

    def _find_tables(self, root):
        """Find table elements in XML tree"""
        tables = []
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag.lower() in ("tbl", "table"):
                tables.append(elem)
        return tables

    def _get_table_rows(self, tbl_elem):
        """Get row elements from table"""
        rows = []
        for child in tbl_elem.iter():
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag.lower() in ("tr", "row"):
                rows.append(child)
        return rows

    def _get_row_cells(self, row_elem):
        """Get cell elements from row"""
        cells = []
        for child in row_elem:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag.lower() in ("tc", "cell", "td"):
                cells.append(child)
        return cells

    def _set_cell_value(self, tbl_elem, row, col, value):
        """Set text value of a table cell"""
        rows = self._get_table_rows(tbl_elem)
        if row >= len(rows):
            return False

        cells = self._get_row_cells(rows[row])
        if col >= len(cells):
            return False

        cell = cells[col]
        # Find and modify text element in cell
        text_set = False
        for elem in cell.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag.lower() in ("t",):
                elem.text = value
                text_set = True
                break

        # If no text element found, try to set via run element
        if not text_set:
            for elem in cell.iter():
                tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if tag.lower() in ("run", "r"):
                    for sub in elem:
                        sub_tag = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                        if sub_tag.lower() in ("t",):
                            sub.text = value
                            text_set = True
                            break
                    if not text_set:
                        # Create t element if missing
                        ns = elem.tag.split("}")[0] + "}" if "}" in elem.tag else ""
                        t_elem = ET.SubElement(elem, f"{ns}t")
                        t_elem.text = value
                        text_set = True
                    break

        return text_set


def main():
    import argparse

    parser = argparse.ArgumentParser(description="HWPX file editor")
    subparsers = parser.add_subparsers(dest="command", help="Command")

    # replace
    p_replace = subparsers.add_parser("replace", help="Replace text")
    p_replace.add_argument("filepath", help="HWPX file path")
    p_replace.add_argument("--find", required=True, help="Text to find")
    p_replace.add_argument("--replace", required=True, help="Replacement text")
    p_replace.add_argument("--output", help="Output file path")

    # cell
    p_cell = subparsers.add_parser("cell", help="Edit table cell")
    p_cell.add_argument("filepath", help="HWPX file path")
    p_cell.add_argument("--table", type=int, required=True, help="Table index (0-based)")
    p_cell.add_argument("--row", type=int, required=True, help="Row (0-based)")
    p_cell.add_argument("--col", type=int, required=True, help="Column (0-based)")
    p_cell.add_argument("--value", required=True, help="New value")
    p_cell.add_argument("--output", help="Output file path")

    # fill-table
    p_fill = subparsers.add_parser("fill-table", help="Bulk fill table data")
    p_fill.add_argument("filepath", help="HWPX file path")
    p_fill.add_argument("--table", type=int, required=True, help="Table index (0-based)")
    p_fill.add_argument("--data", required=True, help="JSON 2D array")
    p_fill.add_argument("--start-row", type=int, default=0, help="Start row")
    p_fill.add_argument("--start-col", type=int, default=0, help="Start column")
    p_fill.add_argument("--output", help="Output file path")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    editor = HwpxEditor(args.filepath)
    output = getattr(args, "output", None) or args.filepath

    if args.command == "replace":
        result = editor.replace_text(args.find, args.replace, output)
    elif args.command == "cell":
        result = editor.edit_table_cell(args.table, args.row, args.col, args.value, output)
    elif args.command == "fill-table":
        data = json.loads(args.data)
        result = editor.fill_table(args.table, data, args.start_row, args.start_col, output)

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
