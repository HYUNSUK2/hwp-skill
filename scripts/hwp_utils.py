#!/usr/bin/env python3
"""HWP/HWPX unified utility - file type detection and unified CLI entry point"""
import os
import sys
import json
import argparse
import zipfile
import struct


def detect_file_type(filepath):
    """Auto-detect file type (.hwp / .hwpx / unknown)"""
    if not os.path.exists(filepath):
        return "not_found"

    ext = os.path.splitext(filepath)[1].lower()

    # Extension-based detection
    if ext == ".hwpx":
        return "hwpx"
    if ext == ".hwp":
        # Verify OLE2 magic bytes
        try:
            with open(filepath, "rb") as f:
                magic = f.read(8)
                if magic[:4] == b"\xd0\xcf\x11\xe0":
                    return "hwp"
        except Exception:
            pass
        return "hwp"

    # Content-based detection if no extension
    try:
        with open(filepath, "rb") as f:
            magic = f.read(8)
            if magic[:4] == b"\xd0\xcf\x11\xe0":
                return "hwp"
            if magic[:4] == b"PK\x03\x04":
                # Check if ZIP contains HWPX structure
                try:
                    with zipfile.ZipFile(filepath, "r") as zf:
                        names = zf.namelist()
                        if any("content.xml" in n.lower() for n in names):
                            return "hwpx"
                except Exception:
                    pass
    except Exception:
        pass

    return "unknown"


def get_file_info(filepath):
    """Return basic file information"""
    if not os.path.exists(filepath):
        return {"error": f"File not found: {filepath}"}

    file_type = detect_file_type(filepath)
    stat = os.stat(filepath)

    info = {
        "path": os.path.abspath(filepath),
        "filename": os.path.basename(filepath),
        "type": file_type,
        "size_bytes": stat.st_size,
        "size_human": _human_size(stat.st_size),
    }

    if file_type == "hwpx":
        try:
            with zipfile.ZipFile(filepath, "r") as zf:
                info["hwpx_entries"] = zf.namelist()
        except Exception as e:
            info["hwpx_error"] = str(e)

    return info


def _human_size(size_bytes):
    for unit in ["B", "KB", "MB", "GB"]:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def main():
    parser = argparse.ArgumentParser(description="HWP/HWPX unified tool")
    subparsers = parser.add_subparsers(dest="command", help="Command")

    # info
    p_info = subparsers.add_parser("info", help="Show file info")
    p_info.add_argument("filepath", help="HWP/HWPX file path")

    # text
    p_text = subparsers.add_parser("text", help="Extract text")
    p_text.add_argument("filepath", help="HWP/HWPX file path")

    # tables
    p_tables = subparsers.add_parser("tables", help="Extract tables")
    p_tables.add_argument("filepath", help="HWP/HWPX file path")

    # meta
    p_meta = subparsers.add_parser("meta", help="Show metadata")
    p_meta.add_argument("filepath", help="HWP/HWPX file path")

    # replace (hwpx only)
    p_replace = subparsers.add_parser("replace", help="Replace text (HWPX only)")
    p_replace.add_argument("filepath", help="HWPX file path")
    p_replace.add_argument("--find", required=True, help="Text to find")
    p_replace.add_argument("--replace", required=True, help="Replacement text")
    p_replace.add_argument("--output", help="Output file path (overwrites original if omitted)")

    # cell (hwpx only)
    p_cell = subparsers.add_parser("cell", help="Edit table cell (HWPX only)")
    p_cell.add_argument("filepath", help="HWPX file path")
    p_cell.add_argument("--table", type=int, required=True, help="Table index (0-based)")
    p_cell.add_argument("--row", type=int, required=True, help="Row (0-based)")
    p_cell.add_argument("--col", type=int, required=True, help="Column (0-based)")
    p_cell.add_argument("--value", required=True, help="New value")
    p_cell.add_argument("--output", help="Output file path")

    # fill-table (hwpx only)
    p_fill = subparsers.add_parser("fill-table", help="Bulk fill table data (HWPX only)")
    p_fill.add_argument("filepath", help="HWPX file path")
    p_fill.add_argument("--table", type=int, required=True, help="Table index (0-based)")
    p_fill.add_argument("--data", required=True, help="JSON 2D array data")
    p_fill.add_argument("--start-row", type=int, default=0, help="Start row")
    p_fill.add_argument("--start-col", type=int, default=0, help="Start column")
    p_fill.add_argument("--output", help="Output file path")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    filepath = args.filepath
    file_type = detect_file_type(filepath)

    if file_type == "not_found":
        print(json.dumps({"error": f"File not found: {filepath}"}, ensure_ascii=False))
        sys.exit(1)

    if args.command == "info":
        print(json.dumps(get_file_info(filepath), ensure_ascii=False, indent=2))
        return

    # Read commands
    if args.command in ("text", "tables", "meta"):
        if file_type == "hwp":
            from hwp_reader import HwpReader
            reader = HwpReader(filepath)
            if args.command == "text":
                result = reader.read_text()
            elif args.command == "tables":
                result = reader.read_tables()
            elif args.command == "meta":
                result = reader.read_metadata()
        elif file_type == "hwpx":
            from hwpx_reader import HwpxReader
            reader = HwpxReader(filepath)
            if args.command == "text":
                result = reader.read_text()
            elif args.command == "tables":
                result = reader.read_tables()
            elif args.command == "meta":
                result = reader.read_metadata()
        else:
            result = {"error": f"Unsupported file type: {file_type}"}

        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    # Edit commands (hwpx only)
    if args.command in ("replace", "cell", "fill-table"):
        if file_type != "hwpx":
            print(json.dumps({"error": "Edit operations are only supported for HWPX files. HWP files are read-only."}, ensure_ascii=False))
            sys.exit(1)

        from hwpx_editor import HwpxEditor
        editor = HwpxEditor(filepath)
        output = getattr(args, "output", None) or filepath

        if args.command == "replace":
            result = editor.replace_text(args.find, args.replace, output)
        elif args.command == "cell":
            result = editor.edit_table_cell(args.table, args.row, args.col, args.value, output)
        elif args.command == "fill-table":
            data = json.loads(args.data)
            result = editor.fill_table(args.table, data, args.start_row, args.start_col, output)

        print(json.dumps(result, ensure_ascii=False, indent=2))
        return


if __name__ == "__main__":
    main()
