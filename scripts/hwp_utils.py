#!/usr/bin/env python3
"""HWP/HWPX 통합 유틸리티 - 파일 감지 및 통합 CLI 엔트리포인트"""
import os
import sys
import json
import argparse
import zipfile
import struct


def detect_file_type(filepath):
    """파일 타입을 자동 감지한다 (.hwp / .hwpx / unknown)"""
    if not os.path.exists(filepath):
        return "not_found"

    ext = os.path.splitext(filepath)[1].lower()

    # 확장자 기반 1차 판별
    if ext == ".hwpx":
        return "hwpx"
    if ext == ".hwp":
        # OLE2 매직 바이트 확인
        try:
            with open(filepath, "rb") as f:
                magic = f.read(8)
                if magic[:4] == b"\xd0\xcf\x11\xe0":
                    return "hwp"
        except Exception:
            pass
        return "hwp"

    # 확장자 없으면 내용 기반 판별
    try:
        with open(filepath, "rb") as f:
            magic = f.read(8)
            if magic[:4] == b"\xd0\xcf\x11\xe0":
                return "hwp"
            if magic[:4] == b"PK\x03\x04":
                # ZIP인지 확인 후 HWPX 구조 확인
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
    """파일 기본 정보를 반환한다"""
    if not os.path.exists(filepath):
        return {"error": f"파일을 찾을 수 없습니다: {filepath}"}

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
    parser = argparse.ArgumentParser(description="HWP/HWPX 통합 도구")
    subparsers = parser.add_subparsers(dest="command", help="명령어")

    # info
    p_info = subparsers.add_parser("info", help="파일 정보 조회")
    p_info.add_argument("filepath", help="HWP/HWPX 파일 경로")

    # text
    p_text = subparsers.add_parser("text", help="텍스트 추출")
    p_text.add_argument("filepath", help="HWP/HWPX 파일 경로")

    # tables
    p_tables = subparsers.add_parser("tables", help="표 추출")
    p_tables.add_argument("filepath", help="HWP/HWPX 파일 경로")

    # meta
    p_meta = subparsers.add_parser("meta", help="메타데이터 조회")
    p_meta.add_argument("filepath", help="HWP/HWPX 파일 경로")

    # replace (hwpx only)
    p_replace = subparsers.add_parser("replace", help="텍스트 치환 (HWPX만)")
    p_replace.add_argument("filepath", help="HWPX 파일 경로")
    p_replace.add_argument("--find", required=True, help="찾을 텍스트")
    p_replace.add_argument("--replace", required=True, help="바꿀 텍스트")
    p_replace.add_argument("--output", help="출력 파일 경로 (미지정시 원본 덮어쓰기)")

    # cell (hwpx only)
    p_cell = subparsers.add_parser("cell", help="표 셀 수정 (HWPX만)")
    p_cell.add_argument("filepath", help="HWPX 파일 경로")
    p_cell.add_argument("--table", type=int, required=True, help="표 인덱스 (0부터)")
    p_cell.add_argument("--row", type=int, required=True, help="행 (0부터)")
    p_cell.add_argument("--col", type=int, required=True, help="열 (0부터)")
    p_cell.add_argument("--value", required=True, help="새 값")
    p_cell.add_argument("--output", help="출력 파일 경로")

    # fill-table (hwpx only)
    p_fill = subparsers.add_parser("fill-table", help="표 데이터 일괄 입력 (HWPX만)")
    p_fill.add_argument("filepath", help="HWPX 파일 경로")
    p_fill.add_argument("--table", type=int, required=True, help="표 인덱스 (0부터)")
    p_fill.add_argument("--data", required=True, help="JSON 데이터 (2차원 배열)")
    p_fill.add_argument("--start-row", type=int, default=0, help="시작 행")
    p_fill.add_argument("--start-col", type=int, default=0, help="시작 열")
    p_fill.add_argument("--output", help="출력 파일 경로")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    filepath = args.filepath
    file_type = detect_file_type(filepath)

    if file_type == "not_found":
        print(json.dumps({"error": f"파일을 찾을 수 없습니다: {filepath}"}, ensure_ascii=False))
        sys.exit(1)

    if args.command == "info":
        print(json.dumps(get_file_info(filepath), ensure_ascii=False, indent=2))
        return

    # 읽기 명령 분기
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
            result = {"error": f"지원하지 않는 파일 형식입니다: {file_type}"}

        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    # 수정 명령 (hwpx만)
    if args.command in ("replace", "cell", "fill-table"):
        if file_type != "hwpx":
            print(json.dumps({"error": "수정 기능은 HWPX 파일만 지원합니다. HWP 파일은 읽기만 가능합니다."}, ensure_ascii=False))
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
