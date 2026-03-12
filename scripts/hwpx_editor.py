#!/usr/bin/env python3
"""HWPX 파일 수정 - ZIP 압축 해제 → XML 수정 → 재압축"""
import os
import sys
import json
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET


# HWPX XML 네임스페이스
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
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {filepath}")

    def _find_content_files(self, zf):
        """content XML 파일들을 찾는다"""
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
        """임시 디렉토리를 다시 ZIP으로 압축한다"""
        # 원본의 파일 순서를 유지하기 위해 원본 목록을 먼저 가져온다
        with zipfile.ZipFile(self.filepath, "r") as orig_zf:
            original_names = orig_zf.namelist()

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in original_names:
                file_path = os.path.join(temp_dir, name)
                if os.path.exists(file_path):
                    zf.write(file_path, name)

            # 원본에 없는 새 파일 추가 (있을 경우)
            for root, dirs, files in os.walk(temp_dir):
                for f in files:
                    full_path = os.path.join(root, f)
                    arcname = os.path.relpath(full_path, temp_dir)
                    if arcname not in original_names:
                        zf.write(full_path, arcname)

    def replace_text(self, find_text, replace_text, output_path=None):
        """HWPX 파일에서 텍스트를 치환한다"""
        if output_path is None:
            output_path = self.filepath

        try:
            temp_dir = tempfile.mkdtemp()
            replace_count = 0

            # ZIP 압축 해제
            with zipfile.ZipFile(self.filepath, "r") as zf:
                zf.extractall(temp_dir)
                content_files = self._find_content_files(zf)

            # 각 콘텐츠 파일에서 텍스트 치환
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
                    "message": f"'{find_text}'를 찾을 수 없습니다",
                    "replacements": 0,
                }

            # 재압축
            self._repack_zip(temp_dir, output_path)
            shutil.rmtree(temp_dir)

            return {
                "success": True,
                "message": f"'{find_text}' → '{replace_text}' 치환 완료",
                "replacements": replace_count,
                "output": output_path,
            }
        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return {"success": False, "error": str(e)}

    def _replace_text_in_element(self, elem, find_text, replace_text):
        """XML 요소 내 텍스트를 재귀적으로 치환한다"""
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
        """HWPX 파일의 표 셀을 수정한다"""
        if output_path is None:
            output_path = self.filepath

        try:
            temp_dir = tempfile.mkdtemp()

            with zipfile.ZipFile(self.filepath, "r") as zf:
                zf.extractall(temp_dir)
                content_files = self._find_content_files(zf)

            # 테이블 찾기
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
                                "error": f"셀 ({row}, {col})을 찾을 수 없습니다",
                            }
                        break
                    current_table_idx += 1

                if modified:
                    break

            if not modified:
                shutil.rmtree(temp_dir)
                return {
                    "success": False,
                    "error": f"표 인덱스 {table_idx}를 찾을 수 없습니다 (총 {current_table_idx}개)",
                }

            self._repack_zip(temp_dir, output_path)
            shutil.rmtree(temp_dir)

            return {
                "success": True,
                "message": f"표[{table_idx}] 셀({row},{col}) = '{value}' 수정 완료",
                "output": output_path,
            }
        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return {"success": False, "error": str(e)}

    def fill_table(self, table_idx, data, start_row=0, start_col=0, output_path=None):
        """HWPX 파일의 표에 데이터를 일괄 입력한다"""
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
                    "error": f"표 인덱스 {table_idx}를 찾을 수 없습니다",
                }

            self._repack_zip(temp_dir, output_path)
            shutil.rmtree(temp_dir)

            return {
                "success": True,
                "message": f"표[{table_idx}]에 {cells_filled}개 셀 입력 완료",
                "cells_filled": cells_filled,
                "output": output_path,
            }
        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return {"success": False, "error": str(e)}

    def _find_tables(self, root):
        """XML 트리에서 테이블 요소를 찾는다"""
        tables = []
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag.lower() in ("tbl", "table"):
                tables.append(elem)
        return tables

    def _get_table_rows(self, tbl_elem):
        """테이블에서 행 요소들을 가져온다"""
        rows = []
        for child in tbl_elem.iter():
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag.lower() in ("tr", "row"):
                rows.append(child)
        return rows

    def _get_row_cells(self, row_elem):
        """행에서 셀 요소들을 가져온다"""
        cells = []
        for child in row_elem:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag.lower() in ("tc", "cell", "td"):
                cells.append(child)
        return cells

    def _set_cell_value(self, tbl_elem, row, col, value):
        """테이블 셀의 텍스트를 설정한다"""
        rows = self._get_table_rows(tbl_elem)
        if row >= len(rows):
            return False

        cells = self._get_row_cells(rows[row])
        if col >= len(cells):
            return False

        cell = cells[col]
        # 셀 내부의 텍스트 요소를 찾아서 수정
        text_set = False
        for elem in cell.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag.lower() in ("t",):
                elem.text = value
                text_set = True
                break

        # 텍스트 요소가 없으면 직접 설정 시도
        if not text_set:
            # 가장 깊은 paragraph/run을 찾아서 텍스트 설정
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
                        # t 요소가 없으면 생성
                        # 네임스페이스 추론
                        ns = elem.tag.split("}")[0] + "}" if "}" in elem.tag else ""
                        t_elem = ET.SubElement(elem, f"{ns}t")
                        t_elem.text = value
                        text_set = True
                    break

        return text_set


def main():
    import argparse

    parser = argparse.ArgumentParser(description="HWPX 파일 수정 도구")
    subparsers = parser.add_subparsers(dest="command", help="명령어")

    # replace
    p_replace = subparsers.add_parser("replace", help="텍스트 치환")
    p_replace.add_argument("filepath", help="HWPX 파일 경로")
    p_replace.add_argument("--find", required=True, help="찾을 텍스트")
    p_replace.add_argument("--replace", required=True, help="바꿀 텍스트")
    p_replace.add_argument("--output", help="출력 파일 경로")

    # cell
    p_cell = subparsers.add_parser("cell", help="표 셀 수정")
    p_cell.add_argument("filepath", help="HWPX 파일 경로")
    p_cell.add_argument("--table", type=int, required=True, help="표 인덱스 (0부터)")
    p_cell.add_argument("--row", type=int, required=True, help="행 (0부터)")
    p_cell.add_argument("--col", type=int, required=True, help="열 (0부터)")
    p_cell.add_argument("--value", required=True, help="새 값")
    p_cell.add_argument("--output", help="출력 파일 경로")

    # fill-table
    p_fill = subparsers.add_parser("fill-table", help="표 데이터 일괄 입력")
    p_fill.add_argument("filepath", help="HWPX 파일 경로")
    p_fill.add_argument("--table", type=int, required=True, help="표 인덱스 (0부터)")
    p_fill.add_argument("--data", required=True, help="JSON 2차원 배열")
    p_fill.add_argument("--start-row", type=int, default=0, help="시작 행")
    p_fill.add_argument("--start-col", type=int, default=0, help="시작 열")
    p_fill.add_argument("--output", help="출력 파일 경로")

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
