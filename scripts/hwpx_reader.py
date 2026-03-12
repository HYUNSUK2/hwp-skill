#!/usr/bin/env python3
"""HWPX 파일 읽기 - ZIP + XML 파싱"""
import os
import sys
import json
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict


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
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "urn:oasis:names:tc:opendocument:xmlns:container",
    "odf": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
}

# 네임스페이스 등록
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class HwpxReader:
    def __init__(self, filepath):
        self.filepath = filepath
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {filepath}")

    def _open_zip(self):
        return zipfile.ZipFile(self.filepath, "r")

    def _find_content_files(self, zf):
        """content XML 파일들을 찾는다 (section0.xml, section1.xml, ...)"""
        content_files = []
        for name in sorted(zf.namelist()):
            lower = name.lower()
            if "contents/" in lower and lower.endswith(".xml") and "section" in lower:
                content_files.append(name)
        # section 파일이 없으면 content.xml 찾기
        if not content_files:
            for name in zf.namelist():
                lower = name.lower()
                if lower.endswith("content.xml") or lower.endswith("contents.xml"):
                    content_files.append(name)
        return content_files

    def _extract_text_from_element(self, elem):
        """XML 요소에서 텍스트를 재귀적으로 추출한다"""
        texts = []
        # 직접 텍스트
        if elem.text and elem.text.strip():
            texts.append(elem.text.strip())

        for child in elem:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            # 텍스트 요소 — 직접 텍스트 추출
            if tag in ("t", "T"):
                if child.text:
                    texts.append(child.text)
            # 무시할 요소
            elif tag in ("linesegarray", "lineBreak", "LineSeg"):
                pass
            # 문단 — 내부 탐색 후 줄바꿈 추가
            elif tag in ("p", "P"):
                sub = self._extract_text_from_element(child)
                if sub:
                    texts.append(sub)
                texts.append("\n")
            # 줄바꿈
            elif tag == "br":
                texts.append("\n")
            # 그 외 — 재귀 탐색
            else:
                sub = self._extract_text_from_element(child)
                if sub:
                    texts.append(sub)

            if child.tail and child.tail.strip():
                texts.append(child.tail.strip())

        return "".join(texts)

    def read_text(self):
        """전체 텍스트를 추출한다"""
        try:
            with self._open_zip() as zf:
                content_files = self._find_content_files(zf)
                if not content_files:
                    return {"error": "콘텐츠 파일을 찾을 수 없습니다", "files": zf.namelist()}

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
        """표 데이터를 추출한다"""
        try:
            with self._open_zip() as zf:
                content_files = self._find_content_files(zf)
                if not content_files:
                    return {"error": "콘텐츠 파일을 찾을 수 없습니다"}

                tables = []
                table_idx = 0

                for cf in content_files:
                    xml_data = zf.read(cf)
                    root = ET.fromstring(xml_data)

                    # 모든 테이블 요소 찾기
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
        """XML 트리에서 테이블 요소를 재귀적으로 찾는다"""
        tables = []
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag.lower() in ("tbl", "table"):
                tables.append(elem)
        return tables

    def _parse_table(self, tbl_elem):
        """테이블 요소를 2차원 배열로 파싱한다"""
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
        """문서 메타데이터를 읽는다"""
        try:
            with self._open_zip() as zf:
                meta = {"filename": os.path.basename(self.filepath)}

                # META-INF/manifest.xml 확인
                for name in zf.namelist():
                    lower = name.lower()
                    if "meta" in lower and lower.endswith(".xml"):
                        try:
                            xml_data = zf.read(name)
                            root = ET.fromstring(xml_data)
                            meta[name] = self._extract_metadata_fields(root)
                        except Exception:
                            pass

                # header.xml 확인
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
        """메타데이터 XML에서 필드를 추출한다"""
        fields = {}
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if elem.text and elem.text.strip():
                key = tag.lower()
                fields[key] = elem.text.strip()
        return fields

    def read_structure(self):
        """문서 구조를 트리 형태로 반환한다"""
        try:
            with self._open_zip() as zf:
                structure = {"files": {}}
                for name in sorted(zf.namelist()):
                    info = zf.getinfo(name)
                    structure["files"][name] = {
                        "size": info.file_size,
                        "compressed": info.compress_size,
                    }
                    # XML 파일이면 루트 태그 정보 추가
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

    parser = argparse.ArgumentParser(description="HWPX 파일 읽기 도구")
    parser.add_argument("command", choices=["text", "tables", "meta", "structure"],
                        help="명령어: text, tables, meta, structure")
    parser.add_argument("filepath", help="HWPX 파일 경로")
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
