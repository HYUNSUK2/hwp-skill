#!/usr/bin/env python3
"""HWP 바이너리 (v5) 파일 읽기 - olefile 기반"""
import os
import sys
import json
import struct
import zlib


class HwpReader:
    """HWP v5 바이너리 파일 리더 (olefile 기반)"""

    def __init__(self, filepath):
        self.filepath = filepath
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {filepath}")

    def read_text(self):
        """HWP 파일에서 텍스트를 추출한다"""
        try:
            import olefile
        except ImportError:
            return {
                "error": "olefile 패키지가 필요합니다. 설치: pip install olefile",
            }

        try:
            ole = olefile.OleFileIO(self.filepath)
            texts = []

            # BodyText 스트림에서 텍스트 추출
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
                    "message": "텍스트를 추출할 수 없습니다. 파일이 암호화되었거나 손상되었을 수 있습니다.",
                }

            return {"text": "\n\n".join(texts), "sections": len(texts)}
        except Exception as e:
            return {"error": str(e)}

    def _extract_text_from_bodytext(self, data):
        """BodyText 스트림 바이너리에서 텍스트를 추출한다"""
        # HWP v5 BodyText는 압축되어 있을 수 있다
        try:
            decompressed = zlib.decompress(data, -15)
        except Exception:
            decompressed = data

        texts = []
        # HWP 레코드 구조 파싱
        offset = 0
        while offset < len(decompressed):
            if offset + 4 > len(decompressed):
                break

            # 레코드 헤더 (4바이트)
            header = struct.unpack_from("<I", decompressed, offset)[0]
            tag_id = header & 0x3FF
            level = (header >> 10) & 0x3FF
            size = (header >> 20) & 0xFFF

            offset += 4

            if size == 0xFFF:
                # 확장 크기 (4바이트 추가)
                if offset + 4 > len(decompressed):
                    break
                size = struct.unpack_from("<I", decompressed, offset)[0]
                offset += 4

            if offset + size > len(decompressed):
                break

            record_data = decompressed[offset: offset + size]
            offset += size

            # HWPTAG_PARA_TEXT = 67 (문단 텍스트)
            if tag_id == 67:
                text = self._parse_para_text(record_data)
                if text.strip():
                    texts.append(text)

        return "\n".join(texts)

    def _parse_para_text(self, data):
        """문단 텍스트 레코드를 파싱한다"""
        chars = []
        i = 0
        while i < len(data) - 1:
            code = struct.unpack_from("<H", data, i)[0]
            i += 2

            if code == 0:
                break
            elif code < 32:
                # 제어 문자
                if code == 13:  # 줄바꿈
                    chars.append("\n")
                elif code == 10:  # 탭
                    chars.append("\t")
                elif code in (1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23):
                    # 인라인/확장 제어 문자 - 추가 바이트 건너뛰기
                    if code in (1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23):
                        # 확장 제어: 추가 데이터를 건너뛴다
                        pass
                # 다른 제어 문자는 무시
            else:
                try:
                    chars.append(chr(code))
                except (ValueError, OverflowError):
                    pass

        return "".join(chars)

    def read_metadata(self):
        """HWP 파일 메타데이터를 읽는다"""
        try:
            import olefile
        except ImportError:
            return {"error": "olefile 패키지가 필요합니다. 설치: pip install olefile"}

        try:
            ole = olefile.OleFileIO(self.filepath)
            meta = {"filename": os.path.basename(self.filepath)}

            # OLE 메타데이터
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

            # FileHeader 스트림에서 HWP 버전 정보
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

            # 스트림 목록
            meta["streams"] = ["/".join(s) for s in ole.listdir()]

            ole.close()
            return meta
        except Exception as e:
            return {"error": str(e)}

    def read_tables(self):
        """HWP 파일에서 표를 추출한다 (제한적 지원)"""
        try:
            import olefile
        except ImportError:
            return {"error": "olefile 패키지가 필요합니다. 설치: pip install olefile"}

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

                # 레코드 파싱하여 표 구조 추출
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

                    # HWPTAG_TABLE = 78 (표 시작)
                    if tag_id == 78:
                        if current_table:
                            tables.append(current_table)
                        current_table = []
                        in_table = True

                    # HWPTAG_PARA_TEXT = 67 (문단 텍스트, 표 내부 셀)
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
                    "note": "HWP 바이너리에서 표의 행/열 구조를 정확히 복원하기 어렵습니다. 셀 내용만 순서대로 표시됩니다.",
                })

            return {
                "tables": result_tables,
                "count": len(result_tables),
                "note": "HWP 바이너리 표 추출은 제한적입니다. 정확한 행/열 구조가 필요하면 HWPX 포맷을 사용하세요.",
            }
        except Exception as e:
            return {"error": str(e)}


def main():
    import argparse

    parser = argparse.ArgumentParser(description="HWP 바이너리 파일 읽기 도구")
    parser.add_argument("command", choices=["text", "tables", "meta"],
                        help="명령어: text, tables, meta")
    parser.add_argument("filepath", help="HWP 파일 경로")
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
