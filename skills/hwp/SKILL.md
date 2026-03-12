---
name: hwp
description: macOS/Linux에서 HWP 및 HWPX 파일을 읽고 수정하는 도구. HWP 파일 텍스트 추출, 표 읽기, 메타데이터 조회, HWPX 파일 텍스트 치환 및 표 수정이 가능하다. 사용자가 HWP 파일을 열거나, HWP 내용을 확인하거나, HWP 문서를 수정하거나, 한글 파일을 다루려 할 때 이 스킬을 사용한다.
---

# HWP/HWPX 파일 도구

macOS/Linux에서 HWP(한글) 파일을 읽고 수정하는 Python 스크립트 도구.

## 지원 범위

| 기능 | .hwp (바이너리) | .hwpx (XML) |
|------|:---:|:---:|
| 텍스트 추출 | O | O |
| 표 추출 | O (셀 내용만) | O (행/열 구조 포함) |
| 메타데이터 | O | O |
| 텍스트 치환 | X | O |
| 표 셀 수정 | X | O |

## 사전 요구사항

의존성이 설치되어 있지 않으면 먼저 설치한다:

```bash
pip install olefile lxml cryptography
```

또는 설치 스크립트 실행:

```bash
bash ${CLAUDE_PLUGIN_ROOT}/scripts/install.sh
```

## 사용법

모든 스크립트는 `${CLAUDE_PLUGIN_ROOT}/scripts/` 경로에 있다.
통합 엔트리포인트 `hwp_utils.py`를 사용하거나, 개별 스크립트를 직접 호출할 수 있다.

### 파일 정보 확인

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py info /path/to/file.hwp
```

자동으로 HWP/HWPX 포맷을 감지하고 파일 정보를 JSON으로 출력한다.

### 텍스트 추출

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py text /path/to/file.hwp
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py text /path/to/file.hwpx
```

파일 타입을 자동 감지하여 적절한 리더를 사용한다.

### 표 추출

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py tables /path/to/file.hwpx
```

표 데이터를 JSON으로 출력한다. HWPX는 행/열 구조가 포함되고, HWP는 셀 내용만 순서대로 표시된다.

### 메타데이터 조회

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py meta /path/to/file.hwp
```

### 텍스트 치환 (HWPX만)

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py replace /path/to/file.hwpx \
  --find "홍길동" --replace "김철수" --output /path/to/output.hwpx
```

`--output`을 생략하면 원본 파일을 덮어쓴다. 원본 보존이 필요하면 반드시 `--output`을 지정하거나, 실행 전에 원본을 백업한다.

### 표 셀 수정 (HWPX만)

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py cell /path/to/file.hwpx \
  --table 0 --row 1 --col 2 --value "수정할 값" --output /path/to/output.hwpx
```

- `--table`: 표 인덱스 (0부터 시작)
- `--row`, `--col`: 셀 위치 (0부터 시작)

### 표 데이터 일괄 입력 (HWPX만)

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py fill-table /path/to/file.hwpx \
  --table 0 --data '[["이름","나이"],["홍길동","30"],["김철수","25"]]' \
  --output /path/to/output.hwpx
```

- `--data`: JSON 2차원 배열
- `--start-row`, `--start-col`: 입력 시작 위치 (기본 0,0)

## 워크플로우 가이드

### HWP 파일 내용 확인

1. `info` 명령으로 파일 타입 확인
2. `text` 명령으로 텍스트 추출
3. `tables` 명령으로 표 확인 (있다면)

### HWPX 파일 수정

1. `text` 또는 `tables`로 현재 내용 확인
2. `replace`, `cell`, 또는 `fill-table`로 수정
3. 수정 결과를 `text` 또는 `tables`로 확인

### 주의사항

- HWP (바이너리) 파일은 읽기만 가능하다. 수정이 필요하면 HWPX로 변환 후 작업한다.
- HWPX 수정 시 원본 백업을 권장한다. `--output`으로 별도 파일에 저장하면 안전하다.
- 모든 출력은 JSON 형식이다.
- `olefile` 패키지가 없으면 HWP 바이너리 읽기가 실패한다. 에러 메시지에 설치 방법이 안내된다.
