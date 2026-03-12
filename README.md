# HWP Tools

macOS/Linux에서 HWP 및 HWPX(한글) 파일을 읽고 수정하는 Claude Code 플러그인.

## About

한글(HWP)은 한국에서 가장 널리 사용되는 문서 편집기이지만, macOS/Linux에서는 파일을 다루기 어렵습니다. HWP Tools는 Claude Code 스킬로 동작하며, Python 스크립트를 통해 HWP/HWPX 파일의 텍스트 추출, 표 읽기, 텍스트 치환, 표 수정 등을 지원합니다.

## Features

| 기능 | .hwp (바이너리) | .hwpx (XML) |
|------|:---:|:---:|
| 텍스트 추출 | O | O |
| 표 추출 | O (셀 내용만) | O (행/열 구조 포함) |
| 메타데이터 조회 | O | O |
| 텍스트 치환 | - | O |
| 표 셀 수정 | - | O |
| 표 일괄 입력 | - | O |

## Installation

### 1. 마켓플레이스 추가

Claude Code에서:

```
/plugin marketplace add HYUNSUK2/hwp-skill
```

### 2. 플러그인 설치

```
/plugin install hwp@hwp-tools-marketplace
```

### 3. Python 의존성 설치

```bash
pip install olefile lxml cryptography
```

### Requirements

- Python 3.8+
- Claude Code (플러그인 지원 버전)

## Usage

플러그인 설치 후 Claude Code에서 HWP/HWPX 파일을 다루면 자동으로 스킬이 트리거됩니다.

### 텍스트 추출

```
이 HWP 파일의 내용을 읽어줘: /path/to/document.hwp
```

### 표 확인

```
이 HWPX 파일에 있는 표를 보여줘: /path/to/document.hwpx
```

### 텍스트 치환

```
이 HWPX 파일에서 "홍길동"을 "김철수"로 바꿔줘: /path/to/document.hwpx
```

### 표 수정

```
이 HWPX 파일의 첫 번째 표에서 2행 3열을 "수정값"으로 바꿔줘
```

## CLI 직접 사용

스킬 없이 Python 스크립트를 직접 사용할 수도 있습니다:

```bash
# 파일 정보
python scripts/hwp_utils.py info document.hwpx

# 텍스트 추출
python scripts/hwp_utils.py text document.hwp

# 표 추출
python scripts/hwp_utils.py tables document.hwpx

# 텍스트 치환 (HWPX만)
python scripts/hwp_utils.py replace document.hwpx --find "홍길동" --replace "김철수" --output output.hwpx

# 표 셀 수정 (HWPX만)
python scripts/hwp_utils.py cell document.hwpx --table 0 --row 1 --col 2 --value "새 값" --output output.hwpx

# 표 일괄 입력 (HWPX만)
python scripts/hwp_utils.py fill-table document.hwpx --table 0 \
  --data '[["이름","나이"],["홍길동","30"]]' --output output.hwpx
```

## Project Structure

```
hwp-skill/
├── .claude-plugin/
│   └── plugin.json           # 플러그인 매니페스트
├── skills/
│   └── hwp/
│       └── SKILL.md          # 스킬 정의
└── scripts/
    ├── requirements.txt      # Python 의존성
    ├── install.sh            # 의존성 설치 스크립트
    ├── hwp_utils.py          # 통합 CLI 엔트리포인트
    ├── hwp_reader.py         # HWP 바이너리 읽기 (olefile)
    ├── hwpx_reader.py        # HWPX 읽기 (ZIP + XML)
    └── hwpx_editor.py        # HWPX 수정
```

## Limitations

- **HWP 바이너리 수정 불가**: HWP v5 포맷은 읽기만 지원합니다. 수정이 필요하면 HWPX로 변환 후 작업하세요.
- **HWP 표 구조 제한**: HWP 바이너리에서 표의 행/열 구조를 정확히 복원하기 어렵습니다. 셀 내용만 순서대로 표시됩니다.
- **새 문서 생성 미지원**: 기존 파일의 읽기/수정만 지원합니다.
- **암호화된 파일 미지원**: 암호가 설정된 HWP 파일은 읽을 수 없습니다.

## License

MIT
