---
name: hwp
description: Read and edit HWP/HWPX (Korean Hangul Word Processor) files on macOS/Linux. Supports text extraction, table reading, metadata inspection, text replacement, and table cell editing. Use this skill when the user wants to open, read, edit, or inspect HWP or HWPX files, or work with Korean Hangul documents.
---

# HWP/HWPX File Tools

Python scripts for reading and editing HWP (Hangul Word Processor) files on macOS/Linux.

## Supported Features

| Feature | .hwp (binary) | .hwpx (XML) |
|---------|:---:|:---:|
| Text extraction | O | O |
| Table extraction | O (cell content only) | O (with row/col structure) |
| Metadata | O | O |
| Text replacement | X | O |
| Table cell editing | X | O |

## Prerequisites

Install dependencies if not already present:

```bash
pip install olefile lxml cryptography
```

Or run the install script:

```bash
bash ${CLAUDE_PLUGIN_ROOT}/scripts/install.sh
```

## Usage

All scripts are located at `${CLAUDE_PLUGIN_ROOT}/scripts/`.
Use the unified entry point `hwp_utils.py` or call individual scripts directly.

### File Info

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py info /path/to/file.hwp
```

Automatically detects HWP/HWPX format and outputs file info as JSON.

### Text Extraction

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py text /path/to/file.hwp
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py text /path/to/file.hwpx
```

Auto-detects file type and uses the appropriate reader.

### Table Extraction

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py tables /path/to/file.hwpx
```

Outputs table data as JSON. HWPX includes row/col structure; HWP lists cell contents in order.

### Metadata

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py meta /path/to/file.hwp
```

### Text Replacement (HWPX only)

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py replace /path/to/file.hwpx \
  --find "old text" --replace "new text" --output /path/to/output.hwpx
```

If `--output` is omitted, the original file is overwritten. Always use `--output` or back up the original to preserve it.

### Table Cell Editing (HWPX only)

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py cell /path/to/file.hwpx \
  --table 0 --row 1 --col 2 --value "new value" --output /path/to/output.hwpx
```

- `--table`: Table index (0-based)
- `--row`, `--col`: Cell position (0-based)

### Bulk Table Fill (HWPX only)

```bash
python ${CLAUDE_PLUGIN_ROOT}/scripts/hwp_utils.py fill-table /path/to/file.hwpx \
  --table 0 --data '[["Name","Age"],["Kim","30"],["Lee","25"]]' \
  --output /path/to/output.hwpx
```

- `--data`: JSON 2D array
- `--start-row`, `--start-col`: Starting position (default 0,0)

## Workflow Guide

### Reading HWP Files

1. Use `info` to check file type
2. Use `text` to extract text
3. Use `tables` to read tables (if any)

### Editing HWPX Files

1. Use `text` or `tables` to inspect current content
2. Use `replace`, `cell`, or `fill-table` to modify
3. Verify changes with `text` or `tables`

### Important Notes

- HWP (binary) files are read-only. Convert to HWPX for editing.
- Back up HWPX files before editing. Use `--output` to save to a separate file.
- All output is in JSON format.
- If `olefile` is not installed, HWP binary reading will fail with an installation guide in the error message.
