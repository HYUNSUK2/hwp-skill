# HWP Tools

A Claude Code plugin to read and edit HWP/HWPX (Korean Hangul Word Processor) files on macOS/Linux.

## About

HWP is the most widely used document format in South Korea, but handling HWP files on macOS/Linux has been difficult. HWP Tools is a Claude Code skill that uses Python scripts to extract text, read tables, replace text, and edit table cells in HWP/HWPX files.

## Features

| Feature | .hwp (binary) | .hwpx (XML) |
|---------|:---:|:---:|
| Text extraction | O | O |
| Table extraction | O (cell content only) | O (with row/col structure) |
| Metadata | O | O |
| Text replacement | - | O |
| Table cell editing | - | O |
| Bulk table fill | - | O |

## Installation

### 1. Add marketplace

In Claude Code:

```
/plugin marketplace add HYUNSUK2/hwp-skill
```

### 2. Install plugin

```
/plugin install hwp@hwp-tools-marketplace
```

### 3. Install Python dependencies

```bash
pip install olefile lxml cryptography
```

### Requirements

- Python 3.8+
- Claude Code (with plugin support)

## Usage

After installation, the skill triggers automatically when you work with HWP/HWPX files in Claude Code.

### Extract text

```
Read the contents of this file: /path/to/document.hwp
```

### Read tables

```
Show me the tables in /path/to/document.hwpx
```

### Replace text (HWPX only)

```
Replace "홍길동" with "김철수" in /path/to/document.hwpx
```

### Edit table cell (HWPX only)

```
Change row 2, column 3 of the first table to "100" in /path/to/file.hwpx
```

## CLI Usage

You can also use the Python scripts directly without the skill:

```bash
# File info
python scripts/hwp_utils.py info document.hwpx

# Extract text
python scripts/hwp_utils.py text document.hwp

# Read tables
python scripts/hwp_utils.py tables document.hwpx

# Replace text (HWPX only)
python scripts/hwp_utils.py replace document.hwpx --find "old" --replace "new" --output output.hwpx

# Edit table cell (HWPX only)
python scripts/hwp_utils.py cell document.hwpx --table 0 --row 1 --col 2 --value "new value" --output output.hwpx

# Bulk fill table (HWPX only)
python scripts/hwp_utils.py fill-table document.hwpx --table 0 \
  --data '[["Name","Age"],["Kim","30"]]' --output output.hwpx
```

## Project Structure

```
hwp-skill/
├── .claude-plugin/
│   └── plugin.json           # Plugin manifest
├── skills/
│   └── hwp/
│       └── SKILL.md          # Skill definition
└── scripts/
    ├── requirements.txt      # Python dependencies
    ├── install.sh            # Dependency installer
    ├── hwp_utils.py          # Unified CLI entry point
    ├── hwp_reader.py         # HWP binary reader (olefile)
    ├── hwpx_reader.py        # HWPX reader (ZIP + XML)
    └── hwpx_editor.py        # HWPX editor
```

## Limitations

- **HWP binary is read-only**: HWP v5 format only supports reading. Convert to HWPX for editing.
- **HWP table structure is limited**: Row/column structure cannot be fully reconstructed from HWP binary. Cell contents are listed in order.
- **No new document creation**: Only reading/editing of existing files is supported.
- **Encrypted files not supported**: Password-protected HWP files cannot be read.

## License

MIT
