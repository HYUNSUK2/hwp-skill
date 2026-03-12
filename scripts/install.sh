#!/bin/bash
# HWP Tools 의존성 설치
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
pip install -r "$SCRIPT_DIR/requirements.txt"
