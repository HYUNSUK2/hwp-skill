#!/bin/bash
# Install HWP Tools dependencies
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
pip install -r "$SCRIPT_DIR/requirements.txt"
