#!/usr/bin/env bash
set -euo pipefail

pip3 install --quiet openpyxl pillow
python3 scripts/build-image-manifest.py
