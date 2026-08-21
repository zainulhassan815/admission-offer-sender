#!/bin/bash
# Build a standalone app. Run on the OS you want to ship for — PyInstaller cannot cross-compile.
set -e
cd "$(dirname "$0")"
rm -rf build dist
uv run --with pyinstaller pyinstaller \
  --name "Admission Offer Sender" \
  --onedir --windowed --clean --noconfirm \
  --add-data "templates:templates" \
  --add-data "static:static" \
  --add-data "offer-letter.docx:." \
  --collect-data docx \
  app.py
echo
echo "Built: dist/Admission Offer Sender.app"
