#!/usr/bin/env bash
set -euo pipefail

WORK_DIR=".build-macos"
DIST_DIR="release-macos"

if [ -d "${WORK_DIR}" ]; then
  chmod -R u+rwX "${WORK_DIR}" 2>/dev/null || true
  rm -rf "${WORK_DIR}"
fi
if [ -d "${DIST_DIR}" ]; then
  chmod -R u+rwX "${DIST_DIR}" 2>/dev/null || true
  rm -rf "${DIST_DIR}"
fi

python3 -m pip install -r requirements.txt
python3 -m PyInstaller \
  --clean \
  --noconfirm \
  --workpath "${WORK_DIR}" \
  --distpath "${DIST_DIR}" \
  --noconsole \
  --windowed \
  --name "Generate Accurate XML" \
  --osx-bundle-identifier "com.local.generate-accurate-xml" \
  app.py

echo
echo "Build selesai."
echo "Aplikasi macOS ada di: release-macos/Generate Accurate XML.app"
