#!/usr/bin/env bash
set -euo pipefail

APP_NAME="Generate Accurate XML"
IDENTIFIER="com.local.generate-accurate-xml"
VERSION="$(date +%Y.%m.%d.%H%M%S)"
WORK_DIR=".build-macos-${VERSION}"
DIST_DIR="release-macos-${VERSION}"
PAYLOAD_ROOT=".pkg-root"

if [ -d "${WORK_DIR}" ]; then
  chmod -R u+rwX "${WORK_DIR}" 2>/dev/null || true
  rm -rf "${WORK_DIR}"
fi
if [ -d "${DIST_DIR}" ]; then
  chmod -R u+rwX "${DIST_DIR}" 2>/dev/null || true
  rm -rf "${DIST_DIR}"
fi
if [ -d "${PAYLOAD_ROOT}" ]; then
  chmod -R u+rwX "${PAYLOAD_ROOT}" 2>/dev/null || true
  rm -rf "${PAYLOAD_ROOT}"
fi

python3 -m pip install -r requirements.txt
python3 -m PyInstaller \
  --clean \
  --noconfirm \
  --workpath "${WORK_DIR}" \
  --distpath "${DIST_DIR}" \
  --noconsole \
  --windowed \
  --name "${APP_NAME}" \
  --osx-bundle-identifier "${IDENTIFIER}" \
  app.py

rm -f "${DIST_DIR}/${APP_NAME}.pkg"

mkdir -p "${PAYLOAD_ROOT}/Applications"
ditto "${DIST_DIR}/${APP_NAME}.app" "${PAYLOAD_ROOT}/Applications/${APP_NAME}.app"

pkgbuild \
  --root "${PAYLOAD_ROOT}" \
  --identifier "${IDENTIFIER}" \
  --version "${VERSION}" \
  --install-location "/" \
  "${DIST_DIR}/${APP_NAME}.pkg"

echo
echo "Build selesai."
echo "Installer macOS ada di: ${DIST_DIR}/${APP_NAME}.pkg"
