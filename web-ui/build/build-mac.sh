#!/usr/bin/env bash
# Build macOS .app for testing on Ravi's machine.
# Run: ./build-mac.sh
set -euo pipefail

cd "$(dirname "$0")/.."
ROOT="$(pwd)"

echo "==> 1. Build React frontend"
cd "$ROOT/frontend"
[ -d node_modules ] || npm install
npm run build

echo "==> 2. Install Python build deps"
cd "$ROOT/backend"
python3 -m pip install --quiet -r requirements.txt
python3 -m pip install --quiet pyinstaller

echo "==> 3. Run PyInstaller"
cd "$ROOT/build"
rm -rf build dist
python3 -m PyInstaller accrual.spec --noconfirm

echo "==> Done. App: $ROOT/build/dist/7t Accrual Updater.app"
echo "   Run: open '$ROOT/build/dist/7t Accrual Updater.app'"
