#!/bin/zsh
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

PYTHON_BIN="${PYTHON_BIN:-python3}"
VENV_DIR="$ROOT/.venv-macos"

if [[ ! -d "$VENV_DIR" ]]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r "$ROOT/requirements-macos.txt"
python -m PyInstaller --clean --noconfirm "$ROOT/供需协同工具V3.0_macOS.spec"

APP_PATH="$ROOT/dist/供需协同工具V3.0.app"
codesign --force --deep --sign - "$APP_PATH"

echo ""
echo "构建完成：$APP_PATH"
echo "首次运行请在 Finder 中右键应用并选择‘打开’。"
