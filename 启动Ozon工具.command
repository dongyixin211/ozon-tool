#!/bin/bash
set -e
cd "$(dirname "$0")"

PYTHON_CMD=""
CODEX_PYTHON="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

supports_tk_gui() {
  "$1" - <<'PY' >/dev/null 2>&1
import tkinter as tk
root = tk.Tk()
patch = root.tk.call("info", "patchlevel")
root.destroy()
major, minor, *_ = [int(part) for part in patch.split(".")]
raise SystemExit(0 if (major, minor) >= (8, 6) else 1)
PY
}

for candidate in \
  "$CODEX_PYTHON" \
  "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3" \
  "/Library/Frameworks/Python.framework/Versions/3.11/bin/python3" \
  "/opt/homebrew/bin/python3" \
  "/usr/local/bin/python3" \
  "$(command -v python3 2>/dev/null || true)" \
  "$(command -v python 2>/dev/null || true)"
do
  if [ -n "$candidate" ] && [ -x "$candidate" ] && supports_tk_gui "$candidate"; then
    PYTHON_CMD="$candidate"
    break
  fi
done

if [ -z "$PYTHON_CMD" ]; then
  echo "没有找到可正常显示 GUI 的 Python。"
  echo "当前 macOS 系统自带 Python 的 Tk 版本过旧，可能导致窗口空白。"
  echo "请安装 Python 3.11 或 3.12（python.org 或 Homebrew），然后重新启动。"
  echo "下载地址: https://www.python.org/downloads/macos/"
  read -r -p "按回车键关闭窗口..."
  exit 1
fi

export TK_SILENCE_DEPRECATION=1
echo "使用 Python: $PYTHON_CMD"

if ! "$PYTHON_CMD" -c "import requests; from PIL import Image; import openpyxl" >/dev/null 2>&1; then
  echo "首次运行，正在安装依赖..."
  "$PYTHON_CMD" -m pip install -r requirements.txt
fi

echo "正在启动 Ozon 工具..."
"$PYTHON_CMD" tool/app.py
echo
read -r -p "程序已关闭，按回车键关闭窗口..."
