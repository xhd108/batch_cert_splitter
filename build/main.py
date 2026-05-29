#!/usr/bin/env python3
"""
PyInstaller 入口点
打包后双击可执行文件时，此脚本被调用。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# ── 修正 sys.path，让 app/ 和 tools/ 均可被 import ─────────────
if getattr(sys, "frozen", False):
    _base = Path(sys._MEIPASS)  # type: ignore[attr-defined]
else:
    _base = Path(__file__).resolve().parent.parent

sys.path.insert(0, str(_base / "app"))
sys.path.insert(0, str(_base / "tools"))

# ── 确保用户数据目录存在 ───────────────────────────────────────
_data = Path.home() / ".cert_splitter"
for _d in (_data / "db", _data / "output", _data / "logs"):
    _d.mkdir(parents=True, exist_ok=True)

os.environ.setdefault("CERT_SPLITTER_DATA", str(_data))

# ── 启动 Flask 服务（含自动打开浏览器） ─────────────────────────
from server import main  # noqa: E402

if __name__ == "__main__":
    main()
