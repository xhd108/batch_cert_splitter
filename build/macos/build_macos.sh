#!/usr/bin/env bash
# macOS 打包脚本
# 需要：pip install pyinstaller pyobjc-framework-Vision pyobjc-framework-Quartz
# 可选：brew install create-dmg

set -euo pipefail

ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
VERSION="${VERSION:-1.0.0}"
ARCH="$(uname -m)"   # arm64 or x86_64

echo "=== 批签发拆分工具 macOS 打包 v${VERSION} (${ARCH}) ==="

cd "$ROOT"

# 清理旧产物
rm -rf dist build/__pycache__

# PyInstaller 打包
pyinstaller build/cert_splitter.spec \
    --distpath dist \
    --workpath build/_pyinstaller_work \
    --noconfirm

APP_PATH="dist/批签发拆分工具.app"

if [ ! -d "$APP_PATH" ]; then
    echo "ERROR: 未找到 .app 包，PyInstaller 可能失败"
    exit 1
fi

echo "✓ .app 打包完成: $APP_PATH"

# ── 生成 DMG ─────────────────────────────────────────────────
if command -v create-dmg &>/dev/null; then
    DMG_NAME="批签发拆分工具_v${VERSION}_macOS_${ARCH}.dmg"
    create-dmg \
        --volname "批签发拆分工具 v${VERSION}" \
        --window-pos 200 120 \
        --window-size 600 400 \
        --icon-size 128 \
        --icon "批签发拆分工具.app" 150 180 \
        --hide-extension "批签发拆分工具.app" \
        --app-drop-link 450 180 \
        "dist/${DMG_NAME}" \
        "$APP_PATH"
    echo "✓ DMG 已生成: dist/${DMG_NAME}"
else
    echo "⚠ create-dmg 未安装，跳过 DMG 生成。安装方法：brew install create-dmg"
fi

echo "=== 完成 ==="
