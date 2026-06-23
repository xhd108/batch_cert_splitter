#!/usr/bin/env bash
# Kylin / UOS 打包脚本（在 Docker 容器内运行）
# 生成 .deb 安装包，兼容 aarch64 和 x86_64

set -euo pipefail

ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
VERSION="${VERSION:-1.0.0}"
ARCH="$(uname -m)"   # aarch64 or x86_64

echo "=== 批签发拆分工具 Linux 打包 v${VERSION} (${ARCH}) ==="

# 安装系统依赖
apt-get update -qq
apt-get install -y --no-install-recommends \
    tesseract-ocr tesseract-ocr-chi-sim \
    libgl1 libglib2.0-0 libgomp1 \
    ruby ruby-dev rubygems build-essential

gem install fpm --no-document

pip install --no-cache-dir \
    flask pymupdf openpyxl requests \
    zxing-cpp opencv-python-headless pyinstaller \
    pytesseract pillow

cd "$ROOT"
rm -rf dist build/__pycache__

pyinstaller build/cert_splitter.spec \
    --distpath dist \
    --workpath build/_pyinstaller_work \
    --noconfirm

# 复制 Tesseract 语言包进程序目录
mkdir -p "dist/批签发拆分工具/tessdata"
cp /usr/share/tesseract-ocr/*/tessdata/{chi_sim,eng}.traineddata \
   "dist/批签发拆分工具/tessdata/" 2>/dev/null || true

# 打包 .deb
VER="${VERSION#v}"   # strip leading 'v' if present
DEB_NAME="cert-splitter_v${VER}_linux_${ARCH}.deb"
fpm -s dir -t deb \
    --name "cert-splitter" \
    --version "$VERSION" \
    --architecture "$( [ "$ARCH" = "aarch64" ] && echo arm64 || echo amd64 )" \
    --maintainer "批签发工具 <noreply@example.com>" \
    --description "批签发证明拆分工具" \
    --url "https://github.com/xhd108/batch_cert_splitter" \
    --after-install build/linux/postinstall.sh \
    "dist/批签发拆分工具/=/opt/cert-splitter/"

mv *.deb "dist/${DEB_NAME}" 2>/dev/null || true
echo "✓ .deb 已生成: dist/${DEB_NAME}"
