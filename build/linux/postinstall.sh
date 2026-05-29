#!/bin/bash
# 安装后脚本：创建桌面启动器
cat > /usr/share/applications/cert-splitter.desktop << 'EOF'
[Desktop Entry]
Version=1.0
Type=Application
Name=批签发证明拆分工具
Comment=批签发大 PDF 自动拆分工具
Exec=/opt/cert-splitter/批签发拆分工具
Icon=/opt/cert-splitter/icon.png
Terminal=false
Categories=Office;
EOF
chmod 644 /usr/share/applications/cert-splitter.desktop
