#!/bin/bash
# 安装后脚本：创建启动脚本 + 桌面启动器

# ASCII 命名的启动脚本，避免中文路径问题
cat > /opt/cert-splitter/start.sh << 'STARTEOF'
#!/bin/bash
# 设置库搜索路径（兼容麒麟/UOS 不同版本）
export LD_LIBRARY_PATH="/opt/cert-splitter:$LD_LIBRARY_PATH"

LOG="$HOME/.cert_splitter/logs/startup.log"
mkdir -p "$(dirname "$LOG")"

# 启动程序，输出到日志
exec "/opt/cert-splitter/批签发拆分工具" >> "$LOG" 2>&1
STARTEOF
chmod +x /opt/cert-splitter/start.sh

# 创建桌面启动器（Exec 使用 ASCII 路径）
cat > /usr/share/applications/cert-splitter.desktop << 'EOF'
[Desktop Entry]
Version=1.0
Type=Application
Name=批签发证明拆分工具
Comment=批签发大 PDF 自动拆分工具
Exec=/opt/cert-splitter/start.sh
Icon=/opt/cert-splitter/icon.png
Terminal=false
Categories=Office;
StartupNotify=false
EOF
chmod 644 /usr/share/applications/cert-splitter.desktop

# 刷新桌面数据库（部分桌面环境需要）
update-desktop-database /usr/share/applications/ 2>/dev/null || true
