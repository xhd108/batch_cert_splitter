# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec 文件
用法：
  pyinstaller build/cert_splitter.spec
"""
import sys
from pathlib import Path

ROOT = Path(SPECPATH).parent  # 项目根目录

block_cipher = None

a = Analysis(
    [str(ROOT / "build" / "main.py")],
    pathex=[
        str(ROOT / "app"),
        str(ROOT / "tools"),
    ],
    binaries=[],
    datas=[
        # 静态文件打包进 _MEIPASS/static/
        (str(ROOT / "app" / "static"), "static"),
    ],
    hiddenimports=[
        # Flask 相关
        "flask",
        "werkzeug",
        "jinja2",
        "click",
        "itsdangerous",
        # 核心依赖
        "fitz",
        "openpyxl",
        "requests",
        "urllib3",
        "charset_normalizer",
        # OCR — macOS
        "Vision",
        "Quartz",
        # OCR — Tesseract（Windows/Linux）
        "pytesseract",
        "PIL",
        "PIL.Image",
        # QR — zxing-cpp（首选，支持 ECI）+ cv2（兜底）
        "zxingcpp",
        "cv2",
        # SQLite（内置，通常不需要，显式列出防遗漏）
        "sqlite3",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter",
        "matplotlib",
        "scipy",
        "pandas",
        "IPython",
        "notebook",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="批签发拆分工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # 不显示黑色命令行窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Windows 图标（如有）
    # icon=str(ROOT / "build" / "windows" / "icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="批签发拆分工具",
)

# macOS .app bundle
app = BUNDLE(
    coll,
    name="批签发拆分工具.app",
    # icon=str(ROOT / "build" / "macos" / "icon.icns"),
    bundle_identifier="cn.gov.cdc.cert-splitter",
    info_plist={
        "NSHighResolutionCapable": True,
        "LSUIElement": False,
        "CFBundleShortVersionString": "1.0.0",
        "NSCameraUsageDescription": "",
    },
)
