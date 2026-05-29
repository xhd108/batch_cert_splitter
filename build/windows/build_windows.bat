@echo off
REM Windows 打包脚本
REM 需要：pip install pyinstaller pytesseract pillow opencv-python-headless
REM 需要：安装 Tesseract-OCR https://github.com/UB-Mannheim/tesseract/wiki
REM 可选：Inno Setup https://jrsoftware.org/isdl.php

setlocal EnableDelayedExpansion

set VERSION=%VERSION%
if "%VERSION%"=="" set VERSION=1.0.0

set ROOT=%~dp0..\..
cd /d "%ROOT%"

echo === 批签发拆分工具 Windows 打包 v%VERSION% ===

REM 清理
if exist dist rmdir /s /q dist
if exist build\_pyinstaller_work rmdir /s /q build\_pyinstaller_work

REM PyInstaller 打包
pyinstaller build\cert_splitter.spec ^
    --distpath dist ^
    --workpath build\_pyinstaller_work ^
    --noconfirm

if not exist "dist\批签发拆分工具" (
    echo ERROR: 打包失败，未找到输出目录
    exit /b 1
)

echo ✓ 打包完成: dist\批签发拆分工具

REM ── 复制 Tesseract 语言包 ──────────────────────────────────
REM 如果 Tesseract 已安装在默认位置，将 chi_sim.traineddata 复制进包
set TESS_DIR=C:\Program Files\Tesseract-OCR\tessdata
if exist "%TESS_DIR%\chi_sim.traineddata" (
    mkdir "dist\批签发拆分工具\tessdata" 2>nul
    copy "%TESS_DIR%\chi_sim.traineddata" "dist\批签发拆分工具\tessdata\" >nul
    copy "%TESS_DIR%\eng.traineddata"     "dist\批签发拆分工具\tessdata\" >nul
    echo ✓ Tesseract 语言包已复制
) else (
    echo ⚠ 未找到 Tesseract 语言包，OCR 功能可能受限
)

REM ── Inno Setup 制作安装包 ──────────────────────────────────
set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist %ISCC% (
    %ISCC% "build\windows\installer.iss" /DMyAppVersion=%VERSION%
    echo ✓ 安装包已生成
) else (
    echo ⚠ 未找到 Inno Setup，跳过安装包生成
)

echo === 完成 ===
