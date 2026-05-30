# 批签发证明拆分工具

> 面向基层疾控中心免疫规划科的批签发 PDF 自动拆分工具。
> 双击启动，浏览器操作，无需编程经验。

**介绍页** → https://xhd108.github.io/batch_cert_splitter/

---

## 功能

- **二维码优先识别**：扫描每份证明首页 QR，直连国家药监局数据库获取批号、疫苗名称、生产企业、批签发号
- **OCR 智能兜底**：QR 失败时自动 OCR；macOS 使用系统 Vision 框架，Windows / 国产系统使用 Tesseract
- **透视矫正预处理**：自动矫正手机拍摄的梯形变形，提升扫描件识别率
- **在线编辑复核**：识别结果在浏览器表格内直接修改，颜色标注置信度（QR / OCR 高 / OCR 中 / 手动）
- **一键拆分下载**：各批号 PDF 自动命名，打包为 ZIP；同步生成 Excel 索引表
- **本地历史记录**：拆分结果写入本机 SQLite，可随时查询历史批次，数据不离本机

---

## 平台支持

| 平台 | 要求 | 安装包 |
|---|---|---|
| Windows | Windows 10 / 11（64 位） | `.exe` 安装程序 |
| macOS | macOS 12 及以上（Apple Silicon / Intel） | `.dmg` |
| 麒麟 Kylin | V10（aarch64 / x86_64） | `.deb` |
| UOS | UOS 20 及以上（aarch64 / x86_64） | `.deb` |

前往 [Releases](https://github.com/xhd108/batch_cert_splitter/releases) 下载对应平台安装包。

---

## 使用流程

```
上传批签发合并 PDF
    ↓
自动逐页识别（二维码 → OCR 兜底），实时显示进度
    ↓
浏览器内复核结果表格，可直接编辑批号、页码等字段
    ↓
点击"确认并拆分"，下载 ZIP 压缩包 + Excel 索引
```

---

## 本地开发

### 环境要求

- Python 3.10+
- macOS：`pyobjc-framework-Vision`（系统 OCR）
- Windows / Linux：`tesseract-ocr`（含中文语言包）

### 安装依赖

```bash
pip install flask pymupdf openpyxl requests zxing-cpp opencv-python pyobjc-framework-Vision
```

Windows / Linux 额外安装：

```bash
pip install pytesseract pillow
# Ubuntu/Kylin/UOS：
sudo apt install tesseract-ocr tesseract-ocr-chi-sim
```

### 启动 Web 服务

```bash
python app/server.py
# 自动打开浏览器 http://127.0.0.1:5050
```

### 命令行模式（高级）

```bash
# 自动识别，生成复核表
python tools/batch_cert_splitter.py detect --pdf 批签发合并.pdf

# 按复核表拆分
python tools/batch_cert_splitter.py split \
  --pdf 批签发合并.pdf \
  --index 批签发合并_复核表.xlsx \
  --out output/
```

---

## 打包发布

推送 `vX.Y.Z` tag 后，GitHub Actions 自动在四个平台编译并发布 Release：

```bash
git tag v1.0.0
git push origin v1.0.0
```

本地手动打包（macOS）：

```bash
bash build/macos/build_macos.sh
```

---

## 目录结构

```
batch_cert_splitter/
├── tools/
│   └── batch_cert_splitter.py   核心逻辑（识别、拆分、入库）
├── app/
│   ├── server.py                Flask 本地 Web 服务
│   └── static/                  前端（HTML / CSS / JS）
├── build/
│   ├── cert_splitter.spec       PyInstaller 配置
│   ├── macos/                   macOS 打包脚本
│   ├── windows/                 Windows 打包脚本 + Inno Setup
│   └── linux/                   Kylin / UOS 打包脚本
├── docs/                        GitHub Pages 介绍页
├── db/                          SQLite 数据库（运行时生成）
└── .github/workflows/
    └── release.yml              四平台自动构建 CI
```

---

## 识别说明

| 标记 | 来源 | 说明 |
|---|---|---|
| 🔵 QR | 国家药监局 API | 字段完全准确，无需复核 |
| 🟢 OCR 高 | OCR + 正则 | 批号置信度高，可直接使用 |
| 🟡 OCR 中 | OCR + 正则 | 建议人工核查 |
| 🔴 手动 | 未识别 | 扫描质量过低，需手动填写 |

---

## 反馈与贡献

发现问题请在 [Issues](https://github.com/xhd108/batch_cert_splitter/issues) 提交，附上操作系统版本和问题描述。
