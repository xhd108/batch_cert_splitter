#!/usr/bin/env python3
"""
批签发大 PDF 拆分工具 v1

子命令：
  init      初始化 SQLite 数据库
  template  生成人工填写用的 Excel 索引模板
  split     按索引表拆分 PDF，写入数据库，生成拆分报告
  list      查看已入库的单证记录

用法示例：
  python3 tools/batch_cert_splitter.py init
  python3 tools/batch_cert_splitter.py template
  python3 tools/batch_cert_splitter.py split --pdf 批签发.pdf --index 索引表.xlsx
  python3 tools/batch_cert_splitter.py list
"""
from __future__ import annotations

import argparse
import logging
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# ── 项目根目录（tools/ 的上级） ───────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_PATH      = PROJECT_ROOT / "db" / "cert.db"
OUTPUT_DIR   = PROJECT_ROOT / "output"
LOG_DIR      = PROJECT_ROOT / "logs"

# 页数不在此区间则标记为待复核
NORMAL_PAGE_RANGE = (2, 3)

# ── 日志 ──────────────────────────────────────────────────────
def _setup_logger() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"splitter_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logger = logging.getLogger("cert_splitter")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger

log = _setup_logger()


# ── 数据库 ────────────────────────────────────────────────────
def get_conn() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    with get_conn() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS batch_files (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            original_filename TEXT    NOT NULL,
            original_path    TEXT    NOT NULL UNIQUE,
            total_pages      INTEGER,
            import_time      TEXT,
            status           TEXT    DEFAULT 'pending'
        );

        CREATE TABLE IF NOT EXISTS cert_index (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_file_id    INTEGER NOT NULL REFERENCES batch_files(id),
            batch_no         TEXT    NOT NULL,
            vaccine_name     TEXT,
            manufacturer     TEXT,
            cert_no          TEXT,
            start_page       INTEGER NOT NULL,
            end_page         INTEGER NOT NULL,
            page_count       INTEGER,
            output_pdf       TEXT,
            review_status    TEXT    DEFAULT 'pending',
            split_time       TEXT,
            notes            TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_cert_batch_no ON cert_index(batch_no);
        CREATE INDEX IF NOT EXISTS idx_cert_file     ON cert_index(batch_file_id);
        """)
    log.info(f"数据库已就绪: {DB_PATH}")


# ── Excel 模板 ────────────────────────────────────────────────
TEMPLATE_COLS = [
    ("批号",     "batch_no",      "必填，与疫苗标签批号一致"),
    ("疫苗名称", "vaccine_name",  "如：冻干人用狂犬病疫苗（Vero细胞）"),
    ("生产企业", "manufacturer",  "如：长春卡介苗研究所"),
    ("批签发号", "cert_no",       "如：2024S02345"),
    ("起始页",   "start_page",    "必填，大PDF中该证明首页页码（从1起）"),
    ("结束页",   "end_page",      "必填，大PDF中该证明末页页码（含末页）"),
    ("备注",     "notes",         "选填"),
]

def cmd_template(args: argparse.Namespace) -> None:
    out = Path(args.out)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "批签发索引"

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    hint_fill   = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    hdr_font    = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    hint_font   = Font(name="Calibri", italic=True, color="595959", size=9)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 第1行：列名
    for col, (label, _, _) in enumerate(TEMPLATE_COLS, 1):
        c = ws.cell(row=1, column=col, value=label)
        c.fill   = header_fill
        c.font   = hdr_font
        c.alignment = center

    # 第2行：填写说明
    for col, (_, _, hint) in enumerate(TEMPLATE_COLS, 1):
        c = ws.cell(row=2, column=col, value=hint)
        c.fill = hint_fill
        c.font = hint_font
        c.alignment = Alignment(wrap_text=True)

    # 示例数据行
    ws.append(["202501AB", "冻干人用狂犬病疫苗", "长春卡介苗研究所", "2025S00123", 1, 2, ""])
    ws.append(["202501CD", "冻干人用狂犬病疫苗", "长春卡介苗研究所", "2025S00124", 3, 5, "3页，已复核"])

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 20
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 30

    wb.save(out)
    log.info(f"索引模板已生成: {out}")
    print(f"\n请用 Excel 打开 {out}，填写每个批号对应的起止页，保存后再运行 split。")
    print("  第1行：列名（勿删勿改）")
    print("  第2行：说明（可删）")
    print("  第3行起：每行一个批号\n")


# ── 读取 Excel 索引 ───────────────────────────────────────────
def _read_index_excel(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # 找列头行（第1行，跳过说明行）
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel 为空")

    header_map = {h: i for i, h in enumerate(rows[0]) if h}
    required = {"批号", "起始页", "结束页"}
    missing = required - set(header_map.keys())
    if missing:
        raise ValueError(f"索引表缺少必填列：{missing}")

    def _get(row, name):
        idx = header_map.get(name)
        return row[idx] if idx is not None else None

    records = []
    for row_num, row in enumerate(rows[1:], start=2):
        batch_no = str(_get(row, "批号") or "").strip()
        if not batch_no or batch_no.lower() in ("填写说明", "hint", "备注"):
            continue  # 跳过空行和说明行
        try:
            start = int(_get(row, "起始页"))
            end   = int(_get(row, "结束页"))
        except (TypeError, ValueError):
            raise ValueError(f"第 {row_num} 行：起始页/结束页必须为整数，批号={batch_no!r}")
        records.append({
            "batch_no":      batch_no,
            "vaccine_name":  str(_get(row, "疫苗名称") or "").strip() or None,
            "manufacturer":  str(_get(row, "生产企业") or "").strip() or None,
            "cert_no":       str(_get(row, "批签发号") or "").strip() or None,
            "start_page":    start,
            "end_page":      end,
            "notes":         str(_get(row, "备注") or "").strip() or None,
        })
    return records


# ── 质控 ─────────────────────────────────────────────────────
def _validate(records: list[dict], total_pages: int) -> list[str]:
    errors = []
    seen_ranges: list[tuple[int, int, str]] = []

    for i, r in enumerate(records, 1):
        s, e, bn = r["start_page"], r["end_page"], r["batch_no"]
        if s < 1:
            errors.append(f"行{i} 批号={bn}：起始页 {s} < 1")
        if e > total_pages:
            errors.append(f"行{i} 批号={bn}：结束页 {e} 超出 PDF 总页数 {total_pages}")
        if s > e:
            errors.append(f"行{i} 批号={bn}：起始页 {s} > 结束页 {e}")
        for ps, pe, pb in seen_ranges:
            if s <= pe and e >= ps:
                errors.append(f"行{i} 批号={bn} [{s}-{e}] 与 批号={pb} [{ps}-{pe}] 页码重叠")
        seen_ranges.append((s, e, bn))

    # 检查页码覆盖是否有空缺
    if records:
        all_pages = set()
        for r in records:
            all_pages.update(range(r["start_page"], r["end_page"] + 1))
        for pg in range(1, total_pages + 1):
            if pg not in all_pages:
                errors.append(f"警告：第 {pg} 页未被任何批号覆盖（可能遗漏）")

    return errors


# ── 核心拆分 ─────────────────────────────────────────────────
def cmd_split(args: argparse.Namespace) -> None:
    pdf_path   = Path(args.pdf).resolve()
    index_path = Path(args.index).resolve()
    out_dir    = Path(args.out_dir).resolve() if args.out_dir else OUTPUT_DIR / pdf_path.stem

    if not pdf_path.exists():
        log.error(f"PDF 文件不存在: {pdf_path}")
        sys.exit(1)
    if not index_path.exists():
        log.error(f"索引文件不存在: {index_path}")
        sys.exit(1)

    out_dir.mkdir(parents=True, exist_ok=True)
    init_db()

    # ── 读索引 ──────────────────────────────────────────────
    log.info(f"读取索引: {index_path}")
    records = _read_index_excel(index_path)
    log.info(f"共 {len(records)} 条记录")

    # ── 打开 PDF ────────────────────────────────────────────
    log.info(f"打开 PDF: {pdf_path}")
    src = fitz.open(str(pdf_path))
    total_pages = len(src)
    log.info(f"PDF 总页数: {total_pages}")

    # ── 质控 ────────────────────────────────────────────────
    errors = _validate(records, total_pages)
    hard_errors = [e for e in errors if not e.startswith("警告")]
    for e in errors:
        if e.startswith("警告"):
            log.warning(e)
        else:
            log.error(e)
    if hard_errors:
        log.error(f"发现 {len(hard_errors)} 个错误，终止拆分。请修正索引表后重试。")
        sys.exit(1)

    # ── 注册源文件 ──────────────────────────────────────────
    now = datetime.now().isoformat(timespec="seconds")
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT OR IGNORE INTO batch_files "
            "(original_filename, original_path, total_pages, import_time, status) "
            "VALUES (?,?,?,?,'pending')",
            (pdf_path.name, str(pdf_path), total_pages, now)
        )
        if cur.lastrowid:
            file_id = cur.lastrowid
        else:
            file_id = conn.execute(
                "SELECT id FROM batch_files WHERE original_path=?", (str(pdf_path),)
            ).fetchone()["id"]

    # ── 逐条拆分 ────────────────────────────────────────────
    results = []
    for r in records:
        s, e = r["start_page"], r["end_page"]   # 1-indexed
        page_count = e - s + 1

        # 文件命名：批号_批签发号.pdf 或 批号.pdf
        stem = r["batch_no"]
        if r["cert_no"]:
            stem = f"{r['batch_no']}_{r['cert_no']}"
        out_pdf = out_dir / f"{stem}.pdf"

        # 复核状态：页数异常则标为 abnormal
        review = "ok"
        notes_parts = [r["notes"]] if r["notes"] else []
        if not (NORMAL_PAGE_RANGE[0] <= page_count <= NORMAL_PAGE_RANGE[1]):
            review = "abnormal"
            notes_parts.append(f"页数={page_count}，超出正常范围{NORMAL_PAGE_RANGE}，需人工复核")

        # 用 PyMuPDF 提取页面（0-indexed）
        writer = fitz.open()
        writer.insert_pdf(src, from_page=s - 1, to_page=e - 1)
        writer.save(str(out_pdf))
        writer.close()
        log.info(f"  拆分: {r['batch_no']} [{s}-{e}页] → {out_pdf.name}  {review}")

        results.append({**r,
                        "page_count":   page_count,
                        "output_pdf":   str(out_pdf),
                        "review_status": review,
                        "notes":        "；".join(notes_parts) or None})

    # ── 写入数据库 ───────────────────────────────────────────
    split_time = datetime.now().isoformat(timespec="seconds")
    with get_conn() as conn:
        for r in results:
            conn.execute(
                "INSERT INTO cert_index "
                "(batch_file_id, batch_no, vaccine_name, manufacturer, cert_no, "
                " start_page, end_page, page_count, output_pdf, review_status, split_time, notes) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                (file_id, r["batch_no"], r["vaccine_name"], r["manufacturer"],
                 r["cert_no"], r["start_page"], r["end_page"], r["page_count"],
                 r["output_pdf"], r["review_status"], split_time, r["notes"])
            )
        conn.execute("UPDATE batch_files SET status='split' WHERE id=?", (file_id,))

    # ── 拆分报告 ────────────────────────────────────────────
    _print_report(pdf_path, total_pages, results, out_dir)


def _print_report(pdf_path: Path, total_pages: int, results: list[dict], out_dir: Path) -> None:
    ok_count       = sum(1 for r in results if r["review_status"] == "ok")
    abnormal_count = sum(1 for r in results if r["review_status"] == "abnormal")

    sep = "─" * 60
    print(f"\n{sep}")
    print(f"  拆分报告")
    print(sep)
    print(f"  源文件     : {pdf_path.name}")
    print(f"  PDF总页数  : {total_pages}")
    print(f"  拆分总数   : {len(results)}")
    print(f"  正常       : {ok_count}")
    print(f"  待复核     : {abnormal_count}")
    print(f"  输出目录   : {out_dir}")
    print(sep)
    print(f"  {'批号':<20} {'页':<8} {'页数':>4}  状态")
    print(f"  {'─'*20} {'─'*8} {'─'*4}  ──────")
    for r in results:
        flag = "⚠ 待复核" if r["review_status"] == "abnormal" else "✓"
        print(f"  {r['batch_no']:<20} {r['start_page']}-{r['end_page']:<5} {r['page_count']:>4}  {flag}")
    print(sep)
    if abnormal_count:
        print("  ⚠ 以下批号页数异常，请人工核查：")
        for r in results:
            if r["review_status"] == "abnormal":
                print(f"    - {r['batch_no']}  {r['start_page']}-{r['end_page']}页  ({r['notes']})")
    print()


# ── list 命令 ────────────────────────────────────────────────
def cmd_list(args: argparse.Namespace) -> None:
    init_db()
    limit = args.limit
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT c.batch_no, c.vaccine_name, c.cert_no, c.start_page, c.end_page, "
            "       c.page_count, c.review_status, b.original_filename "
            "FROM cert_index c JOIN batch_files b ON b.id=c.batch_file_id "
            "ORDER BY c.id DESC LIMIT ?",
            (limit,)
        ).fetchall()
    if not rows:
        print("数据库中暂无记录。")
        return
    print(f"\n最近 {len(rows)} 条记录：")
    print(f"  {'批号':<20} {'疫苗名称':<20} {'批签发号':<15} {'页':<8} {'页数':>4}  状态       源文件")
    print("  " + "─" * 95)
    for r in rows:
        print(f"  {r['batch_no']:<20} {(r['vaccine_name'] or ''):<20} "
              f"{(r['cert_no'] or ''):<15} "
              f"{r['start_page']}-{r['end_page']:<4} {r['page_count']:>4}  "
              f"{r['review_status']:<10} {r['original_filename']}")
    print()


# ── CLI ──────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(
        description="批签发大PDF拆分工具 v1",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("init", help="初始化 SQLite 数据库")

    tp = sub.add_parser("template", help="生成 Excel 索引模板")
    tp.add_argument("--out", default="批签发索引模板.xlsx", help="输出文件名")

    sp = sub.add_parser("split", help="按索引表拆分 PDF")
    sp.add_argument("--pdf",     required=True, help="大 PDF 路径")
    sp.add_argument("--index",   required=True, help="已填写的 Excel 索引表路径")
    sp.add_argument("--out-dir", default=None,  help="拆分 PDF 输出目录（默认 output/<pdf名>/）")

    lp = sub.add_parser("list", help="查看已入库的单证记录")
    lp.add_argument("--limit", type=int, default=50, help="最多显示条数（默认50）")

    args = parser.parse_args()

    if args.cmd == "init":
        init_db()
    elif args.cmd == "template":
        cmd_template(args)
    elif args.cmd == "split":
        cmd_split(args)
    elif args.cmd == "list":
        cmd_list(args)


if __name__ == "__main__":
    main()
