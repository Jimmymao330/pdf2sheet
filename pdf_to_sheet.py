#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
會議資料 PDF →（自動找出分辦表頁面）→ docling JSON → 結構化 → Google 試算表

步驟：
  零、只讀 PDF 文字層，找出「主席裁（指）示暨決議事項執行情形分辦表」所在頁面
  一、只把這幾頁交給 docling 解析為 JSON（存於 --json-dir）
  二、讀取 JSON，合併跨頁斷開的表格列，整理成結構化資料
  三、依「列管案號｜主席裁示｜承辦單位｜最新辦理情形｜列管建議｜來源 PDF 名稱」
      送到 Google Apps Script 網頁應用程式（Code.gs），由它寫入試算表

用法：
  pdf2sheet init                    # 首次使用：設定 Apps Script 網址與密碼
  pdf2sheet ./pdfs --dry-run        # 只檢查解析結果、不上傳
  pdf2sheet ./pdfs                  # 解析並寫入試算表
  pdf2sheet ./pdfs --pages 11-17    # 自動找頁失敗時手動指定

  未用 pipx 安裝時，把 pdf2sheet 換成 python pdf_to_sheet.py

安裝：
  pipx install --python 3.12 <專案資料夾>
  （或 python -m pip install docling；掃描檔需 OCR 時另裝 ocrmac）
"""

from __future__ import annotations

import argparse
import csv
import difflib
import json
import logging
import os
import re
import ssl
import statistics
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

log = logging.getLogger("pdf_to_sheet")

SHEET_HEADER = ["列管案號", "主席裁示", "承辦單位", "最新辦理情形", "列管建議", "來源 PDF 名稱"]

# Google 試算表單一儲存格上限 50,000 字元
CELL_CHAR_LIMIT = 50000

# 內部欄位鍵 → 表頭關鍵字（表頭文字去空白後比對，依序判斷，先符合者優先）
COLUMN_KEYWORDS: list[tuple[str, tuple[str, ...]]] = [
    ("case_no", ("案號",)),
    ("seq", ("序號", "序")),
    ("directive", ("裁示", "裁（指）示", "裁(指)示", "決議")),
    ("unit", ("承辦", "單位")),
    ("status", ("辦理情形", "執行情形", "辦理")),
    ("suggestion", ("建議", "列管")),
]
# 找不到表頭時，6 欄表格的預設順序
DEFAULT_COLUMN_ORDER = ["seq", "case_no", "directive", "unit", "status", "suggestion"]

CASE_NO_RE = re.compile(r"\d{5,8}(?:\s*[-－–]\s*\d+)?")


# ──────────────────────────────────────────────────────────────
# 一、PDF → docling JSON
# ──────────────────────────────────────────────────────────────

def build_converter(use_ocr: bool):
    try:
        import docling  # noqa: F401
    except ImportError:
        sys.exit(
            "找不到 docling。目前執行的 Python 是：\n"
            f"  {sys.executable}\n"
            "請用同一個 Python 安裝：\n"
            f"  \"{sys.executable}\" -m pip install docling"
        )
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode
    from docling.document_converter import DocumentConverter, PdfFormatOption

    opts = PdfPipelineOptions()
    opts.do_table_structure = True
    opts.table_structure_options.mode = TableFormerMode.ACCURATE
    # 使用 PDF 內嵌文字對應儲存格，中文較不易被 OCR 誤判
    opts.table_structure_options.do_cell_matching = True
    opts.generate_page_images = False
    opts.generate_picture_images = False
    opts.do_ocr = use_ocr
    if use_ocr:
        # macOS 用內建 Vision（需 ocrmac）；Windows / 未安裝 ocrmac 時改用 EasyOCR
        use_mac = False
        if sys.platform == "darwin":
            try:
                import ocrmac  # noqa: F401
                use_mac = True
            except ImportError:
                log.warning("未安裝 ocrmac，改用 EasyOCR（首次使用會下載 OCR 模型）")
        if use_mac:
            from docling.datamodel.pipeline_options import OcrMacOptions
            opts.ocr_options = OcrMacOptions(lang=["zh-Hant", "en-US"])
        else:
            from docling.datamodel.pipeline_options import EasyOcrOptions
            opts.ocr_options = EasyOcrOptions(lang=["ch_tra", "en"])

    return DocumentConverter(
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)}
    )


def parse_pages_option(spec: str) -> list[int]:
    """「11-17,20」→ [11, 12, …, 17, 20]"""
    pages: set[int] = set()
    for part in spec.replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"(\d+)\s*[-~～]\s*(\d+)", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            pages.update(range(min(a, b), max(a, b) + 1))
        elif part.isdigit():
            pages.add(int(part))
        else:
            raise ValueError(f"--pages 格式錯誤：{part}（範例：11-17,20）")
    return sorted(pages)


# 分辦表標題：主席裁（指）示暨決議事項執行情形（分辦表）
TITLE_RE = re.compile(r"裁[（(]?指?[）)]?示暨決議事項執行情形")
# 分辦表的表頭關鍵字；頁面文字去空白後比對（表頭可能被拆成「列管／案號」兩列，故分開判斷）
TABLE_PAGE_KEYWORDS = ("案號", "承辦", "辦理情形")
# 續頁沒有表頭時的判斷依據：有勾選框或列管案號，且不是新的章節開頭
CONTINUATION_RE = re.compile(r"[除續]管|\d{7}(?:-\d+)?")
NEW_SECTION_RE = re.compile(r"^\s*[壹貳參肆伍陸柒捌玖拾]+、|工作報告|委員名冊|設置要點", re.M)
MAX_TABLE_PAGES = 80


def locate_table_pages(pdf_path: Path) -> list[int] | None:
    """只讀 PDF 文字層（不經 AI 模型，160 頁約 1 秒），找出分辦表所在頁碼（從 1 起算）。
    找不到時回傳 None。"""
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        raw: list[str] = []
        for i in range(len(pdf)):
            try:
                raw.append(pdf[i].get_textpage().get_text_range() or "")
            except Exception:
                raw.append("")
    finally:
        pdf.close()

    if sum(len(squash(t)) for t in raw) < 20 * max(len(raw), 1):
        log.warning("%s 幾乎沒有文字層（可能是掃描檔），無法自動找頁，請用 --pages 指定", pdf_path.name)
        return None

    flat = [squash(t) for t in raw]

    def is_table_page(i: int) -> bool:
        return all(k in flat[i] for k in TABLE_PAGE_KEYWORDS)

    def has_repeated_header(i: int) -> bool:
        # 分辦表每頁頂端都會重複表頭「序｜列管案號｜主席裁（指）示暨決議事項…」；
        # 後面的工作報告即使也提到案號、承辦、辦理情形，頁首也不會是這個表頭
        top = flat[i][:200]
        return is_table_page(i) and "案號" in top and "決議事項" in top

    def is_continuation(i: int) -> bool:
        head = raw[i].strip()[:300]
        return bool(CONTINUATION_RE.search(flat[i])) and not NEW_SECTION_RE.search(head)

    starts = [i for i in range(len(raw)) if is_table_page(i) and TITLE_RE.search(flat[i])]
    if not starts:
        # 標題可能是圖片或寫法不同：退而求其次，找第一個有分辦表表頭的頁面
        starts = [i for i in range(len(raw)) if is_table_page(i)][:1]
        if starts:
            log.warning("%s 找不到分辦表標題，改以表頭判斷（第 %d 頁起），請確認結果",
                        pdf_path.name, starts[0] + 1)
    if not starts:
        return None

    pages: set[int] = set()
    for s in starts:
        if s in pages:
            continue
        j = s
        while j < len(raw) and j - s < MAX_TABLE_PAGES:
            if j == s or has_repeated_header(j) or is_continuation(j):
                pages.add(j)
                j += 1
            else:
                break
    return sorted(p + 1 for p in pages)


def format_pages(pages: list[int]) -> str:
    out, start = [], None
    for i, p in enumerate(pages):
        if start is None:
            start = p
        if i + 1 == len(pages) or pages[i + 1] != p + 1:
            out.append(str(start) if start == p else f"{start}-{p}")
            start = None
    return ",".join(out)


def extract_pages(pdf_path: Path, pages: list[int], dest: Path) -> None:
    import pypdfium2 as pdfium

    src = pdfium.PdfDocument(str(pdf_path))
    new = pdfium.PdfDocument.new()
    try:
        new.import_pages(src, [p - 1 for p in pages])
        new.save(str(dest))
    finally:
        new.close()
        src.close()


def remap_page_numbers(node: Any, mapping: dict[int, int]) -> Any:
    """docling 解析的是只含分辦表的暫存 PDF，頁碼從 1 起算；改回原始 PDF 的頁碼。"""
    if isinstance(node, dict):
        out = {}
        for k, v in node.items():
            if k == "page_no" and isinstance(v, int):
                out[k] = mapping.get(v, v)
            elif k == "pages" and isinstance(v, dict):
                out[k] = {str(mapping.get(int(pk), int(pk))) if str(pk).isdigit() else pk:
                          remap_page_numbers(pv, mapping) for pk, pv in v.items()}
            else:
                out[k] = remap_page_numbers(v, mapping)
        return out
    if isinstance(node, list):
        return [remap_page_numbers(v, mapping) for v in node]
    return node


def resolve_pages(pdf: Path, pages_opt: str) -> list[int] | None:
    """回傳要解析的頁碼；None 代表整份 PDF。"""
    if pages_opt == "all":
        return None
    if pages_opt != "auto":
        return parse_pages_option(pages_opt)
    try:
        found = locate_table_pages(pdf)
    except Exception as exc:
        log.warning("%s 自動找頁失敗：%s", pdf.name, exc)
        found = None
    return found if found else []


def convert_pdfs(pdf_paths: list[Path], json_dir: Path, use_ocr: bool, force: bool,
                 pages_opt: str = "auto") -> list[Path]:
    import tempfile

    json_dir.mkdir(parents=True, exist_ok=True)
    converter = None
    out: list[Path] = []
    for pdf in pdf_paths:
        target = json_dir / f"{pdf.stem}.json"
        try:
            pages = resolve_pages(pdf, pages_opt)
        except ValueError as exc:
            log.error("%s", exc)
            return out
        if pages == []:
            log.warning("%s 找不到「主席裁（指）示暨決議事項執行情形分辦表」，已略過"
                        "（可用 --pages 11-17 手動指定頁碼，或 --pages all 解析整份）", pdf.name)
            continue
        page_label = "全部" if pages is None else format_pages(pages)
        log.info("%s：分辦表位於第 %s 頁", pdf.name, page_label)

        # 頁碼範圍沒變、PDF 也沒更新，就沿用上次的 JSON
        if not force and target.exists() and target.stat().st_mtime >= pdf.stat().st_mtime:
            try:
                with target.open(encoding="utf-8") as f:
                    cached = (json.load(f).get("_pdf2sheet") or {}).get("pages", "全部")
            except (OSError, json.JSONDecodeError):
                cached = None
            if cached == page_label:
                log.info("  略過轉檔（已有相同頁碼的 JSON）：%s", target.name)
                out.append(target)
                continue

        if converter is None:
            log.info("初始化 docling（首次執行會下載模型，請稍候）…")
            converter = build_converter(use_ocr)
        log.info("  docling 解析中…")
        try:
            if pages is None:
                doc = converter.convert(str(pdf)).document.export_to_dict()
            else:
                # Windows 上 docling 可能仍持有暫存檔，清除失敗不影響結果
                with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
                    part = Path(tmp) / f"{pdf.stem}.pdf"
                    extract_pages(pdf, pages, part)
                    doc = converter.convert(str(part)).document.export_to_dict()
                doc = remap_page_numbers(doc, {i + 1: p for i, p in enumerate(pages)})
        except Exception as exc:  # 單檔失敗不影響其他檔案
            log.error("解析失敗 %s：%s", pdf.name, exc)
            continue
        doc.setdefault("origin", {})["filename"] = pdf.name
        doc["_pdf2sheet"] = {"pages": page_label}
        with target.open("w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        out.append(target)
    return out


# ──────────────────────────────────────────────────────────────
# 二、讀取 JSON 並結構化
# ──────────────────────────────────────────────────────────────

@dataclass
class Record:
    case_no: str = ""
    seq: str = ""
    parts: dict[str, list[str]] = field(
        default_factory=lambda: {"directive": [], "unit": [], "status": [], "suggestion": []}
    )
    source_pdf: str = ""
    pages: list[int] = field(default_factory=list)

    def add(self, key: str, text: str) -> None:
        if key in self.parts and text.strip():
            self.parts[key].append(text)


def resolve_ref(doc: dict, ref: str) -> dict | None:
    # "#/tables/3" → doc["tables"][3]
    node: Any = doc
    for part in ref.lstrip("#/").split("/"):
        if isinstance(node, list):
            node = node[int(part)]
        else:
            node = node.get(part)
        if node is None:
            return None
    return node


def tables_in_reading_order(doc: dict) -> list[dict]:
    """依 body 的閱讀順序取出表格；若結構異常則改以頁碼排序。"""
    ordered: list[dict] = []
    seen: set[str] = set()

    def walk(node: dict) -> None:
        for child in node.get("children", []):
            ref = child.get("$ref", "")
            if ref in seen:
                continue
            seen.add(ref)
            item = resolve_ref(doc, ref)
            if item is None:
                continue
            if ref.startswith("#/tables/"):
                ordered.append(item)
            walk(item)

    walk(doc.get("body", {}))
    tables = doc.get("tables", [])
    if len(ordered) != len(tables):
        def page_of(t: dict) -> tuple[int, float]:
            prov = (t.get("prov") or [{}])[0]
            bbox = prov.get("bbox") or {}
            return prov.get("page_no", 0), -bbox.get("t", 0)
        ordered = sorted(tables, key=page_of)
    return ordered


def squash(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def table_grid(table: dict) -> list[list[dict]]:
    data = table.get("data", {})
    grid = data.get("grid")
    if grid:
        return grid
    # 舊版 docling 沒有 grid，用 table_cells 重建
    n_rows, n_cols = data.get("num_rows", 0), data.get("num_cols", 0)
    rebuilt = [[{"text": "", "column_header": False} for _ in range(n_cols)] for _ in range(n_rows)]
    for cell in data.get("table_cells", []):
        for r in range(cell["start_row_offset_idx"], cell["end_row_offset_idx"]):
            for c in range(cell["start_col_offset_idx"], cell["end_col_offset_idx"]):
                if r < n_rows and c < n_cols:
                    rebuilt[r][c] = cell
    return rebuilt


def is_header_row(row: list[dict]) -> bool:
    texts = [squash(c.get("text", "")) for c in row]
    non_empty = [t for t in texts if t]
    if not non_empty:
        return False
    if all(c.get("column_header") for c in row if squash(c.get("text", ""))):
        return True
    # 沒標 column_header，但內容就是表頭字樣（例如重複出現在表格中段）；
    # 表頭儲存格都很短，藉此避免把提到「辦理情形」的內文列誤判為表頭
    if any(len(t) > 15 for t in non_empty):
        return False
    joined = "".join(texts)
    hits = sum(1 for kw in ("序號", "案號", "裁", "承辦", "辦理情形", "建議") if kw in joined)
    return hits >= 3 and not CASE_NO_RE.search(joined)


def map_columns(header_rows: list[list[dict]], n_cols: int) -> dict[int, str] | None:
    if not header_rows:
        return None
    labels = []
    for c in range(n_cols):
        # 兩列式表頭（「列管」/「案號」）要上下串起來；跨欄儲存格去重
        pieces: list[str] = []
        for row in header_rows:
            if c < len(row):
                t = squash(row[c].get("text", ""))
                if t and (not pieces or pieces[-1] != t):
                    pieces.append(t)
        labels.append("".join(pieces))

    mapping: dict[int, str] = {}
    used: set[str] = set()
    for c, label in enumerate(labels):
        for key, kws in COLUMN_KEYWORDS:
            if key in used:
                continue
            if any(kw in label for kw in kws):
                mapping[c] = key
                used.add(key)
                break
    return mapping if "case_no" in used else None


def cell_text(row: list[dict], col: int, row_idx: int, is_key_col: bool,
              page_no: int = 0, pdf: "PdfLineReader | None" = None) -> str:
    if col < 0 or col >= len(row):
        return ""
    cell = row[col]
    # 跨列合併的儲存格會在 grid 每一列重複出現；非起始列不重複收錄文字
    start = cell.get("start_row_offset_idx")
    if start is not None and start != row_idx and not is_key_col:
        return ""
    text = cell.get("text", "") or ""
    if pdf is not None and not is_key_col:
        return pdf.cell_text(page_no, cell, text)
    return text


class PdfLineReader:
    """docling 會把儲存格內所有換行都變成空白，無法分辨「自動折行」與「真正換行」。
    這裡回到原始 PDF，依儲存格座標取出每一行文字：
    行尾明顯沒寫到儲存格右緣的行，就是段落結束，該處保留換行。"""

    def __init__(self, pdf_path: Path):
        import pypdfium2 as pdfium  # docling 已一併安裝
        self.path = pdf_path
        self.pdf = pdfium.PdfDocument(str(pdf_path))
        self._pages: dict[int, tuple[float, list[tuple[str, float, float, float, float]]]] = {}
        self.used = self.fallback = 0

    def close(self) -> None:
        self.pdf.close()

    def _chars(self, page_no: int):
        if page_no not in self._pages:
            page = self.pdf[page_no - 1]
            tp = page.get_textpage()
            chars = []
            for i in range(tp.count_chars()):
                ch = tp.get_text_range(i, 1)
                if not ch or ch.isspace():
                    continue
                left, bottom, right, top = tp.get_charbox(i, loose=True)
                if right > left and top > bottom:
                    chars.append((ch, left, bottom, right, top))
            self._pages[page_no] = (page.get_height(), chars)
        return self._pages[page_no]

    def cell_text(self, page_no: int, cell: dict, fallback: str) -> str:
        bbox = cell.get("bbox")
        if not bbox or not page_no or not squash(fallback):
            return fallback
        try:
            height, chars = self._chars(page_no)
        except Exception as exc:
            log.debug("讀取 %s 第 %s 頁失敗：%s", self.path.name, page_no, exc)
            self.fallback += 1
            return fallback

        if str(bbox.get("coord_origin", "TOPLEFT")).upper() == "TOPLEFT":
            top, bottom = height - bbox["t"], height - bbox["b"]
        else:
            top, bottom = bbox["t"], bbox["b"]
        left, right, pad = bbox["l"], bbox["r"], 1.0
        inside = [c for c in chars
                  if left - pad <= (c[1] + c[3]) / 2 <= right + pad
                  and bottom - pad <= (c[2] + c[4]) / 2 <= top + pad]
        if not inside:
            self.fallback += 1
            return fallback

        lines = [(ln.text, ln.r, ln.h) for ln in group_lines(inside)]

        # 與 docling 文字差異過大（例如掃描檔、座標偏移）就不採用；勾選框等符號不列入比較
        pdf_plain = "".join(t for t, _, _ in lines)
        if difflib.SequenceMatcher(None, squash_text(pdf_plain), squash_text(fallback)).ratio() < 0.9:
            self.fallback += 1
            return fallback
        self.used += 1
        return lines_to_text(lines, max(r for _, r, _ in lines))

    def page_size(self, page_no: int) -> tuple[float, float]:
        page = self.pdf[page_no - 1]
        return page.get_width(), page.get_height()

    def _darkness(self, page_no: int, x0: float, y0: float, x1: float, y1: float) -> float:
        """區域（左上原點、單位 pt）的平均深度 0～1；用於勾選框是向量圖形、不在文字層的 PDF。"""
        scale = 3
        if not hasattr(self, "_renders"):
            self._renders: dict[int, Any] = {}
        if page_no not in self._renders:
            self._renders[page_no] = self.pdf[page_no - 1].render(scale=scale, grayscale=True).to_pil().convert("L")
        img = self._renders[page_no]
        box = tuple(int(round(v * scale)) for v in (x0, y0, x1, y1))
        if box[2] <= box[0] or box[3] <= box[1]:
            return 0.0
        region = img.crop(box)
        hist = region.histogram()
        total = sum(hist) or 1
        return sum((255 - v) * n for v, n in enumerate(hist)) / total / 255

    def suggestion_text(self, page_no: int, lines: list["Line"]) -> str:
        """列管建議欄：文字層有勾選符號就直接用；沒有的話量測選項左側方框的深淺。"""
        text = "\n".join(ln.text for ln in lines)
        found = list(_OPTION_RE.finditer(squash(text)))
        if not found or any(m.group(1) and not re.match(r"[\w（）()、，,]", m.group(1)) for m in found):
            return text
        # 文字層沒有方框：每個選項各在一行時，量測該行文字左側 1.2 個字寬的區域
        option_lines = [ln for ln in lines if _OPTION_RE.search(squash(ln.text))]
        if len(option_lines) < 2:
            return text
        try:
            dark = [self._darkness(page_no, ln.l - ln.h * 1.3, ln.t, ln.l - ln.h * 0.1, ln.b)
                    for ln in option_lines]
        except Exception as exc:
            log.debug("勾選框影像判斷失敗：%s", exc)
            return text
        best = max(dark)
        others = sorted(dark)[-2] if len(dark) > 1 else 0
        out = []
        for ln, d in zip(option_lines, dark):
            mark = "■" if d == best and best > 0.25 and best > others * 1.5 else "□"
            out.append(mark + squash(ln.text))
        return "\n".join(out)

    def lines_in(self, page_no: int, x0: float, x1: float, y0: float, y1: float) -> list["Line"]:
        """取出區域內的文字行；座標為左上角原點（與 docling TOPLEFT 相同），y 向下遞增。"""
        height, chars = self._chars(page_no)
        inside = [c for c in chars
                  if x0 <= (c[1] + c[3]) / 2 < x1 and y0 <= height - (c[2] + c[4]) / 2 < y1]
        lines = group_lines(inside)
        for ln in lines:
            ln.t, ln.b = height - ln.t, height - ln.b  # 轉成左上角原點
        return lines


@dataclass
class Line:
    text: str
    l: float
    r: float
    t: float  # PDF 原始座標（左下原點）時 t > b；轉換後 t < b
    b: float
    h: float


def group_lines(chars: list[tuple[str, float, float, float, float]]) -> list[Line]:
    """把字元（字, left, bottom, right, top；左下原點）依 y 座標分行，由上而下、行內由左而右。"""
    chars = sorted(chars, key=lambda c: (-(c[2] + c[4]) / 2, c[1]))
    groups: list[list[tuple]] = []
    last_cy = None
    for c in chars:
        cy, h = (c[2] + c[4]) / 2, c[4] - c[2]
        if groups and abs(last_cy - cy) < h * 0.5:
            groups[-1].append(c)
        else:
            groups.append([c])
            last_cy = cy

    lines = []
    for g in groups:
        g.sort(key=lambda c: c[1])
        h = statistics.median(c[4] - c[2] for c in g)
        text = g[0][0]
        for prev, cur in zip(g, g[1:]):
            # PDF 中英文單字間的空白常沒有字元框，用字距補回
            if cur[1] - prev[3] > h * 0.25 and prev[0].isascii() and cur[0].isascii():
                text += " "
            text += cur[0]
        lines.append(Line(text, min(c[1] for c in g), max(c[3] for c in g),
                          max(c[4] for c in g), min(c[2] for c in g), h))
    return lines


def lines_to_text(lines: list[tuple[str, float, float]], line_right: float) -> str:
    """(文字, 行尾 x, 字高) → 文字；行尾明顯沒寫到右緣處視為段落結束。"""
    out = ""
    for i, (text, r, h) in enumerate(lines):
        out += text
        ends_para = line_right - r > h * 1.5
        if i < len(lines) - 1:
            nxt = lines[i + 1][0]
            if ends_para or re.match(BULLET, nxt) or re.match(HEADING, nxt):
                out += "\n"
            elif re.search(r"[A-Za-z0-9]$", text) and re.match(r"[A-Za-z0-9]", nxt):
                out += " "
        elif ends_para and len(lines) > 1:
            out += "\n"  # 最後一行沒寫滿 → 段落已結束，跨頁接續時不與下一頁黏在一起
    return out


def squash_text(text: str) -> str:
    # 只保留文字與數字，用於比對內容是否一致
    return re.sub(r"[^\w]", "", text or "")


def open_pdf_reader(pdf_path: Path | None, source: str) -> PdfLineReader | None:
    if pdf_path is None:
        log.warning("找不到原始 PDF「%s」，換行改以文字規則判斷（較不精確）", source)
        return None
    try:
        return PdfLineReader(pdf_path)
    except ImportError:
        log.warning("缺少 pypdfium2（python -m pip install pypdfium2），換行改以文字規則判斷")
    except Exception as exc:
        log.warning("無法開啟 %s：%s，換行改以文字規則判斷", pdf_path.name, exc)
    return None


def normalize_case_no(text: str) -> str:
    m = CASE_NO_RE.search(squash(text).replace("－", "-").replace("–", "-"))
    return re.sub(r"\s+", "", m.group(0)) if m else ""


def parse_document(doc: dict, source_name: str, pdf: PdfLineReader | None = None) -> list[Record]:
    """有原始 PDF 時，以「列管案號在頁面上的位置」切分案件（不受 docling 表格列切法影響）；
    結果與 docling 表格差異過大時，退回依 docling 表格列切分。"""
    if pdf is None:
        return parse_document_grid(doc, source_name)
    # docling 表格列切分的結果只拿來核對；確定採用時才重跑一次以顯示警告
    grid_records = parse_document_grid(doc, source_name, pdf, quiet=True)

    def fallback() -> list[Record]:
        return parse_document_grid(doc, source_name, pdf)

    try:
        geo_records = parse_document_geometric(doc, source_name, pdf)
    except Exception as exc:
        log.warning("%s 以版面位置切分失敗（%s），改用 docling 表格列", source_name, exc)
        return fallback()

    grid_cases = {r.case_no for r in grid_records if r.case_no}
    geo_cases = {r.case_no for r in geo_records if r.case_no}
    missing = sorted(grid_cases - geo_cases)
    if not geo_records or missing:
        log.warning("%s 版面位置切分漏掉案號 %s，改用 docling 表格列（請特別檢查案件分界）",
                    source_name, "、".join(missing) or "（全部）")
        return fallback()

    def all_text(recs: list[Record]) -> str:
        return squash_text("".join("".join(v) for r in recs for k, v in r.parts.items() if k != "suggestion"))

    ratio = difflib.SequenceMatcher(None, all_text(geo_records), all_text(grid_records), autojunk=False).ratio()
    if ratio < 0.85:
        log.warning("%s PDF 文字層與 docling 結果差異過大（相似度 %.0f%%），改用 docling 表格列",
                    source_name, ratio * 100)
        return fallback()
    extra = sorted(geo_cases - grid_cases)
    if extra:
        log.info("  版面位置另外找到 docling 漏掉的案號：%s", "、".join(extra))
    log.info("  以版面位置切分案件（依列管案號在頁面上的位置）")
    return geo_records


HEADER_WORDS = {
    "序", "號", "序號", "列管", "案號", "列管案號", "案號號", "列管案號號",
    "主席裁（指）示", "主席裁(指)示", "暨決議事項", "主席裁（指）示暨決議事項", "主席裁(指)示暨決議事項",
    "承辦", "單位", "承辦單位", "最新辦理情形", "辦理情形", "建議", "列管建議",
}
REQUIRED_KEYS = {"case_no", "directive", "unit", "status", "suggestion"}
CONTENT_KEYS = ("directive", "unit", "status", "suggestion")


def bbox_tl(bbox: dict, page_h: float) -> tuple[float, float, float, float]:
    """docling bbox → (l, t, r, b)，左上角原點。"""
    if str(bbox.get("coord_origin", "TOPLEFT")).upper() == "TOPLEFT":
        return bbox["l"], bbox["t"], bbox["r"], bbox["b"]
    return bbox["l"], page_h - bbox["t"], bbox["r"], page_h - bbox["b"]


def column_bounds(grid: list[list[dict]], mapping: dict[int, str], page_w: float,
                  page_h: float) -> dict[str, tuple[float, float]]:
    """依 docling 儲存格位置推算每一欄的左右邊界（相鄰兩欄取中點）。"""
    n_cols = max(len(r) for r in grid)
    ls: dict[int, list[float]] = {c: [] for c in range(n_cols)}
    rs: dict[int, list[float]] = {c: [] for c in range(n_cols)}
    for row in grid:
        for c, cell in enumerate(row):
            bb = cell.get("bbox")
            if not bb or (cell.get("col_span") or 1) != 1 or not squash(cell.get("text", "")):
                continue
            l, _, r, _ = bbox_tl(bb, page_h)
            ls[c].append(l)
            rs[c].append(r)
    cols = sorted((statistics.median(ls[c]), statistics.median(rs[c]), c) for c in range(n_cols) if ls[c])
    bounds: dict[str, tuple[float, float]] = {}
    for i, (l, r, c) in enumerate(cols):
        left = 0.0 if i == 0 else (cols[i - 1][1] + l) / 2
        right = page_w if i == len(cols) - 1 else (r + cols[i + 1][0]) / 2
        if c in mapping:
            bounds[mapping[c]] = (left, right)
    return bounds


def page_furniture_top(doc: dict, page_no: int, page_h: float) -> float | None:
    """頁尾（頁碼等）最上緣，用來排除頁碼被當成內容。"""
    tops = []
    for item in doc.get("texts", []):
        if item.get("label") not in ("page_footer",):
            continue
        for prov in item.get("prov") or []:
            if prov.get("page_no") == page_no and prov.get("bbox"):
                tops.append(bbox_tl(prov["bbox"], page_h)[1])
    return min(tops) if tops else None


def parse_document_geometric(doc: dict, source_name: str, pdf: PdfLineReader) -> list[Record]:
    records: list[Record] = []
    current: Record | None = None
    last_mapping: dict[int, str] | None = None
    last_bounds: dict[str, tuple[float, float]] | None = None

    # 同一頁可能被 docling 切成好幾張表：每頁只處理一次，欄位以格子最多的那張為準，上下範圍取聯集
    by_page: dict[int, list[dict]] = {}
    for table in tables_in_reading_order(doc):
        prov = (table.get("prov") or [{}])[0]
        if table_grid(table) and prov.get("page_no") and prov.get("bbox"):
            by_page.setdefault(prov["page_no"], []).append(table)

    for page_no in sorted(by_page):
        tables = by_page[page_no]
        table = max(tables, key=lambda t: sum(len(r) for r in table_grid(t)))
        grid = table_grid(table)
        prov = table["prov"][0]
        n_cols = max(len(r) for r in grid)
        page_w, page_h = pdf.page_size(page_no)

        header_idx = [i for i, row in enumerate(grid) if is_header_row(row)]
        leading = []
        for i in header_idx:
            if i == len(leading):
                leading.append(i)
        mapping = map_columns([grid[i] for i in leading], n_cols)
        if mapping is None:
            if last_mapping and max(last_mapping) < n_cols:
                mapping = last_mapping
            elif n_cols == len(DEFAULT_COLUMN_ORDER):
                mapping = dict(enumerate(DEFAULT_COLUMN_ORDER))
            else:
                continue
        last_mapping = mapping

        bounds = column_bounds(grid, mapping, page_w, page_h)
        # docling 偶爾少抓一欄（例如列管建議併進辦理情形）：沿用前一頁完整的欄位邊界
        if not REQUIRED_KEYS <= bounds.keys():
            if last_bounds is None:
                raise RuntimeError(f"第 {page_no} 頁欄位不完整")
            bounds = last_bounds
        else:
            last_bounds = bounds

        # 表格內容的上下範圍：表頭下緣 ～ 表格下緣（不含頁尾）
        _, tbl_t, _, tbl_b = bbox_tl(prov["bbox"], page_h)
        for other in tables:
            _, ot, _, ob = bbox_tl(other["prov"][0]["bbox"], page_h)
            tbl_t, tbl_b = min(tbl_t, ot), max(tbl_b, ob)
        body_top = tbl_t
        for i in leading:
            for cell in grid[i]:
                # 只採用短字串的表頭格，避免 docling 把第一列內容誤標成表頭時吃掉內容
                if cell.get("bbox") and 0 < len(squash(cell.get("text", ""))) <= 12:
                    body_top = max(body_top, bbox_tl(cell["bbox"], page_h)[3])
        body_bottom = tbl_b + 3
        footer_top = page_furniture_top(doc, page_no, page_h)
        if footer_top is not None and footer_top > tbl_b:
            body_bottom = min(max(body_bottom, tbl_b), footer_top)

        def col_lines(key: str) -> list[Line]:
            if key not in bounds:
                return []
            x0, x1 = bounds[key]
            out = []
            for ln in pdf.lines_in(page_no, x0, x1, body_top - 1, body_bottom):
                s = squash(ln.text)
                if not s or s in HEADER_WORDS:
                    continue
                if s.isdigit() and len(s) <= 3 and key != "seq" and body_bottom - ln.b < ln.h * 2:
                    continue  # 頁碼
                out.append(ln)
            return out

        # 1. 找出本頁每一案的起始位置：列管案號（必要時合併被折成兩行的案號），以及序號
        starts: list[dict] = []
        case_lines = col_lines("case_no")
        i = 0
        while i < len(case_lines):
            ln = case_lines[i]
            text = squash(ln.text)
            if re.search(r"\d{5,8}[-－–]$", text) and i + 1 < len(case_lines):
                text += squash(case_lines[i + 1].text)
                i += 1
            cn = normalize_case_no(text)
            if cn:
                starts.append({"t": ln.t, "h": ln.h, "case_no": cn, "seq": ""})
            i += 1
        for ln in col_lines("seq"):
            s = squash(ln.text)
            if not re.fullmatch(r"\d{1,3}", s):
                continue
            near = [st for st in starts if abs(st["t"] - ln.t) < max(st["h"], ln.h) * 1.2]
            if near:
                near[0]["seq"] = s
            else:
                log.warning("%s 第 %s 頁序號 %s 找不到列管案號，已建立空白案號的資料列",
                            source_name, page_no, s)
                starts.append({"t": ln.t, "h": ln.h, "case_no": "", "seq": s})
        starts.sort(key=lambda st: st["t"])

        # 2. 建立分段：起始位置以上屬於上一頁延續的案件
        segments: list[tuple[float, Record | None]] = [(float("-inf"), current)]
        for st in starts:
            if st["case_no"] and current is not None and st["case_no"] == current.case_no:
                continue  # 續頁重複列出同一案號
            current = Record(case_no=st["case_no"], seq=st["seq"], source_pdf=source_name)
            records.append(current)
            segments.append((st["t"] - st["h"] * 0.5, current))

        # 3. 每一欄的每一行，依垂直位置歸到所屬案件
        for key in CONTENT_KEYS:
            lines = col_lines(key)
            if not lines:
                continue
            right_limit = max(ln.r for ln in lines)
            buckets: dict[int, list[Line]] = {}
            for ln in lines:
                idx = max(j for j, (top, _) in enumerate(segments) if top <= ln.t)
                buckets.setdefault(idx, []).append(ln)
            for idx, lns in buckets.items():
                rec = segments[idx][1]
                if key == "suggestion":
                    text = pdf.suggestion_text(page_no, lns)
                else:
                    text = lines_to_text([(ln.text, ln.r, ln.h) for ln in lns], right_limit)
                if rec is None:
                    log.warning("%s 第 %s 頁表格開頭有不屬於任何案件的內容，已略過：%s",
                                source_name, page_no, squash(text)[:30])
                    continue
                rec.add(key, text)
                if page_no not in rec.pages:
                    rec.pages.append(page_no)
        for _, rec in segments[1:]:
            if rec is not None and page_no not in rec.pages:
                rec.pages.append(page_no)

    for rec in records:
        rec.pages.sort()
    return records


def parse_document_grid(doc: dict, source_name: str, pdf: PdfLineReader | None = None,
                        quiet: bool = False) -> list[Record]:
    warn = (lambda *a, **k: None) if quiet else log.warning
    records: list[Record] = []
    current: Record | None = None
    last_mapping: dict[int, str] | None = None

    for table in tables_in_reading_order(doc):
        grid = table_grid(table)
        if not grid:
            continue
        n_cols = max(len(r) for r in grid)
        page_no = ((table.get("prov") or [{}])[0]).get("page_no", 0)

        header_idx = [i for i, row in enumerate(grid) if is_header_row(row)]
        # 只把表格開頭連續的表頭列拿來判斷欄位（中段重複表頭僅略過）
        leading = []
        for i in header_idx:
            if i == len(leading):
                leading.append(i)
        mapping = map_columns([grid[i] for i in leading], n_cols)
        if mapping is None:
            # 續頁表格沒有表頭：沿用上一張表，或採預設欄序
            if last_mapping and max(last_mapping) < n_cols:
                mapping = last_mapping
            elif n_cols == len(DEFAULT_COLUMN_ORDER):
                mapping = dict(enumerate(DEFAULT_COLUMN_ORDER))
            else:
                warn("%s 第 %s 頁表格無法辨識欄位，已略過", source_name, page_no)
                continue
        last_mapping = mapping
        key_to_col = {v: k for k, v in mapping.items()}

        for r_idx, row in enumerate(grid):
            if r_idx in header_idx:
                continue
            raw_case = cell_text(row, key_to_col["case_no"], r_idx, True)
            case_no = normalize_case_no(raw_case)
            seq = squash(cell_text(row, key_to_col.get("seq", -1), r_idx, True))

            values = {
                key: cell_text(row, col, r_idx, False, page_no, pdf)
                for col, key in mapping.items()
                if key in ("directive", "unit", "status", "suggestion")
            }
            if not case_no and not seq and not any(squash(v) for v in values.values()):
                continue  # 空白列

            # 判斷是否為跨頁 / 跨列的延續：案號空白，或與目前案件相同
            is_continuation = current is not None and (not case_no or case_no == current.case_no)
            # 案號沒抓到、但序號換了 → 視為新案件，而不是併入上一筆
            if is_continuation and not case_no and seq and seq != current.seq:
                warn("%s 第 %s 頁序號 %s 找不到列管案號，已建立空白案號的資料列",
                            source_name, page_no, seq)
                current = Record(case_no="", seq=seq, source_pdf=source_name)
                records.append(current)
                is_continuation = True
            if not is_continuation:
                if not case_no:
                    warn("%s 第 %s 頁有一列找不到列管案號，內容：%s",
                                source_name, page_no, squash(" ".join(values.values()))[:40])
                    continue
                current = Record(case_no=case_no, seq=seq, source_pdf=source_name)
                records.append(current)

            for key, text in values.items():
                current.add(key, text)
            if page_no and page_no not in current.pages:
                current.pages.append(page_no)

    return records


# ─── 文字清理：分段、半形標點轉全形、中英數之間加空格 ───

HAN = r"㐀-䶿一-鿿豈-﫿"
# 條列起頭：一、 （一） (一) 1、 1. （1） ①
BULLET = (
    r"(?:[一二三四五六七八九十]{1,3}、"
    r"|[（(][一二三四五六七八九十]{1,3}[）)]"
    r"|\d{1,2}[、．](?!\d)"
    r"|\d{1,2}\.(?!\d)"
    r"|[（(]\d{1,2}[）)]"
    r"|[①-⑳])"
)
# 「115年4月9日列管事項」這類小標題，前後都要換行
HEADING = (
    r"\d{2,3}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日[^\s，,。；;]{0,15}?"
    r"(?:列管事項|決議事項|裁示事項|指示事項)"
)
SENTENCE_END = r"[。；;：:！!？?」』）)]"

PUNCT_MAP = {",": "，", ";": "；", ":": "：", "?": "？", "!": "！",
             "(": "（", ")": "）", "[": "［", "]": "］", "~": "～"}


def to_fullwidth(text: str) -> str:
    # 網址不轉換
    parts = re.split(r"(https?://\S+)", text)
    for i in range(0, len(parts), 2):
        parts[i] = _to_fullwidth(parts[i])
    return "".join(parts)


def _to_fullwidth(s: str) -> str:
    out = []
    quote_open = True
    for i, ch in enumerate(s):
        prev = s[i - 1] if i else ""
        nxt = s[i + 1] if i + 1 < len(s) else ""
        if ch in ",:" and prev.isdigit() and nxt.isdigit():
            out.append(ch)  # 千分位 2,383、時間 10:30 不轉
        elif ch == ".":
            if prev.isdigit() and nxt.isdigit():
                out.append(ch)  # 小數 83.5
            elif prev and re.match(rf"[{HAN}）」』]", prev):
                out.append("。")
            elif prev.isdigit() and (not nxt or re.match(rf"[{HAN}\s]", nxt)):
                out.append("．")  # 條列「1.」
            else:
                out.append(ch)  # 英文縮寫等
        elif ch == '"':
            out.append("「" if quote_open else "」")
            quote_open = not quote_open
        else:
            out.append(PUNCT_MAP.get(ch, ch))
    return "".join(out)


def add_spacing(s: str) -> str:
    # 中文字與英數字之間加空格；緊鄰標點符號者不加
    s = re.sub(rf"([{HAN}])(?=[A-Za-z0-9])", r"\1 ", s)
    s = re.sub(rf"([A-Za-z0-9%])(?=[{HAN}])", r"\1 ", s)
    return s


def clean_paragraphs(text: str) -> str:
    if not text:
        return ""
    t = text.replace("\r\n", "\n").replace("\r", "\n").replace("　", " ")
    t = re.sub(r"[ \t\f\v]+", " ", t).strip()

    # 在條列符號前分行：前面是空白或句末標點時才算新段落
    t = re.sub(rf"(?:(?<={SENTENCE_END})|\s+)\s*(?={BULLET})", "\n", t)
    # 小標題獨立成行（位於行首或句末標點之後）
    t = re.sub(rf"(?:^|(?<={SENTENCE_END}))[ ]*({HEADING})[ ]*", r"\n\1\n", t, flags=re.M)

    lines = []
    for line in t.split("\n"):
        # PDF 折行留下的空白一律移除，只保留英數字與英數字之間的空白，再重新加空格
        line = re.sub(r"(?<![A-Za-z0-9]) +| +(?![A-Za-z0-9])", "", line.strip())
        line = add_spacing(to_fullwidth(line))
        if line:
            lines.append(line)
    return "\n".join(lines)


def join_fragments(fragments: list[str]) -> str:
    # 跨頁片段：前一段以換行結尾就直接接上；否則以空白相接，交給 clean_paragraphs 接續
    buf = ""
    for f in fragments:
        if not f or not f.strip():
            continue
        if buf and not buf.endswith("\n"):
            buf += " "
        buf += f
    return clean_paragraphs(buf)


# 未勾選的方框；其他出現在選項前的符號（█ ▓ ■ ☑ ✔ ● 等）一律視為已勾選
UNCHECKED = "□☐◻◽▢○◯⬜"
SUGGESTION_OPTIONS = ("解除列管", "繼續列管", "自行列管", "除管", "續管")
_OPTION_RE = re.compile(rf"(.?)({'|'.join(SUGGESTION_OPTIONS)})")


def parse_suggestion(text: str) -> str:
    """「█除管 □續管」「▓除管 □續管」→「除管」；無法判斷時保留原文。"""
    t = squash(text)
    if not t:
        return ""
    found = list(_OPTION_RE.finditer(t))
    if not found:
        return clean_paragraphs(text)
    checked = []
    for m in found:
        mark = m.group(1)
        # 前一個字是文字（例如上一個選項的「管」）代表沒有勾選框
        if mark and mark not in UNCHECKED and not re.match(r"[\w（）()、，,]", mark):
            checked.append(m.group(2))
    if checked:
        return "\n".join(dict.fromkeys(checked))
    if all(m.group(1) in UNCHECKED for m in found if m.group(1)) and any(m.group(1) for m in found):
        return ""  # 全部未勾選
    log.warning("無法判斷列管建議的勾選狀態：%s", t)
    return clean_paragraphs(text)


def to_rows(records: list[Record]) -> list[list[Any]]:
    rows = []
    for rec in records:
        case_value: Any = int(rec.case_no) if rec.case_no.isdigit() else rec.case_no
        row = [
            case_value,
            join_fragments(rec.parts["directive"]),
            join_fragments(rec.parts["unit"]).replace("\n", ""),
            join_fragments(rec.parts["status"]),
            parse_suggestion(" ".join(rec.parts["suggestion"])),
            rec.source_pdf,
        ]
        for i, v in enumerate(row):
            if isinstance(v, str) and len(v) > CELL_CHAR_LIMIT:
                log.warning("案號 %s 第 %d 欄超過 %d 字，已截斷", rec.case_no, i + 1, CELL_CHAR_LIMIT)
                row[i] = v[: CELL_CHAR_LIMIT - 20] + "\n…（內容過長已截斷）"
        rows.append(row)
    return rows


# ──────────────────────────────────────────────────────────────
# 三、送到 Google Apps Script 網頁應用程式寫入試算表
# ──────────────────────────────────────────────────────────────

BATCH_SIZE = 100


def ssl_context() -> ssl.SSLContext:
    # python.org 版 Python 在 macOS 上常缺根憑證；有裝 certifi 就用它
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


def post_json(url: str, payload: dict | None, retries: int = 3) -> dict:
    """payload 為 None 時送 GET（用來測試部署是否成功）。"""
    data = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
    ctx = ssl_context()
    for attempt in range(1, retries + 1):
        # Apps Script 會先回 302 轉址到 googleusercontent.com，urllib 會自動改用 GET 取回結果
        req = urllib.request.Request(
            url, data=data, method="GET" if data is None else "POST",
            headers={"Content-Type": "application/json; charset=utf-8"},
        )
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=120) as resp:
                text = resp.read().decode("utf-8")
            try:
                return json.loads(text)
            except json.JSONDecodeError:
                # 拿到 HTML 通常是部署權限不是「任何人」，或網址不是 /exec 結尾
                raise RuntimeError("網頁應用程式沒有回傳 JSON，請確認部署設定為「任何人」可存取、"
                                   "網址以 /exec 結尾。回應開頭：" + text[:200])
        except (urllib.error.URLError, TimeoutError) as exc:
            reason = getattr(exc, "reason", exc)
            if isinstance(reason, ssl.SSLCertVerificationError) or isinstance(exc, ssl.SSLCertVerificationError):
                raise RuntimeError("SSL 憑證驗證失敗：請執行 python -m pip install certifi；"
                                   "python.org 版 Python 也可執行「Install Certificates.command」")
            if isinstance(exc, urllib.error.HTTPError) and exc.code == 404:
                raise RuntimeError("找不到網頁應用程式（404），請確認網址正確且已部署")
            if attempt == retries:
                raise RuntimeError(f"連線失敗：{exc}")
            log.warning("連線失敗（第 %d 次），%d 秒後重試：%s", attempt, attempt * 3, exc)
            time.sleep(attempt * 3)
    raise RuntimeError("unreachable")


def upload(rows: list[list[Any]], webapp: str, token: str, worksheet: str, mode: str) -> None:
    if mode == "replace" and not rows:
        log.warning("沒有資料，replace 模式不執行以免清空工作表")
        return
    total_updated = total_appended = 0
    for start in range(0, len(rows), BATCH_SIZE):
        batch = rows[start:start + BATCH_SIZE]
        # replace 只在第一批清空，後續批次改為新增
        batch_mode = mode if (mode != "replace" or start == 0) else "append"
        result = post_json(webapp, {
            "token": token, "sheet": worksheet, "mode": batch_mode,
            "rows": [[("" if v is None else v) for v in r] for r in batch],
        })
        if not result.get("ok"):
            raise RuntimeError(f"Apps Script 回報錯誤：{result.get('error')}")
        total_updated += result.get("updated", 0)
        total_appended += result.get("appended", 0)
        log.info("已送出 %d / %d 列", min(start + BATCH_SIZE, len(rows)), len(rows))
    log.info("試算表已更新 %d 列、新增 %d 列", total_updated, total_appended)



# ──────────────────────────────────────────────────────────────
# 設定檔與 init 指令
# ──────────────────────────────────────────────────────────────

def config_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "pdf2sheet" / "config.json"


def load_config() -> dict:
    path = config_path()
    if not path.exists():
        return {}
    try:
        with path.open(encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        log.warning("設定檔無法讀取（%s）：%s，請重新執行 pdf2sheet init", path, exc)
        return {}


def save_config(cfg: dict) -> Path:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    if os.name != "nt":
        os.chmod(path, 0o600)  # 只有自己能讀，避免密碼外洩
    return path


def run_init(argv: list[str]) -> int:
    import getpass

    ap = argparse.ArgumentParser(prog="pdf2sheet init", description="設定 Apps Script 網址與密碼")
    ap.add_argument("--show", action="store_true", help="顯示目前設定（密碼遮蔽）")
    ap.add_argument("--no-test", action="store_true", help="不測試連線，直接存檔")
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    cfg = load_config()
    if args.show:
        print(f"設定檔：{config_path()}")
        if not cfg:
            print("（尚未設定，請執行 pdf2sheet init）")
            return 0
        tok = cfg.get("token", "")
        print(f"網址：{cfg.get('webapp', '')}")
        print(f"密碼：{tok[:2] + '*' * max(len(tok) - 2, 0) if tok else '（未設定）'}")
        print(f"工作表：{cfg.get('worksheet', '列管表')}")
        return 0

    print("設定 pdf2sheet（直接按 Enter 保留目前的值）\n")
    try:
        while True:
            cur = cfg.get("webapp", "")
            url = input(f"Apps Script 網頁應用程式網址{f' [{cur}]' if cur else ''}：").strip() or cur
            if not url:
                print("  請輸入網址（https://script.google.com/macros/s/…/exec）")
                continue
            if url.endswith("/dev"):
                print("  這是測試用的 /dev 網址，請改用「部署」後取得的 /exec 網址")
                continue
            if not url.startswith("https://script.google.com/"):
                print("  網址應以 https://script.google.com/ 開頭")
                continue
            break

        while True:
            cur = cfg.get("token", "")
            tok = getpass.getpass(f"密碼（與 Code.gs 的 TOKEN 相同，輸入時不顯示）{' [已設定]' if cur else ''}：").strip() or cur
            if tok:
                break
            print("  請輸入密碼")

        cur = cfg.get("worksheet", "列管表")
        ws = input(f"工作表名稱 [{cur}]：").strip() or cur
    except (KeyboardInterrupt, EOFError):
        print("\n已取消，未變更設定")
        return 1

    if not args.no_test:
        print("\n測試連線…")
        try:
            if not post_json(url, None).get("ok"):
                raise RuntimeError("網頁應用程式回應異常")
            # 送出 0 筆資料驗證密碼；工作表不存在時會順便建立並加上表頭
            result = post_json(url, {"token": tok, "sheet": ws, "mode": "upsert", "rows": []})
            if not result.get("ok"):
                raise RuntimeError(f"Apps Script 回報錯誤：{result.get('error')}")
        except RuntimeError as exc:
            print(f"✗ {exc}")
            print("設定未儲存。確認後請重新執行 pdf2sheet init（或加 --no-test 略過測試）")
            return 1
        print(f"✓ 連線成功，已確認工作表「{ws}」")

    path = save_config({"webapp": url, "token": tok, "worksheet": ws})
    print(f"\n已儲存設定：{path}")
    print("之後只要執行：pdf2sheet <PDF 資料夾>")
    return 0


# ──────────────────────────────────────────────────────────────

def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if argv and argv[0] == "init":
        return run_init(argv[1:])

    ap = argparse.ArgumentParser(
        prog="pdf2sheet",
        description="會議資料 PDF 中的「主席裁（指）示暨決議事項執行情形分辦表」→ Google 試算表",
        epilog="首次使用請先執行：pdf2sheet init",
    )
    ap.add_argument("input_dir", type=Path, help="放 PDF 的資料夾")
    ap.add_argument("--pages", default="auto",
                    help="要解析的頁碼：auto（預設，自動找分辦表）、all（整份）、或如 11-17,20")
    ap.add_argument("--json-dir", type=Path, help="docling JSON 輸出資料夾（預設：<input_dir>/_docling_json）")
    ap.add_argument("--webapp", help="Apps Script 網址（預設讀 GAS_WEBAPP_URL 或 init 設定）")
    ap.add_argument("--token", help="Apps Script 密碼（預設讀 GAS_TOKEN 或 init 設定）")
    ap.add_argument("--worksheet", help="工作表名稱（預設讀 init 設定，否則為「列管表」）")
    ap.add_argument("--mode", choices=["upsert", "append", "replace"], default="upsert",
                    help="upsert：同 PDF 同案號覆寫，其餘新增（預設）；append：全部新增；replace：清空後重寫")
    ap.add_argument("--recursive", action="store_true", help="一併處理子資料夾內的 PDF")
    ap.add_argument("--ocr", action="store_true", help="掃描型 PDF 啟用 OCR（macOS 使用內建 Vision）")
    ap.add_argument("--force", action="store_true", help="即使已有 JSON 也重新解析 PDF")
    ap.add_argument("--from-json", action="store_true", help="跳過步驟一，直接讀 --json-dir 內現有的 JSON")
    ap.add_argument("--dry-run", action="store_true", help="不上傳，只輸出 structured.csv / structured.json 供檢查")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")

    # 優先順序：指令參數 > 環境變數 > init 設定檔
    cfg = load_config()
    webapp = args.webapp or os.environ.get("GAS_WEBAPP_URL") or cfg.get("webapp")
    token = args.token or os.environ.get("GAS_TOKEN") or cfg.get("token")
    worksheet = args.worksheet or cfg.get("worksheet") or "列管表"

    input_dir: Path = args.input_dir.expanduser().resolve()
    if not input_dir.is_dir():
        log.error("找不到資料夾：%s", input_dir)
        return 1
    json_dir = (args.json_dir or input_dir / "_docling_json").expanduser().resolve()

    # 步驟一
    if args.from_json:
        json_paths = sorted(p for p in json_dir.glob("*.json") if p.name != "structured.json")
    else:
        pattern = "**/*" if args.recursive else "*"
        pdfs = sorted(p for p in input_dir.glob(pattern)
                      if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith(".")
                      and json_dir not in p.parents)
        if not pdfs:
            log.error("資料夾內沒有 PDF：%s", input_dir)
            return 1
        log.info("共 %d 份 PDF", len(pdfs))
        json_paths = convert_pdfs(pdfs, json_dir, args.ocr, args.force, args.pages)

    # 步驟二
    pattern = "**/*" if args.recursive else "*"
    pdf_index = {p.name: p for p in input_dir.glob(pattern)
                 if p.is_file() and p.suffix.lower() == ".pdf"}
    all_records: list[Record] = []
    for jp in json_paths:
        with jp.open(encoding="utf-8") as f:
            doc = json.load(f)
        source = (doc.get("origin") or {}).get("filename") or f"{jp.stem}.pdf"
        pdf = open_pdf_reader(pdf_index.get(source) or pdf_index.get(f"{jp.stem}.pdf"), source)
        try:
            recs = parse_document(doc, source, pdf)
        finally:
            if pdf:
                pdf.close()
        log.info("%s：%d 筆案件", source, len(recs))
        for r in recs:
            if len(r.pages) > 1:
                log.info("  案號 %s 跨頁合併（第 %s 頁）", r.case_no, "、".join(map(str, r.pages)))
        all_records.extend(recs)

    rows = to_rows(all_records)

    out_csv = json_dir / "structured.csv"
    with out_csv.open("w", encoding="utf-8-sig", newline="") as f:  # utf-8-sig 讓 Excel/Numbers 正確顯示中文
        w = csv.writer(f)
        w.writerow(SHEET_HEADER)
        w.writerows(rows)
    with (json_dir / "structured.json").open("w", encoding="utf-8") as f:
        json.dump([dict(zip(SHEET_HEADER, r)) | {"_pages": rec.pages}
                   for r, rec in zip(rows, all_records)], f, ensure_ascii=False, indent=2)
    log.info("結構化結果已存：%s", out_csv)

    # 步驟三
    if args.dry_run:
        log.info("dry-run：未上傳 Google 試算表")
        return 0
    if not webapp or not token:
        log.error("尚未設定 Apps Script 網址與密碼：請先執行 pdf2sheet init（或加 --dry-run 只檢查）")
        return 1
    try:
        upload(rows, webapp, token, worksheet, args.mode)
    except RuntimeError as exc:
        log.error("%s", exc)
        return 1
    log.info("完成")
    return 0


if __name__ == "__main__":
    sys.exit(main())
