#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF（會議主席裁示列管表）→ docling JSON → 結構化 → Google 試算表

步驟：
  一、以 docling 將資料夾內每份 PDF 解析為 JSON（存於 --json-dir）
  二、讀取 JSON，合併跨頁斷開的表格列，整理成結構化資料
  三、依「列管案號｜主席裁示｜承辦單位｜最新辦理情形｜列管建議｜來源 PDF 名稱」
      送到 Google Apps Script 網頁應用程式（Code.gs），由它寫入試算表

用法：
  python3 pdf_to_sheet.py ./pdfs \
      --webapp "https://script.google.com/macros/s/XXXX/exec" \
      --token "change-me" \
      --worksheet "列管表"

  # 只想先檢查解析結果、不上傳：
  python3 pdf_to_sheet.py ./pdfs --dry-run

安裝：
  python3 -m pip install docling
  （掃描檔需 OCR 時另裝：python3 -m pip install ocrmac）
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


def convert_pdfs(pdf_paths: list[Path], json_dir: Path, use_ocr: bool, force: bool) -> list[Path]:
    json_dir.mkdir(parents=True, exist_ok=True)
    converter = None
    out: list[Path] = []
    for pdf in pdf_paths:
        target = json_dir / f"{pdf.stem}.json"
        if not force and target.exists() and target.stat().st_mtime >= pdf.stat().st_mtime:
            log.info("略過轉檔（已有較新的 JSON）：%s", target.name)
            out.append(target)
            continue
        if converter is None:
            log.info("初始化 docling（首次執行會下載模型，請稍候）…")
            converter = build_converter(use_ocr)
        log.info("docling 解析中：%s", pdf.name)
        try:
            result = converter.convert(str(pdf))
        except Exception as exc:  # 單檔失敗不影響其他檔案
            log.error("解析失敗 %s：%s", pdf.name, exc)
            continue
        with target.open("w", encoding="utf-8") as f:
            json.dump(result.document.export_to_dict(), f, ensure_ascii=False, indent=2)
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

        # 依 y 座標分行（由上而下），行內依 x 排序
        inside.sort(key=lambda c: (-(c[2] + c[4]) / 2, c[1]))
        groups: list[list[tuple]] = []
        last_cy = None
        for c in inside:
            cy, h = (c[2] + c[4]) / 2, c[4] - c[2]
            if groups and abs(last_cy - cy) < h * 0.5:
                groups[-1].append(c)
            else:
                groups.append([c])
                last_cy = cy

        lines = []  # (文字, 行尾 x, 字高)
        for g in groups:
            g.sort(key=lambda c: c[1])
            h = statistics.median(c[4] - c[2] for c in g)
            text = g[0][0]
            for prev, cur in zip(g, g[1:]):
                # PDF 中英文單字間的空白常沒有字元框，用字距補回
                if cur[1] - prev[3] > h * 0.25 and prev[0].isascii() and cur[0].isascii():
                    text += " "
                text += cur[0]
            lines.append((text, max(c[3] for c in g), h))

        # 與 docling 文字差異過大（例如掃描檔、座標偏移）就不採用
        pdf_plain = "".join(t for t, _, _ in lines)
        if difflib.SequenceMatcher(None, squash(pdf_plain), squash(fallback)).ratio() < 0.9:
            self.fallback += 1
            return fallback
        self.used += 1

        line_right = max(r for _, r, _ in lines)
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
                log.warning("%s 第 %s 頁表格無法辨識欄位，已略過", source_name, page_no)
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
                log.warning("%s 第 %s 頁序號 %s 找不到列管案號，已建立空白案號的資料列",
                            source_name, page_no, seq)
                current = Record(case_no="", seq=seq, source_pdf=source_name)
                records.append(current)
                is_continuation = True
            if not is_continuation:
                if not case_no:
                    log.warning("%s 第 %s 頁有一列找不到列管案號，內容：%s",
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


CHECKED = "█■☑☒✔✓✅▇◼●◉"
UNCHECKED = "□☐◻○"


def parse_suggestion(text: str) -> str:
    """「█除管 □續管」→「除管」；無法判斷時保留原文。"""
    t = squash(text)
    if not t:
        return ""
    picked = re.findall(rf"[{CHECKED}]([^{CHECKED}{UNCHECKED}]+)", t)
    if picked:
        return "\n".join(p.strip() for p in picked)
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


def post_json(url: str, payload: dict, retries: int = 3) -> dict:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    ctx = ssl_context()
    for attempt in range(1, retries + 1):
        # Apps Script 會先回 302 轉址到 googleusercontent.com，urllib 會自動改用 GET 取回結果
        req = urllib.request.Request(
            url, data=data, method="POST",
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
        except ssl.SSLCertVerificationError:
            raise RuntimeError("SSL 憑證驗證失敗：請執行 /Applications/Python 3.x/Install Certificates.command，"
                               "或 python3 -m pip install certifi")
        except (urllib.error.URLError, TimeoutError) as exc:
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

def main() -> int:
    ap = argparse.ArgumentParser(description="PDF 列管表 → docling JSON → Google 試算表")
    ap.add_argument("input_dir", type=Path, help="放 PDF 的資料夾")
    ap.add_argument("--json-dir", type=Path, help="docling JSON 輸出資料夾（預設：<input_dir>/_docling_json）")
    ap.add_argument("--webapp", default=os.environ.get("GAS_WEBAPP_URL"),
                    help="Apps Script 網頁應用程式網址（…/exec），或設環境變數 GAS_WEBAPP_URL")
    ap.add_argument("--token", default=os.environ.get("GAS_TOKEN"),
                    help="與 Code.gs 中 TOKEN 相同的密碼，或設環境變數 GAS_TOKEN")
    ap.add_argument("--worksheet", default="列管表", help="工作表名稱（不存在會自動建立）")
    ap.add_argument("--mode", choices=["upsert", "append", "replace"], default="upsert",
                    help="upsert：同 PDF 同案號覆寫，其餘新增（預設）；append：全部新增；replace：清空後重寫")
    ap.add_argument("--recursive", action="store_true", help="一併處理子資料夾內的 PDF")
    ap.add_argument("--ocr", action="store_true", help="掃描型 PDF 啟用 OCR（macOS 使用內建 Vision）")
    ap.add_argument("--force", action="store_true", help="即使已有 JSON 也重新解析 PDF")
    ap.add_argument("--from-json", action="store_true", help="跳過步驟一，直接讀 --json-dir 內現有的 JSON")
    ap.add_argument("--dry-run", action="store_true", help="不上傳，只輸出 structured.csv / structured.json 供檢查")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")

    input_dir: Path = args.input_dir.expanduser().resolve()
    if not input_dir.is_dir():
        log.error("找不到資料夾：%s", input_dir)
        return 1
    json_dir = (args.json_dir or input_dir / "_docling_json").expanduser().resolve()

    # 步驟一
    if args.from_json:
        json_paths = sorted(json_dir.glob("*.json"))
    else:
        pattern = "**/*" if args.recursive else "*"
        pdfs = sorted(p for p in input_dir.glob(pattern)
                      if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith("."))
        if not pdfs:
            log.error("資料夾內沒有 PDF：%s", input_dir)
            return 1
        log.info("共 %d 份 PDF", len(pdfs))
        json_paths = convert_pdfs(pdfs, json_dir, args.ocr, args.force)

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
                if pdf.fallback:
                    log.info("  %d 個儲存格無法對應 PDF 行位置，改用 docling 文字", pdf.fallback)
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
    if not args.webapp or not args.token:
        log.error("請以 --webapp 與 --token 指定 Apps Script 網址與密碼（或加 --dry-run）")
        return 1
    try:
        upload(rows, args.webapp, args.token, args.worksheet, args.mode)
    except RuntimeError as exc:
        log.error("%s", exc)
        return 1
    log.info("完成")
    return 0


if __name__ == "__main__":
    sys.exit(main())
