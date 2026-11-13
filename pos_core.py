"""
POS Monthly Data Transformation Core  (v2.0)
============================================

Pure-Python / pandas core (NO Streamlit dependency).

Pipeline
--------
    POS files -> parse filename -> detect template -> transform (standard schema)
              -> Customer Mapping -> SKU Mapping -> Final Report
              -> reconciliation / submission tracking / export helpers

Policies (unchanged from v1)
----------------------------
* SKU mapping is NAME-BASED ONLY:
      RAW_SKU -> norm_key() -> SKU_KEY -> SKU Mapping["SKU Name"]
  Product codes / barcodes are never mapping keys and never appear in output.
  They may be inspected internally only for template detection, continuation
  rows, or excluding subtotal/header rows.
* Final Report prices come ONLY from SKU Mapping RSP
      單價 = RSP, 建議售價 = RSP, 總價 = QTY x RSP
* Responsible sales always comes from Customer Mapping (never the file name).
* No aggregation: every source row is preserved in the Final Report.

See CHANGE_RECORD.md for every change made relative to v1 and why.
"""

from __future__ import annotations

import csv
import io
import re
import time
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

__version__ = "2.0.0"


# ============================================================
# CONSTANTS
# ============================================================

ROW = "__ROW__"  # internal column: original 1-based Excel / CSV row number

OUTPUT_COLUMNS = [
    "地區", "業務員", "合約等級", "合約編號", "合約名稱", "店家編號", "店家名稱",
    "銷量(瓶)", "單價", "總價", "建議售價", "價格帶", "製造商", "品牌",
    "統一品項名稱", "品項", "SKU Mapping Status", "容量", "品類", "年", "月",
]

# Extra columns appended in the "trace" version of the report (D6)
TRACE_COLUMNS = [
    "SOURCE_FILE", "SOURCE_SHEET", "SOURCE_ROW", "TEMPLATE_TYPE",
    "FILE_OUTLET_NO", "CUSTOMER_MAPPING_STATUS", "CUSTOMER_MATCH_METHOD",
    "RSP_STATUS", "QTY_UNIT", "UNIT_STATUS", "QTY_FLAG",
    "DUPLICATE_SUBMISSION", "RAW_UNIT_PRICE", "RAW_TOTAL_PRICE",
]

CUSTOMER_REQUIRED = [
    "REGION", "Sales ID", "Sales", "CONTRACT TYPE", "Contract JDE",
    "Contract NAME", "Outlet No", "Outlet NAME", "Rawdata Name",
]

SKU_REQUIRED = [
    "Mapping Name", "RSP", "Price Band", "Manufacture", "Brand",
    "SKU Name", "SIZE", "CATEGORY",
]

STANDARD_HEADER_MAPPING = {
    "SKU": [
        "貨品名稱", "商品名稱", "品項名稱", "品名", "商品品名", "產品名稱",
        "SKU Name", "Product Name", "產品品名", "單品名稱",
    ],
    "QTY": [
        "數量", "數 量", "銷量", "銷量 Quantity", "Quantity", "QTY", "瓶數", "銷售量",
    ],
    "CUSTOMER": ["客戶名稱", "客戶", "店家名稱", "門市名稱", "Customer Name"],
    "UNIT_PRICE": ["單價", "售價", "價格", "Unit Price", "Price"],
    "TOTAL_PRICE": [
        "小計", "小計金額", "總價", "總金額", "銷貨金額", "金額", "Amount", "Total", "Total Price",
    ],
    "BRAND": ["品牌", "品牌名稱", "品牌名稱 Brand Name", "Brand", "Brand Name"],
    "SIZE": ["容量", "規格", "規格 Size (ml)", "Size", "Size (ml)"],
    "CATEGORY": ["品類", "品類 Category", "Category"],
}

# --- Shared alias lists: used by BOTH detection and transformation (B1) ---
T1_SKU_ALIASES = [
    "品項名稱 SKU Name", "SKU Name", "品項名稱", "品名", "商品名稱", "品名規格", "產品品名",
]
T1_QTY_ALIASES = [
    "銷量 Quantity", "Quantity", "銷量", "銷售量", "數量", "瓶數", "淨銷數量",
]
BRAND_ALIASES = ["品牌名稱 Brand Name", "品牌名稱", "Brand Name"]
SIZE_ALIASES = ["規格 Size (ml)", "Size (ml)", "規格", "容量"]
CATEGORY_ALIASES = ["品類 Category", "Category", "品類"]
T15_SKU_ALIASES = ["產品名稱", "商品名稱", "貨品名稱", "品項名稱", "品名"]
T15_QTY_ALIASES = ["銷售量", "數量", "銷量", "Quantity", "QTY"]

# Structural (internal-only) product-code header aliases
STRUCT_CODE_ALIASES = [
    "產品編號", "產品代號", "商品編號", "商品代號", "商品號碼",
    "貨品編號", "貨品代號", "貨品號碼", "品號", "SKU Code",
]
T16_NAME_ALIASES = [
    "產品品名", "產品名稱", "商品名稱", "商品品名", "貨品名稱",
    "品項名稱", "品名", "SKU Name", "Product Name",
]


class TemplateError(ValueError):
    """Expected, recoverable template problem (missing header / columns / data)."""


# ============================================================
# CONFIG
# ============================================================

@dataclass
class PipelineConfig:
    # B5: one NaN-QTY policy for ALL templates.
    drop_nan_qty: bool = True
    # A2: flag (not drop) quantities whose absolute value exceeds this.
    qty_abs_max: float = 100000
    # A3: "flag" = keep all rows & flag; "keep_first" / "keep_last" = drop the other files' rows.
    duplicate_policy: str = "flag"
    # A4: unit -> bottles multiplier. Units not listed here and in review_units
    # are flagged NEEDS_REVIEW and NOT converted. Add e.g. {"箱": 6} once known.
    unit_factors: Dict[str, float] = field(
        default_factory=lambda: {"瓶": 1.0, "罐": 1.0, "支": 1.0}
    )
    review_units: Tuple[str, ...] = ("箱", "盒", "組", "包", "桶")
    # A7: sheets whose name contains one of these are skipped in multi-sheet templates.
    summary_sheet_keywords: Tuple[str, ...] = (
        "彙總", "匯總", "總表", "summary", "total", "合計", "總計",
    )
    # T7 special customers: keep only rows whose 單位 is in t7_unit_allow.
    t7_unit_filter_keywords: Tuple[str, ...] = ("駱北", "全億")
    t7_unit_allow: Tuple[str, ...] = ("瓶", "盒", "組")
    # A8: "reject" = a Rawdata Name shared by several outlets is NOT used for fallback;
    #     "first"  = v1 behaviour (silently take the first).
    ambiguous_rawdata_fallback: str = "reject"


DEFAULT_CONFIG = PipelineConfig()


# ============================================================
# NORMALISATION HELPERS
# ============================================================

def norm_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = unicodedata.normalize("NFKC", str(value)).strip()
    return re.sub(r"\s+", " ", text)


def norm_key(value) -> str:
    text = norm_text(value).upper()
    return re.sub(r"[^\w\u4e00-\u9fff]+", "", text, flags=re.UNICODE)


def norm_col(value) -> str:
    text = norm_text(value).lower()
    text = re.sub(r"[_\-]+", " ", text)
    return re.sub(r"\s+", " ", text)


def norm_id_key(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):  # 31020549.0 -> 31020549
        text = text.split(".")[0]
    return re.sub(r"[^\w]+", "", text).upper()


def clean_file_label(filename) -> str:
    stem = Path(str(filename)).stem
    return re.sub(r"\(\d+\)$", "", stem).strip()


INVALID_SKU_KEYS = {
    norm_key(x) for x in [
        "商品名稱", "商品", "品名", "品項名稱", "產品名稱", "產品品名", "貨品名稱",
        "品名規格", "條碼", "商品條碼", "Barcode", "單位", "數量", "銷量", "金額",
        "售價", "小計", "合計", "總計", "TOTAL", "以下空白",
    ]
}

# ---------------------------------------------------------------- numbers

_UNIT_STRIP_RE = re.compile(
    r"(瓶|支|罐|箱|盒|包|桶|組|入|個|PCS|PC|BOTTLES|BOTTLE|NT\$|NTD|\$|元)", re.I
)
_UNIT_FIND_RE = re.compile(r"(瓶|支|罐|箱|盒|包|桶|組)")
_PLAIN_NUMBER_RE = re.compile(r"[-+]?(?:\d+\.?\d*|\.\d+)")


def parse_number(value) -> float:
    """
    Strict single-cell number parser (replaces v1 numeric() regex-extract and T8 to_number).

    Accepts : 12  1,200  26瓶  2,408瓶  12.5  (5) -> -5
    Rejects : barcode-like integers (8-14 digits, also when stored as float 4713268580974.0),
              free text such as "700cc", "2026-09", "約12"  -> NaN   (A2, C2)
    """
    if value is None or isinstance(value, bool):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        v = float(value)
        if np.isnan(v) or np.isinf(v):
            return np.nan
        if v.is_integer() and 1e7 <= abs(v) < 1e14:  # barcode-like
            return np.nan
        return v
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return np.nan
    text = norm_text(value)
    if not text:
        return np.nan
    text = text.replace(",", "")
    text = _UNIT_STRIP_RE.sub("", text)
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d{8,14}(?:\.0+)?", text):  # barcode-like
        return np.nan
    m = re.fullmatch(r"\((\d+\.?\d*|\.\d+)\)", text)  # accounting negative
    if m:
        return -float(m.group(1))
    if _PLAIN_NUMBER_RE.fullmatch(text):
        return float(text)
    return np.nan


def parse_unit(value) -> str:
    """Unit character found inside a quantity cell text ('' when none / numeric)."""
    if isinstance(value, str):
        m = _UNIT_FIND_RE.search(unicodedata.normalize("NFKC", value))
        if m:
            return m.group(1)
    return ""


def numeric(series: pd.Series) -> pd.Series:
    return series.map(parse_number).astype("float64")


# ============================================================
# FILE NAME PARSER
# ============================================================

def normalize_period_key(value) -> str:
    """202609 -> 202609 ; ROC 11509 -> 202609 ; otherwise ''."""
    digits = re.sub(r"\D", "", norm_text(value))
    if len(digits) == 6:
        y, m = int(digits[:4]), int(digits[4:])
        if 1900 <= y <= 2200 and 1 <= m <= 12:
            return f"{y}{m:02d}"
    elif len(digits) == 5:
        roc, m = int(digits[:3]), int(digits[3:])
        if 100 <= roc <= 199 and 1 <= m <= 12:  # C5: ROC year sanity check
            return f"{roc + 1911}{m:02d}"
    return ""


def parse_filename(filename) -> dict:
    name = Path(str(filename)).name
    stem_raw = Path(name).stem
    stem = clean_file_label(name)
    parts = [p.strip() for p in stem.split("_") if p.strip()]

    sales_id = parts[0] if len(parts) > 0 else ""
    contract_jde = parts[1] if len(parts) > 1 else ""
    period = parts[2] if len(parts) > 2 else ""
    outlet_no = parts[3] if len(parts) > 3 else ""
    rawdata_name = "_".join(parts[4:]) if len(parts) > 4 else ""

    key = normalize_period_key(period)
    year = int(key[:4]) if key else None
    month = int(key[4:]) if key else None

    ok = all([sales_id, contract_jde, outlet_no, year, month])
    return {
        "FILE_SALES_ID": sales_id,
        "FILE_CONTRACT_JDE": contract_jde,
        "FILE_PERIOD": period,
        "FILE_OUTLET_NO": outlet_no,
        "FILE_RAWDATA_NAME": rawdata_name,
        "YEAR": year,
        "MONTH": month,
        "FILENAME_STATUS": "OK" if ok else "CHECK",
        "FILE_COPY_SUFFIX": stem != stem_raw,  # name ended with "(1)" etc.
    }


# ============================================================
# SOURCE FILE  (read once, cache everything)  - D4 / B11 / B12
# ============================================================

CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp950", "big5")


def decode_csv_bytes(data: bytes) -> str:
    last = None
    for enc in CSV_ENCODINGS:
        try:
            return data.decode(enc)
        except UnicodeDecodeError as exc:
            last = exc
    raise ValueError(
        f"CSV 無法解碼（已嘗試 {', '.join(CSV_ENCODINGS)}）：{last}"
    )


def read_csv_raw(data: bytes) -> pd.DataFrame:
    """Header-less, ragged-safe CSV reader -> DataFrame of strings/NaN."""
    text = decode_csv_bytes(data)
    rows = [[c.strip() for c in r] for r in csv.reader(io.StringIO(text))]
    if not rows:
        return pd.DataFrame()
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    df = pd.DataFrame(rows, dtype=object)
    return df.mask(df == "", np.nan)


class SourceFile:
    """Immutable wrapper around one uploaded file. Parsed lazily and cached."""

    def __init__(self, name: str, data: bytes):
        self.name = Path(str(name)).name
        self.data = bytes(data)
        self.suffix = Path(self.name).suffix.lower()
        self.stem_text = norm_text(Path(self.name).stem)
        self._sheets: Optional[Dict[str, pd.DataFrame]] = None

    # -------------------------------------------------- construction
    @classmethod
    def from_any(cls, source) -> "SourceFile":
        if isinstance(source, cls):
            return source
        if isinstance(source, (str, Path)):
            p = Path(source)
            return cls(p.name, p.read_bytes())
        if isinstance(source, tuple) and len(source) == 2:
            return cls(source[0], source[1])
        name = getattr(source, "name", None)
        if name is None:
            raise TypeError(
                "source must be a path, (filename, bytes), SourceFile, "
                "or a file-like object with .name"
            )
        if hasattr(source, "getvalue"):
            data = source.getvalue()
        else:
            try:
                source.seek(0)
            except Exception:
                pass
            data = source.read()
        return cls(name, data)

    # -------------------------------------------------- readers
    def _engines(self) -> list:
        if self.suffix == ".xls":
            return ["xlrd", "calamine", None]
        if self.suffix in (".xlsx", ".xlsm"):
            return ["openpyxl", "calamine", None]
        return [None]

    def read_excel(self, **kwargs):
        """pd.read_excel with engine fall-back (xlrd / openpyxl / calamine / auto)."""
        last = None
        for engine in self._engines():
            try:
                return pd.read_excel(io.BytesIO(self.data), engine=engine, **kwargs)
            except Exception as exc:  # ImportError, XLRDError, BadZipFile, ...
                last = exc
        raise ValueError(f"無法讀取 Excel（{self.name}）：{last}")

    def sheets(self) -> Dict[str, pd.DataFrame]:
        """{sheet_name: header-less DataFrame}  (CSV -> {'CSV': df})."""
        if self._sheets is None:
            if self.suffix == ".csv":
                self._sheets = {"CSV": read_csv_raw(self.data)}
            elif self.suffix in (".xlsx", ".xls", ".xlsm"):
                result = self.read_excel(sheet_name=None, header=None)
                if not isinstance(result, dict) or not result:
                    raise ValueError(f"{self.name} 找不到任何工作表。")
                self._sheets = {str(k): v for k, v in result.items()}
            else:
                raise ValueError(f"不支援的檔案格式：{self.suffix}")
        return self._sheets

    def first_sheet(self) -> Tuple[str, pd.DataFrame]:
        name = next(iter(self.sheets()))
        return name, self.sheets()[name]


def as_source_file(source) -> SourceFile:
    return SourceFile.from_any(source)


# ============================================================
# TABLE HELPERS
# ============================================================

def _unique_headers(headers: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    out = []
    for i, h in enumerate(headers):
        h = h or f"Unnamed: {i}"
        if h in seen:
            seen[h] += 1
            h = f"{h}.{seen[h]}"
        else:
            seen[h] = 0
        out.append(h)
    return out


def promote_header(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    """Use raw row `header_row` as header. Adds ROW = original 1-based row number (B8)."""
    if header_row >= len(raw):
        raise TemplateError("表頭列超出資料範圍。")
    headers = _unique_headers([norm_text(v) for v in raw.iloc[header_row].tolist()])
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df[ROW] = df.index + 1
    return df.reset_index(drop=True)


def drop_blank_rows(df: pd.DataFrame) -> pd.DataFrame:
    data_cols = [c for c in df.columns if c != ROW]
    if not data_cols:
        return df
    tmp = df[data_cols].replace(r"^\s*$", np.nan, regex=True)
    return df.loc[tmp.notna().any(axis=1)].reset_index(drop=True)


def first_matching_column(df, keywords, exclude=(), contains=True):
    """
    Column lookup by header text (norm_key based, so '數 量' == '數量').
    Priority follows the ORDER OF `keywords` (C1): exact match for every keyword first,
    then (optionally) contains-match for every keyword.
    """
    cols = [c for c in df.columns if c != ROW and c not in exclude]
    keys = {c: norm_key(c) for c in cols}
    kk = [norm_key(k) for k in keywords if norm_key(k)]
    for k in kk:
        for c in cols:
            if keys[c] == k:
                return c
    if contains:
        for k in kk:
            if len(k) < 2:
                continue
            for c in cols:
                if k in keys[c]:
                    return c
    return None


def find_header_row(raw, required_fields, max_rows=30):
    """First row in which EVERY required logical field has an exact (norm_key) alias match."""
    alias_keys = {
        f: {norm_key(a) for a in STANDARD_HEADER_MAPPING.get(f, [])}
        for f in required_fields
    }
    for i in range(min(max_rows, len(raw))):
        values = {norm_key(v) for v in raw.iloc[i].tolist() if norm_text(v)}
        if all(alias_keys[f] & values for f in required_fields):
            return i
    return None


def _row_values(raw: pd.DataFrame, i: int) -> List[str]:
    return [norm_text(v) for v in raw.iloc[i].tolist()]


def select_sheets(src: SourceFile, cfg: PipelineConfig, notes: list):
    """Multi-sheet templates: skip summary sheets (A7) and report which sheets are used."""
    chosen, skipped = [], []
    for name, raw in src.sheets().items():
        low = name.lower()
        if any(k.lower() in low for k in cfg.summary_sheet_keywords):
            skipped.append(name)
        else:
            chosen.append((name, raw))
    if skipped:
        notes.append(f"已略過彙總工作表：{', '.join(skipped)}")
    if not chosen:  # everything looked like a summary -> use everything
        chosen = list(src.sheets().items())
    return chosen


STD_DEFAULTS = {
    "RAW_ROW_CUSTOMER": "", "RAW_QTY": "", "QTY_UNIT": "",
    "RAW_BRAND": "", "RAW_SIZE": "", "RAW_CATEGORY": "",
    "RAW_UNIT_PRICE": np.nan, "RAW_TOTAL_PRICE": np.nan,
    "SOURCE_QTY_COLUMN": "", "SOURCE_SHEET": "",
}


def table_frame(
    df, sheet, sku_col, qty_col, *, brand_col=None, size_col=None, category_col=None,
    customer_col=None, unit_price_col=None, total_price_col=None, unit_col=None,
):
    """Standard rows from a promoted-header table."""
    out = pd.DataFrame(index=df.index)
    out["RAW_SKU"] = df[sku_col].map(norm_text)
    out["QTY"] = numeric(df[qty_col])
    out["RAW_QTY"] = df[qty_col].map(norm_text)
    out["QTY_UNIT"] = (
        df[unit_col].map(norm_text) if unit_col is not None else df[qty_col].map(parse_unit)
    )
    out["SOURCE_ROW"] = df[ROW]
    out["SOURCE_SHEET"] = sheet
    out["SOURCE_QTY_COLUMN"] = str(qty_col)
    if brand_col is not None:
        out["RAW_BRAND"] = df[brand_col].map(norm_text)
    if size_col is not None:
        out["RAW_SIZE"] = df[size_col].map(norm_text)
    if category_col is not None:
        out["RAW_CATEGORY"] = df[category_col].map(norm_text)
    if customer_col is not None:
        out["RAW_ROW_CUSTOMER"] = df[customer_col].map(norm_text)
    if unit_price_col is not None:
        out["RAW_UNIT_PRICE"] = numeric(df[unit_price_col])
    if total_price_col is not None:
        out["RAW_TOTAL_PRICE"] = numeric(df[total_price_col])
    return out[out["RAW_SKU"].ne("")].reset_index(drop=True)


def make_output(src, meta, template_type, frame, cfg, notes=None):
    """
    Single place that turns template-specific rows into the standard schema and applies the
    cross-template policies:
      - remove header / subtotal / footer labels                      (B3)
      - ONE missing-QTY policy for every template                     (B5)
      - unit handling (convert / flag)                                (A4)
      - QTY sanity flag (negative / out of range)                     (A2)
    Non-row information is returned in out.attrs["notes"].
    """
    notes = list(notes or [])
    out = frame.copy().reset_index(drop=True)

    for col, default in STD_DEFAULTS.items():
        if col not in out.columns:
            out[col] = default
    if "SOURCE_ROW" not in out.columns:
        out["SOURCE_ROW"] = range(1, len(out) + 1)

    out["RAW_SKU"] = out["RAW_SKU"].map(norm_text)
    out["QTY"] = pd.to_numeric(out["QTY"], errors="coerce")
    out["QTY_UNIT"] = out["QTY_UNIT"].map(norm_text)

    invalid = out["RAW_SKU"].eq("") | out["RAW_SKU"].map(norm_key).isin(INVALID_SKU_KEYS)
    if invalid.any():
        notes.append(f"已排除 {int(invalid.sum())} 筆表頭 / 合計 / 空白品項列")
    out = out.loc[~invalid]

    nan_qty = out["QTY"].isna()
    if nan_qty.any():
        if cfg.drop_nan_qty:
            notes.append(f"已排除 {int(nan_qty.sum())} 筆數量空白或無法解析的列")
            out = out.loc[~nan_qty]
        else:
            notes.append(f"保留 {int(nan_qty.sum())} 筆數量空白的列")
    out = out.reset_index(drop=True)

    if out.empty:
        raise TemplateError("轉換後沒有任何有效的 SKU / 數量資料。")

    # ---- unit handling (A4)
    out["QTY_ORIGINAL"] = out["QTY"]
    factor = out["QTY_UNIT"].map(lambda u: cfg.unit_factors.get(u))
    has_factor = factor.notna()
    out.loc[has_factor, "QTY"] = out.loc[has_factor, "QTY"] * factor[has_factor].astype(float)

    def unit_status(u):
        if u == "":
            return "ASSUMED_BOTTLE"
        if u in cfg.unit_factors:
            return "OK" if float(cfg.unit_factors[u]) == 1.0 else "CONVERTED"
        if u in cfg.review_units:
            return "NEEDS_REVIEW"
        return "UNKNOWN_UNIT"

    out["UNIT_STATUS"] = out["QTY_UNIT"].map(unit_status)
    n_review = int(out["UNIT_STATUS"].isin(["NEEDS_REVIEW", "UNKNOWN_UNIT"]).sum())
    if n_review:
        notes.append(f"{n_review} 筆單位非「瓶」且尚未設定換算係數（UNIT_STATUS=NEEDS_REVIEW）")

    # ---- qty flags (A2)
    flag = np.where(
        out["QTY"].abs() > cfg.qty_abs_max, "OUT_OF_RANGE",
        np.where(out["QTY"] < 0, "NEGATIVE", ""),
    )
    out["QTY_FLAG"] = flag
    n_flag = int((out["QTY_FLAG"] != "").sum())
    if n_flag:
        notes.append(f"{n_flag} 筆數量為負數或超出合理範圍（QTY_FLAG）")

    # ---- multiple row-level customers inside one file (B7)
    rc = out["RAW_ROW_CUSTOMER"].map(norm_text)
    if rc.replace("", np.nan).nunique() > 1:
        notes.append(
            "檔案內含多個逐列客戶名稱；系統仍以檔名店家歸屬，請確認是否為單一店家檔案"
        )

    # ---- assemble columns
    out["SOURCE_FILE"] = src.name
    out["TEMPLATE_TYPE"] = template_type
    for k, v in meta.items():
        out[k] = v
    out["SOURCE_PRICE_COLUMN"] = "RSP_FROM_SKU_MAPPING"

    head = ["SOURCE_FILE", "SOURCE_SHEET", "SOURCE_ROW", "TEMPLATE_TYPE"] + list(meta.keys())
    rest = [c for c in out.columns if c not in head]
    out = out[head + rest]
    out.attrs["notes"] = notes
    return out


# ============================================================
# TEMPLATE TRANSFORMS  (each: (src, cfg) -> DataFrame)
# ============================================================

def transform_template_1(src, cfg):
    """T1 Standard POS: header on row 1 (xlsx / csv)."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    df = drop_blank_rows(promote_header(raw, 0))
    sku = first_matching_column(df, T1_SKU_ALIASES)
    qty = first_matching_column(df, T1_QTY_ALIASES, exclude=[sku])
    if sku is None or qty is None:
        raise TemplateError(f"Template 1 找不到必要欄位。 SKU={sku}, Quantity={qty}")
    frame = table_frame(
        df, sheet, sku, qty,
        brand_col=first_matching_column(df, BRAND_ALIASES),
        size_col=first_matching_column(df, SIZE_ALIASES),
        category_col=first_matching_column(df, CATEGORY_ALIASES),
    )
    return make_output(src, meta, "T1_STANDARD_POS", frame, cfg)


def _to_year_month(value):
    """Cell -> (year, month) or None. Numbers are NOT treated as epoch timestamps (A5)."""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        ts = pd.Timestamp(value)
        return (ts.year, ts.month) if pd.notna(ts) else None
    if isinstance(value, (int, float, np.integer, np.floating)) and not pd.isna(value):
        key = normalize_period_key(str(int(value)))
        return (int(key[:4]), int(key[4:])) if key else None
    text = norm_text(value)
    if not text:
        return None
    ts = pd.to_datetime(text, errors="coerce")
    if pd.notna(ts):
        return ts.year, ts.month
    key = normalize_period_key(text)
    return (int(key[:4]), int(key[4:])) if key else None


def transform_template_2(src, cfg):
    """T2 Monthly actual sales (horizontal): 商品名稱 + the 實銷 column of the filename month."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    if len(raw) < 4:
        raise TemplateError("Template 2 資料列不足。")

    header_row = None
    for i in range(min(20, len(raw))):
        values = _row_values(raw, i)
        if "商品代號" in values and "商品名稱" in values:
            header_row = i
            break
    if header_row is None:
        raise TemplateError("Template 2 找不到「商品代號 / 商品名稱」表頭。")

    headers = _row_values(raw, header_row)
    dates = raw.iloc[max(0, header_row - 1)].tolist()
    ff, current = [], None
    for v in dates:
        if pd.notna(v) and norm_text(v):
            current = v
        ff.append(current)

    name_idx = headers.index("商品名稱")
    notes = []
    idx = None
    target = (meta["YEAR"], meta["MONTH"])
    for i, h in enumerate(headers):
        if "實銷" in h and meta["YEAR"] is not None and _to_year_month(ff[i]) == target:
            idx = i
            break
    if idx is None:  # A5: fallback is allowed but is now reported
        idx = next((i for i, h in enumerate(headers) if "實銷" in h), None)
        if idx is None:
            raise TemplateError("Template 2 找不到「實銷」欄位。")
        notes.append(
            f"⚠ 找不到與檔名月份 {meta['FILE_PERIOD']} 對應的「實銷」欄，已改用第一個「實銷」欄，請人工確認月份"
        )

    data = raw.iloc[header_row + 1:]
    frame = pd.DataFrame({
        "RAW_SKU": data.iloc[:, name_idx].map(norm_text),
        "QTY": numeric(data.iloc[:, idx]),
        "RAW_QTY": data.iloc[:, idx].map(norm_text),
        "QTY_UNIT": data.iloc[:, idx].map(parse_unit),
        "SOURCE_ROW": data.index + 1,
        "SOURCE_SHEET": sheet,
        "SOURCE_QTY_COLUMN": "實銷",
    })
    frame = frame[frame["RAW_SKU"].ne("")]
    return make_output(src, meta, "T2_MONTHLY_ACTUAL", frame, cfg, notes)


def transform_template_3_multisheet(src, cfg):
    """T3A multi-sheet detail: 品名規格 + 數量2/數量 on every sheet."""
    meta = parse_filename(src.name)
    notes, frames = [], []
    for name, raw in select_sheets(src, cfg, notes):
        if raw.empty or len(raw) < 2:
            continue
        df = drop_blank_rows(promote_header(raw, 0))
        sku = first_matching_column(df, ["品名規格"])
        qty = first_matching_column(df, ["數量2", "數量"])
        if sku is None or qty is None:
            continue
        frames.append(table_frame(
            df, name, sku, qty,
            customer_col=first_matching_column(df, ["名稱"], contains=False),
            unit_price_col=first_matching_column(df, ["單價", "價格"]),
            total_price_col=first_matching_column(df, ["總價", "金額"]),
        ))
    if not frames:
        raise TemplateError("Template 3 多工作表找不到「品名規格 / 數量2」。")
    notes.append("已合併工作表：" + ", ".join(sorted({f['SOURCE_SHEET'].iloc[0] for f in frames if len(f)})))
    return make_output(src, meta, "T3_TRANSACTION_MULTISHEET",
                       pd.concat(frames, ignore_index=True), cfg, notes)


def transform_template_3_salesdetail(src, cfg):
    """T3B single-sheet sales detail: 貨品名稱 + 數量."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    header_row = None
    for i in range(min(30, len(raw))):
        values = _row_values(raw, i)
        if "貨品名稱" in values and any("數" in v and "量" in v for v in values):
            header_row = i
            break
    if header_row is None:
        raise TemplateError("Template 3 銷貨明細找不到表頭。")
    df = drop_blank_rows(promote_header(raw, header_row))
    sku = first_matching_column(df, ["貨品名稱"])
    qty = first_matching_column(df, ["數 量", "數量"])
    if sku is None or qty is None:
        raise TemplateError("Template 3 銷貨明細找不到貨品名稱或數量。")
    frame = table_frame(
        df, sheet, sku, qty,
        customer_col=first_matching_column(df, ["客戶名稱"]),
        unit_price_col=first_matching_column(df, ["單 價", "單價"]),
        total_price_col=first_matching_column(df, ["總 價", "總價"]),
    )
    return make_output(src, meta, "T3_SALES_DETAIL", frame, cfg)


_T4_FOOTER_RE = re.compile(
    r"^(以下空白|合計|總計|小計|第\s*\d+\s*頁|頁次.*|page\s*\d+.*|共\s*\d+\s*頁|"
    r"\d{2,4}[-/.年]\d{1,2}[-/.月]\d{1,2}.*)$", re.I
)


def transform_template_4(src, cfg):
    """T4 麥金酒 multi-page report: QTY = column right of 商品名稱; wrapped names are re-joined."""
    meta = parse_filename(src.name)
    notes, records = [], []

    for sheet, raw in select_sheets(src, cfg, notes):
        if raw is None or raw.empty:
            continue
        header_row = None
        for i in range(min(20, len(raw))):
            values = _row_values(raw, i)
            if "商品貨號" in values and "商品名稱" in values:
                header_row = i
                break
        if header_row is None:
            continue
        headers = _row_values(raw, header_row)
        code_idx = headers.index("商品貨號")
        name_idx = headers.index("商品名稱")
        qty_idx = name_idx + 1
        if qty_idx >= raw.shape[1]:
            raise TemplateError(f"Template 4｜{sheet} 找不到數量欄位。")

        last = None  # continuation target (reset on every sheet)
        for r in range(header_row + 1, len(raw)):
            row = raw.iloc[r]
            marker = norm_text(row.iloc[code_idx])  # structural marker only
            sku = norm_text(row.iloc[name_idx])
            qty = parse_number(row.iloc[qty_idx])

            if sku and norm_key(sku) == norm_key("商品名稱"):  # repeated page header
                continue
            if sku and not np.isnan(qty):
                records.append({
                    "RAW_SKU": sku, "QTY": qty,
                    "RAW_QTY": norm_text(row.iloc[qty_idx]),
                    "QTY_UNIT": parse_unit(row.iloc[qty_idx]),
                    "SOURCE_ROW": r + 1, "SOURCE_SHEET": sheet,
                    "SOURCE_QTY_COLUMN": "商品名稱右側欄位",
                })
                last = len(records) - 1
                continue
            if marker and not sku and np.isnan(qty) and last is not None:
                if _T4_FOOTER_RE.match(marker):
                    continue
                records[last]["RAW_SKU"] += marker  # wrapped product name
    if not records:
        raise TemplateError("Template 4 找不到有效商品資料。")
    return make_output(src, meta, "T4_MAIJIN_MULTI_PAGE", pd.DataFrame(records), cfg, notes)


def _bilingual(values, cn, en, alts):
    return any((cn in v and en in v) or v in alts for v in values)


def _t5_header_flags(values):
    return {
        "sku": _bilingual(values, "品項名稱", "SKU Name", {"品項名稱"}),
        "qty": _bilingual(values, "銷量", "Quantity", {"銷量"}),
        "brand": _bilingual(values, "品牌名稱", "Brand Name", {"品牌名稱"}),
        "size": _bilingual(values, "規格", "Size", {"規格", "容量"}),
        "category": _bilingual(values, "品類", "Category", {"品類"}),
    }


def transform_template_5(src, cfg):
    """T5 POS template with header within the first 10 rows (e.g. row 2)."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    header_row = None
    for i in range(min(10, len(raw))):
        f = _t5_header_flags(_row_values(raw, i))
        if f["sku"] and f["qty"]:
            header_row = i
            break
    if header_row is None:
        raise TemplateError("Template 5 找不到「品項名稱 SKU Name / 銷量 Quantity」表頭。")
    df = drop_blank_rows(promote_header(raw, header_row))
    sku = first_matching_column(df, ["品項名稱 SKU Name", "品項名稱", "SKU Name"])
    qty = first_matching_column(df, ["銷量 Quantity", "銷量", "Quantity"], exclude=[sku])
    if sku is None or qty is None:
        raise TemplateError(f"Template 5 找不到必要欄位。 SKU={sku}, Quantity={qty}")
    frame = table_frame(
        df, sheet, sku, qty,
        brand_col=first_matching_column(df, BRAND_ALIASES),
        size_col=first_matching_column(df, SIZE_ALIASES),
        category_col=first_matching_column(df, CATEGORY_ALIASES),
    )
    return make_output(src, meta, "T5_POS_HEADER_ROW_2", frame, cfg)


def transform_template_6(src, cfg):
    """T6 松賀: no header, fixed columns 日期|客戶號碼|客戶名稱|商品號碼|商品名稱|單價|數量|小計."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    raw = raw.dropna(how="all")  # index preserved -> true row numbers
    if raw.empty:
        raise TemplateError("Template 6 檔案沒有資料。")
    if raw.shape[1] < 8:
        raise TemplateError(f"Template 6 欄位數不足。 Expected >= 8, Actual={raw.shape[1]}")
    frame = pd.DataFrame({
        "RAW_SKU": raw.iloc[:, 4].map(norm_text),
        "QTY": numeric(raw.iloc[:, 6]),
        "RAW_QTY": raw.iloc[:, 6].map(norm_text),
        "QTY_UNIT": raw.iloc[:, 6].map(parse_unit),
        "RAW_ROW_CUSTOMER": raw.iloc[:, 2].map(norm_text),
        "RAW_CUSTOMER_NO": raw.iloc[:, 1].map(norm_text),
        "RAW_TRANSACTION_DATE": raw.iloc[:, 0].map(norm_text),
        "RAW_UNIT_PRICE": numeric(raw.iloc[:, 5]),
        "RAW_TOTAL_PRICE": numeric(raw.iloc[:, 7]),
        "SOURCE_ROW": raw.index + 1,
        "SOURCE_SHEET": sheet,
        "SOURCE_QTY_COLUMN": "Column 7｜數量",
    })
    frame = frame[frame["RAW_SKU"].ne("")]
    return make_output(src, meta, "T6_NO_HEADER_TRANSACTION", frame, cfg)


def transform_template_7(src, cfg):
    """T7 Generic SKU + QTY (header found by alias within first 30 rows; xlsx / csv)."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    if raw.empty:
        raise TemplateError("Template 7 檔案沒有資料。")
    header_row = find_header_row(raw, ["SKU", "QTY"], max_rows=30)
    if header_row is None:
        raise TemplateError("Template 7 找不到 SKU / QTY 類型表頭。")
    df = drop_blank_rows(promote_header(raw, header_row))
    sku = first_matching_column(df, STANDARD_HEADER_MAPPING["SKU"])
    qty = first_matching_column(df, STANDARD_HEADER_MAPPING["QTY"], exclude=[sku])
    if sku is None or qty is None:
        raise TemplateError(f"Template 7 找不到必要欄位。 SKU={sku}, Quantity={qty}")

    notes, unit_col = [], None
    if any(k in src.stem_text for k in cfg.t7_unit_filter_keywords):
        unit_col = first_matching_column(df, ["單位", "Unit"], contains=False)
        if unit_col is None:
            raise TemplateError("Template 7｜駱北 / 全億 找不到「單位」欄位。")
        keep = df[unit_col].map(norm_text).isin(cfg.t7_unit_allow)
        dropped = int((~keep).sum())
        if dropped:
            notes.append(f"駱北/全億：已排除 {dropped} 筆單位不在 {'/'.join(cfg.t7_unit_allow)} 的列")
        df = df.loc[keep].reset_index(drop=True)

    frame = table_frame(
        df, sheet, sku, qty, unit_col=unit_col,
        customer_col=first_matching_column(df, STANDARD_HEADER_MAPPING["CUSTOMER"]),
        unit_price_col=first_matching_column(df, STANDARD_HEADER_MAPPING["UNIT_PRICE"], exclude=[sku, qty]),
        total_price_col=first_matching_column(df, STANDARD_HEADER_MAPPING["TOTAL_PRICE"], exclude=[sku, qty]),
        brand_col=first_matching_column(df, STANDARD_HEADER_MAPPING["BRAND"]),
        size_col=first_matching_column(df, STANDARD_HEADER_MAPPING["SIZE"]),
        category_col=first_matching_column(df, STANDARD_HEADER_MAPPING["CATEGORY"]),
    )
    return make_output(src, meta, "T7_TRANSACTION_DETAIL", frame, cfg, notes)


_T8_UNITS = {"瓶", "罐", "支", "箱", "盒", "組"}


def transform_template_8(src, cfg):
    """T8 唯君: free-form rows. Unit cell -> first number after it = QTY; text before it = SKU."""
    meta = parse_filename(src.name)
    notes, frames = [], []

    def is_numeric_only(value):
        text = norm_text(value).replace(",", "").replace("$", "").strip()
        return bool(re.fullmatch(r"[-+]?\d*\.?\d+", text))

    for sheet, raw in select_sheets(src, cfg, notes):
        if raw is None or raw.empty:
            continue
        records = []
        for r in range(len(raw)):
            row = raw.iloc[r]
            unit_col = next((i for i, v in enumerate(row) if norm_text(v) in _T8_UNITS), None)
            if unit_col is None:
                continue
            unit_value = norm_text(row.iloc[unit_col])

            qty_col, qty_value = None, np.nan
            for i in range(unit_col + 1, len(row)):
                v = parse_number(row.iloc[i])
                if np.isnan(v) or abs(v) > 100000:
                    continue
                qty_col, qty_value = i, v
                break
            if qty_col is None:
                continue

            parts = []
            for i in range(unit_col):
                value = norm_text(row.iloc[i])
                if (not value or value in _T8_UNITS or is_numeric_only(value)
                        or norm_key(value) in INVALID_SKU_KEYS
                        or re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", value)):
                    continue
                if not parts or parts[-1] != value:
                    parts.append(value)
            sku = norm_text(" ".join(parts))
            if not sku or norm_key(sku) in INVALID_SKU_KEYS:
                continue

            amount_col, amount = None, np.nan
            for i in range(qty_col + 1, len(row)):
                v = parse_number(row.iloc[i])
                if not np.isnan(v):
                    amount_col, amount = i, v

            records.append({
                "RAW_SKU": sku, "QTY": qty_value, "QTY_UNIT": unit_value,
                "RAW_QTY": norm_text(row.iloc[qty_col]),
                "RAW_TOTAL_PRICE": amount, "SOURCE_ROW": r + 1, "SOURCE_SHEET": sheet,
                "SOURCE_UNIT": unit_value, "SOURCE_UNIT_COLUMN": unit_col,
                "SOURCE_QTY_COLUMN": f"col{qty_col + 1}",
                "SOURCE_AMOUNT_COLUMN": "" if amount_col is None else amount_col,
            })
        if records:
            frames.append(pd.DataFrame(records))
    if not frames:
        raise TemplateError("Template 8 所有 Sheet 都沒有找到有效的 SKU / Unit / QTY 資料。")
    return make_output(src, meta, "T8_WEIJUN_ROW_LEVEL",
                       pd.concat(frames, ignore_index=True), cfg, notes)


def transform_template_9(src, cfg):
    """T9 新店李 / Diageo: 品號 + 品名規格 + 淨銷數量."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    header_row = None
    for i in range(min(20, len(raw))):
        keys = {norm_key(v) for v in raw.iloc[i].tolist() if norm_text(v)}
        if {norm_key("品號"), norm_key("品名規格"), norm_key("淨銷數量")} <= keys:
            header_row = i
            break
    if header_row is None:
        raise TemplateError("Template 9 找不到「品號 / 品名規格 / 淨銷數量」表頭。")
    df = drop_blank_rows(promote_header(raw, header_row))
    sku = first_matching_column(df, ["品名規格"])
    qty = first_matching_column(df, ["淨銷數量"])
    if sku is None:
        raise TemplateError("Template 9 找不到「品名規格」欄位。")
    if qty is None:
        raise TemplateError("Template 9 找不到「淨銷數量」欄位。")
    return make_output(src, meta, "T9_DIAGEO_NET_SALES", table_frame(df, sheet, sku, qty), cfg)


def transform_template_10(src, cfg):
    """T10 名酒集: no header; column 1 = SKU, column 2 = QTY (after dropping empty rows/cols)."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        raise TemplateError("Template 10 檔案沒有資料。")
    if raw.shape[1] < 2:
        raise TemplateError(f"Template 10 欄位不足。 Expected >= 2, Actual={raw.shape[1]}")
    frame = pd.DataFrame({
        "RAW_SKU": raw.iloc[:, 0].map(norm_text),
        "QTY": numeric(raw.iloc[:, 1]),
        "RAW_QTY": raw.iloc[:, 1].map(norm_text),
        "QTY_UNIT": raw.iloc[:, 1].map(parse_unit),
        "SOURCE_ROW": raw.index + 1,
        "SOURCE_SHEET": sheet,
        "SOURCE_QTY_COLUMN": "Column 2",
    })
    frame = frame[frame["RAW_SKU"].ne("")]
    return make_output(src, meta, "T10_NO_HEADER_SKU_QTY", frame, cfg)


def transform_template_11(src, cfg):
    """T11 酒江村: header row 1; 品牌 column is used as the product name, 數量 = QTY."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    df = drop_blank_rows(promote_header(raw, 0))
    sku = first_matching_column(df, ["品牌"])
    qty = first_matching_column(df, ["數量"], exclude=[sku])
    if sku is None or qty is None:
        raise TemplateError("Template 11 找不到「品牌 / 數量」欄位。")
    return make_output(src, meta, "T11_JIUJIANGCUN", table_frame(df, sheet, sku, qty), cfg)


def transform_template_12(src, cfg):
    """T12 廣茂隆: POS template whose QTY header is blank -> QTY = last column (validated)."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    df = drop_blank_rows(promote_header(raw, 0))
    sku = first_matching_column(df, ["品項名稱 SKU Name", "品項名稱"])
    if sku is None:
        raise TemplateError("Template 12 找不到品項名稱。")
    data_cols = [c for c in df.columns if c != ROW]
    qty = data_cols[-1]
    non_blank = df[qty].map(norm_text).ne("")
    if non_blank.any():
        share = df.loc[non_blank, qty].map(parse_number).notna().mean()
        if share < 0.5:  # C8: last column does not look like a quantity column
            raise TemplateError(
                f"Template 12 最後一欄「{qty}」看起來不是數量欄（數值比例 {share:.0%}）。"
            )
    frame = table_frame(
        df, sheet, sku, qty,
        brand_col=first_matching_column(df, ["品牌名稱 Brand Name", "品牌名稱"]),
        size_col=first_matching_column(df, ["規格 Size (ml)", "規格"]),
        category_col=first_matching_column(df, ["品類 Category", "品類"]),
    )
    frame["SOURCE_QTY_COLUMN"] = "LAST_COLUMN"
    return make_output(src, meta, "T12_POS_BLANK_QTY_HEADER", frame, cfg,
                       [f"T12：使用最後一欄「{qty}」作為數量"])


_MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


def transform_template_13(src, cfg):
    """T13 橡木桶: column B = SKU, QTY = month column (JAN..DEC) of the filename month."""
    meta = parse_filename(src.name)
    if meta["MONTH"] is None:
        raise TemplateError("Template 13｜橡木桶 無法從檔名取得月份。")
    month_col = _MONTHS[meta["MONTH"] - 1]
    sheet, raw = src.first_sheet()
    if raw.empty:
        raise TemplateError("Template 13｜橡木桶 檔案沒有資料。")

    header_row = None
    for i in range(min(30, len(raw))):
        found = {v.upper() for v in _row_values(raw, i)} & set(_MONTHS)
        if len(found) >= 6:
            header_row = i
            break
    if header_row is None:
        raise TemplateError("Template 13｜橡木桶 找不到 JAN～DEC 月份 Header。")
    headers = [v.upper() for v in _row_values(raw, header_row)]
    if month_col not in headers:
        raise TemplateError(f"Template 13｜橡木桶 找不到當月份欄位：{month_col}")
    notes = []
    if headers.count(month_col) > 1:
        notes.append(f"{month_col} 欄位重複出現 {headers.count(month_col)} 次，已使用第一個")
    qty_idx = headers.index(month_col)

    data = raw.iloc[header_row + 1:]
    marker = data.iloc[:, 0].map(norm_text)  # structural marker only
    sku = data.iloc[:, 1].map(norm_text)
    qty = numeric(data.iloc[:, qty_idx])
    valid = marker.ne("") & sku.ne("") & qty.notna()
    upper = marker.str.upper()
    for kw in ("BOOTLES", "BOTTLES", "TOTAL", "NO."):
        valid &= ~upper.str.startswith(kw)
    frame = pd.DataFrame({
        "RAW_SKU": sku[valid], "QTY": qty[valid],
        "RAW_QTY": data.iloc[:, qty_idx].map(norm_text)[valid],
        "QTY_UNIT": "",
        "SOURCE_ROW": data.index[valid] + 1, "SOURCE_SHEET": sheet,
        "SOURCE_QTY_COLUMN": month_col,
    })
    return make_output(src, meta, "T13_OAK_BARREL", frame, cfg, notes)


_D_CODE_RE = re.compile(r"^D\d{2,4}(?:-[0-9A-Za-z]+)+$", re.I)  # D001-00-070-001


def transform_template_14(src, cfg):
    """T14 美滿商行 (.xls): column D = SKU, column F = QTY. D-code rows are excluded (A1)."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    if raw.empty:
        raise TemplateError("Template 14 檔案沒有資料。")
    if raw.shape[1] < 6:
        raise TemplateError(f"Template 14 欄位不足。目前只有 {raw.shape[1]} 欄，至少需要到 F 欄。")
    sku = raw.iloc[:, 3].map(norm_text)
    # A1: exclude ONLY code-like values (D001-00-070-001), NOT every name starting with "D".
    is_code = sku.map(lambda s: bool(_D_CODE_RE.match(s.replace(" ", ""))))
    notes = []
    if is_code.any():
        notes.append(f"已排除 {int(is_code.sum())} 筆 D 開頭商品代碼列")
    frame = pd.DataFrame({
        "RAW_SKU": sku, "QTY": numeric(raw.iloc[:, 5]),
        "RAW_QTY": raw.iloc[:, 5].map(norm_text),
        "QTY_UNIT": raw.iloc[:, 5].map(parse_unit),
        "SOURCE_ROW": raw.index + 1, "SOURCE_SHEET": sheet, "SOURCE_QTY_COLUMN": "F",
    })
    frame = frame[sku.ne("") & ~is_code]
    return make_output(src, meta, "T14_MEIMAN_FIXED_D_F", frame, cfg, notes)


def transform_template_15(src, cfg):
    """T15 精湛酒藏: multi-sheet, header row 1, SKU + QTY aliases."""
    meta = parse_filename(src.name)
    notes, frames = [], []
    for name, raw in select_sheets(src, cfg, notes):
        if raw.empty or len(raw) < 2:
            continue
        df = drop_blank_rows(promote_header(raw, 0))
        sku = first_matching_column(df, T15_SKU_ALIASES)
        qty = first_matching_column(df, T15_QTY_ALIASES, exclude=[sku])
        if sku is None or qty is None:
            continue
        frames.append(table_frame(df, name, sku, qty))
    if not frames:
        raise TemplateError("Template 15 沒有找到可用的 SKU / QTY Sheet。")
    return make_output(src, meta, "T15_MULTI_SHEET_SKU_QTY",
                       pd.concat(frames, ignore_index=True), cfg, notes)


def _store_header_matches(store_key: str, header_values) -> bool:
    if len(store_key) < 2:
        return False
    for v in header_values:
        k = norm_key(v)
        if not k:
            continue
        if k == store_key or store_key in k or (len(k) >= 3 and k in store_key):
            return True
    return False


def _find_store_column(df, store_key, exclude):
    if len(store_key) < 2:
        return None
    cols = [c for c in df.columns if c != ROW and c not in exclude]
    for c in cols:
        if norm_key(c) == store_key:
            return c
    for c in cols:
        k = norm_key(c)
        if k and (store_key in k or (len(k) >= 3 and k in store_key)):
            return c
    return None


def transform_template_16(src, cfg):
    """T16 Store sales matrix: product name + the column of THIS store (B4)."""
    meta = parse_filename(src.name)
    sheet, raw = src.first_sheet()
    if raw.empty:
        raise TemplateError("Template 16 檔案沒有資料。")

    code_keys = {norm_key(a) for a in STRUCT_CODE_ALIASES}
    name_keys = {norm_key(a) for a in T16_NAME_ALIASES}
    header_row = None
    for i in range(min(30, len(raw))):
        keys = {norm_key(v) for v in raw.iloc[i].tolist() if norm_text(v)}
        if keys & code_keys and keys & name_keys:
            header_row = i
            break
    if header_row is None:
        raise TemplateError("Template 16 找不到「產品編號 + 產品品名」類型 Header。")

    df = drop_blank_rows(promote_header(raw, header_row))
    code_col = first_matching_column(df, STRUCT_CODE_ALIASES)
    sku_col = first_matching_column(df, T16_NAME_ALIASES, exclude=[code_col])
    if code_col is None or sku_col is None:
        raise TemplateError(f"Template 16 找不到產品編號/品名欄位。 Columns={list(df.columns)}")

    notes = []
    store_key = norm_key(meta["FILE_RAWDATA_NAME"])
    qty_col = _find_store_column(df, store_key, exclude=[code_col, sku_col])
    if qty_col is None:  # fallback: column right of the product name (v1 behaviour) + warning
        cols = [c for c in df.columns if c != ROW]
        pos = cols.index(sku_col) + 1
        if pos >= len(cols):
            raise TemplateError("Template 16 產品品名右邊沒有欄位，因此無法取得銷量。")
        qty_col = cols[pos]
        notes.append(
            f"⚠ T16：表頭找不到與檔名店名「{meta['FILE_RAWDATA_NAME']}」對應的欄位，"
            f"已改用品名右側欄位「{qty_col}」，請確認是否為該店"
        )

    marker = df[code_col].map(norm_text).map(norm_key)
    invalid_codes = {norm_key(x) for x in ["TOTAL", "合計", "總計", "NO", "NO.", "產品編號", "商品編號", "品號"]}
    frame = table_frame(df, sheet, sku_col, qty_col)
    keep_rows = set(df.loc[~marker.isin(invalid_codes), ROW])
    frame = frame[frame["SOURCE_ROW"].isin(keep_rows)]
    return make_output(src, meta, "T16_STORE_SALES_MATRIX", frame, cfg, notes)


# ============================================================
# TEMPLATE DETECTION
# ============================================================

def _find_rows(raw, predicate, max_rows):
    for i in range(min(max_rows, len(raw))):
        if predicate(_row_values(raw, i)):
            return i
    return None


def _detect_generic(raw) -> Optional[str]:
    """Shared by CSV and Excel (B2): T7 first, then T1 (same order everywhere)."""
    if find_header_row(raw, ["SKU", "QTY"], max_rows=30) is not None:
        return "TEMPLATE_7"
    df = promote_header(raw, 0)
    sku = first_matching_column(df, T1_SKU_ALIASES)
    qty = first_matching_column(df, T1_QTY_ALIASES, exclude=[sku])
    if sku is not None and qty is not None:
        return "TEMPLATE_1"
    return None


def detect_template_type(src) -> str:
    src = as_source_file(src)
    name = src.stem_text

    # ---- filename / suffix based
    if src.suffix == ".xls" and "美滿商行" in name:
        return "TEMPLATE_14"
    if "橡木桶" in name:
        return "TEMPLATE_13"
    if "合歡" in name:
        return "TEMPLATE_16"

    # ---- CSV
    if src.suffix == ".csv":
        try:
            return _detect_generic(src.first_sheet()[1]) or "UNKNOWN"
        except Exception:
            return "UNKNOWN"
    if src.suffix not in (".xlsx", ".xls", ".xlsm"):
        return "UNKNOWN"

    try:
        sheets = src.sheets()
        raw = next(iter(sheets.values()))
    except Exception:
        return "UNKNOWN"
    n_sheets = len(sheets)

    def safe(fn):
        try:
            return bool(fn())
        except Exception:
            return False

    # T4 麥金酒
    def t4():
        head_text = " ".join(norm_text(v) for v in raw.iloc[:20].to_numpy().flatten())
        has_header = _find_rows(raw, lambda v: "商品貨號" in v and "商品名稱" in v, 30) is not None
        return has_header and ("麥金酒" in name or "麥金酒" in head_text or "期間商品銷售統計表" in head_text)
    if safe(t4):
        return "TEMPLATE_4"

    # T5 POS template
    def t5():
        def pred(v):
            f = _t5_header_flags(v)
            return f["sku"] and f["qty"] and f["brand"] and (f["size"] or f["category"])
        return _find_rows(raw, pred, 15) is not None
    if safe(t5):
        return "TEMPLATE_5"

    if "松賀" in name and raw.shape[1] >= 8:
        return "TEMPLATE_6"
    if "唯君" in name:
        return "TEMPLATE_8"

    # T9 (moved BEFORE T3A so a multi-sheet T9 file is not taken as T3A - C6)
    def t9():
        need = {norm_key("品號"), norm_key("品名規格"), norm_key("淨銷數量")}
        for i in range(min(20, len(raw))):
            if need <= {norm_key(v) for v in raw.iloc[i].tolist() if norm_text(v)}:
                return True
        return False
    if safe(t9):
        return "TEMPLATE_9"

    # T3A multi-sheet
    def t3a():
        if n_sheets <= 1:
            return False
        keys = [norm_key(c) for c in promote_header(raw, 0).columns if c != ROW]
        return any("品名規格" in k for k in keys) and any("數量" in k for k in keys)
    if safe(t3a):
        return "TEMPLATE_3"

    # T2
    def t2():
        found_header = found_actual = False
        for i in range(min(30, len(raw))):
            v = _row_values(raw, i)
            found_header |= "商品代號" in v and "商品名稱" in v
            found_actual |= any("實銷" in x for x in v)
        return found_header and found_actual
    if safe(t2):
        return "TEMPLATE_2"

    # T3B
    if safe(lambda: _find_rows(
            raw, lambda v: "貨品名稱" in v and any("數" in x and "量" in x for x in v), 40) is not None):
        return "TEMPLATE_3"

    if "名酒集" in name and raw.shape[1] >= 2:
        return "TEMPLATE_10"

    # T11
    def t11():
        keys = {norm_key(v) for v in raw.iloc[0].tolist() if norm_text(v)}
        return {norm_key("國際條碼"), norm_key("品牌"), norm_key("數量")} <= keys
    if safe(t11):
        return "TEMPLATE_11"

    # T12
    def t12():
        keys = [norm_key(v) for v in raw.iloc[0].tolist() if norm_text(v)]
        return ("廣茂隆" in name and any(norm_key("品項名稱") in k for k in keys) and raw.shape[1] >= 10)
    if safe(t12):
        return "TEMPLATE_12"

    if "精湛酒藏" in name and n_sheets > 1:
        return "TEMPLATE_15"

    # T16 by header: code + name header, NO normal QTY header, and a header matching this store
    def t16():
        code_keys = {norm_key(a) for a in STRUCT_CODE_ALIASES}
        name_keys = {norm_key(a) for a in T16_NAME_ALIASES}
        qty_keys = {norm_key(a) for a in STANDARD_HEADER_MAPPING["QTY"]}
        for i in range(min(30, len(raw))):
            keys = {norm_key(v) for v in raw.iloc[i].tolist() if norm_text(v)}
            if keys & code_keys and keys & name_keys:
                if keys & qty_keys:
                    return False
                store_key = norm_key(parse_filename(src.name)["FILE_RAWDATA_NAME"])
                return _store_header_matches(store_key, _row_values(raw, i))
        return False
    if safe(t16):
        return "TEMPLATE_16"

    try:
        generic = _detect_generic(raw)
    except Exception:
        generic = None
    return generic or "UNKNOWN"


# ============================================================
# AUTO TRANSFORM
# ============================================================

TEMPLATE_LABELS = {
    "TEMPLATE_1": "Template 1｜標準 PRT_POS Template",
    "TEMPLATE_2": "Template 2｜月份實銷表",
    "TEMPLATE_3A": "Template 3A｜多工作表銷售明細",
    "TEMPLATE_3B": "Template 3B｜單工作表銷貨明細",
    "TEMPLATE_4": "Template 4｜麥金酒期間商品銷售統計表",
    "TEMPLATE_5": "Template 5｜POS 公版_HEADER on 2nd row",
    "TEMPLATE_6": "Template 6｜(松賀) no HEADER",
    "TEMPLATE_7": "Template 7｜(七七會社) Generic SKU / QTY",
    "TEMPLATE_8": "Template 8｜唯君系列",
    "TEMPLATE_9": "Template 9｜新店李淨銷報表",
    "TEMPLATE_10": "Template 10｜名酒集無表頭 SKU / QTY",
    "TEMPLATE_11": "Template 11｜酒江村商品銷量",
    "TEMPLATE_12": "Template 12｜POS 公版無 QTY Header",
    "TEMPLATE_13": "Template 13｜橡木桶",
    "TEMPLATE_14": "Template 14｜美滿商行商品彙總",
    "TEMPLATE_15": "Template 15｜Multi-sheet SKU / QTY",
    "TEMPLATE_16": "Template 16｜Store Sales Matrix",
}

TEMPLATE_FUNCTIONS: Dict[str, Callable] = {
    "TEMPLATE_1": transform_template_1,
    "TEMPLATE_2": transform_template_2,
    "TEMPLATE_4": transform_template_4,
    "TEMPLATE_5": transform_template_5,
    "TEMPLATE_6": transform_template_6,
    "TEMPLATE_7": transform_template_7,
    "TEMPLATE_8": transform_template_8,
    "TEMPLATE_9": transform_template_9,
    "TEMPLATE_10": transform_template_10,
    "TEMPLATE_11": transform_template_11,
    "TEMPLATE_12": transform_template_12,
    "TEMPLATE_13": transform_template_13,
    "TEMPLATE_14": transform_template_14,
    "TEMPLATE_15": transform_template_15,
    "TEMPLATE_16": transform_template_16,
}


def auto_transform_template(src, template_type=None, cfg: Optional[PipelineConfig] = None):
    cfg = cfg or DEFAULT_CONFIG
    src = as_source_file(src)
    template_type = template_type or detect_template_type(src)

    if template_type == "TEMPLATE_3":
        if len(src.sheets()) > 1:
            try:  # A6: only an EXPECTED template mismatch may fall back to 3B
                return transform_template_3_multisheet(src, cfg), TEMPLATE_LABELS["TEMPLATE_3A"]
            except TemplateError:
                pass
        return transform_template_3_salesdetail(src, cfg), TEMPLATE_LABELS["TEMPLATE_3B"]

    func = TEMPLATE_FUNCTIONS.get(template_type)
    if func is None:
        raise TemplateError(f"無法辨識此檔案的 Template 格式：{template_type}")
    return func(src, cfg), TEMPLATE_LABELS[template_type]


def transform_file(source, template_type=None, cfg: Optional[PipelineConfig] = None):
    """Detect + transform ONE file -> (DataFrame, template_label, template_code)."""
    src = as_source_file(source)
    code = template_type or detect_template_type(src)
    data, label = auto_transform_template(src, code, cfg)
    data = data.copy()
    data.attrs = {}
    return data, label, code


# ============================================================
# MAPPING
# ============================================================

def resolve_required(df, required):
    lookup = {norm_key(c): c for c in df.columns}
    resolved, missing = {}, []
    for name in required:
        key = norm_key(name)
        if key in lookup:
            resolved[name] = lookup[key]
        else:
            missing.append(name)
    return resolved, missing


def read_mapping_file(source, preferred_sheet=None) -> pd.DataFrame:
    """Mapping file (csv / xlsx / xls) -> DataFrame with the first row as header."""
    src = as_source_file(source)
    if src.suffix == ".csv":
        raw = read_csv_raw(src.data)
        df = promote_header(raw, 0).drop(columns=[ROW])
        return df
    xl = pd.ExcelFile(io.BytesIO(src.data), engine=src._engines()[0])
    sheet = preferred_sheet if preferred_sheet and preferred_sheet in xl.sheet_names else xl.sheet_names[0]
    return src.read_excel(sheet_name=sheet, header=0)


def prepare_customer_mapping(df):
    """Return (mapping, duplicates). duplicates has DUP_TYPE = OUTLET_NO | RAWDATA_NAME (A8)."""
    resolved, missing = resolve_required(df, CUSTOMER_REQUIRED)
    if missing:
        raise ValueError("Customer Mapping 缺少欄位：" + ", ".join(missing))
    mapping = pd.DataFrame({col: df[src] for col, src in resolved.items()})
    mapping["OUTLET_KEY"] = mapping["Outlet No"].map(norm_id_key)
    mapping["SALES_KEY"] = mapping["Sales ID"].map(norm_id_key)
    mapping["CONTRACT_KEY"] = mapping["Contract JDE"].map(norm_id_key)
    mapping["RAWDATA_KEY"] = mapping["Rawdata Name"].map(norm_key)

    dup_outlet = mapping[
        mapping["OUTLET_KEY"].ne("") & mapping.duplicated("OUTLET_KEY", keep=False)
    ].assign(DUP_TYPE="OUTLET_NO")

    m = mapping[mapping["RAWDATA_KEY"].ne("")]
    ambiguous = m.groupby("RAWDATA_KEY")["OUTLET_KEY"].transform("nunique") > 1
    dup_raw = m[ambiguous].assign(DUP_TYPE="RAWDATA_NAME")

    return mapping, pd.concat([dup_outlet, dup_raw], ignore_index=True)


SKU_ATTR_COLUMNS = ["Mapping Name", "RSP", "Price Band", "Manufacture", "Brand", "SIZE", "CATEGORY"]


def prepare_sku_mapping(df):
    """Return (mapping_for_join, duplicates). duplicates carries CONFLICT=True when attributes differ (A9)."""
    df = df.copy()
    df.columns = [norm_text(c) for c in df.columns]
    resolved, missing = resolve_required(df, SKU_REQUIRED)
    if missing:
        raise ValueError(
            "SKU Mapping 缺少欄位：" + ", ".join(missing) + f"\n目前欄位：{list(df.columns)}"
        )
    mapping = pd.DataFrame({name: df[resolved[name]] for name in SKU_REQUIRED})
    if "SKU_MAPPING_SOURCE_FILE" in df.columns:
        mapping["SKU_MAPPING_SOURCE_FILE"] = df["SKU_MAPPING_SOURCE_FILE"]
    mapping["SKU_KEY"] = mapping["SKU Name"].map(norm_key)  # JOIN KEY = SKU Name
    mapping = mapping[mapping["SKU_KEY"].ne("")].copy()

    dup = mapping[mapping.duplicated("SKU_KEY", keep=False)].copy()
    if not dup.empty:
        nun = (
            dup[SKU_ATTR_COLUMNS]
            .apply(lambda s: s.map(norm_text))
            .assign(SKU_KEY=dup["SKU_KEY"])
            .groupby("SKU_KEY")[SKU_ATTR_COLUMNS]
            .transform("nunique")
        )
        dup["CONFLICT"] = (nun.max(axis=1) > 1).values
    else:
        dup["CONFLICT"] = pd.Series(dtype=bool)

    join = mapping.drop_duplicates("SKU_KEY", keep="first").reset_index(drop=True)
    return join, dup


def load_customer_mapping(source, preferred_sheet=None):
    return prepare_customer_mapping(read_mapping_file(source, preferred_sheet))


def load_sku_mapping(sources, preferred_sheet=None):
    """
    One file or a list of files -> (mapping_for_join, duplicates, upload_log).
    A failed file is logged and skipped; at least one file must succeed.
    """
    if not isinstance(sources, (list, tuple)) or (
        isinstance(sources, tuple) and len(sources) == 2 and isinstance(sources[1], (bytes, bytearray))
    ):
        sources = [sources]
    frames, logs = [], []
    for s in sources:
        try:
            src = as_source_file(s)
        except Exception as exc:
            logs.append({"檔案": str(s), "狀態": "FAILED", "筆數": 0, "錯誤": f"{type(exc).__name__}: {exc}"})
            continue
        try:
            df = read_mapping_file(src, preferred_sheet)
            if df is None or df.empty:
                raise ValueError("SKU Mapping 沒有資料")
            df = df.copy()
            df["SKU_MAPPING_SOURCE_FILE"] = src.name
            frames.append(df)
            logs.append({"檔案": src.name, "狀態": "SUCCESS", "筆數": len(df), "錯誤": ""})
        except Exception as exc:
            logs.append({"檔案": src.name, "狀態": "FAILED", "筆數": 0, "錯誤": f"{type(exc).__name__}: {exc}"})
    log = pd.DataFrame(logs)
    if not frames:
        raise ValueError("所有 SKU Mapping 檔案都讀取失敗：" + "; ".join(log["錯誤"].astype(str)))
    join, dup = prepare_sku_mapping(pd.concat(frames, ignore_index=True, sort=False))
    return join, dup, log


def _norm_series(s: pd.Series, f) -> pd.Series:
    return s.map(f) if len(s) else s.astype(object)


def attach_customer_mapping(raw, customer_map, cfg: Optional[PipelineConfig] = None):
    """
    Match order: 1) FILE_OUTLET_NO <-> Outlet No   2) FILE_RAWDATA_NAME <-> Rawdata Name (fallback)
    Adds CUSTOMER_MATCH_METHOD (OUTLET / RAWDATA_FALLBACK / NONE) (A8) and vectorised audit checks.
    """
    cfg = cfg or DEFAULT_CONFIG
    source = raw.copy().reset_index(drop=True)
    source["FILE_OUTLET_KEY"] = source["FILE_OUTLET_NO"].map(norm_id_key)
    source["FILE_RAWDATA_KEY"] = source["FILE_RAWDATA_NAME"].map(norm_key)

    map_cols = CUSTOMER_REQUIRED + ["OUTLET_KEY", "RAWDATA_KEY", "SALES_KEY", "CONTRACT_KEY"]

    by_outlet = (
        customer_map[customer_map["OUTLET_KEY"].ne("")]
        .drop_duplicates("OUTLET_KEY", keep="first")[map_cols]
    )
    raw_pool = customer_map[customer_map["RAWDATA_KEY"].ne("")]
    if cfg.ambiguous_rawdata_fallback == "reject":
        n_out = raw_pool.groupby("RAWDATA_KEY")["OUTLET_KEY"].transform("nunique")
        raw_pool = raw_pool[n_out <= 1]
    by_raw = raw_pool.drop_duplicates("RAWDATA_KEY", keep="first")[map_cols]

    m1 = source[["FILE_OUTLET_KEY"]].merge(
        by_outlet, how="left", left_on="FILE_OUTLET_KEY", right_on="OUTLET_KEY"
    )[map_cols]
    m2 = source[["FILE_RAWDATA_KEY"]].merge(
        by_raw, how="left", left_on="FILE_RAWDATA_KEY", right_on="RAWDATA_KEY"
    )[map_cols]

    ok1 = m1["Outlet No"].map(norm_text).ne("").values
    ok2 = m2["Outlet No"].map(norm_text).ne("").values
    use2 = (~ok1) & ok2
    chosen = m1.copy()
    chosen.loc[use2, map_cols] = m2.loc[use2, map_cols]

    mapped = ok1 | use2
    merged = pd.concat([source, chosen], axis=1)
    merged["CUSTOMER_MATCH_METHOD"] = np.where(ok1, "OUTLET", np.where(use2, "RAWDATA_FALLBACK", "NONE"))
    merged["CUSTOMER_MAPPING_STATUS"] = np.where(mapped, "MAPPED", "UNMAPPED")
    merged["SUBMISSION_CUSTOMER_ID"] = merged["FILE_OUTLET_KEY"]

    merged["RESPONSIBLE_SALES_ID"] = merged["Sales ID"].fillna("").map(norm_id_key)
    merged["RESPONSIBLE_SALES"] = merged["Sales"].fillna("").map(norm_text)
    has_sales = merged["Sales ID"].map(norm_id_key).ne("")
    merged["RESPONSIBLE_SALES_STATUS"] = np.where(mapped & has_sales, "MAPPED", "UNMAPPED")
    merged["CUSTOMER_OWNER_STATUS"] = merged["RESPONSIBLE_SALES_STATUS"]  # alias kept for compatibility

    def check(file_col, map_col, func):
        same = merged[file_col].fillna("").map(func) == merged[map_col].map(func)
        return np.where(mapped, np.where(same, "MATCH", "MISMATCH"), "UNMAPPED")

    merged["FILE_SALES_ID_AUDIT"] = merged["FILE_SALES_ID"].fillna("").map(norm_id_key)
    merged["SALES_ID_CHECK"] = check("FILE_SALES_ID", "Sales ID", norm_id_key)
    merged["CONTRACT_CHECK"] = check("FILE_CONTRACT_JDE", "Contract JDE", norm_id_key)
    merged["OUTLET_CHECK"] = check("FILE_OUTLET_NO", "Outlet No", norm_id_key)
    merged["RAWDATA_NAME_CHECK"] = check("FILE_RAWDATA_NAME", "Rawdata Name", norm_key)
    return merged


def attach_sku_mapping(df, sku_map):
    """
    LEFT JOIN RAW_SKU -> SKU Name.  Status is based on the JOIN result (A9):
      MAPPED     joined and Mapping Name present
      INCOMPLETE joined but Mapping Name blank
      UNMAPPED   no join
    RSP_STATUS: OK / MISSING / UNMAPPED
    """
    out = df.copy()
    out["SKU_KEY"] = out["RAW_SKU"].map(norm_key)
    cols = ["SKU_KEY", "SKU Name", "Mapping Name", "RSP", "Price Band",
            "Manufacture", "Brand", "SIZE", "CATEGORY"]
    out = out.merge(sku_map[cols], how="left", on="SKU_KEY", indicator="_SKU_MERGE")
    joined = out["_SKU_MERGE"].eq("both")
    has_name = out["Mapping Name"].map(norm_text).ne("")
    out["SKU_MAPPING_STATUS"] = np.where(joined & has_name, "MAPPED", np.where(joined, "INCOMPLETE", "UNMAPPED"))
    rsp_ok = numeric(out["RSP"]).notna()
    out["RSP_STATUS"] = np.where(joined, np.where(rsp_ok, "OK", "MISSING"), "UNMAPPED")
    return out.drop(columns=["_SKU_MERGE"])


# ============================================================
# FINAL REPORT
# ============================================================

def safe_column(df, name, default=pd.NA):
    if name in df.columns:
        return df[name]
    return pd.Series([default] * len(df), index=df.index)


def build_report(detail, include_trace=False):
    """
    One report row per detail row (no aggregation). Prices come ONLY from RSP.
    店家編號 = file outlet no, EXCEPT when the customer was matched by Rawdata-Name fallback:
    then the mapped Outlet No is used so that 店家編號 and 店家名稱 describe the same store (A8).
    """
    d = detail
    rep = pd.DataFrame(index=d.index)
    method = safe_column(d, "CUSTOMER_MATCH_METHOD", "OUTLET")
    file_outlet = safe_column(d, "FILE_OUTLET_NO", "")
    mapped_outlet = safe_column(d, "Outlet No", "").map(norm_id_key)

    rep["地區"] = safe_column(d, "REGION")
    rep["業務員"] = safe_column(d, "Sales")
    rep["合約等級"] = safe_column(d, "CONTRACT TYPE")
    rep["合約編號"] = safe_column(d, "Contract JDE")
    rep["合約名稱"] = safe_column(d, "Contract NAME")
    rep["店家編號"] = file_outlet.where(method.ne("RAWDATA_FALLBACK"), mapped_outlet)
    rep["店家名稱"] = safe_column(d, "Outlet NAME")

    qty = pd.to_numeric(safe_column(d, "QTY"), errors="coerce")
    rsp = numeric(safe_column(d, "RSP", np.nan))
    rep["銷量(瓶)"] = qty
    rep["單價"] = rsp
    rep["總價"] = qty * rsp
    rep["建議售價"] = rsp
    rep["價格帶"] = safe_column(d, "Price Band")
    rep["製造商"] = safe_column(d, "Manufacture")
    rep["品牌"] = safe_column(d, "Brand")
    rep["統一品項名稱"] = safe_column(d, "Mapping Name")
    rep["品項"] = safe_column(d, "RAW_SKU")
    rep["SKU Mapping Status"] = (
        safe_column(d, "SKU_MAPPING_STATUS").fillna("UNMAPPED").astype(str).str.strip().str.upper()
        .replace({"": "UNMAPPED", "NAN": "UNMAPPED", "NONE": "UNMAPPED", "<NA>": "UNMAPPED"})
    )
    rep["容量"] = safe_column(d, "SIZE")
    rep["品類"] = safe_column(d, "CATEGORY")
    rep["年"] = pd.to_numeric(safe_column(d, "YEAR"), errors="coerce").astype("Int64")
    rep["月"] = pd.to_numeric(safe_column(d, "MONTH"), errors="coerce").astype("Int64")

    for t in TRACE_COLUMNS:
        rep[t] = safe_column(d, t, "") if include_trace else None
    if not include_trace:
        rep = rep[OUTPUT_COLUMNS]
    else:
        rep = rep[OUTPUT_COLUMNS + TRACE_COLUMNS]

    rep = rep.reset_index(drop=True).sort_values(
        by=["年", "月", "店家編號", "品項"], na_position="last", kind="stable"
    ).reset_index(drop=True)
    rep.attrs = {}
    return rep


def build_store_reconciliation(detail, report_trace):
    """Source qty per (year, month, file outlet) vs report qty. Needs the trace report."""
    src = (
        detail.assign(_q=pd.to_numeric(detail["QTY"], errors="coerce"))
        .groupby(["YEAR", "MONTH", "FILE_OUTLET_NO"], dropna=False, as_index=False)
        .agg(SOURCE_ROWS=("_q", "size"), SOURCE_QTY=("_q", "sum"))
        .rename(columns={"YEAR": "年", "MONTH": "月", "FILE_OUTLET_NO": "店家編號"})
    )
    out = (
        report_trace.assign(_q=pd.to_numeric(report_trace["銷量(瓶)"], errors="coerce"),
                            店家編號_FILE=report_trace["FILE_OUTLET_NO"])
        .groupby(["年", "月", "店家編號_FILE"], dropna=False, as_index=False)
        .agg(店家名稱=("店家名稱", "first"), OUTPUT_ROWS=("_q", "size"), OUTPUT_QTY=("_q", "sum"))
        .rename(columns={"店家編號_FILE": "店家編號"})
    )
    for df in (src, out):
        df["年"] = df["年"].astype("Int64")
        df["月"] = df["月"].astype("Int64")
    rec = src.merge(out, how="outer", on=["年", "月", "店家編號"])
    for c in ["SOURCE_ROWS", "OUTPUT_ROWS", "SOURCE_QTY", "OUTPUT_QTY"]:
        rec[c] = pd.to_numeric(rec[c], errors="coerce").fillna(0)
    rec["QTY_DIFF"] = rec["OUTPUT_QTY"] - rec["SOURCE_QTY"]
    rec["MERGED_ROWS"] = rec["SOURCE_ROWS"] - rec["OUTPUT_ROWS"]
    rec["STATUS"] = rec["QTY_DIFF"].map(lambda x: "PASS" if abs(x) < 1e-6 else "CHECK")
    return rec


# ============================================================
# PIPELINE
# ============================================================

def _source_name(source) -> str:
    if isinstance(source, (str, Path)):
        return Path(source).name
    if isinstance(source, tuple) and source:
        return Path(str(source[0])).name
    return Path(str(getattr(source, "name", "unknown"))).name


def _submission_key(row) -> str:
    if row.get("YEAR") is not None and not pd.isna(row.get("YEAR")):
        period = f"{int(row['YEAR'])}{int(row['MONTH']):02d}"
    else:
        period = norm_text(row.get("FILE_PERIOD"))
    return f"{period}|{norm_id_key(row.get('FILE_OUTLET_NO'))}"


def run_transformation_pipeline(
    pos_files: Iterable,
    customer_mapping=None,
    sku_mapping=None,
    customer_mapping_sheet=None,
    sku_mapping_sheet=None,
    config: Optional[PipelineConfig] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    raise_on_all_failed: bool = False,
):
    """
    Batch pipeline. Returns dict:
      raw_all, detail, final_report, final_report_trace, upload_log,
      customer_duplicates, sku_duplicates, sku_upload_log, store_reconciliation,
      customer_map, sku_map, warnings, config
    A failing file never stops the batch; if ALL fail the upload_log is still returned (D5).
    """
    cfg = config or DEFAULT_CONFIG
    files = list(pos_files)
    frames, logs, warnings = [], [], []

    for idx, source in enumerate(files):
        t0 = time.perf_counter()
        name = _source_name(source)
        meta = parse_filename(name)
        row = {
            "檔案": name, "辨識格式": "", "Template Code": "ERROR", "狀態": "FAILED",
            "警告": "", "錯誤": "", "原始筆數": 0, "重複提交": "",
            "Detect 秒數": 0.0, "Transform 秒數": 0.0, "總秒數": 0.0,
            "業務代號": meta["FILE_SALES_ID"], "合約編號": meta["FILE_CONTRACT_JDE"],
            "年月": meta["FILE_PERIOD"], "店家編號": meta["FILE_OUTLET_NO"],
            "分店/Rawdata Name": meta["FILE_RAWDATA_NAME"], "檔名格式": meta["FILENAME_STATUS"],
        }
        t_detect = t_transform = 0.0
        try:
            src = SourceFile.from_any(source)
            td = time.perf_counter()
            code = detect_template_type(src)
            t_detect = time.perf_counter() - td
            row["Template Code"] = code
            if code == "UNKNOWN":
                raise TemplateError("Unable to detect template type（無法辨識 Template）")
            tt = time.perf_counter()
            data, label = auto_transform_template(src, code, cfg)
            t_transform = time.perf_counter() - tt
            notes = list(data.attrs.get("notes", []))
            data = data.copy()
            data.attrs = {}
            if meta["FILENAME_STATUS"] != "OK":
                notes.append("檔名格式需確認（業務代號_合約_年月_店家編號_店名）")
            if meta["FILE_COPY_SUFFIX"]:
                notes.append("檔名含重複下載後綴，如 (1)")
            data["FILE_INDEX"] = idx
            frames.append(data)
            row.update({"辨識格式": label, "狀態": "SUCCESS", "原始筆數": len(data),
                        "警告": " | ".join(notes), "_idx": idx})
        except Exception as exc:
            row["錯誤"] = f"{exc}" if isinstance(exc, ValueError) else f"{type(exc).__name__}: {exc}"
        row["Detect 秒數"] = round(t_detect, 3)
        row["Transform 秒數"] = round(t_transform, 3)
        row["總秒數"] = round(time.perf_counter() - t0, 3)
        logs.append(row)
        if progress_callback:
            progress_callback(idx + 1, len(files), name)

    upload_log = pd.DataFrame(logs)
    result = {
        "raw_all": None, "detail": None, "final_report": None, "final_report_trace": None,
        "upload_log": upload_log, "customer_duplicates": pd.DataFrame(),
        "sku_duplicates": pd.DataFrame(), "sku_upload_log": pd.DataFrame(),
        "store_reconciliation": pd.DataFrame(), "customer_map": None, "sku_map": None,
        "warnings": warnings, "config": cfg,
    }
    if not frames:
        msg = "所有 POS 檔案都轉換失敗"
        warnings.append(msg)
        if raise_on_all_failed:
            raise ValueError(msg)
        return result

    raw_all = pd.concat(frames, ignore_index=True, sort=False)

    # ---------- A3: duplicate submissions (same period + outlet, different files)
    file_keys = (
        raw_all.drop_duplicates("FILE_INDEX")[["FILE_INDEX", "YEAR", "MONTH", "FILE_PERIOD", "FILE_OUTLET_NO"]]
        .assign(KEY=lambda d: d.apply(_submission_key, axis=1))
        .sort_values("FILE_INDEX")
    )
    counts = file_keys.groupby("KEY")["FILE_INDEX"].transform("nunique")
    dup_idx = set(file_keys.loc[counts > 1, "FILE_INDEX"])
    raw_all["DUPLICATE_SUBMISSION"] = np.where(raw_all["FILE_INDEX"].isin(dup_idx), "DUPLICATE", "")
    excluded = set()
    if dup_idx and cfg.duplicate_policy in ("keep_first", "keep_last"):
        for _, g in file_keys[file_keys["FILE_INDEX"].isin(dup_idx)].groupby("KEY"):
            ordered = list(g["FILE_INDEX"])
            keep = ordered[0] if cfg.duplicate_policy == "keep_first" else ordered[-1]
            excluded |= {i for i in ordered if i != keep}
        raw_all = raw_all[~raw_all["FILE_INDEX"].isin(excluded)].reset_index(drop=True)

    if dup_idx:
        warnings.append(
            f"偵測到 {len(dup_idx)} 個檔案屬於重複提交（同月份 + 同店家編號）"
            + (f"；依 {cfg.duplicate_policy} 已排除 {len(excluded)} 個檔案" if excluded else "；目前僅標記，尚未排除")
        )
    ul = upload_log
    if "_idx" in ul.columns:
        for i in dup_idx:
            ul.loc[ul["_idx"] == i, "重複提交"] = "DUPLICATE"
        for i in excluded:
            m = ul["_idx"] == i
            ul.loc[m, "警告"] = ul.loc[m, "警告"].astype(str).str.cat(
                [f"已依重複提交政策排除({cfg.duplicate_policy})"], sep=" | ").str.strip(" |")
        ul = ul.drop(columns=["_idx"])
    result["upload_log"] = ul.reset_index(drop=True)

    result["raw_all"] = raw_all
    detail = raw_all.copy()

    if customer_mapping is not None:
        cmap, cdup = load_customer_mapping(customer_mapping, customer_mapping_sheet)
        detail = attach_customer_mapping(detail, cmap, cfg)
        result["customer_map"], result["customer_duplicates"] = cmap, cdup
        if not cdup.empty:
            warnings.append(f"Customer Mapping 有 {len(cdup)} 筆重複（Outlet No / Rawdata Name）")
    if sku_mapping is not None:
        smap, sdup, slog = load_sku_mapping(sku_mapping, sku_mapping_sheet)
        detail = attach_sku_mapping(detail, smap)
        result["sku_map"], result["sku_duplicates"], result["sku_upload_log"] = smap, sdup, slog
        if not sdup.empty and bool(sdup["CONFLICT"].any()):
            warnings.append("SKU Mapping 有重複 SKU Name 且屬性不一致（CONFLICT），目前使用第一筆")
    result["detail"] = detail

    if customer_mapping is not None and sku_mapping is not None:
        trace = build_report(detail, include_trace=True)
        result["final_report_trace"] = trace
        result["final_report"] = trace[OUTPUT_COLUMNS].copy()
        result["store_reconciliation"] = build_store_reconciliation(detail, trace)
    return result


# ============================================================
# SUBMISSION TRACKING
# ============================================================

def available_periods(upload_log: pd.DataFrame) -> List[str]:
    keys = upload_log["年月"].map(normalize_period_key)
    return sorted({k for k in keys if k}, reverse=True)


def build_submission_tracking(upload_log, customer_map, selected_period) -> dict:
    """Expected (Customer Mapping) vs actual uploads for one period (key = YYYYMM|Outlet)."""
    expected = customer_map[
        ["Sales ID", "Sales", "Contract JDE", "Contract NAME", "Outlet No", "Outlet NAME", "Rawdata Name"]
    ].copy()
    expected = expected[expected["Outlet No"].map(norm_id_key).ne("")].copy()
    expected["PERIOD_KEY"] = selected_period
    expected["OUTLET_TRACKING_KEY"] = expected["Outlet No"].map(norm_id_key)
    expected["RESPONSIBLE_SALES_ID"] = expected["Sales ID"].map(norm_id_key)
    expected["RESPONSIBLE_SALES"] = expected["Sales"].fillna("").map(norm_text)
    expected["EXPECTED_KEY"] = expected["PERIOD_KEY"] + "|" + expected["OUTLET_TRACKING_KEY"]
    expected = expected.drop_duplicates("EXPECTED_KEY", keep="first").reset_index(drop=True)

    actual = upload_log.copy()
    actual["PERIOD_KEY"] = actual["年月"].map(normalize_period_key)
    actual["OUTLET_TRACKING_KEY"] = actual["店家編號"].map(norm_id_key)
    actual["UPLOAD_KEY"] = actual["PERIOD_KEY"] + "|" + actual["OUTLET_TRACKING_KEY"]
    month = actual[actual["PERIOD_KEY"] == selected_period].copy()
    month["_P"] = month["狀態"].map({"SUCCESS": 2, "FAILED": 1}).fillna(0).astype(int)
    best = (
        month.sort_values(["UPLOAD_KEY", "_P"], ascending=[True, False])
        .drop_duplicates("UPLOAD_KEY", keep="first").drop(columns=["_P"]).reset_index(drop=True)
    )
    need = ["UPLOAD_KEY", "PERIOD_KEY", "檔案", "辨識格式", "狀態", "錯誤", "原始筆數",
            "年月", "檔名格式", "業務代號", "合約編號", "店家編號", "分店/Rawdata Name"]
    for c in need:
        if c not in best.columns:
            best[c] = pd.NA

    detail = expected.merge(best[need], how="left", left_on="EXPECTED_KEY", right_on="UPLOAD_KEY",
                            suffixes=("", "_UPLOAD"))
    detail["SUBMISSION_STATUS"] = np.where(
        detail["檔案"].isna(), "MISSING", np.where(detail["狀態"].eq("SUCCESS"), "SUCCESS", "FAILED"))

    matched = set(detail["UPLOAD_KEY"].dropna().astype(str))
    unmatched = best[~best["UPLOAD_KEY"].astype(str).isin(matched)].copy()

    sales = (
        detail.groupby(["Sales ID", "Sales"], dropna=False, as_index=False)
        .agg(應繳=("EXPECTED_KEY", "count"),
             成功=("SUBMISSION_STATUS", lambda x: int((x == "SUCCESS").sum())),
             失敗=("SUBMISSION_STATUS", lambda x: int((x == "FAILED").sum())),
             未繳=("SUBMISSION_STATUS", lambda x: int((x == "MISSING").sum())))
    )
    sales["完成率_RAW"] = (sales["成功"] / sales["應繳"]).fillna(0)
    sales["完成率"] = sales["完成率_RAW"].map(lambda x: f"{x:.1%}")
    if not sales.empty:
        sales = sales.sort_values(["完成率_RAW", "未繳"], ascending=[True, False]).reset_index(drop=True)

    n_exp = len(expected)
    n_ok = int((detail["SUBMISSION_STATUS"] == "SUCCESS").sum())
    return {
        "expected_submission": expected, "actual_upload_best": best,
        "submission_detail": detail, "sales_submission": sales, "unmatched_uploads": unmatched,
        "expected_count": n_exp, "success_count": n_ok,
        "failed_count": int((detail["SUBMISSION_STATUS"] == "FAILED").sum()),
        "missing_count": int((detail["SUBMISSION_STATUS"] == "MISSING").sum()),
        "unmatched_count": len(unmatched),
        "completion": (n_ok / n_exp) if n_exp else 0.0,
    }


# ============================================================
# EXPORT
# ============================================================

_ILLEGAL_EXCEL_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def clean_excel_value(value):
    try:
        if pd.isna(value):
            return value
    except (TypeError, ValueError):
        pass
    if isinstance(value, str):
        value = _ILLEGAL_EXCEL_RE.sub("", value)
        if len(value) > 32767:
            value = value[:32767]
    return value


def clean_dataframe_for_excel(df):
    if df is None:
        return pd.DataFrame()
    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame(df)
    out = df.copy()
    out.columns = [clean_excel_value(str(c)) for c in out.columns]
    for c in out.columns:
        if pd.api.types.is_object_dtype(out[c]) or pd.api.types.is_string_dtype(out[c]):
            out[c] = out[c].map(clean_excel_value)
    out.attrs = {}
    return out


def excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    used = set()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = re.sub(r"[:\\/?*\[\]]", "_", clean_excel_value(str(name)))[:31] or "Sheet"
            base, n = safe, 1
            while safe in used:
                n += 1
                safe = f"{base[:28]}_{n}"
            used.add(safe)
            clean_dataframe_for_excel(df).to_excel(writer, sheet_name=safe, index=False)
    return buf.getvalue()


AUDIT_EXPORT_COLUMNS = [
    "SOURCE_FILE", "SOURCE_SHEET", "SOURCE_ROW", "TEMPLATE_TYPE",
    "FILE_PERIOD", "FILE_SALES_ID", "FILE_SALES_ID_AUDIT", "FILE_CONTRACT_JDE",
    "FILE_OUTLET_NO", "FILE_RAWDATA_NAME", "DUPLICATE_SUBMISSION",
    "SUBMISSION_CUSTOMER_ID", "RESPONSIBLE_SALES_ID", "RESPONSIBLE_SALES",
    "RESPONSIBLE_SALES_STATUS", "CUSTOMER_OWNER_STATUS",
    "Sales ID", "Sales", "Contract JDE", "Contract NAME", "Outlet No", "Outlet NAME", "Rawdata Name",
    "CUSTOMER_MAPPING_STATUS", "CUSTOMER_MATCH_METHOD",
    "SALES_ID_CHECK", "OUTLET_CHECK", "CONTRACT_CHECK", "RAWDATA_NAME_CHECK",
    "RAW_ROW_CUSTOMER", "RAW_SKU", "RAW_BRAND", "RAW_SIZE", "RAW_CATEGORY",
    "RAW_QTY", "QTY_ORIGINAL", "QTY_UNIT", "UNIT_STATUS", "QTY", "QTY_FLAG",
    "SKU_MAPPING_STATUS", "RSP_STATUS", "Mapping Name", "RSP", "Price Band",
    "Manufacture", "Brand", "SKU Name", "SIZE", "CATEGORY",
    "RAW_UNIT_PRICE", "RAW_TOTAL_PRICE", "SOURCE_QTY_COLUMN", "SOURCE_PRICE_COLUMN",
]


def existing_columns(df, columns):
    return [c for c in columns if c in df.columns]


def build_export_frames(result: dict, submission: Optional[dict] = None) -> dict:
    """All DataFrames used by the Excel exports (full workbook + unmapped workbook)."""
    detail = result["detail"]
    cols = existing_columns(detail, AUDIT_EXPORT_COLUMNS)
    audit = detail[cols].copy().reset_index(drop=True)

    cust_un = detail["CUSTOMER_MAPPING_STATUS"].eq("UNMAPPED")
    sku_un = detail["SKU_MAPPING_STATUS"].ne("MAPPED")  # UNMAPPED + INCOMPLETE
    sub = lambda m: detail.loc[m, cols].copy().reset_index(drop=True)

    own_cols = existing_columns(detail, [
        "SOURCE_FILE", "FILE_PERIOD", "FILE_SALES_ID_AUDIT", "FILE_OUTLET_NO", "FILE_CONTRACT_JDE",
        "FILE_RAWDATA_NAME", "Outlet No", "Outlet NAME", "RESPONSIBLE_SALES_ID", "RESPONSIBLE_SALES",
        "CUSTOMER_MAPPING_STATUS", "CUSTOMER_MATCH_METHOD", "RESPONSIBLE_SALES_STATUS",
        "CUSTOMER_OWNER_STATUS", "SALES_ID_CHECK", "OUTLET_CHECK", "CONTRACT_CHECK", "RAWDATA_NAME_CHECK"])
    ownership = detail[own_cols].drop_duplicates("SOURCE_FILE").reset_index(drop=True).rename(columns={
        "SOURCE_FILE": "檔案", "FILE_PERIOD": "年月", "FILE_SALES_ID_AUDIT": "檔名業務代號",
        "FILE_OUTLET_NO": "檔名Customer ID", "FILE_CONTRACT_JDE": "檔名合約編號",
        "FILE_RAWDATA_NAME": "檔名店家名稱", "Outlet No": "Mapping Customer ID",
        "Outlet NAME": "Mapping 店家名稱", "RESPONSIBLE_SALES_ID": "負責業務代號",
        "RESPONSIBLE_SALES": "負責業務", "CUSTOMER_MAPPING_STATUS": "Customer Mapping",
        "CUSTOMER_MATCH_METHOD": "配對方式", "RESPONSIBLE_SALES_STATUS": "負責業務 Mapping",
        "CUSTOMER_OWNER_STATUS": "業務歸屬狀態", "SALES_ID_CHECK": "Sales ID Check",
        "OUTLET_CHECK": "Customer ID Check", "CONTRACT_CHECK": "Contract Check",
        "RAWDATA_NAME_CHECK": "店名 Check"})
    owner_un = ownership[ownership["業務歸屬狀態"].eq("UNMAPPED")].reset_index(drop=True)

    frames = {
        "all_unmapped": sub(cust_un | sku_un), "customer_unmapped": sub(cust_un),
        "sku_unmapped": sub(sku_un), "both_unmapped": sub(cust_un & sku_un),
        "audit": audit, "ownership": ownership, "owner_unmapped": owner_un,
    }
    s = submission or {}
    sheets = {
        "Final Report": result["final_report"],
        "Final Report (Trace)": result["final_report_trace"],
        "Upload Log": result["upload_log"],
        "Submission Tracker": s.get("submission_detail", pd.DataFrame()),
        "Sales Submission": s.get("sales_submission", pd.DataFrame()),
        "All Unmapped": frames["all_unmapped"], "Customer Unmapped": frames["customer_unmapped"],
        "SKU Unmapped": frames["sku_unmapped"], "Both Unmapped": frames["both_unmapped"],
        "Audit Detail": audit, "Sales Ownership Audit": ownership, "Owner Unmapped": owner_un,
        "Unmatched Upload": s.get("unmatched_uploads", pd.DataFrame()),
        "Customer Duplicates": result["customer_duplicates"],
        "SKU Duplicates": result["sku_duplicates"],
        "Store Reconciliation": result["store_reconciliation"],
    }
    required = {"Final Report", "Upload Log"}
    frames["workbook_sheets"] = {
        k: v for k, v in sheets.items()
        if k in required or (isinstance(v, pd.DataFrame) and not v.empty)
    }
    un_sheets = {
        "All Unmapped": frames["all_unmapped"], "Customer Unmapped": frames["customer_unmapped"],
        "SKU Unmapped": frames["sku_unmapped"], "Both Unmapped": frames["both_unmapped"],
        "Owner Unmapped": owner_un,
    }
    un_sheets = {k: v for k, v in un_sheets.items() if not v.empty}
    frames["unmapped_workbook_sheets"] = un_sheets or {
        "Unmapped": pd.DataFrame({"Status": ["No unmapped records"]})}
    return frames


# ============================================================
# ZIP / FOLDER LOADERS
# ============================================================

SUPPORTED_SUFFIXES = (".xlsx", ".xls", ".csv")


def extract_pos_zip(data: bytes):
    """ZIP bytes -> ([(filename, bytes)], [ignored names])."""
    files, ignored = [], []
    with zipfile.ZipFile(io.BytesIO(data), "r") as z:
        for info in z.infolist():
            fname = info.filename
            if info.is_dir() or "__MACOSX" in fname or Path(fname).name.startswith((".", "~$")):
                continue
            if Path(fname).suffix.lower() not in SUPPORTED_SUFFIXES:
                ignored.append(fname)
                continue
            # ZIPs made on Windows often store non-UTF-8 names as cp437 -> try to repair
            name = Path(fname).name
            try:
                if not (info.flag_bits & 0x800):
                    name = fname.encode("cp437").decode("cp950")
                    name = Path(name).name
            except Exception:
                name = Path(fname).name
            files.append((name, z.read(info)))
    return files, ignored


def collect_pos_inputs(path) -> list:
    """Folder or .zip -> list of (filename, bytes)."""
    p = Path(path)
    if p.is_dir():
        return [(f.name, f.read_bytes()) for f in sorted(p.iterdir())
                if f.suffix.lower() in SUPPORTED_SUFFIXES and not f.name.startswith(("~$", "."))]
    if p.suffix.lower() == ".zip":
        return extract_pos_zip(p.read_bytes())[0]
    return [(p.name, p.read_bytes())]


# ============================================================
# CLI
# ============================================================

def _main():
    import argparse

    ap = argparse.ArgumentParser(description="POS monthly transformation (no Streamlit)")
    ap.add_argument("--pos", required=True, help="folder, .zip or single POS file")
    ap.add_argument("--customer", required=True, help="Customer Mapping file")
    ap.add_argument("--sku", required=True, nargs="+", help="SKU Mapping file(s)")
    ap.add_argument("--out", default="POS_Monthly_Report.xlsx")
    ap.add_argument("--customer-sheet", default="Customer data")
    ap.add_argument("--sku-sheet", default="SKU data")
    ap.add_argument("--duplicate-policy", default="flag", choices=["flag", "keep_first", "keep_last"])
    ap.add_argument("--period", default=None, help="YYYYMM for submission tracking (default: latest)")
    a = ap.parse_args()

    cfg = PipelineConfig(duplicate_policy=a.duplicate_policy)
    res = run_transformation_pipeline(
        collect_pos_inputs(a.pos), a.customer, list(a.sku),
        customer_mapping_sheet=a.customer_sheet, sku_mapping_sheet=a.sku_sheet, config=cfg,
        progress_callback=lambda i, n, name: print(f"[{i}/{n}] {name}"),
    )
    print(res["upload_log"][["檔案", "Template Code", "狀態", "原始筆數", "警告", "錯誤"]].to_string())
    for w in res["warnings"]:
        print("WARNING:", w)
    if res["final_report"] is None:
        raise SystemExit("No final report produced.")
    periods = available_periods(res["upload_log"])
    sub = None
    if periods:
        sub = build_submission_tracking(res["upload_log"], res["customer_map"], a.period or periods[0])
    frames = build_export_frames(res, sub)
    Path(a.out).write_bytes(excel_bytes(frames["workbook_sheets"]))
    print("Saved:", a.out, "rows:", len(res["final_report"]))


if __name__ == "__main__":
    _main()
