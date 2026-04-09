#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PL Scraper / Filler (single-file, no external file dependencies)

MAX_ABS_MILLION_SANITY = 1e9  # outlier guard

Purpose
- Update the input workbook by filling ONLY blank cells (no overwrite).
- If an existing value differs from the best source value by > ±2 (million JPY units),
  mark the cell ORANGE as WARN (do not overwrite).
- Newly filled cells are colored YELLOW.
- If any ORANGE cell exists in the latest 5 quarters (FY2024Q4..FY2025Q4 => q_idx 27..31)
  across any target block, set EY=1 for that row.

Source priority (strict)
1) EDINET
2) TDNet LIVE (https://www.release.tdnet.info/inbs/I_main_00.html) (direct, zip attachments)
3) Kabutan

Important constraints
- Left blocks = single-quarter values.
- Many filings provide cumulative (YTD) values:
    Q1 = as-is
    Q2+ = current cumulative - previous cumulative (within same FY, same source)
- Right blocks (EZ:KE) hold cumulative values as-is (never derived from left blocks).
- Do NOT back-calculate right cumulative blocks from left single blocks.

Units
- Values are stored in "million JPY (百万円)" in the workbook.
- If a filing discloses in thousand JPY (千円), convert to million JPY by dividing by 1000.
  (Keep decimals; do NOT integer-round.)

Workbook (expected)
- Data starts at row 4.
- Columns (left, single):
  * 経常利益: S:AX (2018Q1..2025Q4)
  * 売上高合計: AY:CD (2018Q1..2025Q4)
  * 親会社株主に帰属する当期純利益: CE:DJ (2018Q1..2025Q4)
  * 減価償却費: DK:DV (2023Q1..2025Q4)
  * 売上総利益: DW:EL (2022Q1..2025Q4)
  * 販売費及び一般管理費: EM:EX (2023Q1..2025Q4)
- Column EY: WARN flag
- Right cumulative blocks (created if absent):
  * 経常利益: EZ:GE (2018Q1..2025Q4)
  * 売上高合計: GF:HK (2018Q1..2025Q4)
  * 親会社株主に帰属する当期純利益: HL:IQ (2018Q1..2025Q4)
  * 減価償却費: IR:JC (2023Q1..2025Q4)
  * 売上総利益: JD:JS (2022Q1..2025Q4)
  * 販売費及び一般管理費: JT:KE (2023Q1..2025Q4)

Changelog
- v5_38: Do NOT suppress WARN when the best candidate value is 'derived' (e.g., cross-source cumulative).

CLI
- --input / --output
- --limit for smoke test
- --log-csv to write action log + summary
"""

from __future__ import annotations

# R7: simplify naming/version (file name avoids multiple "v" tokens)
__version__ = "r35"


def _preflight_runtime_checks() -> None:
    """Fail-fast checks for internal NameError regressions.

    This runs before any scraping so we don't waste minutes before crashing.
    It is side-effect free.
    """
    # Ensure priority maps exist
    _ = SRC_PRIORITY.get('edinet', 0)
    _ = SOURCE_PRIORITY.get('edinet', 0)
    # Exercise merge path that previously raised NameError
    try:
        import datetime as _dt
        dummy = FilingRecord(
            ticker='0000', source='edinet', filing_date=_dt.date(2000, 1, 1), title='preflight',
            fy_end_year=2000, fy_end_month=3, quarter_no=1,
            metric_cum={'keijo': 0.0}, metric_single={'keijo': 0.0}, metric_instant={'sisan_total': 0.0},
        )
        merge_records_with_instant([dummy])
        # Exercise meta_set API (both preferred and legacy call styles)
        import openpyxl as _ox
        from types import SimpleNamespace as _SN
        _wb = _ox.Workbook()
        _meta_sh, _meta_row_map = ensure_meta_sheet(_wb)
        _key = meta_make_key('0000', 'keijo', 'cum', 0, 'A')
        _vp = _SN(value=1.0, source='edinet', priority=3, filing_date=_dt.date(2000, 1, 1), title='preflight', derived=False)
        meta_set(_meta_sh, _meta_row_map, _key, '0000', 'keijo', 'cum', 0, 'A', _vp)
        meta_set(_meta_sh, _meta_row_map, _key, src='edinet', priority=3, filing_date=str(_dt.date(2000, 1, 1)), title='preflight', derived=False, value=1.0)

    except Exception as e:
        raise RuntimeError(f'[INTERNAL BUG] preflight failed: {e!r}')



def fy_start_month_from_fy_end_month(fy_end_month: int) -> int:
    """Return fiscal-year start month (1-12) given fiscal-year end month (1-12).
    Example: fy_end=3 => start=4, fy_end=12 => start=1.
    """
    try:
        m = int(fy_end_month)
    except Exception:
        return 1
    if m < 1 or m > 12:
        return 1
    return (m % 12) + 1

import argparse
import csv
import datetime as dt
import io
import json
import os
import re
import sys
import time
import math
import calendar
import traceback
import zipfile
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Any, Iterable, Set

import requests

# Optional: use curl_cffi to bypass WAF/bot-detection for TDNet/Kabutan when needed
try:
    from curl_cffi import requests as cffi_requests  # type: ignore
except Exception:
    cffi_requests = None
from bs4 import BeautifulSoup
from bs4 import FeatureNotFound
import warnings
from bs4 import XMLParsedAsHTMLWarning
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string

# Runtime constant: outlier guard for iXBRL numeric conversion (million JPY).
# NOTE: In r12 this existed only inside the module docstring, causing NameError at runtime.

# Script revision tag (for user troubleshooting)
__REV__ = 'r24'

MAX_ABS_MILLION_SANITY = float(os.environ.get('MAX_ABS_MILLION_SANITY', '1e9'))


def read_dotenv_value(name: str, *, search_dirs: Optional[List[Path]] = None) -> str:
    """Best-effort `.env` reader for local execution.

    This script is often launched from notebooks/subprocesses where process env and
    the workspace `.env` can drift. We only need a tiny subset here, so avoid
    requiring python-dotenv.
    """
    dirs = search_dirs or [Path.cwd(), Path(__file__).resolve().parent]
    seen: Set[str] = set()
    for d in dirs:
        try:
            env_path = (d / ".env").resolve()
        except Exception:
            continue
        key = str(env_path).lower()
        if key in seen or not env_path.exists():
            continue
        seen.add(key)
        try:
            for raw in env_path.read_text(encoding="utf-8").splitlines():
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                if k.strip() != name:
                    continue
                return v.strip().strip('"').strip("'")
        except Exception:
            continue
    return ""


def edinet_api_key_candidates(preferred_key: str) -> List[str]:
    """Return unique non-empty EDINET API key candidates in priority order."""
    vals = [
        (preferred_key or "").strip(),
        (os.environ.get("EDINET_API_KEY", "") or "").strip(),
        read_dotenv_value("EDINET_API_KEY"),
    ]
    out: List[str] = []
    seen: Set[str] = set()
    for v in vals:
        if not v or v in seen:
            continue
        seen.add(v)
        out.append(v)
    return out




# -----------------------
# Byte decoding helpers
# -----------------------
def decode_bytes_guess(b: bytes) -> str:
    """Decode bytes into text with a robust encoding guess.

    TDNet/EDINET ZIPs may contain XBRL (utf-8) and/or iXBRL HTML (often CP932).
    We avoid hard dependency on chardet; try common encodings first, then
    fall back to charset-normalizer (bundled with requests) if available.
    """
    if b is None:
        return ""
    # Fast paths
    for enc in (
        "utf-8-sig",
        "utf-8",
        "cp932",
        "shift_jis",
        "euc_jp",
        "iso2022_jp",
        "utf-16",
        "utf-16le",
        "utf-16be",
    ):
        try:
            return b.decode(enc)
        except Exception:
            continue
    # Best-effort via charset_normalizer if present
    try:
        from charset_normalizer import from_bytes  # type: ignore
        m = from_bytes(b).best()
        if m is not None:
            return str(m)
    except Exception:
        pass
    # Last resort (keeps going; caller should handle partial text)
    return b.decode("utf-8", errors="replace")

# -----------------------
# Constants / Styles
# -----------------------
JST = dt.timezone(dt.timedelta(hours=9))
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

TIMEOUT = 25


# Fixed lookback windows (no disable switches)
EDINET_LOOKBACK_DAYS = 365
TDNET_LOOKBACK_DAYS = 180

# TDNet GitHub archive (user-managed mirror for older TDNet data)
# - TDNet LIVE (release.tdnet.info) is often only searchable for a limited retention window.
# - This script can optionally use a GitHub repository as an archive for TDNet XBRL ZIP/PDF.
#   (We only parse XBRL ZIPs; PDF parsing is not supported.)
TDNET_GH_DEFAULT_OWNER = "yukizi1113"
TDNET_GH_DEFAULT_REPO = "tdnet"
TDNET_GH_DEFAULT_BRANCH = "main"
TDNET_GH_DEFAULT_XBRL_DIR = "XBRL"
TDNET_GH_DEFAULT_PDF_DIR = "tekigikaizi"
TDNET_GH_DEFAULT_FROM_DATE = dt.date(2025, 12, 17)  # inclusive
YELLOW_FILL = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
ORANGE_FILL = PatternFill(fill_type="solid", start_color="F4B183", end_color="F4B183")

BASE_FY_START_YEAR = 2018
MAX_Q_IDX = 31  # FY2018Q1..FY2025Q4

WARN_COL = "EY"
RECENT5_QIDX = set([27, 28, 29, 30, 31])  # FY2024Q4..FY2025Q4


METRIC_JA = {
    "keijo": "経常利益",
    "uriage": "売上高合計",
    "saishu": "親会社株主に帰属する当期純利益",
    "genka": "減価償却費",
    "gross": "売上総利益",
    "sga": "販売費及び一般管理費",
    "opcf": "営業活動によるキャッシュ・フロー",
    "ad": "広告宣伝費",
    "rnd": "研究開発費",
    "capex_ppe": "有形固定資産の取得（投資CF）",
}

# Left single blocks (fixed by workbook)
LEFT_BLOCKS = {
    "keijo": ("S", "AX"),      # 32, FY2018Q1..FY2025Q4
    "uriage": ("AY", "CD"),    # 32
    "saishu": ("CE", "DJ"),    # 32
    "genka": ("DK", "DV"),     # 12, FY2023Q1..FY2025Q4
    "gross": ("DW", "EL"),     # 16, FY2022Q1..FY2025Q4
    "sga": ("EM", "EX"),       # 12, FY2023Q1..FY2025Q4
}

# Right cumulative blocks (fixed by spec; may not exist in workbook yet)
RIGHT_BLOCKS = {
    "keijo": ("EZ", "GE"),     # 32
    "uriage": ("GF", "HK"),    # 32
    "saishu": ("HL", "IQ"),    # 32
    "genka": ("IR", "JC"),     # 12
    "gross": ("JD", "JS"),     # 16
    "sga": ("JT", "KE"),       # 12
}

# Quarter index start for partial-range metrics
QSTART = {
    "keijo": 0,
    "uriage": 0,
    "saishu": 0,
    "gross": 16,  # FY2022Q1
    "genka": 20,  # FY2023Q1
    "sga": 20,    # FY2023Q1
}



# Metrics that must never be negative (both single and cumulative)
NONNEG_METRICS = {"uriage", "sga", "genka"}
SRC_PRIORITY = {"edinet": 3, "tdnet": 2, "kabutan": 1}
SOURCE_PRIORITY = SRC_PRIORITY  # alias (R8)
DERIVE_CUM_FROM_SINGLE = False  # derive cumulative series from single-quarter values (source-internal)
DERIVE_CUM_FROM_SINGLE_SOURCES = {"kabutan"}  # Kabutan quarterly table is single-quarter; build FY cumulative for right blocks

# v5_41: allow higher-priority derived single (from cumulative differences) to override lower-priority direct single.
# This fixes false ORANGE cases like 3776/3904 where Kabutan direct single blocked EDINET-derived single.
SINGLE_DERIVE_OVERRIDE_DIRECT_IF_HIGHER_PRIORITY = True


# -----------------------
# Meta sheet (provenance) for right cumulative blocks
# -----------------------
# We persist the source/priority used to populate each RIGHT_BLOCKS cell so that:
#   - A later run does NOT WARN based on a lower-priority source when a higher-priority value
#     already exists in the sheet but can no longer be fetched (e.g., TDNet 30-day retention).
#   - In the future, RIGHT_BLOCKS may be updated (overwritten) ONLY when the new source has
#     higher priority, and never by Kabutan.
META_SHEET_NAME = "__pl_meta"
META_HEADER = [
    "key", "ticker", "metric", "kind", "q_idx", "col",
    "src", "priority", "derived", "filing_date", "title", "value", "updated_at"
]


def meta_make_key(ticker: str, metric: str, kind: str, qidx: int, col: str) -> str:
    return f"{ticker}|{metric}|{kind}|{int(qidx)}|{col}"


def ensure_meta_sheet(wb):
    # Create/load a hidden sheet storing per-cell provenance.
    if META_SHEET_NAME in wb.sheetnames:
        sh = wb[META_SHEET_NAME]
    else:
        sh = wb.create_sheet(META_SHEET_NAME)
        try:
            sh.sheet_state = "hidden"
        except Exception:
            pass
        sh.append(META_HEADER)

    # Build key -> row index map.
    # NOTE (r28):
    # - openpyxl's Worksheet.max_row can become O(n) on huge sheets because it scans internal cell dict.
    # - We call it ONCE here to seed an append pointer, and never call it again per-cell.
    row_map = {}
    try:
        maxr = sh.max_row
    except Exception:
        maxr = 1
    for rr in range(2, maxr + 1):
        k = sh.cell(rr, 1).value
        if k is None or str(k).strip() == "":
            continue
        row_map[str(k)] = rr

    # Next row pointer for appends (avoid repeated max_row calls)
    row_map["__next_row__"] = int(maxr) + 1
    # If meta writing triggers MemoryError, we disable it and continue filling values.
    row_map["__disabled__"] = False
    return sh, row_map


def meta_get(meta_sh, meta_row_map: dict, key: str):
    rr = meta_row_map.get(key)
    if not rr:
        return None
    def gv(i):
        return meta_sh.cell(rr, i).value
    return {
        "src": gv(7),
        "priority": gv(8),
        "derived": gv(9),
        "filing_date": gv(10),
        "title": gv(11),
        "value": gv(12),
        "updated_at": gv(13),
    }


def meta_set(meta_sh, meta_row_map: dict, key: str,
             ticker: str = None, metric: str = None, kind: str = None,
             qidx: int = None, col: str = None, vp=None, **kwargs):
    """Write provenance metadata for a cell.

    Backward/forward compatible:
      - Preferred: meta_set(meta_sh, meta_row_map, key, ticker, metric, kind, qidx, col, vp)
      - Legacy:   meta_set(meta_sh, meta_row_map, key, src=..., priority=..., filing_date=..., title=..., derived=..., value=...)
    """
    # Parse components from key when omitted
    if (ticker is None or metric is None or kind is None or qidx is None or col is None) and isinstance(key, str) and '|' in key:
        parts = key.split('|')
        if len(parts) >= 5:
            ticker = ticker or parts[0]
            metric = metric or parts[1]
            kind = kind or parts[2]
            try:
                qidx = int(qidx if qidx is not None else parts[3])
            except Exception:
                qidx = int(qidx or 0)
            col = col or parts[4]

    # Build a minimal VP from kwargs when vp is not provided
    if vp is None and kwargs:
        from types import SimpleNamespace
        vp = SimpleNamespace(
            value=float(kwargs.get('value', 0.0) or 0.0),
            source=kwargs.get('src', None) or kwargs.get('source', None),
            priority=int(kwargs.get('priority', 0) or 0),
            filing_date=kwargs.get('filing_date', None),
            title=kwargs.get('title', None),
            derived=bool(kwargs.get('derived', False)),
        )

    # Upsert
    # If meta logging is disabled (e.g., after MemoryError), keep running without meta.
    if meta_row_map.get("__disabled__"):
        return

    try:
        rr = meta_row_map.get(key)
        if not rr:
            rr = meta_row_map.get("__next_row__")
            if rr is None:
                rr = 2
            meta_row_map["__next_row__"] = int(rr) + 1

            meta_row_map[key] = rr
            meta_sh.cell(rr, 1).value = key
            meta_sh.cell(rr, 2).value = ticker
            meta_sh.cell(rr, 3).value = metric
            meta_sh.cell(rr, 4).value = kind
            meta_sh.cell(rr, 5).value = int(qidx or 0)
            meta_sh.cell(rr, 6).value = col
    except MemoryError:
        # r28: meta is optional; never abort the entire scrape due to provenance logging.
        meta_row_map["__disabled__"] = True
        return

    try:
        meta_sh.cell(rr, 7).value = getattr(vp, 'source', None) if vp is not None else None
        meta_sh.cell(rr, 8).value = int(getattr(vp, 'priority', 0) or 0) if vp is not None else 0
        meta_sh.cell(rr, 9).value = bool(getattr(vp, 'derived', False)) if vp is not None else False
        meta_sh.cell(rr, 10).value = getattr(vp, 'filing_date', None) if vp is not None else None
        meta_sh.cell(rr, 11).value = getattr(vp, 'title', None) if vp is not None else None
        meta_sh.cell(rr, 12).value = float(getattr(vp, 'value', 0.0) or 0.0) if vp is not None else 0.0
        import datetime as _dt
        meta_sh.cell(rr, 13).value = _dt.datetime.now().isoformat(timespec='seconds')
    except MemoryError:
        meta_row_map["__disabled__"] = True
        return
    except Exception:
        try:
            meta_sh.cell(rr, 13).value = ""
        except Exception:
            pass
# -----------------------

# Data structures
# -----------------------
@dataclass
class FilingRecord:
    ticker: str
    source: str               # edinet/tdnet/kabutan
    filing_date: Optional[dt.date]
    title: str
    fy_end_year: Optional[int]   # e.g. 2026 for "2026年3月期"
    fy_end_month: Optional[int]  # 3 for "3月期"
    quarter_no: Optional[int]    # 1..4
    metric_cum: Dict[str, float]   # million JPY (百万円)
    metric_single: Dict[str, float]  # million JPY (百万円)
    metric_instant: Dict[str, float] = field(default_factory=dict)  # instant values (e.g., assets) in million JPY (百万円)

@dataclass
class ValuePoint:
    value: float
    source: str
    priority: int
    filing_date: Optional[dt.date]
    title: str
    derived: bool = False


def better_point(cur: Optional[ValuePoint], cand: Optional[ValuePoint]) -> Optional[ValuePoint]:
    """Choose better ValuePoint among two candidates.
    Priority order:
      1) higher priority (EDINET > TDNet > Kabutan)
      2) non-derived beats derived
      3) newer filing_date beats older (if available)
      4) otherwise keep existing (stable)
    """
    if cur is None:
        return cand
    if cand is None:
        return cur
    if cand.priority != cur.priority:
        return cand if cand.priority > cur.priority else cur
    # prefer non-derived
    if cand.derived != cur.derived:
        return cand if (not cand.derived) and cur.derived else cur
    # prefer newer filing_date when both present
    if cand.filing_date and cur.filing_date and cand.filing_date != cur.filing_date:
        return cand if cand.filing_date > cur.filing_date else cur
    if cand.filing_date and not cur.filing_date:
        return cand
    # tie-breaker: keep cur
    return cur



def better_point_for_single(cur: Optional[ValuePoint], cand: Optional[ValuePoint]) -> Optional[ValuePoint]:
    """Choose better ValuePoint for SINGLE-quarter values.

    v5_31 policy:
      - Prefer NON-derived (direct) single-quarter values over derived values.
        (User decision: Kabutan direct 3m actual > (cum diff) derived single.)
      - Then follow normal source priority (EDINET > TDNet > Kabutan) and recency when available.
    """
    if cur is None:
        return cand
    if cand is None:
        return cur
    if cand.priority != cur.priority:
        return cand if cand.priority > cur.priority else cur
    # prefer non-derived
    if cand.derived != cur.derived:
        return cand if (not cand.derived) and cur.derived else cur
    # prefer newer filing_date when both present
    if cand.filing_date and cur.filing_date and cand.filing_date != cur.filing_date:
        return cand if cand.filing_date > cur.filing_date else cur
    if cand.filing_date and not cur.filing_date:
        return cand
    return cur
def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)

def iter_cols(c1: str, c2: str) -> List[str]:
    i1 = column_index_from_string(c1)
    i2 = column_index_from_string(c2)
    return [get_column_letter(i) for i in range(i1, i2 + 1)]

def normalize_ticker(v: Any) -> Optional[str]:
    """Normalize ticker/securities code from Excel/strings.

    Supports:
    - Traditional JP stock codes: 4 digits (e.g., 9760)
    - New alphanumeric codes: 3 digits + 1 letter (e.g., 142A)
      (Also tolerates 4 digits + 1 letter in case it appears.)
    - Excel numeric artifacts like 9760.0
    """
    if v is None:
        return None
    s = str(v).strip().upper()
    if not s:
        return None

    # Excel numeric like 9760.0
    if re.fullmatch(r"\d+\.0", s):
        s = s.split(".")[0]

    # Remove spaces / full-width spaces
    s = re.sub(r"[\s\u3000]+", "", s)

    # Exact formats
    if re.fullmatch(r"\d{4}", s):
        return s
    if re.fullmatch(r"\d{3}[A-Z0-9]", s):
        return s
    if re.fullmatch(r"\d{4}[A-Z]", s):
        return s

    # Fallback: find first plausible code in text
    m = re.search(r"(\d{4}|\d{3}[A-Z0-9])", s)
    return m.group(1) if m else None

def normalize_security_code(v: Any) -> Optional[str]:
    """Normalize a securities code found inside an XBRL ZIP.

    TDNet file names often embed codes like '142A0' or '97600' (trailing 0).
    This returns the 4-char code without the trailing 0 when applicable.
    """
    if v is None:
        return None
    s = str(v).strip().upper()
    if not s:
        return None
    s = re.sub(r"[\s\u3000]+", "", s)

    # Common TDNet embedding: 142A0 / 97600
    m = re.fullmatch(r"(\d{3}[A-Z0-9]|\d{4})0", s)
    if m:
        return m.group(1)

    # Plain code
    if re.fullmatch(r"\d{4}", s):
        return s
    if re.fullmatch(r"\d{3}[A-Z0-9]", s):
        return s
    if re.fullmatch(r"\d{4}[A-Z]", s):
        return s

    # Fallback
    m2 = re.search(r"(\d{4}|\d{3}[A-Z0-9])0?", s)
    return normalize_security_code(m2.group(1)) if m2 else None



def read_ticker_whitelist(args) -> Optional[Set[str]]:
    """Return a set of whitelisted tickers if --tickers / --tickers-file is provided, else None."""
    tickers: Set[str] = set()

    s = getattr(args, "tickers", "") or ""
    if s.strip():
        # split by comma/space/newline
        for part in re.split(r"[,\s]+", s.strip()):
            if not part:
                continue
            t = normalize_ticker(part)
            if t:
                tickers.add(t)

    fp = getattr(args, "tickers_file", "") or ""
    if fp.strip():
        try:
            with open(fp, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    # allow CSV first column
                    part = line.split(",")[0].strip()
                    t = normalize_ticker(part)
                    if t:
                        tickers.add(t)
        except Exception as e:
            raise RuntimeError(f"Failed to read --tickers-file: {fp}: {e}")

    return tickers if tickers else None

def securities_code_matches_expected(code_in_zip: Any, expected_ticker: Any) -> bool:
    """Return True if a code found in ZIP matches expected ticker (supports alphanumeric tickers)."""
    a = normalize_security_code(code_in_zip)
    b = normalize_ticker(expected_ticker)
    if not a or not b:
        return False
    return a == b


def parse_num(v: Any) -> Optional[float]:
    """Parse Japanese-style numeric strings to float. Returns None if blank."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return None
        return float(v)
    s = str(v).strip()
    # Excel formulas (strings starting with '=') should be treated as blank; we can't evaluate them here.
    if s.startswith('='):
        return None
    if not s:
        return None
    if s in {"-", "—", "―", "－", "ー", "N/A", "na", "NA"}:
        return None
    neg = False
    if s.startswith(("△", "▲")):
        neg = True
        s = s[1:]
    # parentheses negative
    if "(" in s and ")" in s:
        neg = True
        s = s.replace("(", "").replace(")", "")
    s = s.replace(",", "").replace(" ", "").replace("　", "")
    s = s.replace("−", "-").replace("－", "-")
    if s.startswith("-"):
        neg = True
        s = s[1:]
    s = re.sub(r"[^\d.]", "", s)
    if not s:
        return None
    try:
        x = float(s)
        return -x if neg else x
    except Exception:
        return None


def parse_num_first_token(v: Any) -> Optional[float]:
    """Parse the FIRST numeric token from a string (safer for iXBRL get_text noise).

    This is only used as a fallback when the normal parse_num() yields an implausible huge value.
    """
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return parse_num(v)
    s = str(v).strip()
    if s.startswith('='):
        return None
    if not s:
        return None
    if s in {"-", "—", "―", "－", "ー", "N/A", "na", "NA"}:
        return None
    neg = False
    if s.startswith(("△", "▲")):
        neg = True
        s = s[1:]
    # parentheses negative
    if "(" in s and ")" in s:
        neg = True
        s = s.replace("(", "").replace(")", "")
    s = s.replace(",", "").replace(" ", "").replace("　", "")
    s = s.replace("−", "-").replace("－", "-")
    if s.startswith("-"):
        neg = True
        s = s[1:]
    m = re.search(r"\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        x = float(m.group(0))
        return -x if neg else x
    except Exception:
        return None

def managed_fill_rgb(cell) -> str:
    try:
        rgb = (cell.fill.start_color.rgb or "").upper()
    except Exception:
        rgb = ""
    return rgb

def is_managed_fill(cell) -> bool:
    rgb = managed_fill_rgb(cell)
    return (
        cell.fill is not None and cell.fill.fill_type == "solid" and
        (rgb.endswith("FFF2CC") or rgb.endswith("F4B183"))
    )


def mark_managed_fill(cell, fill) -> None:
    """Assign a managed fill (yellow/orange)."""
    try:
        cell.fill = fill
    except Exception:
        pass



def clear_managed_fill(cell) -> None:
    if is_managed_fill(cell):
        cell.fill = PatternFill(fill_type=None)

def ensure_requests_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    return s

def make_soup(html_text: str) -> BeautifulSoup:
    """Robust BeautifulSoup parser selection.

    Prefer lxml if available (faster, more tolerant). Falls back to html.parser.
    """
    if not html_text:
        return BeautifulSoup("", "html.parser")
    try:
        return BeautifulSoup(html_text, "lxml")
    except FeatureNotFound:
        return BeautifulSoup(html_text, "html.parser")

HTTP_FAILS: List[Tuple[str, str]] = []
HTTP_FAIL_LIMIT = 200
HTTP_FAIL_COUNT = 0

class SourceAccessError(RuntimeError):
    def __init__(self, source: str, url: str, detail: str = ""):
        msg = f"Required data source '{source}' is not accessible: {url}"
        if detail:
            msg += f" ({detail})"
        super().__init__(msg)
        self.source = source
        self.url = url
        self.detail = detail


def safe_get(
    session: requests.Session,
    url: str,
    *,
    source: str = "",
    must: bool = False,
    allow_404: bool = False,
    log_tag: str = "",
    retries: int = 2,
    backoff: float = 0.7,
    **kwargs
) -> Optional[requests.Response]:
    """HTTP GET with strict/fail-fast option.

    When must=True (used for REQUIRED sources), any non-200 response (except allowed 404)
    or request exception will raise SourceAccessError after retries.
    """
    global HTTP_FAIL_COUNT
    last_detail = ""
    for attempt in range(retries + 1):
        try:
            r = session.get(url, timeout=kwargs.pop("timeout", TIMEOUT), **kwargs)
            if r.status_code == 200:
                return r
            if allow_404 and r.status_code == 404:
                return None

            last_detail = f"status={r.status_code}"
        except Exception as e:
            last_detail = f"exc={type(e).__name__}:{e}"

        # retry?
        if attempt < retries:
            time.sleep(backoff * (2 ** attempt))
            continue

        # final failure
        HTTP_FAIL_COUNT += 1
        if len(HTTP_FAILS) < HTTP_FAIL_LIMIT:
            tag = log_tag or source or "http"
            HTTP_FAILS.append((url, f"{tag} {last_detail}".strip()))

        if must:
            raise SourceAccessError(source or "unknown", url, last_detail)
        return None


# -----------------------
# Required-source HTML fetch helpers (with WAF fallback)
# -----------------------

BLOCK_HINTS = [
    "access denied",
    "forbidden",
    "you have been blocked",
    "unusual traffic",
    "attention required",
    "cloudflare",
    "captcha",
    "enable javascript",
    "bot detection",
    "security check",
]

TDNET_OK_MARKERS = [
    "tdnet", "release.tdnet.info", "/inbs/", "i_list_", "i_main_", "kjcode", "kjtitle",
    "適時開示", "東京証券取引所", "jpx",
]

KABUTAN_OK_MARKERS = [
    "kabutan.jp", "fin_quarter_result_d", "fin_quarter_result", "fin_quarter_result_c", "fin_year_result_d", "fin_year_result", "財務", "業績", "決算",
]

def _lower_ascii_safe(s: str) -> str:
    try:
        return (s or "").lower()
    except Exception:
        return ""

def _find_block_hint(html_lower: str) -> Optional[str]:
    for h in BLOCK_HINTS:
        if h in html_lower:
            return h
    return None

def fetch_html_required(
    session: requests.Session,
    url: str,
    *,
    source: str,
    ok_markers: List[str],
    allow_404: bool = False,
    timeout: int = TIMEOUT,
    log_tag: str = "",
    referer: str = "https://www.jpx.co.jp/",
) -> Optional[str]:
    """Fetch HTML for REQUIRED sources with WAF fallback.

    - Uses requests.Session first.
    - If HTML does not contain any ok_markers (case-insensitive), tries curl_cffi (if available)
      to bypass bot-detection/WAF.
    - If still not OK, writes a dump HTML file and raises SourceAccessError with a useful hint.
    - If allow_404=True and the response is 404, returns None.
    """
    # 1) primary (requests)
    r = safe_get(session, url, source=source, must=True, allow_404=allow_404, log_tag=log_tag, timeout=timeout)
    if r is None:
        if allow_404:
            return None
        raise SourceAccessError(source, url, "status=404")

    try:
        # TDNet is often cp932/shift-jis; requests' apparent_encoding is usually right
        r.encoding = r.apparent_encoding or r.encoding or "utf-8"
    except Exception:
        pass

    body = r.text or ""
    body_l = _lower_ascii_safe(body)

    if any(mk.lower() in body_l for mk in ok_markers):
        return body

    # 2) fallback (curl_cffi)
    if cffi_requests is not None:
        try:
            headers = {
                "User-Agent": UA,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
                "Referer": referer,
            }
            rr = cffi_requests.get(url, headers=headers, timeout=timeout, impersonate="chrome124")
            if rr.status_code == 404 and allow_404:
                return None
            if rr.status_code != 200:
                raise SourceAccessError(source, url, f"curl_cffi_status={rr.status_code}")

            bb = rr.text or ""
            bl = _lower_ascii_safe(bb)
            if any(mk.lower() in bl for mk in ok_markers):
                return bb

            hint = _find_block_hint(bl) or "unexpected_html_content"
            dump_path = os.path.join(os.getcwd(), f"_{source}_dump.html")
            try:
                with open(dump_path, "wb") as f:
                    f.write(rr.content)
            except Exception:
                pass
            raise SourceAccessError(source, url, f"{hint} (dump={dump_path})")
        except SourceAccessError:
            raise
        except Exception as e:
            last = f"curl_cffi_exc={type(e).__name__}:{e}"
            raise SourceAccessError(source, url, last)

    # 3) final: dump and raise
    hint = _find_block_hint(body_l) or "unexpected_html_content"
    dump_path = os.path.join(os.getcwd(), f"_{source}_dump.html")
    try:
        with open(dump_path, "wb") as f:
            f.write(r.content)
    except Exception:
        pass
    raise SourceAccessError(source, url, f"{hint} (dump={dump_path}; install curl_cffi to retry)")


def safe_date_from_any(x: Any) -> Optional[dt.date]:
    if x is None:
        return None
    s = str(x).strip()
    m = re.search(r"(20\d{2})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            return None
    return None


def yyyymm_to_period_end_date(yyyymm: str) -> Optional[dt.date]:
    """Convert 'YYYY/MM' to the last date of that month."""
    m = re.match(r"^(\d{4})/(\d{2})$", str(yyyymm).strip())
    if not m:
        return None
    y = int(m.group(1)); mo = int(m.group(2))
    if mo < 1 or mo > 12:
        return None
    last = calendar.monthrange(y, mo)[1]
    return dt.date(y, mo, last)

def fiscal_quarter_end_date(fy_end_year: int, quarter_no: int, fy_end_month: int) -> Optional[dt.date]:
    yyyymm = fiscal_quarter_end_yyyymm(fy_end_year, quarter_no, fy_end_month)
    if not yyyymm:
        return None
    return yyyymm_to_period_end_date(yyyymm)

def compute_announced_cutoff_qidx(records: List[FilingRecord], today: Optional[dt.date] = None) -> Optional[int]:
    """Compute the latest announced quarter index for this ticker.

    Purpose: prevent filling 'future/unannounced' quarters (e.g. FY Q4 not yet released),
    while not becoming overly conservative when EDINET/TDNet retrieval is incomplete.

    Rules:
      - Consider ALL sources (EDINET / TDNet / Kabutan) but never allow a quarter whose
        quarter-end date is in the future relative to `today`.
      - Prefer official filings (EDINET/TDNet). Kabutan is fallback only when no official filing exists.
    """
    if today is None:
        today = dt.date.today()

    filing_qidx: List[int] = []
    other_qidx: List[int] = []
    for rec in records:
        qidx = record_to_qidx(rec)
        if qidx is None:
            continue

        # Never treat future quarter-ends as 'announced'
        if rec.fy_end_year and rec.fy_end_month and rec.quarter_no:
            endd = fiscal_quarter_end_date(rec.fy_end_year, rec.quarter_no, rec.fy_end_month)
            if endd and endd > today:
                continue

        if rec.source in ("edinet", "tdnet"):
            filing_qidx.append(qidx)
        else:
            other_qidx.append(qidx)

    max_f = max(filing_qidx) if filing_qidx else None
    max_o = max(other_qidx) if other_qidx else None

    if max_f is not None:
        return max_f
    return max_o


def fiscal_quarter_end_yyyymm(fy_end_year: int, quarter_no: int, fy_end_month: int) -> Optional[str]:
    """Return calendar YYYY/MM for that fiscal quarter end."""
    if quarter_no < 1 or quarter_no > 4:
        return None
    if fy_end_month < 1 or fy_end_month > 12:
        return None
    start_m = (fy_end_month % 12) + 1
    end_m = ((start_m - 1) + 3 * quarter_no - 1) % 12 + 1
    end_y = fy_end_year if end_m <= fy_end_month else (fy_end_year - 1)
    return f"{end_y:04d}/{end_m:02d}"

def qidx_from_fy_q(fy_start_year: int, quarter_no: int) -> int:
    return (fy_start_year - BASE_FY_START_YEAR) * 4 + (quarter_no - 1)

def fy_start_year_from_end(fy_end_year: int, fy_end_month: Optional[int] = None) -> int:
    """Convert FY end-year/month to the sheet's FY label (start-year).

    The workbook labels quarters by fiscal-year START year (e.g. '2025年度' for FY2025).
    If the fiscal year starts in January (fy_end_month==12), then start-year == end-year.
    Otherwise start-year == end-year - 1.
    """
    if not fy_end_month:
        return fy_end_year - 1
    start_m = (int(fy_end_month) % 12) + 1
    return fy_end_year if start_m == 1 else (fy_end_year - 1)

def parse_title_for_fy_end_q(title: str) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    """Parse '2026年3月期 第3四半期' style."""
    if not title:
        return None, None, None
    fw = str.maketrans("０１２３４５６７８９", "0123456789")
    t = title.translate(fw).replace(" ", "").replace("　", "")
    fy_end_year = None
    fy_end_month = None
    q = None

    m = re.search(r"(20\d{2})年(\d{1,2})月期", t)
    if m:
        fy_end_year = int(m.group(1))
        fy_end_month = int(m.group(2))

    mq = re.search(r"第([1234])四半期", t)
    if mq:
        q = int(mq.group(1))
    # heuristics
    tl = t.lower()
    if q is None and ("中間期" in t or "半期" in t):
        q = 2
    if q is None and ("通期" in t or "本決算" in t or "年度決算" in t):
        q = 4
    if q is None and ("financialresults" in tl or "決算短信" in t):
        # Guard against false Q4 on quarterly titles where q-number extraction failed.
        # Only infer Q4 when the title is not a quarter release.
        if ("四半期" not in t) and ("quarter" not in tl):
            q = 4
    return fy_end_year, fy_end_month, q




def extract_fy_end_and_quarter_from_ixbrl_zip(zip_bytes: bytes) -> Optional[Tuple[int, int, int]]:
    '''
    Extract (fy_end_year, fy_end_month, quarter_no) from iXBRL contents in a TDNet/EDINET ZIP.

    Why:
      Some TDNet index titles can be misleading (e.g., display year), causing FY shift bugs.
      The ZIP itself usually contains authoritative metadata such as tse-ed-t:FiscalYearEnd and
      tse-ed-t:QuarterlyPeriod in the Summary iXBRL.
    '''
    try:
        import zipfile, io, re
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except Exception:
        return None

    names = zf.namelist()
    cand = [n for n in names if n.lower().endswith((".htm", ".html", ".xhtml"))]

    def _rank(n: str) -> Tuple[int, int]:
        nl = n.lower()
        return (0 if ("summary" in nl) else 1, len(nl))

    cand.sort(key=_rank)

    fy_end = None  # (year, month)
    q_no = None

    fy_patterns = [
        re.compile(r'FiscalYearEnd[^>]*>\s*([0-9]{4})[-/\.]([0-9]{2})[-/\.]([0-9]{2})\s*<', re.I),
        re.compile(r'name\s*=\s*["\'][^"\']*:FiscalYearEnd["\'][^>]*>\s*([0-9]{4})[-/\.]([0-9]{2})[-/\.]([0-9]{2})\s*<', re.I),
    ]
    q_patterns = [
        re.compile(r'QuarterlyPeriod[^>]*>\s*([1-4])\s*<', re.I),
        re.compile(r'name\s*=\s*["\'][^"\']*:QuarterlyPeriod["\'][^>]*>\s*([1-4])\s*<', re.I),
        re.compile(r'Quarter[^>]*Period[^>]*>\s*([1-4])\s*<', re.I),
    ]

    for n in cand[:60]:
        try:
            txt = zf.read(n).decode("utf-8", errors="ignore")
        except Exception:
            continue

        tl = txt.lower()
        if ("fiscalyearend" not in tl) and ("quarterlyperiod" not in tl) and ("quarter" not in tl):
            continue

        if fy_end is None:
            for rp in fy_patterns:
                mm = rp.search(txt)
                if mm:
                    y = int(mm.group(1))
                    mth = int(mm.group(2))
                    fy_end = (y, mth)
                    break

        if q_no is None:
            for rp in q_patterns:
                mm = rp.search(txt)
                if mm:
                    q_no = int(mm.group(1))
                    break

        if fy_end and q_no:
            break

    if not fy_end or not q_no:
        return None
    return (int(fy_end[0]), int(fy_end[1]), int(q_no))

def infer_quarter_from_period_span(period_start: Optional[dt.date], period_end: Optional[dt.date]) -> Optional[int]:
    """Infer quarter_no from months between fiscal start and end (EDINET often provides start/end)."""
    if not period_start or not period_end:
        return None
    # assume period_start is FY start; difference in months approximates quarter end.
    months = (period_end.year - period_start.year) * 12 + (period_end.month - period_start.month) + 1
    if months <= 3:
        return 1
    if months <= 6:
        return 2
    if months <= 9:
        return 3
    if months <= 12:
        return 4
    return None

def infer_fy_end_month_from_period_start(period_start: Optional[dt.date]) -> Optional[int]:
    if not period_start:
        return None
    m = period_start.month
    # FY end month is month before FY start month
    return 12 if m == 1 else (m - 1)

def infer_fy_end_year_from_period_end(period_end: Optional[dt.date], fy_end_month: Optional[int]) -> Optional[int]:
    if not period_end or not fy_end_month:
        return None
    # If quarter end month is after fy_end_month, FY ends next calendar year
    return period_end.year if period_end.month <= fy_end_month else (period_end.year + 1)

# -----------------------
# Workbook header creation
# -----------------------
def ensure_right_blocks_headers(ws) -> None:
    # Touch KE1 to expand worksheet
    _ = ws["KE1"]

    def write_block(metric: str, c1: str, c2: str) -> None:
        cols = iter_cols(c1, c2)
        ws[f"{c1}1"] = METRIC_JA[metric] + "（累積）"
        if metric in ("keijo", "uriage", "saishu"):
            start_year = 2018
        elif metric == "gross":
            start_year = 2022
        else:
            start_year = 2023
        for i, col in enumerate(cols):
            fy = start_year + (i // 4)
            q = (i % 4) + 1
            ws[f"{col}2"] = f"{fy}年度第{q}四半期"
            # row3 stays free

    for metric, (c1, c2) in RIGHT_BLOCKS.items():
        write_block(metric, c1, c2)

    # WARN column header
    if ws[f"{WARN_COL}1"].value is None:
        ws[f"{WARN_COL}1"] = "WARN_FLAG"
    if ws[f"{WARN_COL}2"].value is None:
        ws[f"{WARN_COL}2"] = "直近5Qで差異±2超"

# -----------------------
# Metric mapping (XBRL name heuristics)
# -----------------------
def split_camel_tokens(name: str) -> List[str]:
    """
    Split a CamelCase XBRL local-name into tokens.

    Example:
      "NetSalesSummaryOfBusinessResults" -> ["Net","Sales","Summary","Of","Business","Results"]
      "ValuationDifferenceOnAvailableForSaleSecuritiesNetOfTaxOCI" -> ["Valuation","Difference","On","Available","For","Sale","Securities","Net","Of","Tax","OCI"]
    """
    if not name:
        return []
    # keep acronyms together (OCI, EPS, JPY, etc.)
    return re.findall(r"[A-Z]+(?=[A-Z][a-z]|[0-9]|$)|[A-Z]?[a-z]+|[0-9]+", name)

def metric_from_qname(qname: str) -> Optional[str]:
    """
    Map iXBRL qname -> internal metric key.

    v5.18 FIX:
    - Avoid naive substring matching that can cause false positives (e.g. "...ForSaleSecurities" -> "sales").
    - Prefer robust token-based matching for revenue tags.
    - Support IFRS issuers where "経常利益" is not present; in that case we use "税引前利益"
      (profit before tax) as the substitute for the 'keijo' metric.
    """
    if not qname:
        return None

    local = (qname.split(":")[-1] or "").strip()
    n = local.lower()

    # v16: synthetic metric support (for derived/summed facts)
    if n.startswith("synthcapexppe"):
        return "capex_ppe"

    toks = split_camel_tokens(local)
    tl = [t.lower() for t in toks]

    # ---- sales / revenue (uriage) ----
    # Common in J-GAAP (jppfs): NetSales / NetSalesSummaryOfBusinessResults...
    if len(tl) >= 2 and tl[0] == "net" and tl[1] == "sales":
        return "uriage"

    # IFRS: RevenueIFRS / Revenue2IFRS etc.
    # NOTE: exclude RevenueFromExternalCustomers... because it is typically segment disclosure, not top-line.
    if len(tl) >= 3 and tl[0] == "revenue" and tl[1] == "from" and tl[2] == "external":
        pass
    else:
        if len(tl) >= 1 and tl[0] == "revenue":
            return "uriage"

    # Banking etc: OperatingRevenues / OperatingRevenue
    if len(tl) >= 2 and tl[0] == "operating" and tl[1] in ("revenues", "revenue"):
        return "uriage"

    # Some issuers may use SalesSummaryOfBusinessResults (rare). Allow this exact token pattern only.
    if len(tl) >= 5 and tl[:5] == ["sales", "summary", "of", "business", "results"]:
        return "uriage"

    # Finance: GrossOperatingRevenue(s)
    if len(tl) >= 3 and tl[0] == "gross" and tl[1] == "operating" and tl[2] in ("revenue", "revenues"):
        return "uriage"

    # As a final fallback, allow exact local-name match only
    if n in ("netsales", "operatingrevenues", "operatingrevenue", "revenue", "sales"):
        return "uriage"
    # ---- ordinary income / profit (keijo) ----
    # IMPORTANT (r24): Avoid substring traps and bank-specific revenue tags.
    # - 'ExtraordinaryProfit' contains 'ordinaryprofit' as a substring. It is NOT 経常利益.
    # - Banks may disclose 'OrdinaryIncomeBNK' which is revenue-like (経常収益).
    #   Exclude BNK variants from keijo unless it is explicitly a profit concept (..Loss/..ProfitLoss).
    ln = n

    if (('extraordinary' in ln) and (('ordinaryprofit' in ln) or ('ordinaryincome' in ln))):
        # do not map
        pass
    else:
        is_bnk = ('bnk' in ln) or ln.endswith('bnk')
        if any(k in ln for k in ('ordinaryincomeloss', 'ordinaryprofitloss', 'ordinaryprofit')):
            return 'keijo'
        if ('ordinaryincome' in ln) and (not is_bnk):
            return 'keijo'

    # IFRS / some issuers: ProfitBeforeIncomeTaxes / ProfitLossBeforeTax...
    # Map these to 'keijo' as the best available substitute when '経常利益' is not defined.
    if any(k in ln for k in [
        'profitbeforeincometaxes', 'incomebeforeincometaxes',
        'profitlossbeforetax', 'profitlossbeforetaxifrs',
        'profitlossbeforetaxandextraordinaryitems',
        'profitbeforetax', 'incomebeforetax'
    ]):
        return 'keijo'
    # ---- net income (saishu) ----
    # r24: Some filers (esp. non-consolidated or small issuers) use plain 'ProfitLoss' / 'NetIncomeLoss'
    # for the bottom line. Treat those as saishu, but keep it conservative to avoid picking before-tax concepts.
    if 'profitlossattributabletoownersofparent' in ln or 'netincomelossattributabletoownersofparent' in ln:
        return 'saishu'
    # conservative: only accept exact bottom-line concepts
    if ln in ('profitloss', 'profitlossifrs', 'netincomeloss', 'netincome', 'netincomeifrs'):
        return 'saishu'
    if ln.endswith('profitloss') and 'summaryofbusinessresults' in ln:
        return 'saishu'

    # ---- gross profit ----
    if "grossprofit" in n:
        return "gross"
    if "operatinggrossprofit" in n or "grossoperatingprofit" in n:
        return "gross"
    if "grossoperatingrevenue" in n or "grossoperatingrevenues" in n:
        return "gross"
    if "netoperatingrevenue" in n or "netoperatingrevenues" in n:
        return "gross"
    if "pureoperatingrevenue" in n or "pureoperatingrevenues" in n:
        return "gross"

    # ---- SG&A ----
    # r26: Broaden tag coverage. Some issuers (esp. financials) use different concept names.
    if "sellinggeneralandadministrativeexpenses" in n or "sgaexpenses" in n:
        return "sga"
    if "generalandadministrativeexpenses" in n or "generaladministrativeexpenses" in n:
        return "sga"
    if "administrativeexpenses" in n and ("expense" in n or "expenses" in n):
        return "sga"
    # Finance/banking: OperatingExpenses is often the closest counterpart to SG&A.
    if n in ("operatingexpenses", "operatingexpense"):
        return "sga"

    # ---- depreciation ----
    if any(k in n for k in [
        "depreciation", "depreciationandamortisation", "depreciationandamortization",
        "depreciationexpense", "depreciationofpropertyplantandequipment"
    ]):
        return "genka"

    # ---- operating cash flow (opcf) ----
    if any(k in n for k in [
        "netcashprovidedbyusedinoperatingactivities",
        "netcashprovidedbyusedinoperatingactivitiesifrs",
        "netcashprovidedbyusedinoperatingactivitiesusgaap",
        "netcashprovidedbyusedinoperatingactivitiesconsolidated",
        "netcashprovidedbyusedinoperatingactivitiesnonconsolidated",
        "netcashprovidedbyusedinoperatingactivitiescontinuingoperations",
        "netcashprovidedbyusedinoperatingactivitiescontinuingoperationsifrs"
    ]):
        return "opcf"

    # ---- advertising expenses (ad) ----
    # Common: AdvertisingExpensesSGA, AdvertisingAndSalesPromotionExpenses, etc.
    if "advertising" in n and any(k in n for k in ("expense", "expenses", "cost", "costs")):
        return "ad"
    if ("salespromotion" in n or "salespromotionexpenses" in n) and any(k in n for k in ("expense", "expenses", "cost", "costs")):
        return "ad"
    if "promotion" in n and "advertising" in n:
        return "ad"

    # ---- research and development (rnd) ----
    if "researchanddevelopment" in n and any(k in n for k in ("expense", "expenses", "cost", "costs")):
        return "rnd"
    if "researchanddevelopmentexpenses" in n:
        return "rnd"

    # ---- purchase of PPE / capex (capex_ppe) ----
    if any(k in n for k in [
        "purchaseofpropertyplantandequipment",
        "purchaseofpropertyplantandequipmentifrs",
        "purchaseofpropertyplantandequipmentandintangibleassets",
        "paymentsforpurchaseofpropertyplantandequipment",
        "paymentsforpurchasesofexplorationandevaluationassetsinvcf",
        "paymentsforpurchasesofdevelopmentandproductionassetsinvcf",
        "paymentsforpurchasesofexplorationandevaluationassets",
        "paymentsforpurchasesofdevelopmentandproductionassets"

    ]):
        return "capex_ppe"

    return None


def instant_metric_from_qname(qname: str) -> Optional[str]:
    """Map iXBRL qname -> internal instant metric key (e.g., assets).

    IMPORTANT:
      - Avoid misclassifying partial categories like CurrentAssets / NoncurrentAssets as total Assets.
      - Keep mapping conservative; prefer exact matches and a small allowlist of known total-assets variants.
    """
    if not qname:
        return None
    local = (qname.split(":")[-1] or "").strip()
    n = local.lower()

    # Exact / canonical total assets
    if n in ("assets", "assetsifrs", "totalassets", "totalassetsifrs"):
        return "assets"

    # Known variants that still represent total assets
    if n in ("consolidatedassets", "consolidatedtotalassets", "assetsconsolidated"):
        return "assets"


    # Summary-of-business-results variants (EDINET '業績サマリー' contexts).
    # Some issuers expose total assets only via e.g. TotalAssetsUSGAAPSummaryOfBusinessResults.
    if ("summaryofbusinessresults" in n) and ("totalassets" in n):
        return "assets"

    # Conservative suffix fallback: allow only when it clearly indicates "total" concept.
    if n.endswith("assets"):
        # Exclude obvious partial buckets that are NOT total assets
        if any(k in n for k in (
            "current", "noncurrent", "intangible", "fixed", "financial", "other",
            "inventory", "receivable", "cash", "investment", "property",
            "rightofuse", "deferred", "tax", "segment", "heldforsale",
            "disposalgroup", "note", "of"
        )):
            return None
        # Allow only when local name contains an explicit "total" signal
        if any(k in n for k in ("total", "consolidated", "sum")):
            return "assets"

    return None


def choose_best_instant_by_context(
    cand: List[Tuple[str, str, float, str]],
    prefer_consolidated: bool,
) -> Dict[str, float]:
    """
    Choose best INSTANT fact per metric from candidates.
    cand: list of (qname, ctxRef, value_million, filename)
    """
    best: Dict[str, Tuple[Tuple[int, int, int, int, int], float]] = {}

    def score_ctx(ctx: str, fn: str) -> Tuple[int, int, int, int, int]:
        l = (ctx or "").lower()
        lf = (fn or "").lower()
        # avoid member/segment contexts
                # avoid member/segment contexts (but do NOT penalize a pure 'ConsolidatedMember' context)
        tokens = re.split(r"[_\s]+", l)
        member_tokens = [t for t in tokens if (('member' in t) or ('segment' in t) or ('reportable' in t) or ('businesssegment' in t))]
        if not member_tokens:
            member = 0
        elif len(member_tokens) == 1 and member_tokens[0] in ('consolidatedmember','consolidated_member'):
            member = 0
        else:
            member = 1
        # instant-like contexts are preferred
        instant = 0 if ("instant" in l or "asof" in l) else 1
        # consolidated preference
        cons = 0
        if prefer_consolidated:
            cons = 0 if ("consolidated" in l or "consolidated" in lf) else 1
        else:
            cons = 0 if ("nonconsolidated" in l or "nonconsolidated" in lf) else 1
        # current vs prior
        current = 0 if ("current" in l or "this" in l) else 1
        # file preference (statement-like over summary)
        summary = 1 if any(k in lf for k in ("summary", "gaiyou", "overview", "highlight", "gaikyo")) else 0
        return (member, instant, cons, current, summary)

    for qn, ctx, v, fn in cand:
        m = instant_metric_from_qname(qn)
        if not m:
            continue
        try:
            fv = float(v)
        except Exception:
            continue
        if fv < 0 and m == "assets":
            continue
        s = score_ctx(ctx, fn)
        cur = best.get(m)
        if cur is None or s < cur[0]:
            best[m] = (s, fv)

    return {k: v for k, (s, v) in best.items()}


def parse_zip_instant_metrics(
    zip_bytes: bytes,
    prefer_consolidated: bool,
    log: List[List[Any]],
    who: str,
    expected_ticker: Optional[str] = None,
) -> Dict[str, float]:
    """
    Parse EDINET/TDNet iXBRL ZIP -> INSTANT metrics (e.g., total assets).

    NOTE:
    - Instant metrics are treated as "as of period end" values for the filing's quarter.
    - Values are converted to million JPY (百万円) via extract_inline_facts().
    """
    if expected_ticker:
        try:
            code_in_zip = extract_securities_code_from_zip(zip_bytes)
            if code_in_zip and not securities_code_matches_expected(code_in_zip, expected_ticker):
                return {}
        except Exception:
            pass

    all_facts: List[Tuple[str, str, float, str]] = []
    for fn, txt in iter_html_candidates_from_zip(zip_bytes):
        try:
            facts = extract_inline_facts(txt)
        except Exception as e:
            log.append([None, None, 'zip_parse_error', None, None, None, None, None, None, who, True, None, f'{who}:extract_inline_facts_error:{type(e).__name__}:{e}'])
            facts = []
        if not facts:
            continue
        for qn, ctx, v in facts:
            if not qn or not ctx:
                continue
            # Keep only likely instant metrics to reduce work
            if instant_metric_from_qname(qn):
                all_facts.append((qn, ctx, float(v), fn))

    if not all_facts:
        return {}

    out = choose_best_instant_by_context(all_facts, prefer_consolidated=prefer_consolidated)
    return out


def detect_unit_scale_to_million(text: str) -> float:
    """Return multiplier to convert disclosed unit to million JPY."""
    if not text:
        return 1.0
    head = text[:60000]  # r26: widen for TDNet headers  # enough for header + first tables
    # prioritize explicit unit markers
    has_m = re.search(r"単位[:：]?\s*百万円", head) or re.search(r"\(百万円\)", head)
    has_k = re.search(r"単位[:：]?\s*千円", head) or re.search(r"\(千円\)", head)
    has_y = re.search(r"単位[:：]?\s*円", head) or re.search(r"\(円\)", head)

    if has_m:
        return 1.0
    if has_k and not has_m:
        return 1.0 / 1000.0
    if has_y and not has_m and not has_k:
        return 1.0 / 1_000_000.0
    # fallback: if "千円" appears a lot and "百万円" not, assume thousand
    if ("千円" in head) and ("百万円" not in head):
        return 1.0 / 1000.0
    return 1.0

def extract_inline_facts(html_text: str) -> List[Tuple[str, str, float]]:
    """
    Return list of (qname, contextref, value_in_million).

    IMPORTANT:
    - For iXBRL (EDINET/TDNet), inline numeric facts often carry a `scale` attribute (power of 10).
      The visible text is scaled, while the unitRef is typically JPY (yen).
      We must convert to "million JPY" consistently as:
          value_million = parsed_text * 10^(scale) / 1,000,000   (when unitRef is JPY)
      This avoids relying on fragile "単位:千円/百万円" header detection which can be outside the first pages.

    - If unitRef is missing (rare), we fall back to disclosure-unit detection from the HTML text.
    """
    soup = make_soup(html_text)
    out: List[Tuple[str, str, float]] = []
    unit_scale_fallback = detect_unit_scale_to_million(html_text)

    tags = soup.find_all(attrs={"name": True, "contextref": True})
    for t in tags:
        qn = t.get("name", "")
        ctx = t.get("contextref", "")
        if not qn or not ctx:
            continue

        # v16 FIX: Exclude non-numeric inline facts (ix:nonNumeric) and TextBlock qnames.
        # These can contain note numbers like "※3" which would otherwise be parsed as numeric and cause false WARNs.
        tname = str(getattr(t, "name", "") or "").lower()
        if "nonnumeric" in tname:
            continue
        local = (qn.split(":")[-1] or "").strip().lower()
        if local.endswith("textblock"):
            continue

        lctx = (ctx or "").lower()
        if "prior" in lctx or "previous" in lctx:
            continue

        txt = t.get_text(" ", strip=True)
        v = parse_num(txt)
        if v is None:
            continue

        # iXBRL negative values can be represented via `sign` attribute (e.g., sign="-") rather than a leading minus/△.
        try:
            sign_attr = (t.get("sign") or t.get("signRef") or t.get("signref") or "").strip()
        except Exception:
            sign_attr = ""
        if sign_attr:
            if sign_attr in {"-", "−", "－"}:
                if float(v) > 0:
                    v = -float(v)
            elif sign_attr in {"△", "▲"}:
                if float(v) > 0:
                    v = -float(v)

        # scale attribute is common in iXBRL (e.g., scale="3" for 千円表示 with JPY unit)
        try:
            scale = int(str(t.get("scale") or "0").strip() or "0")
        except Exception:
            scale = 0

        unitref = (t.get("unitref") or t.get("unitRef") or "").strip()
        u = unitref.lower()

        # Monetary JPY facts
        if "jpy" in u and "pershare" not in u and "per_share" not in u and "per-share" not in u:
            # r26: Some TDNet packages omit iXBRL `scale` even when the report is in 百万円/千円.
            # When scale==0, prefer the disclosure-unit detection from HTML headers.
            try:
                val_m_scale = float(v) * (10 ** (scale - 6))
            except OverflowError:
                val_m_scale = float(v) * math.pow(10.0, float(scale - 6))

            if scale == 0 and unit_scale_fallback not in (1.0 / 1_000_000.0,):
                try:
                    val_m_unit = float(v) * float(unit_scale_fallback)
                except Exception:
                    val_m_unit = val_m_scale
                # Choose the unit-based value unless it is clearly absurd and the scale-based one is sane.
                if abs(val_m_unit) > MAX_ABS_MILLION_SANITY and abs(val_m_scale) <= MAX_ABS_MILLION_SANITY:
                    val_m = val_m_scale
                else:
                    val_m = val_m_unit
            else:
                val_m = val_m_scale

            # Outlier guard: if the parsed text accidentally concatenated multiple numbers (common iXBRL noise),
            # try a safer "first numeric token" parse once; otherwise drop the fact.
            if abs(val_m) > MAX_ABS_MILLION_SANITY:
                v2 = parse_num_first_token(txt)
                if v2 is None:
                    continue
                # apply iXBRL sign attribute to fallback value
                if sign_attr:
                    if sign_attr in {"-", "−", "－"}:
                        if float(v2) > 0:
                            v2 = -float(v2)
                    elif sign_attr in {"△", "▲"}:
                        if float(v2) > 0:
                            v2 = -float(v2)
                try:
                    val_m2 = float(v2) * (10 ** (scale - 6))
                except OverflowError:
                    val_m2 = float(v2) * math.pow(10.0, float(scale - 6))
                if abs(val_m2) > MAX_ABS_MILLION_SANITY:
                    continue
                val_m = val_m2

            out.append((qn, ctx, val_m))
            continue

        # If unitRef is missing, accept ONLY numeric facts that map to known metrics.
        # Rationale: ix:nonNumeric / TextBlock often embeds note numbers (e.g., "※3") and would create false facts.
        # Some filings may omit unitRef on monetary facts; for those, we fall back to disclosure-unit scale.
        if not unitref:
            if "nonfraction" not in tname:
                continue
            if (metric_from_qname(qn) is None) and (instant_metric_from_qname(qn) is None):
                continue
            val_m = float(v) * unit_scale_fallback
            if abs(val_m) > MAX_ABS_MILLION_SANITY:
                v2 = parse_num_first_token(txt)
                if v2 is None:
                    continue
                # apply sign attribute to fallback too
                if sign_attr:
                    if sign_attr in {"-", "−", "－"}:
                        if float(v2) > 0:
                            v2 = -float(v2)
                    elif sign_attr in {"△", "▲"}:
                        if float(v2) > 0:
                            v2 = -float(v2)
                val_m2 = float(v2) * unit_scale_fallback
                if abs(val_m2) > MAX_ABS_MILLION_SANITY:
                    continue
                val_m = val_m2
            out.append((qn, ctx, val_m))
            continue

        # Non-monetary units are ignored.
        continue

    return out

def choose_best_by_context(
    facts: List[Tuple[str, str, float]],
    prefer_consolidated: bool = True
) -> Tuple[Dict[str, float], Dict[str, float]]:
    """Split iXBRL facts into (cumulative, single) using contextRef heuristics.

    v5.4 FIX:
    - Do NOT output a 'single' value unless the contextRef strongly indicates quarter duration.
      Otherwise the downstream merge would treat YTD (累計) as a single-quarter value and overwrite
      the derived single (cum diff).

    v5.5 FIX (2026-02-18):
    - Some industries (notably banks) disclose BOTH:
        * OrdinaryIncome (経常収益/経常収入 に相当することがある; revenue-like)
        * OrdinaryIncomeLoss / OrdinaryProfitLoss (経常利益)
      in the SAME statement contextRef.
      Historically we mapped both into the same metric 'keijo', which made tie-breaking depend on
      source order inside the iXBRL. That caused wrong EDINET adoption and false WARN (orange cells)
      even when the Excel value matched the income statement.
      => We keep "context suitability" as the primary selector, and add a *name-based tie-break*
         only when context score ties.
    """
    cum: Dict[str, float] = {}
    single: Dict[str, float] = {}

    # Keep qname so we can break ties deterministically
    cand: Dict[str, List[Tuple[str, str, float]]] = {}
    for qn, ctx, v in facts:
        m = metric_from_qname(qn)
        if not m:
            continue
        cand.setdefault(m, []).append((qn, ctx, v))

    def name_penalty(metric: str, qname: str) -> int:
        """Lower is better. Only used as a *tie-breaker* after context score."""
        local = ((qname or "").split(":")[-1] or "").lower()

        if metric == "keijo":
            # Prefer profit concepts over revenue-like concepts.
            if any(k in local for k in ("ordinaryincomeloss", "ordinaryprofitloss", "ordinaryprofit")):
                return 0
            # IFRS substitute (when 経常利益 is not defined)
            if any(k in local for k in ("profitbeforeincometaxes", "profitlossbeforetax", "profitbeforetax")):
                return 1
            # Ambiguous: can be 経常利益 for non-banks, but can be 経常収益 for banks.
            if "ordinaryincome" in local:
                return 2
            return 3

        if metric == "eigyo":
            if any(k in local for k in ("operatingincomeloss", "operatingprofitloss", "operatingprofit")):
                return 0
            if "operatingincome" in local:
                return 1
            return 2

        # Default: no tie-break
        return 0

    def score_ctx(ctx: str, kind: str, qn: str, metric: str) -> Tuple[int, int, int, int]:
        l = (ctx or "").lower()

        noncon = 1 if "nonconsolidated" in l else 0
        cons_score = noncon if prefer_consolidated else (1 - noncon)

        cur_score = 0 if "current" in l else 1

        if kind == "cum":
            if any(k in l for k in [
                "currentytdduration", "currentyeartodateduration", "currentyearduration",
                "currentduration", "currentaccumulated", "currentaccumulatedduration",
                "currentinterimduration", "currentinterimyeartodateduration",
            ]):
                kind_score = 0
            else:
                kind_score = 2
        else:
            if any(k in l for k in [
                "currentquarterduration", "currentquarter", "currentq",
                "quarterduration", "q1duration", "q2duration", "q3duration", "q4duration",
            ]):
                kind_score = 0
            else:
                kind_score = 2

        # Add name tie-break only at the end (context first)
        return (cons_score, cur_score, kind_score, name_penalty(metric, qn))

    for m, vals in cand.items():
        best_c = None
        best_s = None
        for qn, ctx, v in vals:
            sc = score_ctx(ctx, "cum", qn, m)
            if best_c is None or sc < best_c[0]:
                best_c = (sc, v)

            ss = score_ctx(ctx, "single", qn, m)
            if ss[2] == 0:
                if best_s is None or ss < best_s[0]:
                    best_s = (ss, v)

        if best_c is not None:
            cum[m] = float(best_c[1])
        if best_s is not None:
            single[m] = float(best_s[1])

    return cum, single

def extract_qualitative_table_metrics(html_text: str) -> Dict[str, float]:
    """
    Parse qualitative HTML tables for metrics (mostly for depreciation/gross/sga).
    Picks the rightmost numeric cell in the row.
    Keeps unit scaling based on detected unit.
    """
    out: Dict[str, float] = {}
    if not html_text:
        return out
    unit_scale = detect_unit_scale_to_million(html_text)
    soup = make_soup(html_text)

    label_to_metric = [
        ("減価償却費", "genka"),
        ("売上総利益", "gross"),
        ("販売費及び一般管理費", "sga"),
        ("販管費", "sga"),
    ]
    revenue_labels = ("売上高合計", "売上収益合計", "営業収益合計", "売上高", "売上収益", "営業収益")
    cogs_labels = ("売上原価合計", "売上原価")

    def extract_nums(s: str) -> List[float]:
        vals: List[float] = []
        if not s:
            return vals
        t = s.replace("△", "-").replace("▲", "-").replace("−", "-").replace("－", "-")
        # parenthesis negative
        for m in re.finditer(r"\(([\d,]+(?:\.\d+)?)\)", t):
            try:
                vals.append(-float(m.group(1).replace(",", "")))
            except Exception:
                pass
        for m in re.finditer(r"(?<![\d.])-?[\d,]+(?:\.\d+)?", t):
            try:
                vals.append(float(m.group(0).replace(",", "")))
            except Exception:
                pass
        return vals

    def pick_amount_like(nums: List[float]) -> Optional[float]:
        if not nums:
            return None
        # Prefer large magnitudes when both amount and ratio(%) are mixed in the same row.
        large = [v for v in nums if abs(v) >= 500]
        if large:
            return float(large[-1])
        return float(nums[-1])

    rev_candidate: Optional[Tuple[int, float]] = None
    cogs_candidate: Optional[Tuple[int, float]] = None

    for tr in soup.find_all("tr"):
        tds = tr.find_all(["td", "th"])
        if not tds:
            continue
        row_texts = [c.get_text(" ", strip=True) for c in tds]
        row_join = " ".join(row_texts)
        head = row_texts[0] if row_texts else ""
        head_compact = re.sub(r"\s+", "", head or "")

        row_nums: List[float] = []
        for c in tds[1:] if len(tds) >= 2 else tds:
            row_nums.extend(extract_nums(c.get_text(" ", strip=True)))
        if not row_nums:
            row_nums = extract_nums(row_join)

        for jp, metric in label_to_metric:
            if jp in row_join:
                v = pick_amount_like(row_nums)
                if v is not None:
                    out[metric] = v * unit_scale

        # Last-resort ingredients for gross proxy from qualitative table:
        # gross ~= revenue - cost of sales.
        if row_nums:
            if any(lbl in head_compact for lbl in revenue_labels):
                # Penalize likely segment/detail rows; prefer exact top-level labels.
                rank = 0 if any(head_compact.startswith(lbl) for lbl in revenue_labels) else 1
                if any(ng in head_compact for ng in ("セグメント", "部門", "内訳", "構成比")):
                    rank += 3
                v = pick_amount_like(row_nums)
                if v is not None and (rev_candidate is None or rank < rev_candidate[0]):
                    rev_candidate = (rank, v)
            if any(lbl in head_compact for lbl in cogs_labels):
                rank = 0 if any(head_compact.startswith(lbl) for lbl in cogs_labels) else 1
                if any(ng in head_compact for ng in ("セグメント", "部門", "内訳", "構成比")):
                    rank += 3
                v = pick_amount_like(row_nums)
                if v is not None and (cogs_candidate is None or rank < cogs_candidate[0]):
                    cogs_candidate = (rank, v)

    if "gross" not in out and rev_candidate is not None and cogs_candidate is not None:
        out["gross"] = (float(rev_candidate[1]) - float(cogs_candidate[1])) * unit_scale
    return out

# -----------------------
# TDNet LIVE index prefetch
# -----------------------
TDNET_BASE = "https://www.release.tdnet.info/inbs"

# -----------------------
# TDNet GitHub archive helpers
# -----------------------
def _gh_tree_url(owner: str, repo: str, branch: str) -> str:
    return f"https://api.github.com/repos/{owner}/{repo}/git/trees/{branch}?recursive=1"

def _gh_raw_url(owner: str, repo: str, branch: str, path: str) -> str:
    # raw.githubusercontent.com serves binary content directly
    p = path.lstrip('/')
    return f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{p}"

def _extract_date_from_path(path: str) -> Optional[dt.date]:
    if not path:
        return None
    s = path
    # YYYYMMDD (optionally with separators/underscore) – pick the first plausible date
    m = re.search(r"(20\d{2})[\-/_.]?(\d{2})[\-/_.]?(\d{2})", s)
    if m:
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1 <= mo <= 12 and 1 <= d <= 31:
                return dt.date(y, mo, d)
        except Exception:
            return None
    return None

def _extract_ticker_from_path(path: str) -> Optional[str]:
    """Infer ticker code from GitHub path/filename.

    r27 fix:
    - The r25/r26 regex used `(?!\\w)` boundary, which *rejected* tickers followed by '_' (common in this repo),
      resulting in kept=0 and missing older TDNet XBRL from GitHub.
    - We now treat boundaries as "non-alphanumeric" (underscore is allowed), and avoid matching date directories.
    """
    if not path:
        return None
    s = str(path).upper()

    # Prefer basename (usually "7735_...zip") if present.
    base = os.path.basename(s)
    for target in (base, s):
        # Accept 4-digit or 3-digit+alnum codes, optionally with trailing 0 (TDNet style).
        # Boundary: not [0-9A-Z] on both sides (underscore is OK).
        m = re.search(r"(?<![0-9A-Z])(\d{4}0?|\d{3}[A-Z0-9]0?)(?![0-9A-Z])", target)
        if not m:
            continue
        code = m.group(1)
        if len(code) == 5 and code.endswith('0'):
            code = code[:-1]
        return normalize_ticker(code)

    return None

def tdnet_github_build_index(
    session: requests.Session,
    tickers: set,
    *,
    owner: str,
    repo: str,
    branch: str,
    xbrl_dir: str,
    from_date: Optional[dt.date],
    log: List[List[Any]],
    must: bool = False,
) -> Dict[str, List[TdnetIndexItem]]:
    """Build TDNet index from a GitHub archive repo.

    We rely on GitHub REST API (git trees) to list files recursively, then use raw.githubusercontent.com
    URLs to download ZIPs. This is used as a fallback when TDNet LIVE doesn't provide older data.

    Notes:
    - We only use XBRL ZIPs. PDF files in 'tekigikaizi' are not parsed by this script.
    - Filenames/paths are heuristically parsed to infer ticker and disclosure date.
    """
    index: Dict[str, List[TdnetIndexItem]] = {t: [] for t in tickers}
    if not owner or not repo or not branch or not xbrl_dir:
        return index

    url = _gh_tree_url(owner, repo, branch)
    headers = {"Accept": "application/vnd.github+json"}
    r = safe_get(session, url, source="github", must=must, log_tag="github_tree", headers=headers, timeout=TIMEOUT)
    if not r:
        log.append(["", "", "tdnet_github_index", f"unavailable owner={owner} repo={repo}"])
        return index
    try:
        j = r.json()
    except Exception as e:
        log.append(["", "", "tdnet_github_index", f"invalid_json:{type(e).__name__}"])
        return index
    tree = j.get("tree") if isinstance(j, dict) else None
    if not isinstance(tree, list):
        log.append(["", "", "tdnet_github_index", "missing_tree_list"])
        return index

    total = 0
    kept = 0
    prefix = xbrl_dir.strip('/').rstrip('/') + '/'
    for it in tree:
        try:
            if not isinstance(it, dict):
                continue
            if it.get('type') != 'blob':
                continue
            p = str(it.get('path') or '')
            if not p.startswith(prefix):
                continue
            if not p.lower().endswith('.zip'):
                continue
            total += 1
            tkr = _extract_ticker_from_path(p)
            if not tkr or tkr not in index:
                continue
            d = _extract_date_from_path(p)
            if from_date and d and d < from_date:
                continue
            gh_name = os.path.splitext(os.path.basename(p))[0]
            if not tdnet_is_usable_results_title(gh_name):
                continue
            # If we cannot infer a date, keep but push to the end (date.min).
            if d is None:
                d = dt.date.min
            raw = _gh_raw_url(owner, repo, branch, p)
            title = f"[GH]{os.path.basename(p)}"
            index[tkr].append(TdnetIndexItem(ticker=tkr, title=title, zip_url=raw, disclosure_date=d, origin='github'))
            kept += 1
        except Exception:
            continue

    # sort & dedup
    for t, arr in index.items():
        arr.sort(key=lambda x: (x.disclosure_date, x.zip_url), reverse=True)
        seen = set()
        uniq = []
        for it in arr:
            if it.zip_url in seen:
                continue
            seen.add(it.zip_url)
            uniq.append(it)
        index[t] = uniq

    log.append(["", "", "tdnet_github_index", f"owner={owner} repo={repo} branch={branch} total_zip={total} kept={kept}"])
    try:
        print(f"[TDNET_GH] index built: owner={owner} repo={repo} branch={branch} total_zip={total} kept={kept}")
    except Exception:
        pass
    return index

@dataclass
class TdnetIndexItem:
    ticker: str
    title: str
    zip_url: str
    disclosure_date: dt.date
    origin: str = "live"  # live: release.tdnet.info, github: GitHub archive mirror

_TDNET_FIN_RESULTS_PATTERNS = [
    re.compile(r"決算短信"),
    re.compile(r"\bFinancial\s+Results\b", re.IGNORECASE),
    re.compile(r"\bConsolidated\s+Financial\s+Results\b", re.IGNORECASE),
    re.compile(r"\bQuarterly\s+Financial\s+Results\b", re.IGNORECASE),
]

_TDNET_EXCLUDE_TITLE_PATTERNS = [
    # Not actual earnings releases (they often contain only revised forecasts).
    re.compile(r"業績予想"),
    re.compile(r"予想.*修正"),
    re.compile(r"修正.*予想"),
    re.compile(r"Forecast", re.IGNORECASE),
    re.compile(r"Revision", re.IGNORECASE),
]

def tdnet_is_financial_results_title(title: str) -> bool:
    t = (title or "").replace("\xa0", " ").replace("\u3000", " ").strip()
    if not t:
        return False
    return any(p.search(t) for p in _TDNET_FIN_RESULTS_PATTERNS)

def tdnet_is_usable_results_title(title: str) -> bool:
    t = (title or "").replace("\xa0", " ").replace("\u3000", " ").strip()
    if not t:
        return False
    if not tdnet_is_financial_results_title(t):
        return False
    if any(p.search(t) for p in _TDNET_EXCLUDE_TITLE_PATTERNS):
        return False
    return True

def tdnet_list_pages_for_day(day: dt.date) -> List[str]:
    # TDNet list pages: start from I_list_001_YYYYMMDD.html. Additional pages may or may not exist.
    daystr = day.strftime("%Y%m%d")
    return [f"{TDNET_BASE}/I_list_001_{daystr}.html"]

def tdnet_fetch_today_main(session: requests.Session) -> Optional[str]:
    url = f"{TDNET_BASE}/I_main_00.html"
    html = fetch_html_required(session, url, source="tdnet", ok_markers=TDNET_OK_MARKERS, timeout=TIMEOUT, log_tag="tdnet_main")
    return html

def tdnet_parse_list_html_to_items(html_text: str, disclosure_date: dt.date, tickers_filter: Optional[set] = None) -> List[TdnetIndexItem]:
    """
    Parse a TDNet daily list HTML (I_list_XXX_YYYYMMDD.html) and return items.

    TDNet's table often contains:
      - td.kjCode  : securities code (sometimes 5 digits ending with 0; use first 4 digits)
      - td.kjTitle : title with link(s) to PDF and/or XBRL ZIP
    We prefer a .zip URL when present; otherwise keep the first document URL.
    """
    items: List[TdnetIndexItem] = []
    if not html_text:
        return items

    soup = make_soup(html_text)

    # Prefer the table that has kjCode/kjTitle cells
    candidate_trs = []
    for tbl in soup.find_all("table"):
        if tbl.find("td", class_=re.compile(r"\bkjCode\b")) and tbl.find("td", class_=re.compile(r"\bkjTitle\b")):
            candidate_trs = tbl.find_all("tr")
            break
    if not candidate_trs:
        candidate_trs = soup.find_all("tr")

    def code_to_4digits(raw: str) -> Optional[str]:
        if not raw:
            return None
        fw = str.maketrans("０１２３４５６７８９", "0123456789")
        s = raw.translate(fw)
        s = re.sub(r"[^0-9]", "", s)
        if len(s) >= 4:
            return s[:4]
        return None

    for tr in candidate_trs:
        # code
        td_code = tr.find("td", class_=re.compile(r"\bkjCode\b"))
        ticker = None
        if td_code:
            ticker = code_to_4digits(td_code.get_text(" ", strip=True))
        if not ticker:
            # fallback: pick 4 digits possibly followed by 0 (5-digit style)
            txt_all = tr.get_text(" ", strip=True)
            mcode = re.search(r"\b(\d{4})0?\b", txt_all)
            if mcode:
                ticker = mcode.group(1)
        if not ticker:
            continue
        if tickers_filter is not None and ticker not in tickers_filter:
            continue

        # title
        title = ""
        td_title = tr.find("td", class_=re.compile(r"\bkjTitle\b"))
        if td_title:
            title = td_title.get_text(" ", strip=True) or ""
        if not title:
            title = tr.get_text(" ", strip=True)

        # TDNet daily list includes many disclosure types; limit to financial results-type titles
        # to reduce unnecessary fetches and avoid non-XBRL PDFs.
        if not tdnet_is_usable_results_title(title):
            continue

        # url preference: ZIP > first link
        zip_url = None
        first_url = None
        for a in tr.find_all("a", href=True):
            href = (a.get("href") or "").strip()
            if not href:
                continue
            full = href if href.startswith("http") else f"{TDNET_BASE}/{href.lstrip('/')}"
            if first_url is None:
                first_url = full
            if href.lower().endswith(".zip") or ".zip?" in href.lower():
                zip_url = full
                break

        url = zip_url or first_url
        if not url:
            continue

        items.append(TdnetIndexItem(ticker=ticker, title=title, zip_url=url, disclosure_date=disclosure_date))

    return items


def tdnet_resolve_zip_from_detail(session: requests.Session, detail_url: str) -> Optional[str]:
    """Resolve a TDNet XBRL ZIP URL from a TDNet list/detail URL.

    TDNet list rows sometimes link directly to a PDF (e.g., .../XXXXXXXXXXXX.pdf). In many cases the
    corresponding XBRL ZIP is available at the same path with the extension changed to ".zip".
    Importantly, a PDF URL is *not* an HTML detail page — so we must NOT pass it to fetch_html_required().
    """
    if not detail_url:
        return None

    u = detail_url.strip()
    ul = u.lower()

    # Already a ZIP link
    if ul.endswith(".zip") or ".zip?" in ul:
        return u

    # If the list gives a PDF link, try same-id ZIP by extension swap.
    # Do NOT treat a PDF as HTML. Also, avoid aborting the whole run on a missing ZIP (404) by probing
    # with a lightweight ranged request (must=False).
    if ul.endswith(".pdf") or ".pdf?" in ul:
        cand = re.sub(r"\.pdf(\?.*)?$", ".zip", u, flags=re.IGNORECASE)
        if cand and cand != u:
            try:
                # Range request: read only the first bytes to verify ZIP signature (PK..)
                headers = {"Range": "bytes=0-7", "User-Agent": UA, "Accept": "*/*", "Referer": "https://www.jpx.co.jp/"}
                rr = session.get(cand, headers=headers, timeout=TIMEOUT, stream=True)
                try:
                    if rr.status_code in (200, 206):
                        head = rr.raw.read(4) if rr.raw else (rr.content or b"")[:4]
                        if head.startswith(b"PK"):
                            return cand
                    # 404 / others => no ZIP for this disclosure; treat as "no data" (not a source outage).
                finally:
                    try:
                        rr.close()
                    except Exception:
                        pass
            except Exception:
                # If probe itself fails, do not raise here; downstream will still have EDINET/Kabutan.
                # A genuine TDNet outage will be caught by preflight_required_sources().
                return None
        return None

    # Otherwise, treat as an HTML detail page and look for a ZIP anchor.
    html = fetch_html_required(
        session,
        u,
        source="tdnet",
        ok_markers=TDNET_OK_MARKERS,
        timeout=TIMEOUT,
        log_tag="tdnet_detail",
    )
    if html is None:
        return None

    soup = make_soup(html)
    for a in soup.find_all("a", href=True):
        href = (a.get("href") or "").strip()
        if not href:
            continue
        if href.lower().endswith(".zip") or ".zip?" in href.lower():
            return href if href.startswith("http") else f"{TDNET_BASE}/{href.lstrip('/')}"
    return None

def tdnet_build_index(session: requests.Session, tickers: set, lookback_days: int, log: List[List[Any]], *, gh_enable: bool = True, gh_owner: str = TDNET_GH_DEFAULT_OWNER, gh_repo: str = TDNET_GH_DEFAULT_REPO, gh_branch: str = TDNET_GH_DEFAULT_BRANCH, gh_xbrl_dir: str = TDNET_GH_DEFAULT_XBRL_DIR, gh_from_date: Optional[dt.date] = TDNET_GH_DEFAULT_FROM_DATE) -> Dict[str, List[TdnetIndexItem]]:
    """
    Prefetch TDNet list pages once and build ticker->items mapping.
    """
    index: Dict[str, List[TdnetIndexItem]] = {t: [] for t in tickers}
    today = dt.datetime.now(JST).date()

    # Parse today's main page as extra direct source
    main_html = tdnet_fetch_today_main(session)
    if main_html:
        for it in tdnet_parse_list_html_to_items(main_html, today, tickers_filter=tickers):
            if it.ticker in index:
                index[it.ticker].append(it)

    # Parse daily list pages
    for dd in range(0, lookback_days):
        day = today - dt.timedelta(days=dd)
        daystr = day.strftime("%Y%m%d")
        page = 1
        while True:
            url = f"{TDNET_BASE}/I_list_{page:03d}_{daystr}.html"
            html = fetch_html_required(session, url, source="tdnet", ok_markers=TDNET_OK_MARKERS, allow_404=True, timeout=TIMEOUT, log_tag="tdnet_list")
            if html is None:
                break
            items = tdnet_parse_list_html_to_items(html, day, tickers_filter=tickers)
            for it in items:
                if it.ticker in index:
                    index[it.ticker].append(it)
            page += 1
    # Dedup by zip_url
    for t, arr in index.items():
        seen = set()
        uniq = []
        # newest first by disclosure_date
        arr.sort(key=lambda x: (x.disclosure_date, x.zip_url), reverse=True)
        for it in arr:
            if it.zip_url in seen:
                continue
            seen.add(it.zip_url)
            uniq.append(it)
        index[t] = uniq


    # Optional: merge GitHub archive items (used when TDNet LIVE does not provide older files).
    if gh_enable:
        try:
            gh_idx = tdnet_github_build_index(
                session,
                tickers,
                owner=gh_owner,
                repo=gh_repo,
                branch=gh_branch,
                xbrl_dir=gh_xbrl_dir,
                from_date=gh_from_date,
                log=log,
                must=False,
            )
            for t, arr in gh_idx.items():
                if t not in index:
                    index[t] = []
                index[t].extend(arr)
        except Exception as e:
            log.append(["", "", "tdnet_github_index", f"exception:{type(e).__name__}:{e}"])

    # Re-dedup/sort after merging (live + github)
    for t, arr in index.items():
        seen = set()
        uniq = []
        arr.sort(key=lambda x: (x.disclosure_date, x.zip_url), reverse=True)
        for it in arr:
            if it.zip_url in seen:
                continue
            seen.add(it.zip_url)
            uniq.append(it)
        index[t] = uniq

    log.append(["", "", "tdnet_index", f"lookback_days={lookback_days} tickers={len(tickers)}"])
    return index

# -----------------------
# EDINET index prefetch
# -----------------------
EDINET_DOC_URL = "https://api.edinet-fsa.go.jp/api/v2/documents.json"
EDINET_DOC_URL_CANDIDATES = [
    "https://api.edinet-fsa.go.jp/api/v2/documents.json",
    "https://disclosure.edinet-fsa.go.jp/api/v2/documents.json",
]

@dataclass
class EdinetIndexItem:
    ticker: str               # 4-digit
    doc_id: str
    title: str
    submit_date: Optional[dt.date]
    period_start: Optional[dt.date]
    period_end: Optional[dt.date]

def edinet_fetch_docs_by_day(session: requests.Session, day: dt.date, api_key: str) -> Optional[List[Dict[str, Any]]]:
    # EDINET documents list (type=2). Key is passed both via query parameter and headers for robustness.
    params = {"date": day.strftime("%Y-%m-%d"), "type": 2, "Subscription-Key": api_key}
    if not api_key:
        raise RuntimeError(
            "EDINET APIキーが未設定です。環境変数 EDINET_API_KEY を設定するか、--edinet-api-key で渡してください。"
        )
    headers = {
        "Subscription-Key": api_key,
        "Ocp-Apim-Subscription-Key": api_key,
    }

    last_err: Optional[str] = None
    for url in EDINET_DOC_URL_CANDIDATES:
        r = safe_get(
            session,
            url,
            params=params,
            headers=headers,
            source="edinet",
            must=False,
            log_tag="edinet_docs",
        )
        if not r:
            last_err = "no_response"
            continue
        try:
            j = r.json()
        except Exception as e:
            last_err = f"invalid_json:{type(e).__name__}:{e}"
            continue
        res = j.get("results") if isinstance(j, dict) else None
        meta_status = str(j.get("metadata", {}).get("status", "")) if isinstance(j, dict) else ""
        if isinstance(res, list) and (meta_status == "200" or meta_status == ""):
            return res
        last_err = f"missing_results_or_bad_status(status={meta_status})"

    raise SourceAccessError("edinet", EDINET_DOC_URL_CANDIDATES[0], last_err or "unreachable")


def edinet_build_index(session: requests.Session, tickers: set, lookback_days: int, api_key: str, log: List[List[Any]]) -> Dict[str, List[EdinetIndexItem]]:
    """
    Prefetch EDINET document list once and build ticker->items mapping.
    Notes:
    - EDINET secCode is usually 5 digits (e.g., 13010). We match prefix by 4-digit ticker.
    """
    index: Dict[str, List[EdinetIndexItem]] = {t: [] for t in tickers}
    today = dt.datetime.now(JST).date()

    for dd in range(0, lookback_days):
        day = today - dt.timedelta(days=dd)
        docs = edinet_fetch_docs_by_day(session, day, api_key)
        if not docs:
            continue
        for d in docs:
            sec = str(d.get("secCode") or "").strip()
            if not sec or len(sec) < 4:
                continue
            ticker = sec[:4]
            if ticker not in index:
                continue
            title = str(d.get("docDescription") or "")
            if not title:
                continue
            # Filter to likely financial filings that may include PL
            if not any(k in title for k in ["四半期", "半期", "有価証券報告書", "決算", "短信"]):
                continue
            doc_id = str(d.get("docID") or "").strip()
            if not doc_id:
                continue
            submit = safe_date_from_any(d.get("submitDateTime") or d.get("submitDate"))
            pstart = safe_date_from_any(d.get("periodStart"))
            pend = safe_date_from_any(d.get("periodEnd"))
            index[ticker].append(EdinetIndexItem(
                ticker=ticker, doc_id=doc_id, title=title,
                submit_date=submit, period_start=pstart, period_end=pend
            ))

    # Dedup by doc_id, keep latest submit_date
    for t, arr in index.items():
        arr.sort(key=lambda x: (x.submit_date or dt.date.min), reverse=True)
        seen = set()
        uniq = []
        for it in arr:
            if it.doc_id in seen:
                continue
            seen.add(it.doc_id)
            uniq.append(it)
        index[t] = uniq

    log.append(["", "", "edinet_index", f"lookback_days={lookback_days} tickers={len(tickers)}"])
    return index

def edinet_download_zip(session: requests.Session, doc_id: str, api_key: str) -> Optional[bytes]:
    """Download XBRL ZIP (type=1) from EDINET.

    EDINET is sometimes reachable via different base hosts. To avoid false negatives, we try:
      - https://api.edinet-fsa.go.jp
      - https://disclosure.edinet-fsa.go.jp

    Returns bytes on success, or None on failure (caller may fall back to TDNet/Kabutan).
    """
    params = {"type": 1, "Subscription-Key": api_key}
    headers = {}
    if api_key:
        headers["Subscription-Key"] = api_key
        headers["Ocp-Apim-Subscription-Key"] = api_key

    for base in ("https://api.edinet-fsa.go.jp", "https://disclosure.edinet-fsa.go.jp"):
        url = f"{base}/api/v2/documents/{doc_id}"
        try:
            r = safe_get(session, url, params=params, headers=headers, source="edinet", must=False, log_tag="edinet_zip", timeout=TIMEOUT)
        except Exception:
            continue
        ctype = (r.headers.get("content-type") or "").lower()
        if r.ok and (("zip" in ctype) or ("octet-stream" in ctype) or (len(r.content or b"") > 1000)):
            return r.content
    return None


def require_zip_bytes(source: str, url: str, content: bytes) -> bytes:
    """Validate that downloaded bytes look like a ZIP archive.

    Some endpoints may return status=200 but serve an HTML block page or error message.
    For TDNet/EDINET, we must treat that as a REQUIRED source failure and abort immediately.
    """
    b = content or b""
    if not b.startswith(b"PK"):
        dump_path = os.path.join(os.getcwd(), f"_{source}_nonzip_dump.bin")
        try:
            with open(dump_path, "wb") as f:
                f.write(b[:2_000_000])  # cap at 2MB
        except Exception:
            pass
        raise SourceAccessError(source, url, f"non_zip_content (dump={dump_path})")
    return b

def iter_html_candidates_from_zip(zip_bytes: bytes) -> Iterable[Tuple[str, str]]:
    """Iterate HTML candidates from an XBRL ZIP.

    IMPORTANT:
    - Prefer iXBRL attachments when available, BUT always include 'qualitative' HTML
      (e.g., TDNet 'qualitative.htm') because depreciation (減価償却費) is sometimes
      disclosed only in CF-note tables, especially for Q1/Q3.
    """
    try:
        z = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except Exception:
        return []
    try:
        names = z.namelist()
    except MemoryError:
        # r28: if memory is tight, skip parsing this ZIP rather than crashing the whole run
        try:
            z.close()
        except Exception:
            pass
        return []

    # Prefer iXBRL html first
    ix = [n for n in names if n.lower().endswith(("-ixbrl.htm", "-ixbrl.html", "-ixbrl.xhtml"))]

    # Always include qualitative html even if not iXBRL
    qual = [
        n for n in names
        if n.lower().endswith((".htm", ".html", ".xhtml"))
        and ("qualitative" in n.lower() or "定性的" in n)
    ]

    if not ix:
        ix = [n for n in names if n.lower().endswith((".htm", ".html", ".xhtml"))]

    # keep stable order with uniqueness
    cand_names: List[str] = []
    seen = set()
    for fn in (sorted(ix) + sorted(qual)):
        if fn in seen:
            continue
        seen.add(fn)
        cand_names.append(fn)

    for fn in cand_names:
        try:
            info = z.getinfo(fn)
            # Skip very large HTML files to avoid MemoryError (they are usually appendices).
            if getattr(info, "file_size", 0) and info.file_size > 6_000_000:
                continue
        except Exception:
            pass
        try:
            raw = z.read(fn)
        except Exception:
            continue
        txt = decode_bytes_guess(raw)
        # quick filter: must contain inline facts, or qualitative keyword
        low = txt.lower()
        if (
            ("contextref" in low and "name=" in low)
            or ("ix:" in low)
            or ("qualitative" in fn.lower())
            or ("定性的" in txt[:2000])
        ):
            yield fn, txt




def extract_securities_code_from_zip(zip_bytes: bytes) -> Optional[str]:
    """Best-effort extraction of a securities code from a TDNet/EDINET XBRL ZIP.

    - TDNet iXBRL often contains <ix:nonNumeric name="tse-ed-t:SecuritiesCode">....</ix:nonNumeric>
      (alphanumeric codes like 142A are represented as 142A + hidden '0' span).
    - Some EDINET docs contain jpdei_cor:SecurityCodeDEI or similar.
    - As a fallback, TDNet file names often embed code like '142A0' or '97600'.

    Returns normalized code string (e.g., '9760', '142A') or None if not detectable.
    """
    try:
        z = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except Exception:
        return None
    try:
        # 1) Look for inline or XML tag containing SecuritiesCode*
        for name in z.namelist():
            ln = name.lower()
            if not (ln.endswith(".xbrl") or ln.endswith(".xml") or ln.endswith(".htm") or ln.endswith(".html")):
                continue
            try:
                info = z.getinfo(name)
                if info.file_size > 2_000_000:
                    continue
                raw = z.read(name)
            except Exception:
                continue
            txt = raw.decode("utf-8", errors="ignore")

            # iXBRL / HTML: name="...:SecuritiesCode" or name="...:SecuritiesCodeDEI"
            m = re.search(r'name="[^\"]*:(?:SecurityCodeDEI|SecuritiesCodeDEI|SecuritiesCode|SecurityCode)"[^>]*>(.*?)</', txt, flags=re.I | re.S)
            if m:
                inner = re.sub(r"<[^>]+>", "", m.group(1) or "").strip()
                code = normalize_security_code(inner)
                if code:
                    return code

            # XML: <...SecuritiesCodeDEI>9760</...> or <...SecuritiesCode>142A</...>
            m2 = re.search(r"<(?:[^:>]+:)?(?:SecurityCodeDEI|SecuritiesCodeDEI|SecuritiesCode|SecurityCode)[^>]*>\s*([0-9A-Za-z]{3,6})\s*<", txt, flags=re.I)
            if m2:
                code = normalize_security_code(m2.group(1))
                if code:
                    return code

        # 2) Filename heuristic: TDNet often embeds '142A0' or '97600'
        for name in z.namelist():
            m3 = re.search(r"-(\d{3}[A-Z0-9])0-", name.upper())
            if m3:
                return m3.group(1)
            m4 = re.search(r"-(\d{4})0-", name)
            if m4:
                return m4.group(1)
    finally:
        try:
            z.close()
        except Exception:
            pass
    return None
    try:
        # 1) XBRL/IXBRL tag
        for name in z.namelist():
            ln = name.lower()
            if not (ln.endswith(".xbrl") or ln.endswith(".xml") or ln.endswith(".htm") or ln.endswith(".html")):
                continue
            try:
                info = z.getinfo(name)
                if info.file_size > 2_000_000:
                    continue
                raw = z.read(name)
            except Exception:
                continue
            txt = raw.decode("utf-8", errors="ignore")
            m = re.search(r"(?:jpdei_cor:)?SecuritiesCodeDEI[^0-9]{0,200}([0-9]{4})", txt)
            if not m:
                m = re.search(r"<[^>]*SecuritiesCodeDEI[^>]*>\s*([0-9]{4})\s*<", txt)
            if m:
                try:
                    return int(m.group(1))
                except Exception:
                    return None

        # 2) Filename heuristic: TDNet often embeds '15140' (4-digit + trailing 0) or '1514'
        for name in z.namelist():
            m2 = re.search(r"-(\d{4})0-", name)
            if m2:
                return int(m2.group(1))
            m3 = re.search(r"-(\d{4})-", name)
            if m3:
                return int(m3.group(1))
    finally:
        try:
            z.close()
        except Exception:
            pass
    return None

def parse_zip_metrics(zip_bytes: bytes, prefer_consolidated: bool, log: List[List[Any]], who: str, expected_ticker: Optional[str] = None) -> Tuple[Dict[str, float], Dict[str, float]]:
    """
    Parse EDINET/TDNet iXBRL ZIP -> metric values (cumulative & single).

    v5_25 FIX:
    - Collect facts from ALL candidate iXBRL/HTML files in the ZIP and choose the best fact per metric.
    - Selection order is driven by:
        (i) consolidated preference, (ii) avoiding Member/segment contexts,
        (iii) context suitability (cum vs single), (iv) tag-name priority,
        (v) file preference (financial statements > summary).
    - Keep units in million JPY (百万円) using iXBRL scale/unitRef conversion (see extract_inline_facts).
    """
    # Optional mismatch check (DEI)
    if expected_ticker:
        try:
            code_in_zip = extract_securities_code_from_zip(zip_bytes)
            # Only reject when we can confidently extract the code and it mismatches.
            if code_in_zip and not securities_code_matches_expected(code_in_zip, expected_ticker):
                return {}, {}
        except Exception:
            pass

    qual_metrics: Dict[str, float] = {}
    all_facts: List[Tuple[str, str, float, str]] = []  # (qname, ctx, value_million, filename)

    for fn, txt in iter_html_candidates_from_zip(zip_bytes):
        # qualitative tables (TDNet) can contain depreciation/gross/SGA
        try:
            if fn and ('qualitative' in fn.lower() or '定性的' in fn):
                qm = extract_qualitative_table_metrics(txt)
                if qm:
                    qual_metrics.update(qm)
        except Exception:
            pass

        try:
            facts = extract_inline_facts(txt)
        except Exception as e:
            log.append([None, None, 'zip_parse_error', None, None, None, None, None, None, who, True, None, f'{who}:extract_inline_facts_error:{type(e).__name__}:{e}'])
            facts = []
        if not facts:
            continue
        for qn, ctx, v in facts:
            if not qn or not ctx:
                continue
            all_facts.append((qn, ctx, float(v), fn))

    if not all_facts:
        return {}, {}

    # Detect IFRS-ish filings (affects tag preference for 'keijo')
    # NOTE: do NOT treat jpigp (Japan GAAP namespace) as IFRS.
    doc_ifrs = any(('ifrs' in (qn or '').lower()) or (':ifrs' in (qn or '').lower()) for qn, _, _, _ in all_facts)

    def file_flags(fn: str) -> Tuple[int, int]:
        """Lower is better. Prefer statement-like files; penalize obvious summaries."""
        l = (fn or '').lower()
        summary = 1 if any(k in l for k in ('summary', 'gaiyou', 'overview', 'highlight', 'gaikyo')) else 0
        fs = 0 if ('/0105' in l or '_0105' in l or 'honbun' in l or 'publicdoc' in l) else 1
        return (summary, fs)
    # v16/r24: Synthetic CAPEX (PPE acquisition) sum for issuers where CF splits PPE into components.
    #
    # r24 extension (bugfix):
    # - Some filers split '有形固定資産の取得による支出' into PPE + other components such as
    #     * 賃貸固定資産の取得による支出 (rental fixed assets)
    #     * その他の固定資産の取得による支出
    # - If Excel already has the correct *sum* within tolerance, we must not create false WARN.
    #
    # We synthesize a context-aligned sum when we see PPE + (rent/other) in the same context.
    capex_syn: List[Tuple[str, str, float, str]] = []
    try:
        tmp: Dict[str, Dict[str, Tuple[float, str]]] = {}  # ctx -> key -> (v, fn)
        for qn, ctx, v, fn in all_facts:
            lq = (qn or '').lower()
            if 'invcf' not in lq:
                continue

            if 'paymentsforpurchasesofexplorationandevaluationassets' in lq:
                tmp.setdefault(ctx, {})['explor'] = (float(v), fn)
            elif 'paymentsforpurchasesofdevelopmentandproductionassets' in lq:
                tmp.setdefault(ctx, {})['dev'] = (float(v), fn)

            # Base PPE
            elif 'purchaseofpropertyplantandequipment' in lq or 'paymentsforpurchaseofpropertyplantandequipment' in lq:
                tmp.setdefault(ctx, {})['ppe'] = (float(v), fn)

            # Rental / real estate for rent
            elif ('realestatesforrent' in lq or 'realestateforrent' in lq or 'rentalfixedassets' in lq or 'noncurrentassetsforrent' in lq) and ('purchase' in lq or 'paymentsfor' in lq):
                tmp.setdefault(ctx, {})['rent'] = (float(v), fn)

            # Other noncurrent / fixed assets
            elif ('othernoncurrentassets' in lq or 'otherfixedassets' in lq) and ('purchase' in lq or 'paymentsforpurchase' in lq):
                tmp.setdefault(ctx, {})['other'] = (float(v), fn)

            # Some JP-GAAP iXBRL uses IncreaseInOtherAssetsInvCF for the same line.
            elif 'increaseinotherassets' in lq and 'invcf' in lq:
                tmp.setdefault(ctx, {})['other'] = (float(v), fn)

        def _best_fn(fns: List[str]) -> str:
            return min(fns, key=file_flags) if fns else ''

        for ctx, d in tmp.items():
            # Exploration + development + PPE (legacy INPEX style)
            if all(k in d for k in ('explor', 'dev', 'ppe')):
                sum_v = float(d['explor'][0]) + float(d['dev'][0]) + float(d['ppe'][0])
                fns = [d['explor'][1], d['dev'][1], d['ppe'][1]]
                capex_syn.append(('SYNTHCapexPpeSum', ctx, sum_v, _best_fn(fns)))
                continue

            # PPE + optional components (rent/other)
            if 'ppe' in d and not (('explor' in d) or ('dev' in d)):
                parts: List[Tuple[str, Tuple[float, str]]] = [('ppe', d['ppe'])]
                if 'rent' in d:
                    parts.append(('rent', d['rent']))
                if 'other' in d:
                    parts.append(('other', d['other']))
                if len(parts) >= 2:
                    sum_v = sum(float(vv[0]) for _, vv in parts)
                    fns = [vv[1] for _, vv in parts]
                    capex_syn.append(('SYNTHCapexPpeSum', ctx, sum_v, _best_fn(fns)))
    except Exception:
        capex_syn = []

    if capex_syn:
        all_facts.extend(capex_syn)

    def value_penalty(metric: str, v: float) -> int:
        """Heuristic penalty. Lower is better."""
        try:
            fv = float(v)
        except Exception:
            return 0
        if metric == 'uriage' and fv < 0:
            return 5
        if metric == 'genka' and fv < 0:
            return 2
        if metric in ('gross', 'sga') and fv < 0:
            return 1
        return 0

    def name_priority(metric: str, qname: str) -> int:
        """Lower is better."""
        l = (qname or '').lower()

        if metric == 'uriage':
            penalty = 0
            if ('netsalesof' in l) or ('revenueof' in l) or ('operatingrevenueof' in l) or ('totalrevenueof' in l):
                penalty += 4
            if any(k in l for k in (
                'completedconstruction', 'constructioncontract', 'constructioncontracts',
                'uncompletedconstruction', 'construction'
            )):
                penalty += 5
            if any(k in l for k in ('segment', 'reportablesegment', 'businesssegment')):
                penalty += 2

            base = 9
            if doc_ifrs:
                if 'revenue2ifrs' in l:
                    base = 0
                elif 'revenueifrs' in l:
                    base = 1
                elif 'operatingrevenue' in l:
                    base = 2
                elif 'totalrevenue' in l:
                    base = 3
                elif 'revenue' in l and 'fromexternalcustomers' not in l:
                    base = 4
                elif 'netsales' in l and 'summary' not in l and 'netsalesof' not in l:
                    base = 5
                elif 'netsales' in l:
                    base = 6
                elif 'sales' in l:
                    base = 7
                return min(9, base + penalty)

            if 'netsales' in l and 'summary' not in l and 'netsalesof' not in l:
                base = 0
            elif 'netsales' in l and 'netsalesof' not in l:
                base = 1
            elif 'netsalesof' in l:
                base = 4
            elif 'operatingrevenues' in l:
                base = 5
            elif 'operatingrevenue' in l:
                base = 6
            elif 'totaloperatingrevenue' in l:
                base = 6
            elif 'revenues' in l:
                base = 7
            elif 'revenue' in l and 'fromexternalcustomers' not in l:
                base = 7
            elif 'sales' in l:
                base = 8
            return min(9, base + penalty)

        if metric == 'keijo':
            # v18+v22:
            # - Many JP-GAAP filers use OrdinaryIncome as 経常利益 (ordinary profit), and we still want to
            #   prefer it over IFRS-like "profit before tax" when both are present (to avoid false WARNs).
            # - However, some industries (notably banks) may disclose BOTH:
            #     * OrdinaryIncome (経常収益/収入に相当することがある; revenue-like)
            #     * OrdinaryIncomeLoss / OrdinaryProfitLoss (経常利益)
            #   under the SAME contextRef. In that case we must prefer the *Loss/ProfitLoss* concept.
            #
            # Rule:
            #   OrdinaryIncomeLoss / OrdinaryProfitLoss : base=0 (best)
            #   OrdinaryIncome (ambiguous)              : base=1
            #   ProfitBeforeTax variants                : base=2
            # Plus: penalize "...SummaryOfBusinessResults" to avoid KPI tables.
            base = 9
            if any(k in l for k in ('ordinaryincomeloss', 'ordinaryprofitloss')):
                base = 0
            elif 'ordinaryprofit' in l:
                base = 0
            elif 'ordinaryincome' in l:
                base = 1
            elif any(k in l for k in (
                'profitlossbeforetax',
                'profitbeforetax',
                'profitbeforeincometaxes',
                'incomebeforetax',
                'incomebeforeincometaxes',
            )):
                base = 2
            if 'summaryofbusinessresults' in l:
                base += 4
            return base

        if metric == 'opcf':
            # v18: Prefer CF statement figures over "...SummaryOfBusinessResults" variants.
            base = 9
            if 'netcashprovidedbyusedinoperatingactivities' in l:
                base = 0
            if 'summaryofbusinessresults' in l:
                base += 4
            return base

        if metric == 'saishu':
            # Prefer profit attributable to owners of parent from the *financial statements*
            # over "...SummaryOfBusinessResults" (KPI tables) variants.
            base = 9
            if 'profitlossattributabletoownersofparent' in l:
                base = 0
            elif 'netincomelossattributabletoownersofparent' in l:
                base = 1
            elif 'profitloss' in l or 'netincome' in l:
                base = 5
            if 'summaryofbusinessresults' in l:
                base += 4
            return base

        if metric == 'gross':
            if 'operatinggrossprofit' in l or 'grossoperatingprofit' in l:
                if any(k in l for k in ['segment', 'summary', 'bysegment', 'member']):
                    return 5
                return 0
            if 'grossprofit' in l:
                if any(k in l for k in ['completedconstruction', 'construction', 'oncompletedconstruction']):
                    return 5
                if any(k in l for k in ['segment', 'summary', 'bysegment', 'member']):
                    return 6
                if l.endswith('grossprofit') or l.endswith(':grossprofit') or 'grossprofittotal' in l:
                    return 0
                if 'grossprofitloss' in l:
                    return 1
                return 2
            if any(k in l for k in ['grossoperatingrevenue', 'grossoperatingrevenues', 'netoperatingrevenue', 'netoperatingrevenues', 'pureoperatingrevenue', 'pureoperatingrevenues']):
                if any(k in l for k in ['segment', 'summary', 'bysegment', 'member']):
                    return 7
                return 3
            return 9

        if metric == 'rnd':
            # v16: Prefer R&D expense on the face of P/L if available; otherwise use note disclosure.
            # Priority: P/L line item > "included in SG&A" note > "R&D activities" (often segment split) > others.
            if 'researchanddevelopment' in l:
                if 'includedin' in l:
                    return 1
                if 'researchanddevelopmentactivities' in l:
                    return 5
                if any(k in l for k in ('expense', 'expenses', 'cost', 'costs')):
                    return 0
                return 6
            return 9

        if metric == 'capex_ppe':
            # r24: Prefer synthetic sum (split disclosures), then standard PPE purchase.
            if 'synthcapexppe' in l:
                return 0
            if 'purchaseofpropertyplantandequipment' in l and 'invcf' in l:
                return 1
            if 'paymentsforpurchaseofpropertyplantandequipment' in l:
                return 2
            # Rental / real estates for rent acquisition
            if ('realestatesforrent' in l or 'realestateforrent' in l or 'rentalfixedassets' in l or 'noncurrentassetsforrent' in l) and ('purchase' in l or 'paymentsfor' in l):
                return 4
            # Other noncurrent assets acquisitions sometimes need to be summed into capex bucket
            if ('purchaseofothernoncurrentassets' in l or 'paymentsforpurchaseofothernoncurrentassets' in l or 'increaseinotherassets' in l) and 'invcf' in l:
                return 5
            if 'paymentsforpurchasesofdevelopmentandproductionassets' in l and 'invcf' in l:
                return 6
            if 'paymentsforpurchasesofexplorationandevaluationassets' in l and 'invcf' in l:
                return 6
            return 9

        if metric == 'sga':
            # v16: Prefer "供給販売費、販売費及び一般管理費合計" when available (SupplyAndSales...).
            if 'supplyandsales' in l and 'sellinggeneralandadministrativeexpenses' in l:
                if any(k in l for k in ['segment', 'summary', 'member']):
                    return 3
                return 0
            if 'sellinggeneralandadministrativeexpenses' in l:
                if any(k in l for k in ['segment', 'summary', 'member']):
                    return 4
                return 1
            if 'sgaexpenses' in l:
                return 2
            if 'operatingexpenses' in l:
                return 3
            if 'generalandadministrativeexpenses' in l or 'generaladministrativeexpenses' in l or 'administrativeexpenses' in l:
                return 3
            return 9

        if metric == 'genka':
            if 'accumulateddepreciation' in l:
                return 99
            cf_keys = ('opecf', 'cashflow', 'cashflows', 'operatingactivities', 'cashflowsfromoperatingactivities')
            if any(k in l for k in cf_keys):
                if 'depreciation' in l and ('amortization' in l or 'amortisation' in l):
                    return 0
                return 1
            return 50

        return 5

    def score_ctx(ctx: str, kind: str, metric: str, qname: str) -> Tuple[int, int, int, int, int, int]:
        """Lower is better.

        r24 notes:
        - Member/segment penalty must NOT mistakenly treat NonConsolidatedMember as ConsolidatedMember (substring trap).
        - Some filings omit underscores in contextRef; use substring guards rather than token-only logic.
        - For non-consolidated statements, prefer plain ProfitLoss over AttributableToOwnersOfParent.
        """
        l = (ctx or '').lower()
        noncon = 1 if 'nonconsolidated' in l else 0
        cons_score = noncon if prefer_consolidated else (1 - noncon)

        # Member/segment context penalty
        member_penalty = 0
        if 'nonconsolidatedmember' in l:
            member_penalty = 1
        # r30: include plural segment tokens (e.g., OtherReportableSegmentsMember),
        # which previously slipped through and caused segment values to win.
        elif ('segment' in l) and (('member' in l) or ('axis' in l)):
            member_penalty = 1
        else:
            # ConsolidatedMember (and generic ResultMember) are treated as default axis
            member_penalty = 0

        cur_score = 0 if 'current' in l else 1
        forecast_penalty = 1 if ('forecast' in l or 'planmember' in l) else 0

        is_cum_ctx = any(k in l for k in (
            'currentytdduration',
            'currentyeartodateduration',
            'interimperiodduration',
            'currentinterimduration',
            'currentinterimyeartodateduration',
            'currentaccumulated',
            'accumulated',
            'ytd',
            'year',
        ))
        is_single_ctx = any(k in l for k in (
            'currentquarterduration',
            'quarterduration',
            'currentquarter',
            'firstquarter',
            'secondquarter',
            'thirdquarter',
            'fourthquarter',
            'q1duration',
            'q2duration',
            'q3duration',
            'q4duration',
        ))

        if kind == 'cum':
            if is_cum_ctx:
                kind_score = 0
            elif is_single_ctx:
                kind_score = 2
            else:
                kind_score = 1
        else:
            if is_single_ctx:
                kind_score = 0
            elif is_cum_ctx:
                kind_score = 2
            else:
                kind_score = 1

        tag_ctx_penalty = 0
        lq = (qname or '').lower()
        if metric == 'saishu' and ('nonconsolidated' in l) and ('attributabletoownersofparent' in lq):
            tag_ctx_penalty = 1

        return (cons_score, forecast_penalty, member_penalty, cur_score, kind_score, tag_ctx_penalty)

    def is_trusted_depr_candidate(qname: str, ctx: str, fn: str) -> bool:
        """Accept depreciation only from cash-flow-note style disclosures.

        User requirement:
        - Depreciation (genka) must NOT come from P/L line items.
        - Use cash flow statement / cash flow note disclosures only.

        In practice, valid inline facts are usually OpeCF-style concepts on EDINET.
        TDNet often discloses the correct number only in qualitative CF-note tables;
        those are handled separately via qual_metrics and should not be blocked by
        unrelated inline facts such as DepreciationSGA or AccumulatedDepreciation.
        """
        lq = (qname or '').lower()
        lc = (ctx or '').lower()
        lf = (fn or '').lower()

        # Explicitly reject known non-CF depreciation concepts.
        if any(k in lq for k in (
            'accumulateddepreciation',
            'depreciationsga',
            'depreciationnoe',
            'depreciationsegmentinformation',
            'reserveforadvanceddepreciation',
            'reversalofreserveforadvanceddepreciation',
        )):
            return False

        cf_keys = ('opecf', 'cashflow', 'cashflows', 'operatingactivities')
        if any(k in lq for k in cf_keys):
            return True
        if any(k in lc for k in cf_keys):
            return True
        if ('cashflow' in lf or 'cashflows' in lf) and 'depreciation' in lq:
            return True
        return False

    def pick_best(metric: str, kind: str) -> Optional[Tuple[float, str, str, str]]:
        best = None
        best_score = None
        for qn, ctx, v, fn in all_facts:
            m = metric_from_qname(qn)
            if m != metric:
                continue
            if metric == 'genka' and not is_trusted_depr_candidate(qn, ctx, fn):
                continue
            sc = score_ctx(ctx, kind, metric, qn) + (name_priority(metric, qn),) + (value_penalty(metric, v),) + file_flags(fn)
            if best_score is None or sc < best_score:
                best_score = sc
                best = (v, qn, ctx, fn)
        return best

    metrics = {'uriage', 'keijo', 'saishu', 'gross', 'sga', 'genka', 'opcf', 'ad', 'rnd', 'capex_ppe'}
    metric_cum: Dict[str, float] = {}
    metric_single: Dict[str, float] = {}

    for mkey in metrics:
        b = pick_best(mkey, 'cum')
        if b is not None:
            v, qn, ctx, fn = b
            metric_cum[mkey] = float(v)
        b2 = pick_best(mkey, 'single')
        if b2 is not None:
            v, qn, ctx, fn = b2
            metric_single[mkey] = float(v)

    # Financial-format fallback for gross:
    # Some TDNet iXBRL uses "営業総利益/純営業収益/経常利益" style concepts instead of GrossProfit.
    if 'gross' not in metric_cum:
        has_financial_terms = any(
            any(k in (qn or '').lower() for k in (
                'operatingrevenue', 'operatingrevenues',
                'netoperatingrevenue', 'netoperatingrevenues',
                'grossoperatingrevenue', 'grossoperatingrevenues',
                'ordinaryincome', 'ordinaryprofit',
            ))
            for qn, _, _, _ in all_facts
        )

        def gross_proxy_rank(qname: str) -> Optional[int]:
            l = (qname or '').lower()
            if 'summaryofbusinessresults' in l:
                return None
            if 'operatinggrossprofit' in l or 'grossoperatingprofit' in l:
                return 0
            if any(k in l for k in (
                'grossoperatingrevenue', 'grossoperatingrevenues',
                'netoperatingrevenue', 'netoperatingrevenues',
                'pureoperatingrevenue', 'pureoperatingrevenues',
            )):
                return 1
            if has_financial_terms and any(k in l for k in ('ordinaryincomeloss', 'ordinaryprofitloss', 'ordinaryprofit')):
                return 2
            if has_financial_terms and ('ordinaryincome' in l) and ('bnk' not in l):
                return 6
            return None

        def pick_gross_proxy(kind: str) -> Optional[Tuple[float, str, str, str]]:
            best = None
            best_score = None
            for qn, ctx, v, fn in all_facts:
                rank = gross_proxy_rank(qn)
                if rank is None:
                    continue
                sc = score_ctx(ctx, kind, 'gross', qn) + (rank,) + (value_penalty('gross', v),) + file_flags(fn)
                if best_score is None or sc < best_score:
                    best_score = sc
                    best = (v, qn, ctx, fn)
            return best

        b = pick_gross_proxy('cum')
        if b is not None:
            v, qn, ctx, fn = b
            metric_cum['gross'] = float(v)
        b2 = pick_gross_proxy('single')
        if b2 is not None:
            v, qn, ctx, fn = b2
            metric_single['gross'] = float(v)

    # Derived fallback for gross when direct/proxy tags are unavailable.
    # Priority:
    #   1) InsuranceServiceResult (insurers, IFRS17)
    #   2) OperatingRevenue - CostOfSales - CostOfFinancingOperations
    #   3) OperatingRevenue - FinanceCosts - OtherOperatingExpenses
    #   4) Revenue - CostOfSales
    # These are used only when gross is still missing.
    def pick_component(kind: str, ranker, ctx_exact: Optional[str] = None) -> Optional[Tuple[float, str, str, str]]:
        best = None
        best_score = None
        for qn, ctx, v, fn in all_facts:
            if ctx_exact is not None and ctx != ctx_exact:
                continue
            rank = ranker(qn)
            if rank is None:
                continue
            sc = score_ctx(ctx, kind, 'gross', qn) + (rank,) + (value_penalty('gross', v),) + file_flags(fn)
            if best_score is None or sc < best_score:
                best_score = sc
                best = (v, qn, ctx, fn)
        return best

    def rank_revenue_like(qname: str) -> Optional[int]:
        l = (qname or '').lower()
        if 'summaryofbusinessresults' in l:
            return None
        if any(k in l for k in ('fromexternalcustomers', 'intersegment')):
            return None
        if 'totalnetrevenues' in l:
            return 0
        if 'totaloperatingrevenue' in l or 'totaloperatingrevenues' in l:
            return 1
        if 'operatingrevenue' in l or 'operatingrevenues' in l:
            return 2
        if 'netsales' in l:
            return 3
        if 'revenueifrs' in l or l.endswith(':revenue') or l.endswith('revenue'):
            return 4
        return None

    def rank_cost_of_sales_like(qname: str) -> Optional[int]:
        l = (qname or '').lower()
        if 'summaryofbusinessresults' in l:
            return None
        if 'costofsales' in l:
            return 0
        if 'costofrevenues' in l or 'costofrevenue' in l:
            return 1
        if 'costofgoodssold' in l:
            return 2
        return None

    def rank_insurance_service_result(qname: str) -> Optional[int]:
        l = (qname or '').lower()
        if 'summaryofbusinessresults' in l:
            return None
        if 'insuranceserviceresult' in l:
            return 0
        if 'insuranceserviceprofitloss' in l or 'insuranceserviceprofit' in l:
            return 1
        return None

    def rank_cost_of_financing_operations(qname: str) -> Optional[int]:
        l = (qname or '').lower()
        if 'summaryofbusinessresults' in l:
            return None
        if 'costoffinancingoperations' in l:
            return 0
        if 'financialexpensesrelatedtofinancialbusiness' in l:
            return 1
        return None

    def rank_finance_cost_like(qname: str) -> Optional[int]:
        l = (qname or '').lower()
        if 'summaryofbusinessresults' in l:
            return None
        if 'financecosts' in l:
            return 0
        if 'financialexpenses' in l:
            return 1
        if 'otherfinancecosts' in l or 'otherfinancialexpense' in l:
            return 2
        return None

    def rank_other_operating_expense_like(qname: str) -> Optional[int]:
        l = (qname or '').lower()
        if 'summaryofbusinessresults' in l:
            return None
        if 'otherexpensesoperatingexpenses' in l:
            return 0
        if 'otheroperatingexpenses' in l:
            return 1
        if 'otherexpensesifrs' in l:
            return 2
        return None

    def fill_gross_by_formula(kind: str, bucket: Dict[str, float]) -> None:
        if 'gross' in bucket:
            return
        ins = pick_component(kind, rank_insurance_service_result)
        if ins is not None:
            bucket['gross'] = float(ins[0])
            return

        rev = pick_component(kind, rank_revenue_like)
        cogs = pick_component(kind, rank_cost_of_sales_like)
        if cogs is not None:
            ctx = cogs[2]
            rev_same = pick_component(kind, rank_revenue_like, ctx_exact=ctx)
            finbiz_same = pick_component(kind, rank_cost_of_financing_operations, ctx_exact=ctx)
            if rev_same is not None and finbiz_same is not None:
                bucket['gross'] = float(rev_same[0]) - float(cogs[0]) - float(finbiz_same[0])
                return
            if rev_same is not None:
                bucket['gross'] = float(rev_same[0]) - float(cogs[0])
                return

        oth = pick_component(kind, rank_other_operating_expense_like)
        if oth is not None:
            ctx = oth[2]
            rev_same = pick_component(kind, rank_revenue_like, ctx_exact=ctx)
            fin_same = pick_component(kind, rank_finance_cost_like, ctx_exact=ctx)
            if rev_same is not None and fin_same is not None:
                bucket['gross'] = float(rev_same[0]) - float(fin_same[0]) - float(oth[0])
                return

        if rev is not None and cogs is not None:
            bucket['gross'] = float(rev[0]) - float(cogs[0])
            return

    fill_gross_by_formula('cum', metric_cum)
    fill_gross_by_formula('single', metric_single)

    if qual_metrics:
        # For depreciation, CF-note table disclosure is the authoritative source.
        # Always prefer it over inline facts so P/L/BS depreciation concepts never win.
        if 'genka' in qual_metrics:
            metric_cum['genka'] = float(qual_metrics['genka'])
            metric_single.pop('genka', None)
        for k in ('gross', 'sga'):
            if k not in metric_cum and k in qual_metrics:
                metric_cum[k] = float(qual_metrics[k])
    return metric_cum, metric_single


def sanitize_single_vs_cum(metric_cum: dict, metric_single: dict, quarter_no: int, tol: float = 2.0) -> None:
    """Remove obviously misclassified 'single' values that are actually cumulative.

    TDNet/EDINET iXBRL sometimes expose quarter-duration contexts whose numeric fact equals the
    YTD cumulative value (esp. revenue). For Q2+ this should NOT happen; if it does, it causes
    left-side single-quarter cells to be filled with cumulative values.

    Rule: for Q2+ only, if a metric appears in both cum and single and they are within tol,
    drop the single entry and rely on cumulative-diff derivation.
    """
    try:
        q = int(quarter_no or 0)
    except Exception:
        q = 0
    if q <= 1:
        return
    drop = []
    for k, sv in list(metric_single.items()):
        if k not in metric_cum:
            continue
        cv = metric_cum.get(k)
        try:
            if cv is None or sv is None:
                continue
            if abs(float(sv) - float(cv)) <= (tol + 1e-9):
                drop.append(k)
        except Exception:
            continue
    for k in drop:
        metric_single.pop(k, None)

# -----------------------
# Kabutan
# -----------------------
def kabutan_fetch_html(session: requests.Session, ticker: str) -> Optional[str]:
    """Fetch Kabutan finance page (REQUIRED source) with WAF-aware validation.

    We must not silently accept bot-detection/blocked HTML pages (status=200 but missing finance blocks),
    because that would lead to 'no data' rows even though Kabutan is reachable in a browser.
    This uses the same marker-based validation + curl_cffi fallback as TDNet preflight.
    """
    url = f"https://kabutan.jp/stock/finance?code={ticker}"
    html = fetch_html_required(
        session,
        url,
        source="kabutan",
        ok_markers=KABUTAN_OK_MARKERS,
        timeout=TIMEOUT,
        log_tag="kabutan_finance",
        referer="https://kabutan.jp/",
    )
    return html

@dataclass
class KabutanRow:
    """One row parsed from Kabutan finance tables."""
    period_end_ym: str                 # "YYYY/MM" (period END)
    kind: str                          # "single" or "cum" or "annual"
    metric: Dict[str, float]           # already in million JPY
    raw_label: str = ""
    duration_months: Optional[int] = None   # 3/6/9/12 when label is a range


def parse_period_label_to_range(label: str) -> Tuple[Optional[str], Optional[int], Optional[int], Optional[int]]:
    """
    Parse Kabutan period labels like:
      - "24.07-09" (3m, end=2024/09)
      - "25.10-03" (6m crossing year, end=2026/03)
      - "2025.03"  (annual, end=2025/03)
    Returns:
      (end_ym "YYYY/MM", start_month, end_month, duration_months)
    """
    if not label:
        return (None, None, None, None)
    t = label.replace("\xa0", " ").replace("\u3000", " ").strip()
    t = re.sub(r"^\s*[予]\s*", "", t)  # drop forecast mark if present
    # quarterly/half-year range: YY.MM-MM
    m = re.search(r"(\d{2,4})\.(\d{1,2})-(\d{1,2})", t)
    if m:
        yy = int(m.group(1))
        sm = int(m.group(2))
        em = int(m.group(3))
        if yy < 100:
            sy = 2000 + yy
        else:
            sy = yy
        ey = sy if em >= sm else (sy + 1)
        dur = ((em - sm) % 12) + 1
        end_ym = f"{ey}/{em:02d}"
        return (end_ym, sm, em, dur)

    # annual: YYYY.MM or YY.MM
    m2 = re.search(r"(\d{2,4})\.(\d{1,2})\b", t)
    if m2:
        yy = int(m2.group(1))
        mm = int(m2.group(2))
        if yy < 100:
            yy = 2000 + yy
        end_ym = f"{yy}/{mm:02d}"
        return (end_ym, None, mm, 12)

    return (None, None, None, None)



def kabutan_extract_active_q(soup: BeautifulSoup) -> Optional[int]:
    """Extract active quarter number (1..4) from Kabutan's 3m actual section.

    Kabutan pages can contain multiple 'kessan_s_bar' images (3m actual, cumulative, etc.).
    To avoid quarter-shift bugs, we anchor to the '3ヵ月決算【実績】' menu and read the
    bar image immediately following that menu.
    """
    try:
        menu = None

        # Prefer the 3m actual block near the anchor.
        a = soup.select_one("a#shihanki_name")
        if a is not None:
            menu = a.find_next("div", class_=lambda c: c and ("cap1gyousekishuusei" in c.split()) and ("fin_menu" in c.split()))
            if menu is not None:
                t = menu.get_text(" ", strip=True)
                if ("3ヵ月決算" not in t) or ("実績" not in t):
                    menu = None

        if menu is None:
            for div in soup.select("div.cap1gyousekishuusei.fin_menu"):
                t = div.get_text(" ", strip=True)
                if ("3ヵ月決算" in t) and ("実績" in t):
                    menu = div
                    break

        if menu is None:
            return None

        # The bar image is typically in the next div.cap2.cap2s
        bar_div = menu.find_next("div", class_=lambda c: c and ("cap2" in c.split()) and ("cap2s" in c.split()))
        img = None
        if bar_div is not None:
            img = bar_div.find("img", src=re.compile(r"kessan_s_bar"))
        if img is None:
            img = menu.find_next("img", src=re.compile(r"kessan_s_bar"))
        if img is None:
            return None

        src = img.get("src") or ""
        qs = re.findall(r"(\d)Q", src)
        if not qs:
            return None
        qv = int(qs[-1])
        return qv if (1 <= qv <= 4) else None
    except Exception:
        return None


def kabutan_parse(html: str) -> Tuple[Optional[int], List[KabutanRow], List[KabutanRow], List[KabutanRow]]:
    """Parse Kabutan performance tables.

    v5_31 policy:
      - Use ONLY '3ヵ月決算【実績】' rows (single-quarter actuals) from Kabutan for SINGLE-quarter values.
      - Do NOT use '第n四半期累計決算【実績】' tables from Kabutan (avoid cumulative-table inconsistencies).
      - Annual rows may be used only as Q4 cumulative fallback and FY-end-month hint.
    """
    soup = BeautifulSoup(html, "html.parser")

    def parse_one_table(tbl: Any, kind_hint: str) -> List[KabutanRow]:
        out: List[KabutanRow] = []
        try:
            rows = tbl.select("tr")
        except Exception:
            return out
        if not rows:
            return out

        # header row: detect column indices by header text
        header = [c.get_text(strip=True).replace(" ", "") for c in rows[0].find_all(["th", "td"])]
        col_map: Dict[str, int] = {}
        for i, h in enumerate(header):
            if not h:
                continue
            if "売上高" in h:
                col_map["uriage"] = i
            if "経常" in h:
                col_map["keijo"] = i
            if ("最終" in h) or ("純利益" in h) or ("当期利益" in h):
                col_map["saishu"] = i
            if ("売上原価" in h) or (h == "原価"):
                col_map["genka"] = i
            if ("売上総利益" in h) or (h == "粗利"):
                col_map["gross"] = i
            if "販管費" in h:
                col_map["sga"] = i
            if ("減価償却" in h) or ("償却" in h):
                col_map["depr"] = i

        for tr in rows[1:]:
            cells = tr.find_all(["th", "td"])
            if len(cells) < 2:
                continue
            head0 = cells[0].get_text(" ", strip=True)
            if not head0:
                continue
            # skip forecast rows
            if ("予" in head0) or head0.strip().startswith("予") or ("予想" in head0):
                continue

            pr = parse_period_label_to_range(head0)
            if not pr:
                continue
            period_end_ym, _sm, _em, dur_m = pr

            # If table is known to be 3m (single-quarter) and label is plain 'YYYY.MM' (no range),
            # interpret it as quarter-end (3 months) instead of annual.
            if kind_hint == "single" and ("-" not in head0) and dur_m == 12:
                dur_m = 3

            metric: Dict[str, float] = {}
            for key, idx in col_map.items():
                if idx < len(cells):
                    v = parse_num(cells[idx].get_text(" ", strip=True))
                    if v is not None:
                        metric[key] = float(v)
            if metric:
                out.append(KabutanRow(raw_label=head0, period_end_ym=period_end_ym, duration_months=dur_m, metric=metric, kind=kind_hint))
        return out

    # annual (実績)
    annual_rows: List[KabutanRow] = []
    for div in soup.select("div.fin_year_result, div.fin_year_result_d"):
        for tbl in div.select("table"):
            annual_rows.extend(parse_one_table(tbl, kind_hint="annual"))

    # quarterly (3m 実績): parse tables in fin_quarter_result(_d) and keep ONLY dur==3
    single_rows_all: List[KabutanRow] = []
    for div in soup.select("div.fin_quarter_result, div.fin_quarter_result_d"):
        for tbl in div.select("table"):
            single_rows_all.extend(parse_one_table(tbl, kind_hint="single"))
    single_rows = [rr for rr in single_rows_all if rr.duration_months == 3]

    # Kabutan cumulative tables are intentionally not used in v5_31.
    cum_rows: List[KabutanRow] = []

    # FY end month hint
    fy_end_m_hint: Optional[int] = None
    # Prefer: infer FY end month using the active quarter indicator ("1Q..4Q") shown on Kabutan 3m table bar.
    # This is more reliable than month-mode heuristics and prevents quarter-shift bugs when FY end month is ambiguous.
    try:
        active_q: Optional[int] = kabutan_extract_active_q(soup)
        if active_q is not None and single_rows:
            def _ym_key(ym: str) -> int:
                try:
                    y, m = ym.split("/")
                    return int(y) * 100 + int(m)
                except Exception:
                    return -1
            latest = max((rr for rr in single_rows if rr.period_end_ym), key=lambda rr: _ym_key(rr.period_end_ym))
            m_end = int(latest.period_end_ym.split("/")[1])
            cand = ((m_end + 3 * (4 - active_q) - 1) % 12) + 1
            months = [int(rr.period_end_ym.split("/")[1]) for rr in single_rows if rr.period_end_ym]
            if months:
                allowed = {cand, ((cand - 3 - 1) % 12) + 1, ((cand - 6 - 1) % 12) + 1, ((cand - 9 - 1) % 12) + 1}
                sc = sum(1 for mm in months if mm in allowed)
                if sc >= max(2, int(len(months) * 0.6)):
                    fy_end_m_hint = cand
    except Exception:
        pass

    if fy_end_m_hint is None:
        try:
            # prefer latest ACTUAL annual row (ignore forecast)
            annual_ends = []
            for rr in annual_rows:
                if rr.raw_label.strip().startswith("予"):
                    continue
                if rr.duration_months == 12 and rr.period_end_ym:
                    annual_ends.append(rr.period_end_ym)
            if annual_ends:
                ym = sorted(annual_ends)[-1]
                fy_end_m_hint = int(ym.split("/")[1])
        except Exception:
            fy_end_m_hint = None

    if fy_end_m_hint is None:
        # infer from single rows (mode on quarter end months)
        try:
            months = [int(rr.period_end_ym.split("/")[1]) for rr in single_rows if rr.period_end_ym]
            if months:
                best_E, best_score = None, -1
                for E in range(1, 13):
                    allowed = {E, ((E - 3 - 1) % 12) + 1, ((E - 6 - 1) % 12) + 1, ((E - 9 - 1) % 12) + 1}
                    sc = sum(1 for m in months if m in allowed)
                    if sc > best_score:
                        best_E, best_score = E, sc
                if best_E is not None and best_score > 0:
                    fy_end_m_hint = best_E
        except Exception:
            fy_end_m_hint = None

    return (fy_end_m_hint, single_rows, cum_rows, annual_rows)
def kabutan_rows_to_records(
    ticker: str,
    fy_end_month_hint: Optional[int],
    single_rows: List[KabutanRow],
    cum_rows: List[KabutanRow],
    annual_rows: List[KabutanRow],
) -> List[FilingRecord]:
    """
    Convert Kabutan rows to FilingRecord(s).

    v5_25 FIX:
      - Keep cumulative rows as metric_cum (right block), single rows as metric_single (left block).
      - Annual rows contribute to Q4 cumulative when the period end month matches FY end month.
    """
    out: List[FilingRecord] = []
    if not (single_rows or cum_rows or annual_rows):
        return out

    fy_end_m = fy_end_month_hint

    # fall back inference (should rarely trigger because kabutan_parse already tries)
    if fy_end_m is None:
        months = []
        for rr in (single_rows + cum_rows + annual_rows):
            try:
                months.append(int(rr.period_end_ym.split("/")[1]))
            except Exception:
                pass
        if months:
            best_E, best_score = None, -1
            for E in range(1, 13):
                allowed = {E, ((E - 3 - 1) % 12) + 1, ((E - 6 - 1) % 12) + 1, ((E - 9 - 1) % 12) + 1}
                sc = sum(1 for m in months if m in allowed)
                if sc > best_score:
                    best_score, best_E = sc, E
            if best_E is not None and best_score > 0:
                fy_end_m = best_E

    def to_fy_q(end_ym: str) -> Tuple[Optional[int], Optional[int]]:
        """Map period end (YYYY/MM) to (fy_end_year, quarter_no)."""
        if not fy_end_m:
            return (None, None)
        try:
            y = int(end_ym.split("/")[0])
            m = int(end_ym.split("/")[1])
        except Exception:
            return (None, None)
        start_m = (fy_end_m % 12) + 1
        delta = (m - start_m) % 12
        if delta % 3 != 2:
            return (None, None)
        q = (delta // 3) + 1
        fy_end_y = y if m <= fy_end_m else (y + 1)
        return (fy_end_y, q)

    # --- Single-quarter rows (left)
    for rr in single_rows:
        fy_end_y, q = to_fy_q(rr.period_end_ym)
        if not (fy_end_y and fy_end_m and q):
            continue
        out.append(FilingRecord(
            ticker=ticker,
            source="kabutan",
            filing_date=None,
            title="Kabutan 3m actual",
            fy_end_year=fy_end_y,
            fy_end_month=fy_end_m,
            quarter_no=q,
            metric_cum={},
            metric_single=rr.metric,
        ))

    # --- Cumulative rows (right)
    for rr in cum_rows:
        fy_end_y, q = to_fy_q(rr.period_end_ym)
        if not (fy_end_y and fy_end_m and q):
            continue
        out.append(FilingRecord(
            ticker=ticker,
            source="kabutan",
            filing_date=None,
            title="Kabutan cumulative actual",
            fy_end_year=fy_end_y,
            fy_end_month=fy_end_m,
            quarter_no=q,
            metric_cum=rr.metric,
            metric_single={},
        ))

    # --- Annual rows => Q4 cumulative (when matches FY end month)
    for rr in annual_rows:
        if not fy_end_m:
            continue
        try:
            m = int(rr.period_end_ym.split("/")[1])
        except Exception:
            continue
        if m != fy_end_m:
            continue
        try:
            fy_end_y = int(rr.period_end_ym.split("/")[0])
        except Exception:
            continue
        out.append(FilingRecord(
            ticker=ticker,
            source="kabutan",
            filing_date=None,
            title="Kabutan annual actual",
            fy_end_year=fy_end_y,
            fy_end_month=fy_end_m,
            quarter_no=4,
            metric_cum=rr.metric,
            metric_single={},
        ))

    return out


def record_to_qidx(rec: FilingRecord) -> Optional[int]:
    if not rec.fy_end_year or not rec.quarter_no:
        return None
    fy_start = fy_start_year_from_end(rec.fy_end_year, rec.fy_end_month)
    qidx = qidx_from_fy_q(fy_start, rec.quarter_no)
    if qidx < 0 or qidx > MAX_Q_IDX:
        return None
    return qidx

def merge_records(records: List[FilingRecord]) -> Tuple[Dict[Tuple[str, int], ValuePoint], Dict[Tuple[str, int], ValuePoint]]:
    """
    Build best maps:
      cum_best[(metric, q_idx)] -> ValuePoint (million JPY)
      single_best[(metric, q_idx)] -> ValuePoint (million JPY)
    Derivations:
      - single from cumulative within same source & FY (Q2+ = cum - prev_cum)
      - cumulative from singles within same source & FY (sum Q1..Qn) [NOT from left sheet]
    """
    cum_best: Dict[Tuple[str, int], ValuePoint] = {}
    single_best: Dict[Tuple[str, int], ValuePoint] = {}

    # direct cumulative candidates by (metric, qidx, source) for safe cross-source single derivation
    cum_direct_by_key: Dict[Tuple[str, int, str], ValuePoint] = {}

    # collect direct points
    for rec in records:
        qidx = record_to_qidx(rec)
        if qidx is None:
            continue
        pr = SRC_PRIORITY.get(rec.source, 0)
        for m, v in rec.metric_cum.items():
            if v is None:
                continue
            try:
                fv = float(v)
            except Exception:
                continue
            if m in NONNEG_METRICS and fv < -1e-9:
                continue
            vp = ValuePoint(value=float(fv), source=rec.source, priority=pr, filing_date=rec.filing_date, title=rec.title, derived=False)
            key = (m, qidx)
            cum_best[key] = better_point(cum_best.get(key), vp)
            cum_direct_by_key[(m, qidx, rec.source)] = better_point(cum_direct_by_key.get((m, qidx, rec.source)), vp)
        for m, v in rec.metric_single.items():
            if v is None:
                continue
            try:
                fv = float(v)
            except Exception:
                continue
            if m in NONNEG_METRICS and fv < -1e-9:
                continue
            vp = ValuePoint(value=float(fv), source=rec.source, priority=pr, filing_date=rec.filing_date, title=rec.title, derived=False)
            key = (m, qidx)
            single_best[key] = better_point(single_best.get(key), vp)

    # derive single from cumulative per (source, metric, fy_end_year, fy_end_month)
    groups_cum: Dict[Tuple[str, str, int, int], Dict[int, ValuePoint]] = {}
    for rec in records:
        if not rec.fy_end_year or not rec.fy_end_month or not rec.quarter_no:
            continue
        for m, v in rec.metric_cum.items():
            if v is None:
                continue
            try:
                fv = float(v)
            except Exception:
                continue
            if m in NONNEG_METRICS and fv < -1e-9:
                continue
            key = (rec.source, m, rec.fy_end_year, rec.fy_end_month)
            pr = SRC_PRIORITY.get(rec.source, 0)
            vp = ValuePoint(value=float(fv), source=rec.source, priority=pr, filing_date=rec.filing_date, title=rec.title, derived=False)
            bucket = groups_cum.setdefault(key, {})
            bucket[rec.quarter_no] = better_point(bucket.get(rec.quarter_no), vp)

    for (src, m, fy_end_y, fy_end_m), g in groups_cum.items():
        pr = SRC_PRIORITY.get(src, 0)
        # Q1
        if 1 in g:
            qidx = qidx_from_fy_q(fy_start_year_from_end(fy_end_y, fy_end_m), 1)
            vp = ValuePoint(value=g[1].value, source=src, priority=pr, filing_date=g[1].filing_date, title=g[1].title, derived=True)
            single_best[(m, qidx)] = better_point_for_single(single_best.get((m, qidx)), vp)
        # Q2..Q4
        for q in (2, 3, 4):
            if q in g and (q - 1) in g:
                qidx = qidx_from_fy_q(fy_start_year_from_end(fy_end_y, fy_end_m), q)
                sv = g[q].value - g[q - 1].value
                if m in NONNEG_METRICS and float(sv) < -1e-9:
                    continue
                # keep decimals
                fd = g[q].filing_date or g[q - 1].filing_date
                vp = ValuePoint(value=sv, source=src, priority=pr, filing_date=fd, title=g[q].title, derived=True)
                single_best[(m, qidx)] = better_point_for_single(single_best.get((m, qidx)), vp)

        

    # --- Derive cumulative from singles within source (Kabutan only) ---
    # Because v5_30 does not use Kabutan cumulative tables, we build Kabutan cumulative series
    # from Kabutan 3m actual values as a LOW-CONFIDENCE derived fallback.
    # This is used ONLY to fill blank right-block cumulative cells; it must never override any direct cumulative.
    groups_single: Dict[Tuple[str, str, int, int], Dict[int, ValuePoint]] = {}
    for rec in records:
        if not rec.fy_end_year or not rec.fy_end_month or not rec.quarter_no:
            continue
        if rec.source not in DERIVE_CUM_FROM_SINGLE_SOURCES:
            continue
        pr = SRC_PRIORITY.get(rec.source, 0)
        for m, v in rec.metric_single.items():
            if v is None:
                continue
            try:
                fv = float(v)
            except Exception:
                continue
            if m in NONNEG_METRICS and fv < -1e-9:
                continue
            key = (rec.source, m, rec.fy_end_year, rec.fy_end_month)
            bucket = groups_single.setdefault(key, {})
            vp = ValuePoint(value=float(fv), source=rec.source, priority=pr, filing_date=rec.filing_date, title=rec.title, derived=False)
            bucket[rec.quarter_no] = better_point(bucket.get(rec.quarter_no), vp)

    for (src, m, fy_end_y, fy_end_m), g in groups_single.items():
        # build cumulative progressively only when all previous quarters exist
        fy_start = fy_start_year_from_end(fy_end_y, fy_end_m)
        running = 0.0
        for q in (1, 2, 3, 4):
            if q not in g:
                break
            running += float(g[q].value)
            qidx = qidx_from_fy_q(fy_start, q)
            if qidx < 0 or qidx > MAX_Q_IDX:
                continue
            if (m, qidx) in cum_best:
                continue  # direct cumulative already exists
            if m in NONNEG_METRICS and running < -1e-9:
                continue
            vp = ValuePoint(
                value=float(running),
                source="derived",
                priority=0,
                filing_date=g[q].filing_date,
                title=f"derived_cum_from_single(src={src})",
                derived=True,
            )
            cum_best[(m, qidx)] = better_point(cum_best.get((m, qidx)), vp)

    # --- Safe cross-source single derivation from DIRECT cumulative anchors ---
    # Purpose:
    #   When higher-priority sources provide cumulative but no reliable single, compute single as (cum - prev_cum).
    # v5_31 SAFETY:
    #   - Never override an existing DIRECT single-quarter value (derived=False).
    #   - Use only DIRECT (non-derived) cumulative anchors.
    #   - Current quarter cumulative is the anchor (higher priority wins).

    def _best_direct_cum(metric: str, qidx: int, prefer_src: str = None):
        if qidx < 0:
            return None
        if prefer_src:
            vp0 = cum_direct_by_key.get((metric, qidx, prefer_src))
            if vp0 and (not vp0.derived) and vp0.source != "derived":
                return vp0
        best = None
        for src in ("edinet", "tdnet", "kabutan"):
            vp = cum_direct_by_key.get((metric, qidx, src))
            if not vp:
                continue
            if vp.derived or vp.source == "derived":
                continue
            best = better_point(best, vp)
        return best

    for m in METRIC_JA.keys():
        for qidx in range(0, MAX_Q_IDX + 1):
            q_in_fy = (qidx % 4) + 1
            if q_in_fy == 1:
                continue

            existing_s = single_best.get((m, qidx))

            # v5_41 FIX:
            # Do NOT block a higher-priority derived single (from EDINET/TDNet cumulative) just because a direct
            # single exists from a lower-priority source (e.g., Kabutan).
            # Optimization: if the existing direct single is already >= the priority of the source cumulative,
            # derived cannot win -> skip.


            cur_c = _best_direct_cum(m, qidx)
            if not cur_c:
                continue

            if existing_s is not None and (not existing_s.derived):
                # Legacy v5_40 behavior: any direct single blocks deriving from cum-diff.
                # v5_41 default: allow override when the derived candidate comes from a higher-priority source.
                if not SINGLE_DERIVE_OVERRIDE_DIRECT_IF_HIGHER_PRIORITY:
                    continue
                try:
                    if int(existing_s.priority) >= int(cur_c.priority):
                        continue
                except Exception:
                    # if priority is missing/invalid, keep the safer behavior (derive allowed)
                    pass
            prev_c = _best_direct_cum(m, qidx - 1, prefer_src=cur_c.source) or _best_direct_cum(m, qidx - 1)
            if not prev_c:
                # Allow derived cumulative at qidx-1 if it was constructed from a direct prev cumulative + a direct single.
                cand_prev = cum_best.get((m, qidx - 1))
                if cand_prev and cand_prev.source == "derived" and ("derived_cum(" in (cand_prev.title or "")):
                    prev_c = cand_prev

            if not prev_c:
                continue
            sv = cur_c.value - prev_c.value
            if m in NONNEG_METRICS and float(sv) < -1e-9:
                continue
            vp = ValuePoint(
                value=float(sv),
                source=cur_c.source,
                priority=int(cur_c.priority or 0),
                filing_date=cur_c.filing_date or prev_c.filing_date,
                title=f"derived_single(cur_cum:{cur_c.source}-prev_cum:{prev_c.source})",
                derived=True,
            )
            single_best[(m, qidx)] = better_point_for_single(single_best.get((m, qidx)), vp)

    # derive cumulative values ONLY in the allowed case (cross-source):
    # If cumulative(q) is missing but cumulative(q-1) exists (MUST be a direct/source value)
    # and single(q) exists (source value), we may set:
    #   cum(q) = cum(q-1) + single(q)
    # This is marked as derived and given the LOWEST priority so any direct cumulative beats it.
    for m in METRIC_JA.keys():
        for qidx in range(0, MAX_Q_IDX + 1):
            q_in_fy = (qidx % 4) + 1
            if q_in_fy == 1:
                continue
            if (m, qidx) in cum_best:
                continue
            prev = cum_best.get((m, qidx - 1))
            cur_s = single_best.get((m, qidx))
            if not prev or not cur_s:
                continue
            if prev.derived:
                continue
            if cur_s.derived:
                continue
            cumv = prev.value + cur_s.value
            if m in NONNEG_METRICS and float(cumv) < -1e-9:
                continue
            vp = ValuePoint(
                value=cumv,
                source="derived",
                priority=0,
                filing_date=cur_s.filing_date or prev.filing_date,
                title=f"derived_cum(prev:{prev.source} + single:{cur_s.source})",
                derived=True,
            )
            cum_best[(m, qidx)] = better_point(cum_best.get((m, qidx)), vp)

    # NOTE: Cross-source single derivation from per-quarter BEST cumulative points is DISABLED.
    # It caused regressions where consecutive cumulative quarters came from different sources
    # (or from previously-derived cumulative), producing incorrect single values (e.g. negative sales).
    # It caused regressions where consecutive cumulative quarters came from different sources
    # (or from previously-derived cumulative), producing incorrect single values (e.g. negative sales).

    # We may have cumulative (YTD) from a higher-priority source (EDINET/TDNet) while only having
    # prior cumulative from a lower-priority source (often Kabutan-derived from 3m table).
    # In that case, we can still derive single-quarter = cur_cum - prev_cum.
    # This is especially important for Q4 single where EDINET provides FY cumulative but 3Q cumulative is absent.
    for (m, qidx), cur_c in list(cum_best.items()):
        if cur_c is None:
            continue
        # Only trust DIRECT cumulative as the anchor.
        if cur_c.derived:
            continue
        q_in_fy = qidx % 4 + 1
        if q_in_fy == 1:
            continue

        existing_s = single_best.get((m, qidx))
        # If we already have a DIRECT single-quarter value (from any source),
        # do NOT replace it with a cross-source cum-diff.
        #
        # Reason: Kabutan/TDNet/EDINET may provide a reliable direct "3m actual" for the quarter.
        # Cross-source cum-diff (e.g., EDINET FY cum minus a Kabutan-derived prior cum) is useful
        # only as a fallback when no direct single exists, and can create regressions (false WARNs)
        # if the prior cumulative is incomplete/misaligned.
        if existing_s is not None and (not existing_s.derived):
            continue
        prev_c = cum_best.get((m, qidx - 1))
        if prev_c is None:
            # Fallback: Q1 single equals Q1 cumulative (useful for deriving Q2 single from H1 cumulative).
            prev_s = single_best.get((m, qidx - 1))
            if prev_s is not None:
                prev_c = prev_s
        if prev_c is None:
            continue

        try:
            s_val = float(cur_c.value) - float(prev_c.value)
        except Exception:
            continue

        if m in NONNEG_METRICS and s_val < -1e-6:
            continue

        new_vp = ValuePoint(
            value=s_val,
            source=cur_c.source,
            priority=cur_c.priority,
            filing_date=cur_c.filing_date or prev_c.filing_date,
            title=f"cross_source_single({cur_c.source}_cum - {prev_c.source}_cum)",
            derived=True,
        )
        single_best[(m, qidx)] = better_point_for_single(existing_s, new_vp)

    return cum_best, single_best


def merge_records_with_instant(
    records: List[FilingRecord],
) -> Tuple[Dict[Tuple[str, int], ValuePoint], Dict[Tuple[str, int], ValuePoint], Dict[Tuple[str, int], ValuePoint]]:
    """
    Extended merge:
      - cum_best, single_best: same as merge_records()
      - instant_best: best instant values (e.g., total assets) per (metric, q_idx)
    """
    cum_best, single_best = merge_records(records)

    instant_best: Dict[Tuple[str, int], ValuePoint] = {}
    for rec in records:
        qidx = record_to_qidx(rec)
        if qidx is None:
            continue
        mm = getattr(rec, "metric_instant", None) or {}
        for m, v in mm.items():
            if v is None:
                continue
            try:
                fv = float(v)
            except Exception:
                continue
            vp = ValuePoint(
                value=fv,
                source=rec.source,
                priority=SOURCE_PRIORITY.get(rec.source, 0),
                filing_date=rec.filing_date,
                title=rec.title,
                derived=False,
            )
            instant_best[(m, qidx)] = better_point(instant_best.get((m, qidx)), vp)

    return cum_best, single_best, instant_best


# -----------------------
# Workbook filling
# -----------------------
def validate_workbook_layout(ws) -> None:
    # Basic sanity checks
    for col in ["A", "B"]:
        if ws[f"{col}1"].value is None:
            raise ValueError(f"Workbook header missing at {col}1")

    # Ensure left blocks exist
    for m, (c1, c2) in LEFT_BLOCKS.items():
        cols = iter_cols(c1, c2)
        if not cols:
            raise ValueError(f"Invalid left block range {m}: {c1}:{c2}")
        # check header row2 start/end
        if ws[f"{c1}2"].value is None or ws[f"{c2}2"].value is None:
            raise ValueError(f"Left block headers missing for {m} at row2 ({c1}2 or {c2}2)")

def fill_row(
    ws,
    r: int,
    cum_best: Dict[Tuple[str, int], ValuePoint],
    single_best: Dict[Tuple[str, int], ValuePoint],
    warn_tol: float,
    log_rows: List[List[Any]],
    counters: Dict[str, int],
    src_counts: Dict[str, int],
    fy_end_month: Optional[int] = None,
    kessanki_col_by_qidx: Optional[Dict[int, str]] = None,
    announced_cutoff_qidx: Optional[int] = None,
    today: Optional[dt.date] = None,
    meta_sh=None,
    meta_row_map: Optional[Dict[str, int]] = None,
) -> None:
    """
    Apply to one row.

    Policy:
    - Never overwrite existing numeric values.
    - Fill blanks (YELLOW).
    - If a source value exists and the current value differs by > warn_tol (±2 is OK),
      mark ORANGE (WARN) but keep the number.
    - If a cell was previously ORANGE by this script and now falls within tolerance,
      clear the managed fill (back to default).

    Additional safety:
    - For Q2+ single-quarter WARN checks, if the required cumulative series (current+prev)
      in the right cumulative block is missing/blank, we treat it as "cannot judge" and do NOT WARN.
      (This avoids false ORANGE when the quarter single is expected to be validated via cum-diff.)
    """

    def is_orange_fill(cell) -> bool:
        try:
            rgb = (cell.fill.start_color.rgb or "").upper()
        except Exception:
            rgb = ""
        return (cell.fill is not None and cell.fill.fill_type == "solid" and rgb.endswith("F4B183"))

    ticker = normalize_ticker(ws[f"A{r}"].value) or str(ws[f"A{r}"].value)

    if today is None:
        today = dt.date.today()

    # Determine whether a given quarter index is announced/available.
    # We never fill or WARN for quarters beyond the latest announced quarter (cutoff),
    # and we also hard-stop quarters whose period end is in the future.
    def is_unannounced_qidx(qidx: int) -> bool:
        if announced_cutoff_qidx is not None and qidx > int(announced_cutoff_qidx):
            return True
        if fy_end_month:
            try:
                start_m_local = (int(fy_end_month) % 12) + 1
            except Exception:
                start_m_local = None
            if start_m_local:
                fy_start_local = BASE_FY_START_YEAR + (qidx // 4)
                q_local = (qidx % 4) + 1
                fy_end_year_local = fy_start_local if start_m_local == 1 else (fy_start_local + 1)
                endd = fiscal_quarter_end_date(int(fy_end_year_local), int(q_local), int(fy_end_month))
                if endd and endd > today:
                    return True
        return False


    # ---- Fill/cleanup 決算期 (YYYY/MM) columns ----
    if fy_end_month and kessanki_col_by_qidx:
        try:
            start_m = (int(fy_end_month) % 12) + 1
        except Exception:
            start_m = None
        if start_m:
            # 1) Cleanup: if a 決算期 cell was previously filled by this script (managed fill)
            #    but now falls into an unannounced/future quarter, clear it back to blank.
            for qidx, col in kessanki_col_by_qidx.items():
                cell = ws[f"{col}{r}"]
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    continue
                if cell.value in (None, "", "-", "—", "―", "－", "ー"):
                    continue
                if is_managed_fill(cell) and is_unannounced_qidx(int(qidx)):
                    oldv = cell.value
                    cell.value = None
                    clear_managed_fill(cell)
                    log_rows.append([r, ticker, "clear_future", "kessanki", "meta", int(qidx), col, oldv, "", "managed", True, "", ""])

            # 2) Fill blanks only for announced quarters.
            for qidx, col in kessanki_col_by_qidx.items():
                if announced_cutoff_qidx is None:
                    continue
                if is_unannounced_qidx(int(qidx)):
                    continue
                cell = ws[f"{col}{r}"]
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    continue
                if cell.value not in (None, "", "-", "—", "―", "－", "ー"):
                    continue
                fy_start = BASE_FY_START_YEAR + (int(qidx) // 4)
                q = (int(qidx) % 4) + 1
                fy_end_year = fy_start if start_m == 1 else (fy_start + 1)
                yyyymm = fiscal_quarter_end_yyyymm(int(fy_end_year), int(q), int(fy_end_month))
                if yyyymm:
                    cell.value = yyyymm
                    cell.fill = YELLOW_FILL
                    counters["filled_cells"] += 1
                    log_rows.append([r, ticker, "fill", "kessanki", "meta", int(qidx), col, "", yyyymm, "derived", True, "", ""])

    def qidx_to_right_col(metric: str, qidx: int) -> Optional[str]:
        if metric not in RIGHT_BLOCKS:
            return None
        c1, c2 = RIGHT_BLOCKS[metric]
        cols = iter_cols(c1, c2)
        q0 = QSTART.get(metric, 0)
        i = qidx - q0
        if i < 0 or i >= len(cols):
            return None
        return cols[i]

    def qidx_to_left_col(metric: str, qidx: int) -> Optional[str]:
        if metric not in LEFT_BLOCKS:
            return None
        c1, c2 = LEFT_BLOCKS[metric]
        cols = iter_cols(c1, c2)
        q0 = QSTART.get(metric, 0)
        i = int(qidx) - int(q0)
        if i < 0 or i >= len(cols):
            return None
        return cols[i]

    def can_judge_single(metric: str, qidx: int) -> bool:
        # If we don't have a right cumulative block for this metric, allow judgement.
        if metric not in RIGHT_BLOCKS:
            return True
        q_in_fy = (qidx % 4) + 1
        if q_in_fy == 1:
            return True
        col_prev = qidx_to_right_col(metric, qidx - 1)
        col_cur = qidx_to_right_col(metric, qidx)
        if col_prev and col_cur:
            if (parse_num(ws[f"{col_prev}{r}"].value) is not None) and (parse_num(ws[f"{col_cur}{r}"].value) is not None):
                return True
        # If sheet is blank but we DO have both cumulative points from sources, we can judge.
        if cum_best.get((metric, qidx - 1)) is not None and cum_best.get((metric, qidx)) is not None:
            return True
        return False

    def apply_cell(col: str, qidx: int, metric: str, kind: str, vp: ValuePoint):
        """kind: 'single' or 'cum'"""
        cell = ws[f"{col}{r}"]
        # Skip (and optionally clean) unannounced/future quarters
        if is_unannounced_qidx(int(qidx)):
            # Never touch formulas
            if isinstance(cell.value, str) and cell.value.startswith('='):
                return
            # If this cell was previously filled by this script (managed fill), clear it back to blank
            if is_managed_fill(cell) and cell.value not in (None, "", "-", "—", "―", "－", "ー"):
                oldv = cell.value
                cell.value = None
                clear_managed_fill(cell)
                log_rows.append([r, ticker, "clear_future", metric, kind, int(qidx), col, oldv, "", "managed", True, "", ""])
            return

        cur = parse_num(cell.value)
        if cur is None:
            cell.value = float(vp.value)
            cell.fill = YELLOW_FILL
            counters["filled_cells"] += 1
            src_counts[vp.source] = src_counts.get(vp.source, 0) + 1
            log_rows.append([r, ticker, "fill", metric, kind, qidx, col, "", vp.value, vp.source, vp.derived, vp.filing_date, vp.title])
            if kind == "cum" and metric in RIGHT_BLOCKS and meta_sh is not None and meta_row_map is not None:
                k = meta_make_key(ticker, metric, kind, int(qidx), col)
                try:
                    meta_set(meta_sh, meta_row_map, k, ticker, metric, kind, int(qidx), col, vp)
                except MemoryError:
                    # r28: provenance logging must not stop filling
                    try:
                        meta_row_map["__disabled__"] = True
                    except Exception:
                        pass
                except Exception:
                    # keep going even if meta logging fails for any reason
                    pass
            return

        diff = float(cur) - float(vp.value)
        if abs(diff) > (warn_tol + 1e-6):
            # Low-confidence sources: avoid false positive WARNs.
            # Right-block provenance guard: if an existing RIGHT_BLOCKS cell was previously sourced
            # from a higher-priority source, we must NOT WARN based on a lower-priority source (e.g., TDNet expired).
            if kind == "cum" and metric in RIGHT_BLOCKS and meta_sh is not None and meta_row_map is not None:
                k = meta_make_key(ticker, metric, kind, int(qidx), col)
                prev = meta_get(meta_sh, meta_row_map, k)
                prev_pri = None
                prev_src = None
                if prev is not None:
                    try:
                        prev_pri = int(prev.get('priority') or 0)
                    except Exception:
                        prev_pri = 0
                    prev_src = str(prev.get('src') or '')
                else:
                    # If this cell is script-managed but we don't know the old source (older versions),
                    # we conservatively protect against Kabutan-based WARN/overwrite.
                    if vp.source == "kabutan" and is_managed_fill(cell):
                        if is_orange_fill(cell):
                            clear_managed_fill(cell)
                        log_rows.append([r, ticker, "skip_warn", metric, kind, qidx, col, cur, vp.value, vp.source, vp.derived, vp.filing_date, "unknown_managed_guard_against_kabutan"])
                        return
                if prev_pri is not None and prev_pri > int(getattr(vp, 'priority', 0) or 0):
                    if is_orange_fill(cell):
                        clear_managed_fill(cell)
                    log_rows.append([r, ticker, "skip_warn", metric, kind, qidx, col, cur, vp.value, vp.source, vp.derived, vp.filing_date, f"lower_source_than_existing(prev={prev_src},pri={prev_pri})"])
                    return

                # Controlled overwrite (RIGHT_BLOCKS only): overwrite ONLY if
                # - the cell is script-managed (managed fill OR meta exists), and
                # - the new source has strictly higher priority, and
                # - the new source is NOT Kabutan.
                if vp.source != "kabutan":
                    managed = is_managed_fill(cell) or (prev is not None)
                    prev_pri_eff = prev_pri if prev_pri is not None else (2 if is_managed_fill(cell) else 0)
                    if managed and int(getattr(vp, 'priority', 0) or 0) > int(prev_pri_eff or 0):
                        oldv = cur
                        cell.value = float(vp.value)
                        cell.fill = YELLOW_FILL
                        counters["overwritten_cells"] = counters.get("overwritten_cells", 0) + 1
                        src_counts[vp.source] = src_counts.get(vp.source, 0) + 1
                        meta_set(meta_sh, meta_row_map, k, ticker, metric, kind, int(qidx), col, vp)
                        log_rows.append([r, ticker, "overwrite", metric, kind, qidx, col, oldv, vp.value, vp.source, vp.derived, vp.filing_date, vp.title])
                        return


            # For single-quarter cells, if the cell matches the cumulative-difference identity,
            # treat it as OK even if a low-confidence direct-single source (e.g., Kabutan) disagrees.
            # This prevents false-positive ORANGE when:
            #   single_input == (cum[q] - cum[q-1]) but Kabutan 3m actual differs.
            if kind == "single" and metric in RIGHT_BLOCKS and can_judge_single(metric, qidx):
                try:
                    q_in_fy = (int(qidx) % 4) + 1
                except Exception:
                    q_in_fy = None

                exp = None  # expected single (million JPY)
                # Prefer sheet's right cumulative cells (already filled earlier in this loop).
                col_cur = qidx_to_right_col(metric, int(qidx)) if q_in_fy is not None else None
                col_prev = qidx_to_right_col(metric, int(qidx) - 1) if (q_in_fy is not None and q_in_fy != 1) else None

                v_cur = parse_num(ws[f"{col_cur}{r}"].value) if col_cur else None
                if q_in_fy == 1:
                    if v_cur is not None:
                        exp = float(v_cur)
                    else:
                        vp_c = cum_best.get((metric, int(qidx)))
                        if vp_c is not None:
                            exp = float(vp_c.value)
                else:
                    v_prev = parse_num(ws[f"{col_prev}{r}"].value) if col_prev else None
                    if v_prev is not None and v_cur is not None:
                        exp = float(v_cur) - float(v_prev)
                    else:
                        vp_prev = cum_best.get((metric, int(qidx) - 1))
                        vp_cur = cum_best.get((metric, int(qidx)))
                        if vp_prev is not None and vp_cur is not None:
                            exp = float(vp_cur.value) - float(vp_prev.value)

                if exp is not None:
                    try:
                        if abs(float(cur) - float(exp)) <= (warn_tol + 1e-6):
                            # OK by cumulative-difference identity -> do not WARN
                            if is_orange_fill(cell):
                                clear_managed_fill(cell)
                            log_rows.append([r, ticker, "skip_warn", metric, kind, qidx, col, cur, exp, "derived", True, "", "ok_by_cum_diff"])
                            return
                    except Exception:
                        pass

            # For single-quarter cells, skip WARN when required cum series is missing.
            
            if kind == "single" and not can_judge_single(metric, qidx):
                if is_orange_fill(cell):
                    clear_managed_fill(cell)
                log_rows.append([r, ticker, "skip_warn", metric, kind, qidx, col, cur, vp.value, vp.source, vp.derived, vp.filing_date, "unjudgeable_single(missing_cum)"])
                return
            cell.fill = ORANGE_FILL
            counters["warn_cells"] += 1
            src_counts[vp.source] = src_counts.get(vp.source, 0) + 1
            log_rows.append([r, ticker, "warn", metric, kind, qidx, col, cur, vp.value, vp.source, vp.derived, vp.filing_date, vp.title])
        else:
            # If previously managed ORANGE and now OK, clear it.
            if is_orange_fill(cell):
                clear_managed_fill(cell)


    def cleanup_cell(col: str, qidx: int, metric: str, kind: str):
        cell = ws[f"{col}{r}"]
        if isinstance(cell.value, str) and cell.value.startswith('='):
            return
        if cell.value in (None, "", "-", "—", "―", "－", "ー"):
            return
        if is_managed_fill(cell):
            oldv = cell.value
            cell.value = None
            clear_managed_fill(cell)
            log_rows.append([r, ticker, "clear_future", metric, kind, int(qidx), col, oldv, "", "managed", True, "", ""])

    # Right cumulative (fill first so single-quarter WARN decisions can rely on cum presence)
    for metric, (c1, c2) in RIGHT_BLOCKS.items():
        cols = iter_cols(c1, c2)
        q0 = QSTART.get(metric, 0)
        for i, col in enumerate(cols):
            qidx = q0 + i
            if qidx < 0 or qidx > MAX_Q_IDX:
                continue
            if is_unannounced_qidx(int(qidx)):
                cleanup_cell(col, int(qidx), metric, "cum")
                continue
            vp = cum_best.get((metric, qidx))
            if not vp:
                continue
            apply_cell(col, qidx, metric, "cum", vp)

    # Left singles
    for metric, (c1, c2) in LEFT_BLOCKS.items():
        cols = iter_cols(c1, c2)
        q0 = QSTART.get(metric, 0)
        for i, col in enumerate(cols):
            qidx = q0 + i
            if qidx < 0 or qidx > MAX_Q_IDX:
                continue
            if is_unannounced_qidx(int(qidx)):
                cleanup_cell(col, int(qidx), metric, "single")
                continue
            vp = single_best.get((metric, qidx))
            if not vp:
                continue
            apply_cell(col, qidx, metric, "single", vp)

    # Derived fill from (sheet cumulative + sheet single) for gross cumulative cells.
    # Example: if JQ (cum Q2) and EK (Q3 single) exist but JR is blank, fill JR=JQ+EK.
    # This is intentionally gross-only and blank-only to avoid broad behavioral changes.
    def fill_cum_from_sheet_single(metric: str) -> None:
        if metric not in LEFT_BLOCKS or metric not in RIGHT_BLOCKS:
            return
        c1, c2 = RIGHT_BLOCKS[metric]
        cum_cols = iter_cols(c1, c2)
        q0 = QSTART.get(metric, 0)
        for i, col in enumerate(cum_cols):
            qidx = q0 + i
            if qidx < 0 or qidx > MAX_Q_IDX:
                continue
            if is_unannounced_qidx(int(qidx)):
                continue

            cell = ws[f"{col}{r}"]
            if isinstance(cell.value, str) and cell.value.startswith('='):
                continue
            if parse_num(cell.value) is not None:
                continue
            if (cell.value is not None) and (str(cell.value).strip() not in ("", "-")):
                continue

            col_cur_single = qidx_to_left_col(metric, int(qidx))
            if not col_cur_single:
                continue
            v_cur_single = parse_num(ws[f"{col_cur_single}{r}"].value)
            if v_cur_single is None:
                continue

            q_in_fy = (int(qidx) % 4) + 1
            exp = None
            if q_in_fy == 1:
                exp = float(v_cur_single)
            else:
                col_prev_cum = qidx_to_right_col(metric, int(qidx) - 1)
                v_prev_cum = parse_num(ws[f"{col_prev_cum}{r}"].value) if col_prev_cum else None
                if v_prev_cum is None:
                    continue
                exp = float(v_prev_cum) + float(v_cur_single)

            if exp is None:
                continue

            cell.value = float(exp)
            cell.fill = YELLOW_FILL
            counters["filled_cells"] += 1
            src_counts["derived"] = src_counts.get("derived", 0) + 1
            log_rows.append([r, ticker, "fill", metric, "cum", int(qidx), col, "", exp, "derived", True, "", "derived_cum_from_sheet_single"])

    fill_cum_from_sheet_single("gross")

    # Derived fill from cumulative-on-sheet for gross single-quarter cells.
    # Example: if JQ (cum Q2) and EI (Q1 single) exist but EJ is blank, fill EJ=JQ-EI.
    def fill_single_from_sheet_cum(metric: str) -> None:
        if metric not in LEFT_BLOCKS or metric not in RIGHT_BLOCKS:
            return
        s1, s2 = LEFT_BLOCKS[metric]
        single_cols = iter_cols(s1, s2)
        q0 = QSTART.get(metric, 0)
        for i, col in enumerate(single_cols):
            qidx = q0 + i
            if qidx < 0 or qidx > MAX_Q_IDX:
                continue
            if is_unannounced_qidx(int(qidx)):
                continue

            cell = ws[f"{col}{r}"]
            if isinstance(cell.value, str) and cell.value.startswith('='):
                continue
            if parse_num(cell.value) is not None:
                continue
            if (cell.value is not None) and (str(cell.value).strip() not in ("", "-")):
                continue

            col_cur_cum = qidx_to_right_col(metric, int(qidx))
            if not col_cur_cum:
                continue
            v_cur_cum = parse_num(ws[f"{col_cur_cum}{r}"].value)
            if v_cur_cum is None:
                continue

            q_in_fy = (int(qidx) % 4) + 1
            exp = None
            if q_in_fy == 1:
                exp = float(v_cur_cum)
            else:
                col_prev_cum = qidx_to_right_col(metric, int(qidx) - 1)
                v_prev_cum = parse_num(ws[f"{col_prev_cum}{r}"].value) if col_prev_cum else None
                if v_prev_cum is not None:
                    exp = float(v_cur_cum) - float(v_prev_cum)
                else:
                    # If cumulative (q-1) is absent, derive via cum(q) - sum(singles before q in same FY).
                    fy_start_qidx = int(qidx) - (q_in_fy - 1)
                    vals: List[float] = []
                    ok = True
                    for qq in range(fy_start_qidx, int(qidx)):
                        j = qq - q0
                        if j < 0 or j >= len(single_cols):
                            ok = False
                            break
                        pv = parse_num(ws[f"{single_cols[j]}{r}"].value)
                        if pv is None:
                            ok = False
                            break
                        vals.append(float(pv))
                    if ok:
                        exp = float(v_cur_cum) - float(sum(vals))

            if exp is None:
                continue

            cell.value = float(exp)
            cell.fill = YELLOW_FILL
            counters["filled_cells"] += 1
            src_counts["derived"] = src_counts.get("derived", 0) + 1
            log_rows.append([r, ticker, "fill", metric, "single", int(qidx), col, "", exp, "derived", True, "", "derived_single_from_sheet_cum"])

    fill_single_from_sheet_cum("gross")

    # EY flag: if any ORANGE exists in recent5 quarters (scan the sheet, not only this run)
    latest_orange = False
    for blocks in (LEFT_BLOCKS, RIGHT_BLOCKS):
        for metric, (c1, c2) in blocks.items():
            cols = iter_cols(c1, c2)
            q0 = QSTART.get(metric, 0)
            for i, col in enumerate(cols):
                qidx = q0 + i
                if qidx in RECENT5_QIDX and is_orange_fill(ws[f"{col}{r}"]):
                    latest_orange = True
                    break
            if latest_orange:
                break
        if latest_orange:
            break

    ws[f"{WARN_COL}{r}"].value = 1 if latest_orange else None
    if latest_orange:
        counters["warn_rows"] += 1


# -----------------------
# Main flow

# -----------------------

def preflight_required_sources(session: requests.Session, edinet_api_key: str, sample_ticker: str, *, gh_enable: bool = True, gh_owner: str = TDNET_GH_DEFAULT_OWNER, gh_repo: str = TDNET_GH_DEFAULT_REPO, gh_branch: str = TDNET_GH_DEFAULT_BRANCH, gh_xbrl_dir: str = TDNET_GH_DEFAULT_XBRL_DIR) -> str:
    """Fail-fast verification that all REQUIRED data sources are reachable.

    This runs BEFORE any row processing. If any REQUIRED source is not reachable OR the content
    is obviously not the expected shape, the script exits with an error because the output would
    be meaningless under the user's strict requirement.

    v5.35 FIX:
    - The v5.34 edit accidentally broke indentation, executing the EDINET probe at module scope
      and raising NameError (edinet_api_key not defined).
    - Also align EDINET auth header to Ocp-Apim-Subscription-Key (same as the main EDINET client).
    - Make EDINET preflight robust to "today not indexed yet" 404 by probing a short backward window.
    """
    # --- EDINET ---
    # EDINET documents.json is indexed by 'date'. Depending on EDINET batch timing,
    # today's date can legitimately return 404 ("Not Found") even though the API is healthy.
    # Probe a short backward window and accept the first day that returns the normal shape.
    day0 = dt.datetime.now(JST).date()
    probe_days = [day0 - dt.timedelta(days=i) for i in range(0, 8)]  # today .. 7 days back

    used_url: Optional[str] = None
    last_resp: Optional[requests.Response] = None
    last_payload: Any = None
    ok = False
    ok_day: Optional[dt.date] = None
    working_key = ""
    key_candidates = edinet_api_key_candidates(edinet_api_key)

    for key in key_candidates:
        headers = {
            "Subscription-Key": key,
            "Ocp-Apim-Subscription-Key": key,
        }
        for url in EDINET_DOC_URL_CANDIDATES:
            used_url = url
            for day in probe_days:
                params = {"date": day.strftime("%Y-%m-%d"), "type": 2, "Subscription-Key": key}
                try:
                    resp = session.get(url, params=params, headers=headers, timeout=25)
                    last_resp = resp
                    # Try decode payload for diagnostics
                    payload = None
                    try:
                        payload = resp.json()
                    except Exception:
                        # keep a small prefix to avoid huge dumps
                        payload = (resp.text or "")[:2000]
                    last_payload = payload

                    if resp.status_code == 200 and isinstance(payload, dict):
                        meta = payload.get("metadata", {}) if isinstance(payload.get("metadata"), dict) else {}
                        meta_status = str(meta.get("status", ""))
                        # Success shape
                        if "results" in payload and meta_status == "200":
                            ok = True
                            ok_day = day
                            working_key = key
                            break

                        # If EDINET returned 200 but with an error-ish shape, keep probing other candidates/days.
                        continue

                    # Benign "date not yet indexed" case: keep probing backward
                    if resp.status_code == 404:
                        # Some 404 responses are JSON with metadata.status=404; some are plain.
                        continue

                    # Other status: keep probing other days/hosts; final failure will dump diagnostics.
                    continue

                except Exception as e:
                    last_payload = {"exception": type(e).__name__, "message": str(e), "day": day.strftime("%Y-%m-%d"), "url": url}
                    continue

            if ok:
                break
        if ok:
            break

    if not ok:
        dump_json_path = str(Path.cwd() / "_edinet_dump.json")
        try:
            with open(dump_json_path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "note": "EDINET preflight failed. This file contains the last response payload and the probe window.",
                        "used_url": used_url,
                        "candidate_keys": len(key_candidates),
                        "probe_days": [d.strftime("%Y-%m-%d") for d in probe_days],
                        "ok_day": ok_day.strftime("%Y-%m-%d") if ok_day else None,
                        "http_status": getattr(last_resp, "status_code", None),
                        "http_headers": dict(getattr(last_resp, "headers", {}) or {}),
                        "last_payload": last_payload,
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )
        except Exception:
            # best-effort raw dump
            try:
                if last_resp is not None:
                    with open(dump_json_path, "wb") as f:
                        f.write(last_resp.content or b"")
            except Exception:
                pass

        raise SourceAccessError("edinet", used_url or EDINET_DOC_URL_CANDIDATES[0], f"unexpected_json_shape_or_unreachable (dump={dump_json_path})")

    # --- TDNet (REQUIRED) ---
    # TDNet LIVE (release.tdnet.info) sometimes has a short public retention window.
    # If LIVE is not reachable (or blocked), we allow a fallback to the GitHub archive index probe.
    tdnet_live_ok = False
    try:
        day = dt.datetime.now(JST).date()
        daystr = day.strftime("%Y%m%d")
        tdnet_list = f"{TDNET_BASE}/I_list_001_{daystr}.html"
        _ = fetch_html_required(session, tdnet_list, source="tdnet", ok_markers=TDNET_OK_MARKERS, timeout=25, log_tag="tdnet_preflight")
        tdnet_main = f"{TDNET_BASE}/I_main_00.html"
        _ = fetch_html_required(session, tdnet_main, source="tdnet", ok_markers=TDNET_OK_MARKERS, timeout=25, log_tag="tdnet_preflight_main")
        tdnet_live_ok = True
    except Exception:
        tdnet_live_ok = False

    gh_ok = False
    if gh_enable:
        try:
            # GitHub probe: can we list the repo tree?
            url = _gh_tree_url(gh_owner, gh_repo, gh_branch)
            headers = {"Accept": "application/vnd.github+json"}
            r = safe_get(session, url, source="github", must=False, log_tag="github_preflight", headers=headers, timeout=TIMEOUT)
            if r is not None and r.status_code == 200:
                try:
                    j = r.json()
                    tree = j.get("tree") if isinstance(j, dict) else None
                    if isinstance(tree, list):
                        # best-effort: ensure there is at least one ZIP under the XBRL directory
                        pref = gh_xbrl_dir.strip('/').rstrip('/') + '/'
                        gh_ok = any(isinstance(x, dict) and (x.get('type')=='blob') and str(x.get('path','')).startswith(pref) and str(x.get('path','')).lower().endswith('.zip') for x in tree)
                        if not gh_ok:
                            # If tree exists but no ZIP matched (maybe structure differs), still accept connectivity.
                            gh_ok = True
                except Exception:
                    gh_ok = True
        except Exception:
            gh_ok = False

    if not (tdnet_live_ok or gh_ok):
        raise SourceAccessError("tdnet", TDNET_BASE, "both tdnet_live and github_archive unavailable")

    # Kabutan is not preflight-required; it is a fallback source.
    return working_key


# -----------------------------
# v5_41: 3-file support helpers
# -----------------------------

def _col_range(start_col: str, end_col: str) -> List[str]:
    a = column_index_from_string(start_col)
    b = column_index_from_string(end_col)
    return [get_column_letter(i) for i in range(a, b + 1)]


def _cell_to_yyyymm(v: Any) -> Optional[str]:
    if v is None:
        return None
    try:
        if isinstance(v, (dt.datetime, dt.date)):
            return f"{int(v.year):04d}/{int(v.month):02d}"
    except Exception:
        pass
    s = str(v).strip()
    if not s:
        return None
    # accept YYYY/MM or YYYY-M
    m = re.match(r"(\d{4})[/-](\d{1,2})", s)
    if m:
        yy = int(m.group(1))
        mm = int(m.group(2))
        if 1 <= mm <= 12:
            return f"{yy:04d}/{mm:02d}"
    return None




def _infer_yyyymm_from_period_header(ws: "openpyxl.worksheet.worksheet.Worksheet", col_letter: str, fy_end_month: int, header_row: int = 2) -> Optional[str]:
    """Infer a quarter/half/year period-end 'YYYY/MM' from the template header.

    The HALF/ANNUAL templates use labels like '2022年度前期' / '2022年度通期' in row2.
    Empirically (based on existing filled YYYY/MM cells), that '年度' refers to the *fiscal-year start year*.
    With fy_end_month, we can deterministically compute the period end month/year.

    Returns:
      'YYYY/MM' or None if not inferable.
    """
    try:
        hv = ws[f"{col_letter}{header_row}"].value
    except Exception:
        return None
    if hv is None:
        return None
    s = str(hv).strip()
    if not s:
        return None
    m = re.search(r"(\d{4})\s*年度\s*(前期|通期)", s)
    if not m:
        return None
    start_year = int(m.group(1))
    kind = m.group(2)

    try:
        fy_end_month = int(fy_end_month)
    except Exception:
        return None
    if fy_end_month < 1 or fy_end_month > 12:
        return None

    # fiscal year starts at the month after FY end month
    start_month = (fy_end_month % 12) + 1

    if kind == "前期":
        # end at start_month + 5 months
        add = 5
        end_month = ((start_month + add - 1) % 12) + 1
        end_year = start_year + ((start_month + add - 1) // 12)
        return f"{end_year:04d}/{end_month:02d}"

    # 通期
    end_year = start_year if fy_end_month == 12 else (start_year + 1)
    return f"{end_year:04d}/{fy_end_month:02d}"


def _period_col_to_qidx(ws: "openpyxl.worksheet.worksheet.Worksheet", r: int, col_letter: str, y2q: Dict[str, int], fy_end_month: int) -> Optional[int]:
    """Resolve a period column (C.. etc) to qidx using:
      1) cell value in the period column (preferred)
      2) fallback inference from the row2 header label if the cell is blank/invalid
    """
    ym = _cell_to_yyyymm(ws[f"{col_letter}{r}"].value)
    if not ym:
        ym = _infer_yyyymm_from_period_header(ws, col_letter, fy_end_month, header_row=2)
    return y2q.get(ym) if ym else None

def _build_yyyymm_to_qidx_map(fy_end_month: int) -> Dict[str, int]:
    """
    Build mapping from YYYY/MM (quarter-end) -> qidx for a given FY end month.
    Uses the same BASE_FY_START_YEAR..MAX_Q_IDX quarter timeline as PL.xlsx.
    """
    out: Dict[str, int] = {}
    for qidx in range(0, MAX_Q_IDX + 1):
        fy = BASE_FY_START_YEAR + (qidx // 4)
        q = (qidx % 4) + 1
        start_m = fy_start_month_from_fy_end_month(fy_end_month)
        fy_end_y = fy if start_m == 1 else (fy + 1)
        yyyymm = fiscal_quarter_end_yyyymm(fy_end_y, q, fy_end_month)
        out[str(yyyymm)] = qidx
    return out


def _infer_fy_end_month_from_period_cols(ws, r: int, period_cols: List[str]) -> Optional[int]:
    """
    Fallback inference if scraping is unavailable:
    - If headers contain '通期'/'第4四半期', use that column's YYYY/MM month.
    - Else, take the mode of months in the period cols (works when most are Q4).
    """
    try:
        for col in period_cols:
            h = ws[f"{col}2"].value
            if isinstance(h, str) and (("通期" in h) or ("第4" in h) or ("Q4" in h.upper())):
                ym = _cell_to_yyyymm(ws[f"{col}{r}"].value)
                if ym:
                    return int(ym.split("/")[1])
    except Exception:
        pass
    mm: List[int] = []
    for col in period_cols:
        ym = _cell_to_yyyymm(ws[f"{col}{r}"].value)
        if ym:
            try:
                mm.append(int(ym.split("/")[1]))
            except Exception:
                pass
    if not mm:
        return None
    from collections import Counter
    return Counter(mm).most_common(1)[0][0]


def _apply_value_cell(
    ws,
    meta_sh,
    meta_row_map: Dict[str, int],
    r: int,
    ticker: str,
    col: str,
    metric: str,
    kind: str,
    qidx: int,
    vp: Optional[ValuePoint],
    warn_tol: float,
    log_rows: List[List[Any]],
    counters: Dict[str, int],
    src_counts: Dict[str, int],
    workbook_tag: str,
    allow_overwrite_if_higher_priority: bool = True,
) -> None:
    """
    Apply vp.value to ws[col+r] following the same rules as PL:
      - Fill blanks with yellow
      - Warn (orange) only if abs(diff) > warn_tol
      - Clear orange if now within tolerance (only if managed)
      - Never overwrite user edits; only overwrite if the cell is script-managed AND new priority is higher.
    """
    cell = ws[f"{col}{r}"]

    # skip formula cells
    if isinstance(cell.value, str) and cell.value.strip().startswith("="):
        log_rows.append([r, ticker, "skip", metric, kind, qidx, col, str(cell.value), "", "", False, "", f"{workbook_tag}:formula"])
        return

    # no source -> skip (do not color)
    if vp is None:
        log_rows.append([r, ticker, "no_source_metric", metric, kind, qidx, col, cell.value, "", "", False, "", f"{workbook_tag}:"])
        return

    key = meta_make_key(ticker, metric, kind, qidx, col)
    meta = meta_get(meta_sh, meta_row_map, key)

    prev_pri = meta.get("priority") if meta else None
    prev_src = meta.get("src") if meta else None

    sheet_num = parse_num(cell.value)
    src_num = float(vp.value)

    # Protect against downgrade: if the cell is script-managed and previously higher-priority than current vp,
    # do not overwrite and also suppress WARNs.
    if prev_pri is not None:
        try:
            if int(prev_pri) > int(vp.priority):
                # If orange was set by us earlier, clear it (we now regard comparison unreliable).
                if is_managed_fill(cell) and managed_fill_rgb(cell) == ORANGE_FILL.start_color.rgb.upper():
                    clear_managed_fill(cell)
                log_rows.append([r, ticker, "skip", metric, kind, qidx, col, sheet_num, src_num, vp.source, bool(vp.derived), vp.filing_date, f"{workbook_tag}:prev_pri>{vp.priority}"])
                return
        except Exception:
            pass

    # If blank -> fill
    if sheet_num is None:
        cell.value = src_num
        # If the source has decimals (e.g., TDNet values like -828.338), ensure Excel displays them.
        try:
            if isinstance(src_num, float) and abs(src_num - round(src_num)) > 1e-9:
                cell.number_format = '0.###'
        except Exception:
            pass
        cell.fill = YELLOW_FILL
        mark_managed_fill(cell, YELLOW_FILL)
        meta_set(meta_sh, meta_row_map, key, ticker, metric, kind, qidx, col, vp)
        counters["filled_cells"] = counters.get("filled_cells", 0) + 1
        src_counts[vp.source] = src_counts.get(vp.source, 0) + 1
        log_rows.append([r, ticker, "fill", metric, kind, qidx, col, "", src_num, vp.source, bool(vp.derived), vp.filing_date, f"{workbook_tag}:{vp.title}"])
        return

    # Optionally overwrite script-managed cells only when new is higher priority
    if allow_overwrite_if_higher_priority and is_managed_fill(cell) and managed_fill_rgb(cell) == YELLOW_FILL.start_color.rgb.upper():
        try:
            if prev_pri is not None and int(vp.priority) > int(prev_pri):
                cell.value = src_num
                meta_set(meta_sh, meta_row_map, key, ticker, metric, kind, qidx, col, vp)
                counters["overwritten_cells"] = counters.get("overwritten_cells", 0) + 1
                src_counts[vp.source] = src_counts.get(vp.source, 0) + 1
                log_rows.append([r, ticker, "overwrite", metric, kind, qidx, col, sheet_num, src_num, vp.source, bool(vp.derived), vp.filing_date, f"{workbook_tag}:{vp.title}"])
                return
        except Exception:
            pass

    # Judge & warn
    diff = abs(float(sheet_num) - float(src_num))
    if diff <= float(warn_tol):
        if is_managed_fill(cell) and managed_fill_rgb(cell) == ORANGE_FILL.start_color.rgb.upper():
            clear_managed_fill(cell)
            log_rows.append([r, ticker, "clear_warn", metric, kind, qidx, col, sheet_num, src_num, vp.source, bool(vp.derived), vp.filing_date, f"{workbook_tag}:{vp.title}"])
        else:
            log_rows.append([r, ticker, "ok", metric, kind, qidx, col, sheet_num, src_num, vp.source, bool(vp.derived), vp.filing_date, f"{workbook_tag}:{vp.title}"])
        return

    # diff > tol -> warn (orange), but do not overwrite numeric value
    cell.fill = ORANGE_FILL
    mark_managed_fill(cell, ORANGE_FILL)
    meta_set(meta_sh, meta_row_map, key, ticker, metric, kind, qidx, col, vp)
    counters["warn_cells"] = counters.get("warn_cells", 0) + 1
    log_rows.append([r, ticker, "warn", metric, kind, qidx, col, sheet_num, src_num, vp.source, bool(vp.derived), vp.filing_date, f"{workbook_tag}:{vp.title}"])


def update_half_cum_workbook(
    input_path: str,
    output_path: str,
    records_cache: Dict[str, List[FilingRecord]],
    warn_tol: float,
    log_rows: List[List[Any]],
    ticker_whitelist: Optional[Set[str]] = None,
) -> Dict[str, int]:
    """
    Update データ取得_半期累積.xlsx (EDINET>TDNet only, reuse scraped records from PL run).
    - Fill blanks with yellow
    - Mark suspicious existing values with orange (abs diff > warn_tol)
    - Set AQ=1 if any orange in the latest 3 half-periods (1.5 years)
    """
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    meta_sh, meta_row_map = ensure_meta_sheet(wb)

    period_cols = _col_range("C", "J")      # 8
    keijo_cols  = _col_range("K", "R")      # 8
    opcf_cols   = _col_range("S", "Z")      # 8
    assets_cols = _col_range("AA", "AH")    # 8
    saishu_cols = _col_range("AI", "AP")    # 8
    flag_col = "AQ"

    counters: Dict[str, int] = {"processed_rows": 0, "filled_cells": 0, "warn_cells": 0, "overwritten_cells": 0}
    src_counts: Dict[str, int] = {}
    last3_idx = list(range(len(period_cols) - 3, len(period_cols)))  # rightmost 3

    for r in range(4, ws.max_row + 1):
        ticker = normalize_ticker(ws[f"A{r}"].value)
        if not ticker:
            continue
        if ticker_whitelist is not None and ticker not in ticker_whitelist:
            continue
        recs_all = records_cache.get(ticker, [])
        recs = [rec for rec in recs_all if rec.source in ("edinet", "tdnet")]
        if not recs:
            log_rows.append([r, ticker, "no_source_data", "", "", "", "", "", "", "", False, "", "[HALF]"])
            continue

        # FY end month: prefer scraped records, else sheet inference
        fy_end_m = None
        try:
            from collections import Counter
            mm = [rec.fy_end_month for rec in recs if rec.fy_end_month]
            if mm:
                fy_end_m = Counter(mm).most_common(1)[0][0]
        except Exception:
            fy_end_m = None
        if not fy_end_m:
            fy_end_m = _infer_fy_end_month_from_period_cols(ws, r, period_cols)
        if not fy_end_m:
            log_rows.append([r, ticker, "skip", "", "no_fy_end_month", "", "", "", "", "", False, "", "[HALF]"])
            continue

        y2q = _build_yyyymm_to_qidx_map(int(fy_end_m))
        cum_best, single_best, inst_best = merge_records_with_instant(recs)

        counters["processed_rows"] += 1

        qidx_list: List[Optional[int]] = []
        for pc in period_cols:
            qidx_list.append(_period_col_to_qidx(ws, r, pc, y2q, int(fy_end_m)))

        # apply values
        for i, qidx in enumerate(qidx_list):
            if qidx is None:
                continue
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, keijo_cols[i], "keijo", "cum", qidx, cum_best.get(("keijo", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[HALF]")
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, opcf_cols[i], "opcf", "cum", qidx, cum_best.get(("opcf", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[HALF]")
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, saishu_cols[i], "saishu", "cum", qidx, cum_best.get(("saishu", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[HALF]")
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, assets_cols[i], "assets", "instant", qidx, inst_best.get(("assets", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[HALF]")

        # flag AQ
        has_orange = False
        for i in last3_idx:
            for col in (keijo_cols[i], opcf_cols[i], assets_cols[i], saishu_cols[i]):
                c = ws[f"{col}{r}"]
                if managed_fill_rgb(c) == ORANGE_FILL.start_color.rgb.upper():
                    has_orange = True
                    break
            if has_orange:
                break

        flag_cell = ws[f"{flag_col}{r}"]
        if has_orange:
            if parse_num(flag_cell.value) is None:
                flag_cell.value = 1
                flag_cell.fill = YELLOW_FILL
                mark_managed_fill(flag_cell, YELLOW_FILL)
        else:
            # clear if managed and set by us
            if is_managed_fill(flag_cell) and managed_fill_rgb(flag_cell) == YELLOW_FILL.start_color.rgb.upper() and str(flag_cell.value).strip() == "1":
                flag_cell.value = None
                clear_managed_fill(flag_cell)

    wb.save(output_path)
    return counters


def update_annual_workbook(
    input_path: str,
    output_path: str,
    records_cache: Dict[str, List[FilingRecord]],
    warn_tol: float,
    log_rows: List[List[Any]],
    ticker_whitelist: Optional[Set[str]] = None,
) -> Dict[str, int]:
    """
    Update 年次_データ取得.xlsx (EDINET>TDNet only, reuse scraped records from PL run).
    - assets (instant), sga/ad/rnd (cum), capex_ppe (cum; usually negative)
    - Set AA=1 if any orange in the latest 3 periods
    """
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    meta_sh, meta_row_map = ensure_meta_sheet(wb)

    period_cols = _col_range("C", "F")      # 4
    assets_cols = _col_range("G", "J")      # 4
    sga_cols    = _col_range("K", "N")      # 4
    ad_cols     = _col_range("O", "R")      # 4
    rnd_cols    = _col_range("S", "V")      # 4
    capex_cols  = _col_range("W", "Z")      # 4
    flag_col = "AA"

    counters: Dict[str, int] = {"processed_rows": 0, "filled_cells": 0, "warn_cells": 0, "overwritten_cells": 0}
    src_counts: Dict[str, int] = {}

    last3_idx = list(range(len(period_cols) - 3, len(period_cols)))  # rightmost 3

    for r in range(4, ws.max_row + 1):
        ticker = normalize_ticker(ws[f"A{r}"].value)
        if not ticker:
            continue
        if ticker_whitelist is not None and ticker not in ticker_whitelist:
            continue
        recs_all = records_cache.get(ticker, [])
        recs = [rec for rec in recs_all if rec.source in ("edinet", "tdnet")]
        if not recs:
            log_rows.append([r, ticker, "no_source_data", "", "", "", "", "", "", "", False, "", "[ANNUAL]"])
            continue

        fy_end_m = None
        try:
            from collections import Counter
            mm = [rec.fy_end_month for rec in recs if rec.fy_end_month]
            if mm:
                fy_end_m = Counter(mm).most_common(1)[0][0]
        except Exception:
            fy_end_m = None
        if not fy_end_m:
            fy_end_m = _infer_fy_end_month_from_period_cols(ws, r, period_cols)
        if not fy_end_m:
            log_rows.append([r, ticker, "skip", "", "no_fy_end_month", "", "", "", "", "", False, "", "[ANNUAL]"])
            continue

        y2q = _build_yyyymm_to_qidx_map(int(fy_end_m))
        cum_best, single_best, inst_best = merge_records_with_instant(recs)
        counters["processed_rows"] += 1

        qidx_list: List[Optional[int]] = []
        for pc in period_cols:
            qidx_list.append(_period_col_to_qidx(ws, r, pc, y2q, int(fy_end_m)))

        for i, qidx in enumerate(qidx_list):
            if qidx is None:
                continue
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, assets_cols[i], "assets", "instant", qidx, inst_best.get(("assets", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[ANNUAL]")
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, sga_cols[i], "sga", "cum", qidx, cum_best.get(("sga", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[ANNUAL]")
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, ad_cols[i], "ad", "cum", qidx, cum_best.get(("ad", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[ANNUAL]")
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, rnd_cols[i], "rnd", "cum", qidx, cum_best.get(("rnd", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[ANNUAL]")
            _apply_value_cell(ws, meta_sh, meta_row_map, r, ticker, capex_cols[i], "capex_ppe", "cum", qidx, cum_best.get(("capex_ppe", qidx)), warn_tol, log_rows, counters, src_counts, workbook_tag="[ANNUAL]")

        has_orange = False
        for i in last3_idx:
            for col in (assets_cols[i], sga_cols[i], ad_cols[i], rnd_cols[i], capex_cols[i]):
                c = ws[f"{col}{r}"]
                if managed_fill_rgb(c) == ORANGE_FILL.start_color.rgb.upper():
                    has_orange = True
                    break
            if has_orange:
                break

        flag_cell = ws[f"{flag_col}{r}"]
        if has_orange:
            if parse_num(flag_cell.value) is None:
                flag_cell.value = 1
                flag_cell.fill = YELLOW_FILL
                mark_managed_fill(flag_cell, YELLOW_FILL)
        else:
            if is_managed_fill(flag_cell) and managed_fill_rgb(flag_cell) == YELLOW_FILL.start_color.rgb.upper() and str(flag_cell.value).strip() == "1":
                flag_cell.value = None
                clear_managed_fill(flag_cell)

    wb.save(output_path)
    return counters

def run(args) -> None:
    t0 = time.time()

    # Identity stamp (helps avoid confusion when output filenames contain v19/v20 etc)
    print(f"[SCRIPT] {os.path.basename(__file__)} (__version__={__version__})")

    global SINGLE_DERIVE_OVERRIDE_DIRECT_IF_HIGHER_PRIORITY
    SINGLE_DERIVE_OVERRIDE_DIRECT_IF_HIGHER_PRIORITY = (not bool(getattr(args, 'legacy_v5_40_single', False)))
    if getattr(args, 'legacy_v5_40_single', False):
        print('[MODE] legacy v5_40 single policy (direct single blocks cum-diff derived single).')
    else:
        print('[MODE] v5_41 single policy (allow higher-priority derived single to override lower-priority direct single).')

    # Fail-fast internal checks (no network)
    _preflight_runtime_checks()


    if not os.path.exists(args.input):
        raise FileNotFoundError(args.input)

    wb = openpyxl.load_workbook(args.input)
    ws = wb.active

    validate_workbook_layout(ws)
    ensure_right_blocks_headers(ws)

    # Provenance store for RIGHT_BLOCKS (prevents Kabutan from overriding/invalidating higher-priority data)
    meta_sh, meta_row_map = ensure_meta_sheet(wb)
    # R7: cache scraped FilingRecord list per ticker to reuse for half/annual workbooks (avoid double scraping)
    #     v6 forgot to initialize this, causing NameError + filled=0.
    records_cache: Dict[str, List[FilingRecord]] = {}
    # Map '決算期' columns: qidx -> Excel col letter (fill blank YYYY/MM values)
    kessanki_col_by_qidx: Dict[int, str] = {}
    for c in range(1, ws.max_column + 1):
        v0 = ws.cell(1, c).value
        if not (isinstance(v0, str) and v0.strip().startswith("決算期")):
            continue
        lab = ws.cell(2, c).value
        if not isinstance(lab, str):
            continue
        m = re.match(r"(\d{4})年度第([1-4])四半期", lab)
        if not m:
            continue
        fy = int(m.group(1))
        q = int(m.group(2))
        qidx = qidx_from_fy_q(fy, q)
        kessanki_col_by_qidx[qidx] = get_column_letter(c)

    start_row = max(4, args.start_row)
    rows_all = list(range(start_row, ws.max_row + 1))

    ticker_whitelist = read_ticker_whitelist(args)
    if ticker_whitelist is not None:
        rows_all = [r for r in rows_all if (normalize_ticker(ws[f"A{r}"].value) in ticker_whitelist)]

    rows = rows_all
    if args.limit and args.limit > 0:
        rows = rows[:args.limit]


    # --- FY end month hints from the workbook (used for Kabutan mapping to avoid time-series shift) ---
    # Determine from the rightmost valid period cell in C:R.
    # If the rightmost non-empty cell is not parseable as YYYY/MM with numeric month, skip it and continue left.
    # This handles fiscal-year change rows (e.g., old FY-end in N, new FY-end in O/P).
    def infer_fy_end_month_from_sheet_row(r: int) -> Optional[int]:
        period_cols = _col_range("C", "R")
        c0 = column_index_from_string("C")
        for col in reversed(period_cols):
            v = ws[f"{col}{r}"].value
            if v in (None, "", "-", "—", "―", "－", "ー"):
                continue

            month_end: Optional[int] = None
            if isinstance(v, (dt.date, dt.datetime)):
                mm = int(v.month)
                month_end = mm if 1 <= mm <= 12 else None
            else:
                s = str(v).strip()
                # Strictly require trailing YYYY/MM or YYYY-M; if trailing chars exist, treat as invalid.
                m = re.match(r"^(\d{4})[/-](\d{1,2})$", s)
                if m:
                    mm = int(m.group(2))
                    month_end = mm if 1 <= mm <= 12 else None

            if month_end is None:
                continue

            idx = column_index_from_string(col) - c0  # C=0 .. R=15
            if idx < 0 or idx > 15:
                continue
            q = (idx % 4) + 1
            # quarter-end month -> FY-end month
            fy_end_m = ((month_end + 3 * (4 - q) - 1) % 12) + 1
            return fy_end_m
        return None

    fy_end_month_hint_by_ticker: Dict[str, Optional[int]] = {}
    for r in rows:
        t = normalize_ticker(ws[f"A{r}"].value)
        if t:
            fy_end_month_hint_by_ticker[t] = infer_fy_end_month_from_sheet_row(r)


    # gather tickers for index prefetch
    tickers: List[str] = []
    for r in rows:
        t = normalize_ticker(ws[f"A{r}"].value)
        if t:
            tickers.append(t)
    ticker_set = set(tickers)

    session = ensure_requests_session()

    # Fail-fast: ensure all 3 sources (EDINET/TDNet/Kabutan) are reachable.
    sample_ticker = next((t for t in ticker_set if re.fullmatch(r"\d{4}", t)), next(iter(ticker_set), "1301"))
    args.edinet_api_key = preflight_required_sources(session, args.edinet_api_key, sample_ticker, gh_enable=(not args.tdnet_github_disable), gh_owner=args.tdnet_github_owner, gh_repo=args.tdnet_github_repo, gh_branch=args.tdnet_github_branch, gh_xbrl_dir=args.tdnet_github_xbrl_dir)


    # logs
    log_rows: List[List[Any]] = []
    log_rows.append(["row", "ticker", "action", "metric", "kind", "q_idx", "col", "sheet_value", "src_value", "src", "derived", "filing_date", "title"])

    counters = {
        "processed_rows": 0,
        "skipped_rows": 0,
        "filled_cells": 0,
        "overwritten_cells": 0,
        "warn_cells": 0,
        "warn_rows": 0,
        "no_source_rows": 0,
        "errors": 0,
    }
    src_counts: Dict[str, int] = {}

    # Prefetch indexes once (always scrape EDINET / TDNet / Kabutan)
    if not args.edinet_api_key:
        raise RuntimeError(
            "EDINET API key is required. Set environment variable EDINET_API_KEY "
            "or pass --edinet-api-key <KEY>."
        )
    edinet_index: Dict[str, List[EdinetIndexItem]] = edinet_build_index(session, ticker_set, EDINET_LOOKBACK_DAYS, args.edinet_api_key, log_rows)
    tdnet_index: Dict[str, List[TdnetIndexItem]] = tdnet_build_index(session, ticker_set, TDNET_LOOKBACK_DAYS, log_rows, gh_enable=(not args.tdnet_github_disable), gh_owner=args.tdnet_github_owner, gh_repo=args.tdnet_github_repo, gh_branch=args.tdnet_github_branch, gh_xbrl_dir=args.tdnet_github_xbrl_dir, gh_from_date=safe_date_from_any(args.tdnet_github_from_date) or TDNET_GH_DEFAULT_FROM_DATE)
    # cache downloaded zips
    tdnet_zip_cache: Dict[str, bytes] = {}
    tdnet_detail_cache: Dict[str, str] = {}
    edinet_zip_cache: Dict[str, bytes] = {}


    def collect_records_for_ticker(ticker: str) -> List[FilingRecord]:
        """Always scrape EDINET, TDNet, and Kabutan for the given ticker."""
        recs: List[FilingRecord] = []

        # EDINET (priority 1)
        for it in edinet_index.get(ticker, [])[:args.max_filings_per_ticker]:
            z = edinet_zip_cache.get(it.doc_id)
            if z is None:
                z = edinet_download_zip(session, it.doc_id, args.edinet_api_key)
                if z:
                    edinet_zip_cache[it.doc_id] = z
            if not z:
                continue
            # Reject ZIPs whose embedded securities code does not match the target ticker (fix: wrong-company ZIP pickup)
            sec_in_zip = extract_securities_code_from_zip(z)
            if sec_in_zip is not None and normalize_ticker(sec_in_zip) != ticker:
                log_rows.append([r, ticker, "skip", "", "zip_mismatch", "", "", "", "", "edinet", False, "", f"zip_sec={sec_in_zip}"])
                continue
            prefer_con = True

            metric_cum, metric_single = parse_zip_metrics(z, prefer_consolidated=prefer_con, log=log_rows, who="edinet", expected_ticker=ticker)
            metric_instant = parse_zip_instant_metrics(z, prefer_consolidated=prefer_con, log=log_rows, who="edinet_instant", expected_ticker=ticker)
            if not metric_cum and not metric_single:
                continue

            fy_end_y, fy_end_m, q = parse_title_for_fy_end_q(it.title)
            title_no_space = (it.title or "").replace(" ", "").replace("　", "")
            if q is None:
                q = infer_quarter_from_period_span(it.period_start, it.period_end)
            if q is None:
                # Some EDINET half/annual titles do not carry explicit quarter labels and
                # periodStart/periodEnd can represent a full FY span (q cannot be inferred by span).
                # In that case, use form-name semantics as a safe fallback.
                if "半期報告書" in title_no_space:
                    q = 2
                elif "有価証券報告書" in title_no_space:
                    q = 4
            if fy_end_m is None:
                # For half/annual EDINET docs, periodEnd month is a stronger FY-end hint than periodStart.
                if it.period_end and (("半期報告書" in title_no_space) or ("有価証券報告書" in title_no_space)):
                    fy_end_m = int(it.period_end.month)
                else:
                    fy_end_m = infer_fy_end_month_from_period_start(it.period_start)
            if fy_end_m is None and it.period_end:
                fy_end_m = int(it.period_end.month)
            if fy_end_y is None:
                fy_end_y = infer_fy_end_year_from_period_end(it.period_end, fy_end_m)

            if not (fy_end_y and fy_end_m and q):
                continue

            sanitize_single_vs_cum(metric_cum, metric_single, q, tol=args.warn_tolerance)

            recs.append(FilingRecord(
                ticker=ticker,
                source="edinet",
                filing_date=it.submit_date,
                title=it.title,
                fy_end_year=fy_end_y,
                fy_end_month=fy_end_m,
                quarter_no=q,
                metric_cum=metric_cum,
                metric_single=metric_single,
                metric_instant=metric_instant,
            ))

        # TDNet LIVE (priority 2)
        for it in tdnet_index.get(ticker, [])[:args.max_filings_per_ticker]:
            if not tdnet_is_usable_results_title(it.title):
                continue
            url = it.zip_url
            if not (url.lower().endswith('.zip') or '.zip?' in url.lower()):
                url = tdnet_detail_cache.get(it.zip_url) or tdnet_resolve_zip_from_detail(session, it.zip_url)
                if url:
                    tdnet_detail_cache[it.zip_url] = url
            if not url:
                continue
            z = tdnet_zip_cache.get(url)
            resp = None
            if z is None:
                src_tag = "tdnet_github" if getattr(it, "origin", "live") == "github" else "tdnet"
                resp = safe_get(session, url, source=src_tag, must=False, allow_404=True, log_tag="tdnet_zip")  # TDNet zip may disappear (404); treat as optional
                if resp:
                    z = require_zip_bytes("tdnet", url, resp.content or b"")
                    tdnet_zip_cache[url] = z
            if not z:
                continue
            sec_in_zip = extract_securities_code_from_zip(z)
            if sec_in_zip is not None and normalize_ticker(sec_in_zip) != ticker:
                log_rows.append(["", ticker, "skip", "", "zip_mismatch", "", "", "", "", "tdnet", False, "", f"zip_sec={sec_in_zip}"])
                continue
            prefer_con = True
            metric_cum, metric_single = parse_zip_metrics(z, prefer_consolidated=prefer_con, log=log_rows, who="tdnet", expected_ticker=ticker)
            metric_instant = parse_zip_instant_metrics(z, prefer_consolidated=prefer_con, log=log_rows, who="tdnet_instant", expected_ticker=ticker)
            if not metric_cum and not metric_single:
                continue

            meta = extract_fy_end_and_quarter_from_ixbrl_zip(z)
            if meta:
                fy_end_y, fy_end_m, q = meta
            else:
                fy_end_y, fy_end_m, q = parse_title_for_fy_end_q(it.title)

            if not (fy_end_y and fy_end_m and q):
                # If we cannot determine FY/Q reliably, skip to avoid time-series shift.
                continue

            sanitize_single_vs_cum(metric_cum, metric_single, q, tol=args.warn_tolerance)

            recs.append(FilingRecord(
                ticker=ticker,
                source="tdnet",
                filing_date=it.disclosure_date,
                title=it.title,
                fy_end_year=fy_end_y,
                fy_end_month=fy_end_m,
                quarter_no=q,
                metric_cum=metric_cum,
                metric_single=metric_single,
                metric_instant=metric_instant,
            ))

        # Kabutan (priority 3)
        html = kabutan_fetch_html(session, ticker)
        if html:
            fy_end_m_hint, k_single, k_cum, k_annual = kabutan_parse(html)
            hint = fy_end_m_hint or fy_end_month_hint_by_ticker.get(ticker)
            recs.extend(kabutan_rows_to_records(ticker, hint, k_single, k_cum, k_annual))

        return recs

    # Process rows
    total = len(rows)
    for idx, r in enumerate(rows, 1):
        ticker = normalize_ticker(ws[f"A{r}"].value)
        if not ticker:
            counters["skipped_rows"] += 1
            log_rows.append([r, str(ws[f"A{r}"].value), "skip", "", "", "", "", "", "", "", "", "", "invalid_ticker"])
            continue

        try:
            counters["processed_rows"] += 1
            recs = collect_records_for_ticker(ticker)
            records_cache[ticker] = recs
            if not recs:
                counters["no_source_rows"] += 1
                log_rows.append([r, ticker, "no_source_data", "", "", "", "", "", "", "", "", "", ""])
                continue


            # Determine FY end month for this row:
            # 1) Prefer workbook hint from existing 決算期 cells (if any)
            # 2) Fallback to mode of scraped records (EDINET/TDNet/Kabutan)
            fy_end_m_row = infer_fy_end_month_from_sheet_row(r)
            if not fy_end_m_row:
                try:
                    mm = [rec.fy_end_month for rec in recs if rec.fy_end_month]
                    if mm:
                        from collections import Counter
                        fy_end_m_row = Counter(mm).most_common(1)[0][0]
                except Exception:
                    fy_end_m_row = None
            cum_best, single_best = merge_records(recs)
            cutoff_qidx = compute_announced_cutoff_qidx(recs, today=dt.date.today())
            fill_row(ws, r, cum_best, single_best, warn_tol=args.warn_tolerance, log_rows=log_rows, counters=counters, src_counts=src_counts, fy_end_month=fy_end_m_row, kessanki_col_by_qidx=kessanki_col_by_qidx, announced_cutoff_qidx=cutoff_qidx, today=dt.date.today(), meta_sh=meta_sh, meta_row_map=meta_row_map)

        except SourceAccessError as e:
            # Under strict requirements, abort immediately if any required source becomes inaccessible.
            eprint(str(e))
            raise
        except Exception as e:
            counters["errors"] += 1
            tb = traceback.format_exc()
            dump_path = ""
            try:
                base = os.path.splitext(args.output)[0]
                mode = getattr(args, "error_dump", "single") or "single"
                if mode == "per_row":
                    dump_path = f"{base}_error_{ticker}_{r}.txt"
                    with open(dump_path, "w", encoding="utf-8") as f:
                        f.write(tb)
                elif mode == "single":
                    dump_path = getattr(args, "error_dump_file", "") or f"{base}_errors.txt"
                    with open(dump_path, "a", encoding="utf-8") as f:
                        f.write(f"\n===== ERROR ticker={ticker} row={r} =====\n")
                        f.write(tb)
                else:
                    dump_path = ""
            except Exception:
                dump_path = ""
            msg = repr(e) + (f" (trace={dump_path})" if dump_path else "")
            log_rows.append([r, ticker, "error", "", "", "", "", "", "", "", "", "", msg])

        if args.progress_every and (idx % args.progress_every == 0):
            print(f"[{idx}/{total}] ticker={ticker} processed_rows={counters['processed_rows']} filled={counters['filled_cells']} overwritten={counters.get('overwritten_cells',0)} warn={counters['warn_cells']}")

    # Save workbook
    wb.save(args.output)

    # ---- v5_41: optional additional workbooks (reuse scraped EDINET/TDNet data; no double scraping) ----
    if getattr(args, "half_input", ""):
        half_in = args.half_input
        half_out = args.half_output or (str(Path(half_in).with_suffix("")) + "_out.xlsx")
        c_half = update_half_cum_workbook(
            input_path=half_in,
            output_path=half_out,
            records_cache=records_cache,
            warn_tol=float(args.warn_tolerance),
            log_rows=log_rows,
            ticker_whitelist=ticker_whitelist,
        )
        print(f"[HALF] saved: {half_out}  processed_rows={c_half.get('processed_rows',0)} filled={c_half.get('filled_cells',0)} warn={c_half.get('warn_cells',0)} overwrite={c_half.get('overwritten_cells',0)}")

    if getattr(args, "annual_input", ""):
        ann_in = args.annual_input
        ann_out = args.annual_output or (str(Path(ann_in).with_suffix("")) + "_out.xlsx")
        c_ann = update_annual_workbook(
            input_path=ann_in,
            output_path=ann_out,
            records_cache=records_cache,
            warn_tol=float(args.warn_tolerance),
            log_rows=log_rows,
            ticker_whitelist=ticker_whitelist,
        )
        print(f"[ANNUAL] saved: {ann_out}  processed_rows={c_ann.get('processed_rows',0)} filled={c_ann.get('filled_cells',0)} warn={c_ann.get('warn_cells',0)} overwrite={c_ann.get('overwritten_cells',0)}")

    print(f"Saved: {args.output}")

    # Write log
    if args.log_csv:
        # append summary lines
        log_rows.append(["", "", "SUMMARY", "", "", "", "", "", "", "", "", "", ""])
        for k, v in counters.items():
            log_rows.append(["", "", "summary_counter", k, "", "", "", "", v, "", "", "", ""])
        for src, v in sorted(src_counts.items(), key=lambda x: (-x[1], x[0])):
            log_rows.append(["", "", "summary_source", src, "", "", "", "", v, "", "", "", ""])
        # HTTP failures (sample)
        log_rows.append(["", "", "summary_http_fails", "count", "", "", "", "", len(HTTP_FAILS), "", "", "", ""])
        for (u, msg) in HTTP_FAILS[:20]:
            log_rows.append(["", "", "http_fail", u, msg, "", "", "", "", "", "", "", ""])
        with open(args.log_csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerows(log_rows)
        print(f"Log: {args.log_csv}")

    elapsed = time.time() - t0
    print(f"Done. elapsed={elapsed:.1f}s  processed_rows={counters['processed_rows']}  filled_cells={counters['filled_cells']}  overwritten_cells={counters.get('overwritten_cells',0)}  warn_cells={counters['warn_cells']}  warn_rows={counters['warn_rows']}  errors={counters['errors']}")

    # Fail the process if any row-level errors happened (prevents silent 'success').
    if counters.get('errors', 0) > 0:
        eprint(f"FATAL: {counters['errors']} errors occurred during processing. Output may be incomplete.")
        if args.log_csv:
            eprint(f"See log CSV: {args.log_csv}")
        raise SystemExit(2)


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="PL scraper/filler (single-file)")
    p.add_argument("--input", required=True, help="input xlsx")
    p.add_argument("--output", required=True, help="output xlsx")
    p.add_argument("--start-row", type=int, default=4)
    p.add_argument("--limit", type=int, default=0, help="0 means all rows")

    p.add_argument("--tickers", default="", help="Comma/space separated ticker whitelist (e.g., 1301,142A). If set, only these tickers are processed.")
    p.add_argument("--tickers-file", default="", help="Path to a text/CSV file containing tickers (one per line). If set, only these tickers are processed.")
    p.add_argument("--progress-every", type=int, default=10)
    p.add_argument("--log-csv", default="", help="log csv path")
    p.add_argument("--warn-tolerance", type=float, default=2.0, help="abs diff <= this is allowed (default ±2)")
    p.add_argument("--legacy-v5-40-single", action="store_true", help="keep v5_40 single policy: any direct single blocks cum-diff derived single (may cause false ORANGE).")
    p.add_argument("--error-dump", choices=["single","per_row","none"], default="single",
                       help="Error traceback dump mode. single: append all tracebacks to one file (<output_base>_errors.txt by default). per_row: create one <output_base>_error_<ticker>_<row>.txt per error. none: do not dump tracebacks to files.")
    p.add_argument("--error-dump-file", default="",
                       help="When --error-dump=single, write/append tracebacks to this file path (default: <output_base>_errors.txt).")
    
    # Optional: additional workbooks (v5_41)
    p.add_argument("--half-input", default="", help="path to データ取得_半期累積.xlsx (optional)")
    p.add_argument("--half-output", default="", help="output path for half-cumulative workbook (optional; defaults to input basename + _out.xlsx)")
    p.add_argument("--annual-input", default="", help="path to 年次_データ取得.xlsx (optional)")
    p.add_argument("--annual-output", default="", help="output path for annual workbook (optional; defaults to input basename + _out.xlsx)")


    # Required to scrape EDINET (always scraped; no disable switches)
    p.add_argument("--edinet-api-key", default=os.environ.get("EDINET_API_KEY", ""), help="EDINET API key (or set env EDINET_API_KEY)")

    # Limit downloads per ticker per source to control runtime
    p.add_argument("--max-filings-per-ticker", type=int, default=6, help="limit downloads per ticker per source")

    # TDNet GitHub archive fallback (optional, enabled by default).
    p.add_argument("--tdnet-github-disable", action="store_true", help="Disable TDNet GitHub archive fallback.")
    p.add_argument("--tdnet-github-owner", default=TDNET_GH_DEFAULT_OWNER, help="GitHub owner for TDNet archive repo (default: yukizi1113)")
    p.add_argument("--tdnet-github-repo", default=TDNET_GH_DEFAULT_REPO, help="GitHub repo for TDNet archive (default: tdnet)")
    p.add_argument("--tdnet-github-branch", default=TDNET_GH_DEFAULT_BRANCH, help="GitHub branch/ref for TDNet archive (default: main)")
    p.add_argument("--tdnet-github-xbrl-dir", default=TDNET_GH_DEFAULT_XBRL_DIR, help="Path in repo that contains TDNet XBRL ZIPs (default: XBRL)")
    p.add_argument("--tdnet-github-from-date", default=str(TDNET_GH_DEFAULT_FROM_DATE), help="Only use GitHub archive files on/after this date (YYYY-MM-DD). Default: 2025-12-17")

    return p

if __name__ == "__main__":
    args = build_argparser().parse_args()
    run(args)
