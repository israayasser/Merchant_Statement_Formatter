from __future__ import annotations

from pathlib import Path
from typing import Iterable
import re
import zipfile
import tempfile
import shutil
from datetime import datetime, time

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule


CANONICAL_HEADERS = [
    "Merchant Location MID", "Terminal ID", "Settlement Date", "Trxn Date",
    "Trxn Time", "Batch No", "Invoice No", "Auth Code", "Card No",
    "Card Group", "Trxn Type", "Gross Amount", "Discount Amount",
    "VAT Amount", "Net Amount"
]

HEADER_ALIASES = {
    "merchant location mid": "Merchant Location MID",
    "terminal id": "Terminal ID",
    "settlem date": "Settlement Date",
    "settlement date": "Settlement Date",
    "trxn date": "Trxn Date",
    "trxn time": "Trxn Time",
    "batch no": "Batch No",
    "invoice no": "Invoice No",
    "auth code": "Auth Code",
    "card no": "Card No",
    "card group": "Card Group",
    "trxn type": "Trxn Type",
    "gross amount": "Gross Amount",
    "discount amount": "Discount Amount",
    "vat amount": "VAT Amount",
    "net amount": "Net Amount",
}


def _norm_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).replace("\n", " ").replace("\r", " ").lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def _header_score(row: Iterable) -> int:
    cells = {_norm_text(v) for v in row if _norm_text(v)}
    wanted = set(HEADER_ALIASES)
    return len(cells & wanted)


def _is_transaction_header(row: Iterable) -> bool:
    # We deliberately detect by semantic content, not row/column numbers.
    return _header_score(row) >= 8


def _find_transaction_blocks(df: pd.DataFrame):
    header_rows = [i for i in range(len(df)) if _is_transaction_header(df.iloc[i].tolist())]
    blocks = []
    for pos, header_idx in enumerate(header_rows):
        end = header_rows[pos + 1] if pos + 1 < len(header_rows) else len(df)
        blocks.append((header_idx, end))
    return blocks


def _map_header_columns(header_row: pd.Series) -> dict[str, int]:
    # Keep the semantic anchor positions from the header.  The source report
    # uses merged cells, so the actual data value can be one or two columns
    # beside the visible header anchor.
    mapping = {}
    for idx, value in enumerate(header_row.tolist()):
        key = _norm_text(value)
        if key in HEADER_ALIASES:
            mapping[HEADER_ALIASES[key]] = idx
    return mapping


def _header_bands(mapping: dict[str, int], width: int) -> dict[str, tuple[int, int]]:
    anchors = sorted((idx, name) for name, idx in mapping.items())
    bands = {}
    for pos, (anchor, name) in enumerate(anchors):
        prev_anchor = anchors[pos - 1][0] if pos else 0
        next_anchor = anchors[pos + 1][0] if pos + 1 < len(anchors) else width
        left = max(0, (prev_anchor + anchor) // 2)
        right = min(width, (anchor + next_anchor + 1) // 2)
        bands[name] = (left, right)
    return bands


def _extract_band_value(row: pd.Series, band: tuple[int, int], anchor: int):
    left, right = band
    candidates = []
    for idx in range(left, right):
        value = row.iloc[idx]
        if value is not None and not (isinstance(value, float) and pd.isna(value)) and str(value).strip() != "":
            candidates.append((abs(idx - anchor), value))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]

def _clean_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().date()
    if isinstance(value, datetime):
        return value.date()
    try:
        return pd.to_datetime(value, dayfirst=True, errors="coerce").date()
    except Exception:
        return None


def _clean_time(value):
    if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
        return None
    if isinstance(value, pd.Timestamp):
        return value.time()
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    s = str(value).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s)
    if m:
        return time(int(m.group(1)), int(m.group(2)), int(m.group(3) or 0))
    try:
        parsed = pd.to_datetime(s, errors="coerce")
        return None if pd.isna(parsed) else parsed.time()
    except Exception:
        return None


def _clean_number(value):
    if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    s = str(value).strip().replace(",", "")
    if s in {"-", "—", "–"}:
        return None
    try:
        return float(s)
    except ValueError:
        return value


def _row_has_transaction_data(values: list) -> bool:
    # Semantic checks: at least MID/terminal + transaction type + amount-like field.
    text = [_norm_text(v) for v in values]
    nonempty = sum(bool(v) for v in text)
    if nonempty < 3:
        return False
    has_type = any(v in {"sales draft", "credit voucher", "refund", "cash advance", "void", "purchase"} or "sales" in v for v in text)
    numericish = any(re.match(r"^-?\d+(?:\.\d+)?$", v.replace(",", "")) for v in text if v)
    return has_type and numericish


def clean_statement(source_path: str | Path) -> tuple[pd.DataFrame, dict]:
    source_path = Path(source_path)
    df = pd.read_excel(source_path, header=None, dtype=object)
    blocks = _find_transaction_blocks(df)
    records = []
    detected_headers = 0

    for header_idx, end_idx in blocks:
        mapping = _map_header_columns(df.iloc[header_idx])
        if len(mapping) < 8:
            continue
        bands = _header_bands(mapping, len(df.columns))
        detected_headers += 1
        for r in range(header_idx + 1, end_idx):
            row = df.iloc[r]
            values = [
                _extract_band_value(row, bands[h], mapping[h]) if h in mapping else None
                for h in CANONICAL_HEADERS
            ]
            if not _row_has_transaction_data(values):
                continue
            rec = dict(zip(CANONICAL_HEADERS, values))
            rec["Settlement Date"] = _clean_date(rec["Settlement Date"])
            rec["Trxn Date"] = _clean_date(rec["Trxn Date"])
            rec["Trxn Time"] = _clean_time(rec["Trxn Time"])
            for col in ["Gross Amount", "Discount Amount", "VAT Amount", "Net Amount"]:
                rec[col] = _clean_number(rec[col])
            records.append(rec)

    result = pd.DataFrame(records, columns=CANONICAL_HEADERS)
    if not result.empty:
        result["Trxn DateTime"] = [
            datetime.combine(d, t) if pd.notna(d) and pd.notna(t) else None
            for d, t in zip(result["Trxn Date"], result["Trxn Time"])
        ]
        # Put DateTime after Time while retaining source columns.
        cols = CANONICAL_HEADERS[:4] + ["Trxn Time", "Trxn DateTime"] + CANONICAL_HEADERS[5:]
        result = result[cols]

    stats = {
        "source": source_path.name,
        "source_rows": len(df),
        "source_columns": len(df.columns),
        "detected_header_blocks": detected_headers,
        "transactions": len(result),
    }
    return result, stats


def _style_sheet(ws, freeze="A2", table_name="TransactionsTable"):
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    header_fill = PatternFill("solid", fgColor="1565C0")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    ws.row_dimensions[1].height = 30

    date_cols = {"Settlement Date", "Trxn Date"}
    time_cols = {"Trxn Time"}
    dt_cols = {"Trxn DateTime"}
    amount_cols = {"Gross Amount", "Discount Amount", "VAT Amount", "Net Amount"}
    headers = {cell.value: cell.column for cell in ws[1]}
    for name, col in headers.items():
        for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2, max_row=ws.max_row):
            for c in cell:
                if name in date_cols:
                    c.number_format = "dd/mm/yyyy"
                elif name in time_cols:
                    c.number_format = "hh:mm:ss"
                elif name in dt_cols:
                    c.number_format = "dd/mm/yyyy hh:mm:ss"
                elif name in amount_cols:
                    c.number_format = '#,##0.000'
    widths = {
        "Merchant Location MID": 24, "Terminal ID": 16, "Settlement Date": 16,
        "Trxn Date": 16, "Trxn Time": 13, "Trxn DateTime": 21, "Batch No": 27,
        "Invoice No": 24, "Auth Code": 13, "Card No": 22, "Card Group": 13,
        "Trxn Type": 20, "Gross Amount": 17, "Discount Amount": 18,
        "VAT Amount": 15, "Net Amount": 17,
    }
    for name, col in headers.items():
        ws.column_dimensions[get_column_letter(col)].width = widths.get(name, 18)

    # Use a normal worksheet AutoFilter instead of an Excel Table.
    # This avoids Excel repair warnings caused by table metadata while
    # keeping filtering, freeze panes, and the normal worksheet formatting.


def dataframe_to_workbook(df: pd.DataFrame, source_name: str, stats: dict) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"
    for c, name in enumerate(df.columns, 1):
        ws.cell(1, c, name)
    for r_idx, row in enumerate(df.itertuples(index=False, name=None), 2):
        for c_idx, value in enumerate(row, 1):
            if pd.isna(value):
                value = None
            ws.cell(r_idx, c_idx, value)
    _style_sheet(ws, table_name="TransactionsTable")

    info = wb.create_sheet("Processing Info")
    info.append(["Source File", source_name])
    info.append(["Source Rows", stats["source_rows"]])
    info.append(["Detected Transaction Header Blocks", stats["detected_header_blocks"]])
    info.append(["Transactions Output", stats["transactions"]])
    info.append(["Processing", "Dynamic/content-based parsing; no fixed row count."])
    info.column_dimensions["A"].width = 34
    info.column_dimensions["B"].width = 55
    for cell in info[1]:
        cell.font = Font(bold=True)
    return wb


def save_cleaned_file(source_path: str | Path, output_path: str | Path) -> dict:
    df, stats = clean_statement(source_path)
    if df.empty:
        raise ValueError("No transaction rows were detected. The file may use a different statement layout.")
    wb = dataframe_to_workbook(df, Path(source_path).name, stats)
    wb.save(output_path)
    return stats


def extract_inputs(paths: list[str | Path], temp_dir: Path) -> list[Path]:
    files = []
    for p in paths:
        p = Path(p)
        if p.suffix.lower() == ".zip":
            with zipfile.ZipFile(p) as z:
                for name in z.namelist():
                    if Path(name).suffix.lower() in {".xls", ".xlsx", ".xlsm"} and not name.endswith("/"):
                        target = temp_dir / Path(name).name
                        target.parent.mkdir(parents=True, exist_ok=True)
                        with z.open(name) as src, open(target, "wb") as dst:
                            shutil.copyfileobj(src, dst)
                        files.append(target)
        elif p.suffix.lower() in {".xls", ".xlsx", ".xlsm"}:
            files.append(p)
    return files


def process_files(paths: list[str | Path], output_dir: str | Path, mode="separate", group_size=2, progress=None):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="merchant_cleaner_") as td:
        files = extract_inputs(paths, Path(td))
        if not files:
            raise ValueError("No Excel files were found.")
        results = []
        total = len(files)

        cleaned = []
        for i, src in enumerate(files, 1):
            df, stats = clean_statement(src)
            if df.empty:
                raise ValueError(f"No transaction rows detected in: {src.name}")
            cleaned.append((src, df, stats))
            if progress:
                progress(i / max(total, 1) * (0.75 if mode != "separate" else 0.9), f"Cleaning {i}/{total}: {src.name}")

        if mode == "separate":
            for i, (src, df, stats) in enumerate(cleaned, 1):
                out = output_dir / f"{src.stem}_Clean.xlsx"
                dataframe_to_workbook(df, src.name, stats).save(out)
                results.append(out)
                if progress:
                    progress(0.9 + 0.1 * i / total, f"Saved {out.name}")
        else:
            chunks = [cleaned[i:i + group_size] for i in range(0, len(cleaned), group_size)] if mode == "group" else [cleaned]
            for idx, chunk in enumerate(chunks, 1):
                frames = [x[1] for x in chunk]
                merged = pd.concat(frames, ignore_index=True)
                source_names = [x[0].name for x in chunk]
                stats = {"source_rows": sum(x[2]["source_rows"] for x in chunk), "detected_header_blocks": sum(x[2]["detected_header_blocks"] for x in chunk), "transactions": len(merged)}
                stem = "Merged" if mode == "merge" else f"Group_{idx:02d}"
                out = output_dir / f"{stem}_Clean.xlsx"
                dataframe_to_workbook(merged, ", ".join(source_names), stats).save(out)
                results.append(out)
                if progress:
                    progress(0.75 + 0.25 * idx / len(chunks), f"Saved {out.name}")
    return results
