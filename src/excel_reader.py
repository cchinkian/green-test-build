"""
Excel reader — multi-sheet, fully dynamic.

Fix 8: IC numbers normalized on both sides of the join (strip dashes/spaces,
        zero-pad to 12) so "880101-14-5678" == "880101145678" always match.
Fix 9: Native Python types (datetime, int, float) preserved through the reader
        so formatters in pdf_engine receive typed values, not pre-cast strings.
"""
import re
import openpyxl
from pathlib import Path


# ── IC normalisation ─────────────────────────────────────────────────────────

def normalize_ic(raw) -> str:
    """Strip non-digits, zero-pad to 12. Empty input returns ''."""
    cleaned = re.sub(r"[^0-9]", "", str(raw))
    return cleaned.zfill(12) if cleaned else ""


# ── Sheet parsing ─────────────────────────────────────────────────────────────

def _sheet_to_dicts(ws) -> list[dict]:
    """
    Parse a worksheet into a list of dicts.
    Native types (datetime, int, float) are preserved — only None → "" and
    bare strings are stripped. Formatters in pdf_engine handle the rest.
    """
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [
        str(h).strip().lower().replace(" ", "_") if h else f"col_{i}"
        for i, h in enumerate(rows[0])
    ]
    result = []
    for row in rows[1:]:
        if not any(row):
            continue
        record = {}
        for k, v in zip(headers, row):
            if v is None:
                record[k] = ""
            elif isinstance(v, str):
                record[k] = v.strip()
            else:
                record[k] = v  # int, float, datetime preserved
        # Normalize ic_number if present
        if "ic_number" in record:
            record["ic_number"] = normalize_ic(record["ic_number"])
        result.append(record)
    return result


# ── Public loaders ────────────────────────────────────────────────────────────

def load_master(path: Path) -> dict[str, dict]:
    """Returns {normalized_ic: client_dict} from the Master sheet."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    clients = _sheet_to_dicts(wb["Master"])
    wb.close()
    return {c["ic_number"]: c for c in clients if c.get("ic_number")}


def get_sheet_headers(path: Path, sheet_name: str) -> list[str]:
    """Column names from a batch sheet, excluding ic_number."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[sheet_name]
    first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    wb.close()
    return [
        str(h).strip().lower().replace(" ", "_")
        for h in first_row
        if h and str(h).strip().lower().replace(" ", "_") != "ic_number"
    ]


def load_batch(path: Path, sheet_name: str,
               master: dict[str, dict]) -> list[dict]:
    """
    Merged records: Master static + batch transaction data.
    Joined by normalized ic_number. Skips rows with no Master match.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available: {wb.sheetnames}"
        )
    rows = _sheet_to_dicts(wb[sheet_name])
    wb.close()
    merged = []
    for row in rows:
        ic = row.get("ic_number", "")
        master_rec = master.get(ic)
        if not master_rec:
            continue
        merged.append({**master_rec, **row})
    return merged


def find_client_in_batch(path: Path, sheet_name: str, ic_number: str) -> dict:
    """Return batch row for one client (by normalized IC), or {} if not found."""
    ic_norm = normalize_ic(ic_number)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return {}
    rows = _sheet_to_dicts(wb[sheet_name])
    wb.close()
    for row in rows:
        if row.get("ic_number") == ic_norm:
            return row
    return {}


def get_master_names(master: dict[str, dict]) -> list[str]:
    return sorted(c.get("name", "") for c in master.values() if c.get("name"))


def sheet_names(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names
