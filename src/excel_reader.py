"""
Excel reader — multi-sheet, fully dynamic.
Master sheet: client static info keyed by ic_number.
Batch sheets: any columns, joined to Master by ic_number.
All column names are auto-discovered — no hardcoding.
"""
import openpyxl
from pathlib import Path


def _sheet_to_dicts(ws) -> list[dict]:
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
        result.append({k: (str(v).strip() if v is not None else "")
                       for k, v in zip(headers, row)})
    return result


def load_master(path: Path) -> dict[str, dict]:
    """Returns {ic_number: client_dict} from Master sheet."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    clients = _sheet_to_dicts(wb["Master"])
    wb.close()
    return {c["ic_number"]: c for c in clients if c.get("ic_number")}


def get_sheet_headers(path: Path, sheet_name: str) -> list[str]:
    """Column names from a batch sheet (excluding ic_number — auto-joined)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    wb.close()
    headers = [
        str(h).strip().lower().replace(" ", "_") if h else ""
        for h in first_row
    ]
    return [h for h in headers if h and h != "ic_number"]


def load_batch(path: Path, sheet_name: str,
               master: dict[str, dict]) -> list[dict]:
    """
    Returns merged records: Master static + batch transaction data.
    Joined by ic_number. Skips batch rows with no matching Master entry.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {wb.sheetnames}"
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


def find_client_in_batch(path: Path, sheet_name: str,
                         ic_number: str) -> dict:
    """Return batch row for one client, or {} if not found."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return {}
    rows = _sheet_to_dicts(wb[sheet_name])
    wb.close()
    for row in rows:
        if row.get("ic_number") == ic_number:
            return row
    return {}


def get_master_names(master: dict[str, dict]) -> list[str]:
    return sorted(c.get("name", "") for c in master.values() if c.get("name"))


def sheet_names(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names
