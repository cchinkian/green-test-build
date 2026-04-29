"""
Excel reader — multi-sheet support.
Master sheet: client static info (one row per client, keyed by ic_number).
Batch sheets: transaction-specific data joined to Master by ic_number.
"""
import openpyxl
from pathlib import Path


def _sheet_to_dicts(ws) -> list[dict]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip().lower().replace(" ", "_") if h else "" for h in rows[0]]
    result = []
    for row in rows[1:]:
        if not any(row):
            continue
        record = {k: (str(v).strip() if v is not None else "") for k, v in zip(headers, row)}
        result.append(record)
    return result


def load_master(path: Path) -> dict[str, dict]:
    """Returns {ic_number: client_dict} from the Master sheet."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Master"]
    clients = _sheet_to_dicts(ws)
    wb.close()
    return {c["ic_number"]: c for c in clients if c.get("ic_number")}


def load_batch_sheet(path: Path, sheet_name: str) -> list[dict]:
    """Returns rows from a named batch sheet (e.g. UT_Subscription)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Sheet '{sheet_name}' not found in {path.name}. "
                         f"Available: {wb.sheetnames}")
    ws = wb[sheet_name]
    rows = _sheet_to_dicts(ws)
    wb.close()
    return rows


def get_master_names(master: dict[str, dict]) -> list[str]:
    """Sorted list of client names from master."""
    return sorted(c.get("name", "") for c in master.values() if c.get("name"))


def get_batch_clients(path: Path, sheet_name: str,
                      master: dict[str, dict]) -> list[dict]:
    """
    Returns merged records for clients in a batch sheet.
    Master static data + batch transaction data, joined by ic_number.
    Clients in batch but missing from Master are skipped with a warning.
    """
    batch_rows = load_batch_sheet(path, sheet_name)
    merged = []
    for row in batch_rows:
        ic = row.get("ic_number", "")
        master_rec = master.get(ic)
        if not master_rec:
            continue  # not in master — skip
        merged_rec = {**master_rec, **row}  # batch overrides master if same key
        merged.append(merged_rec)
    return merged


def sheet_names(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names
