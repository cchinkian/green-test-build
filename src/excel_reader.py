"""
Excel reader — loads client data from clients.xlsx.
Returns list of dicts keyed by column headers (lowercase, spaces→underscores).
"""
import openpyxl
from pathlib import Path


def load_clients(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip().lower().replace(" ", "_") if h else "" for h in rows[0]]
    clients = []
    for row in rows[1:]:
        if not any(row):
            continue
        record = {}
        for key, val in zip(headers, row):
            record[key] = str(val).strip() if val is not None else ""
        clients.append(record)
    wb.close()
    return clients


def get_client_names(clients: list[dict]) -> list[str]:
    return sorted(c.get("name", "") for c in clients if c.get("name"))
