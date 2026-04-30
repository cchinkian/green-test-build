"""
Config loader — pen-drive portable path resolution.

P3-10: find_template() hashes the PDF and warns if it changed since mapping.
P3-11: save_forms() backs up forms.json before overwriting.
P1-3:  health_check() verifies forms_folder + mapped subfolders.
"""
import hashlib
import json
import shutil
import sys
from pathlib import Path


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def _config_path(filename: str) -> Path:
    return _base_dir() / "config" / filename


# ── Settings ──────────────────────────────────────────────────────────────────

def load_settings() -> dict:
    with open(_config_path("settings.json"), encoding="utf-8") as f:
        return json.load(f)


def save_settings(data: dict):
    with open(_config_path("settings.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ── Forms ─────────────────────────────────────────────────────────────────────

def load_forms() -> dict:
    with open(_config_path("forms.json"), encoding="utf-8") as f:
        return json.load(f)


def save_forms(data: dict):
    """P3-11: Back up before overwriting."""
    forms_path = _config_path("forms.json")
    bak_path   = _config_path("forms.json.bak")
    if forms_path.exists():
        shutil.copy(forms_path, bak_path)
    with open(forms_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ── Applications ──────────────────────────────────────────────────────────────

def load_applications() -> list:
    with open(_config_path("applications.json"), encoding="utf-8") as f:
        return json.load(f)


# ── Template PDF ──────────────────────────────────────────────────────────────

def _md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def find_template(settings: dict, template_subfolder: str,
                  form_config: dict | None = None) -> Path:
    """
    Find the single PDF at the top level of a form subfolder.
    Raises if zero or >1 PDF found (compliance guard).
    P3-10: warns if PDF hash differs from stored template_hash.
    Returns (path, hash_changed: bool) but callers can ignore the bool.
    """
    forms_folder = Path(settings["forms_folder"])
    subfolder    = forms_folder / template_subfolder
    pdfs = [p for p in subfolder.iterdir()
            if p.is_file() and p.suffix.lower() == ".pdf"]

    if not pdfs:
        raise FileNotFoundError(
            f"No PDF found in:\n  {subfolder}\n"
            f"Create the folder C:\\Forms\\{template_subfolder}\\ "
            "and place the blank form PDF there."
        )
    if len(pdfs) > 1:
        names = ", ".join(p.name for p in sorted(pdfs))
        raise ValueError(
            f"Multiple PDFs in {subfolder}:\n  {names}\n"
            "Keep exactly ONE PDF at the top level. "
            "Move old versions into a subfolder (e.g. 'old forms\\')."
        )

    pdf_path = pdfs[0]

    # P3-10: hash check
    if form_config:
        stored_hash = form_config.get("template_hash", "")
        if stored_hash:
            current_hash = _md5(pdf_path)
            if current_hash != stored_hash:
                raise TemplateChangedWarning(
                    f"The PDF for this form has changed since coordinates were mapped.\n"
                    f"Form: {form_config.get('name', template_subfolder)}\n"
                    f"File: {pdf_path.name}\n\n"
                    "Re-run CoordPicker for this form to update the field positions.",
                    pdf_path
                )

    return pdf_path


def compute_template_hash(pdf_path: Path) -> str:
    return _md5(pdf_path)


class TemplateChangedWarning(Exception):
    """Raised when the PDF template has changed since coordinates were mapped."""
    def __init__(self, message: str, pdf_path: Path):
        super().__init__(message)
        self.pdf_path = pdf_path


# ── Health check ──────────────────────────────────────────────────────────────

def health_check(settings: dict, forms: dict) -> list[dict]:
    """
    P1-3: Verify forms_folder exists and each mapped form subfolder is present.
    Returns list of {name, status, message} dicts.
    Skips _TEMPLATE and _FIELD_REFERENCE entries.
    """
    results = []
    forms_folder = Path(settings.get("forms_folder", ""))

    if not forms_folder.exists():
        return [{"name": "forms_folder", "status": "error",
                 "message": f"Folder not found: {forms_folder}"}]

    for form_id, form_cfg in forms.items():
        if form_id.startswith("_"):
            continue
        subfolder = form_cfg.get("template_subfolder", "")
        if not subfolder:
            continue
        path = forms_folder / subfolder
        if not path.exists():
            results.append({"name": form_id, "status": "error",
                            "message": f"Missing: {path}"})
        else:
            pdfs = [p for p in path.iterdir()
                    if p.is_file() and p.suffix.lower() == ".pdf"]
            if not pdfs:
                results.append({"name": form_id, "status": "warn",
                                "message": f"No PDF in: {path}"})
            elif len(pdfs) > 1:
                results.append({"name": form_id, "status": "warn",
                                "message": f"Multiple PDFs in: {path}"})
            else:
                results.append({"name": form_id, "status": "ok",
                                "message": str(pdfs[0].name)})
    return results


# ── Data / state paths ────────────────────────────────────────────────────────

def data_path(filename: str) -> Path:
    return _base_dir() / "data" / filename


def log_path() -> Path:
    return _base_dir() / "data" / "fill_log.csv"


def state_path() -> Path:
    return _base_dir() / "data" / "state.json"


def load_state() -> dict:
    p = state_path()
    if p.exists():
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_state(data: dict):
    p = state_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f)
