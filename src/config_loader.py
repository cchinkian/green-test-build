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


def scan_forms_folder(settings: dict) -> list[str]:
    """
    Auto-detect: list all subdirectory names inside forms_folder.
    Returns sorted list of folder names (not full paths).
    Returns [] if forms_folder doesn't exist.
    """
    forms_folder = Path(settings.get("forms_folder", ""))
    if not forms_folder.exists():
        return []
    return sorted(p.name for p in forms_folder.iterdir() if p.is_dir())


def find_template(settings: dict, template_subfolder: str,
                  form_config: dict | None = None,
                  test_mode: bool = False) -> Path:
    """
    Find the single PDF at the top level of a form subfolder.

    Surrender logic (filename-based):
      - forms.json stores template_filename = last known PDF name
      - If current PDF has a DIFFERENT name → form is surrendered
      - Surrendered forms: raises TemplateSurrenderedError unless test_mode=True
      - test_mode=True → returns the new PDF path with a _TEST_ fill warning

    Compliance guard:
      - Zero PDFs → FileNotFoundError
      - >1 PDF → ValueError (must keep exactly one at top level)

    Hash check (content-based, secondary):
      - If filename matches but hash differs → TemplateChangedWarning
    """
    forms_folder = Path(settings["forms_folder"])
    subfolder    = forms_folder / template_subfolder
    pdfs = [p for p in subfolder.iterdir()
            if p.is_file() and p.suffix.lower() == ".pdf"]

    if not pdfs:
        raise FileNotFoundError(
            f"No PDF found in:\n  {subfolder}\n"
            f"Place the blank form PDF in C:\\Forms\\{template_subfolder}\\"
        )
    if len(pdfs) > 1:
        names = ", ".join(p.name for p in sorted(pdfs))
        raise ValueError(
            f"Multiple PDFs in {subfolder}:\n  {names}\n"
            "Keep exactly ONE PDF at the top level. "
            "Move old versions into a subfolder (e.g. 'old forms\\')."
        )

    pdf_path = pdfs[0]

    if form_config:
        stored_name = form_config.get("template_filename", "")
        stored_hash = form_config.get("template_hash", "")
        form_label  = form_config.get("name", template_subfolder)

        # Filename changed → surrender (primary check)
        if stored_name and pdf_path.name != stored_name:
            if not test_mode:
                raise TemplateSurrenderedError(
                    f"Form PDF was replaced — old coordinates are surrendered.\n\n"
                    f"Form:     {form_label}\n"
                    f"Was:      {stored_name}\n"
                    f"Now:      {pdf_path.name}\n\n"
                    "Re-map this form in CoordPicker to restore normal fills.\n"
                    "Or click 'Test Fill' to preview with old coordinates "
                    "(output will be marked _TEST_).",
                    pdf_path, stored_name
                )
            # test_mode: return path anyway (caller adds _TEST_ prefix)
            return pdf_path

        # Hash check — runs when:
        #   (a) stored_name matches current filename (same file, check content), OR
        #   (b) no stored_name (old format without filename tracking, check hash only)
        if stored_hash:
            current_hash = _md5(pdf_path)
            if current_hash != stored_hash:
                raise TemplateChangedWarning(
                    f"PDF content changed since last mapping.\n"
                    f"Form: {form_label}\n"
                    f"File: {pdf_path.name}\n\n"
                    "Re-map in CoordPicker to update field positions.",
                    pdf_path
                )

    return pdf_path


def compute_template_hash(pdf_path: Path) -> str:
    return _md5(pdf_path)


class TemplateChangedWarning(Exception):
    """Same filename, different content — coordinates may be off."""
    def __init__(self, message: str, pdf_path: Path):
        super().__init__(message)
        self.pdf_path = pdf_path


class TemplateSurrenderedError(Exception):
    """PDF filename changed — coordinates are surrendered. Test-only until re-mapped."""
    def __init__(self, message: str, pdf_path: Path, old_filename: str):
        super().__init__(message)
        self.pdf_path    = pdf_path
        self.old_filename = old_filename


# ── Health check ──────────────────────────────────────────────────────────────

def health_check(settings: dict, forms: dict) -> list[dict]:
    """
    Verify forms_folder + each mapped form subfolder.
    Statuses: ok | warn | error | surrendered
    Also reports discovered subfolders not yet in forms.json (unmapped).
    """
    results = []
    forms_folder = Path(settings.get("forms_folder", ""))

    if not forms_folder.exists():
        return [{"name": "forms_folder", "status": "error",
                 "message": f"Folder not found: {forms_folder}"}]

    # Check each registered form
    registered_subfolders = set()
    for form_id, form_cfg in forms.items():
        if form_id.startswith("_"):
            continue
        subfolder = form_cfg.get("template_subfolder", "")
        if not subfolder:
            continue
        registered_subfolders.add(subfolder)
        path = forms_folder / subfolder

        if not path.exists():
            results.append({"name": form_id, "status": "error",
                            "message": f"Subfolder missing: {subfolder}"})
            continue

        pdfs = [p for p in path.iterdir()
                if p.is_file() and p.suffix.lower() == ".pdf"]

        if not pdfs:
            results.append({"name": form_id, "status": "warn",
                            "message": f"No PDF in {subfolder}\\"})
        elif len(pdfs) > 1:
            names = ", ".join(p.name for p in sorted(pdfs))
            results.append({"name": form_id, "status": "warn",
                            "message": f"Multiple PDFs: {names}"})
        else:
            current_name   = pdfs[0].name
            stored_name    = form_cfg.get("template_filename", "")
            if stored_name and current_name != stored_name:
                results.append({
                    "name": form_id, "status": "surrendered",
                    "message": (f"PDF replaced: {stored_name} → {current_name}. "
                                "Re-map in CoordPicker. Test fills allowed.")
                })
            else:
                results.append({"name": form_id, "status": "ok",
                                "message": current_name})

    # Report unregistered subfolders (discovered but not in forms.json)
    for sub in scan_forms_folder(settings):
        if sub not in registered_subfolders:
            results.append({"name": f"[{sub}]", "status": "unmapped",
                            "message": f"Found in C:\\Forms but not in forms.json"})

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
