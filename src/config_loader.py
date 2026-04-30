"""
Config loader — resolves paths relative to .exe so app works from any pen drive.
"""
import json
import sys
from pathlib import Path


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def _config_path(filename: str) -> Path:
    return _base_dir() / "config" / filename


def load_settings() -> dict:
    with open(_config_path("settings.json"), encoding="utf-8") as f:
        return json.load(f)


def load_forms() -> dict:
    with open(_config_path("forms.json"), encoding="utf-8") as f:
        return json.load(f)


def save_forms(data: dict):
    with open(_config_path("forms.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def load_applications() -> list:
    with open(_config_path("applications.json"), encoding="utf-8") as f:
        return json.load(f)


def find_template(settings: dict, template_subfolder: str) -> Path:
    """
    Find the single PDF at the top level of a form subfolder.
    Raises if zero or MORE than one PDF found — ambiguity is a compliance risk.
    Old versions must be kept in a sub-subfolder, not at the top level.
    """
    forms_folder = Path(settings["forms_folder"])
    subfolder    = forms_folder / template_subfolder
    pdfs = [p for p in subfolder.iterdir()
            if p.is_file() and p.suffix.lower() == ".pdf"]

    if not pdfs:
        raise FileNotFoundError(
            f"No PDF found in:\n  {subfolder}\n"
            "Add the blank form template there."
        )
    if len(pdfs) > 1:
        names = ", ".join(p.name for p in sorted(pdfs))
        raise ValueError(
            f"Multiple PDFs in {subfolder}:\n  {names}\n"
            "Keep exactly ONE pdf at the top level. "
            "Move old versions into a subfolder (e.g. 'old forms/')."
        )
    return pdfs[0]


def data_path(filename: str) -> Path:
    return _base_dir() / "data" / filename


def log_path() -> Path:
    return _base_dir() / "data" / "fill_log.csv"
