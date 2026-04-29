"""
Config loader — resolves paths relative to .exe location so app works from any pen drive.
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
    path = _config_path("settings.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_forms() -> dict:
    path = _config_path("forms.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_forms(data: dict):
    path = _config_path("forms.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def load_applications() -> list:
    path = _config_path("applications.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def find_template(settings: dict, template_subfolder: str) -> Path:
    """
    Find the single PDF at the top level of a subfolder inside forms_folder.
    Old versions are kept in sub-subfolders so this only picks the current one.
    """
    forms_folder = Path(settings["forms_folder"])
    subfolder = forms_folder / template_subfolder
    pdfs = [p for p in subfolder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    if not pdfs:
        raise FileNotFoundError(f"No PDF found in {subfolder}")
    if len(pdfs) > 1:
        # If multiple, take the alphabetically last (most likely latest version)
        pdfs.sort()
        return pdfs[-1]
    return pdfs[0]


def data_path(filename: str) -> Path:
    return _base_dir() / "data" / filename
