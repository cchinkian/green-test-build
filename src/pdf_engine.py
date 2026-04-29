"""
PDF engine — overlays text onto flat/scanned PDF forms.

Field source types (set in forms.json per field):
  data     — from Excel (Master or batch sheet); changes every client/transaction
  settings — from settings.json; RM name, branch, staff ID; set once
  fixed    — hardcoded in forms.json; never changes (currency, bank name, etc.)
  auto     — app generates at runtime (today's date, etc.)

Coordinate system: (0,0) = bottom-left, same as PDF/reportlab standard.
"""
import io
from datetime import date as dt_date
from pathlib import Path

from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4
import pypdf


# ── Value resolver ────────────────────────────────────────────────────────────

def _resolve(field: dict, client: dict, settings: dict) -> str:
    """Return the string value to write for a field, based on its source."""
    source = field.get("source", "data")

    if source == "fixed":
        # Value is hardcoded in forms.json — e.g. "MYR", "RHB Bank Berhad"
        return str(field.get("value", ""))

    if source == "settings":
        # Value comes from settings.json — e.g. rm_name, rm_branch, rm_staff_id
        key = field.get("settings_key") or field.get("name", "")
        return str(settings.get(key, field.get("value", "")))

    if source == "auto":
        # App generates the value at runtime
        kind = field.get("auto_type") or field.get("name", "date")
        if kind == "date":
            fmt = settings.get("date_format", "%d/%m/%Y")
            return dt_date.today().strftime(fmt)
        if kind == "year":
            return str(dt_date.today().year)
        if kind == "month":
            return dt_date.today().strftime("%B")
        # Extendable: add more auto types here (ref_no, time, etc.)
        return ""

    # source == "data" (default) — from Excel client/transaction data
    key = field.get("data_key") or field.get("name", "")
    return str(client.get(key, ""))


# ── Overlay builder ───────────────────────────────────────────────────────────

def _make_overlay(pw: float, ph: float,
                  fields: list[dict],
                  client: dict,
                  settings: dict) -> bytes:
    """Create a transparent reportlab page with text at field coordinates."""
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(pw, ph))
    for field in fields:
        value = _resolve(field, client, settings)
        if not value:
            continue
        font      = field.get("font",      settings.get("default_font",      "Helvetica"))
        font_size = field.get("font_size", settings.get("default_font_size", 10))
        c.setFont(font, font_size)
        c.drawString(field["x"], field["y"], value)
    c.save()
    return buf.getvalue()


# ── Form filler ───────────────────────────────────────────────────────────────

def fill_form(template_path: Path,
              form_config: dict,
              client: dict,
              settings: dict,
              output_path: Path):
    """Overlay all field values onto a flat PDF template and save."""
    reader = pypdf.PdfReader(str(template_path))
    writer = pypdf.PdfWriter()

    # Group fields by page (1-based)
    by_page: dict[int, list] = {}
    for field in form_config.get("fields", []):
        pg = field.get("page", 1)
        by_page.setdefault(pg, []).append(field)

    for i, page in enumerate(reader.pages):
        page_num = i + 1
        box = page.mediabox
        pw, ph = float(box.width), float(box.height)

        page_fields = by_page.get(page_num, [])
        if page_fields:
            overlay_bytes = _make_overlay(pw, ph, page_fields, client, settings)
            overlay_page = pypdf.PdfReader(io.BytesIO(overlay_bytes)).pages[0]
            page.merge_page(overlay_page)

        writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)


# ── Bundle filler ─────────────────────────────────────────────────────────────

def fill_bundle(application: dict,
                forms_config: dict,
                client: dict,
                output_folder: Path,
                settings: dict,
                find_template_fn) -> list[Path]:
    """
    Fill all forms in an application bundle for one client.
    Returns list of output PDF paths.
    """
    today_safe = dt_date.today().strftime("%Y%m%d")
    client_name_safe = (client.get("name", "Client")
                        .replace(" ", "_")
                        .replace("/", "-"))
    results = []

    for form_id in application.get("forms", []):
        form_cfg = forms_config.get(form_id)
        if not form_cfg:
            raise ValueError(
                f"Form '{form_id}' not found in forms.json. "
                "Add it or check the spelling."
            )
        template_path = find_template_fn(settings, form_cfg["template_subfolder"])
        filename = f"{client_name_safe}_{form_id}_{today_safe}.pdf"
        output_path = output_folder / filename
        fill_form(template_path, form_cfg, client, settings, output_path)
        results.append(output_path)

    return results
