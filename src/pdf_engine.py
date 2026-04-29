"""
PDF engine — overlays text onto flat/scanned PDF forms using reportlab + pypdf.
Coordinate system: (0,0) = bottom-left of page, same as PDF/reportlab standard.
"""
import io
from datetime import date
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pypdf


def _make_overlay(page_width: float, page_height: float, fields: list[dict], client: dict) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_width, page_height))
    for field in fields:
        value = client.get(field["name"], "")
        if not value:
            continue
        font = field.get("font", "Helvetica")
        size = field.get("font_size", 10)
        c.setFont(font, size)
        c.drawString(field["x"], field["y"], value)
    c.save()
    return buf.getvalue()


def fill_form(template_path: Path, form_config: dict, client: dict, output_path: Path):
    """Overlay client data onto a flat PDF template and save."""
    reader = pypdf.PdfReader(str(template_path))
    writer = pypdf.PdfWriter()

    # Group fields by page number (1-based)
    fields_by_page: dict[int, list] = {}
    for field in form_config.get("fields", []):
        pg = field.get("page", 1)
        fields_by_page.setdefault(pg, []).append(field)

    for i, page in enumerate(reader.pages):
        page_num = i + 1
        box = page.mediabox
        pw = float(box.width)
        ph = float(box.height)

        page_fields = fields_by_page.get(page_num, [])
        if page_fields:
            overlay_bytes = _make_overlay(pw, ph, page_fields, client)
            overlay_reader = pypdf.PdfReader(io.BytesIO(overlay_bytes))
            page.merge_page(overlay_reader.pages[0])

        writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)


def fill_bundle(application: dict, forms_config: dict, client: dict,
                output_folder: Path, settings: dict, find_template_fn):
    """Fill all forms in a bundle for one client. Returns list of output paths."""
    from datetime import date as dt
    today = dt.today().strftime(settings.get("date_format", "%d/%m/%Y").replace("%", "").replace("/", "-"))
    today_safe = dt.today().strftime("%Y%m%d")

    client_name_safe = client.get("name", "Client").replace(" ", "_").replace("/", "-")

    # Inject today's date into client data for forms that need it
    client_with_date = dict(client)
    client_with_date.setdefault("date", dt.today().strftime(settings.get("date_format", "%d/%m/%Y")))

    results = []
    for form_id in application.get("forms", []):
        form_cfg = forms_config.get(form_id)
        if not form_cfg:
            continue
        template_path = find_template_fn(settings, form_cfg["template_subfolder"])
        filename = f"{client_name_safe}_{form_id}_{today_safe}.pdf"
        output_path = output_folder / filename
        fill_form(template_path, form_cfg, client_with_date, output_path)
        results.append(output_path)
    return results
