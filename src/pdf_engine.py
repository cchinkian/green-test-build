"""
PDF engine — coordinate-based text overlay for flat/scanned PDF forms.

Field source types (per-field in forms.json):
  data     — from Excel column (Master or batch sheet)
  settings — from settings.json (rm_name, rm_branch, rm_staff_id)
  fixed    — hardcoded in forms.json value (currency, bank name, etc.)
  auto     — generated at runtime (date, year, month)

P3-9: _shared_fields support — common fields defined once, referenced by name.
P1-4: Required blank fields prefix output with _REVIEW_ and return warnings.

Coordinates: (0,0) = bottom-left. A4 top ≈ y=842.
"""
import csv
import io
import datetime
from pathlib import Path

from reportlab.pdfgen import canvas as rl_canvas
import pypdf


# ── Shared field expansion (P3-9) ─────────────────────────────────────────────

def _expand_shared(fields: list[dict], shared: dict) -> list[dict]:
    """
    Resolve "shared" references at fill time.
    Field with {"shared": "client_name", "page": 1, "x": 120, "y": 700}
    inherits all attributes from _shared_fields["client_name"],
    then the per-form keys (x, y, page) override.
    """
    out = []
    for f in fields:
        ref = f.get("shared")
        if ref:
            base = dict(shared.get(ref, {}))
            base.update(f)           # per-form coords override shared attrs
            base.pop("shared", None)
            out.append(base)
        else:
            out.append(f)
    return out


# ── Value formatter ───────────────────────────────────────────────────────────

def _format_value(raw, fmt: str) -> str:
    s = str(raw).strip() if not isinstance(raw, (datetime.datetime,
                                                   datetime.date)) else ""
    if fmt == "currency_myr":
        try:
            n = float(str(raw).replace(",", "").replace("RM", "").strip())
            return f"RM {n:,.2f}"
        except (ValueError, TypeError):
            return s

    if fmt == "currency_no_symbol":
        try:
            n = float(str(raw).replace(",", "").replace("RM", "").strip())
            return f"{n:,.2f}"
        except (ValueError, TypeError):
            return s

    if fmt in ("date_dmy", "date_dmy_slash"):
        if isinstance(raw, (datetime.datetime, datetime.date)):
            return raw.strftime("%d/%m/%Y")
        for pat in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.datetime.strptime(s, pat).strftime("%d/%m/%Y")
            except ValueError:
                pass
        return s

    if fmt == "date_dmy_long":
        if isinstance(raw, (datetime.datetime, datetime.date)):
            return raw.strftime("%d %b %Y")
        for pat in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.datetime.strptime(s, pat).strftime("%d %b %Y")
            except ValueError:
                pass
        return s

    if fmt == "ic_dashed":
        digits = "".join(c for c in str(raw) if c.isdigit()).zfill(12)
        return f"{digits[:6]}-{digits[6:8]}-{digits[8:]}" if len(digits) == 12 else str(raw)

    if fmt == "phone_dashed":
        digits = "".join(c for c in str(raw) if c.isdigit())
        return f"{digits[:3]}-{digits[3:]}" if len(digits) >= 10 else str(raw)

    if fmt == "uppercase":
        return s.upper()

    if fmt == "integer":
        try:
            return str(int(float(str(raw).replace(",", ""))))
        except (ValueError, TypeError):
            return s

    return s


# ── Value resolver ────────────────────────────────────────────────────────────

def _to_str(value) -> str:
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()


def _resolve(field: dict, client: dict, settings: dict) -> str:
    source = field.get("source", "data")

    if source == "fixed":
        raw = field.get("value", "")
    elif source == "settings":
        key = field.get("settings_key") or field.get("name", "")
        raw = settings.get(key, field.get("value", ""))
    elif source == "auto":
        kind = field.get("auto_type") or field.get("name", "date")
        today = datetime.date.today()
        if kind == "date":
            return today.strftime(settings.get("date_format", "%d/%m/%Y"))
        if kind == "year":
            return str(today.year)
        if kind == "month":
            return today.strftime("%B")
        return ""
    else:  # data
        key = field.get("data_key") or field.get("name", "")
        raw = client.get(key, "")

    if raw == "" or raw is None:
        return ""

    fmt = field.get("format")
    return _format_value(raw, fmt) if fmt else _to_str(raw)


# ── Overlay builder ───────────────────────────────────────────────────────────

def _make_overlay(pw: float, ph: float, fields: list[dict],
                  client: dict, settings: dict) -> tuple[bytes, list[str]]:
    """Returns (pdf_bytes, list_of_blank_required_field_names)."""
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(pw, ph))
    blanks = []
    for field in fields:
        value = _resolve(field, client, settings)
        if not value:
            if field.get("required", False):
                blanks.append(field.get("name", "?"))
            continue
        font      = field.get("font",      settings.get("default_font",      "Helvetica"))
        font_size = field.get("font_size", settings.get("default_font_size", 10))
        c.setFont(font, font_size)
        c.drawString(field["x"], field["y"], value)
    c.save()
    return buf.getvalue(), blanks


# ── Audit log ─────────────────────────────────────────────────────────────────

_LOG_COLS = ["timestamp", "rm_staff_id", "application_id",
             "client_initials", "ic_last4",
             "form_id", "output_file", "status", "blank_fields"]


def _initials(name: str) -> str:
    return "".join(w[0].upper() for w in name.split() if w)


def _write_audit(log_path: Path, row: dict):
    write_header = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=_LOG_COLS)
        if write_header:
            writer.writeheader()
        writer.writerow(row)
        f.flush()


# ── Form filler ───────────────────────────────────────────────────────────────

def fill_form(template_path: Path, form_config: dict,
              client: dict, settings: dict,
              output_path: Path) -> list[str]:
    """
    Overlay all fields onto a flat PDF. Returns blank required field names.
    P1-4: If blanks exist, prefix output filename with _REVIEW_.
    """
    # P3-9: expand shared field references
    shared = settings.get("_shared_fields", {})  # passed via settings for convenience
    all_fields = _expand_shared(form_config.get("fields", []), shared)

    reader = pypdf.PdfReader(str(template_path))
    writer = pypdf.PdfWriter()
    all_blanks = []

    by_page: dict[int, list] = {}
    for field in all_fields:
        by_page.setdefault(field.get("page", 1), []).append(field)

    for i, page in enumerate(reader.pages):
        box = page.mediabox
        pw, ph = float(box.width), float(box.height)
        page_fields = by_page.get(i + 1, [])
        if page_fields:
            overlay_bytes, blanks = _make_overlay(pw, ph, page_fields, client, settings)
            all_blanks.extend(blanks)
            overlay_page = pypdf.PdfReader(io.BytesIO(overlay_bytes)).pages[0]
            page.merge_page(overlay_page)
        writer.add_page(page)

    # P1-4: _REVIEW_ prefix when required fields are blank
    if all_blanks:
        output_path = output_path.parent / f"_REVIEW_{output_path.name}"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

    return all_blanks, output_path


# ── Bundle filler ─────────────────────────────────────────────────────────────

def fill_bundle(application: dict, forms_config: dict,
                client: dict, output_folder: Path,
                settings: dict, find_template_fn,
                log_path: Path | None = None,
                test_mode: bool = False) -> tuple[list[Path], list[str]]:
    """
    Fill all forms in a bundle for one client.
    Returns (output_paths, warning_messages).
    """
    today_safe  = datetime.date.today().strftime("%Y%m%d")
    name_safe   = client.get("name", "Client").replace(" ", "_").replace("/", "-")
    app_id      = application.get("id", "unknown")
    rm_staff_id = settings.get("rm_staff_id", "")
    ic          = client.get("ic_number", "")

    # Load shared fields from forms_config for _expand_shared
    shared = forms_config.get("_shared_fields", {})
    settings_with_shared = {**settings, "_shared_fields": shared}

    results, warnings = [], []

    for form_id in application.get("forms", []):
        form_cfg = forms_config.get(form_id)
        if not form_cfg:
            raise ValueError(
                f"Form '{form_id}' not in forms.json. "
                "Add it via CoordPicker or check spelling in applications.json."
            )
        template_path = find_template_fn(settings, form_cfg["template_subfolder"],
                                         form_cfg)
        prefix        = "_TEST_" if test_mode else ""
        base_name     = f"{prefix}{name_safe}_{form_id}_{today_safe}.pdf"
        output_path   = output_folder / base_name

        blanks, final_path = fill_form(
            template_path, form_cfg, client,
            settings_with_shared, output_path)
        results.append(final_path)

        if blanks:
            status = "WARN"
            msg = (f"'{client.get('name', '?')}' / {form_id}: "
                   f"blank required fields — {', '.join(blanks)}")
            warnings.append(msg)
        else:
            status = "OK"

        if log_path:
            _write_audit(log_path, {
                "timestamp":       datetime.datetime.now().isoformat(timespec="seconds"),
                "rm_staff_id":     rm_staff_id,
                "application_id":  app_id,
                "client_initials": _initials(client.get("name", "")),
                "ic_last4":        str(ic)[-4:] if ic else "",
                "form_id":         form_id,
                "output_file":     final_path.name,
                "status":          status,
                "blank_fields":    "|".join(blanks),
            })

    return results, warnings
