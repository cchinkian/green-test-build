"""
CoordPicker — visual tool to map PDF form field coordinates.

Open a blank form PDF, click where each field's text should appear,
fill in the field details, then save to forms.json.

Output: PDF-space coordinates (bottom-left origin).
  y=0 = bottom of page.  A4 top = y≈842.
  Use these values in forms.json field definitions.
"""
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import config_loader

try:
    import fitz          # PyMuPDF
    from PIL import Image, ImageTk
    HAS_DEPS = True
except ImportError:
    HAS_DEPS = False

ZOOM = 1.5  # display scale factor (1.5 = 150% of actual PDF size)

SOURCE_TYPES  = ["data", "settings", "fixed", "auto"]
FORMAT_OPTS   = ["", "ic_dashed", "currency_myr", "currency_no_symbol",
                 "date_dmy", "date_dmy_long", "phone_dashed",
                 "uppercase", "integer"]
AUTO_TYPES    = ["date", "year", "month"]


class CoordPickerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Coord Picker — Form Field Mapper")
        self.geometry("1050x700")
        self.resizable(True, True)
        self.minsize(800, 560)

        self._doc        = None
        self._page_num   = 1
        self._page_count = 0
        self._page_w     = 0.0
        self._page_h     = 0.0
        self._photo      = None
        self._fields: list[dict] = []

        # Form identity vars
        self._form_id       = tk.StringVar()
        self._form_name_var = tk.StringVar()
        self._subfolder     = tk.StringVar()

        # Field entry vars
        self._click_x    = tk.DoubleVar(value=0.0)
        self._click_y    = tk.DoubleVar(value=0.0)
        self._fname      = tk.StringVar()
        self._source     = tk.StringVar(value="data")
        self._fmt        = tk.StringVar()
        self._auto_type  = tk.StringVar(value="date")
        self._fixed_val  = tk.StringVar()
        self._font_size  = tk.IntVar(value=10)
        self._required   = tk.BooleanVar(value=False)

        self._build_ui()

        if not HAS_DEPS:
            messagebox.showerror(
                "Missing libraries",
                "PyMuPDF and Pillow are required.\n"
                "pip install PyMuPDF Pillow"
            )

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Left control panel
        left = tk.Frame(self, width=310, bg="#f0f0f0")
        left.pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=6)
        left.pack_propagate(False)

        def section(parent, title):
            tk.Label(parent, text=title, bg="#f0f0f0",
                     font=("", 9, "bold"), anchor="w").pack(fill=tk.X, pady=(8, 2))

        def lbl_entry(parent, text, var):
            tk.Label(parent, text=text, bg="#f0f0f0", anchor="w").pack(fill=tk.X)
            tk.Entry(parent, textvariable=var).pack(fill=tk.X, pady=(0, 3))

        # ── Form identity ──
        section(left, "1. Form identity")
        lbl_entry(left, "Form ID (no spaces):", self._form_id)
        lbl_entry(left, "Display name:", self._form_name_var)
        lbl_entry(left, "Template subfolder:", self._subfolder)
        tk.Button(left, text="Open PDF…",
                  command=self._open_pdf).pack(fill=tk.X, pady=(2, 4))

        # Page nav
        nav = tk.Frame(left, bg="#f0f0f0")
        nav.pack(fill=tk.X, pady=(0, 6))
        tk.Button(nav, text="◄", width=3,
                  command=self._prev_page).pack(side=tk.LEFT)
        self.lbl_page = tk.Label(nav, text="No PDF", bg="#f0f0f0", width=14)
        self.lbl_page.pack(side=tk.LEFT, expand=True)
        tk.Button(nav, text="►", width=3,
                  command=self._next_page).pack(side=tk.LEFT)

        ttk.Separator(left, orient="horizontal").pack(fill=tk.X, pady=4)

        # ── Coordinates display ──
        section(left, "2. Click on PDF → coordinates appear here")
        cfrm = tk.Frame(left, bg="#f0f0f0")
        cfrm.pack(fill=tk.X)
        tk.Label(cfrm, text="x:", bg="#f0f0f0").pack(side=tk.LEFT)
        tk.Label(cfrm, textvariable=self._click_x,
                 width=8, relief=tk.SUNKEN).pack(side=tk.LEFT, padx=2)
        tk.Label(cfrm, text="y:", bg="#f0f0f0").pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(cfrm, textvariable=self._click_y,
                 width=8, relief=tk.SUNKEN).pack(side=tk.LEFT, padx=2)

        ttk.Separator(left, orient="horizontal").pack(fill=tk.X, pady=6)

        # ── Field details ──
        section(left, "3. Field details")
        lbl_entry(left, "Field name (Excel column):", self._fname)

        tk.Label(left, text="Source:", bg="#f0f0f0", anchor="w").pack(fill=tk.X)
        src_cb = ttk.Combobox(left, textvariable=self._source,
                              values=SOURCE_TYPES, state="readonly")
        src_cb.pack(fill=tk.X, pady=(0, 3))
        src_cb.bind("<<ComboboxSelected>>", self._on_source_change)

        # Conditional panels
        self._pnl_fixed = tk.Frame(left, bg="#f0f0f0")
        tk.Label(self._pnl_fixed, text="Fixed value:", bg="#f0f0f0",
                 anchor="w").pack(fill=tk.X)
        tk.Entry(self._pnl_fixed,
                 textvariable=self._fixed_val).pack(fill=tk.X, pady=(0, 3))

        self._pnl_auto = tk.Frame(left, bg="#f0f0f0")
        tk.Label(self._pnl_auto, text="Auto type:", bg="#f0f0f0",
                 anchor="w").pack(fill=tk.X)
        ttk.Combobox(self._pnl_auto, textvariable=self._auto_type,
                     values=AUTO_TYPES, state="readonly").pack(fill=tk.X, pady=(0, 3))

        tk.Label(left, text="Format (optional):", bg="#f0f0f0",
                 anchor="w").pack(fill=tk.X)
        ttk.Combobox(left, textvariable=self._fmt,
                     values=FORMAT_OPTS, state="readonly").pack(fill=tk.X, pady=(0, 4))

        bot_row = tk.Frame(left, bg="#f0f0f0")
        bot_row.pack(fill=tk.X, pady=(0, 4))
        tk.Label(bot_row, text="Font size:", bg="#f0f0f0").pack(side=tk.LEFT)
        tk.Spinbox(bot_row, from_=6, to=24, textvariable=self._font_size,
                   width=4).pack(side=tk.LEFT, padx=4)
        tk.Checkbutton(bot_row, text="Required", variable=self._required,
                       bg="#f0f0f0").pack(side=tk.LEFT)

        btn_row = tk.Frame(left, bg="#f0f0f0")
        btn_row.pack(fill=tk.X, pady=4)
        tk.Button(btn_row, text="Add Field", bg="#28a745", fg="white",
                  command=self._add_field).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(btn_row, text="Undo Last",
                  command=self._undo_last).pack(side=tk.LEFT)
        self.lbl_count = tk.Label(btn_row, text="  0 fields",
                                  bg="#f0f0f0", fg="gray")
        self.lbl_count.pack(side=tk.LEFT, padx=6)

        ttk.Separator(left, orient="horizontal").pack(fill=tk.X, pady=4)

        tk.Label(left, text="Fields recorded:", bg="#f0f0f0",
                 anchor="w").pack(fill=tk.X)
        self.txt_list = tk.Text(left, height=9, font=("Courier", 7),
                                state="disabled", wrap=tk.NONE, bg="#ffffff")
        self.txt_list.pack(fill=tk.BOTH, expand=True)

        ttk.Separator(left, orient="horizontal").pack(fill=tk.X, pady=4)
        tk.Button(left, text="Save to forms.json",
                  bg="#1F4E79", fg="white", height=2,
                  command=self._save).pack(fill=tk.X)

        # Right PDF canvas
        right = tk.Frame(self)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                   padx=(0, 6), pady=6)
        self.canvas = tk.Canvas(right, bg="#888888", cursor="crosshair",
                                highlightthickness=0)
        sb_v = ttk.Scrollbar(right, orient="vertical",
                             command=self.canvas.yview)
        sb_h = ttk.Scrollbar(right, orient="horizontal",
                             command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=sb_v.set,
                              xscrollcommand=sb_h.set)
        sb_v.pack(side=tk.RIGHT, fill=tk.Y)
        sb_h.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<MouseWheel>",
                         lambda e: self.canvas.yview_scroll(
                             int(-1*(e.delta/120)), "units"))

        self.lbl_status = tk.Label(self, text="  Open a PDF to start.",
                                   anchor="w", fg="gray")
        self.lbl_status.pack(fill=tk.X, padx=8, pady=(0, 4))

    # ── PDF ───────────────────────────────────────────────────────────────────

    def _open_pdf(self):
        if not HAS_DEPS:
            return
        path = filedialog.askopenfilename(
            title="Open blank form PDF",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")]
        )
        if not path:
            return
        self._doc = fitz.open(path)
        self._page_count = len(self._doc)
        self._page_num = 1
        p = Path(path)
        if not self._subfolder.get():
            self._subfolder.set(p.parent.name)
        if not self._form_id.get():
            self._form_id.set(p.parent.name.lower().replace(" ", "_"))
        self._render()

    def _render(self):
        if not self._doc:
            return
        page = self._doc[self._page_num - 1]
        self._page_w = page.rect.width
        self._page_h = page.rect.height
        pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self._photo = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self._photo)
        self.canvas.configure(scrollregion=(0, 0, pix.width, pix.height))
        for f in self._fields:
            if f["page"] == self._page_num:
                self._dot(f["x"] * ZOOM,
                          (self._page_h - f["y"]) * ZOOM,
                          f["name"])
        self.lbl_page.config(
            text=f"Page {self._page_num} / {self._page_count}")
        self._status(f"PDF: {self._page_w:.0f}×{self._page_h:.0f} pt  |  "
                     "Click to pick coordinate")

    def _prev_page(self):
        if self._page_num > 1:
            self._page_num -= 1
            self._render()

    def _next_page(self):
        if self._doc and self._page_num < self._page_count:
            self._page_num += 1
            self._render()

    # ── Click ─────────────────────────────────────────────────────────────────

    def _on_click(self, e):
        if not self._doc:
            return
        cx = self.canvas.canvasx(e.x)
        cy = self.canvas.canvasy(e.y)
        px = round(cx / ZOOM, 1)
        py = round(self._page_h - cy / ZOOM, 1)
        self._click_x.set(px)
        self._click_y.set(py)
        self._dot(cx, cy, "·")
        self._status(f"x={px}  y={py}  (page {self._page_num})  — "
                     "enter field name and click Add Field")

    def _dot(self, cx, cy, label):
        r = 5
        self.canvas.create_oval(cx-r, cy-r, cx+r, cy+r,
                                 outline="red", width=2, tags="marker")
        self.canvas.create_text(cx+8, cy, text=label, fill="red",
                                 font=("", 7), anchor="w", tags="marker")

    # ── Field management ──────────────────────────────────────────────────────

    def _on_source_change(self, _=None):
        self._pnl_fixed.pack_forget()
        self._pnl_auto.pack_forget()
        src = self._source.get()
        if src == "fixed":
            self._pnl_fixed.pack(fill=tk.X, after=self._pnl_auto)
        elif src == "auto":
            self._pnl_auto.pack(fill=tk.X, after=self._pnl_fixed)

    def _add_field(self):
        name = self._fname.get().strip()
        if not name:
            messagebox.showwarning("Missing", "Enter a field name first.")
            return
        field: dict = {
            "name":      name,
            "page":      self._page_num,
            "x":         self._click_x.get(),
            "y":         self._click_y.get(),
            "source":    self._source.get(),
            "font_size": self._font_size.get(),
        }
        if self._fmt.get():
            field["format"] = self._fmt.get()
        if self._required.get():
            field["required"] = True
        if self._source.get() == "fixed" and self._fixed_val.get().strip():
            field["value"] = self._fixed_val.get().strip()
        if self._source.get() == "auto" and self._auto_type.get():
            field["auto_type"] = self._auto_type.get()

        self._fields.append(field)
        self._dot(self._click_x.get() * ZOOM,
                  (self._page_h - self._click_y.get()) * ZOOM,
                  name)
        self._refresh_list()
        self._fname.set("")
        self._status(f"Added '{name}'  (total: {len(self._fields)} fields)")

    def _undo_last(self):
        if not self._fields:
            return
        removed = self._fields.pop()
        self._refresh_list()
        self._render()
        self._status(f"Removed '{removed['name']}'")

    def _refresh_list(self):
        self.lbl_count.config(text=f"  {len(self._fields)} fields")
        self.txt_list.config(state="normal")
        self.txt_list.delete("1.0", tk.END)
        for f in self._fields:
            line = (f"p{f['page']}  "
                    f"x:{f['x']:<7.1f}  y:{f['y']:<7.1f}  "
                    f"{f['name']:<22} [{f['source']}]")
            if f.get("format"):
                line += f" {f['format']}"
            if f.get("required"):
                line += " *"
            self.txt_list.insert(tk.END, line + "\n")
        self.txt_list.config(state="disabled")

    # ── Save ──────────────────────────────────────────────────────────────────

    def _save(self):
        fid = self._form_id.get().strip()
        if not fid:
            messagebox.showwarning("Missing", "Enter a Form ID.")
            return
        if not self._fields:
            messagebox.showwarning("Empty", "No fields added yet.")
            return
        try:
            existing = config_loader.load_forms()
        except Exception:
            existing = {}

        # Strip template/reference entries
        clean = {k: v for k, v in existing.items()
                 if not k.startswith("_")}
        clean[fid] = {
            "name":               self._form_name_var.get().strip() or fid,
            "template_subfolder": self._subfolder.get().strip(),
            "fields":             self._fields,
        }
        config_loader.save_forms(clean)
        self._status(
            f"Saved {len(self._fields)} fields for '{fid}' → forms.json", "green")
        messagebox.showinfo(
            "Saved",
            f"'{fid}' saved to forms.json.\n"
            f"{len(self._fields)} field(s) across "
            f"{len({f['page'] for f in self._fields})} page(s)."
        )

    def _status(self, msg: str, color: str = "gray"):
        self.lbl_status.config(text=f"  {msg}", fg=color)


if __name__ == "__main__":
    CoordPickerApp().mainloop()
