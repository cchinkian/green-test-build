"""
FormFiller — fully data-driven GUI.

Mode is a UI choice, not a config choice:
  Bulk   — checkbox list from batch sheet; process many clients at once
  Single — pick one client; fields auto-discovered from sheet columns;
           existing row pre-filled if client already in sheet

Adding new fields: add columns to the Excel batch sheet — they appear
in Single mode automatically, no code or config changes needed.

Adding new application: add one entry to applications.json + one sheet
in clients.xlsx — done.
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import date

sys.path.insert(0, str(Path(__file__).parent))

import config_loader
import excel_reader
import pdf_engine

# Master-sheet columns shown in bulk preview (first 3 transaction cols shown too)
_MASTER_COLS = {"name", "ic_number", "phone", "email",
                "address_line1", "address_line2", "city", "state",
                "postcode", "dob", "occupation"}

LABEL_W = 22   # label column width in single-mode form


class FormFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Form Filler")
        self.geometry("720x580")
        self.resizable(True, True)
        self.minsize(600, 460)

        # Data state
        self.settings    = {}
        self.forms       = {}
        self.applications = []
        self.app_map     = {}
        self.master_data = {}
        self.xlsx_path   = None
        self._mode       = tk.StringVar(value="bulk")
        self._bulk_vars  = []      # [(BooleanVar, merged_dict)]
        self._single_entries = {}  # field_name → StringVar

        self._build_ui()
        self._load_all()

    # ── Static UI skeleton ───────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top bar ──
        top = tk.Frame(self, pady=8)
        top.pack(fill=tk.X, padx=12)

        tk.Label(top, text="Application:", width=13, anchor="w").pack(side=tk.LEFT)
        self.var_app = tk.StringVar()
        self.cmb_app = ttk.Combobox(top, textvariable=self.var_app,
                                    state="readonly", width=36)
        self.cmb_app.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_app.bind("<<ComboboxSelected>>", self._on_app_change)

        tk.Button(top, text="↻ Reload", command=self._load_all).pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(top, text="Mode:").pack(side=tk.LEFT, padx=(0, 4))
        tk.Radiobutton(top, text="Bulk", variable=self._mode,
                       value="bulk",   command=self._on_mode_change).pack(side=tk.LEFT)
        tk.Radiobutton(top, text="Single", variable=self._mode,
                       value="single", command=self._on_mode_change).pack(side=tk.LEFT)

        tk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=(0, 4))

        # ── Dynamic middle ──
        self._panel = tk.Frame(self)
        self._panel.pack(fill=tk.BOTH, expand=True, padx=12)

        tk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=4)

        # ── Bottom bar ──
        bot = tk.Frame(self)
        bot.pack(fill=tk.X, padx=12, pady=(0, 8))
        self.btn_fill = tk.Button(bot, text="Fill & Save", width=18, height=2,
                                  bg="#28a745", fg="white",
                                  activebackground="#1e7e34",
                                  command=self._on_fill)
        self.btn_fill.pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(bot, text="Open Output Folder", width=20, height=2,
                  command=self._open_output).pack(side=tk.LEFT)

        self.lbl_status = tk.Label(self, text="  Loading...",
                                   anchor="w", fg="gray")
        self.lbl_status.pack(fill=tk.X, padx=12, pady=(0, 6))

    # ── Data loading ─────────────────────────────────────────────────────────

    def _load_all(self):
        try:
            self.settings     = config_loader.load_settings()
            self.forms        = config_loader.load_forms()
            self.applications = config_loader.load_applications()
            self.app_map      = {a["name"]: a for a in self.applications}
            self.xlsx_path    = config_loader.data_path("clients.xlsx")
            self.master_data  = excel_reader.load_master(self.xlsx_path)

            names = [a["name"] for a in self.applications]
            self.cmb_app["values"] = names
            if names:
                self.cmb_app.current(0)
            self._on_app_change()
            self._set_status(
                f"Loaded {len(self.master_data)} clients, "
                f"{len(self.applications)} applications.", "green")
        except Exception as e:
            self._set_status(f"Load error: {e}", "red")

    # ── Panel switching ───────────────────────────────────────────────────────

    def _on_app_change(self, _=None):
        self._rebuild_panel()

    def _on_mode_change(self):
        self._rebuild_panel()

    def _rebuild_panel(self):
        for w in self._panel.winfo_children():
            w.destroy()
        self._bulk_vars      = []
        self._single_entries = {}

        app = self.app_map.get(self.var_app.get(), {})
        if not app:
            return

        if self._mode.get() == "bulk":
            self._build_bulk_panel(app)
            self.btn_fill.config(text="Bulk Fill")
        else:
            self._build_single_panel(app)
            self.btn_fill.config(text="Fill & Save")

    # ── Bulk panel ────────────────────────────────────────────────────────────

    def _build_bulk_panel(self, app: dict):
        sheet = app.get("data_sheet", "Master")

        try:
            batch = excel_reader.load_batch(
                self.xlsx_path, sheet, self.master_data)
        except ValueError as e:
            tk.Label(self._panel, text=str(e), fg="red",
                     wraplength=680, justify="left").pack(anchor="w", pady=8)
            return

        # Controls row
        ctrl = tk.Frame(self._panel)
        ctrl.pack(fill=tk.X, pady=(4, 6))
        count_lbl = tk.Label(ctrl,
                             text=f"{len(batch)} client(s) in '{sheet}' sheet",
                             fg="gray")
        count_lbl.pack(side=tk.LEFT)
        tk.Button(ctrl, text="Select All",
                  command=self._bulk_select_all).pack(side=tk.RIGHT, padx=(4, 0))
        tk.Button(ctrl, text="Clear All",
                  command=self._bulk_clear_all).pack(side=tk.RIGHT)

        # Scrollable list
        outer = tk.Frame(self._panel)
        outer.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0, borderwidth=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * (e.delta / 120)), "units"))

        # One checkbox per client
        tx_cols = [k for k in (batch[0].keys() if batch else [])
                   if k not in _MASTER_COLS and k != "ic_number"]

        for client in batch:
            var = tk.BooleanVar(value=True)
            # Show name + first 4 transaction columns as summary
            summary_parts = [f"{k}: {client[k]}"
                             for k in tx_cols[:4] if client.get(k)]
            label = f"  {client.get('name', '?'):<28}  " + \
                    "   |   ".join(summary_parts)
            cb = tk.Checkbutton(inner, text=label, variable=var,
                                anchor="w", font=("Courier", 9))
            cb.pack(fill=tk.X, padx=4, pady=1)
            self._bulk_vars.append((var, client))

    def _bulk_select_all(self):
        for var, _ in self._bulk_vars:
            var.set(True)

    def _bulk_clear_all(self):
        for var, _ in self._bulk_vars:
            var.set(False)

    # ── Single panel ──────────────────────────────────────────────────────────

    def _build_single_panel(self, app: dict):
        sheet = app.get("data_sheet", "Master")

        # Client row
        row0 = tk.Frame(self._panel)
        row0.pack(fill=tk.X, pady=4)
        tk.Label(row0, text="Client:", width=LABEL_W, anchor="w").pack(side=tk.LEFT)
        self.var_client = tk.StringVar()
        names = excel_reader.get_master_names(self.master_data)
        self.cmb_client = ttk.Combobox(row0, textvariable=self.var_client,
                                       values=names, state="readonly", width=38)
        self.cmb_client.pack(side=tk.LEFT)
        if names:
            self.cmb_client.current(0)
        self.cmb_client.bind("<<ComboboxSelected>>",
                             lambda _: self._autofill_single(sheet))

        tk.Separator(self._panel, orient="horizontal").pack(
            fill=tk.X, pady=(6, 4))

        # Scrollable field form — columns from batch sheet (minus ic_number)
        try:
            headers = excel_reader.get_sheet_headers(self.xlsx_path, sheet)
        except Exception:
            headers = []

        outer = tk.Frame(self._panel)
        outer.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0, borderwidth=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        if not headers:
            tk.Label(inner,
                     text=f"No extra columns in '{sheet}' sheet.\n"
                          "All data comes from Master.",
                     fg="gray", justify="left").pack(anchor="w", pady=8, padx=4)
        else:
            tk.Label(inner,
                     text=f"Transaction fields from '{sheet}' sheet "
                          f"({len(headers)} columns). "
                          "Auto-filled if client already has a row.",
                     fg="gray", font=("", 8), justify="left"
                     ).pack(anchor="w", padx=4, pady=(2, 6))
            for col in headers:
                r = tk.Frame(inner)
                r.pack(fill=tk.X, pady=2, padx=4)
                label_text = col.replace("_", " ").title() + ":"
                tk.Label(r, text=label_text, width=LABEL_W,
                         anchor="w").pack(side=tk.LEFT)
                var = tk.StringVar()
                tk.Entry(r, textvariable=var, width=40).pack(side=tk.LEFT)
                self._single_entries[col] = var

        # Trigger autofill for default selected client
        self._autofill_single(sheet)

    def _autofill_single(self, sheet: str):
        """Pre-fill transaction fields if client already has a row in sheet."""
        ic = self._ic_for_name(self.var_client.get()
                               if hasattr(self, "var_client") else "")
        if not ic:
            return
        existing = excel_reader.find_client_in_batch(
            self.xlsx_path, sheet, ic)
        for col, var in self._single_entries.items():
            if existing.get(col):
                var.set(existing[col])

    def _ic_for_name(self, name: str) -> str:
        for ic, rec in self.master_data.items():
            if rec.get("name") == name:
                return ic
        return ""

    # ── Fill actions ──────────────────────────────────────────────────────────

    def _on_fill(self):
        app = self.app_map.get(self.var_app.get(), {})
        if not app:
            return
        if self._mode.get() == "bulk":
            self._fill_bulk(app)
        else:
            self._fill_single(app)

    def _fill_bulk(self, app: dict):
        selected = [c for var, c in self._bulk_vars if var.get()]
        if not selected:
            messagebox.showwarning("Nothing selected",
                                   "Tick at least one client.")
            return
        self._run_fill(selected, app,
                       label=f"{len(selected)} client(s)")

    def _fill_single(self, app: dict):
        name = self.var_client.get() if hasattr(self, "var_client") else ""
        ic = self._ic_for_name(name)
        if not ic:
            messagebox.showwarning("No client", "Please select a client.")
            return
        client = dict(self.master_data[ic])
        # Merge manually entered transaction fields
        for col, var in self._single_entries.items():
            v = var.get().strip()
            if v:
                client[col] = v
        self._run_fill([client], app, label=name)

    def _run_fill(self, clients: list[dict], app: dict, label: str = ""):
        fmt = self.settings.get("date_format", "%d/%m/%Y")
        today = date.today().strftime(fmt)
        for c in clients:
            c.setdefault("date", today)
            # Inject RM info from settings
            for k in ("rm_name", "rm_branch", "rm_staff_id"):
                if self.settings.get(k):
                    c.setdefault(k, self.settings[k])

        output_folder = Path(self.settings.get("output_folder", "."))
        self._set_status(f"Filling PDFs for {label}...", "blue")
        self.update_idletasks()

        saved, errors = [], []
        for client in clients:
            try:
                results = pdf_engine.fill_bundle(
                    application=app,
                    forms_config=self.forms,
                    client=client,
                    output_folder=output_folder,
                    settings=self.settings,
                    find_template_fn=config_loader.find_template,
                )
                saved.extend(results)
            except Exception as e:
                errors.append(f"{client.get('name', '?')}: {e}")

        if errors:
            messagebox.showerror("Errors", "\n".join(errors))
        msg = f"Saved {len(saved)} PDF(s)."
        if errors:
            msg += f" {len(errors)} failed."
        self._set_status(msg, "green" if not errors else "orange")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _open_output(self):
        folder = self.settings.get("output_folder", ".")
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Cannot open", str(e))

    def _set_status(self, msg: str, color: str = "gray"):
        self.lbl_status.config(text=f"  {msg}", fg=color)


if __name__ == "__main__":
    FormFillerApp().mainloop()
