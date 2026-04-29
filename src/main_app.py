"""
FormFiller — main GUI app.
Supports three application types:
  single — one client, forms filled from Master + manual fields typed in GUI
  bulk   — multiple clients from a batch sheet, checkbox selection
  bundle — one client, multiple forms, all data from Master (no manual fields)
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path
from datetime import date

sys.path.insert(0, str(Path(__file__).parent))

import config_loader
import excel_reader
import pdf_engine


class FormFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Form Filler")
        self.geometry("680x560")
        self.resizable(True, True)

        self.settings = {}
        self.forms = {}
        self.applications = []
        self.app_map = {}
        self.master_data = {}   # ic_number → dict
        self.xlsx_path = None

        self._build_header()
        self._dynamic_frame = tk.Frame(self)
        self._dynamic_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        self._build_footer()
        self._load_all()

    # ── Header (always visible) ─────────────────────────────────────────────

    def _build_header(self):
        frm = tk.Frame(self)
        frm.pack(fill=tk.X, padx=12, pady=(10, 4))

        tk.Label(frm, text="Application:", width=12, anchor="w").pack(side=tk.LEFT)
        self.var_app = tk.StringVar()
        self.cmb_app = ttk.Combobox(frm, textvariable=self.var_app, state="readonly", width=40)
        self.cmb_app.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_app.bind("<<ComboboxSelected>>", self._on_app_change)
        tk.Button(frm, text="↻", width=3, command=self._load_all).pack(side=tk.LEFT)

        tk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=4)

    # ── Footer (always visible) ─────────────────────────────────────────────

    def _build_footer(self):
        tk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=4)
        frm = tk.Frame(self)
        frm.pack(fill=tk.X, padx=12, pady=(0, 8))
        self.btn_action = tk.Button(frm, text="Fill & Save", width=18, height=2,
                                    bg="#28a745", fg="white", activebackground="#218838",
                                    command=self._on_fill)
        self.btn_action.pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(frm, text="Open Output Folder", width=18, height=2,
                  command=self._open_output).pack(side=tk.LEFT)
        self.lbl_status = tk.Label(self, text="  Loading...", anchor="w", fg="gray")
        self.lbl_status.pack(fill=tk.X, padx=12, pady=(0, 6))

    # ── Data loading ────────────────────────────────────────────────────────

    def _load_all(self):
        try:
            self.settings = config_loader.load_settings()
            self.forms = config_loader.load_forms()
            self.applications = config_loader.load_applications()
            self.app_map = {a["name"]: a for a in self.applications}

            self.xlsx_path = config_loader.data_path("clients.xlsx")
            self.master_data = excel_reader.load_master(self.xlsx_path)

            app_names = [a["name"] for a in self.applications]
            self.cmb_app["values"] = app_names
            if app_names:
                self.cmb_app.current(0)
                self._on_app_change()

            self._set_status(
                f"Loaded {len(self.master_data)} clients, {len(self.applications)} applications.", "green")
        except Exception as e:
            self._set_status(f"Load error: {e}", "red")

    # ── Dynamic panel switching ─────────────────────────────────────────────

    def _on_app_change(self, _=None):
        for w in self._dynamic_frame.winfo_children():
            w.destroy()
        app = self.app_map.get(self.var_app.get(), {})
        app_type = app.get("type", "single")

        if app_type == "bulk":
            self._build_bulk_panel(app)
            self.btn_action.config(text="Bulk Fill")
        else:  # single or bundle
            self._build_single_panel(app)
            self.btn_action.config(text="Fill & Save")

    # ── Single / Bundle panel ───────────────────────────────────────────────

    def _build_single_panel(self, app: dict):
        frm = self._dynamic_frame

        # Client row
        row = tk.Frame(frm)
        row.pack(fill=tk.X, pady=3)
        tk.Label(row, text="Client:", width=14, anchor="w").pack(side=tk.LEFT)
        self.var_client = tk.StringVar()
        names = excel_reader.get_master_names(self.master_data)
        self.cmb_client = ttk.Combobox(row, textvariable=self.var_client,
                                       values=names, state="readonly", width=38)
        self.cmb_client.pack(side=tk.LEFT)
        if names:
            self.cmb_client.current(0)
        self.cmb_client.bind("<<ComboboxSelected>>", self._refresh_single_preview)

        # Manual fields (if any)
        self._manual_entries = {}
        for field in app.get("manual_fields", []):
            r = tk.Frame(frm)
            r.pack(fill=tk.X, pady=2)
            tk.Label(r, text=field["label"] + ":", width=14, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar()
            tk.Entry(r, textvariable=var, width=40).pack(side=tk.LEFT)
            self._manual_entries[field["name"]] = var

        # Preview
        tk.Label(frm, text="Preview:", anchor="w").pack(fill=tk.X, pady=(8, 2))
        self.txt_preview = tk.Text(frm, height=10, state="disabled",
                                   bg="#f8f8f8", relief=tk.FLAT,
                                   wrap=tk.WORD, font=("Courier", 10))
        self.txt_preview.pack(fill=tk.BOTH, expand=True)
        self._refresh_single_preview()

    def _refresh_single_preview(self, _=None):
        if not hasattr(self, "txt_preview"):
            return
        app = self.app_map.get(self.var_app.get(), {})
        ic = self._selected_ic()
        client = self.master_data.get(ic, {})
        lines = [
            f"  Name:     {client.get('name', '-')}",
            f"  IC:       {client.get('ic_number', '-')}",
            f"  Phone:    {client.get('phone', '-')}",
            f"  Email:    {client.get('email', '-')}",
            f"  Address:  {client.get('address_line1', '')} {client.get('address_line2', '')}",
            f"            {client.get('city', '')} {client.get('postcode', '')} {client.get('state', '')}",
            "",
            f"  Forms:    {', '.join(app.get('forms', []))}",
            f"  Output:   {self.settings.get('output_folder', '-')}",
        ]
        self._set_preview("\n".join(lines))

    def _selected_ic(self) -> str:
        name = self.var_client.get() if hasattr(self, "var_client") else ""
        for ic, rec in self.master_data.items():
            if rec.get("name") == name:
                return ic
        return ""

    # ── Bulk panel ──────────────────────────────────────────────────────────

    def _build_bulk_panel(self, app: dict):
        frm = self._dynamic_frame
        sheet = app.get("data_sheet", "")

        try:
            batch_clients = excel_reader.get_batch_clients(
                self.xlsx_path, sheet, self.master_data)
        except Exception as e:
            tk.Label(frm, text=f"Cannot load sheet '{sheet}': {e}",
                     fg="red", wraplength=600, justify="left").pack(anchor="w")
            self._bulk_vars = []
            return

        # Select All / Clear All
        ctrl = tk.Frame(frm)
        ctrl.pack(fill=tk.X, pady=(0, 4))
        tk.Label(ctrl, text=f"Clients in '{sheet}':", width=18, anchor="w").pack(side=tk.LEFT)
        tk.Button(ctrl, text="Select All", command=self._bulk_select_all).pack(side=tk.LEFT, padx=4)
        tk.Button(ctrl, text="Clear All",  command=self._bulk_clear_all).pack(side=tk.LEFT)

        # Scrollable checkbox list
        canvas = tk.Canvas(frm, borderwidth=0, highlightthickness=0)
        sb = ttk.Scrollbar(frm, orient="vertical", command=canvas.yview)
        self._scroll_frm = tk.Frame(canvas)
        self._scroll_frm.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._scroll_frm, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self._bulk_vars = []
        self._bulk_clients = batch_clients
        for client in batch_clients:
            var = tk.BooleanVar(value=True)
            # Build summary line from non-Master columns
            extra = {k: v for k, v in client.items()
                     if k not in ("name", "ic_number", "phone", "email",
                                  "address_line1", "address_line2", "city",
                                  "state", "postcode", "dob", "occupation")
                     and v}
            summary = "  |  ".join(f"{k}: {v}" for k, v in list(extra.items())[:3])
            label = f"{client.get('name', '?')}   {summary}"
            tk.Checkbutton(self._scroll_frm, text=label, variable=var,
                           anchor="w").pack(fill=tk.X, padx=4, pady=1)
            self._bulk_vars.append((var, client))

    def _bulk_select_all(self):
        for var, _ in getattr(self, "_bulk_vars", []):
            var.set(True)

    def _bulk_clear_all(self):
        for var, _ in getattr(self, "_bulk_vars", []):
            var.set(False)

    # ── Fill actions ────────────────────────────────────────────────────────

    def _on_fill(self):
        app = self.app_map.get(self.var_app.get(), {})
        app_type = app.get("type", "single")
        if app_type == "bulk":
            self._fill_bulk(app)
        else:
            self._fill_single(app)

    def _fill_single(self, app: dict):
        ic = self._selected_ic()
        if not ic:
            messagebox.showwarning("No client", "Please select a client.")
            return
        client = dict(self.master_data[ic])
        # Merge manual fields
        for fname, var in self._manual_entries.items():
            client[fname] = var.get().strip()
        client.setdefault("date", date.today().strftime(
            self.settings.get("date_format", "%d/%m/%Y")))
        self._run_fill([client], app)

    def _fill_bulk(self, app: dict):
        selected = [c for var, c in getattr(self, "_bulk_vars", []) if var.get()]
        if not selected:
            messagebox.showwarning("Nothing selected", "Tick at least one client.")
            return
        fmt = self.settings.get("date_format", "%d/%m/%Y")
        for c in selected:
            c.setdefault("date", date.today().strftime(fmt))
        self._run_fill(selected, app)

    def _run_fill(self, clients: list[dict], app: dict):
        output_folder = Path(self.settings.get("output_folder", "."))
        self._set_status(f"Filling {len(clients)} client(s)...", "blue")
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
            messagebox.showerror("Some errors", "\n".join(errors))
        if saved:
            self._set_status(f"Saved {len(saved)} PDF(s) to output folder.", "green")
        else:
            self._set_status("No PDFs saved.", "red")

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _set_preview(self, text: str):
        if hasattr(self, "txt_preview"):
            self.txt_preview.config(state="normal")
            self.txt_preview.delete("1.0", tk.END)
            self.txt_preview.insert(tk.END, text)
            self.txt_preview.config(state="disabled")

    def _open_output(self):
        folder = self.settings.get("output_folder", ".")
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Cannot open folder", str(e))

    def _set_status(self, msg: str, color: str = "gray"):
        self.lbl_status.config(text=f"  {msg}", fg=color)


if __name__ == "__main__":
    app = FormFillerApp()
    app.mainloop()
