"""
FormFiller — main GUI app.
Select client from Excel, pick a form bundle, fill & save PDFs to output folder.
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

# Allow running directly from src/ during dev
sys.path.insert(0, str(Path(__file__).parent))

import config_loader
import excel_reader
import pdf_engine


class FormFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Form Filler")
        self.geometry("620x480")
        self.resizable(True, True)

        self.clients = []
        self.client_map = {}   # name → dict
        self.forms = {}
        self.applications = []
        self.app_map = {}      # name → dict
        self.settings = {}

        self._build_ui()
        self._load_all()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # --- Client row ---
        frm_client = tk.Frame(self)
        frm_client.pack(fill=tk.X, **pad)
        tk.Label(frm_client, text="Client:", width=10, anchor="w").pack(side=tk.LEFT)
        self.var_client = tk.StringVar()
        self.cmb_client = ttk.Combobox(frm_client, textvariable=self.var_client, state="readonly", width=38)
        self.cmb_client.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_client.bind("<<ComboboxSelected>>", self._on_client_change)
        tk.Button(frm_client, text="↻", width=3, command=self._load_all).pack(side=tk.LEFT)

        # --- Bundle row ---
        frm_bundle = tk.Frame(self)
        frm_bundle.pack(fill=tk.X, **pad)
        tk.Label(frm_bundle, text="Bundle:", width=10, anchor="w").pack(side=tk.LEFT)
        self.var_bundle = tk.StringVar()
        self.cmb_bundle = ttk.Combobox(frm_bundle, textvariable=self.var_bundle, state="readonly", width=38)
        self.cmb_bundle.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_bundle.bind("<<ComboboxSelected>>", self._on_bundle_change)

        tk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=4)

        # --- Preview ---
        tk.Label(self, text="Preview:", anchor="w").pack(fill=tk.X, padx=12)
        self.txt_preview = tk.Text(self, height=12, state="disabled", bg="#f8f8f8",
                                   relief=tk.FLAT, wrap=tk.WORD, font=("Courier", 10))
        self.txt_preview.pack(fill=tk.BOTH, expand=True, padx=12, pady=(2, 4))

        tk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=4)

        # --- Buttons ---
        frm_btns = tk.Frame(self)
        frm_btns.pack(fill=tk.X, padx=12, pady=(0, 4))
        self.btn_fill = tk.Button(frm_btns, text="Fill & Save", width=16, height=2,
                                  bg="#28a745", fg="white", activebackground="#218838",
                                  command=self._fill_and_save)
        self.btn_fill.pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(frm_btns, text="Open Output Folder", width=18, height=2,
                  command=self._open_output).pack(side=tk.LEFT)

        # --- Status ---
        self.lbl_status = tk.Label(self, text="Loading...", anchor="w", fg="gray")
        self.lbl_status.pack(fill=tk.X, padx=12, pady=(0, 8))

    def _load_all(self):
        try:
            self.settings = config_loader.load_settings()
            self.forms = config_loader.load_forms()
            self.applications = config_loader.load_applications()
            self.app_map = {a["name"]: a for a in self.applications}

            xlsx_path = config_loader.data_path("clients.xlsx")
            self.clients = excel_reader.load_clients(xlsx_path)
            self.client_map = {c["name"]: c for c in self.clients}

            names = excel_reader.get_client_names(self.clients)
            self.cmb_client["values"] = names
            if names:
                self.cmb_client.current(0)

            app_names = [a["name"] for a in self.applications]
            self.cmb_bundle["values"] = app_names
            if app_names:
                self.cmb_bundle.current(0)

            self._refresh_preview()
            self._set_status(f"Loaded {len(self.clients)} clients, {len(self.applications)} bundles.", "green")
        except Exception as e:
            self._set_status(f"Load error: {e}", "red")

    def _on_client_change(self, _=None):
        self._refresh_preview()

    def _on_bundle_change(self, _=None):
        self._refresh_preview()

    def _refresh_preview(self):
        client_name = self.var_client.get()
        bundle_name = self.var_bundle.get()
        client = self.client_map.get(client_name, {})
        app = self.app_map.get(bundle_name, {})

        lines = []
        lines.append(f"  Name:     {client.get('name', '-')}")
        lines.append(f"  IC:       {client.get('ic_number', '-')}")
        lines.append(f"  Phone:    {client.get('phone', '-')}")
        lines.append(f"  Email:    {client.get('email', '-')}")
        lines.append(f"  Address:  {client.get('address_line1', '')} {client.get('address_line2', '')}")
        lines.append(f"            {client.get('city', '')} {client.get('postcode', '')} {client.get('state', '')}")
        lines.append("")
        form_ids = app.get("forms", [])
        lines.append(f"  Bundle:   {bundle_name}")
        lines.append(f"  Forms:    {', '.join(form_ids) if form_ids else '-'}")
        lines.append(f"  Output:   {self.settings.get('output_folder', '-')}")

        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert(tk.END, "\n".join(lines))
        self.txt_preview.config(state="disabled")

    def _fill_and_save(self):
        client_name = self.var_client.get()
        bundle_name = self.var_bundle.get()
        if not client_name:
            messagebox.showwarning("No client", "Please select a client.")
            return
        if not bundle_name:
            messagebox.showwarning("No bundle", "Please select a form bundle.")
            return

        client = self.client_map.get(client_name, {})
        app = self.app_map.get(bundle_name, {})
        output_folder = Path(self.settings.get("output_folder", "."))

        self._set_status("Filling PDFs...", "blue")
        self.update_idletasks()
        try:
            results = pdf_engine.fill_bundle(
                application=app,
                forms_config=self.forms,
                client=client,
                output_folder=output_folder,
                settings=self.settings,
                find_template_fn=config_loader.find_template,
            )
            names = [p.name for p in results]
            self._set_status(f"Saved {len(results)} PDF(s): {', '.join(names)}", "green")
        except Exception as e:
            self._set_status(f"Error: {e}", "red")
            messagebox.showerror("Fill failed", str(e))

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
