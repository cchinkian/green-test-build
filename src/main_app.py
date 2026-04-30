"""
FormFiller — data-driven PDF form filler.

P1-2: Settings dialog with Browse buttons for forms_folder / output_folder.
P1-3: Startup health check — green/red form subfolder status.
P1-4: Blank required fields → _REVIEW_ filename prefix + blocking warning.
P2-5: "Execute All" single button in bulk mode.
P2-6: Auto-open output folder after fill (settings: auto_open_output).
P2-7: Session memory — restores last application + mode on launch.
P2-8: "Open CoordPicker" button launches CoordPicker.exe.
"""
import os
import sys
import platform
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import date

sys.path.insert(0, str(Path(__file__).parent))

import config_loader
import excel_reader
import pdf_engine
from excel_reader import ExcelLockedError

_MASTER_COLS = {
    "name", "ic_number", "phone", "email",
    "address_line1", "address_line2", "city",
    "state", "postcode", "dob", "occupation",
}
LABEL_W = 22


class FormFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Form Filler")
        self.geometry("740x600")
        self.resizable(True, True)
        self.minsize(620, 480)

        self.settings     = {}
        self.forms        = {}
        self.applications = []
        self.app_map      = {}
        self.master_data  = {}
        self.xlsx_path    = None
        self._mode        = tk.StringVar(value="bulk")
        self._bulk_vars   = []
        self._single_entries = {}

        self._build_ui()
        self._load_all()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Top bar
        top = tk.Frame(self, pady=6)
        top.pack(fill=tk.X, padx=12)

        tk.Label(top, text="Application:", width=13, anchor="w").pack(side=tk.LEFT)
        self.var_app = tk.StringVar()
        self.cmb_app = ttk.Combobox(top, textvariable=self.var_app,
                                    state="readonly", width=34)
        self.cmb_app.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_app.bind("<<ComboboxSelected>>", self._on_app_change)

        tk.Button(top, text="↻", width=3,
                  command=self._load_all).pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(top, text="Mode:").pack(side=tk.LEFT, padx=(0, 4))
        tk.Radiobutton(top, text="Bulk",   variable=self._mode,
                       value="bulk",   command=self._on_mode_change).pack(side=tk.LEFT)
        tk.Radiobutton(top, text="Single", variable=self._mode,
                       value="single", command=self._on_mode_change).pack(side=tk.LEFT)

        # Right-side utility buttons
        tk.Button(top, text="⚙ Settings",
                  command=self._open_settings).pack(side=tk.RIGHT, padx=2)
        tk.Button(top, text="🗂 CoordPicker",
                  command=self._open_coord_picker).pack(side=tk.RIGHT, padx=2)

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=(0, 4))

        # Health check banner (hidden until needed)
        self.frm_health = tk.Frame(self, bg="#fff3cd")
        self.lbl_health = tk.Label(self.frm_health, text="", bg="#fff3cd",
                                   anchor="w", justify="left",
                                   font=("", 8), wraplength=700)
        self.lbl_health.pack(fill=tk.X, padx=8, pady=4)

        # Dynamic panel
        self._panel = tk.Frame(self)
        self._panel.pack(fill=tk.BOTH, expand=True, padx=12)

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=4)

        # Bottom bar
        bot = tk.Frame(self)
        bot.pack(fill=tk.X, padx=12, pady=(0, 4))

        self.btn_exec_all = tk.Button(
            bot, text="▶ Execute All", width=14, height=2,
            bg="#155724", fg="white", activebackground="#0d3d18",
            command=self._execute_all)
        self.btn_exec_all.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_fill = tk.Button(
            bot, text="Fill Selected", width=14, height=2,
            bg="#28a745", fg="white", activebackground="#1e7e34",
            command=self._on_fill)
        self.btn_fill.pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(bot, text="Open Output Folder", width=18, height=2,
                  command=self._open_output).pack(side=tk.LEFT)

        self.lbl_status = tk.Label(self, text="  Loading...",
                                   anchor="w", fg="gray")
        self.lbl_status.pack(fill=tk.X, padx=12, pady=(0, 6))

    # ── Data loading ──────────────────────────────────────────────────────────

    def _load_all(self):
        try:
            self.settings     = config_loader.load_settings()
            self.forms        = config_loader.load_forms()
            self.applications = config_loader.load_applications()
            self.app_map      = {a["name"]: a for a in self.applications}
            self.xlsx_path    = config_loader.data_path("clients.xlsx")
            self.master_data  = excel_reader.load_master(self.xlsx_path)
        except ExcelLockedError as e:
            messagebox.showerror("Excel is open", str(e))
            return
        except Exception as e:
            self._set_status(f"Load error: {e}", "red")
            return

        # Populate app dropdown
        names = [a["name"] for a in self.applications]
        self.cmb_app["values"] = names

        # P2-7: restore session
        state = config_loader.load_state()
        last_app  = state.get("last_app", "")
        last_mode = state.get("last_mode", "bulk")
        self._mode.set(last_mode)
        if last_app in names:
            self.cmb_app.set(last_app)
        elif names:
            self.cmb_app.current(0)

        self._on_app_change()
        self._run_health_check()
        self._set_status(
            f"Loaded {len(self.master_data)} clients, "
            f"{len(self.applications)} applications.", "green")

    def _run_health_check(self):
        """P1-3: Show warning banner if any form subfolders are missing."""
        results = config_loader.health_check(self.settings, self.forms)
        errors  = [r for r in results if r["status"] == "error"]
        warns   = [r for r in results if r["status"] == "warn"]
        ok      = [r for r in results if r["status"] == "ok"]

        if not results:
            self.frm_health.pack_forget()
            return

        parts = []
        if ok:
            parts.append(f"✓ {len(ok)} form(s) ready")
        if warns:
            parts.append(f"⚠ {len(warns)} warning(s)")
        if errors:
            names = ", ".join(r["name"] for r in errors)
            parts.append(f"✗ {len(errors)} missing subfolder(s): {names}")

        msg = "  |  ".join(parts)
        if errors or warns:
            self.lbl_health.config(text=f"  Forms health: {msg}")
            self.frm_health.pack(fill=tk.X, padx=12, pady=(0, 4),
                                 before=self._panel)
        else:
            self.frm_health.pack_forget()

    # ── Panel switching ───────────────────────────────────────────────────────

    def _on_app_change(self, _=None):
        self._save_session()
        self._rebuild_panel()

    def _on_mode_change(self):
        self._save_session()
        self._rebuild_panel()

    def _save_session(self):
        config_loader.save_state({
            "last_app":  self.var_app.get(),
            "last_mode": self._mode.get(),
        })

    def _rebuild_panel(self):
        for w in self._panel.winfo_children():
            w.destroy()
        self._bulk_vars      = []
        self._single_entries = {}

        app = self.app_map.get(self.var_app.get(), {})
        if not app:
            return

        is_bulk = self._mode.get() == "bulk"
        self.btn_exec_all.config(state=tk.NORMAL if is_bulk else tk.DISABLED)
        self.btn_fill.config(text="Fill Selected" if is_bulk else "Fill & Save")

        if is_bulk:
            self._build_bulk_panel(app)
        else:
            self._build_single_panel(app)

    # ── Bulk panel ────────────────────────────────────────────────────────────

    def _build_bulk_panel(self, app: dict):
        sheet = app.get("data_sheet", "Master")
        try:
            batch = excel_reader.load_batch(
                self.xlsx_path, sheet, self.master_data)
        except ExcelLockedError as e:
            messagebox.showerror("Excel is open", str(e))
            return
        except ValueError as e:
            tk.Label(self._panel, text=str(e), fg="red",
                     wraplength=680, justify="left").pack(anchor="w", pady=8)
            return

        ctrl = tk.Frame(self._panel)
        ctrl.pack(fill=tk.X, pady=(4, 4))
        tk.Label(ctrl,
                 text=f"{len(batch)} client(s) in '{sheet}' sheet",
                 fg="gray").pack(side=tk.LEFT)
        tk.Button(ctrl, text="Select All",
                  command=self._bulk_select_all).pack(side=tk.RIGHT, padx=(4, 0))
        tk.Button(ctrl, text="Clear All",
                  command=self._bulk_clear_all).pack(side=tk.RIGHT)

        outer = tk.Frame(self._panel)
        outer.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
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
                            int(-1*(e.delta/120)), "units"))

        tx_cols = [k for k in (batch[0].keys() if batch else [])
                   if k not in _MASTER_COLS and k != "ic_number"]

        for client in batch:
            var = tk.BooleanVar(value=True)
            parts = [f"{k}: {client[k]}" for k in tx_cols[:4]
                     if client.get(k)]
            label = f"  {client.get('name', '?'):<28}  " + "   |   ".join(parts)
            tk.Checkbutton(inner, text=label, variable=var,
                           anchor="w", font=("Courier", 9)
                           ).pack(fill=tk.X, padx=4, pady=1)
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

        row0 = tk.Frame(self._panel)
        row0.pack(fill=tk.X, pady=4)
        tk.Label(row0, text="Client:", width=LABEL_W,
                 anchor="w").pack(side=tk.LEFT)
        self.var_client = tk.StringVar()
        names = excel_reader.get_master_names(self.master_data)
        self.cmb_client = ttk.Combobox(row0, textvariable=self.var_client,
                                       values=names, state="readonly", width=38)
        self.cmb_client.pack(side=tk.LEFT)
        if names:
            self.cmb_client.current(0)
        self.cmb_client.bind("<<ComboboxSelected>>",
                             lambda _: self._autofill_single(sheet))

        ttk.Separator(self._panel, orient="horizontal").pack(
            fill=tk.X, pady=(6, 4))

        try:
            headers = excel_reader.get_sheet_headers(self.xlsx_path, sheet)
        except ExcelLockedError:
            headers = []

        outer = tk.Frame(self._panel)
        outer.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
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
                     text=f"No transaction columns in '{sheet}' sheet.\n"
                          "All data from Master sheet.",
                     fg="gray", justify="left").pack(anchor="w", pady=8, padx=4)
        else:
            tk.Label(inner,
                     text=f"{len(headers)} transaction field(s) from '{sheet}'. "
                          "Auto-filled if client already has a row.",
                     fg="gray", font=("", 8),
                     justify="left").pack(anchor="w", padx=4, pady=(2, 6))
            for col in headers:
                r = tk.Frame(inner)
                r.pack(fill=tk.X, pady=2, padx=4)
                tk.Label(r, text=col.replace("_", " ").title() + ":",
                         width=LABEL_W, anchor="w").pack(side=tk.LEFT)
                var = tk.StringVar()
                tk.Entry(r, textvariable=var, width=40).pack(side=tk.LEFT)
                self._single_entries[col] = var

        self._autofill_single(sheet)

    def _autofill_single(self, sheet: str):
        ic = self._ic_for_name(
            self.var_client.get() if hasattr(self, "var_client") else "")
        if not ic:
            return
        try:
            existing = excel_reader.find_client_in_batch(
                self.xlsx_path, sheet, ic)
        except ExcelLockedError:
            return
        for col, var in self._single_entries.items():
            if existing.get(col) not in ("", None):
                var.set(str(existing[col]))

    def _ic_for_name(self, name: str) -> str:
        for ic, rec in self.master_data.items():
            if rec.get("name") == name:
                return ic
        return ""

    # ── Fill actions ──────────────────────────────────────────────────────────

    def _execute_all(self):
        """P2-5: Select all + fill in one click."""
        self._bulk_select_all()
        self._fill_bulk(self.app_map.get(self.var_app.get(), {}))

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
        self._run_fill(selected, app, label=f"{len(selected)} client(s)")

    def _fill_single(self, app: dict):
        name = self.var_client.get() if hasattr(self, "var_client") else ""
        ic   = self._ic_for_name(name)
        if not ic:
            messagebox.showwarning("No client", "Please select a client.")
            return
        client = dict(self.master_data[ic])
        for col, var in self._single_entries.items():
            v = var.get().strip()
            if v:
                client[col] = v
        self._run_fill([client], app, label=name)

    def _run_fill(self, clients: list[dict], app: dict, label: str = ""):
        fmt   = self.settings.get("date_format", "%d/%m/%Y")
        today = date.today().strftime(fmt)
        for c in clients:
            c.setdefault("date", today)
            for k in ("rm_name", "rm_branch", "rm_staff_id"):
                if self.settings.get(k):
                    c.setdefault(k, self.settings[k])

        output_folder = Path(self.settings.get("output_folder", "."))
        log = config_loader.log_path()

        n = len(clients)
        saved, all_warnings, errors = [], [], []

        for i, client in enumerate(clients, 1):
            self._set_status(
                f"Filling {i}/{n}: {client.get('name', '?')}...", "blue")
            self.update_idletasks()
            try:
                from config_loader import TemplateChangedWarning
                results, warnings = pdf_engine.fill_bundle(
                    application=app,
                    forms_config=self.forms,
                    client=client,
                    output_folder=output_folder,
                    settings=self.settings,
                    find_template_fn=config_loader.find_template,
                    log_path=log,
                )
                saved.extend(results)
                all_warnings.extend(warnings)
            except TemplateChangedWarning as e:
                if messagebox.askyesno(
                    "Form template changed",
                    f"{e}\n\nFill anyway with possibly wrong coordinates?",
                    icon="warning"
                ):
                    # retry without hash check
                    def _no_hash(settings, subfolder, form_cfg=None):
                        return config_loader.find_template(
                            settings, subfolder, None)
                    try:
                        results, warnings = pdf_engine.fill_bundle(
                            application=app,
                            forms_config=self.forms,
                            client=client,
                            output_folder=output_folder,
                            settings=self.settings,
                            find_template_fn=_no_hash,
                            log_path=log,
                        )
                        saved.extend(results)
                        all_warnings.extend(warnings)
                    except Exception as e2:
                        errors.append(f"{client.get('name', '?')}: {e2}")
                else:
                    errors.append(f"{client.get('name', '?')}: skipped (template changed)")
            except Exception as e:
                errors.append(f"{client.get('name', '?')}: {e}")

        # P1-4: blocking warning for blank required fields
        if all_warnings:
            review_count = sum(1 for p in saved if p.name.startswith("_REVIEW_"))
            messagebox.showwarning(
                "Required fields blank",
                f"{len(all_warnings)} form(s) had blank required fields.\n"
                f"Files prefixed with _REVIEW_ need checking before submission.\n\n"
                + "\n".join(all_warnings)
            )

        if errors:
            messagebox.showerror("Errors", "\n".join(errors))

        parts = [f"Saved {len(saved)} PDF(s)."]
        if errors:
            parts.append(f"{len(errors)} error(s).")
        if all_warnings:
            parts.append(f"Check _REVIEW_ files.")
        self._set_status("  ".join(parts),
                         "green" if not errors and not all_warnings else "orange")

        # P2-6: auto-open output folder
        if saved and self.settings.get("auto_open_output", True):
            self._open_output()

    # ── Settings dialog (P1-2) ────────────────────────────────────────────────

    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.geometry("520x360")
        dlg.resizable(False, False)
        dlg.grab_set()

        s = self.settings
        vars_ = {}

        def row(parent, label, key, default="", browse=None):
            frm = tk.Frame(parent)
            frm.pack(fill=tk.X, padx=12, pady=4)
            tk.Label(frm, text=label, width=18, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar(value=s.get(key, default))
            vars_[key] = var
            entry = tk.Entry(frm, textvariable=var, width=32)
            entry.pack(side=tk.LEFT)
            if browse:
                tk.Button(frm, text="Browse…",
                          command=lambda v=var: v.set(
                              filedialog.askdirectory() or v.get())
                          ).pack(side=tk.LEFT, padx=4)
            return var

        tk.Label(dlg, text="Paths", font=("", 10, "bold"),
                 anchor="w").pack(fill=tk.X, padx=12, pady=(10, 2))
        row(dlg, "Forms folder:", "forms_folder",
            r"C:\Forms", browse=True)
        row(dlg, "Output folder:", "output_folder",
            r"C:\Users\%USERNAME%\Documents\Filled Forms", browse=True)

        tk.Label(dlg, text="RM Profile", font=("", 10, "bold"),
                 anchor="w").pack(fill=tk.X, padx=12, pady=(10, 2))
        row(dlg, "RM Name:", "rm_name")
        row(dlg, "Branch:", "rm_branch")
        row(dlg, "Staff ID:", "rm_staff_id")

        tk.Label(dlg, text="Options", font=("", 10, "bold"),
                 anchor="w").pack(fill=tk.X, padx=12, pady=(10, 2))
        frm_opt = tk.Frame(dlg)
        frm_opt.pack(fill=tk.X, padx=12)
        auto_open = tk.BooleanVar(value=s.get("auto_open_output", True))
        tk.Checkbutton(frm_opt, text="Auto-open output folder after fill",
                       variable=auto_open).pack(anchor="w")

        def _save():
            new_settings = dict(s)
            for k, v in vars_.items():
                new_settings[k] = v.get().strip()
            new_settings["auto_open_output"] = auto_open.get()
            config_loader.save_settings(new_settings)
            self.settings = new_settings
            self._run_health_check()
            dlg.destroy()
            self._set_status("Settings saved.", "green")

        btn_frm = tk.Frame(dlg)
        btn_frm.pack(fill=tk.X, padx=12, pady=12)
        tk.Button(btn_frm, text="Save", width=10, bg="#28a745",
                  fg="white", command=_save).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_frm, text="Cancel", width=10,
                  command=dlg.destroy).pack(side=tk.LEFT)

    # ── CoordPicker launcher (P2-8) ───────────────────────────────────────────

    def _open_coord_picker(self):
        if getattr(sys, "frozen", False):
            exe = Path(sys.executable).parent / "CoordPicker.exe"
        else:
            exe = Path(__file__).parent.parent / "dist" / "CoordPicker.exe"

        if exe.exists():
            try:
                os.startfile(str(exe))
            except Exception as e:
                messagebox.showerror("Cannot open", str(e))
        else:
            messagebox.showinfo(
                "CoordPicker not found",
                f"CoordPicker.exe not found at:\n{exe}\n\n"
                "Make sure CoordPicker.exe is in the same folder as FormFiller.exe."
            )

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _open_output(self):
        folder = self.settings.get("output_folder", ".")
        try:
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                import subprocess
                subprocess.Popen(["open", folder])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            messagebox.showerror("Cannot open", str(e))

    def _set_status(self, msg: str, color: str = "gray"):
        self.lbl_status.config(text=f"  {msg}", fg=color)


if __name__ == "__main__":
    FormFillerApp().mainloop()
