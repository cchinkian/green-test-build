"""
Green Test App - Environment Verification
Double-click to test if .exe can run on target PC without admin.
Tests: GUI, file read, file write, PDF detection.
"""
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import platform
import os
import sys

def get_app_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent

def test_read_folder():
    folder = filedialog.askdirectory(title="Select a folder with PDF forms")
    if not folder:
        return

    folder_path = Path(folder)
    pdfs = list(folder_path.rglob("*.pdf"))
    subfolders = [f.name for f in folder_path.iterdir() if f.is_dir()]

    result = f"Folder: {folder}\n"
    result += f"Subfolders: {len(subfolders)}\n"
    for sf in subfolders[:10]:
        result += f"  - {sf}\n"
    result += f"\nPDFs found: {len(pdfs)}\n"
    for pdf in pdfs[:15]:
        size_kb = pdf.stat().st_size / 1024
        result += f"  - {pdf.relative_to(folder_path)} ({size_kb:.0f} KB)\n"
    if len(pdfs) > 15:
        result += f"  ... and {len(pdfs)-15} more\n"

    txt_result.delete("1.0", tk.END)
    txt_result.insert(tk.END, result)
    lbl_read_status.config(text="READ: OK", fg="green")

def test_write_file():
    app_dir = get_app_dir()
    test_file = app_dir / "zzz_write_test.txt"
    try:
        test_file.write_text("Write test OK\n")
        content = test_file.read_text()
        test_file.unlink()
        lbl_write_status.config(text=f"WRITE: OK ({app_dir})", fg="green")
    except PermissionError:
        lbl_write_status.config(text=f"WRITE: BLOCKED ({app_dir})", fg="red")
    except Exception as e:
        lbl_write_status.config(text=f"WRITE: ERROR - {e}", fg="red")

def test_pdf_write():
    folder = filedialog.askdirectory(title="Select output folder (e.g. pen drive)")
    if not folder:
        return
    out = Path(folder) / "zzz_pdf_test.txt"
    try:
        out.write_text("PDF output test OK\n")
        out.unlink()
        lbl_pdf_status.config(text=f"PDF OUTPUT: OK ({folder})", fg="green")
    except Exception as e:
        lbl_pdf_status.config(text=f"PDF OUTPUT: FAILED - {e}", fg="red")

# --- GUI ---
root = tk.Tk()
root.title("Green Test - Form Filler Environment Check")
root.geometry("700x520")
root.resizable(True, True)

# System info
info = f"OS: {platform.system()} {platform.release()} | Python: {platform.python_version()} | Admin: {os.name}"
if platform.system() == "Windows":
    import ctypes
    is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    info += f" | Running as Admin: {is_admin}"
info += f"\nApp location: {get_app_dir()}"

lbl_info = tk.Label(root, text=info, justify=tk.LEFT, anchor="w", wraplength=680)
lbl_info.pack(fill=tk.X, padx=10, pady=(10,5))

# Test buttons frame
frm = tk.Frame(root)
frm.pack(fill=tk.X, padx=10, pady=5)

btn_read = tk.Button(frm, text="1. Test Read Folder", command=test_read_folder, width=20, height=2)
btn_read.grid(row=0, column=0, padx=5, pady=5)
lbl_read_status = tk.Label(frm, text="READ: not tested", fg="gray")
lbl_read_status.grid(row=0, column=1, sticky="w", padx=10)

btn_write = tk.Button(frm, text="2. Test Write (App Dir)", command=test_write_file, width=20, height=2)
btn_write.grid(row=1, column=0, padx=5, pady=5)
lbl_write_status = tk.Label(frm, text="WRITE: not tested", fg="gray")
lbl_write_status.grid(row=1, column=1, sticky="w", padx=10)

btn_pdf = tk.Button(frm, text="3. Test Output Folder", command=test_pdf_write, width=20, height=2)
btn_pdf.grid(row=2, column=0, padx=5, pady=5)
lbl_pdf_status = tk.Label(frm, text="PDF OUTPUT: not tested", fg="gray")
lbl_pdf_status.grid(row=2, column=1, sticky="w", padx=10)

# Results area
tk.Label(root, text="Results:", anchor="w").pack(fill=tk.X, padx=10, pady=(10,0))
txt_result = tk.Text(root, height=15, wrap=tk.WORD)
txt_result.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))

root.mainloop()
