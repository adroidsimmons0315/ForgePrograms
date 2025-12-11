import subprocess, sys, pathlib, tkinter as tk
from tkinter import ttk, messagebox

ROOT = pathlib.Path(__file__).resolve().parents[0]
APPS = {
    "Inventory": ROOT / "Inventory" / "inventory.py",
    "1150 Genner": ROOT / "Genner1150" / "main.py",
    "Scan Out": ROOT / "Inventory" / "scan_out.py",
    "Scan In": ROOT / "Inventory" / "scan_in.py",
    "Import new HR" : ROOT / "Inventory" / "import_hr.py"
}

def run_app(path: pathlib.Path):
    if not path.exists():
        messagebox.showerror("Error", f"Not found: {path}")
        return
    # launch as a separate process, with its folder as CWD so outputs go there
    subprocess.Popen([sys.executable, str(path)], cwd=str(path.parent))

def main():
    root = tk.Tk()
    root.title("Toolbox Launcher")
    root.geometry("360x360")

    frm = ttk.Frame(root, padding=16)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Choose a program to run:", font=("Segoe UI", 12, "bold")).pack(pady=(0,12))

    for name, path in APPS.items():
        ttk.Button(frm, text=name, command=lambda p=path: run_app(p)).pack(fill="x", pady=6)

    ttk.Separator(frm).pack(fill="x", pady=12)
    ttk.Button(frm, text="Quit", command=root.destroy).pack()

    root.mainloop()

if __name__ == "__main__":
    main()