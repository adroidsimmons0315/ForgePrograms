import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from pathlib import Path

def normalize(val: str, width: int = 6) -> str:
    """
    Keep only digits from val and left-pad with zeros to fixed width.
    """
    s = str(val)
    digits = "".join(re.findall(r"\d", s))
    return digits.zfill(width)

def run_initial_inventory(root: tk.Misc | None = None):
    """
    If root is None, creates its own Tk root and mainloop.
    If root is provided, opens a Toplevel attached to it.
    """
    base_dir = Path(__file__).resolve().parent
    output_dir = base_dir / "output"

    # Create root FIRST so dialogs have a parent
    owns_root = False
    if root is None:
        root = tk.Tk()
        root.withdraw()
        owns_root = True

    # Select Excel file
    data_path = filedialog.askopenfilename(
        title="Select Inventory Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")],
        parent=root
    )
    if not data_path:
        if owns_root:
            root.destroy()
        return

    # Load data
    try:
        df_inventory = pd.read_excel(data_path)
    except Exception as e:
        messagebox.showerror("Load error", f"Failed to read Excel file:\n{e}", parent=root)
        if owns_root:
            root.destroy()
        return

    asset_column = "Asset Id"
    if asset_column not in df_inventory.columns:
        messagebox.showerror(
            "Column missing",
            f"'{asset_column}' column not found in {Path(data_path).name}",
            parent=root,
        )
        if owns_root:
            root.destroy()
        return

    # Normalize identifiers in inventory (digits-only, pad to 6)
    df_inventory = df_inventory.copy()
    df_inventory["Normalized"] = df_inventory[asset_column].map(normalize)
    normalized_asset_set = set(df_inventory["Normalized"])

    scanned = set()
    new_items = []

    # UI window
    scan_win = tk.Toplevel(root)
    scan_win.title("Initial Inventory Scan")
    scan_win.geometry("720x560")
    scan_win.minsize(640, 420)

    def close_window():
        if owns_root:
            root.destroy()
        else:
            scan_win.destroy()

    scan_win.protocol("WM_DELETE_WINDOW", close_window)

    container = ttk.Frame(scan_win, padding=12)
    container.pack(fill="both", expand=True)

    ttk.Label(
        container,
        text="Scan or type an asset code, then press Enter. Type 'done' to finish.",
        font=("Segoe UI", 11),
    ).pack(anchor="w", pady=(0, 8))

    # Text area + scrollbar
    wrap = ttk.Frame(container)
    wrap.pack(fill="both", expand=True)
    yscroll = ttk.Scrollbar(wrap, orient="vertical")
    output_text = tk.Text(wrap, width=90, height=22, yscrollcommand=yscroll.set)
    yscroll.config(command=output_text.yview)
    output_text.pack(side="left", fill="both", expand=True)
    yscroll.pack(side="right", fill="y")

    def log(msg: str):
        output_text.insert(tk.END, msg + "\n")
        output_text.see(tk.END)

    # Entry row
    row = ttk.Frame(container)
    row.pack(fill="x", pady=8)
    ttk.Label(row, text="Code:").pack(side="left", padx=(0, 8))
    scan_entry = ttk.Entry(row, width=32, font=("Segoe UI", 14))
    scan_entry.pack(side="left")
    scan_entry.focus()

    # Buttons
    btns = ttk.Frame(container)
    btns.pack(fill="x", pady=(4, 0))

    def handle_scan(event=None):
        code = scan_entry.get().strip()
        scan_entry.delete(0, tk.END)
        if not code:
            return

        if code.lower() == "done":
            summarize()
            return

        normalized_code = normalize(code)

        if normalized_code in normalized_asset_set:
            if normalized_code in scanned:
                log(f"[✓] {code} already scanned.")
            else:
                scanned.add(normalized_code)
                log(f"[✓] {code} found in inventory.")
        else:
            log(f"[X] {code} NOT found in inventory.")
            new_items.append(code)

    def summarize():
        not_scanned_df = df_inventory[~df_inventory["Normalized"].isin(scanned)].drop(
            columns=["Normalized"]
        )
        output_dir.mkdir(parents=True, exist_ok=True)

        not_scanned_file = output_dir / "not_scanned.xlsx"
        new_items_file = output_dir / "new_item.xlsx"

        try:
            not_scanned_df.to_excel(not_scanned_file, index=False)
            pd.DataFrame({"New Items": new_items}).to_excel(new_items_file, index=False)
        except Exception as e:
            messagebox.showerror("Export error", f"Failed to write output files:\n{e}", parent=scan_win)
            return

        log("\n--- Scan Summary ---")
        log(f"Total in inventory: {len(df_inventory)}")
        log(f"Scanned: {len(scanned)}")
        log(f"Missing: {len(not_scanned_df)}")
        log(f"New items: {len(new_items)}")
        log("\nExported to:")
        log(f" - {not_scanned_file}")
        log(f" - {new_items_file}")

        messagebox.showinfo(
            "Export complete", f"Saved:\n{not_scanned_file}\n{new_items_file}", parent=scan_win
        )
        close_window()

    ttk.Button(btns, text="Finish (Done)", command=summarize).pack(side="left")
    ttk.Button(btns, text="Clear Entry", command=lambda: scan_entry.delete(0, tk.END)).pack(
        side="left", padx=8
    )
    ttk.Button(btns, text="Quit All", command=close_window).pack(side="right")

    scan_entry.bind("<Return>", handle_scan)
    scan_win.bind("<Escape>", lambda e: close_window())

    # Start loop in standalone mode
    if owns_root:
        root.deiconify()
        root.mainloop()

if __name__ == "__main__":
    run_initial_inventory()