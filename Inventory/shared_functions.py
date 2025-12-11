# shared_functions.py
import re
import tkinter as tk
from tkinter import ttk, filedialog
from dataclasses import dataclass
from typing import Callable, Iterable
from pathlib import Path
from openpyxl.worksheet.worksheet import Worksheet

import pandas as pd


# ---------- Data models ----------

@dataclass
class InventorySummary:
    total_in_inventory: int
    scanned: int
    missing: int
    new_items: int
    not_scanned_file: Path
    new_items_file: Path


# A small container for the standard scan UI, so all modes can reuse it
@dataclass
class ScanUI:
    window: tk.Toplevel
    container: ttk.Frame
    scan_entry: ttk.Entry
    log: Callable[[str], None]


# ---------- Generic utilities ----------

import re

def normalize(val: str, width: int = 10) -> str:
    """
    Extract digits appearing *after the final letter* in the string.
    Left-pad with zeros to the desired width.
    """
    s = str(val)

    # Find last letter (A-Z, case-insensitive)
    match = re.search(r"([A-Za-z])(?!.*[A-Za-z])", s)
    if match:
        # Take substring after the last letter
        right_side = s[match.end():]
    else:
        # No letters, use entire string
        right_side = s

    # Keep only digits
    digits = "".join(re.findall(r"\d", right_side))

    # Pad to width
    return digits.zfill(width)


def get_output_dir(base_dir: Path | None = None, name: str = "output") -> Path:
    """
    Return the directory where output files should be stored.
    Does NOT create the directory, just returns the path.
    """
    if base_dir is None:
        base_dir = Path(__file__).resolve().parent
    return base_dir / name


def ensure_root(root: tk.Misc | None = None) -> tuple[tk.Misc, bool]:
    """
    Ensure we have a Tk root. If one is created here, it's withdrawn and
    owns_root is True.
    """
    owns_root = False
    if root is None:
        root = tk.Tk()
        root.withdraw()
        owns_root = True
    return root, owns_root


def select_inventory_excel_file(
    root: tk.Misc | None = None,
    title: str = "Select Inventory Excel File",
) -> str | None:
    """
    Show a file-open dialog to choose an inventory Excel file.

    Returns:
        Selected file path as string, or None if user cancels.
    """
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")],
        parent=root,
    )
    return path or None


def load_inventory_dataframe(
    path: str | Path,
    asset_column: str = "Asset Id",
) -> pd.DataFrame:
    """
    Load the inventory Excel file into a DataFrame and verify the asset column exists.

    Raises:
        FileNotFoundError, ValueError, or any pandas-related exceptions.
    """
    path = Path(path)
    df = pd.read_excel(path)

    if asset_column not in df.columns:
        raise ValueError(f"'{asset_column}' column not found in {path.name}")

    return df


def summarize_inventory_scan(
    df_inventory: pd.DataFrame,
    scanned_codes: Iterable[str],
    new_items: Iterable[str],
    output_dir: Path,
    normalized_column: str = "Normalized",
) -> InventorySummary:
    """
    Produce summary files and basic counts for an inventory scan.

    This function is UI-agnostic and can be reused in CLI, GUI, etc.
    """
    if normalized_column not in df_inventory.columns:
        raise ValueError(f"'{normalized_column}' column not found in inventory DataFrame.")

    scanned_set = set(scanned_codes)

    not_scanned_df = df_inventory[
        ~df_inventory[normalized_column].isin(scanned_set)
    ].drop(columns=[normalized_column])

    output_dir.mkdir(parents=True, exist_ok=True)

    not_scanned_file = output_dir / "not_scanned.xlsx"
    new_items_file = output_dir / "new_item.xlsx"

    not_scanned_df.to_excel(not_scanned_file, index=False)
    pd.DataFrame({"New Items": list(new_items)}).to_excel(new_items_file, index=False)

    return InventorySummary(
        total_in_inventory=len(df_inventory),
        scanned=len(scanned_set),
        missing=len(not_scanned_df),
        new_items=len(list(new_items)),
        not_scanned_file=not_scanned_file,
        new_items_file=new_items_file,
    )


# ---------- Shared scan-window GUI ----------

def create_scan_ui(
    root: tk.Misc,
    *,
    title: str,
    instructions: str,
) -> ScanUI:
    """
    Create a standard scan window with:
      - Instructions label
      - Text log area + scrollbar
      - Entry field for the code

    Returns a ScanUI object; caller is responsible for:
      - Creating buttons under ui.container
      - Binding <Return> on ui.scan_entry
      - Handling window closing / ESC, etc.
    """
    window = tk.Toplevel(root)
    window.title(title)
    window.geometry("720x560")
    window.minsize(640, 420)

    container = ttk.Frame(window, padding=12)
    container.pack(fill="both", expand=True)

    ttk.Label(
        container,
        text=instructions,
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

    return ScanUI(
        window=window,
        container=container,
        scan_entry=scan_entry,
        log=log,
    )

def copy_cell_styles(ws: Worksheet, src_row: int, dst_row: int, cols: list[str]) -> None:
    """
    Copy font, border, fill, number_format, protection, and alignment
    from src_row to dst_row for the given column letters.
    """
    for col in cols:
        src = ws[f"{col}{src_row}"]
        dst = ws[f"{col}{dst_row}"]

        dst.font = src.font
        dst.border = src.border
        dst.fill = src.fill
        dst.number_format = src.number_format
        dst.protection = src.protection
        dst.alignment = src.alignment