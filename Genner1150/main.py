import pandas as pd
import tkinter as tk
import shutil
import os
import math
from tkinter import filedialog, messagebox
from fillpdf import fillpdfs
from datetime import datetime
import pymupdf
import fitz
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "Templates"
TEMP_OUTPUT_DIR = BASE_DIR / "Temp_Output"

def makeDataDict(filtered_df, page_capacity, from_to_values, selected_value, page_num, max_page_num):
    """
    This function makes the data dictionary for each page and returns it. This will be used to fill the pdf later
    filtered_df: Dataframe containing data from the excel that are relevant to the SHR
    page_capacity: Max amount of rows on the page
    from_to_values: Array that stores the user entered from and to values
    selected_value: String that the user selected (Issue, Turn-in, Transfer)
    page_num: The current page number that we are building the data dictionary for
    max_page_num: The max pages the end product will have
    """

    data_dict = {}

    # First page unique fields
    if page_num == 1:
        data_dict = {
            "3 TO a LOCATION b CUSTODIAN CODE": from_to_values['to'] or "",
            "4 FROM a LOCATION b CUSTODIAN CODE": from_to_values['from'] or "",
            "6 DOCUMENT NUMBER": selected_value['value'] or "",
            "2 DELIVERY DATE": datetime.today().strftime('%m/%d/%Y'), # The templates had two names, too lazy to fix the field names
            "2 DELIVERY DATE YYYYMMDD": datetime.today().strftime('%m/%d/%Y') # The templates had two names, too lazy to fix the field names
        }
        if (from_to_values['transaction_type'] == "Issue"):
            data_dict["ISSUE"] = "On"
        elif (from_to_values['transaction_type'] == "Turn-in"):
            data_dict["TURNIN"] = "On"
        else:
            data_dict["TRANSFER"] = "On"
    
    # If max page is more than 1
    if page_capacity > 16:
        page_x_of_y = f"PageXof_Y"
        data_dict[page_x_of_y] = max_page_num

    # If current page is more than 1, fill in page number
    if page_num > 1:
        page_x = f"PageX"
        data_dict[page_x] = page_num

        # Row offsets
        if max_page_num == 2:
            offset = 20
        else:
            offset = 20 + (25 * (page_num - 2))

    # Last page uniqueness
    if page_num == max_page_num:
        data_dict["10 DELIVERED BY"] = from_to_values['from']
        data_dict["11 RECEIVED BY"] = from_to_values['to']

    # Iterate through the data previously grabbed from the excel
    for i, (_, row) in enumerate(filtered_df.iterrows()):
        row_number = i + 1

        # Returns if page is at capacity
        if row_number > page_capacity:
            return data_dict

        # Field names
        if page_num > 1:
            item_num = f"Row{row_number}"
            data_dict[item_num] = row_number + offset

        asset_id = f"2 ASSET IDRow{row_number}"
        item_description = f"3 ITEM DESCRIPTION  NAMERow{row_number}"
        stock_number = f"4 STOCK NUMBERRow{row_number}"
        serial_number = f"5 SERIAL NUMBERRow{row_number}"
        manufacturer = f"6 MANUFACTURERRow{row_number}"
        model = f"7 MODELRow{row_number}"
        unit_issue = f"8 UNIT OF ISSUERow{row_number}"
        requested_QTY = f"9 REQUESTED QUANTITYRow{row_number}"
        received_QTY = f"10 RECEIVED QUANTITYRow{row_number}"
        unit_price = f"11 UNIT PRICERow{row_number}"
        total_cost = f"12 TOTAL COSTRow{row_number}"

        # Data
        val_b = str(row.iloc[2]) if not pd.isna(row.iloc[2]) else ""
        val_c = str(row.iloc[7]) if not pd.isna(row.iloc[7]) else ""
        val_d = str(row.iloc[14]) if not pd.isna(row.iloc[14]) else ""
        val_e = str(row.iloc[9]) if not pd.isna(row.iloc[9]) else ""
        val_f = str(row.iloc[13]) if not pd.isna(row.iloc[13]) else ""
        val_h = str(row.iloc[16]) if not pd.isna(row.iloc[16]) else ""
        val_i = f"{float(row.iloc[19]):.2f}" if not pd.isna(row.iloc[19]) else ""

        # Assign to dict
        data_dict[asset_id] = val_b
        data_dict[item_description] = val_c
        data_dict[stock_number] = val_h
        data_dict[serial_number] = val_d
        data_dict[manufacturer] = f"{val_e}"
        data_dict[model] = val_f
        data_dict[unit_issue] = "ea"
        data_dict[requested_QTY] = "1"
        data_dict[received_QTY] = "1"
        data_dict[unit_price] = val_i
        data_dict[total_cost] = val_i

    return data_dict

def writePDF(file_name, save_path, data_dict):
    """
    This function writes the data to the pdf
    file_name: The file name to write to
    save_path: Where to save the file after writing. This is in Temp_Output.
    data_dict: The data dictionary where field names are keys and data is value.
    """
    fillpdfs.write_fillable_pdf(
    input_pdf_path=file_name,
    output_pdf_path=save_path,
    data_dict=data_dict,
    flatten=False  # Keeps form editable
    )
    return

def combineFiles(save_path, file_names):
    """
    Combines the pdfs in Temp_output and save them to the save_path
    save_path: User defined save location
    file_names: Array of all the file names created in the Temp_Output folder
    """
    result = fitz.open()
    pymupdf.TOOLS.mupdf_display_errors(False) # Stops displaying misc errors of duplicate form fields when merging
    for pdf in file_names:
        source_path = TEMP_OUTPUT_DIR / pdf
        try:
            with fitz.open(source_path) as mfile:
                result.insert_pdf(mfile)
        except Exception as e:
            print(f"Warning: Failed to insert {pdf}: {e}")

    result.save(save_path)
    return

def deleteTempOutput():
    """
    Deleted the Temp_out folder and files then recreates it. Ready for the next run of the program.
    """
    shutil.rmtree(TEMP_OUTPUT_DIR, ignore_errors=True)
    TEMP_OUTPUT_DIR.mkdir(exist_ok=True)
    return

def main():
    # Create the root window
    root = tk.Tk()
    root.withdraw()

    # File selection
    excel_path = filedialog.askopenfilename(
        title="Select Inventory Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not excel_path:
        messagebox.showerror("Error", "No Excel file selected.")
        return

    # Read Excel and filter Column I for values containing "SHR"
    try:
        df = pd.read_excel(excel_path)
        column_j = (
            df.iloc[:, 8]
            .dropna()
            .astype(str)
            .unique()
        )
        column_j = [val for val in column_j if "SHR" in val.upper()]
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel: {e}")
        return

    if not column_j:
        messagebox.showerror("Error", "No 'SHR' values found in Column I.")
        return

    # Selection window
    selected_value = {}

    def show_selection_window():
        selection_window = tk.Toplevel(root)
        selection_window.title("Select a Value from Column I (SHR only)")

        tk.Label(selection_window, text="Choose a category containing 'SHR':").pack(pady=5)
        listbox = tk.Listbox(selection_window, width=50, height=15)
        listbox.pack()

        for val in sorted(column_j):
            listbox.insert(tk.END, val)

        def on_select():
            try:
                selected_value['value'] = listbox.get(listbox.curselection())
                selection_window.destroy()
            except:
                messagebox.showerror("Error", "Please select a value.")

        tk.Button(selection_window, text="OK", command=on_select).pack(pady=5)

    show_selection_window()
    root.wait_window(root.winfo_children()[-1])  # Wait for the selection window to close

    # Filter the dataframe
    if 'value' not in selected_value:
        messagebox.showerror("Error", "No selection made.")
        return

    filtered_df = df[df.iloc[:, 8].astype(str).str.upper() == selected_value['value'].upper()]
    if filtered_df.empty:
        messagebox.showerror("Error", "No matching inventory data found.")
        return

    # Ask for save location and filename
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Save Filled DD 1150 As"
    )
    if not save_path:
        messagebox.showerror("Error", "No save location selected.")
        return

    # Prompt for "FROM" and "TO" fields
    from_to_values = {"from": "", "to": "", "transaction_type": ""}

    def prompt_from_to():
        dialog = tk.Toplevel(root)
        dialog.title("Enter FROM, TO, and Transaction Type")

        tk.Label(dialog, text="FROM:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        from_entry = tk.Entry(dialog, width=40)
        from_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(dialog, text="TO:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        to_entry = tk.Entry(dialog, width=40)
        to_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(dialog, text="Transaction Type:").grid(row=2, column=0, padx=10, pady=10, sticky="ne")
        transaction_type_var = tk.StringVar(value="Issue")  # Default selection

        transaction_frame = tk.Frame(dialog)
        transaction_frame.grid(row=2, column=1, sticky="w", pady=5)

        options = ["Issue", "Transfer", "Turn-in"]
        for option in options:
            tk.Radiobutton(transaction_frame, text=option, variable=transaction_type_var, value=option).pack(anchor="w")

        def submit():
            from_to_values["from"] = from_entry.get()
            from_to_values["to"] = to_entry.get()
            from_to_values["transaction_type"] = transaction_type_var.get()
            dialog.destroy()

        tk.Button(dialog, text="OK", command=submit).grid(row=3, column=0, columnspan=2, pady=15)

    prompt_from_to()
    root.wait_window(root.winfo_children()[-1])

    # Display summary
    messagebox.showinfo(
        "Info",
        f"Category: {selected_value['value']}\n"
        f"Save to: {save_path}\n"
        f"FROM: {from_to_values['from']}\n"
        f"TO: {from_to_values['to']}\n"
        f"Items: {len(filtered_df)}"
    )

    # Prepare the PDFs
    # Identify which Template(s) to use and save the filenames to an array
    file_names = []
    num_items = len(filtered_df)
    if num_items <= 16:
        file_name = "Clean_1150_OnePager.pdf"
        file_names.append(file_name)

        max_pages = 1

    elif num_items <= 41:
        file_name = "Clean_1150_Pg1.pdf"
        file_names.append(file_name)

        file_name = "Clean_1150_PgLast.pdf"
        file_names.append(file_name)

        max_pages = 2
            
    else:   
        # 41 is last page and first page capacity combined, 25 is the capacity of each extra middle page
        extra_page_needed = math.ceil((num_items - 41) / 25)

        file_name = "Clean_1150_Pg1.pdf"
        file_names.append(file_name)

        for i in range (0, extra_page_needed):
            file_name = "Clean_1150_ExtraPage"
            ext = "pdf"
            file_name = f"{file_name}_{i}.{ext}"
            file_names.append(file_name)

        file_name = "Clean_1150_PgLast.pdf"
        file_names.append(file_name)

        max_pages = 2 + extra_page_needed

    # Page Capacities [Page 1 only, Page 1, Extra Pages, Last Page]
    page_capacity = [16,20,25,21]

    # Align the data to the field names
    if max_pages == 1:
        data_dict = makeDataDict(filtered_df, page_capacity[0], from_to_values, selected_value, 1, max_pages)
        source_path = TEMPLATES_DIR / file_names[0]
        destination_path = TEMP_OUTPUT_DIR / file_names[0]
        writePDF(source_path, destination_path, data_dict)
        filtered_df = filtered_df.iloc[page_capacity[0]:].reset_index(drop=True)

    elif max_pages == 2:
        data_dict = makeDataDict(filtered_df, page_capacity[1], from_to_values, selected_value, 1, max_pages)
        source_path = TEMPLATES_DIR / file_names[0]
        destination_path = TEMP_OUTPUT_DIR / file_names[0]
        writePDF(source_path, destination_path, data_dict)
        filtered_df = filtered_df.iloc[page_capacity[1]:].reset_index(drop=True)

        data_dict = makeDataDict(filtered_df, page_capacity[3], from_to_values, selected_value, 2, max_pages)
        source_path = TEMPLATES_DIR / file_names[1]
        destination_path = TEMP_OUTPUT_DIR / file_names[1]
        writePDF(source_path, destination_path, data_dict)
        filtered_df = filtered_df.iloc[page_capacity[3]:].reset_index(drop=True)
    
    else:
        data_dict = makeDataDict(filtered_df, page_capacity[1], from_to_values, selected_value, 1, max_pages)
        source_path = TEMPLATES_DIR / file_names[0]
        destination_path = TEMP_OUTPUT_DIR / file_names[0]
        writePDF(source_path, destination_path, data_dict)
        filtered_df = filtered_df.iloc[page_capacity[1]:].reset_index(drop=True)

        for i in range(1, len(file_names) - 1):
            data_dict = makeDataDict(filtered_df, page_capacity[2], from_to_values, selected_value, i+1, max_pages)
            source_path = TEMPLATES_DIR / "Clean_1150_ExtraPage.pdf"
            destination_path = TEMP_OUTPUT_DIR / file_names[i]
            writePDF(source_path, destination_path, data_dict)
            filtered_df = filtered_df.iloc[page_capacity[2]:].reset_index(drop=True)

        data_dict = makeDataDict(filtered_df, page_capacity[3], from_to_values, selected_value, len(file_names), max_pages)
        source_path = TEMPLATES_DIR / file_names[len(file_names) - 1]
        destination_path = TEMP_OUTPUT_DIR / file_names[len(file_names) - 1]
        writePDF(source_path, destination_path, data_dict)
        filtered_df = filtered_df.iloc[page_capacity[3]:].reset_index(drop=True)


    # Combine the form to Save Path and clean the Temp Output
    combineFiles(save_path, file_names)
    deleteTempOutput()


if __name__ == "__main__":
    main()