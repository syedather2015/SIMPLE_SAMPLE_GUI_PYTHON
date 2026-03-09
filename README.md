import math
import pandas as pd
from tkinter import Tk, filedialog, simpledialog, messagebox

# ======================================================
#     SIMPLE GUI SYSTEMATIC SAMPLER (MACRO-EQUIVALENT)
# ======================================================

IDEAL = 730      # Main target
MAX_CAP = 800    # Maximum allowed sample size

# Hide main Tk window
root = Tk()
root.withdraw()

messagebox.showinfo("Systematic Sampler", 
                    "Select the input Excel/CSV file for sampling.")

# ---- FILE PICKER ----
file_path = filedialog.askopenfilename(
    title="Select your data file",
    filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
)

if not file_path:
    messagebox.showerror("Error", "No file selected. Exiting.")
    exit()

# ---- Ask for sheet name (Excel only) ----
sheet_name = None
read_first_sheet = False
if file_path.lower().endswith((".xlsx", ".xls")):
    sheet_name_input = simpledialog.askstring(
        "Sheet Name",
        "Enter sheet name (leave blank to use the FIRST sheet):"
    )
    if sheet_name_input and sheet_name_input.strip():
        sheet_name = sheet_name_input.strip()
    else:
        read_first_sheet = True

# ---- Ask for RPG column name (kept for compatibility) ----
_ = simpledialog.askstring(
    "RPG Column",
    "Enter the RPG column name (as in header):\n(Used only for compatibility; not needed for sampling)"
)

# ---- Load File ----
try:
    if file_path.lower().endswith(".csv"):
        df = pd.read_csv(file_path)
    else:
        if read_first_sheet:
            df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

        if isinstance(df, dict):
            first_sheet = list(df.keys())[0]
            df = df[first_sheet]
except Exception as e:
    messagebox.showerror("Error Loading File", str(e))
    exit()

if not isinstance(df, pd.DataFrame):
    messagebox.showerror("Error", "Loaded object is not a table. Please check the file.")
    exit()

n = len(df)
if n == 0:
    messagebox.showerror("Error", "The file contains no data rows!")
    exit()

# ---- Step Calculation ----
def choose_step(total_rows, target):
    return max(1, math.ceil(total_rows / target))

step_ideal = choose_step(n, IDEAL)
count_if_ideal = math.ceil(n / step_ideal)
step = step_ideal if count_if_ideal <= MAX_CAP else choose_step(n, MAX_CAP)

# ---- Select indices (systematic sample) ----
indices = list(range(0, n, step))[:MAX_CAP]
sample = df.iloc[indices].reset_index(drop=True)

# ---- Save Output ----
output_path = filedialog.asksaveasfilename(
    title="Save Output File As",
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)

if not output_path:
    messagebox.showerror("Error", "Output file not chosen. Exiting.")
    exit()

try:
    sample.to_excel(output_path, index=False)
except Exception as e:
    messagebox.showerror("Error Saving File", str(e))
    exit()

messagebox.showinfo(
    "Success",
    f"Sampling complete!\n\n"
    f"Total rows in source: {n}\n"
    f"Sampled rows: {len(sample)}\n"
    f"Step size used: {step}\n\n"
    f"Saved to:\n{output_path}"
)
