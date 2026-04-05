import pandas as pd
import os
from glob import glob
import tkinter as tk
from tkinter import filedialog, messagebox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DATA_DIR = os.path.join(BASE_DIR, "sample_data")
DEFAULT_INPUT_DIR = SAMPLE_DATA_DIR
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Batch capacities in order
BATCH_SIZES = [30, 30, 30, 20, 20, 20, 20, 10, 10, 10]

# Hide Tkinter main window
root = tk.Tk()
root.withdraw()

# Select input Excel
input_file = filedialog.askopenfilename(
    title="Select the input Excel file",
    initialdir=DEFAULT_INPUT_DIR,
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not input_file:
    messagebox.showinfo("No file selected", "Exiting...")
    exit()

# Select output folder
output_folder = filedialog.askdirectory(
    title="Select folder to save compiled Excel files",
    initialdir=DEFAULT_OUTPUT_DIR
)

# If user cancels
if not output_folder:
    messagebox.showinfo("No folder selected", "Exiting...")
    exit()

# Ensure folder exists
os.makedirs(output_folder, exist_ok=True)

folders_to_search = {
    "Profile 1": os.path.join(SAMPLE_DATA_DIR, "Profile 1", "Parts", "By Name"),
    "Profile 2": os.path.join(SAMPLE_DATA_DIR, "Profile 2", "Parts", "By Name"),
    "Profile 3": os.path.join(SAMPLE_DATA_DIR, "Profile 3", "Parts", "By Name"),
    "Profile 4": os.path.join(SAMPLE_DATA_DIR, "Profile 4", "Parts", "By Name"),
}

df_input = pd.read_excel(input_file)
input_base = os.path.splitext(os.path.basename(input_file))[0]

required_cols = {"Filename", "Quantity"}
if not required_cols.issubset(df_input.columns):
    raise ValueError("Input Excel must contain 'Filename' and 'Quantity' columns.")


# ----------- HELPER FUNCTIONS -----------

def find_file_and_profile(filename_base):
    """
    Finds the correct file in profile folders.
    Matches filename_base exactly at the start, but ensures the next character
    is either a space, dash, or the end of the filename to avoid XF1 matching XF11.
    """
    filename_base = filename_base.strip().lower()  # normalize input
    for profile, folder in folders_to_search.items():
        for f in glob(os.path.join(folder, "*.xlsx")):
            base = os.path.splitext(os.path.basename(f))[0].strip().lower()
            # match if filename_base is exactly at start, but not followed by a number
            if base == filename_base:
                return f, profile
            if base.startswith(filename_base):
                next_char = base[len(filename_base):len(filename_base)+1]
                if next_char in ["", " ", "-"]:  # allow XF1, XF1 - 7 holes
                    return f, profile
    print(f"DEBUG: Could not find file for '{filename_base}' in any profile folder")
    return None, None

def clean_part_dataframe(df):
    # Keep only the columns we need: 'Hole length (mm)' + 'Hole 1' → 'Hole 5'
    expected_cols = ["Hole length (mm)", "Hole 1", "Hole 2", "Hole 3", "Hole 4", "Hole 5"]
    df_clean = df.loc[:, df.columns.intersection(expected_cols)].copy()

    # Convert 'Hole length (mm)' to numeric, coerce errors to NaN
    df_clean["Hole length (mm)"] = pd.to_numeric(df_clean["Hole length (mm)"], errors='coerce')

    # Drop rows where 'Hole length (mm)' is NaN
    df_clean = df_clean.dropna(subset=["Hole length (mm)"])

    # Fill remaining NaN in Hole columns with empty string
    hole_cols = [col for col in df_clean.columns if "Hole" in col]
    df_clean[hole_cols] = df_clean[hole_cols].fillna("")

    return df_clean

# Store parts grouped by profile
compiled_profiles = {profile: [] for profile in folders_to_search.keys()}
skipped_parts = []

# ----------- PROCESS INPUT LIST -----------

for _, row in df_input.iterrows():
    filename_full = row["Filename"]
    quantity = row["Quantity"]

    base_name = os.path.splitext(os.path.basename(filename_full))[0]
    matched_file, profile = find_file_and_profile(base_name)

    if matched_file:
        print(f"\n=== FOUND PART: {base_name} in {profile} ===")

        try:
            df_part = pd.read_excel(matched_file, sheet_name="Sheet1")
            df_clean = clean_part_dataframe(df_part)

            print(f"Rows after cleaning: {len(df_clean)}")
            compiled_profiles[profile].append((base_name, quantity, df_clean))
        except Exception as e:
            print(f"\n### SKIPPED - failed to read {base_name}: {e}")
            skipped_parts.append({
                "Filename": base_name,
                "Quantity": quantity,
                "Reason": f"Failed to read Sheet1: {e}"
            })
    else:
        print(f"\n### SKIPPED - file not found: {base_name}")
        skipped_parts.append({
            "Filename": base_name,
            "Quantity": quantity,
            "Reason": "File not found in any profile folder"
        })


def pack_parts_one_per_slot(parts):
    """
    Packs parts into batches with one part per slot.
    Each batch has slot capacities defined by BATCH_SIZES.
    Largest parts are placed first.
    If a part cannot fit in the current batch, move to the next batch.
    """
    # Sort parts by descending row count
    parts_sorted = sorted(parts, key=lambda x: len(x[2]), reverse=True)
    
    batches = []  # list of batches, each batch = list of slots
    leftovers = []  # parts that couldn't fit in any batch

    for part_name, qty, df_part in parts_sorted:
        rows = len(df_part)
        placed = False

        # Try to place in existing batches
        for batch in batches:
            for i, slot in enumerate(batch):
                slot_capacity = BATCH_SIZES[i]
                if slot is None and rows <= slot_capacity:
                    batch[i] = (part_name, qty, df_part)
                    placed = True
                    print(f"✔ Placed {part_name} ({rows} rows) into slot #{i+1} of batch #{batches.index(batch)+1}")
                    break
            if placed:
                break

        # If not placed, create a new batch
        if not placed:
            batch = [None]*len(BATCH_SIZES)
            for i, cap in enumerate(BATCH_SIZES):
                if rows <= cap:
                    batch[i] = (part_name, qty, df_part)
                    placed = True
                    print(f"✔ Placed {part_name} ({rows} rows) into slot #{i+1} of new batch #{len(batches)+1}")
                    break
            batches.append(batch)
            if not placed:
                leftovers.append((part_name, qty, df_part))
                print(f"❌ Could NOT fit {part_name} ({rows} rows) in a new batch either")

    # Print final status
    for b_idx, batch in enumerate(batches, start=1):
        print(f"\n==== FINAL SLOT STATUS: Batch {b_idx} ====")
        for i, slot in enumerate(batch, start=1):
            if slot:
                print(f"Slot #{i}: {slot[0]} ({len(slot[2])} rows)")
            else:
                print(f"Slot #{i}: empty (Capacity: {BATCH_SIZES[i-1]})")

    if leftovers:
        print(f"\n⚠ {len(leftovers)} parts could not fit in any batch:")
        for p in leftovers:
            print(f" - {p[0]} ({len(p[2])} rows)")

    return batches, leftovers


# ----------- WRITE OUTPUT EXCEL FOR MULTI-BATCH ONE-PART-PER-SLOT (Skip empty slots) -----------

for profile, parts in compiled_profiles.items():
    if not parts:
        continue

    print(f"\n==== PACKING PROFILE: {profile} ====")
    batches, leftovers = pack_parts_one_per_slot(parts)

    save_path = os.path.join(output_folder, f"{input_base} {profile}.xlsx")
    
    # Make sure this 'with' block is inside the loop
    with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
        workbook = writer.book

        for b_idx, batch in enumerate(batches, start=1):
            sheet_name = f"{profile} Batch {b_idx}"
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            row_pos = 0
            
            # Write batch header
            worksheet.write(row_pos, 0, f"Batch {b_idx}")
            row_pos += 1

            # Write input file name under batch header
            worksheet.write(row_pos, 0, f"Input File: {input_base}")
            row_pos += 2  # spacing before slot info

            # Write profile name under input file
            worksheet.write(row_pos, 0, f"Profile: {profile}")
            row_pos += 2  # spacing before slot info

            for slot_idx, slot in enumerate(batch, start=1):
                if slot:  # Only write filled slots
                    part_name, qty, df_part = slot
                    worksheet.write(row_pos, 1, f"Slot #{slot_idx}: Part Name: {part_name}")  # Column B
                    worksheet.write(row_pos, 2, f"Quantity: {qty}")  # Column C
                    row_pos += 1

                    # Adjust column width for Slot and Quantity columns
                    slot_text = f"Slot #{slot_idx}: Part Name: {part_name}"
                    qty_text = f"Quantity: {qty}"
                    worksheet.set_column(1, 1, max(len(slot_text) * 1.2, 20))  # Column B
                    worksheet.set_column(2, 2, max(len(qty_text) * 1.2, 12))    # Column C

                    # Write headers, inserting "No" as first column
                    worksheet.write(row_pos, 0, "No")  # New column
                    for c, colname in enumerate(df_part.columns, start=1):
                        worksheet.write(row_pos, c, colname)
                    header_row = row_pos  # keep track for table range
                    row_pos += 1

                    # Write rows with numbering
                    for r_idx, r in enumerate(df_part.itertuples(index=False), start=1):
                        worksheet.write(row_pos, 0, r_idx)  # "No" column
                        for c, val in enumerate(r, start=1):
                            worksheet.write(row_pos, c, val)
                        row_pos += 1

                    # Add Excel table for this part
                    num_rows = row_pos - header_row
                    num_cols = len(df_part.columns) + 1  # +1 for "No"
                    worksheet.add_table(header_row, 0, row_pos - 1, num_cols - 1,
                                        {'columns': [{'header': "No"}] + [{'header': col} for col in df_part.columns]})

    print(f"\n✔ SAVED PROFILE: {profile} → {save_path}")
    if leftovers:
        print(f"⚠ {len(leftovers)} parts could not fit in any batch for profile {profile}")

if skipped_parts:
    skipped_df = pd.DataFrame(skipped_parts)
    skipped_path = os.path.join(output_folder, f"{input_base}_skipped_parts.xlsx")
    skipped_df.to_excel(skipped_path, index=False)
    print(f"⚠ Skipped parts report saved: {skipped_path}")

messagebox.showinfo("Done", f"All profiles saved in:\n{output_folder}")
