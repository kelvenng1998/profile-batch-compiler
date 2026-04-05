# Profile Batch Compiler

Python automation tool for compiling part data across multiple manufacturing profiles and organizing them into production-ready batch structures.

---

## 🚀 Features

- Reads BOM-style input Excel (`Filename` + `Quantity`)
- Searches part files across multiple profile folders
- Intelligent filename matching (supports variations like `sample_part_a - 24 holes`)
- Cleans and processes hole-based engineering data
- Packs parts into fixed-capacity batch slots
- Generates structured Excel outputs by profile and batch
- Produces skipped/missing parts report for validation

---

## 🧩 Workflow

1. Select input Excel file  
2. Select output folder  
3. System automatically:
   - locates matching part files  
   - cleans hole data  
   - groups parts by profile  
   - packs into batch slots  
   - exports formatted Excel outputs  

---

## 📁 Project Structure

    sample_data/
    ├─ input.xlsx
    ├─ Profile 1/Parts/By Name/
    ├─ Profile 2/Parts/By Name/
    ├─ Profile 3/Parts/By Name/
    └─ Profile 4/Parts/By Name/

---

## 📊 Output

The system generates:

- Profile-based Excel files  
  - input Profile 1.xlsx  
  - input Profile 2.xlsx  

- Each file contains:
  - Batch sheets  
  - Slot assignments  
  - Structured tables  

- Skipped parts report:
  - input_skipped_parts.xlsx  

---

## ⚠️ Error Handling

- Missing part files are logged  
- Invalid or unreadable Excel files are skipped safely  
- Reports are generated for debugging and validation  

---

## 🛠 Tech Stack

- Python  
- pandas  
- openpyxl  
- xlsxwriter  
- Tkinter  

---

## ▶️ How to Run

    pip install -r requirements.txt
    python main.py

---

## 💡 Notes

- Sample data is included for demonstration purposes  
- Ensure part files contain a `Sheet1` worksheet with required columns:
  - Hole length (mm)  
  - Hole 1 → Hole 5  

---

## 📌 Use Case

Designed for manufacturing and engineering workflows to automate:
- Part data compilation  
- Batch planning  
- Production preparation  
