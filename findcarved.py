import os
import pandas as pd

# --- CONFIG ---
folder_path = "./excel_files"   # path to folder containing your Excel files
output_file = "carved_report.xlsx"  # name of the output Excel file

# Columns we need from source Excel files
file_name_col = "File Name"
file_path_col = "File Path"

# Keywords to search for inside File Path
keywords = ["$OrphanedFiles", "$RECYCLE.BIN", "Unallocated Files"]

# --- SCRIPT ---
all_matches = []

# Loop through all Excel files in the folder
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        filepath = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(filepath, usecols=[file_name_col, file_path_col])

            # Check for rows where File Path contains any of the keywords
            mask = df[file_path_col].astype(str).apply(
                lambda x: any(keyword in x for keyword in keywords)
            )
            matches = df[mask].copy()
            matches["SourceFile"] = file

            if not matches.empty:
                all_matches.append(matches)
        except Exception as e:
            print(f"Skipping {file}: {e}")

# Combine and save to Excel
if all_matches:
    result = pd.concat(all_matches, ignore_index=True)
    result.to_excel(output_file, index=False)
    print(f"Carved files report saved to {output_file}")
else:
    print("No matching rows found across files.")
