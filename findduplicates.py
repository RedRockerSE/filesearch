import os
import pandas as pd

# --- CONFIG ---
folder_path = "./excel_files"   # path to folder containing your Excel files
filename_column = "Filename"    # name of the column that contains filenames
output_file = "duplicates_report.xlsx"  # name of the output Excel file

# --- SCRIPT ---
all_data = []

# Loop through all Excel files in the folder
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        filepath = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(filepath, usecols=[filename_column])
            df["SourceFile"] = file  # Keep track of which Excel file this row came from
            all_data.append(df)
        except Exception as e:
            print(f"Skipping {file}: {e}")

# Combine everything into one DataFrame
if all_data:
    combined = pd.concat(all_data, ignore_index=True)

    # Group by filename and collect the unique files each appears in
    grouped = combined.groupby(filename_column)["SourceFile"].unique().reset_index()

    # Keep only filenames that appear in 2 or more different Excel files
    grouped["FileCount"] = grouped["SourceFile"].apply(len)
    duplicates = grouped[grouped["FileCount"] > 1]

    if not duplicates.empty:
        duplicates.to_excel(output_file, index=False)
        print(f"Duplicates report saved to {output_file}")
    else:
        print("No duplicates found across files.")
else:
    print("No Excel files found in the folder.")
