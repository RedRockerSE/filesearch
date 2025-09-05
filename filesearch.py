import pandas as pd
import zipfile
import os
from datetime import datetime
import argparse
from tqdm import tqdm


def search_in_zip(filenames, zip_path):
    results = []
    with zipfile.ZipFile(zip_path, 'r') as z:
        for file_to_find in tqdm(filenames, desc="Searching in ZIP"):
            for zip_info in z.infolist():
                if os.path.basename(zip_info.filename).lower() == file_to_find:
                    results.append({
                        "searched_filename": file_to_find,
                        "found_path": zip_info.filename,
                        "file_size": zip_info.file_size,
                        "compress_size": zip_info.compress_size,
                        "date_time": datetime(*zip_info.date_time),
                        "source": zip_path
                    })
    return results


def search_in_drive(filenames, drive_path):
    results = []
    for root, _, files in tqdm(os.walk(drive_path), desc="Searching drive"):
        for file_to_find in filenames:
            for f in files:
                if f.lower() == file_to_find:
                    full_path = os.path.join(root, f)
                    stats = os.stat(full_path)
                    results.append({
                        "searched_filename": file_to_find,
                        "found_path": full_path,
                        "file_size": stats.st_size,
                        "compress_size": None,
                        "date_time": datetime.fromtimestamp(stats.st_mtime),
                        "source": drive_path
                    })
    return results


def search_files(excel_path, column_name, output_path, zip_path=None, drive_path=None):
    # Read Excel file
    df = pd.read_excel(excel_path)

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in Excel file.")

    filenames = df[column_name].dropna().astype(str).tolist()
    filenames = [f.lower() for f in filenames]  # case-insensitive

    results = []

    if zip_path and drive_path:
        raise ValueError("Please provide only one of --zip or --drive, not both.")
    elif zip_path:
        results.extend(search_in_zip(filenames, zip_path))
    elif drive_path:
        results.extend(search_in_drive(filenames, drive_path))
    else:
        raise ValueError("Either --zip or --drive must be provided.")

    # Convert results to DataFrame
    if results:
        results_df = pd.DataFrame(results)
    else:
        results_df = pd.DataFrame(columns=[
            "searched_filename", "found_path", "file_size",
            "compress_size", "date_time", "source"
        ])

    # Export to CSV or Excel depending on output extension
    if output_path.endswith(".csv"):
        results_df.to_csv(output_path, index=False)
    else:
        results_df.to_excel(output_path, index=False)

    print(f"Results exported to {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Search filenames from Excel in ZIP or drive.")

    parser.add_argument("--excel", required=True, help="Path to Excel file containing filenames")
    parser.add_argument("--column", required=True, help="Column name containing filenames")
    parser.add_argument("--output", required=True, help="Output file path (.csv or .xlsx)")

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--zip", help="Path to ZIP file")
    group.add_argument("--drive", help="Path to drive or folder to search")

    args = parser.parse_args()

    search_files(args.excel, args.column, args.output, zip_path=args.zip, drive_path=args.drive)
