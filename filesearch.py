#!/usr/bin/env python3
import pandas as pd
import zipfile
import os
from datetime import datetime
import argparse
from tqdm import tqdm
import sys

# Try to import pyewf (Kali may use libewf instead)
try:
    import pyewf
except ImportError:
    try:
        import libewf as pyewf
    except ImportError:
        print("[-] Neither pyewf nor libewf found. Install with: sudo apt install python3-libewf")
        sys.exit(1)

import pytsk3

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


def search_in_ewf(filenames, ewf_path):
    import pyewf
    import pytsk3

    results = []

    # Open EWF (E01) image
    ewf_filenames = pyewf.glob(ewf_path)
    ewf_handle = pyewf.handle()
    ewf_handle.open(ewf_filenames)

    # Expose it as a file-like object for pytsk3
    class EWFFileLikeObject(pytsk3.Img_Info):
        def __init__(self, ewf_handle):
            self._ewf_handle = ewf_handle
            super().__init__()

        def read(self, offset, size):
            self._ewf_handle.seek(offset)
            return self._ewf_handle.read(size)

        def get_size(self):
            return self._ewf_handle.get_media_size()

    img_info = EWFFileLikeObject(ewf_handle)
    fs = pytsk3.FS_Info(img_info)

    def search_files(directory, search_term):
        for entry in directory:
            if entry.info.name.name in [b".", b".."]:
                continue

            try:
                fname = entry.info.name.name.decode("utf-8", "ignore")
            except Exception:
                fname = str(entry.info.name.name)

            # Check for filename match
            if search_term.lower() == fname.lower():
                deleted = bool(entry.info.meta and entry.info.meta.flags & pytsk3.TSK_FS_META_FLAG_UNALLOC)
                file_size = entry.info.meta.size if entry.info.meta else 0
                mtime = datetime.fromtimestamp(entry.info.meta.mtime) if entry.info.meta and entry.info.meta.mtime else None
                results.append({
                    "searched_filename": search_term,
                    "found_path": fs.path + '/' + fname,
                    "file_size": file_size,
                    "compress_size": None,
                    "date_time": mtime,
                    "source": ewf_path + (" (Deleted)" if deleted else "")
                })

            # Recurse into subdirectories
            if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
                try:
                    subdir = entry.as_directory()
                    search_files(subdir, search_term)
                except Exception:
                    continue

    root_dir = fs.open_dir("/")
    for file_to_find in tqdm(filenames, desc="Searching in EWF"):
        search_files(root_dir, file_to_find)

    return results

def search_files(excel_path, column_name, output_path, zip_path=None, drive_path=None, ewf_path=None):
    # Read Excel file
    df = pd.read_excel(excel_path)

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in Excel file.")

    filenames = df[column_name].dropna().astype(str).tolist()
    filenames = [f.lower() for f in filenames]  # case-insensitive

    results = []

    if zip_path and drive_path and ewf_path:
        raise ValueError("Please provide only one of --zip, --drive, or --ewf, not all.")
    elif zip_path:
        results.extend(search_in_zip(filenames, zip_path))
    elif drive_path:
        results.extend(search_in_drive(filenames, drive_path))
    elif ewf_path:
        results.extend(search_in_ewf(filenames, ewf_path))
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
    group.add_argument("--ewf", help="Path to EWF image file")

    args = parser.parse_args()

    search_files(args.excel, args.column, args.output, zip_path=args.zip, drive_path=args.drive, ewf_path=args.ewf)
