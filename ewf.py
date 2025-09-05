#!/usr/bin/env python3
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


def open_image(image_path):
    # Open EWF (E01) image
    filenames = pyewf.glob(image_path)
    ewf_handle = pyewf.handle()
    ewf_handle.open(filenames)

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

    return EWFFileLikeObject(ewf_handle)


def search_files(directory, search_term):
    for entry in directory:
        if entry.info.name.name in [b".", b".."]:
            continue

        try:
            fname = entry.info.name.name.decode("utf-8", "ignore")
        except Exception:
            fname = str(entry.info.name.name)

        # Check for filename match
        if search_term.lower() in fname.lower():
            deleted = bool(entry.info.meta and entry.info.meta.flags & pytsk3.TSK_FS_META_FLAG_UNALLOC)
            print(f"[+] Found: {fname} (Deleted: {deleted})")

        # Recurse into subdirectories
        if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
            try:
                subdir = entry.as_directory()
                search_files(subdir, search_term)
            except Exception:
                continue


def main(image_file, search_term):
    img_info = open_image(image_file)
    fs = pytsk3.FS_Info(img_info)

    root_dir = fs.open_dir("/")
    search_files(root_dir, search_term)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <image.E01> <filename>")
        sys.exit(1)

    main(sys.argv[1], sys.argv[2])
