# /// script
# requires-python = ">=3.10"
# dependencies = []
# ///
"""Builds the LibreOffice versions (.ods) of all workbooks in files/.  Run with:  uv run tools/make_ods.py

Needs LibreOffice (`soffice`) on the PATH. Run it after the make_NN_*.py scripts.

Functions that Excel stores with the _xlfn. prefix (here: XLOOKUP) are not recognised by LibreOffice under
that name. In the .ods they are renamed to their ODF name (COM.MICROSOFT.XLOOKUP) and their cached result is
removed, so LibreOffice calculates them when the file is opened (XLOOKUP needs LibreOffice 24.8 or newer).
"""

import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

FILES = Path(__file__).resolve().parent.parent / "files"

CELL = re.compile(r'<table:table-cell\b([^>]*\btable:formula="([^"]*)"[^>]*?)(?:/>|>(.*?)</table:table-cell>)', re.S)
CACHED = re.compile(r'\s(?:office:value-type|office:value|office:string-value|office:boolean-value|calcext:value-type)="[^"]*"')
XLFN = re.compile(r"_xlfn\.([a-z0-9.]+)\(", re.I)


def fix_cell(m):
    attrs = m.group(1)
    if "_xlfn." not in m.group(2):
        return m.group(0)
    attrs = XLFN.sub(lambda f: f"COM.MICROSOFT.{f.group(1).upper()}(", attrs)
    return f"<table:table-cell{CACHED.sub('', attrs)}/>"


def fix_ods(path):
    """Rename _xlfn. functions and drop their cached results; returns the number of cells changed."""
    with zipfile.ZipFile(path) as src:
        content = src.read("content.xml").decode()
        fixed = CELL.sub(fix_cell, content)
        n = sum(1 for m in CELL.finditer(content) if "_xlfn." in m.group(2))
        if not n:
            return 0
        items = [(i, fixed.encode() if i.filename == "content.xml" else src.read(i.filename)) for i in src.infolist()]
    with zipfile.ZipFile(path, "w") as out:
        for info, data in items:  # "mimetype" must stay first and uncompressed
            out.writestr(info, data, zipfile.ZIP_STORED if info.filename == "mimetype" else zipfile.ZIP_DEFLATED)
    return n


def main():
    sources = sorted(FILES.glob("*.xlsx"))
    with tempfile.TemporaryDirectory() as profile:  # separate profile: works even if LibreOffice is open
        subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "ods",
                        "--outdir", str(FILES), *map(str, sources)], check=True, capture_output=True)
    for xlsx in sources:
        ods = xlsx.with_suffix(".ods")
        n = fix_ods(ods)
        print(f"Saved {ods}" + (f"  ({n} cells: _xlfn. functions renamed)" if n else ""))


if __name__ == "__main__":
    main()
