
from typing import List
from pathlib import Path

try:
    from PyPDF2 import PdfMerger
except Exception as e:
    PdfMerger = None

from ..io.files import ensure_dir

def merge_pdfs_in_order(pdf_paths: List[Path], output_path: Path):
    if PdfMerger is None:
        raise RuntimeError("PyPDF2 not available. Install with: pip install PyPDF2")
    if not pdf_paths:
        print("    ! No PDFs to merge.")
        return
    merger = PdfMerger()
    try:
        for p in pdf_paths:
            try:
                merger.append(str(p))
            except Exception as e:
                print(f"    ! Failed to append {p.name}: {e}")
        ensure_dir(output_path.parent)
        merger.write(str(output_path))
    finally:
        merger.close()
    print(f"✓ Combined PDF created: {output_path.name}")
