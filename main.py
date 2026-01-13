
import os
import sys
import time
import shutil
import re
from pathlib import Path
import pandas as pd

# ---------- Optional GUI for folder picking ----------
try:
    import tkinter as tk
    from tkinter import filedialog
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False

# ---------- COM / AutoCAD ----------
try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False

# ---------- PDF merge ----------
try:
    from PyPDF2 import PdfMerger
    PYPDF2_AVAILABLE = True
except Exception:
    PYPDF2_AVAILABLE = False

# ---------- PDF text extraction / data ----------
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except Exception:
    PDFPLUMBER_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    PANDAS_AVAILABLE = False


# =========================
# Helpers
# =========================

def select_directory_gui(prompt_title: str) -> str:
    """Folder picker; falls back to console input."""
    if TK_AVAILABLE:
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askdirectory(title=prompt_title)
        root.update()
        if path:
            return path
    print(prompt_title)
    return input("> ").strip().strip('"')


def ensure_dir(path: Path):
    path.mkdir(parents=True, exist_ok=True)


def list_dwg_files(input_dir: Path, recursive: bool = True) -> list[Path]:
    pattern = "**/*.dwg" if recursive else "*.dwg"
    # Skip temp/hidden files that often cause issues
    return [p for p in input_dir.glob(pattern) if p.is_file() and not p.name.startswith("~")]


def copy_dwg_files(dwg_files: list[Path], dest_dir: Path) -> list[Path]:
    ensure_dir(dest_dir)
    copied = []
    for src in dwg_files:
        dest = dest_dir / src.name
        i = 1
        while dest.exists():
            stem, ext = os.path.splitext(src.name)
            dest = dest_dir / f"{stem} ({i}){ext}"
            i += 1
        shutil.copy2(src, dest)
        copied.append(dest)
    return copied


def sanitize_filename(name: str) -> str:
    """Replace characters illegal in Windows file names."""
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)


# =========================
# COM Retry Layer (critical)
# =========================

RPC_E_CALL_REJECTED = -2147418111  # "Call was rejected by callee."

def com_retry_call(func, *args, retries=50, delay=0.15, desc="", **kwargs):
    """
    Retry wrapper for COM invocations to handle AutoCAD being busy.
    Pumps messages and waits briefly before retrying when we see RPC_E_CALL_REJECTED.
    """
    last_exc = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pythoncom.com_error as e:
            hresult = getattr(e, "hresult", None) or (e.args[0] if e.args else None)
            if hresult == RPC_E_CALL_REJECTED:
                pythoncom.PumpWaitingMessages()
                time.sleep(delay)
                last_exc = e
                continue
            raise
    if desc:
        print(f"    ! COM call retries exhausted for: {desc}")
    if last_exc:
        raise last_exc


def com_get(obj, attr, **kw):
    return com_retry_call(lambda: getattr(obj, attr), desc=f"get {attr}", **kw)


def com_set(obj, attr, value, **kw):
    def setter():
        setattr(obj, attr, value)
        return True
    return com_retry_call(setter, desc=f"set {attr}", **kw)


def com_call(obj, method, *args, **kw):
    def inv():
        return getattr(obj, method)(*args)
    return com_retry_call(inv, desc=f"call {method}", **kw)


# =========================
# AutoCAD PDF Converter
# =========================

class AutoCADPdfConverter:
    """DWG -> individual layout PDFs using AutoCAD COM."""

    def __init__(self, pdf_pc3_name="DWG To PDF.pc3", visible=False):
        if not WIN32_AVAILABLE:
            raise RuntimeError("pywin32 not available. Install with: pip install pywin32")
        self.pdf_pc3_name = pdf_pc3_name
        self.visible = visible
        self.acad = None

    def __enter__(self):
        # Initialize COM in STA (Apartment threaded) – more reliable for AutoCAD
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception:
            pythoncom.CoInitialize()
        # Start AutoCAD
        try:
            try:
                self.acad = win32com.client.gencache.EnsureDispatch("AutoCAD.Application")
            except Exception:
                self.acad = win32com.client.Dispatch("AutoCAD.Application")
            com_set(self.acad, "Visible", self.visible)
            # Probe readiness
            _ = com_get(self.acad, "Version")
        except Exception as e:
            raise RuntimeError("AutoCAD COM interface not available. Ensure AutoCAD is installed and can launch without modal dialogs.") from e
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def open_document(self, dwg_path: Path):
        docs = com_get(self.acad, "Documents")
        # Use single-argument Open for compatibility
        doc = com_call(docs, "Open", str(dwg_path))
        # Wait until document is ready by probing FullName
        for _ in range(100):
            try:
                _ = com_get(doc, "FullName")
                break
            except Exception:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.1)
        return doc

    def close_document(self, doc, save=False):
        try:
            com_call(doc, "Close", bool(save))
        except Exception:
            pass

    def convert_individual(self, dwg_path: Path, out_dir: Path):
        """Create one PDF per Paper Space layout for this DWG. Returns list of PDF paths."""
        generated = []
        doc = None
        try:
            doc = self.open_document(dwg_path)
        except Exception as e:
            print(f"    ! Failed to open {dwg_path.name}: {e}")
            return generated

        try:
            layouts = com_get(doc, "Layouts")
            count = com_get(layouts, "Count")
            for i in range(count):
                layout = com_call(layouts, "Item", i)
                name = com_get(layout, "Name")
                if name.lower() == "model":
                    continue

                # Refresh and set plot device
                try:
                    try:
                        com_call(layout, "RefreshPlotDeviceInfo")
                    except Exception:
                        pass
                    com_set(layout, "ConfigName", self.pdf_pc3_name)
                except Exception:
                    print(f"    ! Plotter '{self.pdf_pc3_name}' unavailable for layout '{name}'")
                    continue

                # Activate this layout
                try:
                    com_set(doc, "ActiveLayout", layout)
                except Exception:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.2)
                    try:
                        com_set(doc, "ActiveLayout", layout)
                    except Exception:
                        print(f"    ! Could not activate layout '{name}'")
                        continue

                pdf_name = f"{dwg_path.stem}__{sanitize_filename(name)}.pdf"
                pdf_path = out_dir / pdf_name

                # Plot
                try:
                    plot = com_get(doc, "Plot")
                    com_call(plot, "PlotToFile", str(pdf_path))
                    # Wait for file to finish writing
                    t0 = time.time()
                    while time.time() - t0 < 30:
                        if pdf_path.exists() and pdf_path.stat().st_size > 0:
                            break
                        pythoncom.PumpWaitingMessages()
                        time.sleep(0.2)
                    if pdf_path.exists() and pdf_path.stat().st_size > 0:
                        print(f"    ✓ {pdf_path.name}")
                        generated.append(pdf_path)
                    else:
                        print(f"    ! Plot finished but no file for layout '{name}'")
                except Exception as e:
                    print(f"    ! Failed to plot layout '{name}': {e}")

        finally:
            if doc is not None:
                self.close_document(doc, save=False)

        return generated


# =========================
# Merge PDFs
# =========================

def merge_pdfs_in_order(pdf_paths: list[Path], output_path: Path):
    if not PYPDF2_AVAILABLE:
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
    print(f"\n✓ Combined PDF created: {output_path.name}")


# =========================
# Sheet List Extraction
# =========================

SHEET_KEYWORDS = (
    "PLAN", "ELEVATION", "SECTION", "DETAIL", "SCHEDULE",
    "P&ID", "PROCESS", "DEMO", "FAB", "TITLE", "GENERAL", "NOTES"
)

def extract_sheet_info(pdf_path: Path, crop_frac: float = 0.22) -> pd.DataFrame:
    """
    Extract Sheet Number, Sheet Name, Revision from the combined PDF.
    - Crops bottom-right area (title block) where possible.
    - Uses heuristics & regex to find values.
    Returns a pandas DataFrame with columns: Page, Sheet Number, Sheet Name, Revision.
    """
    if not PDFPLUMBER_AVAILABLE:
        raise RuntimeError("pdfplumber not available. Install with: pip install pdfplumber")
    if not PANDAS_AVAILABLE:
        raise RuntimeError("pandas not available. Install with: pip install pandas")

    results = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            width, height = page.width, page.height
            crop_box = (width * (1 - crop_frac), height * (1 - crop_frac), width, height)

            # Try cropped title block area; fallback to full page text
            try:
                cropped_page = page.crop(crop_box)
                text = cropped_page.extract_text() or ""
                if not text.strip():
                    text = page.extract_text() or ""
            except Exception:
                text = page.extract_text() or ""

            # Normalize and split lines
            lines = [ln.strip() for ln in (text.split("\n") if text else []) if ln.strip()]

            # Defaults
            sheet_no = "Not Found"
            sheet_name = "Not Found"
            revision = "Not Found"

            # --- Heuristic A: labeled fields (more robust if present) ---
            labeled_patterns = [
                r'(?:SHEET\s*(?:NO|NUMBER)|DWG\s*(?:NO|NUMBER)|DRAWING\s*(?:NO|NUMBER))[:\s]*([A-Z0-9.\-]+)',
                r'(?:REV(?:ISION)?)[:\s\-]*([A-Z0-9.\-]+)',
            ]
            block_text = " ".join(lines)

            # Revision first
            rev_match = re.search(r'\bREV(?:ISION)?[:\s\-]*([A-Z0-9.\-]+)\b', block_text, flags=re.IGNORECASE)
            if rev_match:
                revision = rev_match.group(1).upper()

            # Sheet number via labeled field
            no_match = re.search(r'(?:SHEET\s*(?:NO|NUMBER)|DWG\s*(?:NO|NUMBER)|DRAWING\s*(?:NO|NUMBER))[:\s]*([A-Z0-9.\-]+)', block_text, flags=re.IGNORECASE)
            if no_match:
                sheet_no = no_match.group(1).upper()

            # --- Heuristic B: pattern-based (handles FA-DD 205 etc.) ---
            if sheet_no == "Not Found":
                patterns = [
                    r'\bFA-[A-Z]+(?:\s*[A-Z0-9.\-]+)?\b',     # FA-DD 205, FA-J 8BIRW001
                    r'\b[A-Z]{1,3}\s*-\s*[A-Z0-9.\-]+\b',     # A-101, M-2.1
                    r'\b[A-Z]{1,3}\s*[0-9][A-Z0-9.\-]*\b',    # A101, M2.1
                ]
                for line in reversed(lines):  # Often at bottom-most lines
                    for pat in patterns:
                        m = re.search(pat, line)
                        if m:
                            sheet_no = m.group(0).replace(" ", "").upper()
                            break
                    if sheet_no != "Not Found":
                        break

            # --- Heuristic C: Sheet name candidates (uppercase title near title block) ---
            # Prefer lines containing keywords; otherwise pick the longest ALL-CAPS near bottom.
            candidates = []
            for ln in lines:
                ln_clean = ln.strip()
                # Skip obvious metadata
                if any(k in ln_clean.upper() for k in ("CONFIDENTIAL", "PROJECT", "CLIENT", "DRAWN", "CHECKED", "APPROVED", "SCALE")):
                    continue
                # Capture keyword lines
                if any(k in ln_clean.upper() for k in SHEET_KEYWORDS):
                    candidates.append(ln_clean)
                else:
                    # Heuristic: all caps / mostly caps
                    letters = re.sub(r'[^A-Za-z]', '', ln_clean)
                    if letters and (sum(1 for c in letters if c.isupper()) / max(1, len(letters))) > 0.6 and len(ln_clean) >= 6:
                        candidates.append(ln_clean)

            if candidates:
                # Prefer the last candidate in the cropped area (closer to title block bottom)
                sheet_name = candidates[-1].upper()

            results.append({
                "Page": i,
                "Sheet Number": sheet_no,
                "Sheet Name": sheet_name,
                "Revision": revision
            })

    return pd.DataFrame(results)


def save_sheet_list(df: pd.DataFrame, output_dir: Path, project_name: str):
    """Save the sheet list as CSV (always) and Excel if openpyxl is available."""
    ensure_dir(output_dir)
    csv_path = output_dir / f"{sanitize_filename(project_name)}_Sheet List.csv"
    df.to_csv(csv_path, index=False)
    print(f"✓ Sheet List (CSV) saved: {csv_path.name}")

    # Excel optional
    try:
        # pandas will use openpyxl for .xlsx
        xlsx_path = output_dir / f"{sanitize_filename(project_name)}_Sheet List.xlsx"
        df.to_excel(xlsx_path, index=False, engine="openpyxl")
        print(f"✓ Sheet List (Excel) saved: {xlsx_path.name}")
    except Exception:
        print("    ! Excel export skipped (openpyxl not installed or write error).")


# =========================
# Main
# =========================

def main():
    print("DWG Copier and PDF Converter + Sheet List")
    print("-----------------------------------------")

    # 1) Inputs
    input_dir = Path(select_directory_gui("Select INPUT folder (DWG source)")).resolve()
    output_dir = Path(select_directory_gui("Select OUTPUT folder (PDF target)")).resolve()

    default_project_name = input_dir.name
    project_name = input(f"\nEnter Project Name [{default_project_name}]: ").strip() or default_project_name

    recursive_input = input("\nInclude subfolders when searching for DWG files? (Y/N): ").strip().lower()
    recursive = recursive_input.startswith('y')

    # 2) Prepare folders
    ensure_dir(output_dir)
    dwg_out_dir = output_dir / "DWG"
    individual_pdf_dir = output_dir / "Individual PDFs"
    ensure_dir(dwg_out_dir)
    ensure_dir(individual_pdf_dir)

    # 3) Find + copy DWGs
    print(f"\nScanning for DWGs in: {input_dir}")
    dwg_files = list_dwg_files(input_dir, recursive)
    if not dwg_files:
        print("No DWG files found. Exiting.")
        sys.exit(0)
    print(f"Found {len(dwg_files)} DWG file(s).")

    print(f"\nCopying DWGs to: {dwg_out_dir}")
    copied = copy_dwg_files(dwg_files, dwg_out_dir)
    print(f"Copied {len(copied)} DWG file(s).")

    # 4) Convert with AutoCAD
    if not WIN32_AVAILABLE:
        print("ERROR: pywin32 not installed or AutoCAD COM unavailable.")
        sys.exit(1)
    if not PYPDF2_AVAILABLE:
        print("ERROR: PyPDF2 not installed. Install with: pip install PyPDF2")
        sys.exit(1)

    combined_pdf_path = output_dir / f"{sanitize_filename(project_name)}_Combined.pdf"
    all_pdfs: list[Path] = []

    print("\nStarting AutoCAD for PDF generation...")
    try:
        # visible=True can reduce COM rejections in some environments
        with AutoCADPdfConverter(pdf_pc3_name="DWG To PDF.pc3", visible=False) as converter:
            print("\nGenerating Individual PDFs...")
            total = len(copied)
            for idx, dwg in enumerate(copied, 1):
                print(f"[{idx}/{total}] {dwg.name}")
                generated = converter.convert_individual(dwg, individual_pdf_dir)
                all_pdfs.extend(generated)
    except Exception as e:
        print(f"\nERROR: AutoCAD conversion failed: {e}")
        sys.exit(2)

    # 5) Merge into single combined PDF
    if all_pdfs:
        print(f"\nMerging {len(all_pdfs)} PDFs into one combined file...")
        # Keep order stable by sorting by DWG name then layout name
        all_pdfs_sorted = sorted(all_pdfs, key=lambda p: (p.stem.split("__")[0].lower(), p.stem.lower()))
        try:
            merge_pdfs_in_order(all_pdfs_sorted, combined_pdf_path)
        except Exception as e:
            print(f"    ! Failed to merge PDFs: {e}")
    else:
        print("\nNo PDFs generated to merge.")

    # 6) Extract Sheet List from combined PDF
    if PDFPLUMBER_AVAILABLE and PANDAS_AVAILABLE and combined_pdf_path.exists():
        print("\nExtracting Sheet List from combined PDF...")
        try:
            df = extract_sheet_info(combined_pdf_path, crop_frac=0.22)
            save_sheet_list(df, output_dir, project_name)
        except Exception as e:
            print(f"    ! Sheet List extraction failed: {e}")
    else:
        print("    ! Sheet List extraction skipped (pdfplumber/pandas not available or combined PDF missing).")

    print("\nDone.")


if __name__ == "__main__":
    main()
