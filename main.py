import os
import sys
import time
import shutil
import re
import csv
from pathlib import Path

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


# =========================
# Configuration (fast path for your title block name)
# =========================
# If you know the exact sheet-specific title block name(s), add them here (UPPERCASE).
# This makes the scanner extremely fast because it skips all other blocks immediately.
TARGET_BLOCK_NAMES = {
    # Example from your screenshot (update to your actual block name(s)):
    # "GF MALTA TITLE BLOCK 30X42-TB-ATT",
}

# Joiners (formatting)
SHEETNO_SEPARATOR = "-"             # e.g., "FA-JD-8BIRW002"
TITLE_JOINER = " "                  # e.g., "FAB - I&C BIRW (IRW) TANK DEMO P&ID 2 OF 2"

# Revision index range to scan (R0..R9 => 0..9). Increase if you use more than R9.
REV_INDEX_RANGE = range(0, 10)


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


def list_dwg_files(input_dir: Path, recursive: bool = True) -> list:
    pattern = "**/*.dwg" if recursive else "*.dwg"
    return [p for p in input_dir.glob(pattern) if p.is_file() and not p.name.startswith("~")]


def copy_dwg_files(dwg_files: list, dest_dir: Path) -> list:
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
# PDF Merge
# =========================

def merge_pdfs_in_order(pdf_paths: list, output_path: Path):
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
# Title Block Scanning (fast + robust)
# =========================

def _norm_tag(s: str) -> str:
    """Normalize attribute tag names for matching (upper + strip non-alnum)."""
    return re.sub(r'[^A-Z0-9]', '', (s or '').upper())

def _norm_prompt(s: str) -> str:
    """Normalize attribute prompts similarly."""
    return re.sub(r'[^A-Z0-9]', '', (s or '').upper())

def _norm_name(s: str) -> str:
    """Normalize block name (upper + trim)."""
    return (s or '').upper().strip()


def read_titleblock_from_active_layout_robust(
    doc,
    *,
    # Fast path: match by block name(s)
    target_block_names=None,
    # Sheet number: tags & prompts (we match both)
    sheet_top_tags=("FC-E", "TOPNUMBER", "TOP_NUM", "TOP"),
    sheet_top_prompts=("TOP NUMBER",),
    sheet_bottom_tags=("442C", "BOTTOMNUMBER", "BOTTOM_NUM", "BOTTOM"),
    sheet_bottom_prompts=("BOTTOM NUMBER",),
    # Title lines: prompts TITLE_1..TITLE_5 + a common alias you showed ("ELECTRICAL")
    title_tag_primary=("TITLE_1", "TITLE_2", "TITLE_3", "TITLE_4", "TITLE_5"),
    title_prompt_primary=("TITLE_1", "TITLE_2", "TITLE_3", "TITLE_4", "TITLE_5"),
    title_prompts_alias=("ELECTRICAL",),
    # Revisions: R#NO, R#DATE, R#DESC, R#BY for indices in REV_INDEX_RANGE
    revision_index_range=REV_INDEX_RANGE,
    # Formatting
    sheetno_sep=SHEETNO_SEPARATOR,
    title_joiner=TITLE_JOINER,
):
    """
    Read the sheet-specific title block attributes on the ACTIVE layout (PaperSpace).

    Returns **7 values always**:
        (sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs_dict)

    Logic:
    - If target_block_names provided, filter PaperSpace blocks by name (fast).
    - Else, scan & score blocks against known tags/prompts (robust).
    - Skip XREF blocks where detectable.
    - Match both by TagString and PromptString for sheet/ title fields.
    """
    # Normalize config
    tb_names = {_norm_name(n) for n in (target_block_names or [])} if target_block_names else None

    sheet_top_tag_norm = {_norm_tag(t) for t in sheet_top_tags}
    sheet_top_prompt_norm = {_norm_prompt(p) for p in sheet_top_prompts}

    sheet_bot_tag_norm = {_norm_tag(t) for t in sheet_bottom_tags}
    sheet_bot_prompt_norm = {_norm_prompt(p) for p in sheet_bottom_prompts}

    title_tag_norm_order = [_norm_tag(t) for t in title_tag_primary]          # ordered by TITLE_1..TITLE_5
    title_prompt_norm_order = [_norm_prompt(p) for p in title_prompt_primary] # same positions
    title_prompt_alias_norm = {_norm_prompt(p) for p in title_prompts_alias}  # alias to TITLE_1

    # Outputs
    sheet_no = "Not Found"
    sheet_title = "Not Found"
    rev_no = "Not Found"
    rev_date = "Not Found"
    rev_desc = "Not Found"
    rev_by = "Not Found"
    other_attrs = {}

    # Access PaperSpace
    try:
        ps = com_get(doc, "PaperSpace")
        ps_count = com_get(ps, "Count")
    except Exception:
        return sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs

    def collect_attrs(ent):
        """
        Return list of dicts:
          {"raw_tag","norm_tag","raw_prompt","norm_prompt","value"}
        """
        out = []
        try:
            if not com_get(ent, "HasAttributes"):
                return out
            arr = com_call(ent, "GetAttributes")
        except Exception:
            return out
        for a in arr:
            try:
                raw_tag = com_get(a, "TagString") or ""
            except Exception:
                raw_tag = ""
            try:
                raw_prompt = com_get(a, "PromptString") or ""
            except Exception:
                raw_prompt = ""  # Not always available; safe to blank
            norm_tag = _norm_tag(raw_tag)
            norm_prompt = _norm_prompt(raw_prompt)
            try:
                val = (com_get(a, "TextString") or "").strip()
            except Exception:
                val = ""
            out.append({
                "raw_tag": raw_tag, "norm_tag": norm_tag,
                "raw_prompt": raw_prompt, "norm_prompt": norm_prompt,
                "value": val
            })
        return out

    def analyze_block(attrs):
        """
        Score a block & extract fields:
          - Sheet number top/bottom (by tag OR prompt)
          - Title lines (TITLE_1..TITLE_5 by tag OR prompt + alias into TITLE_1)
          - Revisions R#NO/DATE/DESC/BY
        """
        score = 0
        top, bot = "", ""
        # titles stored by TITLE_1..TITLE_5 normalized tag keys
        titles = {t: "" for t in title_tag_norm_order}
        revs = {}  # idx -> {"NO","DATE","DESC","BY"}

        # First pass: detect sheet number and titles by tag/prompt
        for attr in attrs:
            val = attr["value"]
            if not val:
                continue
            nt = attr["norm_tag"]
            np = attr["norm_prompt"]

            # Sheet top
            if (nt in sheet_top_tag_norm or np in sheet_top_prompt_norm) and not top:
                top = val; score += 3
                continue
            # Sheet bottom
            if (nt in sheet_bot_tag_norm or np in sheet_bot_prompt_norm) and not bot:
                bot = val; score += 3
                continue

            # Titles: tag match in correct order
            if nt in titles and not titles[nt]:
                titles[nt] = val; score += 1
                continue

            # Titles: prompt match (TITLE_1..TITLE_5)
            if np in title_prompt_norm_order:
                idx = title_prompt_norm_order.index(np)
                t_norm = title_tag_norm_order[idx]
                if not titles[t_norm]:
                    titles[t_norm] = val; score += 1
                continue

        # Alias to TITLE_1 (e.g., ELECTRICAL)
        t1 = title_tag_norm_order[0] if title_tag_norm_order else None
        if t1 and not titles[t1]:
            for attr in attrs:
                if not attr["value"]:
                    continue
                if attr["norm_prompt"] in title_prompt_alias_norm or attr["norm_tag"] in {_norm_tag("ELECTRICAL")}:
                    titles[t1] = attr["value"]; score += 1
                    break

        # Revisions R#NO/DATE/DESC/BY (by tag only — prompts are less consistent)
        for attr in attrs:
            val = attr["value"]
            if not val:
                continue
            m = re.match(r'^R(\d+)(NO|DATE|DESC|BY)$', attr["norm_tag"])
            if m:
                idx = int(m.group(1))
                if idx in revision_index_range:
                    kind = m.group(2)
                    if idx not in revs:
                        revs[idx] = {"NO": "", "DATE": "", "DESC": "", "BY": ""}
                    revs[idx][kind] = val
                    score += 1

        return score, top, bot, titles, revs

    best = {
        "score": -1,
        "attrs": [],
        "top": "",
        "bot": "",
        "titles": {t: "" for t in title_tag_norm_order},
        "revs": {},
        "attrs_count": 0,
        "blkname": "",
    }

    def consider_entity(ent):
        nonlocal best
        # Filter by block name if provided
        try:
            blkname = com_get(ent, "Name") or ""
        except Exception:
            blkname = ""
        blkname_norm = _norm_name(blkname)

        # Skip XREF blocks if detectable
        try:
            if com_get(ent, "IsXRef"):
                return
        except Exception:
            pass

        if TARGET_BLOCK_NAMES:
            # Use configured globals unless overridden by param
            tb_names_local = tb_names if tb_names is not None else {_norm_name(n) for n in TARGET_BLOCK_NAMES}
        else:
            tb_names_local = tb_names

        if tb_names_local is not None:
            if blkname_norm not in tb_names_local:
                return  # not our sheet-specific block name

        attrs = collect_attrs(ent)
        if not attrs:
            return

        score, top, bot, titles, revs = analyze_block(attrs)
        if (score > best["score"]) or (score == best["score"] and len(attrs) > best["attrs_count"]):
            best = {
                "score": score,
                "attrs": attrs,
                "top": top,
                "bot": bot,
                "titles": titles,
                "revs": revs,
                "attrs_count": len(attrs),
                "blkname": blkname,
            }

    # Iterate PaperSpace once (fast path if block names are known)
    for i in range(ps_count):
        try:
            ent = com_call(ps, "Item", i)
        except Exception:
            continue
        consider_entity(ent)
        # If block names are enforced and we already found a high-score block,
        # you could break early. We keep scanning in case multiple instances exist.

    # No candidate found
    if best["score"] < 0:
        return sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs

    # Build sheet number
    if best["top"] and best["bot"]:
        sheet_no = f"{best['top']}{sheetno_sep}{best['bot']}"
    elif best["top"]:
        sheet_no = best["top"]
    elif best["bot"]:
        sheet_no = best["bot"]

    # Build title (TITLE_1..TITLE_5)
    title_vals = [best["titles"][t] for t in title_tag_norm_order if best["titles"][t]]
    if title_vals:
        sheet_title = title_joiner.join(title_vals)

    # Pick highest revision index that has NO
    if best["revs"]:
        max_idx = max((idx for idx, d in best["revs"].items() if d.get("NO")), default=None)
        if max_idx is not None:
            d = best["revs"][max_idx]
            rev_no = d.get("NO") or rev_no
            rev_date = d.get("DATE") or rev_date
            rev_desc = d.get("DESC") or rev_desc
            rev_by = d.get("BY") or rev_by

    # Build other_attrs: include everything else from chosen block
    # Exclude used tags/prompts (by normalized forms)
    used_norm_tags = set()
    used_norm_prompts = set()

    used_norm_tags |= sheet_top_tag_norm
    used_norm_prompts |= sheet_top_prompt_norm
    used_norm_tags |= sheet_bot_tag_norm
    used_norm_prompts |= sheet_bot_prompt_norm

    used_norm_tags |= set(title_tag_norm_order)
    used_norm_prompts |= set(title_prompt_norm_order)
    used_norm_prompts |= set(title_prompt_alias_norm)

    for idx in revision_index_range:
        for kind in ("NO", "DATE", "DESC", "BY"):
            used_norm_tags.add(_norm_tag(f"R{idx}{kind}"))

    # Collect remaining attributes using *raw tag name* as the column header
    for a in best["attrs"]:
        if not a["value"]:
            continue
        if a["norm_tag"] in used_norm_tags or a["norm_prompt"] in used_norm_prompts:
            continue
        # Keep first occurrence
        if a["raw_tag"] and a["raw_tag"] not in other_attrs:
            other_attrs[a["raw_tag"]] = a["value"]
        elif not a["raw_tag"]:
            # If raw tag is empty (rare), fallback to prompt as key
            key = a["raw_prompt"] or "ATTR"
            if key not in other_attrs:
                other_attrs[key] = a["value"]

    return sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs


# =========================
# AutoCAD PDF Converter + Drawing List
# =========================

class AutoCADPdfConverter:
    """DWG -> individual layout PDFs + per-layout Drawing List rows."""

    def __init__(self, pdf_pc3_name="DWG To PDF.pc3", visible=False):
        if not WIN32_AVAILABLE:
            raise RuntimeError("pywin32 not available. Install with: pip install pywin32")
        self.pdf_pc3_name = pdf_pc3_name
        self.visible = visible
        self.acad = None

    def __enter__(self):
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception:
            pythoncom.CoInitialize()
        try:
            try:
                self.acad = win32com.client.gencache.EnsureDispatch("AutoCAD.Application")
            except Exception:
                self.acad = win32com.client.Dispatch("AutoCAD.Application")
            com_set(self.acad, "Visible", self.visible)
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
        doc = com_call(docs, "Open", str(dwg_path))  # single-arg Open for compatibility
        # Probe readiness
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

    def convert_individual_and_collect_rows(self, dwg_path: Path, out_dir: Path, rows: list):
        """
        - EXACT plotting behavior as your working version
        - Always writes one row per Paper layout (even if plot fails)
        - Adds Sheet/Title/Revision fields and the remaining attributes as extra columns
        """
        generated = []
        doc = None
        try:
            doc = self.open_document(dwg_path)
        except Exception as e:
            rows.append({
                "DWG File": dwg_path.name,
                "Layout": "",
                "Sheet Number": "Not Found",
                "Sheet Name": "Not Found",
                "Revision Number": "Not Found",
                "Revision Date": "Not Found",
                "Revision Description": "Not Found",
                "Revision By": "Not Found",
                "Plot Successful": "No",
                "Status": "Open Failed",
                "Error": str(e),
                "Individual PDF": ""
            })
            print(f"    ! Failed to open {dwg_path.name}: {e}")
            return generated

        try:
            layouts = com_get(doc, "Layouts")
            count = com_get(layouts, "Count")
            found_paper = False

            for i in range(count):
                layout = com_call(layouts, "Item", i)
                name = com_get(layout, "Name")
                layout_name = str(name)
                if layout_name.lower() == "model":
                    continue
                found_paper = True

                # Refresh & set plot device
                plotter_ok = True
                status = ""
                error = ""
                try:
                    try:
                        com_call(layout, "RefreshPlotDeviceInfo")
                    except Exception:
                        pass
                    com_set(layout, "ConfigName", self.pdf_pc3_name)
                except Exception as e:
                    plotter_ok = False
                    status = "Plotter Unavailable"
                    error = str(e)

                # Activate layout
                activated = False
                if plotter_ok:
                    try:
                        com_set(doc, "ActiveLayout", layout)
                        activated = True
                    except Exception:
                        pythoncom.PumpWaitingMessages()
                        time.sleep(0.2)
                        try:
                            com_set(doc, "ActiveLayout", layout)
                            activated = True
                        except Exception as e2:
                            status = "Activate Failed"
                            error = str(e2)

                # Read title block attributes (fast + robust)
                sheet_no = "Not Found"
                sheet_title = "Not Found"
                rev_no = "Not Found"
                rev_date = "Not Found"
                rev_desc = "Not Found"
                rev_by = "Not Found"
                other_attrs = {}

                if activated:
                    try:
                        sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs = read_titleblock_from_active_layout_robust(
                            doc,
                            target_block_names=TARGET_BLOCK_NAMES or None,
                            sheet_top_tags=("FC-E", "TOPNUMBER", "TOP_NUM", "TOP"),
                            sheet_top_prompts=("TOP NUMBER",),
                            sheet_bottom_tags=("442C", "BOTTOMNUMBER", "BOTTOM_NUM", "BOTTOM"),
                            sheet_bottom_prompts=("BOTTOM NUMBER",),
                            title_tag_primary=("TITLE_1", "TITLE_2", "TITLE_3", "TITLE_4", "TITLE_5"),
                            title_prompt_primary=("TITLE_1", "TITLE_2", "TITLE_3", "TITLE_4", "TITLE_5"),
                            title_prompts_alias=("ELECTRICAL",),
                            revision_index_range=REV_INDEX_RANGE,
                            sheetno_sep=SHEETNO_SEPARATOR,
                            title_joiner=TITLE_JOINER,
                        )
                    except Exception as e:
                        if error:
                            error += f" | Title block read: {e}"
                        else:
                            error = f"Title block read: {e}"

                # Prepare PDF path
                pdf_name = f"{dwg_path.stem}__{sanitize_filename(layout_name)}.pdf"
                pdf_path = out_dir / pdf_name
                plot_success = "No"

                # Plot if possible
                if plotter_ok and activated:
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
                            plot_success = "Yes"
                            status = "Plotted"
                        else:
                            status = "Plot Failed"
                            if not error:
                                error = "Plot finished but no file for layout."
                    except Exception as e:
                        status = "Plot Failed"
                        error = (error + " | " if error else "") + str(e)
                else:
                    if not status:
                        status = "Skipped"

                # Build per-layout row (base fields)
                row = {
                    "DWG File": dwg_path.name,
                    "Layout": layout_name,
                    "Sheet Number": sheet_no,
                    "Sheet Name": sheet_title,
                    "Revision Number": rev_no,
                    "Revision Date": rev_date,
                    "Revision Description": rev_desc,
                    "Revision By": rev_by,
                    "Plot Successful": plot_success,
                    "Status": status or "",
                    "Error": error or "",
                    "Individual PDF": str(pdf_path) if (plot_success == "Yes") else ""
                }

                # Add remaining TB attributes as extra columns
                for k, v in other_attrs.items():
                    if k not in row:
                        row[k] = v

                rows.append(row)

            # If no paper layouts, still list this DWG
            if not found_paper:
                rows.append({
                    "DWG File": dwg_path.name,
                    "Layout": "",
                    "Sheet Number": "Not Found",
                    "Sheet Name": "Not Found",
                    "Revision Number": "Not Found",
                    "Revision Date": "Not Found",
                    "Revision Description": "Not Found",
                    "Revision By": "Not Found",
                    "Plot Successful": "No",
                    "Status": "No Paper Layouts",
                    "Error": "",
                    "Individual PDF": ""
                })

        finally:
            if doc is not None:
                self.close_document(doc, save=False)

        return generated


# =========================
# Drawing List CSV writer (dynamic columns)
# =========================

def write_drawing_list_csv(rows: list, output_csv: Path):
    ensure_dir(output_csv.parent)
    base_cols = [
        "DWG File", "Layout",
        "Sheet Number", "Sheet Name",
        "Revision Number", "Revision Date", "Revision Description", "Revision By",
        "Plot Successful", "Status", "Error", "Individual PDF"
    ]
    # collect union of extra headers
    seen = set(base_cols)
    extra_cols = []
    for r in rows:
        for k in r.keys():
            if k not in seen:
                seen.add(k)
                extra_cols.append(k)
    extra_cols_sorted = sorted(extra_cols)
    headers = base_cols + extra_cols_sorted

    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for r in rows:
            w.writerow(r)

    print(f"✓ Drawing List saved: {output_csv.name}")


# =========================
# Main
# =========================

def main():
    print("DWG Copier and PDF Converter + Drawing List (per layout)")
    print("--------------------------------------------------------")

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
    drawing_list_csv = output_dir / f"{sanitize_filename(project_name)}_Drawing List.csv"

    all_pdfs = []
    drawing_rows = []

    print("\nStarting AutoCAD for PDF generation...")
    try:
        with AutoCADPdfConverter(pdf_pc3_name="DWG To PDF.pc3", visible=False) as converter:
            print("\nGenerating Individual PDFs and building Drawing List...")
            total = len(copied)
            for idx, dwg in enumerate(copied, 1):
                print(f"[{idx}/{total}] {dwg.name}")
                generated = converter.convert_individual_and_collect_rows(dwg, individual_pdf_dir, drawing_rows)
                all_pdfs.extend(generated)
    except Exception as e:
        print(f"\nERROR: AutoCAD conversion failed: {e}")
        # still write whatever rows we have
        if drawing_rows:
            write_drawing_list_csv(drawing_rows, drawing_list_csv)
        sys.exit(2)

    # 5) Merge PDFs into single combined PDF
    if all_pdfs:
        print(f"\nMerging {len(all_pdfs)} PDFs into one combined file...")
        # Keep order stable by DWG name then layout
        all_pdfs_sorted = sorted(all_pdfs, key=lambda p: (p.stem.split("__")[0].lower(), p.stem.lower()))
        try:
            merge_pdfs_in_order(all_pdfs_sorted, combined_pdf_path)
        except Exception as e:
            print(f"    ! Failed to merge PDFs: {e}")
    else:
        print("\nNo PDFs generated to merge.")

    # 6) Write Drawing List (CSV)
    write_drawing_list_csv(drawing_rows, drawing_list_csv)

    print("\nDone.")


if __name__ == "__main__":
    main()