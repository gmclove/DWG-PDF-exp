import re
import shutil
from pathlib import Path
from typing import List, Dict

try:
    import tkinter as tk
    from tkinter import filedialog
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False


def select_directory_gui(prompt_title: str) -> str:
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


def list_dwg_files(input_dir: Path, recursive: bool = True) -> List[Path]:
    pattern = "**/*.dwg" if recursive else "*.dwg"
    return [p for p in input_dir.glob(pattern) if p.is_file() and not p.name.startswith("~")]


def copy_dwg_files(dwg_files: List[Path], dest_dir: Path) -> List[Path]:
    ensure_dir(dest_dir)
    copied = []
    for src in dwg_files:
        dest = dest_dir / src.name
        i = 1
        while dest.exists():
            stem, ext = src.stem, src.suffix
            dest = dest_dir / f"{stem} ({i}){ext}"
            i += 1
        shutil.copy2(src, dest)
        copied.append(dest)
    return copied


def sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)


def write_drawing_list_csv(rows: List[Dict], output_csv: Path):
    ensure_dir(output_csv.parent)
    base_cols = [
        "DWG File", "Layout",
        "Sheet Number", "Sheet Name",
        "Revision Number", "Revision Date", "Revision Description", "Revision By",
        "Plot Successful", "Status", "Error", "Individual PDF"
    ]
    seen = set(base_cols)
    extra_cols = []
    for r in rows:
        for k in r.keys():
            if k not in seen:
                seen.add(k)
                extra_cols.append(k)
    extra_cols_sorted = sorted(extra_cols)
    headers = base_cols + extra_cols_sorted

    import csv
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"✓ Drawing List saved: {output_csv.name}")