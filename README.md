
# DWG Tool – Batch Plot + Drawing List

Modular AutoCAD batch plotter that:
- Copies DWGs from an input folder (optionally recursive) to an output staging folder
- Plots **each Paper layout** to individual PDFs
- Merges PDFs into one **combined** project PDF
- Extracts per-layout **Drawing List** (Sheet Number, Title, Revision No/Date/Desc/By) from the sheet title block
  and writes a CSV with dynamic extra columns for remaining title-block attributes.

## Structure

```
dwgtool/
├─ dwgtool/
│  ├─ __init__.py
│  ├─ app.py                # Orchestrates the flow
│  ├─ config.py             # Defaults + interactive prompts -> TB_CONFIG dict
│  ├─ titleblock/
│  │  ├─ __init__.py
│  │  └─ scanner.py         # Robust + fast title block scanner (tags/prompts + block-name filter)
│  ├─ cad/
│  │  ├─ __init__.py
│  │  └─ plotter.py         # AutoCAD COM (retry layer) + per-layout plotting + row building
│  ├─ pdf/
│  │  ├─ __init__.py
│  │  └─ merge.py           # PDF merging (PyPDF2)
│  └─ io/
│     ├─ __init__.py
│     └─ files.py           # filesystem utils, copying, CSV writer
└─ run.py                   # Entry point (good for PyInstaller)
```

## Requirements

- Windows, AutoCAD installed (DWG To PDF.pc3 available)
- Python 3.9+ (64-bit recommended)

Install in your environment:
```bash
pip install pywin32 PyPDF2
```

## Usage

From the project root (where `run.py` is):
```bash
python run.py
```

The app will prompt you to:
1. Pick input folder and output folder
2. Enter project name (default = input folder name)
3. Choose recursive scan (Y/N)
4. Confirm/override **Title Block Configuration** (block names, tags, prompts, joiners)

## Packaging to EXE (optional)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed run.py
```

The executable will be in `dist/run.exe`.

## Notes

- The title block scanner matches by **TagString** and **PromptString**, and picks the **highest revision index** (R0..R9) that has a `NO` value, including DATE/DESC/BY columns.
- Remaining title-block attributes from the **chosen** sheet block are added as extra CSV columns (raw tag names as headers).
- If AutoCAD COM sometimes rejects calls while busy, you can try setting `visible=True` in the `AutoCADPdfConverter` for debugging runs.
