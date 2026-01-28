
    import sys
    from pathlib import Path

    from .config import prompt_for_titleblock_config
    from .io.files import select_directory_gui, ensure_dir, list_dwg_files, copy_dwg_files, write_drawing_list_csv, sanitize_filename
    from .cad.plotter import AutoCADPdfConverter
    from .pdf.merge import merge_pdfs_in_order

    def main():
        print("DWG Copier and PDF Converter + Drawing List (per layout)")
        print("--------------------------------------------------------")

        input_dir = Path(select_directory_gui("Select INPUT folder (DWG source)")).resolve()
        output_dir = Path(select_directory_gui("Select OUTPUT folder (PDF target)")).resolve()

        default_project_name = input_dir.name
        project_name = input(f"
Enter Project Name [{default_project_name}]: ").strip() or default_project_name

        recursive_input = input("
Include subfolders when searching for DWG files? (Y/N): ").strip().lower()
        recursive = recursive_input.startswith('y')

        print("
Configure Title Block Settings")
        TB_CONFIG = prompt_for_titleblock_config()

        ensure_dir(output_dir)
        dwg_out_dir = output_dir / "DWG"
        individual_pdf_dir = output_dir / "Individual PDFs"
        ensure_dir(dwg_out_dir)
        ensure_dir(individual_pdf_dir)

        print(f"
Scanning for DWGs in: {input_dir}")
        dwg_files = list_dwg_files(input_dir, recursive)
        if not dwg_files:
            print("No DWG files found. Exiting.")
            sys.exit(0)
        print(f"Found {len(dwg_files)} DWG file(s).")

        print(f"
Copying DWGs to: {dwg_out_dir}")
        copied = copy_dwg_files(dwg_files, dwg_out_dir)
        print(f"Copied {len(copied)} DWG file(s).")

        # Dependencies check for merging happens inside merge function
        combined_pdf_path = output_dir / f"{sanitize_filename(project_name)}_Combined.pdf"
        drawing_list_csv = output_dir / f"{sanitize_filename(project_name)}_Drawing List.csv"

        all_pdfs = []
        drawing_rows = []

        try:
            with AutoCADPdfConverter(pdf_pc3_name="DWG To PDF.pc3", visible=False, tb_config=TB_CONFIG) as converter:
                print("
Generating Individual PDFs and building Drawing List...")
                total = len(copied)
                for idx, dwg in enumerate(copied, 1):
                    print(f"[{idx}/{total}] {dwg.name}")
                    generated = converter.convert_individual_and_collect_rows(dwg, individual_pdf_dir, drawing_rows)
                    all_pdfs.extend(generated)
        except Exception as e:
            print(f"
ERROR: AutoCAD conversion failed: {e}")
            if drawing_rows:
                write_drawing_list_csv(drawing_rows, drawing_list_csv)
            sys.exit(2)

        if all_pdfs:
            print(f"
Merging {len(all_pdfs)} PDFs into one combined file...")
            all_pdfs_sorted = sorted(all_pdfs, key=lambda p: (p.stem.split("__")[0].lower(), p.stem.lower()))
            try:
                merge_pdfs_in_order(all_pdfs_sorted, combined_pdf_path)
            except Exception as e:
                print(f"    ! Failed to merge PDFs: {e}")
        else:
            print("
No PDFs generated to merge.")

        write_drawing_list_csv(drawing_rows, drawing_list_csv)
        print("
Done.")

    if __name__ == "__main__":
        main()
