import time
from pathlib import Path
from typing import List, Dict

import pythoncom
import win32com.client

from ..titleblock.scanner import read_titleblock_from_active_layout_robust
from ..io.files import sanitize_filename

# COM retry helpers
RPC_E_CALL_REJECTED = -2147418111


def com_retry_call(func, *args, retries=50, delay=0.15, desc="", **kwargs):
    last_exc = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pythoncom.com_error as e:
            hresult = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
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


class AutoCADPdfConverter:
    def __init__(self, pdf_pc3_name: str = "DWG To PDF.pc3", visible: bool = False, tb_config: Dict = None):
        self.pdf_pc3_name = pdf_pc3_name
        self.visible = visible
        self.tb_config = tb_config or {}
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
            raise RuntimeError(
                "AutoCAD COM interface not available. Ensure AutoCAD is installed and can launch without modal dialogs."
            ) from e
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def open_document(self, dwg_path: Path):
        docs = com_get(self.acad, "Documents")
        doc = com_call(docs, "Open", str(dwg_path))
        for _ in range(100):
            try:
                _ = com_get(doc, "FullName")
                break
            except Exception:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.1)
        return doc

    def close_document(self, doc, save: bool = False):
        try:
            com_call(doc, "Close", bool(save))
        except Exception:
            pass

    def convert_individual_and_collect_rows(self, dwg_path: Path, out_dir: Path, rows: List[Dict]) -> List[Path]:
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

                # Title block read (only once per layout)
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
                            com_get,
                            com_call,
                            self.tb_config,
                        )
                    except Exception as e:
                        if error:
                            error += f" | Title block read: {e}"
                        else:
                            error = f"Title block read: {e}"

                # Plot
                pdf_name = f"{dwg_path.stem}__{sanitize_filename(layout_name)}.pdf"
                pdf_path = out_dir / pdf_name
                plot_success = "No"

                if plotter_ok and activated:
                    try:
                        plot = com_get(doc, "Plot")
                        com_call(plot, "PlotToFile", str(pdf_path))
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

                # Build CSV row
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
                for k, v in other_attrs.items():
                    if k not in row:
                        row[k] = v
                rows.append(row)

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