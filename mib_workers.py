"""Background worker routines for standalone app."""

import os
import tempfile
import shutil
import time
import traceback
import pythoncom

from mib_core import (
    ExcelConst,
    create_excel_app,
    init_excel_app,
    close_excel_safely,
    create_dump_sheet,
    inject_template_sheets,
    process_sheet_cutting,
    delete_unused_sheets,
    detect_sheet_names,
    normalize_excel_range_2d,
    build_mode2_lookup_maps,
    process_mode2_matching_rows,
)


def _finalize_excel(excel):
    """Best-effort Excel shutdown for worker threads."""
    if not excel:
        return

    try:
        while excel.Workbooks.Count > 0:
            excel.Workbooks(1).Close(SaveChanges=False)
    except:
        pass

    try:
        excel.Visible = False
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True
        excel.Quit()
    except:
        pass

    try:
        pythoncom.CoUninitialize()
    except:
        pass

def background_worker_mode1(spek_path, selected, template_path, output_queue):
    """Worker thread untuk Mode 1 - Model Selection"""
    try:
        start_time = time.time()
        
        # Setup paths
        temp_dir = tempfile.gettempdir()
        safe_name = "".join(c for c in selected["name"] if c.isalnum() or c == ' ').strip()
        output_filename = f"Checksheet_{safe_name.replace(' ', '_')}.xlsm"
        output_path = os.path.join(temp_dir, output_filename)
        
        shutil.copy(spek_path, output_path)
        
        # Send progress updates
        output_queue.put({"progress": 10, "message": "🔌 Initializing Excel..."})
        excel = init_excel_app()
        
        output_queue.put({"progress": 20, "message": "📁 Opening workbook..."})
        wb = excel.Workbooks.Open(
            os.path.abspath(output_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )
        excel.Calculation = ExcelConst.CALCULATION_MANUAL
        
        output_queue.put({"progress": 30, "message": "📝 Creating dump sheet..."})
        create_dump_sheet(wb)
        
        output_queue.put({"progress": 40, "message": "📑 Injecting templates..."})
        if os.path.exists(template_path):
            inject_template_sheets(wb, template_path, excel)
        
        output_queue.put({"progress": 50, "message": "✂️ Processing sheets..."})
        process_sheet_cutting(wb, selected["col_index"], excel)
        
        output_queue.put({"progress": 80, "message": "🧹 Cleaning..."})
        delete_unused_sheets(wb)
        
        output_queue.put({"progress": 90, "message": "💾 Saving..."})
        # Ensure formulas stay responsive after user edits in Excel (U -> Y).
        excel.Calculation = ExcelConst.CALCULATION_AUTOMATIC
        try:
            excel.CalculateFullRebuild()
        except:
            try:
                excel.Calculate()
            except:
                pass
        wb.Save()
        wb.Close(SaveChanges=False)
        close_excel_safely(excel)
        
        duration = time.time() - start_time
        
        # Send completion
        output_queue.put({
            "status": "success",
            "progress": 100,
            "message": f"✅ Complete! ({duration:.1f}s)",
            "output_path": output_path,
            "duration": duration
        })
        
    except Exception as e:
        output_queue.put({
            "status": "error",
            "message": str(e),
            "traceback": traceback.format_exc()
        })
    finally:
        # Ensure Excel is properly closed even if error occurs
        excel_obj = excel if 'excel' in locals() else None
        _finalize_excel(excel_obj)



def background_worker_mode2(spek_path, checksheet_path, selected_model, output_queue):
    """Worker thread untuk Mode 2 - Model Selection with Double Validation"""
    try:
        start_time = time.time()
        
        # Detect sheet names based on checksheet filename
        checksheet_filename = os.path.basename(checksheet_path)
        checksheet_sheet_name, mib_sheet_name = detect_sheet_names(checksheet_filename)
        
        output_queue.put({"progress": 10, "message": f"🔌 Initializing Excel... (Using {mib_sheet_name})"})
        
        pythoncom.CoInitialize()
        excel = create_excel_app()
        
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        
        output_queue.put({"progress": 20, "message": f"📖 Reading MIB data from {mib_sheet_name}..."})
        
        # Open Spek file
        wb_spek = excel.Workbooks.Open(os.path.abspath(spek_path), ReadOnly=True)
        ws_private = wb_spek.Worksheets(mib_sheet_name)
        
        # Find model column
        model_col_private = None
        model_col_value = None
        empty_count = 0
        for col_idx in range(11, 200, 1):  # Scan all columns
            header_value = ws_private.Cells(10, col_idx).Value
            if not header_value or str(header_value).strip() == "":
                empty_count += 1
                if empty_count >= 20:  # Stop after 20 consecutive empty columns
                    break
                continue
            
            empty_count = 0  # Reset counter
            if str(header_value).strip() == selected_model:
                model_col_private = col_idx
                model_col_value = col_idx + 4
                break
        
        if not model_col_private:
            raise Exception(f"Model '{selected_model}' not found in {mib_sheet_name}")
        
        # Read OID and Attribute mappings
        last_row_private = ws_private.Cells(ws_private.Rows.Count, 6).End(-4162).Row
        private_oid_value_map = {}
        private_attr_value_map = {}
        mib_oid_to_attr_map = {}
        parent_oid_index = {}
        mib_oids_set = set()
        
        if model_col_value and last_row_private >= 12:
            oid_range = ws_private.Range(ws_private.Cells(12, 6), ws_private.Cells(last_row_private, 6)).Value
            attr_range = ws_private.Range(ws_private.Cells(12, 5), ws_private.Cells(last_row_private, 5)).Value
            value_range = ws_private.Range(ws_private.Cells(12, model_col_value), ws_private.Cells(last_row_private, model_col_value)).Value

            lookup_maps = build_mode2_lookup_maps(
                normalize_excel_range_2d(oid_range),
                normalize_excel_range_2d(attr_range),
                normalize_excel_range_2d(value_range),
            )
            private_oid_value_map = lookup_maps["private_oid_value_map"]
            private_attr_value_map = lookup_maps["private_attr_value_map"]
            mib_oid_to_attr_map = lookup_maps["mib_oid_to_attr_map"]
            parent_oid_index = lookup_maps["parent_oid_index"]
            mib_oids_set = lookup_maps["mib_oids_set"]
        
        wb_spek.Close(False)
        
        output_queue.put({"progress": 40, "message": "⚙️ Processing matching..."})
        
        # Open checksheet
        wb_check = excel.Workbooks.Open(os.path.abspath(checksheet_path))
        ws_check = wb_check.Worksheets(checksheet_sheet_name)
        start_row = 10
        last_row_check = ws_check.Cells(ws_check.Rows.Count, 6).End(-4162).Row
        total_rows = last_row_check - start_row + 1
        
        # Clear old content
        if last_row_check >= start_row:
            ws_check.Range(ws_check.Cells(start_row, 9), ws_check.Cells(last_row_check, 22)).ClearContents()
        
        # Read checksheet data
        check_attr_range = []
        check_col_e_range = []
        check_oid_range = []
        
        if last_row_check >= start_row:
            check_attr_range = normalize_excel_range_2d(
                ws_check.Range(ws_check.Cells(start_row, 4), ws_check.Cells(last_row_check, 4)).Value
            )
            check_col_e_range = normalize_excel_range_2d(
                ws_check.Range(ws_check.Cells(start_row, 5), ws_check.Cells(last_row_check, 5)).Value
            )
            check_oid_range = normalize_excel_range_2d(
                ws_check.Range(ws_check.Cells(start_row, 6), ws_check.Cells(last_row_check, 6)).Value
            )

        def _on_match_progress(idx, total):
            if total <= 0:
                return
            progress_pct = 40 + int((idx / total) * 40)
            output_queue.put({"progress": progress_pct, "message": f"⚙️ Processing {idx}/{total}..."})

        match_result = process_mode2_matching_rows(
            check_oid_range,
            check_attr_range,
            check_col_e_range,
            private_oid_value_map,
            private_attr_value_map,
            mib_oid_to_attr_map,
            parent_oid_index,
            mib_oids_set,
            progress_callback=_on_match_progress,
        )

        output_data_col_i = match_result["output_data_col_i"]
        output_data_col_j = match_result["output_data_col_j"]
        output_data_rest = match_result["output_data_rest"]
        stats = match_result["stats"]
        
        output_queue.put({"progress": 85, "message": "💾 Writing data..."})
        
        # Write data to Excel - single bulk write per column group
        if output_data_col_i and last_row_check >= start_row:
            col_i_data = [[val] for val in output_data_col_i]
            ws_check.Range(ws_check.Cells(start_row, 9), ws_check.Cells(last_row_check, 9)).Value = col_i_data
            
            col_j_data = [[val] for val in output_data_col_j]
            ws_check.Range(ws_check.Cells(start_row, 10), ws_check.Cells(last_row_check, 10)).Value = col_j_data
            
            if len(output_data_rest) > 0:
                for col_offset in range(13):
                    col_data = [[row[col_offset]] for row in output_data_rest]
                    col_index = 11 + col_offset
                    ws_check.Range(ws_check.Cells(start_row, col_index), ws_check.Cells(last_row_check, col_index)).Value = col_data
        
        # Save file
        file_name, file_ext = os.path.splitext(checksheet_path)
        new_path_check = f"{file_name}_{selected_model}{file_ext}"
        
        file_format = 52 if file_ext.lower() == '.xlsm' else 51
        
        excel.Calculate()
        excel.ScreenUpdating = False
        
        wb_check.SaveAs(new_path_check, FileFormat=file_format)
        wb_check.Close(False)
        
        # Restore settings before quit
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        excel.DisplayAlerts = True
        excel.Visible = False  # Ensure hidden before quit
        excel.Quit()
        
        duration = time.time() - start_time
        
        output_queue.put({
            "status": "success",
            "progress": 100,
            "message": f"✅ Complete! ({duration:.1f}s)",
            "output_path": new_path_check,
            "duration": duration,
            "stats": stats
        })
        
    except Exception as e:
        output_queue.put({
            "status": "error",
            "message": str(e),
            "traceback": traceback.format_exc()
        })
    finally:
        # Ensure Excel is properly closed even if error occurs
        excel_obj = excel if 'excel' in locals() else None
        _finalize_excel(excel_obj)


