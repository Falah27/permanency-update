"""Background worker routines for standalone app."""

import os
import tempfile
import shutil
import time
import traceback
import re
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
    extract_numeric_value,
)


def _normalize_oid_for_match(oid_value):
    """Normalize OID for strict matching by converting segment '.x' to '.1'."""
    oid = str(oid_value).strip() if oid_value is not None else ""
    if not oid:
        return ""
    return re.sub(r'\.[xX](?=\.|$)', '.1', oid)


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
        parent_oid_index = {}  # NEW: Maps OID prefix → list of full OIDs for faster parent lookup
        
        if model_col_value and last_row_private >= 12:
            oid_range = ws_private.Range(ws_private.Cells(12, 6), ws_private.Cells(last_row_private, 6)).Value
            attr_range = ws_private.Range(ws_private.Cells(12, 5), ws_private.Cells(last_row_private, 5)).Value
            value_range = ws_private.Range(ws_private.Cells(12, model_col_value), ws_private.Cells(last_row_private, model_col_value)).Value
            
            if oid_range and attr_range and value_range:
                for i in range(len(oid_range)):
                    oid_row = oid_range[i]
                    attr_row = attr_range[i]
                    value_row = value_range[i]
                    
                    oid = oid_row[0] if oid_row and oid_row[0] else None
                    attr_name = attr_row[0] if attr_row and attr_row[0] else None
                    value = value_row[0] if value_row and value_row[0] else None
                    
                    if oid:
                        oid_key = _normalize_oid_for_match(oid)
                        private_oid_value_map[oid_key] = value if value else ""
                        
                        # NEW: Build parent OID index - split by dots for prefix matching
                        oid_parts = oid_key.split('.')
                        for length in range(len(oid_parts), 0, -1):
                            prefix = '.'.join(oid_parts[:length])
                            if prefix not in parent_oid_index:
                                parent_oid_index[prefix] = []
                            if oid_key not in parent_oid_index[prefix]:
                                parent_oid_index[prefix].append(oid_key)
                        
                        if attr_name:
                            attr_normalized = str(attr_name).strip().lower()
                            mib_oid_to_attr_map[oid_key] = attr_normalized
                    
                    if attr_name:
                        attr_key = str(attr_name).strip().lower()
                        # Only map if we have a value
                        if value is not None and str(value).strip() != "":
                            private_attr_value_map[attr_key] = value
        
        # Build OID sets directly from already-read map (no second COM call needed)
        mib_oids_set = set(private_oid_value_map.keys())
        
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
            check_attr_range = ws_check.Range(ws_check.Cells(start_row, 4), ws_check.Cells(last_row_check, 4)).Value
            check_col_e_range = ws_check.Range(ws_check.Cells(start_row, 5), ws_check.Cells(last_row_check, 5)).Value
            check_oid_range = ws_check.Range(ws_check.Cells(start_row, 6), ws_check.Cells(last_row_check, 6)).Value
            
            if check_attr_range is None:
                check_attr_range = []
            elif not isinstance(check_attr_range[0], tuple):
                check_attr_range = [(check_attr_range,)]
            
            if check_col_e_range is None:
                check_col_e_range = []
            elif not isinstance(check_col_e_range[0], tuple):
                check_col_e_range = [(check_col_e_range,)]
                
            if check_oid_range is None:
                check_oid_range = []
            elif not isinstance(check_oid_range[0], tuple):
                check_oid_range = [(check_oid_range,)]
        
        # Process matching with double validation
        output_data_col_i = []
        output_data_col_j = []
        output_data_rest = []
        match_count = 0
        no_support_count = 0
        
        # Statistics tracking
        value_from_exact_oid = 0
        value_from_parent_oid = 0
        value_from_attr = 0
        value_empty = 0
        
        for idx, row in enumerate(check_oid_range, start=1):
            c_oid_raw = row[0] if row else None
            
            c_attr_name = ""
            if idx <= len(check_attr_range):
                attr_row = check_attr_range[idx - 1]
                c_attr_name = str(attr_row[0]).strip().lower() if attr_row and attr_row[0] else ""
            
            col_e_value = ""
            if idx <= len(check_col_e_range):
                col_e_row = check_col_e_range[idx - 1]
                col_e_value = str(col_e_row[0]).strip() if col_e_row and col_e_row[0] else ""
            
            if not c_oid_raw:
                output_data_col_i.append("")
                output_data_col_j.append("")
                output_data_rest.append([""] * 13)
                continue
            
            c_oid = _normalize_oid_for_match(c_oid_raw)
            
            # Check for 範囲外
            if "範囲外" in col_e_value:
                is_match = False
            else:
                is_match = False
                
                # Strategy 1: OID exact match with attribute validation
                if c_oid in mib_oids_set:
                    if c_oid in mib_oid_to_attr_map:
                        mib_attr = mib_oid_to_attr_map[c_oid]
                        if c_attr_name == mib_attr:
                            is_match = True
                    else:
                        is_match = True
                
                # Strategy 2: Improved parent OID match with progressive prefix checking
                if not is_match:
                    c_oid_parts = c_oid.split('.')
                    
                    # Try all possible parent prefixes (from longest to shortest)
                    for length in range(len(c_oid_parts) - 1, 0, -1):
                        parent_prefix = '.'.join(c_oid_parts[:length])
                        
                        # Check if this prefix exists in MIB
                        if parent_prefix in parent_oid_index:
                            # Get all MIB OIDs that match this prefix
                            matching_oids = parent_oid_index[parent_prefix]
                            
                            # Find exact match or best parent
                            for m_oid in matching_oids:
                                if c_oid.startswith(m_oid + ".") or c_oid == m_oid:
                                    # Check attribute validation if available
                                    if m_oid in mib_oid_to_attr_map:
                                        mib_attr = mib_oid_to_attr_map[m_oid]
                                        if c_attr_name == mib_attr:
                                            is_match = True
                                            break
                                    else:
                                        is_match = True
                                        break
                            
                            if is_match:
                                break
                
                # Strategy 3: Fallback to attribute-only match
                if not is_match and c_attr_name and c_attr_name in private_attr_value_map:
                    is_match = True
            
            # Get model value with IMPROVED matching (only for matched rows)
            model_value = ""
            value_source = ""
            
            if is_match:
                # Strategy 1: Exact OID match
                if c_oid in private_oid_value_map:
                    model_value = private_oid_value_map[c_oid]
                    value_source = "exact_oid"
                    if model_value and str(model_value).strip() != "":
                        value_from_exact_oid += 1
                else:
                    # Strategy 2: Find BEST (longest) parent OID match
                    # Go from longest prefix to shortest — first hit IS the best match
                    c_oid_parts = c_oid.split('.')
                    best_match_value = None
                    
                    for length in range(len(c_oid_parts), 0, -1):
                        prefix = '.'.join(c_oid_parts[:length])
                        if prefix in private_oid_value_map:
                            best_match_value = private_oid_value_map[prefix]
                            break
                    
                    if best_match_value is not None:
                        model_value = best_match_value
                        value_source = "parent_oid"
                        if model_value and str(model_value).strip() != "":
                            value_from_parent_oid += 1
                    else:
                        # Strategy 3: Fallback to attribute name match
                        if c_attr_name and c_attr_name in private_attr_value_map:
                            model_value = private_attr_value_map[c_attr_name]
                            value_source = "attribute"
                            if model_value and str(model_value).strip() != "":
                                value_from_attr += 1
                
                # Convert value to string, extract numeric value, and handle None/empty
                if model_value is None or str(model_value).strip() == "":
                    model_value = ""
                    value_empty += 1
                else:
                    model_value = extract_numeric_value(str(model_value).strip())
            
            # Build output data
            if is_match:
                match_count += 1
                output_data_col_i.append("FactoryDefault")
                output_data_col_j.append(model_value if model_value else "")
                row_data = ["", "", "", "○", "", "", "", "○", "", "", "", "○", ""]
            else:
                no_support_count += 1
                output_data_col_i.append("NoSupport")
                output_data_col_j.append("")
                row_data = ["[NA]", "", "", "-", "[NA]", "", "", "-", "[NA]", "", "", "-", ""]
            
            output_data_rest.append(row_data)
            
            # Progress update
            if idx % 100 == 0:
                progress_pct = 40 + int((idx / total_rows) * 40)
                output_queue.put({"progress": progress_pct, "message": f"⚙️ Processing {idx}/{total_rows}..."})
        
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
            "stats": {
                "match_count": match_count,
                "no_support_count": no_support_count,
                "total_rows": total_rows,
                "value_from_exact_oid": value_from_exact_oid,
                "value_from_parent_oid": value_from_parent_oid,
                "value_from_attr": value_from_attr,
                "value_empty": value_empty
            }
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


