import win32com.client as win32
import pythoncom
import os
import time
from datetime import datetime

# ==========================================
# FUNGSI DETEKSI MODEL (C-CODE)
# ==========================================

def detect_models_from_spek(file_spek):
    """
    Deteksi model/C-Code yang tersedia di file Spek
    Returns: list of dict [{"name": "CL86", "col_index": 11}, ...]
    """
    print("\n🔍 Mendeteksi model yang tersedia di file Spek...")
    
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
    except:
        excel = win32.Dispatch("Excel.Application")
    
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    
    models = []
    
    try:
        wb = excel.Workbooks.Open(os.path.abspath(file_spek), ReadOnly=True)
        ws = wb.Worksheets("Public MIB")
        
        # Baca header di row 11, mulai dari kolom K (11) dengan step 5
        # Kolom K(11), P(16), U(21), Z(26), dst
        COL_START = 11
        COL_END = 200
        COL_STEP = 5
        ROW_HEADER = 11
        
        for col_idx in range(COL_START, COL_END, COL_STEP):
            header_value = ws.Cells(ROW_HEADER, col_idx).Value
            if not header_value or str(header_value).strip() == "":
                break
            
            model_name = str(header_value).strip()
            models.append({
                "name": model_name,
                "col_index": col_idx,
                "col_atchi": col_idx + 4  # Kolom 値 (4 kolom ke kanan dari header C-Code)
            })
        
        wb.Close(False)
        print(f"   ✓ Ditemukan {len(models)} model: {[m['name'] for m in models]}")
        
    except Exception as e:
        print(f"   ❌ Error saat deteksi model: {e}")
    finally:
        try:
            excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()  # Cleanup COM
    
    return models

# ==========================================
# FUNGSI UTAMA GENERATE CHECKSHEET
# ==========================================

def generate_checksheet(file_mib, file_check, selected_model_name):
    start_time = time.time()
    print("\n" + "="*70)
    print(f"🚀 MEMULAI PROSES OTOMATISASI CHECKSHEET")
    print(f"📅 Waktu Mulai: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"🎯 Model Terpilih: {selected_model_name}")
    print("="*70)
    
    # Initialize COM for this thread (required for Streamlit/multi-threading)
    pythoncom.CoInitialize()
    
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
    except:
        excel = win32.Dispatch("Excel.Application")
    
    excel.Visible = False       
    excel.DisplayAlerts = False 
    excel.EnableEvents = False  # Matikan event agar macro bawaan tidak jalan otomatis
    
    try:
        path_mib = os.path.abspath(file_mib)
        path_check = os.path.abspath(file_check)
        
        # ===== LANGKAH 1: DETEKSI MODEL DARI FILE SPEK =====
        print("\n🔍 [1/6] Mendeteksi model dari file Spek...")
        step_start = time.time()
        
        # Buka file Spek
        wb_spek = excel.Workbooks.Open(path_mib, ReadOnly=True)
        
        # Ambil nilai dari sheet "EPSON Private MIB" (karena OID-nya match dengan checksheet)
        ws_private = wb_spek.Worksheets("EPSON Private MIB")
        
        # Cari kolom model yang dipilih di Private MIB
        model_col_private = None
        model_col_value = None  # Kolom nilai actual (kolom +4 dari header, yaitu kolom '値')
        for col_idx in range(11, 200, 5):
            header_value = ws_private.Cells(11, col_idx).Value
            if header_value and str(header_value).strip() == selected_model_name:
                model_col_private = col_idx
                model_col_value = col_idx + 4  # Ambil dari kolom ke-5 (offset +4), yaitu kolom '値'
                print(f"   ✓ Model '{selected_model_name}' ditemukan di Private MIB kolom {col_idx}")
                print(f"   ✓ Nilai akan diambil dari kolom {model_col_value} (kolom '値' dengan nilai actual)")
                break
        
        if not model_col_private:
            print(f"   ⚠️ Model '{selected_model_name}' tidak ditemukan di Private MIB")
        
        # Baca OID, Attribute Name, dan nilai dari Private MIB untuk mapping
        last_row_private = ws_private.Cells(ws_private.Rows.Count, 6).End(-4162).Row
        private_oid_value_map = {}  # Map OID ke nilai model
        private_attr_value_map = {}  # Map Attribute Name ke nilai model
        mib_oid_to_attr_map = {}  # Map OID ke Attribute Name (untuk validasi)
        
        if model_col_value and last_row_private >= 12:
            # Baca OID (kolom F/6), Attribute Name (kolom E/5), dan nilai model (kolom O) sekaligus
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
                        # Map OID ke nilai
                        private_oid_value_map[str(oid).strip()] = value if value else ""
                        
                        # Map OID ke Attribute Name (untuk double validation)
                        if attr_name:
                            mib_oid_to_attr_map[str(oid).strip()] = str(attr_name).strip().lower()
                    
                    if attr_name:
                        # Map Attribute Name ke nilai (untuk fallback matching)
                        attr_key = str(attr_name).strip().lower()  # Case-insensitive
                        private_attr_value_map[attr_key] = value if value else ""
            
            print(f"   ✓ Berhasil mapping {len(private_oid_value_map)} OID dan {len(private_attr_value_map)} Attribute dari Private MIB ({time.time()-step_start:.2f}s)")
        
        # ===== LANGKAH 2: BACA FILE MIB (PRIVATE) UNTUK MATCHING =====
        print("\n📖 [2/6] Membaca OID dari MIB Private untuk matching...")
        step_start = time.time()
        ws_mib = ws_private
        print(f"   ✓ Sheet Private MIB ready ({time.time()-step_start:.2f}s)")
        
        # ===== LANGKAH 3: BACA SEMUA OID DARI MIB SEKALIGUS (OPTIMASI!) =====
        print("\n🔍 [3/6] Membaca data OID dari MIB Private...")
        step_start = time.time()
        last_row_mib = ws_mib.Cells(ws_mib.Rows.Count, 6).End(-4162).Row
        
        # OPTIMASI: Baca seluruh kolom OID sekaligus (tidak satu-satu)
        if last_row_mib >= 12:
            mib_range = ws_mib.Range(ws_mib.Cells(12, 6), ws_mib.Cells(last_row_mib, 6)).Value
            # Ubah jadi list dan filter yang kosong
            if mib_range:
                mib_oids_list = [str(row[0]).strip() for row in mib_range if row[0] is not None]
            else:
                mib_oids_list = []
        else:
            mib_oids_list = []
        
        # OPTIMASI: Gunakan set untuk pencarian cepat O(1) vs O(n)
        mib_oids_set = set(mib_oids_list)
        
        # Buat mapping untuk parent OID (untuk cek prefix dengan cepat)
        mib_oids_parents = set()
        for oid in mib_oids_list:
            mib_oids_parents.add(oid)
            # Tambahkan semua parent path
            parts = oid.split('.')
            for i in range(1, len(parts)):
                mib_oids_parents.add('.'.join(parts[:i]))
        
        wb_spek.Close(False)
        print(f"   ✓ Ditemukan {len(mib_oids_list)} OID dari MIB ({time.time()-step_start:.2f}s)")
        
        # ===== LANGKAH 4: BUKA FILE CHECKSHEET =====
        print("\n📋 [4/6] Membuka file Checksheet...")
        step_start = time.time()
        wb_check = excel.Workbooks.Open(path_check)
        ws_check = wb_check.Worksheets("Printer Private")
        start_row = 10  # Row 9 adalah header, data mulai row 10
        last_row_check = ws_check.Cells(ws_check.Rows.Count, 6).End(-4162).Row
        total_rows = last_row_check - start_row + 1
        print(f"   ✓ File Checksheet dibuka: {os.path.basename(path_check)} ({time.time()-step_start:.2f}s)")
        print(f"   ℹ️  Total baris data yang akan diproses: {total_rows}")
        
        # Clear konten lama dulu (kolom I sampai V)
        if last_row_check >= start_row:
            ws_check.Range(ws_check.Cells(start_row, 9), ws_check.Cells(last_row_check, 22)).ClearContents()
            print(f"   ✓ Cleared range I{start_row}:V{last_row_check}")
        
        # ===== LANGKAH 5: PROSES MATCHING OID (OPTIMASI BESAR!) =====
        print("\n⚙️  [5/6] Memproses matching OID dan Attribute Name...")
        step_start = time.time()
        
        # OPTIMASI: Baca semua data dari checksheet sekaligus
        if last_row_check >= start_row:
            # Baca kolom D (Attribute), kolom E (範囲外), dan kolom F (OID)
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
        else:
            check_attr_range = []
            check_col_e_range = []
            check_oid_range = []
        
        # OPTIMASI: Siapkan data output dalam memory (tidak tulis satu-satu ke Excel)
        output_data_col_i = []  # Kolom I: FactoryDefault / NoSupport
        output_data_col_j = []  # Kolom J: Nilai dari model
        output_data_rest = []   # Kolom K-V: data lainnya
        match_count = 0
        no_support_count = 0
        match_by_oid_exact_count = 0  # OID exact match (tanpa peduli attribute)
        match_by_oid_parent_count = 0  # OID parent match (tanpa peduli attribute)
        match_by_attr_only_count = 0  # Match hanya by attribute (OID tidak match)
        oid_match_attr_validated_count = 0  # OID match DAN attribute validated ✓
        oid_rejected_count = 0  # OID match tapi attribute tidak match (false positive)
        
        found_values_count = 0  # Track berapa OID yang dapat nilai dari model
        
        for idx, row in enumerate(check_oid_range, start=1):
            c_oid_raw = row[0] if row else None
            
            # Baca attribute name dari kolom D
            c_attr_name = ""
            if idx <= len(check_attr_range):
                attr_row = check_attr_range[idx - 1]
                c_attr_name = str(attr_row[0]).strip().lower() if attr_row and attr_row[0] else ""
            
            # Baca nilai kolom E untuk cek 範囲外
            col_e_value = ""
            if idx <= len(check_col_e_range):
                col_e_row = check_col_e_range[idx - 1]
                col_e_value = str(col_e_row[0]).strip() if col_e_row and col_e_row[0] else ""
            
            if not c_oid_raw:
                # Baris kosong, biarkan kosong juga
                output_data_col_i.append("")
                output_data_col_j.append("")
                output_data_rest.append([""] * 13)  # K sampai V = 13 kolom
                continue
            
            c_oid = str(c_oid_raw).strip()
            
            # Cek dulu apakah kolom E berisi "範囲外" -> langsung NoSupport
            if "範囲外" in col_e_value:
                is_match = False  # Paksa jadi NoSupport
                match_method = None
                matched_oid = None
            else:
                # STRATEGI MATCHING dengan DOUBLE VALIDATION:
                # 1. Coba match by OID (exact atau parent)
                # 2. Validasi: OID match DAN Attribute Name juga harus match
                # 3. Jika validasi gagal, coba match by Attribute Name saja
                
                is_match = False
                match_method = None
                matched_oid = None
                
                # MATCHING BY OID (dengan validasi attribute)
                if c_oid in mib_oids_set:
                    # OID exact match - validasi attribute name
                    if c_oid in mib_oid_to_attr_map:
                        mib_attr = mib_oid_to_attr_map[c_oid]
                        if c_attr_name == mib_attr:
                            is_match = True
                            match_method = "oid_exact"
                            matched_oid = c_oid
                        else:
                            oid_rejected_count += 1
                    else:
                        # OID match tapi tidak ada attribute di map, accept saja
                        is_match = True
                        match_method = "oid_exact"
                        matched_oid = c_oid
                else:
                    # Cek apakah c_oid adalah child dari salah satu MIB OID (parent match)
                    for m_oid in mib_oids_list:
                        if c_oid.startswith(m_oid + "."):
                            # Parent OID match - validasi attribute name
                            if m_oid in mib_oid_to_attr_map:
                                mib_attr = mib_oid_to_attr_map[m_oid]
                                if c_attr_name == mib_attr:
                                    is_match = True
                                    match_method = "oid_parent"
                                    matched_oid = m_oid
                                    break
                                else:
                                    oid_rejected_count += 1
                            else:
                                # Parent match tapi tidak ada attribute, accept saja
                                is_match = True
                                match_method = "oid_parent"
                                matched_oid = m_oid
                                break
                
                # MATCHING BY ATTRIBUTE NAME (fallback jika OID tidak match atau ditolak)
                if not is_match and c_attr_name and c_attr_name in private_attr_value_map:
                    is_match = True
                    match_method = "attribute"
            
            # Ambil nilai dari model yang dipilih (kolom J) - dari Private MIB
            # Strategi: coba by OID dulu (exact atau parent), kalau gagal coba by Attribute Name
            model_value = ""
            
            # 1. Coba exact match by OID
            if c_oid in private_oid_value_map:
                model_value = private_oid_value_map[c_oid]
                found_values_count += 1
            else:
                # 2. Cari parent OID yang paling spesifik (longest match)
                # Contoh: .28 akan cocok dengan .1 dan ambil nilai yang sama dari .1
                best_parent = None
                best_parent_value = None
                for parent_oid, parent_value in private_oid_value_map.items():
                    if c_oid.startswith(parent_oid + "."):
                        # Ambil parent yang paling panjang (paling spesifik)
                        if best_parent is None or len(parent_oid) > len(best_parent):
                            best_parent = parent_oid
                            best_parent_value = parent_value
                
                if best_parent_value is not None:
                    model_value = best_parent_value
                    found_values_count += 1
                else:
                    # 3. Fallback: coba by Attribute Name
                    if c_attr_name and c_attr_name in private_attr_value_map:
                        model_value = private_attr_value_map[c_attr_name]
                        found_values_count += 1
            
            # Siapkan data sesuai hasil matching
            if is_match:
                match_count += 1
                
                # Track matching method dengan detail
                if match_method == "attribute":
                    match_by_attr_only_count += 1
                elif match_method == "oid_exact":
                    match_by_oid_exact_count += 1
                    oid_match_attr_validated_count += 1  # Exact match pasti validated
                elif match_method == "oid_parent":
                    match_by_oid_parent_count += 1
                    oid_match_attr_validated_count += 1  # Parent match yang lolos validasi
                
                # Kolom I: FactoryDefault
                output_data_col_i.append("FactoryDefault")
                # Kolom J: Nilai dari model
                output_data_col_j.append(model_value if model_value else "")
                # Kolom K(11), L-V: data evaluasi
                row_data = [
                    "",                # Kolom K (11)
                    "",                # Kolom L (12)
                    "",                # Kolom M (13)
                    "○",               # Kolom N (14)
                    "",                # Kolom O (15)
                    "",                # Kolom P (16)
                    "",                # Kolom Q (17)
                    "○",               # Kolom R (18)
                    "",                # Kolom S (19)
                    "",                # Kolom T (20)
                    "",                # Kolom U (21)
                    "○"                # Kolom V (22)
                ]
            else:
                no_support_count += 1
                # Kolom I: NoSupport
                output_data_col_i.append("NoSupport")
                # Kolom J: kosong untuk NoSupport
                output_data_col_j.append("")
                # Kolom K-V
                row_data = [
                    "[NA]",       # Kolom K (11)
                    "",           # Kolom L (12)
                    "",           # Kolom M (13)
                    "-",          # Kolom N (14)
                    "[NA]",       # Kolom O (15)
                    "",           # Kolom P (16)
                    "",           # Kolom Q (17)
                    "-",          # Kolom R (18)
                    "[NA]",       # Kolom S (19)
                    "",           # Kolom T (20)
                    "",           # Kolom U (21)
                    "-"           # Kolom V (22)
                ]
            
            output_data_rest.append(row_data)
            
            # Progress indicator setiap 100 baris
            if idx % 100 == 0:
                print(f"   ⏳ Progress: {idx}/{total_rows} baris ({idx*100//total_rows}%)")
        
        print(f"   ✓ Matching selesai: {match_count} match, {no_support_count} no-support ({time.time()-step_start:.2f}s)")
        print(f"   ℹ️  Detail Matching:")
        print(f"      - OID Exact Match        : {match_by_oid_exact_count}")
        print(f"      - OID Parent Match       : {match_by_oid_parent_count}")
        print(f"      - Attribute Only Match   : {match_by_attr_only_count}")
        print(f"      - OID+Attr Validated ✓   : {oid_match_attr_validated_count}")
        print(f"   ⚠️  OID rejected (attr ≠)   : {oid_rejected_count}")
        print(f"   ℹ️  OID dengan nilai model   : {found_values_count}/{len(check_oid_range)}")
        print(f"   ℹ️  Data siap ditulis: {len(output_data_col_i)} baris")
        
        # Debug: Tampilkan sampel data
        if len(output_data_rest) > 0:
            print(f"   ℹ️  Sampel data row pertama (kolom K-V): {output_data_rest[0]}")
            print(f"   ℹ️  Jumlah kolom per row: {len(output_data_rest[0])}")
        
        # OPTIMASI: Tulis semua data sekaligus ke Excel (BULK WRITE!)
        print("\n💾 [6/6] Menulis hasil ke Checksheet...")
        step_start = time.time()
        
        if output_data_col_i and last_row_check >= start_row:
            # Tulis kolom I (FactoryDefault/NoSupport)
            col_i_data = [[val] for val in output_data_col_i]
            ws_check.Range(ws_check.Cells(start_row, 9), ws_check.Cells(last_row_check, 9)).Value = col_i_data
            print(f"   ✓ Kolom I tertulis: {len(col_i_data)} baris")
            
            # Tulis kolom J (Nilai dari model)
            col_j_data = [[val] for val in output_data_col_j]
            ws_check.Range(ws_check.Cells(start_row, 10), ws_check.Cells(last_row_check, 10)).Value = col_j_data
            print(f"   ✓ Kolom J tertulis: {len(col_j_data)} baris")
            
            # Tulis kolom K sampai V (data evaluasi lainnya)
            # PERBAIKAN: Tulis kolom per kolom untuk memastikan berhasil
            print(f"   ℹ️  Akan menulis {len(output_data_rest)} baris x 12 kolom (K-V)")
            
            try:
                # Extract data per kolom dari output_data_rest
                if len(output_data_rest) > 0:
                    for col_offset in range(12):  # 12 kolom (K sampai V)
                        col_data = [[row[col_offset]] for row in output_data_rest]
                        col_index = 11 + col_offset  # K=11, L=12, ..., V=22
                        ws_check.Range(ws_check.Cells(start_row, col_index), ws_check.Cells(last_row_check, col_index)).Value = col_data
                    print(f"   ✓ Kolom K-V tertulis: {len(output_data_rest)} baris x 12 kolom")
                else:
                    print(f"   ⚠️  Tidak ada data untuk ditulis ke kolom K-V")
            except Exception as e:
                print(f"   ❌ Error menulis kolom K-V: {e}")
                # Coba tulis satu per satu untuk debug
                print(f"   ⚠️  Mencoba tulis satu per satu...")
                for i, row_data in enumerate(output_data_rest):
                    try:
                        row_num = start_row + i
                        ws_check.Range(ws_check.Cells(row_num, 11), ws_check.Cells(row_num, 22)).Value = [row_data]
                    except Exception as e2:
                        print(f"   ❌ Error di baris {row_num}: {e2}")
            
        print(f"   ✓ Data berhasil ditulis ke Excel ({time.time()-step_start:.2f}s)")
        
        # ===== SIMPAN FILE =====
        print("\n💾 Menyimpan file hasil...")
        step_start = time.time()
        file_name, file_ext = os.path.splitext(path_check)
        new_path_check = f"{file_name}_{selected_model_name}_aftergenerate{file_ext}"
        
        if file_ext.lower() == '.xlsm':
            file_format = 52  # xlOpenXMLWorkbookMacroEnabled
        else:
            file_format = 51  # xlOpenXMLWorkbook (.xlsx)
        
        # PENTING: Pastikan data di-commit ke Excel sebelum save
        excel.Calculate()  # Force calculation
        excel.ScreenUpdating = False
        
        try:
            wb_check.SaveAs(new_path_check, FileFormat=file_format)
            wb_check.Close(False)
            print(f"   ✓ File tersimpan: {os.path.basename(new_path_check)} ({time.time()-step_start:.2f}s)")
        except Exception as e:
            print(f"   ❌ Error saat save: {e}")
            wb_check.Close(False)
        print(f"   ✓ File tersimpan ({time.time()-step_start:.2f}s)")
        
        # ===== TUTUP EXCEL =====
        print("\n🔒 Menutup Excel...")
        try:
            excel.Quit()
            print(f"   ✓ Excel ditutup")
        except Exception as e:
            print(f"   ⚠️ Error saat tutup Excel: {e}")
        
        # ===== SELESAI =====
        total_time = time.time() - start_time
        print("\n" + "="*70)
        print("✅ PROSES SELESAI!")
        print("="*70)
        print(f"📊 Ringkasan:")
        print(f"   • Total OID di MIB      : {len(mib_oids_list)}")
        print(f"   • Total baris diproses  : {total_rows}")
        print(f"   • OID match (support)   : {match_count}")
        print(f"     - OID Exact           : {match_by_oid_exact_count}")
        print(f"     - OID Parent          : {match_by_oid_parent_count}")
        print(f"     - Attribute Only      : {match_by_attr_only_count}")
        print(f"     - OID+Attr Validated✓ : {oid_match_attr_validated_count}")
        print(f"   • OID rejected (attr ≠) : {oid_rejected_count}")
        print(f"   • OID no-support        : {no_support_count}")
        print(f"   • Model dipilih         : {selected_model_name}")
        print(f"   • Waktu total           : {total_time:.2f} detik")
        print(f"   • File output           : {os.path.basename(new_path_check)}")
        print("="*70 + "\n")
        
        return new_path_check  # Return path untuk Streamlit integration
        
    except Exception as e:
        print(f"\n❌ TERJADI ERROR: {e}")
        import traceback
        traceback.print_exc()
        try:
            excel.Quit()
        except:
            pass
        raise  # Re-raise untuk error handling di Streamlit
    finally:
        pythoncom.CoUninitialize()  # Cleanup COM

# ==========================================
# MAIN EXECUTION
# ==========================================

if __name__ == "__main__":
    print("\n" + "="*70)
    print("🖨️  MIB CHECKSHEET GENERATOR - With Model Selection")
    print("="*70)
    
    # File paths
    file_spek = "PRZ-24-00355615_RevH.xlsm"
    file_checksheet = "02_初期値評価仕様書_epPrt.xlsm"
    
    # Step 1: Deteksi model yang tersedia
    models = detect_models_from_spek(file_spek)
    
    if not models:
        print("\n❌ Tidak ada model yang ditemukan di file Spek!")
        exit(1)
    
    # Step 2: Tampilkan pilihan model
    print("\n📋 Model yang tersedia:")
    for i, model in enumerate(models, start=1):
        print(f"   {i}. {model['name']}")
    
    # Step 3: User memilih model
    print("\n🎯 Pilih model (masukkan nomor):")
    try:
        pilihan = int(input("   Pilihan: "))
        if pilihan < 1 or pilihan > len(models):
            print(f"❌ Pilihan tidak valid! Harus antara 1-{len(models)}")
            exit(1)
        
        selected_model = models[pilihan - 1]
        print(f"\n✅ Model terpilih: {selected_model['name']}")
        
    except ValueError:
        print("❌ Input tidak valid! Harus berupa angka")
        exit(1)
    
    # Step 4: Generate checksheet
    input("\nTekan Enter untuk mulai generate checksheet...")
    generate_checksheet(file_spek, file_checksheet, selected_model['name'])