import win32com.client as win32
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
        excel.Quit()
    
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
        for col_idx in range(11, 200, 5):
            header_value = ws_private.Cells(11, col_idx).Value
            if header_value and str(header_value).strip() == selected_model_name:
                model_col_private = col_idx
                print(f"   ✓ Model '{selected_model_name}' ditemukan di Private MIB kolom {col_idx}")
                break
        
        if not model_col_private:
            print(f"   ⚠️ Model '{selected_model_name}' tidak ditemukan di Private MIB")
        
        # Baca OID dan nilai dari Private MIB untuk mapping
        last_row_private = ws_private.Cells(ws_private.Rows.Count, 6).End(-4162).Row
        private_oid_value_map = {}  # Map OID ke nilai model (● atau -)
        
        if model_col_private and last_row_private >= 12:
            # Baca OID (kolom F/6) dan nilai model sekaligus
            oid_range = ws_private.Range(ws_private.Cells(12, 6), ws_private.Cells(last_row_private, 6)).Value
            value_range = ws_private.Range(ws_private.Cells(12, model_col_private), ws_private.Cells(last_row_private, model_col_private)).Value
            
            if oid_range and value_range:
                for i in range(len(oid_range)):
                    oid_row = oid_range[i]
                    value_row = value_range[i]
                    
                    oid = oid_row[0] if oid_row and oid_row[0] else None
                    value = value_row[0] if value_row and value_row[0] else None
                    
                    if oid:
                        private_oid_value_map[str(oid).strip()] = value if value else ""
            
            print(f"   ✓ Berhasil mapping {len(private_oid_value_map)} OID dari Private MIB ({time.time()-step_start:.2f}s)")
        
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
        print("\n⚙️  [5/6] Memproses matching OID dan menyiapkan data...")
        step_start = time.time()
        
        # OPTIMASI: Baca semua OID checksheet sekaligus
        if last_row_check >= start_row:
            check_oid_range = ws_check.Range(ws_check.Cells(start_row, 6), ws_check.Cells(last_row_check, 6)).Value
            if check_oid_range is None:
                check_oid_range = []
            elif not isinstance(check_oid_range[0], tuple):
                check_oid_range = [(check_oid_range,)]
        else:
            check_oid_range = []
        
        # OPTIMASI: Siapkan data output dalam memory (tidak tulis satu-satu ke Excel)
        output_data_col_i = []  # Kolom I: FactoryDefault / NoSupport
        output_data_col_j = []  # Kolom J: Nilai dari model
        output_data_rest = []   # Kolom K-V: data lainnya
        match_count = 0
        no_support_count = 0
        
        found_values_count = 0  # Track berapa OID yang dapat nilai dari model
        
        for idx, row in enumerate(check_oid_range, start=1):
            c_oid_raw = row[0] if row else None
            
            if not c_oid_raw:
                # Baris kosong, biarkan kosong juga
                output_data_col_i.append("")
                output_data_col_j.append("")
                output_data_rest.append([""] * 13)  # K sampai V = 13 kolom
                continue
            
            c_oid = str(c_oid_raw).strip()
            
            # OPTIMASI: Cek matching dengan set (lebih cepat!)
            is_match = False
            if c_oid in mib_oids_set:
                is_match = True
            else:
                # Cek apakah c_oid adalah child dari salah satu MIB OID
                for m_oid in mib_oids_list:
                    if c_oid.startswith(m_oid + "."):
                        is_match = True
                        break
            
            # Ambil nilai dari model yang dipilih (kolom J) - dari Private MIB
            # Cari exact match dulu, kalau tidak ada cari parent OID
            model_value = ""
            if c_oid in private_oid_value_map:
                model_value = private_oid_value_map[c_oid]
                found_values_count += 1
            else:
                # Cari parent OID (OID checksheet mungkin lebih panjang)
                for parent_oid, parent_value in private_oid_value_map.items():
                    if c_oid.startswith(parent_oid + "."):
                        model_value = parent_value
                        found_values_count += 1
                        break
            
            # Siapkan data sesuai hasil matching
            if is_match:
                match_count += 1
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
                    "NA",         # Kolom K (11)
                    "",           # Kolom L (12)
                    "",           # Kolom M (13)
                    "-",          # Kolom N (14)
                    "NA",         # Kolom O (15)
                    "",           # Kolom P (16)
                    "",           # Kolom Q (17)
                    "-",          # Kolom R (18)
                    "NA",         # Kolom S (19)
                    "",           # Kolom T (20)
                    "",           # Kolom U (21)
                    "-"           # Kolom V (22)
                ]
            
            output_data_rest.append(row_data)
            
            # Progress indicator setiap 100 baris
            if idx % 100 == 0:
                print(f"   ⏳ Progress: {idx}/{total_rows} baris ({idx*100//total_rows}%)")
        
        print(f"   ✓ Matching selesai: {match_count} match, {no_support_count} no-support ({time.time()-step_start:.2f}s)")
        print(f"   ℹ️  OID dengan nilai model: {found_values_count}/{len(check_oid_range)}")
        
        # OPTIMASI: Tulis semua data sekaligus ke Excel (BULK WRITE!)
        print("\n💾 [6/6] Menulis hasil ke Checksheet...")
        step_start = time.time()
        
        if output_data_col_i and last_row_check >= start_row:
            # Tulis kolom I (FactoryDefault/NoSupport)
            col_i_data = [[val] for val in output_data_col_i]
            ws_check.Range(ws_check.Cells(start_row, 9), ws_check.Cells(last_row_check, 9)).Value = col_i_data
            
            # Tulis kolom J (Nilai dari model)
            col_j_data = [[val] for val in output_data_col_j]
            ws_check.Range(ws_check.Cells(start_row, 10), ws_check.Cells(last_row_check, 10)).Value = col_j_data
            
            # Tulis kolom K sampai V (data evaluasi lainnya)
            ws_check.Range(ws_check.Cells(start_row, 11), ws_check.Cells(last_row_check, 22)).Value = output_data_rest
            
        print(f"   ✓ Data berhasil ditulis ke Excel ({time.time()-step_start:.2f}s)")
        
        # ===== SIMPAN FILE =====
        print("\n💾 Menyimpan file hasil...")
        step_start = time.time()
        file_name, file_ext = os.path.splitext(path_check)
        new_path_check = f"{file_name}_aftergenerate{file_ext}"
        
        if file_ext.lower() == '.xlsm':
            file_format = 52  # xlOpenXMLWorkbookMacroEnabled
        else:
            file_format = 51  # xlOpenXMLWorkbook (.xlsx)
            
        wb_check.SaveAs(new_path_check, FileFormat=file_format)
        wb_check.Close(False)
        print(f"   ✓ File tersimpan ({time.time()-step_start:.2f}s)")
        
        # ===== BUKA FILE HASIL =====
        print("\n📂 Membuka file hasil di Excel...")
        try:
            # Buka file hasil agar user bisa langsung lihat
            excel.Visible = True
            excel.Workbooks.Open(os.path.abspath(new_path_check))
            print(f"   ✓ File dibuka: {new_path_check}")
        except Exception as e:
            print(f"   ⚠️ Tidak bisa auto-open: {e}")
            print(f"   ℹ️  Silakan buka manual: {new_path_check}")
        
        # ===== SELESAI =====
        total_time = time.time() - start_time
        print("\n" + "="*70)
        print("✅ PROSES SELESAI!")
        print("="*70)
        print(f"📊 Ringkasan:")
        print(f"   • Total OID di MIB      : {len(mib_oids_list)}")
        print(f"   • Total baris diproses  : {total_rows}")
        print(f"   • OID match (support)   : {match_count}")
        print(f"   • OID no-support        : {no_support_count}")
        print(f"   • Model dipilih         : {selected_model_name}")
        print(f"   • Waktu total           : {total_time:.2f} detik")
        print(f"   • File output           : {os.path.basename(new_path_check)}")
        print("="*70 + "\n")
        
    except Exception as e:
        print(f"\n❌ TERJADI ERROR: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # Jangan quit Excel agar user bisa lihat hasilnya
        # excel.EnableEvents = True 
        # excel.Quit()
        pass

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