"""
MIB Checksheet Generator - Integrated Version
Menggabungkan 2 mode:
1. Mode C-Code Selection (Original)
2. Mode Model Selection dengan Double Validation (New)
"""

import streamlit as st
import sys
import os
import warnings

# Suppress ScriptRunContext warnings dari import non-Streamlit modules
warnings.filterwarnings('ignore', message='.*ScriptRunContext.*')

# Set page config
st.set_page_config(
    page_title="MIB Checksheet Generator - Integrated",
    page_icon="🖨️",
    layout="wide"
)

# ==========================================
# MAIN UI
# ==========================================

def main():
    # st.title("🖨️ MIB Checksheet Generator - Integrated")
    # st.markdown("Pilih mode operasi yang diinginkan")
    # st.divider()
    
    # Sidebar untuk mode selection
    with st.sidebar:
        st.title("🖨️ MIB Checksheet Generator")
        st.markdown("Pilih mode operasi yang diinginkan")
        st.divider()
        mode = st.radio(
            "Pilih Menu:",
            [
                "📊 Mode 1: MIB Checksheet Overall",
                "🎯 Mode 2: Model Selection + Validation"
            ],
            index=0
        )
        
        st.divider()
        st.markdown("### ℹ️ Info")
        
        if "Mode 1" in mode:
            st.info("""
            **Mode C-Code Selection**
            - Generate checksheet per C-Code
            - Web-based interface
            - Fast C-Code detection
            - Template injection
            """)
        else:
            st.info("""
            **Mode Model Selection**
            - Generate checksheet per Model
            - OID + Attribute double validation
            - Prevents false positives
            - Advanced matching system
            """)
    
    # Display selected mode
    if "Mode 1" in mode:
        st.markdown("## 📊 MIB Checksheet Overall")
        st.caption("Generate checksheet dari MIB Implementation Specification dengan memilih model tertentu")
        st.divider()
        run_mode_ccode()
    else:
        st.markdown("## 🎯 MIB Checksheet Section")
        st.caption("Generate checksheet dengan matching OID + Attribute untuk model tertentu terhadap MIB Implementation Specification")
        st.divider()
        run_mode_model()

# ==========================================
# MODE 1: C-CODE SELECTION
# ==========================================

def run_mode_ccode():
    """Mode 1: Original C-Code selection dari app_v3.py""" 
    # Template Sheet otomatis dari folder project
    template_path = os.path.join(os.path.dirname(__file__), "Template Sheet.xlsm")
    # if os.path.exists(template_path):
    #     st.success(f"✅ Template Sheet: Loaded automatically from project folder")
    # else:
    #     st.error(f"❌ Template Sheet not found at: {template_path}")
    
    uploaded_spek = st.file_uploader("📂 MIB Implementation Specification (.xlsm)", type=["xlsm"], key="mode1_spek")
    
    if not uploaded_spek:
        st.info("👆 Upload MIB Implementation Specification")
        return
    
    # Save uploaded file
    import tempfile
    temp_dir = tempfile.gettempdir()
    spek_path = os.path.join(temp_dir, uploaded_spek.name)
    
    with open(spek_path, "wb") as f:
        f.write(uploaded_spek.getbuffer())
    
    st.success(f"✅ File uploaded: {uploaded_spek.name}")
    
    # Detect C-Codes
    st.markdown("### 🔍 C-Code Detection")
    
    cache_key = f"{uploaded_spek.name}_{uploaded_spek.size}"
    
    if "mode1_ccodes" not in st.session_state or st.session_state.get('mode1_cache_key') != cache_key:
        with st.spinner("⚡ Detecting C-Codes..."):
            try:
                # Import fungsi dari app_v3.py
                sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                from app_v3 import read_ccodes_openpyxl, OPENPYXL_AVAILABLE
                
                if OPENPYXL_AVAILABLE:
                    ccodes = read_ccodes_openpyxl(spek_path)
                    st.session_state.mode1_ccodes = ccodes
                    st.session_state.mode1_cache_key = cache_key
                    st.session_state.mode1_spek_path = spek_path
                    st.success(f"⚡ Detected {len(ccodes)} C-Codes")
                else:
                    st.error("❌ openpyxl not installed. Run: pip install openpyxl")
                    return
                    
            except Exception as e:
                st.error(f"❌ Detection failed: {e}")
                return
    else:
        st.info(f"📦 Using cached detection: {len(st.session_state.mode1_ccodes)} C-Codes")
    
    if not st.session_state.mode1_ccodes:
        st.warning("⚠️ No C-Codes detected")
        return
    
    # C-Code Selection & Generate
    st.markdown("### 🎯 Generate Checksheet")
    
    ccodes_list = [c["name"] for c in st.session_state.mode1_ccodes]
    selected_name = st.selectbox("Select C-Code:", ccodes_list, key="mode1_ccode_select")
    
    if st.button("🚀 Generate Checksheet", type="primary", use_container_width=True, key="mode1_generate"):
        selected = next(c for c in st.session_state.mode1_ccodes if c["name"] == selected_name)
        
        progress_bar = st.progress(0, text="Starting generation...")
        
        with st.status(f"⚙️ Generating {selected_name}...", expanded=True) as status:
            try:
                # Import generate functions
                from app_v3 import (init_excel_app, create_dump_sheet, inject_template_sheets,
                                   process_sheet_cutting, delete_unused_sheets, ExcelConst, 
                                   Config as V3Config, close_excel_safely)
                
                import shutil
                import time
                
                # Prepare output
                safe_name = "".join(c for c in selected["name"] if c.isalnum() or c == ' ').strip()
                output_filename = f"Checksheet_{safe_name.replace(' ', '_')}.xlsm"
                output_path = os.path.join(temp_dir, output_filename)
                
                shutil.copy(st.session_state.mode1_spek_path, output_path)
                
                start_time = time.time()
                
                # Step 1: Initialize Excel
                progress_bar.progress(10, text="🔌 Initializing Excel...")
                st.write("🔌 Initializing Excel engine...")
                excel = init_excel_app()
                
                # Step 2: Open workbook
                progress_bar.progress(20, text="📁 Opening workbook...")
                st.write("📁 Opening workbook...")
                wb = excel.Workbooks.Open(
                    os.path.abspath(output_path),
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True
                )
                
                excel.Calculation = ExcelConst.CALCULATION_MANUAL
                
                # Step 3: Create dump sheet
                progress_bar.progress(30, text="📝 Creating dump sheet...")
                st.write("📝 Creating dump sheet...")
                create_dump_sheet(wb)
                
                # Step 4: Inject templates
                progress_bar.progress(40, text="📑 Injecting template sheets...")
                st.write("📑 Injecting template sheets...")
                if os.path.exists(template_path):
                    inject_template_sheets(wb, template_path, excel)
                    st.write("✅ Template injected")
                
                # Step 5: Process cutting
                progress_bar.progress(50, text="✂️ Processing sheets...")
                st.write("✂️ Processing: cut columns & add evaluation...")
                process_sheet_cutting(wb, selected["col_index"], excel)
                
                # Step 6: Clean sheets
                progress_bar.progress(80, text="🧹 Cleaning...")
                st.write("🧹 Cleaning unused sheets...")
                delete_unused_sheets(wb)
                
                # Step 7: Save
                progress_bar.progress(90, text="💾 Saving...")
                st.write("💾 Saving file...")
                wb.Save()
                wb.Close(SaveChanges=False)
                close_excel_safely(excel)
                
                progress_bar.progress(100, text="✅ Complete!")
                
                duration = time.time() - start_time
                status.update(label=f"✅ Generation complete! ({duration:.1f}s)", state="complete")
                
                # Download button
                if os.path.exists(output_path):
                    file_size = os.path.getsize(output_path) / (1024 * 1024)
                    st.success(f"✅ Generated: {output_filename} ({file_size:.2f} MB)")
                    
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Download Generated Checksheet",
                            data=f,
                            file_name=output_filename,
                            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                            use_container_width=True
                        )
                    
                    st.balloons()
                    
            except Exception as e:
                progress_bar.progress(0, text="❌ Failed!")
                status.update(label="❌ Error occurred!", state="error")
                st.error(f"❌ Generation failed: {e}")
                
                import traceback
                with st.expander("🔍 Error Details"):
                    st.code(traceback.format_exc(), language="text")

# ==========================================
# MODE 2: MODEL SELECTION
# ==========================================

def run_mode_model():
    """Mode 2: Model selection dengan double validation dari app_section.py"""
    
    st.markdown("### 📥 Upload Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        uploaded_spek = st.file_uploader(
            "📂 MIB Implementation Specification (.xlsm)", 
            type=["xlsm"], 
            key="mode2_spek",
            help="File PRZ-*.xlsm yang berisi EPSON Private MIB"
        )
    
    with col2:
        uploaded_checksheet = st.file_uploader(
            "📋 Checksheet Template (.xlsm)", 
            type=["xlsm"], 
            key="mode2_checksheet",
            help="File 初期値評価仕様書 (Printer Private sheet)"
        )
    
    # Jika belum upload
    if not uploaded_spek or not uploaded_checksheet:
        st.info("👆 Upload kedua file untuk mulai")
        
        # with st.expander("ℹ️ About Mode 2"):
        #     st.markdown("""
        #     ### 🎯 Mode 2: Model Selection + Double Validation
            
        #     **Key Features:**
        #     - 🎯 **Model Selection**: Pilih model spesifik dari Private MIB (CL56, CL59, dst)
        #     - ✅ **Double Validation**: OID match DAN Attribute Name match
        #     - 🚫 **False Positive Prevention**: Reject OID match jika Attribute tidak cocok
        #     - 📊 **Detailed Metrics**: Tracking match by OID Exact, Parent, Attribute, Validated
        #     - 🔍 **Advanced Matching**: Support parent OID match untuk index variations (.1.1, dll)
            
        #     **Matching Strategy:**
        #     1. Cek "範囲外" in column E → immediate NoSupport
        #     2. OID exact match + attribute validation
        #     3. OID parent match + attribute validation
        #     4. Reject if OID match but attribute ≠
        #     5. Fallback: attribute-only match
            
        #     **Output Format:**
        #     - Column I: FactoryDefault / NoSupport
        #     - Column J: Model value
        #     - Column K, O, S: `[NA]` for NoSupport (dengan kurung siku)
        #     - Column P, T: Kosong (reserved)
        #     - Column N, R, V: "○" for support, "-" for no-support
            
        #     **File Naming:**
        #     - Format: `{original_name}_{ModelName}_aftergenerate.xlsm`
        #     - Example: `02_初期値評価仕様書_epPrt_CL56_aftergenerate.xlsm`
        #     """)
        
        return
    
    # Save uploaded files
    import tempfile
    temp_dir = tempfile.gettempdir()
    
    spek_path = os.path.join(temp_dir, uploaded_spek.name)
    checksheet_path = os.path.join(temp_dir, uploaded_checksheet.name)
    
    with open(spek_path, "wb") as f:
        f.write(uploaded_spek.getbuffer())
    
    with open(checksheet_path, "wb") as f:
        f.write(uploaded_checksheet.getbuffer())
    
    st.success(f"✅ Files uploaded: {uploaded_spek.name}, {uploaded_checksheet.name}")
    
    # Detect models
    st.markdown("### 🔍 Model Detection")
    
    if "models_detected" not in st.session_state:
        with st.spinner("🔍 Detecting available models from Spek..."):
            try:
                # Import fungsi dari app_section.py (without IO redirection)
                sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                from app_section import detect_models_from_spek
                
                # Call detection function (console output akan muncul di terminal)
                models = detect_models_from_spek(spek_path)
                
                if models:
                    st.session_state.models_detected = models
                    st.session_state.spek_path = spek_path
                    st.session_state.checksheet_path = checksheet_path
                    st.success(f"✅ Detected {len(models)} models: {[m['name'] for m in models]}")
                else:
                    st.error("❌ No models detected in Spek file")
                    return
            
            except Exception as e:
                st.error(f"❌ Error detecting models: {e}")
                st.code(str(e), language="text")
                return
    else:
        models = st.session_state.models_detected
        st.info(f"📦 {len(models)} models detected: {[m['name'] for m in models]}")
    
    # Model selection
    st.markdown("### 🎯 Model Selection")
    
    model_names = [m["name"] for m in st.session_state.models_detected]
    selected_model = st.selectbox(
        "Select Model:",
        model_names,
        help="Pilih model yang ingin di-generate"
    )
    
    # Generation settings
    with st.expander("⚙️ Advanced Settings"):
        st.checkbox("Enable detailed logging", value=True, key="detailed_log")
        st.checkbox("Auto-close Excel after save", value=True, key="auto_close")
        st.slider("Progress update interval (rows)", 50, 200, 100, key="progress_interval")
    
    # Generate button
    st.divider()
    
    if st.button("🚀 Generate Checksheet", type="primary", use_container_width=True):
        st.markdown(f"### ⚙️ Generating for Model: **{selected_model}**")
        
        progress_bar = st.progress(0, text="Starting generation...")
        log_container = st.container()
        
        with log_container:
            with st.status(f"⚙️ Generating {selected_model}...", expanded=True) as status:
                try:
                    # Import generate function (without IO redirection)
                    from app_section import generate_checksheet
                    
                    # Progress updates
                    progress_bar.progress(10, text="🔌 Initializing Excel...")
                    st.write("🔌 Initializing Excel COM application...")
                    
                    progress_bar.progress(20, text="📖 Reading MIB data...")
                    st.write("📖 Opening files and reading MIB data...")
                    
                    progress_bar.progress(40, text="⚙️ Processing matching...")
                    st.write("⚙️ Processing OID + Attribute matching with double validation...")
                    
                    # Run generation (NO IO redirection to avoid COM context issues)
                    import time
                    
                    start_time = time.time()
                    
                    # Call generate_checksheet directly without IO redirection
                    # Console output akan muncul di terminal, bukan di Streamlit UI
                    output_path = generate_checksheet(
                        st.session_state.spek_path,
                        st.session_state.checksheet_path,
                        selected_model
                    )
                    
                    duration = time.time() - start_time
                    
                    progress_bar.progress(90, text="💾 Saving file...")
                    st.write("💾 Saving generated file...")
                    
                    progress_bar.progress(100, text="✅ Complete!")
                    
                    status.update(
                        label=f"✅ Generation complete! ({duration:.1f}s)",
                        state="complete",
                        expanded=False
                    )
                    
                    st.success(f"✅ Checksheet generated successfully in {duration:.1f}s")
                    
                    # Show output file info
                    if output_path and os.path.exists(output_path):
                        file_size = os.path.getsize(output_path) / (1024 * 1024)
                        st.info(f"📁 Output: `{os.path.basename(output_path)}` ({file_size:.2f} MB)")
                        
                        # Download button
                        with open(output_path, "rb") as f:
                            st.download_button(
                                label="📥 Download Generated Checksheet",
                                data=f,
                                file_name=os.path.basename(output_path),
                                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                                use_container_width=True
                            )
                        
                        # Show statistics
                        with st.expander("📊 Generation Statistics"):
                            st.markdown(f"""
                            - **Model**: {selected_model}
                            - **Duration**: {duration:.2f} seconds
                            - **Output Size**: {file_size:.2f} MB
                            - **Spek File**: {uploaded_spek.name}
                            - **Checksheet File**: {uploaded_checksheet.name}
                            """)
                        
                        st.balloons()
                    
                except Exception as e:
                    progress_bar.progress(0, text="❌ Failed!")
                    status.update(label="❌ Error occurred!", state="error", expanded=True)
                    st.error(f"❌ Generation failed: {e}")
                    
                    with st.expander("🔍 Error Details"):
                        st.code(str(e), language="text")
                        import traceback
                        st.code(traceback.format_exc(), language="text")

    # Sidebar info for Mode 2
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📊 Mode 2 Features")
        st.caption("✅ Model Selection")
        st.caption("✅ Double Validation (OID + Attr)")
        st.caption("✅ False Positive Prevention")
        st.caption("✅ Parent OID Matching")
        st.caption("✅ Detailed Statistics")
        
        st.markdown("---")
        st.markdown("### 🎯 Current Session")
        if "models_detected" in st.session_state:
            st.write(f"Models: {len(st.session_state.models_detected)}")
            for model in st.session_state.models_detected:
                st.caption(f"• {model['name']}")

if __name__ == "__main__":
    main()
