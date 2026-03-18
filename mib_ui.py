"""Streamlit UI for standalone app."""

import os
import queue
import tempfile
import threading

import streamlit as st

from mib_core import OPENPYXL_AVAILABLE, read_ccodes_openpyxl, detect_models_from_spek
from mib_workers import background_worker_mode1, background_worker_mode2

st.set_page_config(
    page_title="MIB Checksheet Generator",
    page_icon="🖨️",
    layout="wide"
)

def main():
    with st.sidebar:
        st.title("🖨️ MIB Checksheet Generator")
        st.markdown("Pilih mode operasi yang diinginkan")
        st.divider()
        mode = st.radio(
            "Pilih Menu:",
            [
                "📊 MIB Checksheet Overall",
                "🎯 MIB Checksheet Section"
            ],
            index=0
        )
        
        st.divider()
        st.markdown("### ℹ️ Info")
        
        if "Overall" in mode:
            st.info("""
            **MIB Checksheet Overall**
            - Generate checksheet per Model
            - Fast detection dengan openpyxl
            - Template injection otomatis
            - Background processing
            """)
        else:
            st.info("""
            **MIB Checksheet Section**
            - Generate per Model
            - OID + Attribute validation
            - Prevents false positives
            - Background processing
            """)
    
    if "Overall" in mode:
        st.markdown("## 📊 MIB Checksheet Overall")
        st.caption("Generate checksheet dari MIB Implementation Specification dengan memilih model tertentu")
        st.divider()
        run_mode_ccode()
    else:
        st.markdown("## 🎯 MIB Checksheet Section")
        st.caption("Generate checksheet dengan matching OID + Attribute untuk model tertentu")
        st.divider()
        run_mode_model()

# ==========================================
# MODE 1: MODEL SELECTION
# ==========================================

def run_mode_ccode():
    """Mode 1: Model Selection dengan background processing""" 
    
    # Template Sheet otomatis dari folder project
    template_path = os.path.join(os.path.dirname(__file__), "Template Sheet.xlsm")
    if os.path.exists(template_path):
        st.success(f"✅ Template Sheet: Loaded automatically")
    else:
        st.warning(f"⚠️ Template Sheet not found")
    
    uploaded_spek = st.file_uploader("📂 MIB Implementation Specification (.xlsm)", type=["xlsm"], key="mode1_spek")
    
    if not uploaded_spek:
        st.info("👆 Upload MIB Implementation Specification")
        return
    
    # Save uploaded file
    temp_dir = tempfile.gettempdir()
    spek_path = os.path.join(temp_dir, uploaded_spek.name)
    
    with open(spek_path, "wb") as f:
        f.write(uploaded_spek.getbuffer())
    
    st.success(f"✅ File uploaded: {uploaded_spek.name}")
    
    # Detect Models
    st.markdown("### 🔍 Model Detection")
    
    cache_key = f"{uploaded_spek.name}_{uploaded_spek.size}"
    
    if "mode1_ccodes" not in st.session_state or st.session_state.get('mode1_cache_key') != cache_key:
        if not OPENPYXL_AVAILABLE:
            st.error("❌ openpyxl not installed. Run: pip install openpyxl")
            return
        
        with st.spinner("⚡ Detecting Models..."):
            try:
                ccodes = read_ccodes_openpyxl(spek_path)
                st.session_state.mode1_ccodes = ccodes
                st.session_state.mode1_cache_key = cache_key
                st.session_state.mode1_spek_path = spek_path
                st.success(f"⚡ Detected {len(ccodes)} Models")
            except Exception as e:
                st.error(f"❌ Detection failed: {e}")
                return
    else:
        st.info(f"📦 Using cached detection: {len(st.session_state.mode1_ccodes)} Models")
    
    if not st.session_state.mode1_ccodes:
        st.warning("⚠️ No Models detected")
        return
    
    # Model Selection & Generate
    st.markdown("### 🎯 Generate Checksheet")
    
    ccodes_list = [c["name"] for c in st.session_state.mode1_ccodes]
    selected_name = st.selectbox("Select Model:", ccodes_list, key="mode1_ccode_select")
    
    if st.button("🚀 Generate Checksheet", type="primary", use_container_width=True, key="mode1_generate"):
        selected = next(c for c in st.session_state.mode1_ccodes if c["name"] == selected_name)
        
        # Initialize queue for background worker
        result_queue = queue.Queue()
        
        # Start background thread
        worker_thread = threading.Thread(
            target=background_worker_mode1,
            args=(st.session_state.mode1_spek_path, selected, template_path, result_queue),
            daemon=True
        )
        worker_thread.start()
        
        # UI tracking
        progress_bar = st.progress(0, text="Starting...")
        
        # Poll queue for updates
        with st.status(f"⚙️ Generating {selected_name}...", expanded=True) as status:
            while worker_thread.is_alive() or not result_queue.empty():
                try:
                    update = result_queue.get(timeout=0.1)
                    
                    if "progress" in update:
                        progress_bar.progress(update["progress"], text=update["message"])
                        st.write(update["message"])
                    
                    if update.get("status") == "success":
                        status.update(label=update["message"], state="complete")
                        
                        output_path = update["output_path"]
                        if os.path.exists(output_path):
                            file_size = os.path.getsize(output_path) / (1024 * 1024)
                            st.success(f"✅ Generated: {os.path.basename(output_path)} ({file_size:.2f} MB)")
                            
                            with open(output_path, "rb") as f:
                                st.download_button(
                                    label="📥 Download Generated Checksheet",
                                    data=f,
                                    file_name=os.path.basename(output_path),
                                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                                    use_container_width=True
                                )
                            st.balloons()
                        break
                    
                    elif update.get("status") == "error":
                        status.update(label="❌ Error occurred!", state="error")
                        st.error(f"❌ Generation failed: {update['message']}")
                        with st.expander("🔍 Error Details"):
                            st.code(update.get("traceback", ""), language="text")
                        break
                
                except queue.Empty:
                    continue

# ==========================================
# MODE 2: MODEL SELECTION
# ==========================================

def run_mode_model():
    """Mode 2: Model selection dengan background processing"""
    
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
    
    if not uploaded_spek or not uploaded_checksheet:
        st.info("👆 Upload kedua file untuk mulai")
        return
    
    # Save uploaded files
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
    
    cache_key_mode2 = f"{uploaded_spek.name}_{uploaded_spek.size}_{uploaded_checksheet.name}"
    
    if "models_detected" not in st.session_state or st.session_state.get("mode2_cache_key") != cache_key_mode2:
        with st.spinner("🔍 Detecting available models..."):
            try:
                checksheet_filename = os.path.basename(checksheet_path)
                models = detect_models_from_spek(spek_path, checksheet_filename)
                
                if models:
                    st.session_state.models_detected = models
                    st.session_state.spek_path = spek_path
                    st.session_state.checksheet_path = checksheet_path
                    st.session_state.mode2_cache_key = cache_key_mode2
                    st.success(f"✅ Detected {len(models)} models")
                else:
                    st.error("❌ No models detected")
                    return
            
            except Exception as e:
                st.error(f"❌ Error detecting models: {e}")
                return
    else:
        models = st.session_state.models_detected
        st.info(f"📦 {len(models)} models detected: {[m['name'] for m in models]}")
    
    # Model selection
    st.markdown("### 🎯 Generate Checksheet")
    
    model_names = [m["name"] for m in st.session_state.models_detected]
    selected_model = st.selectbox(
        "Select Model:",
        model_names,
        help="Pilih model yang ingin di-generate"
    )
    
    if st.button("🚀 Generate Checksheet", type="primary", use_container_width=True):
        result_queue = queue.Queue()
        
        worker_thread = threading.Thread(
            target=background_worker_mode2,
            args=(st.session_state.spek_path, st.session_state.checksheet_path, selected_model, result_queue),
            daemon=True
        )
        worker_thread.start()
        
        progress_bar = st.progress(0, text="Starting...")
        
        with st.status(f"⚙️ Generating {selected_model}...", expanded=True) as status:
            while worker_thread.is_alive() or not result_queue.empty():
                try:
                    update = result_queue.get(timeout=0.1)
                    
                    if "progress" in update:
                        progress_bar.progress(update["progress"], text=update["message"])
                        st.write(update["message"])
                    
                    if update.get("status") == "success":
                        status.update(label=update["message"], state="complete")
                        
                        output_path = update["output_path"]
                        if os.path.exists(output_path):
                            file_size = os.path.getsize(output_path) / (1024 * 1024)
                            st.success(f"✅ Generated: {os.path.basename(output_path)} ({file_size:.2f} MB)")
                            
                            # Show detailed stats
                            if "stats" in update:
                                stats = update["stats"]
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.metric("Match", stats['match_count'], help="OID yang di-support")
                                    st.metric("NoSupport", stats['no_support_count'], help="OID tidak di-support")
                                with col2:
                                    st.metric("Total Rows", stats['total_rows'])
                                
                                # Value source breakdown
                                with st.expander("📊 Value Source Details"):
                                    st.markdown(f"""
                                    - **Exact OID Match**: {stats.get('value_from_exact_oid', 0)} values
                                    - **Parent OID Match**: {stats.get('value_from_parent_oid', 0)} values
                                    - **Attribute Match**: {stats.get('value_from_attr', 0)} values
                                    - **Empty Values**: {stats.get('value_empty', 0)} (matched but no value)
                                    """)
                            
                            with open(output_path, "rb") as f:
                                st.download_button(
                                    label="📥 Download Generated Checksheet",
                                    data=f,
                                    file_name=os.path.basename(output_path),
                                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                                    use_container_width=True
                                )
                            st.balloons()
                        break
                    
                    elif update.get("status") == "error":
                        status.update(label="❌ Error occurred!", state="error")
                        st.error(f"❌ Generation failed: {update['message']}")
                        with st.expander("🔍 Error Details"):
                            st.code(update.get("traceback", ""), language="text")
                        break
                
                except queue.Empty:
                    continue

if __name__ == "__main__":
    main()
