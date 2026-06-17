import sys
import io
import os
import time
import subprocess
import datetime
from pathlib import Path
import glob

if sys.platform == "win32":
    try:
        if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding != 'utf-8':
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', newline='')
        if not isinstance(sys.stderr, io.TextIOWrapper) or sys.stderr.encoding != 'utf-8':
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', newline='')
    except (AttributeError, ValueError, OSError):
        pass  # Gracefully continue if wrapping fails

import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Weather Auto Regression",
    layout="wide",
    page_icon="⚡",
    initial_sidebar_state="expanded"
)

# ====================== STYLING ======================
st.markdown("""
<style>
    /* Force dark background everywhere */
    .stApp {
        background: #0f172a !important;
    }
    [data-testid="stAppViewContainer"] {
        background: #0f172a !important;
    }
    [data-testid="stMain"] {
        background: #0f172a !important;
    }
    .main .block-container {
        background: #0f172a !important;
    }

    /* SIDEBAR */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%) !important;
        border-right: 2px solid #14b8a6 !important;
    }
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 2rem !important;
    }

    /* LOGO */
    .logo-container {
        background: rgba(20,184,166,0.1);
        border: 2px solid #14b8a6;
        border-radius: 16px;
        padding: 24px 16px;
        margin-bottom: 24px;
        text-align: center;
    }
    .company-name {
        font-size: 20px;
        font-weight: 700;
        color: #14b8a6 !important;
        font-family: 'Courier New', monospace;
        margin-bottom: 6px;
    }
    .tagline {
        font-size: 12px;
        color: #94a3b8 !important;
        font-weight: 500;
    }
    .sidebar-info {
        background: rgba(245,158,11,0.1);
        border: 1px solid #f59e0b;
        border-radius: 12px;
        padding: 12px 14px;
        font-size: 12px;
        color: #f59e0b !important;
        font-weight: 600;
        text-align: center;
    }

    /* MAIN TITLE */
    .main-title {
        font-size: 48px;
        font-weight: 800;
        background: linear-gradient(135deg, #14b8a6 0%, #f59e0b 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 8px;
        letter-spacing: -1px;
    }

    /* BUTTONS */
    .stButton > button {
        background: linear-gradient(135deg, #14b8a6 0%, #0f766e 100%) !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-weight: 700 !important;
        height: 52px !important;
        border: none !important;
        border-radius: 10px !important;
        font-size: 15px !important;
        letter-spacing: 0.5px !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%) !important;
        transform: translateY(-2px) !important;
    }
    .stButton > button p,
    .stButton > button span {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    /* INPUT FIELDS */
    .stTextInput > div > div > input {
        background-color: #1e293b !important;
        border: 2px solid #334155 !important;
        color: #f1f5f9 !important;
        -webkit-text-fill-color: #f1f5f9 !important;
        border-radius: 8px !important;
    }

    /* ALL LABELS */
    label, label p, label span,
    .stTextInput label, .stTextInput label p,
    .stSelectbox label, .stSelectbox label p,
    .stDateInput label, .stDateInput label p,
    .stCheckbox label, .stCheckbox label p,
    .stCheckbox label span,
    [data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] p,
    [data-testid="stWidgetLabel"] span {
        color: #facc15 !important;
        -webkit-text-fill-color: #facc15 !important;
        font-size: 14px !important;
        font-weight: 600 !important;
    }
    /* ALL PARAGRAPH TEXT */
    p, span, div.stMarkdown p {
    color: #1e293b !important;  /* Dark text for dark bg */
    -webkit-text-fill-color: #1e293b !important;
    }

    /* Ensure dark text on light backgrounds */
    [data-testid="stAppViewContainer"] p,
    [data-testid="stAppViewContainer"] span {
    color: #1e293b !important;
    }

    /* ALERT / INFO / SUCCESS / WARNING BOXES */
    [data-testid="stAlert"] {
        border-radius: 10px !important;
        border-left: 4px solid !important;
    }
    [data-testid="stAlert"] p,
    [data-testid="stAlert"] span,
    [data-testid="stAlert"] strong,
    [data-testid="stAlert"] li,
    [data-testid="stNotification"] p {
        color: #f8fafc !important;
        -webkit-text-fill-color: #f8fafc !important;
        font-weight: 500 !important;
    }
    div[data-testid="stMarkdownContainer"] p,
    div[data-testid="stMarkdownContainer"] span,
    div[data-testid="stMarkdownContainer"] strong,
    div[data-testid="stMarkdownContainer"] li {
        color: #f8fafc !important;
        -webkit-text-fill-color: #f8fafc !important;
    }

    /* TABS */
    [data-testid="stTabs"] button {
        color: #e2e8f0 !important;
        -webkit-text-fill-color: #e2e8f0 !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        border-bottom: 2px solid transparent !important;
        background: transparent !important;
    }
    [data-testid="stTabs"] button p,
    [data-testid="stTabs"] button span {
        color: #e2e8f0 !important;
        -webkit-text-fill-color: #e2e8f0 !important;
        font-size: 15px !important;
        font-weight: 600 !important;
    }
    [data-testid="stTabs"] button[aria-selected="true"] {
        color: #14b8a6 !important;
        -webkit-text-fill-color: #14b8a6 !important;
        border-color: #14b8a6 !important;
    }
    [data-testid="stTabs"] button[aria-selected="true"] p,
    [data-testid="stTabs"] button[aria-selected="true"] span {
        color: #14b8a6 !important;
        -webkit-text-fill-color: #14b8a6 !important;
    }

    /* SELECTBOX */
    [data-testid="stSelectbox"] > div > div {
        background: #1e293b !important;
        color: #f1f5f9 !important;
        -webkit-text-fill-color: #f1f5f9 !important;
        border: 2px solid #334155 !important;
        border-radius: 8px !important;
    }
    [data-testid="stSelectbox"] span,
    [data-testid="stSelectbox"] div[data-baseweb] span {
        color: #f1f5f9 !important;
        -webkit-text-fill-color: #f1f5f9 !important;
    }

    /* DATE INPUT */
    [data-testid="stDateInput"] input {
        background: #1e293b !important;
        color: #f1f5f9 !important;
        -webkit-text-fill-color: #f1f5f9 !important;
        border: 2px solid #334155 !important;
    }

    /* EXPANDER */
    [data-testid="stExpander"] > div:first-child > button {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        color: #f1f5f9 !important;
        -webkit-text-fill-color: #f1f5f9 !important;
        font-weight: 600 !important;
    }
    [data-testid="stExpander"] summary p,
    [data-testid="stExpander"] details summary p {
        color: #f1f5f9 !important;
        -webkit-text-fill-color: #f1f5f9 !important;
    }

    /* PROGRESS BAR */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #14b8a6, #f59e0b) !important;
        border-radius: 8px !important;
    }

    /* STATUS TABLE */
    .status-table {
        background-color: #1e293b;
        border-radius: 12px;
        border: 1px solid #334155;
        padding: 20px;
    }
    .status-table table { width: 100%; border-collapse: separate; border-spacing: 0 8px; }
    .status-table thead { background: rgba(20,184,166,0.1); }
    .status-table th {
        padding: 14px 16px !important;
        color: #14b8a6 !important;
        font-weight: 700;
        font-size: 12px;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }
    .status-table tbody tr {
        background: rgba(30,41,59,0.6);
        border: 1px solid #334155;
        transition: all 0.3s ease;
    }
    .status-table tbody tr:hover {
        background: rgba(20,184,166,0.1);
        transform: translateX(4px);
    }
    .status-table td {
        padding: 14px 16px !important;
        color: #cbd5e1 !important;
        font-size: 13px;
        font-weight: 500;
    }

    /* STATUS BADGES */
    .step-pending  { 
        color: #94a3b8 !important; 
        font-weight: 700;
        background: rgba(148,163,184,0.25) !important;
        padding: 8px 16px !important;
        border-radius: 6px !important;
        display: inline-block !important;
        border-left: 4px solid #94a3b8 !important;
        white-space: nowrap;
    }
    .step-running  { 
        color: #fbbf24 !important; 
        font-weight: 700;
        background: rgba(245,158,11,0.3) !important;
        padding: 8px 16px !important;
        border-radius: 6px !important;
        display: inline-block !important;
        border-left: 4px solid #f59e0b !important;
        white-space: nowrap;
        animation: pulse 1.5s infinite;
    }
    .step-done     { 
        color: #10b981 !important; 
        font-weight: 700;
        background: rgba(16,185,129,0.25) !important;
        padding: 8px 16px !important;
        border-radius: 6px !important;
        display: inline-block !important;
        border-left: 4px solid #10b981 !important;
        white-space: nowrap;
    }
    .step-error    { 
        color: #fca5a5 !important; 
        font-weight: 700;
        background: rgba(239,68,68,0.3) !important;
        padding: 8px 16px !important;
        border-radius: 6px !important;
        display: inline-block !important;
        border-left: 4px solid #ef4444 !important;
        white-space: nowrap;
        animation: shake 0.5s ease-in-out;
    }
    .step-skipped  { 
        color: #14b8a6 !important; 
        font-weight: 700;
        background: rgba(20,184,166,0.25) !important;
        padding: 8px 16px !important;
        border-radius: 6px !important;
        display: inline-block !important;
        border-left: 4px solid #14b8a6 !important;
        white-space: nowrap;
    }

    /* SUCCESS COMPLETION PANEL */
    .success-panel {
        background: linear-gradient(135deg, rgba(16,185,129,0.15), rgba(20,184,166,0.15));
        border: 2px solid #10b981;
        border-radius: 16px;
        padding: 32px;
        margin: 24px 0;
        box-shadow: 0 8px 32px rgba(16,185,129,0.1);
    }
    .success-title {
        font-size: 28px;
        font-weight: 800;
        color: #10b981;
        text-align: center;
        margin-bottom: 8px;
    }
    .success-subtitle {
        font-size: 14px;
        color: #94a3b8;
        text-align: center;
        margin-bottom: 32px;
    }

    /* OUTPUT FILE BLOCKS */
    .output-block {
        background: rgba(30,41,59,0.7);
        border: 2px solid #334155;
        border-radius: 12px;
        padding: 20px;
        transition: all 0.3s ease;
    }
    .output-block:hover {
        background: rgba(20,184,166,0.1);
        border-color: #14b8a6;
        transform: translateY(-2px);
    }
    .output-icon {
        font-size: 32px;
        margin-bottom: 12px;
    }
    .output-name {
        font-size: 16px;
        font-weight: 700;
        color: #f1f5f9;
        margin-bottom: 8px;
    }
    .output-path {
        font-size: 11px;
        color: #94a3b8;
        font-family: monospace;
        word-break: break-all;
    }

    /* DOWNLOAD BUTTON */
    [data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #8b5cf6 0%, #6d28d9 100%) !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }

    /* DIVIDER */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #334155, transparent);
        margin: 24px 0 !important;
    }

    /* ANIMATIONS */
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.6; }
    }
    @keyframes shake {
        0%, 100% { transform: translateX(0); }
        10%, 30%, 50%, 70%, 90% { transform: translateX(-4px); }
        20%, 40%, 60%, 80% { transform: translateX(4px); }
    }
</style>
""", unsafe_allow_html=True)

# ====================== SIDEBAR ======================
with st.sidebar:
    st.markdown("""
    <div class="logo-container">
        <div class="company-name">⚡ 50 HERTZ</div>
        <div class="tagline">Energy Automation Suite</div>
        <div class="tagline">Weather Auto v1.0</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("""
    <div class="sidebar-info">
        📊 All outputs saved to Output Folder
    </div>
    """, unsafe_allow_html=True)

# ====================== PIPELINE DEFINITION ======================
PIPELINE = [
    ("1/12  FTP Download",                "ECM10_FTP_DOWNLOAD.py"),
    ("2/12  NC Data ECM10",               "Nc_data_ECM_10(ECM10).py"),
    ("3/12  ECM10 Master",                "ECM10_MASTER.py"),
    ("4/12  File Monitoring",             "File_Monitoring(ECM10_.py"),
    ("5/12  Filtered Weather Data",       "Filtered_Weather_Data(ECM10).py"),
    ("6/12  Filtered NC File",            "FilteredNc_File(ECM10).py"),
    ("7/12  Weather DB – NC File Match",  "WeatherDB_NcFileMatch(ECM10).py"),
    ("8/12  RE Prod ECM10",               "REProdECM10.py"),
    ("9/12  RE Solar ECM10",              "RESOLAR_ECM10.py"),
    ("10/12 RE Wind ECM10",               "REWIND_ECM10.py"),
    ("11/12 Sync RE DB",                  "SYNC_RE_DB.py"),
    ("12/12 Data Validation",             "Data_Validation.py"),
]

# ====================== SESSION STATE ======================
if "statuses" not in st.session_state:
    st.session_state.statuses = ["pending"] * len(PIPELINE)
if "resume_from" not in st.session_state:
    st.session_state.resume_from = 0
if "pipeline_completed" not in st.session_state:
    st.session_state.pipeline_completed = False

# ====================== SECTION HEADER HELPER ======================
def section_header(icon, text):
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;">'
        f'<div style="width:4px;height:24px;background:linear-gradient(180deg,#14b8a6,#f59e0b);border-radius:2px;flex-shrink:0;"></div>'
        f'<span style="font-size:18px;font-weight:700;color:#ffffff !important;-webkit-text-fill-color:#ffffff !important;">{icon} {text}</span>'
        f'</div>',
        unsafe_allow_html=True
    )

# ====================== HELPER FUNCTIONS ======================
def find_script(script_name, scripts_dir):
    """Find script file handling special characters in filenames."""
    # Try exact match in scripts directory
    if (scripts_dir / script_name).exists():
        return scripts_dir / script_name
    
    # Try in current working directory
    if Path(script_name).exists():
        return Path(script_name)
    
    # Try with wildcard for parentheses (e.g., Nc_data_ECM_10*.py)
    if "(" in script_name and ")" in script_name:
        base_name = script_name.replace("(", "").replace(")", "")
        matches = glob.glob(str(scripts_dir / base_name.replace(".py", "*.py")))
        if matches:
            return Path(matches[0])
    
    return None

def run_python_script(script_name, scripts_dir, output_dir):
    """
    Execute Python script with proper environment and unbuffered output.
    Returns: (success, stdout, stderr, execution_time)
    """
    start_time = time.time()
    
    try:
        script_path = find_script(script_name, scripts_dir)
        if not script_path or not script_path.exists():
            return False, "", f"Script not found: {script_name}", 0
        
        # Set up environment - INHERIT FULL PARENT ENV
        env = os.environ.copy()  # This copies parent process env
        env["OUTPUT_DIR"] = str(output_dir)
        env["PYTHONUNBUFFERED"] = "1"
        
        # ADD: Explicitly set these if MongoDB credentials are in environment
        # (They should be inherited above, but we're being explicit)
        for key in os.environ:
            if 'MONGO' in key or 'DB' in key or 'FTP' in key:
                env[key] = os.environ[key]
        
        st.info(f"📍 Running: {script_path.name}")
        st.info(f"📂 Working Dir: {scripts_dir}")
        st.info(f"💾 Output Dir: {output_dir}")
        
        # Run with full environment inheritance
        result = subprocess.run(
            [sys.executable, str(script_path)],
            capture_output=True,
            text=True,
            cwd=str(scripts_dir),
            env=env,  # Pass the environment
            timeout=3600
        )
        
        execution_time = time.time() - start_time
        
        if result.returncode == 0:
            return True, result.stdout, result.stderr, execution_time
        else:
            error_msg = result.stderr if result.stderr else f"Script exited with code {result.returncode}"
            return False, result.stdout, error_msg, execution_time
    
    except subprocess.TimeoutExpired:
        return False, "", "Script execution timed out (>1 hour)", time.time() - start_time
    except Exception as e:
        return False, "", str(e), time.time() - start_time

def verify_script_completion(output_dir, script_index):
    """
    Verify that the output file was actually generated.
    Add specific file patterns here based on what each script produces.
    """
    # Define expected output patterns for each script
    output_patterns = {
        0: ["*FTP*", "*.nc"],  # FTP Download
        1: ["*.xlsx", "*.xls"],  # NC Data (produces Excel)
        2: ["ECM10_MASTER*"],     # ECM10 Master
        3: ["*file_monitoring*"], # File Monitoring
        4: ["*weather_data*"],    # Filtered Weather
        5: ["*filtered_nc*"],     # Filtered NC
        6: ["*match*"],           # WeatherDB Match
        7: ["*prod*"],            # RE Prod
        8: ["*solar*"],           # RE Solar
        9: ["*wind*"],            # RE Wind
        10: ["*sync*"],           # Sync RE
        11: ["*validation*"],     # Data Validation
    }
    
    if script_index < len(output_patterns):
        for pattern in output_patterns[script_index]:
            files = list(output_dir.glob(pattern))
            if files:
                # Check file is not empty
                for f in files:
                    if f.stat().st_size > 0:
                        return True, f"Output verified: {f.name}"
    
    return False, "Output file not found"

# ====================== MAIN UI ======================
st.markdown('<h1 class="main-title">⚡ Automated Regression Suite</h1>', unsafe_allow_html=True)
st.markdown(
    '<p style="font-size:14px;color:#94a3b8;-webkit-text-fill-color:#94a3b8;font-weight:500;margin-bottom:24px;">'
    'Weather Data Processing Pipeline | Advanced Automation</p>',
    unsafe_allow_html=True
)

# ====================== CONFIGURATION ======================
section_header("📁", "Configuration")

NETWORK_SHARE = r"\\172.16.0.65\Share\24_23_QA_Team\Abhishek_Sinha"
SCRIPTS_DIR = r"\\172.16.0.65\Share\24_23_QA_Team\Abhishek_Sinha\Commit_File"

shared_path = st.text_input(
    "📍 Output Folder Path (Network Share)",
    value=NETWORK_SHARE,
    help="Network share path where ALL pipeline scripts save their outputs."
)

scripts_path = st.text_input(
    "📂 Scripts Directory Path",
    value=SCRIPTS_DIR,
    help="Path to folder containing all Python scripts (.py files)."
)

output_dir = Path(shared_path)
scripts_dir = Path(scripts_path)

# Verify network path is accessible
try:
    if output_dir.exists():
        st.success(f"✅ **Connected to Output Directory:** `{shared_path}`")
    else:
        st.error(f"❌ **Output path not accessible:** `{shared_path}`")
        st.stop()
    
    if scripts_dir.exists():
        st.success(f"✅ **Connected to Scripts Directory:** `{scripts_path}`")
    else:
        st.error(f"❌ **Scripts path not accessible:** `{scripts_path}`")
        st.error("Please verify:")
        st.error("  • Scripts directory exists")
        st.error("  • Path is correct: `\\Commit_File`")
        st.error("  • You have read permissions")
        st.stop()
except Exception as e:
    st.error(f"❌ **Cannot access network paths:** {e}")
    st.stop()

st.markdown("---")

# ====================== FTP SETTINGS ======================
section_header("⚙️", "FTP Download Settings")

ftp_col1, ftp_col2 = st.columns([3, 1])
with ftp_col1:
    ftp_file_path = st.text_input(
        "📂 FTP Downloaded File Path",
        value=str(output_dir / "ECM10_FTP"),
        help="Step 1 will be skipped if this exists"
    )
with ftp_col2:
    force_redownload = st.checkbox("🔁 Force Re-download", value=False)

ftp_path = Path(ftp_file_path)
ftp_already_exists = ftp_path.exists() and not force_redownload

if ftp_already_exists:
    st.success("✅ FTP file exists — Step 1 will be skipped")
elif force_redownload:
    st.warning("🔁 Force re-download enabled")
else:
    st.info("📡 Step 1 (FTP Download) will run normally")

st.markdown("---")

# ====================== EXECUTION PARAMETERS ======================
section_header("⚙️", "Execution Parameters")

param_col1, param_col2 = st.columns(2)
with param_col1:
    data_source = st.selectbox("Data Source", ["NWP_E10.0", "NWP_E10.1", "OBS_Data"], index=0)
with param_col2:
    selected_date = st.date_input("Execution Date", value=datetime.date.today())

resume_from = st.session_state.resume_from
if resume_from > 0 and resume_from < len(PIPELINE):
    st.info(f"⚡ **Resume from Step {resume_from + 1}:** {PIPELINE[resume_from][0].strip()}")

button_col1, button_col2 = st.columns([3, 1])
with button_col1:
    run_button = st.button("🚀 RUN / RESUME REGRESSION SUITE", type="primary", use_container_width=True)
with button_col2:
    reset_button = st.button("🔄 Reset", use_container_width=True)

if reset_button:
    st.session_state.statuses = ["pending"] * len(PIPELINE)
    st.session_state.resume_from = 0
    st.session_state.pipeline_completed = False
    st.rerun()

# ====================== PIPELINE STATUS TABLE ======================
st.markdown("---")
section_header("🔄", "Pipeline Execution Status")

status_placeholder = st.empty()

def render_status_table(statuses):
    icon_map = {
        "pending": ("⏳", "step-pending"),
        "running": ("▶️", "step-running"),
        "done":    ("✅", "step-done"),
        "error":   ("❌", "step-error"),
        "skipped": ("⏭️", "step-skipped"),
    }
    
    # Prepare data for the table in a structured format
    table_rows_data = []
    for i, (label, script) in enumerate(PIPELINE):
        status_key = statuses[i] if i < len(statuses) else "pending"
        icon, css_class = icon_map.get(status_key, ("⏳", "step-pending"))
        
        table_rows_data.append({
            "label": label,
            "script": script,
            "status_text": status_key.upper(),
            "icon": icon,
            "css_class": css_class
        })

    # Generate HTML rows from the structured data
    rows_html = ""
    for row_data in table_rows_data:
        rows_html += (
            f"<tr>"
            f"<td style='padding:12px 16px;'><span class='{row_data['css_class']}'>{row_data['icon']} {row_data['label']}</span></td>"
            f"<td style='padding:12px 16px;font-family:monospace;font-size:12px;color:#94a3b8;'>{row_data['script']}</td>"
            f"<td style='padding:12px 16px;'><span class='{row_data['css_class']}'>{row_data['status_text']}</span></td>"
            f"</tr>"
        )
    return (
        "<div class='status-table'>"
        "<table><thead><tr>"
        "<th>Step</th><th>Script</th><th>Status</th>"
    )

status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)

# ====================== RUN PIPELINE ======================
if run_button:
    overall_progress = st.progress(0)
    log_expander = st.expander("📋 Script Output Logs", expanded=False)
    pipeline_failed = False
    start_from = st.session_state.resume_from

    for idx, (label, script_name) in enumerate(PIPELINE):
        if idx < start_from:
            continue

        if idx == 0 and ftp_already_exists:
            st.session_state.statuses[0] = "skipped"
            st.session_state.resume_from = 1
            status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)
            st.info("⏭️ **Step 1 (FTP Download) skipped** — file already exists")
            done_count = sum(1 for s in st.session_state.statuses if s in ("done", "skipped"))
            overall_progress.progress(int(done_count / len(PIPELINE) * 100))
            continue

        # Mark as running
        st.session_state.statuses[idx] = "running"
        status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)

        # Run the script
        success, stdout, stderr, exec_time = run_python_script(script_name, scripts_dir, output_dir)
        
        # Display logs
        with log_expander:
            st.markdown(f"### **Step {idx + 1}: {label}** (`{script_name}`)")
            st.markdown(f"⏱️ **Execution Time:** {exec_time:.2f} seconds")
            
            if stdout:
                st.markdown("**STDOUT:**")
                st.code(stdout, language="text")
            if stderr:
                st.markdown("**STDERR:**")
                st.code(stderr, language="text")

        if not success:
            st.session_state.statuses[idx] = "error"
            st.session_state.resume_from = idx
            status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)
            st.error(f"❌ **Step {idx + 1} failed:** {label}\n\n**Error:**\n```\n{stderr[:500]}\n```\n\nFix the issue and click **RUN / RESUME** to continue.")
            pipeline_failed = True
            break

        # Verify output was actually generated (optional but recommended)
        if idx > 0:  # Skip verification for FTP download
            verified, msg = verify_script_completion(output_dir, idx)
            if not verified:
                st.warning(f"⚠️ {msg} - Script claims success but output not verified")
            else:
                st.success(f"✅ {msg}")

        # Mark as done
        st.session_state.statuses[idx] = "done"
        st.session_state.resume_from = idx + 1
        status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)
        
        done_count = sum(1 for s in st.session_state.statuses if s in ("done", "skipped"))
        overall_progress.progress(int(done_count / len(PIPELINE) * 100))

    if not pipeline_failed:
        overall_progress.progress(100)
        st.session_state.resume_from = 0
        st.session_state.pipeline_completed = True
        st.rerun()

# ====================== FILE FINDER - DYNAMIC FROM NETWORK SHARE ======================
def find_latest_file(pattern):
    """Find the most recently modified file matching pattern in network share."""
    try:
        if not output_dir.exists():
            return None
        
        matching_files = []
        for f in output_dir.iterdir():
            if f.is_file() and pattern.lower() in f.name.lower():
                matching_files.append(f)
        
        if matching_files:
            return max(matching_files, key=lambda x: x.stat().st_mtime)
    except (PermissionError, OSError):
        pass
    
    return None

def get_output_files():
    """Scan network share for the latest output files generated by pipeline."""
    files_found = {
        "file_monitoring": find_latest_file("file_monitoring"),
        "data_validation": find_latest_file("data_validation"),
        "data_sync": find_latest_file("sync_from_re")
    }
    return files_found

def read_csv_safe(filepath):
    """Read CSV with multiple encoding fallbacks."""
    encodings = ['utf-8', 'utf-8-sig', 'cp1252', 'latin-1']
    
    for encoding in encodings:
        try:
            with open(filepath, 'r', encoding=encoding) as f:
                return pd.read_csv(f)
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    return pd.read_csv(filepath, encoding='utf-8', errors='ignore')

def read_excel_safe(filepath):
    """Read Excel file with error handling."""
    try:
        return pd.read_excel(filepath, engine='openpyxl')
    except Exception:
        try:
            return pd.read_excel(filepath, engine='xlrd')
        except Exception as e:
            st.error(f"Cannot read Excel file: {e}")
            return None

# ====================== SUCCESS COMPLETION PANEL ======================
if st.session_state.pipeline_completed:
    st.markdown("""
    <div class="success-panel">
        <div class="success-title">🎉 All 12 Steps Completed Successfully!</div>
        <div class="success-subtitle">Your weather data processing pipeline has finished. Output files are ready below.</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    
    section_header("📊", "Generated Output Files")
    
    output_files = get_output_files()
    
    col1, col2, col3 = st.columns(3)
    
    # FILE 1: FILE MONITORING
    with col1:
        fp = output_files.get("file_monitoring")
        if fp:
            file_size = fp.stat().st_size / (1024 * 1024)
            st.markdown(f"""
            <div class="output-block">
                <div class="output-icon">📁</div>
                <div class="output-name">File Monitoring</div>
                <div class="output-path">{fp.name}</div>
                <div style="font-size:11px;color:#64748b;margin-top:8px;">📍 {str(fp.parent)}</div>
                <div style="font-size:10px;color:#475569;margin-top:4px;">Size: {file_size:.2f} MB</div>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                if fp.suffix.lower() == ".xlsx":
                    df = pd.read_excel(fp)
                else:
                    df = read_csv_safe(fp)
                st.dataframe(df, use_container_width=True, hide_index=True)
                
                with open(fp, "rb") as f:
                    file_bytes = f.read()

                st.download_button(
                    label=f"⬇️ Download {fp.name}",
                    data=file_bytes,
                    file_name=fp.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="fm_download_success",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)[:100]}")
        else:
            st.markdown("""
            <div class="output-block">
                <div class="output-icon">⏳</div>
                <div class="output-name">File Monitoring</div>
                <div class="output-path">Waiting for file...</div>
            </div>
            """, unsafe_allow_html=True)
            st.info("📂 File Monitoring output not found yet.")
    
    # FILE 2: DATA VALIDATION
    with col2:
        fp = output_files.get("data_validation")
        if fp:
            file_size = fp.stat().st_size / (1024 * 1024)
            st.markdown(f"""
            <div class="output-block">
                <div class="output-icon">✔️</div>
                <div class="output-name">Data Validation</div>
                <div class="output-path">{fp.name}</div>
                <div style="font-size:11px;color:#64748b;margin-top:8px;">📍 {str(fp.parent)}</div>
                <div style="font-size:10px;color:#475569;margin-top:4px;">Size: {file_size:.2f} MB</div>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                df = read_excel_safe(fp)
                if df is not None:
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    
                    with open(fp, 'rb') as f:
                        file_bytes = f.read()
                    
                    st.download_button(
                        label="⬇️ Download Data_Validation.xlsx",
                        data=file_bytes,
                        file_name=fp.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dv_download_success",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)[:100]}")
        else:
            st.markdown("""
            <div class="output-block">
                <div class="output-icon">⏳</div>
                <div class="output-name">Data Validation</div>
                <div class="output-path">Waiting for file...</div>
            </div>
            """, unsafe_allow_html=True)
            st.info("📂 Data Validation output not found yet.")
    
    # FILE 3: DATA SYNC
    with col3:
        fp = output_files.get("data_sync")
        if fp:
            file_size = fp.stat().st_size / (1024 * 1024)
            st.markdown(f"""
            <div class="output-block">
                <div class="output-icon">🔄</div>
                <div class="output-name">Data Sync</div>
                <div class="output-path">{fp.name}</div>
                <div style="font-size:11px;color:#64748b;margin-top:8px;">📍 {str(fp.parent)}</div>
                <div style="font-size:10px;color:#475569;margin-top:4px;">Size: {file_size:.2f} MB</div>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                df = read_excel_safe(fp)
                if df is not None:
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    
                    with open(fp, 'rb') as f:
                        file_bytes = f.read()
                    
                    st.download_button(
                        label="⬇️ Download Sync_from_RE.xlsx",
                        data=file_bytes,
                        file_name=fp.name,
                        
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="sync_download_success",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)[:100]}")
        else:
            st.markdown("""
            <div class="output-block">
                <div class="output-icon">⏳</div>
                <div class="output-name">Data Sync</div>
                <div class="output-path">Waiting for file...</div>
            </div>
            """, unsafe_allow_html=True)
            st.info("📂 Data Sync output not found yet.")
    
    st.markdown("---")
    st.info(f"📂 **Output Directory:** `{output_dir}`")

# ====================== FOOTER ======================
st.markdown("---")
st.markdown(f"""
<div style="text-align:center;color:#64748b;font-size:11px;margin-top:24px;">
    <strong style="color:#94a3b8;">Output Directory:</strong>
    <span style="color:#94a3b8;"> {shared_path}</span><br>
    <strong style="color:#94a3b8;">Last Refreshed:</strong>
    <span style="color:#94a3b8;"> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</span><br>
    <em style="color:#64748b;">Weather Auto Regression Suite v1.0 | 50 Hertz Energy</em>
</div>
""", unsafe_allow_html=True)