import sys
import io
import os
import time
import subprocess
import datetime
import glob
from pathlib import Path
from io import BytesIO

import pandas as pd
import streamlit as st


if sys.platform == "win32":
    try:
        if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding != "utf-8":
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", newline="")
        if not isinstance(sys.stderr, io.TextIOWrapper) or sys.stderr.encoding != "utf-8":
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", newline="")
    except Exception:
        pass


st.set_page_config(
    page_title="Weather Auto Regression",
    layout="wide",
    page_icon="⚡",
    initial_sidebar_state="expanded",
)


st.markdown(
    """
<style>
* {
    box-sizing: border-box;
}

html, body {
    scroll-behavior: smooth;
}

#MainMenu, footer, header {
    visibility: hidden;
}

.stApp {
    background: linear-gradient(135deg, #050810 0%, #0a0e1a 25%, #0f172a 50%, #0a0e1a 75%, #050810 100%) !important;
    color: #e2e8f0 !important;
    overflow-x: hidden;
}

/* Animated background gradient */
@keyframes gradientFlow {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

.stApp {
    background-size: 400% 400%;
    animation: gradientFlow 15s ease infinite;
}

.block-container {
    padding-top: 1rem;
    padding-bottom: 1.5rem;
    max-width: 1500px;
    padding-left: 2rem;
    padding-right: 2rem;
}

/* ============ SIDEBAR STYLING ============ */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0a0e1a 0%, #050810 100%) !important;
    border-right: 1px solid rgba(99, 102, 241, 0.1) !important;
}

[data-testid="stSidebar"] > div {
    padding-top: 1.2rem;
}

/* ============ PREMIUM TOPBAR ============ */
.topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 20px;
    padding: 18px 24px;
    background: linear-gradient(135deg, rgba(15, 23, 42, 0.85), rgba(10, 14, 26, 0.85));
    border: 1px solid rgba(99, 102, 241, 0.15);
    border-radius: 16px;
    box-shadow: 
        0 8px 32px rgba(0, 0, 0, 0.4),
        inset 0 1px 1px rgba(255, 255, 255, 0.05);
    margin-bottom: 28px;
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
}

.brand {
    display: flex;
    align-items: center;
    gap: 14px;
    font-size: 18px;
    font-weight: 900;
    color: #f8fafc;
    letter-spacing: -0.3px;
}

.brand-badge {
    width: 44px;
    height: 44px;
    border-radius: 12px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg, #6366f1 0%, #a78bfa 50%, #10b981 100%);
    box-shadow: 
        0 12px 32px rgba(99, 102, 241, 0.35),
        inset 0 1px 2px rgba(255, 255, 255, 0.2);
    font-size: 22px;
    transition: all 300ms cubic-bezier(0.34, 1.56, 0.64, 1);
}

.brand-badge:hover {
    transform: scale(1.05) rotate(5deg);
    box-shadow: 0 16px 40px rgba(99, 102, 241, 0.45);
}

.navlinks {
    display: flex;
    align-items: center;
    gap: 8px;
    color: #cbd5e1;
    font-size: 13px;
    font-weight: 700;
    white-space: nowrap;
    letter-spacing: 0.5px;
}

.nav-item {
    padding: 10px 16px;
    border-radius: 10px;
    background: transparent;
    cursor: pointer;
    transition: all 250ms cubic-bezier(0.34, 1.56, 0.64, 1);
    position: relative;
}

.nav-item::after {
    content: '';
    position: absolute;
    bottom: 0;
    left: 0;
    width: 0;
    height: 2px;
    background: linear-gradient(90deg, #6366f1, #10b981);
    border-radius: 1px;
    transition: width 300ms ease;
}

.nav-item:hover {
    background: rgba(99, 102, 241, 0.12);
    color: #e2e8f0;
    transform: translateY(-2px);
}

.nav-item:hover::after {
    width: 100%;
}

.nav-item.active {
    background: rgba(99, 102, 241, 0.18);
    color: #fff;
    border: 1px solid rgba(99, 102, 241, 0.4);
}

/* ============ PREMIUM HERO SECTION ============ */
.hero {
    background: linear-gradient(135deg, rgba(15, 23, 42, 0.95), rgba(10, 14, 26, 0.95));
    border: 1px solid rgba(99, 102, 241, 0.18);
    border-radius: 18px;
    padding: 32px 36px;
    margin-bottom: 32px;
    box-shadow: 
        0 12px 40px rgba(0, 0, 0, 0.3),
        inset 0 1px 2px rgba(255, 255, 255, 0.05);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    animation: slideInDown 600ms cubic-bezier(0.34, 1.56, 0.64, 1);
}

@keyframes slideInDown {
    from {
        opacity: 0;
        transform: translateY(-20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.hero-title {
    font-size: 36px;
    line-height: 1.15;
    font-weight: 950;
    color: #f8fafc;
    letter-spacing: -0.8px;
    margin-bottom: 10px;
    background: linear-gradient(135deg, #f8fafc 0%, #cbd5e1 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}

.hero-sub {
    color: #94a3b8;
    font-size: 15px;
    margin-bottom: 20px;
    font-weight: 500;
    letter-spacing: 0.3px;
}

.hero-pills {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
}

.pill {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    padding: 8px 16px;
    border-radius: 10px;
    background: linear-gradient(135deg, rgba(99, 102, 241, 0.12), rgba(16, 185, 129, 0.08));
    border: 1px solid rgba(99, 102, 241, 0.25);
    color: #cbd5e1;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    box-shadow: 0 4px 12px rgba(99, 102, 241, 0.08);
    transition: all 250ms ease;
    cursor: default;
}

.pill:hover {
    background: linear-gradient(135deg, rgba(99, 102, 241, 0.18), rgba(16, 185, 129, 0.12));
    border-color: rgba(99, 102, 241, 0.35);
    transform: translateY(-2px);
    box-shadow: 0 6px 16px rgba(99, 102, 241, 0.12);
}

/* ============ METRIC CARDS - PREMIUM ============ */
.metric-card {
    background: linear-gradient(135deg, rgba(15, 23, 42, 0.98), rgba(20, 30, 48, 0.95));
    border: 1px solid rgba(99, 102, 241, 0.15);
    border-left: 5px solid var(--accent-color, #6366f1);
    border-radius: 14px;
    padding: 22px 24px;
    min-height: 140px;
    box-shadow: 
        0 8px 24px rgba(0, 0, 0, 0.25),
        inset 0 1px 1px rgba(255, 255, 255, 0.03);
    transition: all 300ms cubic-bezier(0.34, 1.56, 0.64, 1);
    position: relative;
    overflow: hidden;
}

.metric-card::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -50%;
    width: 200px;
    height: 200px;
    background: radial-gradient(circle, var(--accent-color, #6366f1) 0%, transparent 70%);
    opacity: 0.05;
    transition: all 400ms ease;
}

.metric-card:hover {
    border-color: rgba(99, 102, 241, 0.35);
    transform: translateY(-6px);
    box-shadow: 
        0 16px 40px rgba(99, 102, 241, 0.15),
        inset 0 1px 1px rgba(255, 255, 255, 0.05);
}

.metric-card:hover::before {
    top: -20%;
    right: -20%;
    opacity: 0.1;
}

.metric-label {
    color: #94a3b8;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    margin-bottom: 10px;
    font-weight: 800;
}

.metric-value {
    color: #f8fafc;
    font-size: 42px;
    font-weight: 950;
    line-height: 1;
    letter-spacing: -1px;
    margin-bottom: 8px;
}

.metric-desc {
    color: #64748b;
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 0.2px;
}

/* ============ SECTION HEADERS ============ */
.section-title {
    color: #f8fafc;
    font-size: 20px;
    font-weight: 850;
    display: flex;
    align-items: center;
    gap: 14px;
    margin: 12px 0 18px 0;
    letter-spacing: -0.3px;
}

.section-bar {
    width: 5px;
    height: 26px;
    border-radius: 5px;
    background: linear-gradient(180deg, #6366f1 0%, #10b981 100%);
    box-shadow: 0 4px 12px rgba(99, 102, 241, 0.25);
}

/* ============ GLASS CARDS ============ */
.glass-card, .panel {
    background: linear-gradient(135deg, rgba(15, 23, 42, 0.95), rgba(20, 30, 48, 0.90));
    border: 1px solid rgba(99, 102, 241, 0.15);
    border-radius: 14px;
    box-shadow: 
        0 8px 24px rgba(0, 0, 0, 0.25),
        inset 0 1px 1px rgba(255, 255, 255, 0.03);
    padding: 20px 24px;
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    transition: all 300ms ease;
}

.glass-card:hover {
    border-color: rgba(99, 102, 241, 0.25);
    box-shadow: 
        0 12px 32px rgba(99, 102, 241, 0.12),
        inset 0 1px 1px rgba(255, 255, 255, 0.05);
    transform: translateY(-2px);
}

/* ============ STATUS PILLS ============ */
.status-pill {
    padding: 7px 14px;
    border-radius: 10px;
    font-size: 11px;
    font-weight: 800;
    display: inline-block;
    white-space: nowrap;
    letter-spacing: 0.6px;
    text-transform: uppercase;
    transition: all 250ms ease;
}

.pill-pending {
    background: linear-gradient(135deg, rgba(107, 114, 128, 0.15), rgba(107, 114, 128, 0.08));
    color: #d1d5db;
    border: 1px solid rgba(107, 114, 128, 0.25);
}

.pill-running {
    background: linear-gradient(135deg, rgba(59, 130, 246, 0.15), rgba(59, 130, 246, 0.08));
    color: #93c5fd;
    border: 1px solid rgba(59, 130, 246, 0.4);
    animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
    box-shadow: 0 0 12px rgba(59, 130, 246, 0.2);
}

.pill-done {
    background: linear-gradient(135deg, rgba(16, 185, 129, 0.15), rgba(16, 185, 129, 0.08));
    color: #86efac;
    border: 1px solid rgba(16, 185, 129, 0.4);
}

.pill-error {
    background: linear-gradient(135deg, rgba(239, 68, 68, 0.15), rgba(239, 68, 68, 0.08));
    color: #fca5a5;
    border: 1px solid rgba(239, 68, 68, 0.4);
}

.pill-skipped {
    background: linear-gradient(135deg, rgba(139, 92, 246, 0.15), rgba(139, 92, 246, 0.08));
    color: #d8b4fe;
    border: 1px solid rgba(139, 92, 246, 0.4);
}

@keyframes pulse {
    0%, 100% { 
        opacity: 1;
        box-shadow: 0 0 12px rgba(59, 130, 246, 0.2);
    }
    50% { 
        opacity: 0.8;
        box-shadow: 0 0 20px rgba(59, 130, 246, 0.35);
    }
}

/* ============ PREMIUM BUTTONS ============ */
.stButton > button {
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    font-weight: 800 !important;
    height: 45px !important;
    box-shadow: 0 12px 32px rgba(99, 102, 241, 0.3) !important;
    font-size: 14px !important;
    letter-spacing: 0.3px !important;
    transition: all 300ms cubic-bezier(0.34, 1.56, 0.64, 1) !important;
}

.stButton > button:hover {
    background: linear-gradient(135deg, #4f46e5 0%, #4338ca 100%) !important;
    box-shadow: 0 16px 40px rgba(99, 102, 241, 0.45) !important;
    transform: translateY(-3px) !important;
}

.stButton > button:active {
    transform: scale(0.97) !important;
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    font-weight: 800 !important;
    height: 45px !important;
    box-shadow: 0 12px 32px rgba(16, 185, 129, 0.3) !important;
}

.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
    box-shadow: 0 16px 40px rgba(16, 185, 129, 0.4) !important;
    transform: translateY(-3px) !important;
}

/* ============ PREMIUM INPUT FIELDS ============ */
.stTextInput > div > div > input,
.stNumberInput input,
.stDateInput input,
.stSelectbox [data-baseweb="select"] > div {
    background-color: rgba(15, 23, 42, 0.95) !important;
    color: #f1f5f9 !important;
    border: 1px solid rgba(99, 102, 241, 0.2) !important;
    border-radius: 12px !important;
    padding: 12px 16px !important;
    font-weight: 500 !important;
    font-size: 14px !important;
    transition: all 300ms ease !important;
}

.stTextInput > div > div > input:focus,
.stNumberInput input:focus,
.stDateInput input:focus {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.15) !important;
    background-color: rgba(20, 30, 48, 0.95) !important;
}

.stNumberInput button {
    background-color: rgba(51, 65, 85, 0.8) !important;
    color: #cbd5e1 !important;
    border: 1px solid rgba(99, 102, 241, 0.2) !important;
    border-radius: 10px !important;
    transition: all 200ms ease !important;
}

.stNumberInput button:hover {
    background-color: rgba(71, 85, 105, 0.9) !important;
    color: #f1f5f9 !important;
    border-color: #6366f1 !important;
}

/* ============ LABELS & TEXT ============ */
label, label p, label span {
    color: #cbd5e1 !important;
    font-weight: 700 !important;
    font-size: 13px !important;
    letter-spacing: 0.4px !important;
}

.section-label {
    color: #6366f1 !important;
    font-weight: 800 !important;
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: 1.2px !important;
    margin-bottom: 14px !important;
}

/* ============ SIDEBAR TEXT ============ */
[data-testid="stSidebar"] *,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span {
    color: #e2e8f0 !important;
}

[data-testid="stSidebar"] input,
[data-testid="stSidebar"] textarea {
    color: #f1f5f9 !important;
    opacity: 1 !important;
}

/* ============ DATA TABLES (native st.dataframe wrapper only) ============ */
div[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
    border: 1px solid rgba(99, 102, 241, 0.15);
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
}

/* ============ CUSTOM STYLED HTML TABLES (used for output file previews) ============ */
.styled-table-wrap {
    overflow-x: auto;
    max-height: 520px;
    overflow-y: auto;
}

.styled-table-wrap table {
    width: 100%;
    border-collapse: collapse;
}

.styled-table-wrap thead tr {
    position: sticky;
    top: 0;
    z-index: 1;
}

.styled-table-wrap th {
    padding: 12px 14px;
    text-align: left;
    color: #94a3b8;
    text-transform: uppercase;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 0.6px;
    background: rgba(15, 23, 42, 0.98);
    border-bottom: 1px solid rgba(99, 102, 241, 0.2);
    white-space: nowrap;
}

.styled-table-wrap td {
    padding: 10px 14px;
    color: #e2e8f0;
    font-size: 13px;
    font-weight: 500;
    border-bottom: 1px solid rgba(99, 102, 241, 0.08);
    white-space: nowrap;
}

.styled-table-wrap tr:hover td {
    background-color: rgba(99, 102, 241, 0.08);
}

.styled-table-wrap tr:nth-child(even) td {
    background-color: rgba(255, 255, 255, 0.015);
}

/* Scrollbar theming for the table wrap */
.styled-table-wrap::-webkit-scrollbar {
    height: 8px;
    width: 8px;
}

.styled-table-wrap::-webkit-scrollbar-track {
    background: rgba(15, 23, 42, 0.4);
    border-radius: 8px;
}

.styled-table-wrap::-webkit-scrollbar-thumb {
    background: rgba(99, 102, 241, 0.4);
    border-radius: 8px;
}

.styled-table-wrap::-webkit-scrollbar-thumb:hover {
    background: rgba(99, 102, 241, 0.6);
}

/* ============ EXCEL PREVIEW (real workbook formatting, incl. ============ */
/* ============ conditional-formatting highlights, via LibreOffice) ======== */
.excel-preview-wrap {
    padding: 16px !important;
}

.excel-preview-scroll {
    overflow-x: auto;
    overflow-y: auto;
    max-height: 560px;
    border-radius: 10px;
    background: #ffffff;
    padding: 4px;
}

.excel-preview-scroll table {
    border-collapse: collapse;
    font-family: Calibri, Arial, sans-serif !important;
    font-size: 13px !important;
}

.excel-preview-scroll td,
.excel-preview-scroll th {
    border: 1px solid #d4d7dd;
    padding: 6px 10px !important;
    white-space: nowrap;
}

.excel-preview-scroll tr:first-child td,
.excel-preview-scroll tr:first-child th {
    position: sticky;
    top: 0;
    z-index: 1;
}

.excel-preview-scroll::-webkit-scrollbar {
    height: 10px;
    width: 10px;
}

.excel-preview-scroll::-webkit-scrollbar-track {
    background: #eef0f4;
    border-radius: 8px;
}

.excel-preview-scroll::-webkit-scrollbar-thumb {
    background: #a5acba;
    border-radius: 8px;
}

.excel-preview-scroll::-webkit-scrollbar-thumb:hover {
    background: #8b93a3;
}

/* ============ PREMIUM SIDEBAR HEADER ============ */
.sidebar-header {
    padding: 20px 18px;
    margin-bottom: 20px;
    border-radius: 14px;
    background: linear-gradient(135deg, rgba(99, 102, 241, 0.18), rgba(16, 185, 129, 0.12));
    border: 1px solid rgba(99, 102, 241, 0.25);
    box-shadow: 0 8px 24px rgba(99, 102, 241, 0.15);
    backdrop-filter: blur(10px);
    -webkit-backdrop-filter: blur(10px);
}

.sidebar-header-icon {
    width: 48px;
    height: 48px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg, #6366f1, #10b981);
    font-size: 26px;
    box-shadow: 0 10px 28px rgba(99, 102, 241, 0.35);
}

.sidebar-header-title {
    font-size: 32px;
    font-weight: 950;
    color: #ffffff;
    letter-spacing: 0.5px;
}

.sidebar-header-subtitle {
    font-size: 12px;
    font-weight: 700;
    color: #cbd5e1;
    letter-spacing: 0.3px;
}

/* ============ FILE CHIPS ============ */
.file-chip {
    display: inline-block;
    padding: 7px 14px;
    border-radius: 10px;
    background: linear-gradient(135deg, rgba(99, 102, 241, 0.15), rgba(99, 102, 241, 0.08));
    border: 1px solid rgba(99, 102, 241, 0.25);
    color: #a5b4fc;
    font-size: 12px;
    font-weight: 700;
    margin: 0 8px 8px 0;
    transition: all 250ms ease;
}

.file-chip:hover {
    background: linear-gradient(135deg, rgba(99, 102, 241, 0.22), rgba(99, 102, 241, 0.12));
    border-color: rgba(99, 102, 241, 0.4);
}

.small-muted {
    color: #94a3b8;
    font-size: 12px;
    font-weight: 500;
    letter-spacing: 0.2px;
}

/* ============ DIVIDERS ============ */
hr {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(99, 102, 241, 0.2), transparent);
    margin: 24px 0 !important;
}

/* ============ SMOOTH ANIMATIONS ============ */
@keyframes fadeIn {
    from {
        opacity: 0;
    }
    to {
        opacity: 1;
    }
}

@keyframes slideIn {
    from {
        opacity: 0;
        transform: translateY(12px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

@keyframes slideInLeft {
    from {
        opacity: 0;
        transform: translateX(-20px);
    }
    to {
        opacity: 1;
        transform: translateX(0);
    }
}

.metric-card {
    animation: slideIn 500ms cubic-bezier(0.34, 1.56, 0.64, 1);
}

.glass-card {
    animation: slideIn 600ms cubic-bezier(0.34, 1.56, 0.64, 1);
}

/* ============ RESPONSIVE DESIGN ============ */
@media (max-width: 1024px) {
    .topbar {
        flex-direction: column;
        gap: 14px;
    }
    
    .navlinks {
        width: 100%;
        justify-content: center;
    }
    
    .block-container {
        padding-left: 1rem;
        padding-right: 1rem;
    }
    
    .hero-title {
        font-size: 28px;
    }
}

@media (max-width: 640px) {
    .hero {
        padding: 24px 20px;
    }
    
    .hero-title {
        font-size: 24px;
    }
    
    .metric-card {
        min-height: 120px;
        padding: 16px 18px;
    }
}
</style>
""",
    unsafe_allow_html=True,
)


NETWORK_SHARE = r"\\172.16.0.65\Share\24_23_QA_Team\Abhishek_Sinha"
SCRIPTS_DIR = r"\\172.16.0.65\Share\24_23_QA_Team\Abhishek_Sinha\Commit_File"

PIPELINE = [
    ("1/12 FTP Download", "ECM10_FTP_DOWNLOAD.py"),
    ("2/12 NC Data ECM10", "Nc_data_ECM_10(ECM10).py"),
    ("3/12 ECM10 Master", "ECM10_MASTER.py"),
    ("4/12 File Monitoring", "File_Monitoring(ECM10_.py"),
    ("5/12 Filtered Weather Data", "Filtered_Weather_Data(ECM10).py"),
    ("6/12 Filtered NC File", "FilteredNc_File(ECM10).py"),
    ("7/12 WeatherDB NcFile Match", "WeatherDB_NcFileMatch(ECM10).py"),
    ("8/12 RE Prod ECM10", "REProdECM10.py"),
    ("9/12 RE Solar ECM10", "RESOLAR_ECM10.py"),
    ("10/12 RE Wind ECM10", "REWIND_ECM10.py"),
    ("11/12 Sync RE DB", "SYNC_RE_DB.py"),
    ("12/12 Data Validation", "Data_Validation.py"),
]

for k, v in {
    "statuses": ["pending"] * len(PIPELINE),
    "resume_from": 0,
    "pipeline_completed": False,
    "latlon_df": None,
    "mismatch_df": None,
    "shared_path": NETWORK_SHARE,
    "scripts_path": SCRIPTS_DIR,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


def section_header(icon, text):
    st.markdown(
        f"""
        <div class="section-title">
            <span class="section-bar"></span>
            <span>{icon} {text}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def find_script(script_name, scripts_dir):
    if (scripts_dir / script_name).exists():
        return scripts_dir / script_name
    if Path(script_name).exists():
        return Path(script_name)
    if "(" in script_name and ")" in script_name:
        base = script_name.replace("(", "").replace(")", "")
        matches = glob.glob(str(scripts_dir / base.replace(".py", "*.py")))
        if matches:
            return Path(matches[0])
    return None


def run_python_script(script_name, scripts_dir, output_dir):
    start = time.time()
    try:
        script_path = find_script(script_name, scripts_dir)
        if not script_path or not script_path.exists():
            return False, "", f"Script not found: {script_name}", 0
        env = os.environ.copy()
        env["OUTPUT_DIR"] = str(output_dir)
        env["PYTHONUNBUFFERED"] = "1"
        result = subprocess.run(
            [sys.executable, str(script_path)],
            capture_output=True,
            text=True,
            cwd=str(scripts_dir),
            env=env,
            timeout=3600,
        )
        elapsed = time.time() - start
        if result.returncode == 0:
            return True, result.stdout, result.stderr, elapsed
        return False, result.stdout, result.stderr or f"Exit code {result.returncode}", elapsed
    except subprocess.TimeoutExpired:
        return False, "", "Script execution timed out", time.time() - start
    except Exception as e:
        return False, "", str(e), time.time() - start


def verify_script_completion(output_dir, idx):
    patterns = {
        0: ["*FTP*", "*.nc"],
        1: ["*.xlsx", "*.xls"],
        2: ["*MASTER*"],
        3: ["*file_monitoring*"],
        4: ["*weather_data*"],
        5: ["*filtered_nc*"],
        6: ["*match*"],
        7: ["*prod*"],
        8: ["*solar*"],
        9: ["*wind*"],
        10: ["*sync*"],
        11: ["*validation*"],
    }
    for pat in patterns.get(idx, []):
        for f in output_dir.glob(pat):
            if f.is_file() and f.stat().st_size > 0:
                return True, "Output file verified successfully"
    return False, "Output file not verified"


def find_latest_file(output_dir, keyword):
    try:
        files = [f for f in output_dir.iterdir() if f.is_file() and keyword.lower() in f.name.lower()]
        return max(files, key=lambda x: x.stat().st_mtime) if files else None
    except Exception:
        return None


def read_any_file(path):
    if path.suffix.lower() in [".xlsx", ".xls"]:
        try:
            # Read the file normally first
            df = pd.read_excel(path, engine="openpyxl", header=0)
        except Exception:
            df = pd.read_excel(path, engine="xlrd", header=0)
        
        # If columns are "Unnamed", use the second row as actual headers
        if any(str(col).startswith('Unnamed') for col in df.columns):
            try:
                # Read with row 1 (second row) as the header
                df = pd.read_excel(path, engine="openpyxl" if path.suffix.lower() in [".xlsx", ".xlsm"] else "xlrd", header=1)
            except Exception:
                pass
    else:
        for enc in ["utf-8", "utf-8-sig", "cp1252", "latin-1"]:
            try:
                df = pd.read_csv(path, encoding=enc)
                break
            except Exception:
                pass
        else:
            df = pd.read_csv(path, encoding="latin-1")
    
    return df


def df_to_excel_bytes(sheets):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            safe = str(name)[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
    bio.seek(0)
    return bio.getvalue()


def detect_latlon_columns(df):
    cols = list(df.columns)
    lat = next((c for c in cols if c.lower() in ["lat", "latitude", "lat_dd"]), None)
    lon = next((c for c in cols if c.lower() in ["lon", "lng", "long", "longitude", "lon_dd"]), None)
    return lat, lon


def mismatch_rows(df, lat_col, lon_col):
    temp = df.copy()
    temp[lat_col] = pd.to_numeric(temp[lat_col], errors="coerce")
    temp[lon_col] = pd.to_numeric(temp[lon_col], errors="coerce")
    bad = temp[temp[lat_col].isna() | temp[lon_col].isna()].copy()
    if not bad.empty:
        bad["Mismatch_Reason"] = "Invalid or missing Lat/Long"
    return bad


def render_status_pill(status):
    cls = {
        "pending": "pill-pending",  
        "running": "pill-running",
        "done": "pill-done",
        "error": "pill-error",
        "skipped": "pill-skipped",
    }.get(status, "pill-pending")
    return f"<span class='status-pill {cls}'>{status.upper()}</span>"


import html as _html_lib
import re
import subprocess as _subprocess
import tempfile
import shutil


def _convert_xlsx_to_html_table(xlsx_path):
    """
    Convert an .xlsx file to HTML using headless LibreOffice. LibreOffice
    fully evaluates conditional-formatting rules (the green/red highlight
    cells your pipeline scripts apply), so the resulting <table> markup
    contains the *actual* computed bgcolor/font color for every cell --
    something openpyxl cannot give us directly, since it only reports the
    static formatting rules, not their evaluated result.

    Returns the inner <table>...</table> HTML string, or None on failure.
    """
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        return None

    with tempfile.TemporaryDirectory() as tmpdir:
        local_copy = Path(tmpdir) / xlsx_path.name
        try:
            shutil.copy2(xlsx_path, local_copy)
        except Exception:
            return None

        try:
            result = _subprocess.run(
                [
                    "soffice", "--headless", "--norestore",
                    "--convert-to", "html:HTML (StarCalc)",
                    "--outdir", tmpdir, str(local_copy),
                ],
                capture_output=True, text=True, timeout=90,
            )
        except Exception:
            return None

        if result.returncode != 0:
            return None

        html_path = local_copy.with_suffix(".html")
        if not html_path.exists():
            return None

        try:
            raw_html = html_path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return None

    match = re.search(r"<table[^>]*>.*</table>", raw_html, re.DOTALL | re.IGNORECASE)
    if not match:
        return None
    table_html = match.group(0)

    table_html = re.sub(r'\sdata-sheets-value="[^"]*"', "", table_html)
    table_html = re.sub(r"<a class=\"comment-indicator\".*?</a>", "", table_html, flags=re.DOTALL)
    table_html = re.sub(r"<comment>.*?</comment>", "", table_html, flags=re.DOTALL)

    return table_html


def render_excel_with_formatting(path, max_rows=None):
    """
    Render an .xlsx file's first sheet as HTML, preserving its native
    cell colors (including evaluated conditional formatting) by running it
    through headless LibreOffice. Wrapped in the dashboard's glass-card
    container with a light inner background, since most Excel highlight
    colors (white/green/red/yellow) are designed for light backgrounds.
    Returns None if conversion isn't possible, so callers can fall back.
    """
    table_html = _convert_xlsx_to_html_table(path)
    if table_html is None:
        return None

    return f"""
    <div class="glass-card excel-preview-wrap">
        <div class="excel-preview-scroll">
            {table_html}
        </div>
    </div>
    """


def render_styled_dataframe(df, max_rows=300):
    """
    Render a pandas DataFrame as a styled HTML table that matches the
    dashboard's dark glass theme. Used as a fallback when the source file
    isn't an .xlsx (so there's no native Excel formatting to preserve),
    e.g. .csv outputs or in-session data like the Lat/Long table.
    """
    if df is None or df.empty:
        return "<div class='glass-card'><span class='small-muted'>No data to display.</span></div>"

    display_df = df.head(max_rows)

    header_html = "".join(
        f"<th>{_html_lib.escape(str(col))}</th>" for col in display_df.columns
    )

    body_rows = []
    for _, row in display_df.iterrows():
        cells = []
        for v in row:
            if pd.isna(v):
                cells.append("<td></td>")
            else:
                cells.append(f"<td>{_html_lib.escape(str(v))}</td>")
        body_rows.append(f"<tr>{''.join(cells)}</tr>")
    rows_html = "".join(body_rows)

    truncated_note = ""
    if len(df) > max_rows:
        truncated_note = (
            f"<div class='small-muted' style='margin-top:10px;'>"
            f"Showing first {max_rows} of {len(df)} rows. Use the download "
            f"button for the full dataset.</div>"
        )

    return f"""
    <div class="glass-card styled-table-wrap">
        <table>
            <thead>
                <tr>{header_html}</tr>
            </thead>
            <tbody>{rows_html}</tbody>
        </table>
    </div>
    {truncated_note}
    """


def render_file_preview(path, df, max_rows=300):
    """
    Top-level helper used by the output tabs: tries to preserve the
    original Excel coloring (fills + evaluated conditional-formatting
    highlights) by converting the file with LibreOffice; falls back to the
    plain dark-themed table renderer for non-Excel files or on any
    conversion error (e.g. LibreOffice unavailable, file locked).
    """
    if path is not None and Path(path).suffix.lower() in (".xlsx", ".xlsm"):
        html_table = render_excel_with_formatting(path)
        if html_table is not None:
            return html_table
    return render_styled_dataframe(df, max_rows=max_rows)


st.markdown(
    """
<div class="topbar">
    <div class="brand">
        <span class="brand-badge">⚡</span>
        <span>Weather Auto Regression</span>
    </div>
    <div class="navlinks">
        <span class="nav-item active">Dashboard</span>
        <span class="nav-item">Pipeline</span>
        <span class="nav-item">Outputs</span>
        <span class="nav-item">Logs</span>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="hero">
    <div class="hero-title">Automated Regression Suite</div>
    <div class="hero-sub">Weather Data Processing Pipeline | Production-Grade Dashboard</div>
    <div class="hero-pills">
        <span class="pill">⚡ Weather Auto v3.0</span>
        <span class="pill">🔄 Multi-Step Pipeline</span>
        <span class="pill">🗺️ Lat/Long Review</span>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("""
    <div style="
        display: flex;
        align-items: center;
        gap: 14px;
        padding: 20px 18px;
        margin-bottom: 20px;
        border-radius: 14px;
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.18), rgba(16, 185, 129, 0.12));
        border: 1px solid rgba(99, 102, 241, 0.25);
        box-shadow: 0 8px 24px rgba(99, 102, 241, 0.15);
    ">
        <div style="
            width: 48px;
            height: 48px;
            border-radius: 12px;
            background: linear-gradient(135deg, #6366f1, #10b981);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 26px;
            box-shadow: 0 10px 28px rgba(99, 102, 241, 0.35);
        ">⚡</div>
        <div>
            <div style="font-size: 32px; font-weight: 950; color: #ffffff; letter-spacing: 0.5px;">50 HERTZ</div>
            <div style="font-size: 12px; font-weight: 700; color: #cbd5e1; letter-spacing: 0.3px;">Energy Automation</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    
    st.markdown("<p class='section-label'>📁 Output Configuration</p>", unsafe_allow_html=True)
    shared_path = st.text_input("Output Path", value=st.session_state.shared_path, key="shared_input")
    st.session_state.shared_path = shared_path
    
    scripts_path = st.text_input("Scripts Path", value=st.session_state.scripts_path, key="scripts_input")
    st.session_state.scripts_path = scripts_path
    
    st.markdown("---")
    
    st.markdown("<p class='section-label'>🎯 Data & Execution</p>", unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])
    with col1:
        data_source = st.selectbox("Data Source", ["NWP_E10.0", "ALLNCM_12.5"], index=0)
    with col2:
        selected_date = st.date_input("Exec Date", value=datetime.date.today(), label_visibility="collapsed")
    
    st.markdown("---")
    
    st.markdown("<p class='section-label'>⚙️ Pipeline Control</p>", unsafe_allow_html=True)
    force_redownload = st.checkbox("Force Re-download", value=False)
    
    st.markdown("---")

output_dir = Path(st.session_state.shared_path)
scripts_dir = Path(st.session_state.scripts_path)

if not output_dir.exists():
    st.error(f"❌ Output path not accessible: {st.session_state.shared_path}")
    st.stop()
if not scripts_dir.exists():
    st.error(f"❌ Scripts path not accessible: {st.session_state.scripts_path}")
    st.stop()

total_steps = len(PIPELINE)
done_steps = sum(1 for s in st.session_state.statuses if s in ("done", "skipped"))
running_steps = sum(1 for s in st.session_state.statuses if s == "running")
error_steps = sum(1 for s in st.session_state.statuses if s == "error")

m1, m2, m3, m4 = st.columns(4)

metric_configs = [
    (m1, "Total Steps", total_steps, "Pipeline length", "#3b82f6"),
    (m2, "Completed", done_steps, "Done + skipped", "#10b981"),
    (m3, "Running", running_steps, "Active step", "#f59e0b"),
    (m4, "Errors", error_steps, "Failed steps", "#ef4444"),
]

for col, label, value, desc, accent_color in metric_configs:
    with col:
        st.markdown(
            f"""
            <div class="metric-card" style="--accent-color: {accent_color}">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
                <div class="metric-desc">{desc}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.markdown("---")

section_header("📁", "Configuration")
c1, c2, c3 = st.columns([3, 1.2, 1])
with c1:
    ftp_file_path = st.text_input("FTP File Path", value=str(output_dir / "ECM10_FTP"))
with c2:
    start_from = st.number_input("Resume", min_value=0, max_value=len(PIPELINE), value=st.session_state.resume_from, step=1)
    st.session_state.resume_from = int(start_from)
with c3:
    st.metric("Pipeline", f"{done_steps}/{total_steps}")

ftp_path = Path(ftp_file_path)
ftp_already_exists = ftp_path.exists() and not force_redownload

section_header("🗺️", "Lat/Long Configuration")
ll1, ll2, ll3 = st.columns([1, 1, 0.8])
with ll1:
    lat_value = st.number_input("Latitude", value=28.6139, format="%.6f", step=0.0001)
with ll2:
    lon_value = st.number_input("Longitude", value=77.2090, format="%.6f", step=0.0001)
with ll3:
    save_latlon = st.button("Save", use_container_width=True)

if save_latlon:
    st.session_state.latlon_df = pd.DataFrame([{"Latitude": lat_value, "Longitude": lon_value}])
    st.success("✅ Lat/Long saved to session.")

if st.session_state.latlon_df is not None:
    st.markdown(render_styled_dataframe(st.session_state.latlon_df), unsafe_allow_html=True)

section_header("🔄", "Pipeline Execution Status")
status_placeholder = st.empty()

def render_status_table(statuses):
    rows = ""
    for i, (label, script) in enumerate(PIPELINE):
        s = statuses[i]
        rows += (
            f"<tr>"
            f"<td style='padding:12px 14px;'>{label}</td>"
            f"<td style='padding:12px 14px;font-family:monospace;color:#cbd5e1;font-size:12px;'>{script}</td>"
            f"<td style='padding:12px 14px;'>{render_status_pill(s)}</td>"
            f"</tr>"
        )
    return f"""
    <div class="glass-card">
        <table style="width:100%;border-collapse:collapse;">
            <thead>
                <tr style="color:#94a3b8;text-align:left;border-bottom:1px solid rgba(99, 102, 241, 0.15);">
                    <th style="padding:12px 14px;font-weight:700;text-transform:uppercase;font-size:11px;letter-spacing:0.5px;">Step</th>
                    <th style="padding:12px 14px;font-weight:700;text-transform:uppercase;font-size:11px;letter-spacing:0.5px;">Script</th>
                    <th style="padding:12px 14px;font-weight:700;text-transform:uppercase;font-size:11px;letter-spacing:0.5px;">Status</th>
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>
    </div>
    """

status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)

run_col, reset_col = st.columns([3, 1])
with run_col:
    run_button = st.button("🚀 RUN / RESUME REGRESSION SUITE", use_container_width=True, type="primary")
with reset_col:
    reset_button = st.button("🔄 Reset", use_container_width=True)

if reset_button:
    st.session_state.statuses = ["pending"] * len(PIPELINE)
    st.session_state.resume_from = 0
    st.session_state.pipeline_completed = False
    st.session_state.mismatch_df = None
    st.rerun()

progress = st.progress(0, text="Pipeline ready")

if run_button:
    logs = st.container()
    failed = False

    for idx, (label, script_name) in enumerate(PIPELINE):
        if idx < st.session_state.resume_from:
            continue

        if idx == 0 and ftp_already_exists:
            st.session_state.statuses[idx] = "skipped"
            st.session_state.resume_from = 1
            status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)
            progress.progress(int((1 / total_steps) * 100), text=f"Skipped: {label}")
            continue

        st.session_state.statuses[idx] = "running"
        status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)
        progress.progress(int((idx / total_steps) * 100), text=f"Running: {label}")

        with logs:
            with st.expander(f"Step {idx+1}: {label}", expanded=True):
                st.info(f"Executing {script_name}")

        success, stdout, stderr, exec_time = run_python_script(script_name, scripts_dir, output_dir)

        with logs:
            with st.expander(f"Logs: {label}", expanded=False):
                st.write(f"Execution time: {exec_time:.2f} sec")
                if stdout:
                    st.code(stdout, language="text")
                if stderr:
                    st.code(stderr, language="text")

        if not success:
            st.session_state.statuses[idx] = "error"
            st.session_state.resume_from = idx
            status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)
            st.error(f"❌ Step failed: {label}")
            failed = True
            break



        st.session_state.statuses[idx] = "done"
        st.session_state.resume_from = idx + 1
        status_placeholder.markdown(render_status_table(st.session_state.statuses), unsafe_allow_html=True)
        progress.progress(int(((idx + 1) / total_steps) * 100), text=f"Completed: {label}")

    if not failed:
        st.session_state.pipeline_completed = True
        st.session_state.resume_from = 0
        progress.progress(100, text="Pipeline completed successfully")
        st.balloons()


def get_output_files():
    return {
        "file_monitoring": find_latest_file(output_dir, "file_monitoring"),
        "data_validation": find_latest_file(output_dir, "data_validation"),
        "data_sync": find_latest_file(output_dir, "sync_from_re"),
        "mismatch": find_latest_file(output_dir, "mismatch"),
    }


if st.session_state.pipeline_completed:
    section_header("📊", "Generated Output Files")
    files = get_output_files()
    tabs = st.tabs(["File Monitoring", "Data Validation", "Data Sync", "Mismatch Data", "Lat/Long"])

    with tabs[0]:
        fp = files["file_monitoring"]
        if fp:
            df = read_any_file(fp)
            st.markdown(f"<div class='file-chip'>{fp.name}</div>", unsafe_allow_html=True)
            st.markdown(render_file_preview(fp, df), unsafe_allow_html=True)
            with open(fp, "rb") as f:
                st.download_button(
                    "Download Excel",
                    f.read(),
                    file_name=fp.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("No File Monitoring file found.")

    with tabs[1]:
        fp = files["data_validation"]
        if fp:
            df = read_any_file(fp)
            st.markdown(f"<div class='file-chip'>{fp.name}</div>", unsafe_allow_html=True)
            st.markdown(render_file_preview(fp, df), unsafe_allow_html=True)
            with open(fp, "rb") as f:
                st.download_button(
                    "Download Excel",
                    f.read(),
                    file_name=fp.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("No Data Validation file found.")

    with tabs[2]:
        fp = files["data_sync"]
        if fp:
            df = read_any_file(fp)
            st.markdown(f"<div class='file-chip'>{fp.name}</div>", unsafe_allow_html=True)
            st.markdown(render_file_preview(fp, df), unsafe_allow_html=True)
            with open(fp, "rb") as f:
                st.download_button(
                    "Download Excel",
                    f.read(),
                    file_name=fp.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("No Sync file found.")

    with tabs[3]:
        fp = files["mismatch"]
        if fp:
            df = read_any_file(fp).copy()
            df.columns = [str(c).strip() for c in df.columns]
            lat_col, lon_col = detect_latlon_columns(df)

            st.markdown(f"<div class='file-chip'>{fp.name}</div>", unsafe_allow_html=True)
            st.markdown(render_file_preview(fp, df), unsafe_allow_html=True)

            if lat_col and lon_col:
                bad = mismatch_rows(df, lat_col, lon_col)
                st.session_state.mismatch_df = bad
                if not bad.empty:
                    st.warning(f"⚠️ {len(bad)} mismatch rows found in Lat/Long.")
                else:
                    st.success("✅ No Lat/Long mismatch found.")

                st.markdown("<p class='section-label' style='margin-top:18px;'>✏️ Edit Data</p>", unsafe_allow_html=True)
                edited = st.data_editor(df, use_container_width=True, hide_index=True, num_rows="dynamic")
                combined = {
                    "Original Data": df,
                    "Mismatch Data": bad if not bad.empty else pd.DataFrame(columns=df.columns),
                    "Updated Data": edited,
                    "LatLong": st.session_state.latlon_df if st.session_state.latlon_df is not None else pd.DataFrame([{"Latitude": lat_value, "Longitude": lon_value}]),
                }
                st.download_button(
                    "Download Mismatch Excel",
                    df_to_excel_bytes(combined),
                    file_name="mismatch_with_latlong.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("Latitude/Longitude columns not detected.")
                st.download_button(
                    "Download Excel",
                    df_to_excel_bytes({"Mismatch Source": df}),
                    file_name="mismatch_source.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("No mismatch file found.")

    with tabs[4]:
        if st.session_state.latlon_df is not None:
            st.markdown(render_styled_dataframe(st.session_state.latlon_df), unsafe_allow_html=True)
            st.download_button(
                "Download LatLong Excel",
                df_to_excel_bytes({"LatLong": st.session_state.latlon_df}),
                file_name="lat_long.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("No Lat/Long saved yet.")


st.markdown("---")
st.markdown(
    f"""
<div style="text-align:center;color:#94a3b8;font-size:12px;padding:24px 0;">
    <div style="margin-bottom:10px;">📂 <strong>Output:</strong> {st.session_state.shared_path}</div>
    <div style="margin-bottom:10px;">⏰ <strong>Last Refreshed:</strong> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
    <div style="font-size:11px;color:#64748b;margin-top:18px;letter-spacing:0.3px;"><em>Weather Auto Regression Suite | Premium Production Dashboard</em></div>
</div>
""",
    unsafe_allow_html=True,
)
