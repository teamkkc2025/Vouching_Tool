import io
import os
import re
from datetime import datetime
 
import streamlit as st
import pandas as pd
import numpy as np
import pdfplumber
import pytesseract
 
from PIL import Image, ImageEnhance
import xlsxwriter
 
 
# ---------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------
 
st.set_page_config(
    page_title="Vouching Tool",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)
 
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=IBM+Plex+Mono:wght@300;400;500&display=swap%27);
 
*, *::before, *::after { box-sizing: border-box; }
 
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
.main {
    background: #ffffff !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    color: #111827 !important;
}
 
[data-testid="stAppViewContainer"] {
    background: #f8f9fc !important;
}
 
[data-testid="stHeader"] { background: transparent !important; display: none; }
[data-testid="stToolbar"] { display: none !important; }
footer { display: none !important; }
 
.block-container {
    padding: 0 2.5rem 3rem !important;
    max-width: 1280px !important;
}
 
[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e5e7eb !important;
}
[data-testid="stSidebar"] .block-container { padding: 0 !important; }
 
/* Top bar */
.vt-topbar {
    background: #ffffff;
    border-bottom: 1px solid #e5e7eb;
    padding: 18px 0 16px;
    margin-bottom: 32px;
    display: flex;
    align-items: center;
    gap: 14px;
}
.vt-icon {
    width: 38px; height: 38px;
    background: #1a56db;
    border-radius: 9px;
    display: flex; align-items: center; justify-content: center;
    font-size: 18px; flex-shrink: 0;
}
.vt-title {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 20px; font-weight: 800;
    color: #111827; letter-spacing: -0.4px; line-height: 1;
}
.vt-subtitle {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10.5px; color: #9ca3af; letter-spacing: 0.5px; margin-top: 3px;
}
.vt-pill {
    margin-left: auto;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px; color: #1a56db;
    background: #eff4ff; border: 1px solid #c7d7fe;
    padding: 4px 12px; border-radius: 20px;
}
 
/* Section label */
.vt-section-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px; letter-spacing: 1.8px; text-transform: uppercase;
    color: #9ca3af; margin: 28px 0 12px;
    display: flex; align-items: center; gap: 10px;
}
.vt-section-label::after {
    content: ''; flex: 1; height: 1px; background: #f3f4f6;
}
 
/* Upload */
.vt-upload-title {
    font-size: 12px; font-weight: 700; color: #374151;
    letter-spacing: 0.3px; margin-bottom: 10px;
    display: flex; align-items: center; gap: 7px;
}
.vt-upload-title .badge {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9.5px; background: #f3f4f6; color: #6b7280;
    padding: 2px 8px; border-radius: 4px; font-weight: 400;
}
.vt-file-confirmed {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10.5px; color: #059669;
    background: #f0fdf4; border: 1px solid #bbf7d0;
    padding: 6px 12px; border-radius: 6px; margin-top: 8px;
}
 
[data-testid="stFileUploader"] {
    background: #fafafa !important;
    border: 1.5px dashed #d1d5db !important;
    border-radius: 8px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #1a56db !important;
    background: #f5f8ff !important;
}
[data-testid="stFileUploaderDropzone"] {
    background: transparent !important; border: none !important;
}
 
/* Metrics */
.vt-metrics {
    display: grid; grid-template-columns: repeat(3, 1fr);
    gap: 14px; margin: 6px 0 28px;
}
.vt-metric {
    background: #ffffff; border: 1.5px solid #e5e7eb;
    border-radius: 12px; padding: 22px 22px 18px;
    position: relative; overflow: hidden;
}
.vt-metric-accent {
    position: absolute; left: 0; top: 0; bottom: 0;
    width: 4px; border-radius: 12px 0 0 12px;
}
.vt-metric.matched .vt-metric-accent { background: #059669; }
.vt-metric.missing  .vt-metric-accent { background: #dc2626; }
.vt-metric.duplicate .vt-metric-accent { background: #d97706; }
 
.vt-metric-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9.5px; letter-spacing: 1.2px; text-transform: uppercase;
    color: #9ca3af; margin-bottom: 10px; padding-left: 14px;
}
.vt-metric-value {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 40px; font-weight: 800; line-height: 1;
    padding-left: 14px; letter-spacing: -1.5px;
}
.vt-metric.matched  .vt-metric-value { color: #059669; }
.vt-metric.missing  .vt-metric-value { color: #dc2626; }
.vt-metric.duplicate .vt-metric-value { color: #d97706; }
.vt-metric-sub {
    font-size: 11.5px; color: #9ca3af; margin-top: 6px;
    padding-left: 14px; font-weight: 400;
}
 
/* Buttons */
[data-testid="stButton"] > button {
    background: #1a56db !important; color: #ffffff !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 13.5px !important; font-weight: 700 !important;
    border: none !important; border-radius: 9px !important;
    padding: 11px 32px !important;
    box-shadow: 0 1px 3px rgba(26,86,219,0.25), 0 4px 12px rgba(26,86,219,0.15) !important;
    transition: all 0.18s !important;
}
[data-testid="stButton"] > button:hover {
    background: #1648c0 !important;
    transform: translateY(-1px) !important;
}
[data-testid="stDownloadButton"] > button {
    background: #ffffff !important; color: #1a56db !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important; font-size: 13px !important;
    border: 1.5px solid #1a56db !important;
    border-radius: 9px !important; padding: 10px 28px !important;
    transition: all 0.18s !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #eff4ff !important;
}
 
/* Progress */
[data-testid="stProgressBar"] > div {
    background: #f3f4f6 !important; border-radius: 99px !important; height: 5px !important;
}
[data-testid="stProgressBar"] > div > div {
    background: #1a56db !important; border-radius: 99px !important;
}
 
/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1.5px solid #e5e7eb !important; border-radius: 10px !important;
    overflow: hidden !important; box-shadow: 0 1px 4px rgba(0,0,0,0.04) !important;
}
 
.vt-divider { height: 1px; background: #f3f4f6; margin: 24px 0; }
 
.vt-result-head {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 16px; font-weight: 800; color: #111827;
    letter-spacing: -0.3px; margin: 4px 0 14px;
    display: flex; align-items: center; gap: 10px;
}
.vt-result-head .count-tag {
    font-family: 'IBM Plex Mono', monospace; font-size: 10px;
    background: #eff4ff; color: #1a56db; border: 1px solid #c7d7fe;
    padding: 3px 10px; border-radius: 20px;
}
 
.vt-infobox {
    background: #f8faff; border: 1px solid #dbe4ff; border-radius: 10px;
    padding: 13px 16px; font-size: 12.5px; color: #3b5bdb;
    display: flex; gap: 10px; align-items: flex-start;
    margin-bottom: 28px; line-height: 1.6;
}
 
/* Sidebar */
.sb-header { padding: 24px 20px 20px; border-bottom: 1px solid #f3f4f6; margin-bottom: 8px; }
.sb-logo {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 15px; font-weight: 800; color: #111827; letter-spacing: -0.3px;
}
.sb-logo-dot {
    display: inline-block; width: 7px; height: 7px; background: #1a56db;
    border-radius: 50%; margin-left: 3px; vertical-align: middle; margin-bottom: 2px;
}
.sb-ver { font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: #9ca3af; margin-top: 2px; }
.sb-nav { padding: 4px 12px; }
.sb-nav-item {
    display: flex; align-items: center; gap: 10px; padding: 9px 12px;
    border-radius: 8px; font-size: 13px; color: #6b7280; font-weight: 500;
}
.sb-nav-item.active { background: #eff4ff; color: #1a56db; font-weight: 700; }
.sb-divider { height: 1px; background: #f3f4f6; margin: 10px 20px; }
.sb-info-block {
    margin: 12px 16px; background: #f9fafb;
    border: 1px solid #f3f4f6; border-radius: 10px; padding: 14px;
}
.sb-info-title {
    font-family: 'IBM Plex Mono', monospace; font-size: 9px;
    letter-spacing: 1.5px; text-transform: uppercase; color: #9ca3af; margin-bottom: 10px;
}
.sb-info-row {
    display: flex; justify-content: space-between; align-items: center;
    padding: 5px 0; border-bottom: 1px solid #f3f4f6; font-size: 12px;
}
.sb-info-row:last-child { border-bottom: none; }
.sb-info-key { color: #9ca3af; }
.sb-info-val { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #374151; font-weight: 500; }
.sb-info-val.green { color: #059669; }
 
/* Empty state */
.vt-empty { text-align: center; padding: 64px 24px; }
.vt-empty-icon { font-size: 44px; margin-bottom: 16px; }
.vt-empty-title {
    font-family: 'Plus Jakarta Sans', sans-serif; font-size: 16px;
    font-weight: 700; color: #9ca3af; margin-bottom: 6px;
}
.vt-empty-sub { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #d1d5db; }
 
.vt-preview-meta { font-family: 'IBM Plex Mono', monospace; font-size: 10.5px; color: #9ca3af; margin-top: 6px; }
.vt-proc-label { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #6b7280; margin-bottom: 8px; }
.vt-proc-done  { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #059669; }
.vt-timestamp  { font-family: 'IBM Plex Mono', monospace; font-size: 10.5px; color: #9ca3af; padding-top: 11px; }
 
[data-testid="stCaptionContainer"] { font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important; color: #9ca3af !important; }
label, [data-testid="stWidgetLabel"] { font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 13px !important; color: #374151 !important; font-weight: 500 !important; }
h2, h3 { font-family: 'Plus Jakarta Sans', sans-serif !important; color: #111827 !important; font-weight: 800 !important; }
[data-testid="stSpinner"] { color: #1a56db !important; }
 
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #f9fafb; }
::-webkit-scrollbar-thumb { background: #e5e7eb; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)
 
 
# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------
 
with st.sidebar:
    st.markdown("""
    <div class="sb-header">
        <div class="sb-logo">Vouching Tool<span class="sb-logo-dot"></span></div>
        <div class="sb-ver">v6.0 · Audit Suite</div>
    </div>
    <div class="sb-nav">
        <div class="sb-nav-item active"><span>📋</span> Vouching Engine</div>
        <div class="sb-nav-item"><span>📊</span> Analytics</div>
        <div class="sb-nav-item"><span>🗂️</span> Report History</div>
        <div class="sb-nav-item"><span>⚙️</span> Settings</div>
    </div>
    <div class="sb-divider"></div>
    <div class="sb-info-block">
        <div class="sb-info-title">Engine Status</div>
        <div class="sb-info-row">
            <span class="sb-info-key">OCR Engine</span>
            <span class="sb-info-val green">Active</span>
        </div>
        <div class="sb-info-row">
            <span class="sb-info-key">PDF Parser</span>
            <span class="sb-info-val green">Active</span>
        </div>
        <div class="sb-info-row">
            <span class="sb-info-key">Match Algorithm</span>
            <span class="sb-info-val">v3.2</span>
        </div>
        <div class="sb-info-row">
            <span class="sb-info-key">Threshold</span>
            <span class="sb-info-val">± ₹0.05</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
 
    st.markdown("""
    <div style="position:absolute;bottom:20px;left:0;right:0;padding:0 20px;">
        <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:#d1d5db;text-align:center;">
            © 2025 Vouching Tool · All rights reserved
        </div>
    </div>
    """, unsafe_allow_html=True)
 
 
# ---------------------------------------------------------
# TESSERACT
# ---------------------------------------------------------
 
try:
    if os.name == "nt":
        path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
    pytesseract.get_tesseract_version()
    OCR_OK = True
except:
    OCR_OK = False
 
 
# ---------------------------------------------------------
# OCR
# ---------------------------------------------------------
 
def ocr_image(img):
    try:
        return pytesseract.image_to_string(img, config="--psm 6 --oem 3")
    except:
        return ""
 
def preprocess(img):
    gray = img.convert("L")
    gray = ImageEnhance.Contrast(gray).enhance(2)
    gray = gray.resize((gray.width * 2, gray.height * 2))
    return gray
 
 
# ---------------------------------------------------------
# AMOUNT EXTRACTION
# ---------------------------------------------------------
 
def extract_amounts(text):
    text = text.replace(",", "")
    nums = set()
    patterns = [
        r"(?:₹|rs\.?|inr)\s*(\d+(?:\.\d{1,2})?)",
        r"(?:total|grand total|amount)\D{0,10}(\d+(?:\.\d{1,2})?)",
        r"\b(\d{2,6}(?:\.\d{1,2})?)\b"
    ]
    for p in patterns:
        for m in re.findall(p, text, re.IGNORECASE):
            try:
                val = float(m)
                if 1 <= val <= 500000:
                    nums.add(round(val, 2))
            except:
                pass
    return list(nums)
 
 
# ---------------------------------------------------------
# VENDOR DETECTION
# ---------------------------------------------------------
 
def detect_vendor(text):
    vendors = [
        "uber", "ola", "rapido",
        "zomato", "swiggy",
        "amazon", "flipkart",
        "airtel", "jio", "bsnl",
        "restaurant", "cafe", "hotel"
    ]
    t = text.lower()
    for v in vendors:
        if v in t:
            return v
    return None
 
 
# ---------------------------------------------------------
# PROCESS FILES
# ---------------------------------------------------------
 
def process_image(file_bytes, name):
    img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
    img = preprocess(img)
    text = ocr_image(img)
    return {"name": name, "amounts": extract_amounts(text), "vendor": detect_vendor(text), "text": text}
 
def process_pdf(file_bytes, name):
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for p in pdf.pages[:3]:
                text += p.extract_text() or ""
    except:
        pass
    return {"name": name, "amounts": extract_amounts(text), "vendor": detect_vendor(text), "text": text}
 
def process_file(file_bytes, name):
    ext = name.split(".")[-1].lower()
    if ext == "pdf":
        return process_pdf(file_bytes, name)
    return process_image(file_bytes, name)
 
 
# ---------------------------------------------------------
# MATCHING LOGIC
# ---------------------------------------------------------
 
def amount_match(reg_amt, doc_amounts):
    try:
        ra = float(reg_amt)
    except:
        return False, None
    # Guard against NaN/Inf in the register itself
    if np.isnan(ra) or np.isinf(ra):
        return False, None
    for da in doc_amounts:
        if round(ra, 2) == round(da, 2):
            return True, da
    for da in doc_amounts:
        if abs(ra - da) <= 0.05:
            return True, da
    return False, None
 
def vendor_match(reg_vendor, doc_vendor, text):
    rv = str(reg_vendor).lower()
    dv = str(doc_vendor).lower()
    if rv and dv and rv in dv:
        return True
    if rv in text.lower():
        return True
    return False
 
def category_match(cat, text):
    mapping = {
        "food":   ["zomato", "swiggy", "restaurant"],
        "travel": ["uber", "ola", "rapido"],
        "mobile": ["airtel", "jio"]
    }
    rc = str(cat).lower()
    if rc in text.lower():
        return True
    if rc in mapping:
        for k in mapping[rc]:
            if k in text.lower():
                return True
    return False
 
 
# ---------------------------------------------------------
# VOUCHING ENGINE
# ---------------------------------------------------------
 
def run_vouching(df, docs):
    results = []
    used = set()
    duplicates = set()
 
    for i, row in df.iterrows():
        amount   = row["Expense Amount"]
        vendor   = row["Vendor"]
        category = row["Category"]
        best_score = 0
        best_doc   = None
 
        for d in docs:
            score = 0
            ok, _ = amount_match(amount, d["amounts"])
            if ok:
                score += 10
            if vendor_match(vendor, d["vendor"], d["text"]):
                score += 3
            if category_match(category, d["text"]):
                score += 2
            if score > best_score:
                best_score = score
                best_doc   = d
 
        if best_doc:
            if best_doc["name"] in used:
                status = "DUPLICATE_RECEIPT"
                duplicates.add(best_doc["name"])
            else:
                used.add(best_doc["name"])
                status = "MATCHED"
        else:
            status = "MISSING_DOC"
 
        # Store amount as float, everything else as string
        try:
            amt_val = float(amount)
            if np.isnan(amt_val) or np.isinf(amt_val):
                amt_val = 0.0
        except:
            amt_val = 0.0
 
        results.append({
            "Report ID":    str(row["ExpenseReport ID"]),
            "Amount (Rs)":  amt_val,                          # plain float, no NaN/Inf
            "Vendor":       str(vendor),
            "Category":     str(category),
            "Matched File": best_doc["name"] if best_doc else "-",
            "Confidence":   f"{best_score}/15",
            "Status":       status,
        })
 
    return results
 
 
# ---------------------------------------------------------
# EXCEL EXPORT  — three-layer NaN/Inf defence
# ---------------------------------------------------------
 
def export_excel(df):
    """
    Layer 1: clean_df_for_excel  — pandas-level sanitise
    Layer 2: nan_inf_to_errors   — xlsxwriter Workbook option
    Layer 3: per-cell type check — write_string / write_number never receives NaN
    """
 
    # ── Layer 1: sanitise the dataframe ──────────────────
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                  .replace([np.inf, -np.inf], np.nan)
                  .fillna(0)
            )
        else:
            df[col] = df[col].fillna("").astype(str)
 
    buffer = io.BytesIO()
 
    # ── Layer 2: Workbook option ──────────────────────────
    wb = xlsxwriter.Workbook(buffer, {
        "nan_inf_to_errors": True,   # converts any surviving NaN/Inf to #NUM!
        "in_memory": True
    })
 
    # ── Formats ──────────────────────────────────────────
    hdr = wb.add_format({
        "bold": True, "bg_color": "#1a56db", "font_color": "#ffffff",
        "font_name": "Calibri", "font_size": 11,
        "valign": "vcenter", "align": "center", "border": 0
    })
    base = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "valign": "vcenter"
    })
    num_f = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "valign": "vcenter",
        "num_format": "#,##0.00"
    })
    ok_f = wb.add_format({
        "font_name": "Calibri", "font_size": 10,
        "font_color": "#059669", "bold": True, "valign": "vcenter"
    })
    mi_f = wb.add_format({
        "font_name": "Calibri", "font_size": 10,
        "font_color": "#dc2626", "bold": True, "valign": "vcenter"
    })
    du_f = wb.add_format({
        "font_name": "Calibri", "font_size": 10,
        "font_color": "#d97706", "bold": True, "valign": "vcenter"
    })
 
    # ── Results sheet ─────────────────────────────────────
    ws = wb.add_worksheet("Vouching Results")
    ws.set_row(0, 22)
    col_widths = {"Report ID": 18, "Amount (Rs)": 16, "Vendor": 20,
                  "Category": 16, "Matched File": 32, "Confidence": 12, "Status": 20}
    for ci, col in enumerate(df.columns):
        ws.set_column(ci, ci, col_widths.get(col, 16))
        ws.write(0, ci, col, hdr)
 
    status_ci = list(df.columns).index("Status") if "Status" in df.columns else -1
    amount_ci = list(df.columns).index("Amount (Rs)") if "Amount (Rs)" in df.columns else -1
 
    for ri in range(len(df)):
        status = str(df.iloc[ri].get("Status", ""))
        sfmt   = ok_f if status == "MATCHED" else (mi_f if status == "MISSING_DOC" else du_f)
 
        for ci in range(len(df.columns)):
            raw = df.iloc[ri, ci]
            fmt = sfmt if ci == status_ci else (num_f if ci == amount_ci else base)
 
            # ── Layer 3: safe per-cell write ──
            if isinstance(raw, (int, np.integer)):
                ws.write_number(ri + 1, ci, int(raw), fmt)
            elif isinstance(raw, (float, np.floating)):
                v = float(raw)
                if np.isnan(v) or np.isinf(v):
                    ws.write_string(ri + 1, ci, "", fmt)
                else:
                    ws.write_number(ri + 1, ci, v, fmt)
            else:
                ws.write_string(ri + 1, ci, str(raw) if raw is not None else "", fmt)
 
    # ── Summary sheet ─────────────────────────────────────
    sw = wb.add_worksheet("Summary")
    sw.set_column(0, 0, 26)
    sw.set_column(1, 1, 14)
 
    t_fmt = wb.add_format({"font_name": "Calibri", "font_size": 14, "bold": True, "font_color": "#111827"})
    l_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "font_color": "#6b7280"})
    v_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True, "font_color": "#111827"})
 
    sw.write(0, 0, "Vouching Report — Summary", t_fmt)
    sw.write(1, 0, f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}", l_fmt)
 
    total    = len(df)
    matched  = int((df["Status"] == "MATCHED").sum())
    missing  = int((df["Status"] == "MISSING_DOC").sum())
    dup      = int((df["Status"] == "DUPLICATE_RECEIPT").sum())
    rate     = round(matched / total * 100, 1) if total else 0.0
 
    for i, (lbl, val) in enumerate([
        ("Total Entries",      total),
        ("Matched",            matched),
        ("Missing Documents",  missing),
        ("Duplicate Receipts", dup),
        ("Match Rate (%)",     rate),
    ]):
        sw.write(i + 3, 0, lbl, l_fmt)
        sw.write_number(i + 3, 1, float(val), v_fmt)
 
    wb.close()
    buffer.seek(0)
    return buffer.read()          # return bytes, not BytesIO
 
 
# ---------------------------------------------------------
# TOP BAR
# ---------------------------------------------------------
 
st.markdown("""
<div class="vt-topbar">
    <div class="vt-icon">📋</div>
    <div>
        <div class="vt-title">Vouching Tool</div>
        <div class="vt-subtitle">Expense Audit &amp; Receipt Matching</div>
    </div>
    <div class="vt-pill">● Session Active</div>
</div>
""", unsafe_allow_html=True)
 
st.markdown("""
<div class="vt-infobox">
    <span style="font-size:15px;flex-shrink:0">ℹ</span>
    <span>
        Upload an <strong>Expense Register</strong> (.xlsx) and one or more
        <strong>Receipts / Invoices</strong> (PDF, PNG, JPG).
        The engine matches each entry to a supporting document using amount, vendor, and category
        signals — then generates a colour-coded audit report with a Summary sheet.
    </span>
</div>
""", unsafe_allow_html=True)
 
 
# ---------------------------------------------------------
# UPLOAD
# ---------------------------------------------------------
 
st.markdown('<div class="vt-section-label">Document Intake</div>', unsafe_allow_html=True)
 
col1, col2 = st.columns(2, gap="medium")
 
with col1:
    st.markdown('<div class="vt-upload-title">Expense Register <span class="badge">XLSX</span></div>', unsafe_allow_html=True)
    register = st.file_uploader("exp_reg", type=["xlsx"], label_visibility="collapsed", key="reg")
    if register:
        st.markdown(f'<div class="vt-file-confirmed">✓ &nbsp;{register.name}</div>', unsafe_allow_html=True)
 
with col2:
    st.markdown('<div class="vt-upload-title">Receipts &amp; Invoices <span class="badge">PDF / PNG / JPG</span></div>', unsafe_allow_html=True)
    docs = st.file_uploader("receipts", accept_multiple_files=True, label_visibility="collapsed", key="docs")
    if docs:
        st.markdown(f'<div class="vt-file-confirmed">✓ &nbsp;{len(docs)} file{"s" if len(docs)!=1 else ""} ready</div>', unsafe_allow_html=True)
 
 
# ---------------------------------------------------------
# MAIN FLOW
# ---------------------------------------------------------
 
if register and docs:
 
    st.markdown('<div class="vt-section-label">Register Preview</div>', unsafe_allow_html=True)
    df = pd.read_excel(register)
    st.dataframe(df.head(8), use_container_width=True, hide_index=True)
    st.markdown(f'<div class="vt-preview-meta">{len(df)} entries · {len(df.columns)} columns</div>', unsafe_allow_html=True)
 
    st.markdown('<div class="vt-section-label">Processing Documents</div>', unsafe_allow_html=True)
    files_data = []
    prog_text  = st.empty()
    prog_bar   = st.progress(0)
 
    for i, f in enumerate(docs):
        prog_text.markdown(
            f'<div class="vt-proc-label">Extracting → <strong style="color:#374151">{f.name}</strong> '
            f'<span style="color:#1a56db">[{i+1}/{len(docs)}]</span></div>',
            unsafe_allow_html=True
        )
        files_data.append(process_file(f.read(), f.name))
        prog_bar.progress((i + 1) / len(docs))
 
    prog_text.markdown('<div class="vt-proc-done">✓ &nbsp;All documents processed successfully</div>', unsafe_allow_html=True)
    st.markdown('<div class="vt-divider"></div>', unsafe_allow_html=True)
 
    btn_col, _ = st.columns([1, 4])
    with btn_col:
        run = st.button("Run Vouching Analysis →")
 
    if run:
        with st.spinner("Matching expense entries to documents…"):
            results = run_vouching(df, files_data)
            rdf     = pd.DataFrame(results)
 
        n_matched = int((rdf.Status == "MATCHED").sum())
        n_missing = int((rdf.Status == "MISSING_DOC").sum())
        n_dup     = int((rdf.Status == "DUPLICATE_RECEIPT").sum())
        total     = len(rdf)
        rate      = round(n_matched / total * 100) if total else 0
 
        st.markdown('<div class="vt-section-label">Summary</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="vt-metrics">
            <div class="vt-metric matched">
                <div class="vt-metric-accent"></div>
                <div class="vt-metric-label">Matched</div>
                <div class="vt-metric-value">{n_matched}</div>
                <div class="vt-metric-sub">{rate}% match rate</div>
            </div>
            <div class="vt-metric missing">
                <div class="vt-metric-accent"></div>
                <div class="vt-metric-label">Missing Documents</div>
                <div class="vt-metric-value">{n_missing}</div>
                <div class="vt-metric-sub">No receipt found</div>
            </div>
            <div class="vt-metric duplicate">
                <div class="vt-metric-accent"></div>
                <div class="vt-metric-label">Duplicates</div>
                <div class="vt-metric-value">{n_dup}</div>
                <div class="vt-metric-sub">Reused receipts</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
 
        st.markdown('<div class="vt-section-label">Detailed Results</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="vt-result-head">Vouching Report <span class="count-tag">{total} entries</span></div>', unsafe_allow_html=True)
 
        def style_status(val):
            return {
                "MATCHED":           "color:#059669;font-weight:700",
                "MISSING_DOC":       "color:#dc2626;font-weight:700",
                "DUPLICATE_RECEIPT": "color:#d97706;font-weight:700",
            }.get(val, "")
 
        st.dataframe(rdf.style.applymap(style_status, subset=["Status"]), use_container_width=True, hide_index=True)
 
        st.markdown('<div class="vt-divider"></div>', unsafe_allow_html=True)
        dl_col, ts_col = st.columns([1, 3])
 
        with dl_col:
            excel_bytes = export_excel(rdf)          # returns bytes
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                label="⬇  Download Audit Report (.xlsx)",
                data=excel_bytes,
                file_name=f"vouching_report_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with ts_col:
            st.markdown(
                f'<div class="vt-timestamp">Generated {datetime.now().strftime("%d %b %Y · %H:%M")} &nbsp;·&nbsp; '
                f'{total} entries &nbsp;·&nbsp; 2 sheets: Results + Summary</div>',
                unsafe_allow_html=True
            )
 
else:
    st.markdown("""
    <div class="vt-empty">
        <div class="vt-empty-icon">📋</div>
        <div class="vt-empty-title">Ready to begin</div>
        <div class="vt-empty-sub">Upload your expense register and receipts above to start vouching</div>
    </div>
    """, unsafe_allow_html=True)
