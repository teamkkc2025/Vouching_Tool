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
# TESSERACT
# ---------------------------------------------------------

import toml

try:
    _cfg_path = os.path.join(os.path.dirname(__file__), "config.toml")
    if os.path.exists(_cfg_path):
        _cfg = toml.load(_cfg_path)
        _tess_path = _cfg.get("tesseract", {}).get("path", "")
        if _tess_path and os.path.exists(_tess_path):
            pytesseract.pytesseract.tesseract_cmd = _tess_path
    elif os.name == "nt":
        path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
    pytesseract.get_tesseract_version()
    OCR_OK = True
except:
    OCR_OK = False
 
 
# ---------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------
 
st.set_page_config(
    page_title="KKC Vouching Tool",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)
 
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700;800;900&family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&display=swap%27);
 
:root {
    --kkc-green:     #5ba632;
    --kkc-green-mid: #6dbc3c;
    --kkc-green-lt:  #e8f5e0;
    --kkc-green-bar: #76b82a;
    --white:         #ffffff;
    --gray-50:       #f7f9f7;
    --gray-100:      #eef2ee;
    --gray-200:      #dde5dd;
    --gray-400:      #9aaa9a;
    --gray-600:      #5a6b5a;
    --gray-800:      #2a3a2a;
    --text:          #1c2b1c;
    --text-mid:      #445544;
    --text-light:    #7a8e7a;
}
 
*, *::before, *::after { box-sizing: border-box; }
 
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
.main {
    background: var(--gray-50) !important;
    font-family: 'Nunito Sans', sans-serif !important;
    color: var(--text) !important;
}
 
[data-testid="stHeader"] { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
footer { display: none !important; }
 
.block-container {
    padding: 0 2.5rem 3rem !important;
    max-width: 1320px !important;
}
 
/* ── SIDEBAR ─────────────────────────────── */
[data-testid="stSidebar"] {
    background: var(--white) !important;
    border-right: 2px solid var(--kkc-green-lt) !important;
}
[data-testid="stSidebar"] .block-container { padding: 0 !important; }
 
.sb-topbar {
    background: var(--kkc-green-bar);
    padding: 0 0 0 0;
    height: 5px;
}
.sb-brand {
    padding: 22px 20px 18px;
    border-bottom: 1px solid var(--gray-100);
}
.sb-logo-text {
    font-family: 'Nunito Sans', sans-serif;
    font-size: 17px; font-weight: 900; letter-spacing: -0.3px;
    color: var(--kkc-green);
    line-height: 1;
}
.sb-logo-amp { color: var(--kkc-green-mid); }
.sb-logo-sub {
    font-size: 10px; font-weight: 400; color: var(--text-light);
    margin-top: 3px; letter-spacing: 0.2px;
}
.sb-logo-sub2 {
    font-size: 9.5px; color: var(--gray-400); margin-top: 1px;
    font-style: italic;
}
 
.sb-nav { padding: 10px 12px; }
.sb-nav-item {
    display: flex; align-items: center; gap: 10px;
    padding: 10px 14px; border-radius: 7px;
    font-size: 13px; color: var(--text-mid); font-weight: 600;
    cursor: pointer; transition: all 0.15s; margin-bottom: 2px;
}
.sb-nav-item.active {
    background: var(--kkc-green-lt);
    color: var(--kkc-green);
    border-left: 3px solid var(--kkc-green);
    padding-left: 11px;
}
.sb-nav-item:not(.active):hover { background: var(--gray-50); }
 
.sb-section-title {
    font-size: 9px; font-weight: 700; letter-spacing: 1.8px;
    text-transform: uppercase; color: var(--gray-400);
    padding: 16px 14px 6px;
}
.sb-divider { height: 1px; background: var(--gray-100); margin: 8px 14px; }
 
.sb-status-block {
    margin: 10px 14px;
    background: var(--gray-50);
    border: 1px solid var(--gray-100);
    border-radius: 10px; padding: 14px 16px;
}
.sb-status-title {
    font-size: 9px; font-weight: 700; letter-spacing: 1.5px;
    text-transform: uppercase; color: var(--gray-400); margin-bottom: 10px;
}
.sb-status-row {
    display: flex; justify-content: space-between; align-items: center;
    padding: 5px 0; border-bottom: 1px solid var(--gray-100); font-size: 11.5px;
}
.sb-status-row:last-child { border-bottom: none; }
.sb-status-key { color: var(--text-light); font-weight: 400; }
.sb-status-val { font-weight: 700; font-size: 11px; color: var(--text-mid); }
.sb-status-val.ok { color: var(--kkc-green); }
 
/* ── TOP HEADER BAR ────────────────────── */
.vt-header {
    background: var(--white);
    border-bottom: 3px solid var(--kkc-green-bar);
    padding: 0;
    margin-bottom: 0;
}
.vt-topstrip {
    background: var(--kkc-green-bar);
    padding: 7px 0;
    display: flex; align-items: center; justify-content: center; gap: 30px;
    font-size: 11px; color: var(--white); font-weight: 600; letter-spacing: 0.2px;
}
.vt-topstrip a { color: var(--white); text-decoration: none; }
.vt-topstrip-sep { opacity: 0.4; }
 
.vt-navbar {
    background: var(--white);
    padding: 14px 0 12px;
    display: flex; align-items: center; gap: 20px;
}
.vt-brand {
    display: flex; flex-direction: column; line-height: 1;
}
.vt-brand-name {
    font-family: 'Nunito Sans', sans-serif;
    font-size: 21px; font-weight: 900; color: var(--kkc-green);
    letter-spacing: -0.5px;
}
.vt-brand-tag {
    font-size: 10.5px; color: var(--text-light); font-weight: 400; margin-top: 2px;
}
.vt-brand-former {
    font-size: 9.5px; color: var(--gray-400); font-style: italic; margin-top: 1px;
}
.vt-nav-divider {
    width: 1px; height: 36px; background: var(--gray-200); flex-shrink: 0;
}
.vt-tool-badge {
    background: var(--kkc-green-lt);
    border: 1.5px solid var(--kkc-green);
    border-radius: 7px; padding: 6px 16px;
    font-size: 13px; font-weight: 800; color: var(--kkc-green);
    letter-spacing: -0.2px;
}
.vt-tool-sub {
    font-size: 10px; color: var(--text-light); font-weight: 400; margin-top: 2px;
    font-family: 'Nunito Sans', sans-serif;
}
.vt-status-pill {
    margin-left: auto;
    background: var(--kkc-green-lt);
    border: 1px solid var(--kkc-green);
    border-radius: 20px; padding: 5px 16px;
    font-size: 10.5px; font-weight: 700; color: var(--kkc-green);
    display: flex; align-items: center; gap: 6px;
}
.vt-status-dot {
    width: 7px; height: 7px; border-radius: 50%;
    background: var(--kkc-green);
    animation: blink 1.8s ease-in-out infinite;
}
@keyframes blink { 0%,100%{opacity:1} 50%{opacity:0.3} }
 
/* ── SECTION LABELS ──────────────────── */
.vt-section {
    font-size: 9.5px; font-weight: 700; letter-spacing: 2px;
    text-transform: uppercase; color: var(--kkc-green);
    margin: 28px 0 14px;
    display: flex; align-items: center; gap: 12px;
}
.vt-section::after {
    content: ''; flex: 1; height: 1.5px; background: var(--kkc-green-lt);
}
 
/* ── INFO BOX ────────────────────────── */
.vt-infobox {
    background: var(--kkc-green-lt);
    border: 1px solid var(--kkc-green);
    border-left: 4px solid var(--kkc-green);
    border-radius: 0 9px 9px 0;
    padding: 14px 18px; font-size: 12.5px; color: var(--gray-800);
    display: flex; gap: 12px; align-items: flex-start;
    margin-bottom: 24px; line-height: 1.65;
}
 
/* ── UPLOAD ──────────────────────────── */
.vt-upload-label {
    font-size: 12px; font-weight: 700; color: var(--text-mid);
    margin-bottom: 8px; display: flex; align-items: center; gap: 8px;
}
.vt-upload-label .badge {
    font-size: 9px; background: var(--gray-100); color: var(--text-light);
    padding: 2px 8px; border-radius: 4px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.5px;
}
.vt-file-ok {
    font-size: 10.5px; font-weight: 700; color: var(--kkc-green);
    background: var(--kkc-green-lt); border: 1px solid var(--kkc-green);
    padding: 7px 14px; border-radius: 7px; margin-top: 8px;
    display: flex; align-items: center; gap: 6px;
}
 
[data-testid="stFileUploader"] {
    background: var(--white) !important;
    border: 2px dashed var(--gray-200) !important;
    border-radius: 9px !important;
    transition: all 0.2s !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--kkc-green) !important;
    background: var(--kkc-green-lt) !important;
}
[data-testid="stFileUploaderDropzone"] {
    background: transparent !important; border: none !important;
}
 
/* ── METRICS ─────────────────────────── */
.vt-metrics {
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 14px; margin: 6px 0 28px;
}
.vt-metric {
    background: var(--white);
    border: 1px solid var(--gray-200);
    border-top: 4px solid transparent;
    border-radius: 10px; padding: 22px 20px 18px;
    position: relative;
    transition: box-shadow 0.2s, transform 0.2s;
}
.vt-metric:hover {
    box-shadow: 0 6px 24px rgba(91,166,50,0.1);
    transform: translateY(-2px);
}
.vt-metric.matched  { border-top-color: var(--kkc-green); }
.vt-metric.mismatch { border-top-color: #e07b00; }
.vt-metric.missing  { border-top-color: #d13030; }
.vt-metric.duplicate{ border-top-color: #b8860b; }
 
.vt-metric-label {
    font-size: 9.5px; font-weight: 700; letter-spacing: 1.5px;
    text-transform: uppercase; color: var(--text-light); margin-bottom: 10px;
}
.vt-metric-value {
    font-family: 'Libre Baskerville', serif;
    font-size: 48px; font-weight: 700; line-height: 1;
    letter-spacing: -2px;
}
.vt-metric.matched  .vt-metric-value { color: var(--kkc-green); }
.vt-metric.mismatch .vt-metric-value { color: #e07b00; }
.vt-metric.missing  .vt-metric-value { color: #d13030; }
.vt-metric.duplicate .vt-metric-value { color: #b8860b; }
.vt-metric-sub { font-size: 11px; color: var(--text-light); margin-top: 7px; font-weight: 400; }
 
/* ── MISMATCH ALERT ─────────────────── */
.vt-alert {
    border-radius: 9px; padding: 16px 20px;
    font-size: 12.5px; display: flex; gap: 14px;
    align-items: flex-start; margin-bottom: 24px; line-height: 1.6;
}
.vt-alert.warn {
    background: #fff8f0;
    border: 1px solid #f5c57a;
    border-left: 4px solid #e07b00;
    color: #6b3a00;
}
.vt-alert.info {
    background: var(--kkc-green-lt);
    border: 1px solid var(--kkc-green);
    border-left: 4px solid var(--kkc-green);
    color: var(--gray-800);
}
 
.mismatch-table {
    width: 100%; margin-top: 12px; border-collapse: collapse;
    font-size: 11.5px;
}
.mismatch-table th {
    text-align: left; font-size: 9px; font-weight: 700;
    letter-spacing: 1.3px; text-transform: uppercase;
    color: #a06020; padding: 4px 10px 8px 0;
}
.mismatch-table td {
    padding: 6px 10px 6px 0;
    border-top: 1px solid #fde8c8;
}
.mismatch-table tr:first-child td { border-top: none; }
.val-reg  { color: #d13030; font-weight: 700; }
.val-doc  { color: var(--kkc-green); font-weight: 700; }
.val-diff { color: #e07b00; font-weight: 700; }
 
/* ── BUTTONS ─────────────────────────── */
[data-testid="stButton"] > button {
    background: var(--kkc-green) !important;
    color: var(--white) !important;
    font-family: 'Nunito Sans', sans-serif !important;
    font-size: 13.5px !important; font-weight: 800 !important;
    border: none !important; border-radius: 8px !important;
    padding: 12px 34px !important;
    box-shadow: 0 2px 8px rgba(91,166,50,0.3) !important;
    transition: all 0.18s !important;
    letter-spacing: 0.2px !important;
}
[data-testid="stButton"] > button:hover {
    background: #4e9429 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 14px rgba(91,166,50,0.4) !important;
}
[data-testid="stDownloadButton"] > button {
    background: var(--white) !important;
    color: var(--kkc-green) !important;
    font-family: 'Nunito Sans', sans-serif !important;
    font-weight: 800 !important; font-size: 13px !important;
    border: 2px solid var(--kkc-green) !important;
    border-radius: 8px !important; padding: 11px 28px !important;
    transition: all 0.18s !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: var(--kkc-green-lt) !important;
}
 
/* ── PROGRESS ─────────────────────────── */
[data-testid="stProgressBar"] > div {
    background: var(--gray-100) !important;
    border-radius: 99px !important; height: 5px !important;
}
[data-testid="stProgressBar"] > div > div {
    background: var(--kkc-green) !important;
    border-radius: 99px !important;
}
 
/* ── TABLE ───────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--gray-200) !important;
    border-radius: 10px !important; overflow: hidden !important;
    box-shadow: 0 1px 6px rgba(0,0,0,0.04) !important;
}
 
/* ── MISC ─────────────────────────────── */
.vt-divider { height: 1px; background: var(--gray-100); margin: 24px 0; }
.vt-result-head {
    font-size: 16px; font-weight: 800; color: var(--text);
    letter-spacing: -0.3px; margin: 4px 0 16px;
    display: flex; align-items: center; gap: 12px;
}
.vt-result-head .count-tag {
    font-size: 10px; font-weight: 700;
    background: var(--kkc-green-lt); color: var(--kkc-green);
    border: 1px solid var(--kkc-green);
    padding: 3px 12px; border-radius: 20px;
}
.vt-preview-meta { font-size: 10.5px; color: var(--text-light); margin-top: 6px; }
.vt-proc-label   { font-size: 11px; color: var(--text-light); margin-bottom: 8px; }
.vt-proc-done    { font-size: 11px; color: var(--kkc-green); font-weight: 700; }
.vt-timestamp    { font-size: 10.5px; color: var(--text-light); padding-top: 11px; }
 
.vt-empty { text-align: center; padding: 72px 24px; }
.vt-empty-icon { font-size: 52px; margin-bottom: 16px; }
.vt-empty-title {
    font-size: 17px; font-weight: 800; color: var(--gray-400); margin-bottom: 8px;
}
.vt-empty-sub { font-size: 11.5px; color: var(--gray-200); }
 
/* ── FOOTER ─────────────────────────── */
.vt-footer {
    background: var(--white);
    border-top: 1px solid var(--gray-100);
    padding: 14px 0;
    text-align: center;
    font-size: 11px; color: var(--text-light);
    margin-top: 48px;
}
.vt-footer strong { color: var(--kkc-green); }
 
label, [data-testid="stWidgetLabel"] {
    font-family: 'Nunito Sans', sans-serif !important;
    font-size: 13px !important; color: var(--text-mid) !important; font-weight: 700 !important;
}
[data-testid="stCaptionContainer"] { font-size: 11px !important; color: var(--text-light) !important; }
h2, h3 { font-family: 'Nunito Sans', sans-serif !important; color: var(--text) !important; font-weight: 800 !important; }
[data-testid="stSpinner"] { color: var(--kkc-green) !important; }
 
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--gray-50); }
::-webkit-scrollbar-thumb { background: var(--gray-200); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--kkc-green); }
</style>
""", unsafe_allow_html=True)
 
 
# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------
 
with st.sidebar:
    st.markdown("""
    <div class="sb-topbar"></div>
    <div class="sb-brand">
        <div class="sb-logo-text">kkc <span class="sb-logo-amp">&</span> associates llp</div>
        <div class="sb-logo-sub">Chartered Accountants</div>
        <div class="sb-logo-sub2">(Formerly Khimji Kunverji & Co LLP)</div>
    </div>
 
    <div class="sb-section-title">Navigation</div>
    <div class="sb-nav">
        <div class="sb-nav-item active">🔍&nbsp; Vouching Engine</div>
        <div class="sb-nav-item">📊&nbsp; Analytics</div>
        <div class="sb-nav-item">🗂️&nbsp; Report History</div>
        <div class="sb-nav-item">⚙️&nbsp; Settings</div>
    </div>
 
    <div class="sb-divider"></div>
    <div class="sb-section-title">Engine Status</div>
    <div class="sb-status-block">
        <div class="sb-status-row">
            <span class="sb-status-key">OCR Engine</span>
            <span class="sb-status-val ok">● Active</span>
        </div>
        <div class="sb-status-row">
            <span class="sb-status-key">PDF Parser</span>
            <span class="sb-status-val ok">● Active</span>
        </div>
        <div class="sb-status-row">
            <span class="sb-status-key">Match Algorithm</span>
            <span class="sb-status-val">v4.0</span>
        </div>
        <div class="sb-status-row">
            <span class="sb-status-key">Amount Tolerance</span>
            <span class="sb-status-val">± ₹0.50</span>
        </div>
        <div class="sb-status-row">
            <span class="sb-status-key">ID Matching</span>
            <span class="sb-status-val ok">Enabled</span>
        </div>
    </div>
 
    <div style="position:absolute;bottom:16px;left:0;right:0;text-align:center;">
        <div style="font-size:10px;color:#ccc;">
            © 2025 KKC & Associates LLP
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
# AMOUNT EXTRACTION (prioritised)
# ---------------------------------------------------------
 
def extract_amounts_with_context(text):
    raw = text.replace(",", "")
    results = []
 
    p1_patterns = [
        r"(?:grand\s+total|total\s+amount|invoice\s+total|amount\s+due|net\s+payable)"
        r"[\s:₹Rs\.]*(\d+(?:\.\d{1,2})?)",
        r"(?:total)[^\n₹\d]{0,8}(?:₹|rs\.?|inr)?\s*(\d{4,}(?:\.\d{1,2})?)",
    ]
    for p in p1_patterns:
        for m in re.findall(p, raw, re.IGNORECASE):
            try:
                v = float(m)
                if 1 <= v <= 10_000_000:
                    results.append({"amount": round(v, 2), "label": "Grand Total", "priority": 1})
            except:
                pass
 
    p2_patterns = [
        r"(?:sub\s*total|taxable|basic|consultancy|professional\s+fees?)"
        r"[\s:₹Rs\.]*(\d+(?:\.\d{1,2})?)",
        r"(?:₹|rs\.?|inr)\s*(\d+(?:\.\d{1,2})?)",
    ]
    for p in p2_patterns:
        for m in re.findall(p, raw, re.IGNORECASE):
            try:
                v = float(m)
                if 1 <= v <= 10_000_000:
                    results.append({"amount": round(v, 2), "label": "Line Item", "priority": 2})
            except:
                pass
 
    for m in re.findall(r"\b(\d{4,7}(?:\.\d{1,2})?)\b", raw):
        try:
            v = float(m)
            if 1 <= v <= 10_000_000:
                results.append({"amount": round(v, 2), "label": "Number", "priority": 3})
        except:
            pass
 
    seen = {}
    for r in results:
        k = r["amount"]
        if k not in seen or r["priority"] < seen[k]["priority"]:
            seen[k] = r
    return list(seen.values())
 
def extract_amounts(text):
    return [r["amount"] for r in extract_amounts_with_context(text)]
 
 
# ---------------------------------------------------------
# INVOICE ID EXTRACTION
# ---------------------------------------------------------
 
def extract_invoice_ids(text):
    ids = set()
    patterns = [
        r"invoice\s*(?:no\.?|#|number)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9/\-]{3,30})",
        r"(?:bill|ref|receipt|voucher)\s*(?:no\.?|#)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9/\-]{3,30})",
        r"\b([A-Z]{2,6}/\d{2}-\d{2}/\d{4,6})\b",
        r"\b([A-Z]{2,6}-\d{4,8})\b",
    ]
    for p in patterns:
        for m in re.findall(p, text, re.IGNORECASE):
            ids.add(m.strip().upper())
    return ids
 
 
# ---------------------------------------------------------
# VENDOR DETECTION
# ---------------------------------------------------------
 
def detect_vendor(text):
    keywords = [
        "uber", "ola", "rapido",
        "zomato", "swiggy", "restaurant", "cafe",
        "amazon", "flipkart",
        "airtel", "jio", "bsnl", "vodafone",
        "hotel", "lodge",
    ]
    t = text.lower()
    for v in keywords:
        if v in t:
            return v
    lines = [l.strip() for l in text.split("\n") if len(l.strip()) > 4][:5]
    for l in lines:
        if any(w in l.lower() for w in ["ltd", "llp", "pvt", "inc", "co.", "corp"]):
            return l[:60]
    return None
 
 
# ---------------------------------------------------------
# PROCESS FILES
# ---------------------------------------------------------
 
def process_image(file_bytes, name):
    img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
    img = preprocess(img)
    text = ocr_image(img)
    amts = extract_amounts_with_context(text)
    return {"name": name, "amounts_detail": amts, "amounts": [a["amount"] for a in amts],
            "invoice_ids": extract_invoice_ids(text), "vendor": detect_vendor(text), "text": text}
 
def process_pdf(file_bytes, name):
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for p in pdf.pages[:5]:
                text += (p.extract_text() or "") + "\n"
    except:
        pass
    amts = extract_amounts_with_context(text)
    return {"name": name, "amounts_detail": amts, "amounts": [a["amount"] for a in amts],
            "invoice_ids": extract_invoice_ids(text), "vendor": detect_vendor(text), "text": text}
 
def process_file(file_bytes, name):
    ext = name.split(".")[-1].lower()
    if ext == "pdf":
        return process_pdf(file_bytes, name)
    return process_image(file_bytes, name)
 
 
# ---------------------------------------------------------
# MATCHING LOGIC (v4)
# ---------------------------------------------------------
 
AMOUNT_TOLERANCE = 0.50
 
def best_amount_from_doc(doc):
    if not doc["amounts_detail"]:
        return None
    return sorted(doc["amounts_detail"], key=lambda x: x["priority"])[0]["amount"]
 
def amount_match_result(reg_amt, doc):
    try:
        ra = float(reg_amt)
        if np.isnan(ra) or np.isinf(ra):
            return "not_found", None, None
    except:
        return "not_found", None, None
    doc_amt = best_amount_from_doc(doc)
    if doc_amt is None:
        return "not_found", None, None
    diff = abs(ra - doc_amt)
    if diff <= AMOUNT_TOLERANCE:
        return "exact", doc_amt, diff
    return "mismatch", doc_amt, diff
 
def id_match(report_id, doc):
    rid = str(report_id).strip().upper()
    if not rid:
        return False
    if rid in doc["text"].upper():
        return True
    for did in doc["invoice_ids"]:
        if rid == did or rid in did or did in rid:
            return True
    return False
 
def vendor_match_score(reg_vendor, doc):
    rv = str(reg_vendor).lower().strip()
    dv = str(doc["vendor"] or "").lower().strip()
    text = doc["text"].lower()
    if not rv:
        return 0
    rv_tokens = set(rv.replace("&", " ").replace(".", " ").split())
    score = 0
    for tok in rv_tokens:
        if len(tok) > 2 and tok in text:
            score += 1
    if dv and rv and (rv in dv or dv in rv):
        score += 3
    return score
 
def run_vouching(df, docs):
    results = []
    used = set()
 
    for i, row in df.iterrows():
        amount    = row["Expense Amount"]
        vendor    = row["Vendor"]
        category  = row["Category"]
        report_id = row["ExpenseReport ID"]
 
        candidates = []
        for d in docs:
            id_matched = id_match(report_id, d)
            vend_score = vendor_match_score(vendor, d)
            amt_status, doc_amt, diff = amount_match_result(amount, d)
            score = 0
            if id_matched: score += 20
            score += min(vend_score, 5)
            if amt_status == "exact": score += 10
            elif amt_status == "mismatch": score += 3
            candidates.append({"doc": d, "score": score, "id_matched": id_matched,
                                "amt_status": amt_status, "doc_amt": doc_amt, "diff": diff})
 
        candidates.sort(key=lambda x: -x["score"])
        best = candidates[0] if candidates and candidates[0]["score"] > 0 else None
 
        try:
            amt_val = float(amount)
            if np.isnan(amt_val) or np.isinf(amt_val):
                amt_val = 0.0
        except:
            amt_val = 0.0
 
        if best is None:
            status = "MISSING_DOC"
            matched_file = "-"; doc_amount = "-"; amount_diff = "-"; confidence = "0/35"
        else:
            matched_file = best["doc"]["name"]
            doc_amount   = f"₹{best['doc_amt']:,.2f}" if best["doc_amt"] is not None else "-"
            amount_diff  = f"₹{best['diff']:,.2f}"    if best["diff"]    is not None else "-"
            confidence   = f"{best['score']}/35"
            if matched_file in used:
                status = "DUPLICATE_RECEIPT"
            elif best["amt_status"] == "exact":
                status = "MATCHED"
            elif best["amt_status"] == "mismatch":
                status = "AMOUNT_MISMATCH"
            else:
                status = "MATCHED" if best["id_matched"] else "MISSING_DOC"
            if status != "DUPLICATE_RECEIPT":
                used.add(matched_file)
 
        results.append({
            "Report ID":       str(report_id),
            "Register Amount": amt_val,
            "Document Amount": doc_amount,
            "Difference":      amount_diff,
            "Vendor":          str(vendor),
            "Category":        str(category),
            "Matched File":    matched_file,
            "ID Match":        "✓" if (best and best["id_matched"]) else "✗",
            "Confidence":      confidence,
            "Status":          status,
        })
 
    return results
 
 
# ---------------------------------------------------------
# EXCEL EXPORT
# ---------------------------------------------------------
 
def export_excel(df):
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = (pd.to_numeric(df[col], errors="coerce")
                         .replace([np.inf, -np.inf], np.nan).fillna(0))
        else:
            df[col] = df[col].fillna("").astype(str)
 
    buffer = io.BytesIO()
    wb = xlsxwriter.Workbook(buffer, {"nan_inf_to_errors": True, "in_memory": True})
 
    KKC_GREEN = "#5ba632"
    hdr  = wb.add_format({"bold": True, "bg_color": KKC_GREEN, "font_color": "#ffffff",
                           "font_name": "Calibri", "font_size": 11,
                           "valign": "vcenter", "align": "center"})
    base = wb.add_format({"font_name": "Calibri", "font_size": 10, "valign": "vcenter"})
    num_f= wb.add_format({"font_name": "Calibri", "font_size": 10, "valign": "vcenter",
                           "num_format": "₹#,##0.00"})
    ok_f = wb.add_format({"font_name": "Calibri", "font_size": 10,
                           "font_color": KKC_GREEN, "bold": True, "valign": "vcenter"})
    mi_f = wb.add_format({"font_name": "Calibri", "font_size": 10,
                           "font_color": "#d13030", "bold": True, "valign": "vcenter"})
    mm_f = wb.add_format({"font_name": "Calibri", "font_size": 10,
                           "font_color": "#e07b00", "bold": True, "valign": "vcenter"})
    du_f = wb.add_format({"font_name": "Calibri", "font_size": 10,
                           "font_color": "#b8860b", "bold": True, "valign": "vcenter"})
 
    ws = wb.add_worksheet("Vouching Results")
    ws.set_tab_color(KKC_GREEN)
    ws.set_row(0, 24)
    ws.hide_gridlines(2)
 
    col_widths = {"Report ID": 20, "Register Amount": 18, "Document Amount": 18,
                  "Difference": 14, "Vendor": 24, "Category": 16,
                  "Matched File": 34, "ID Match": 10, "Confidence": 12, "Status": 22}
    for ci, col in enumerate(df.columns):
        ws.set_column(ci, ci, col_widths.get(col, 16))
        ws.write(0, ci, col, hdr)
 
    status_ci  = list(df.columns).index("Status")         if "Status"          in df.columns else -1
    reg_amt_ci = list(df.columns).index("Register Amount") if "Register Amount" in df.columns else -1
 
    for ri in range(len(df)):
        status = str(df.iloc[ri].get("Status", ""))
        sfmt = {"MATCHED": ok_f, "MISSING_DOC": mi_f, "AMOUNT_MISMATCH": mm_f, "DUPLICATE_RECEIPT": du_f}.get(status, base)
        for ci in range(len(df.columns)):
            raw = df.iloc[ri, ci]
            fmt = sfmt if ci == status_ci else (num_f if ci == reg_amt_ci else base)
            if isinstance(raw, (int, np.integer)):
                ws.write_number(ri+1, ci, int(raw), fmt)
            elif isinstance(raw, (float, np.floating)):
                v = float(raw)
                ws.write_string(ri+1, ci, "", fmt) if (np.isnan(v) or np.isinf(v)) else ws.write_number(ri+1, ci, v, fmt)
            else:
                ws.write_string(ri+1, ci, str(raw) if raw is not None else "", fmt)
 
    sw = wb.add_worksheet("Summary")
    sw.set_tab_color("#a8d870")
    sw.set_column(0, 0, 28); sw.set_column(1, 1, 16)
    sw.hide_gridlines(2)
 
    t_fmt = wb.add_format({"font_name": "Calibri", "font_size": 15, "bold": True, "font_color": "#1c2b1c"})
    l_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "font_color": "#7a8e7a"})
    v_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True, "font_color": "#1c2b1c"})
    g_fmt = wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True, "font_color": KKC_GREEN})
 
    sw.write(0, 0, "KKC Vouching Report — Summary", t_fmt)
    sw.write(1, 0, f"KKC & Associates LLP  ·  Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}", l_fmt)
 
    total    = len(df)
    matched  = int((df["Status"] == "MATCHED").sum())
    missing  = int((df["Status"] == "MISSING_DOC").sum())
    mismatch = int((df["Status"] == "AMOUNT_MISMATCH").sum())
    dup      = int((df["Status"] == "DUPLICATE_RECEIPT").sum())
    rate     = round(matched / total * 100, 1) if total else 0.0
 
    for i, (lbl, val, fmt) in enumerate([
        ("Total Entries",      total,    v_fmt),
        ("Matched",            matched,  g_fmt),
        ("Amount Mismatches",  mismatch, v_fmt),
        ("Missing Documents",  missing,  v_fmt),
        ("Duplicate Receipts", dup,      v_fmt),
        ("Match Rate (%)",     rate,     g_fmt),
    ]):
        sw.write(i+3, 0, lbl, l_fmt)
        sw.write_number(i+3, 1, float(val), fmt)
 
    wb.close(); buffer.seek(0)
    return buffer.read()
 
 
# ---------------------------------------------------------
# PAGE HEADER (KKC styled)
# ---------------------------------------------------------
 
st.markdown("""
<div class="vt-header">
    <div class="vt-topstrip">
        <span>📞 +91 22 6143 7333</span>
        <span class="vt-topstrip-sep">|</span>
        <span>✉ info@kkcllp.in</span>
        <span class="vt-topstrip-sep">|</span>
        <span>Vouching Tool — Internal Audit Suite</span>
    </div>
    <div class="vt-navbar" style="padding-left:0.5rem;padding-right:0.5rem">
        <div class="vt-brand">
            <div class="vt-brand-name">kkc &amp; associates llp</div>
            <div class="vt-brand-tag">Chartered Accountants</div>
            <div class="vt-brand-former">(Formerly Khimji Kunverji &amp; Co LLP)</div>
        </div>
        <div class="vt-nav-divider"></div>
        <div>
            <div class="vt-tool-badge">🔍 Vouching Engine</div>
            <div class="vt-tool-sub">Expense Audit &amp; Receipt Matching</div>
        </div>
        <div class="vt-status-pill">
            <div class="vt-status-dot"></div>
            Session Active
        </div>
    </div>
</div>
""", unsafe_allow_html=True)
 
st.markdown("""
<div class="vt-infobox">
    <span style="font-size:16px;flex-shrink:0">ℹ</span>
    <span>
        Upload an <strong>Expense Register</strong> (.xlsx) and one or more
        <strong>Receipts / Invoices</strong> (PDF, PNG, JPG).
        The engine matches each entry by <strong>Invoice ID → Vendor → Amount</strong>.
        <strong>AMOUNT_MISMATCH</strong> flags entries where a document was found but amounts differ —
        discrepancies are never silently passed as matched.
    </span>
</div>
""", unsafe_allow_html=True)
 
 
# ---------------------------------------------------------
# UPLOAD
# ---------------------------------------------------------
 
st.markdown('<div class="vt-section">Document Intake</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2, gap="medium")
 
with col1:
    st.markdown('<div class="vt-upload-label">Expense Register <span class="badge">XLSX</span></div>', unsafe_allow_html=True)
    register = st.file_uploader("exp_reg", type=["xlsx"], label_visibility="collapsed", key="reg")
    if register:
        st.markdown(f'<div class="vt-file-ok">✓ &nbsp;{register.name}</div>', unsafe_allow_html=True)
 
with col2:
    st.markdown('<div class="vt-upload-label">Receipts &amp; Invoices <span class="badge">PDF / PNG / JPG</span></div>', unsafe_allow_html=True)
    docs = st.file_uploader("receipts", accept_multiple_files=True, label_visibility="collapsed", key="docs")
    if docs:
        st.markdown(f'<div class="vt-file-ok">✓ &nbsp;{len(docs)} file{"s" if len(docs)!=1 else ""} ready</div>', unsafe_allow_html=True)
 
 
# ---------------------------------------------------------
# MAIN FLOW
# ---------------------------------------------------------
 
if register and docs:
 
    st.markdown('<div class="vt-section">Register Preview</div>', unsafe_allow_html=True)
    df = pd.read_excel(register)
    st.dataframe(df.head(8), use_container_width=True, hide_index=True)
    st.markdown(f'<div class="vt-preview-meta">{len(df)} entries · {len(df.columns)} columns</div>', unsafe_allow_html=True)
 
    st.markdown('<div class="vt-section">Processing Documents</div>', unsafe_allow_html=True)
    files_data = []
    prog_text  = st.empty()
    prog_bar   = st.progress(0)
 
    for i, f in enumerate(docs):
        prog_text.markdown(
            f'<div class="vt-proc-label">Extracting → <strong>{f.name}</strong> [{i+1}/{len(docs)}]</div>',
            unsafe_allow_html=True
        )
        files_data.append(process_file(f.read(), f.name))
        prog_bar.progress((i + 1) / len(docs))
 
    prog_text.markdown('<div class="vt-proc-done">✓ &nbsp;All documents processed successfully</div>', unsafe_allow_html=True)
    st.markdown('<div class="vt-divider"></div>', unsafe_allow_html=True)
 
    btn_col, _ = st.columns([1, 4])
    with btn_col:
        run = st.button("🔍  Run Vouching Analysis")
 
    if run:
        with st.spinner("Matching expense entries to documents…"):
            results = run_vouching(df, files_data)
            rdf     = pd.DataFrame(results)
 
        n_matched  = int((rdf.Status == "MATCHED").sum())
        n_missing  = int((rdf.Status == "MISSING_DOC").sum())
        n_mismatch = int((rdf.Status == "AMOUNT_MISMATCH").sum())
        n_dup      = int((rdf.Status == "DUPLICATE_RECEIPT").sum())
        total      = len(rdf)
        rate       = round(n_matched / total * 100) if total else 0
 
        st.markdown('<div class="vt-section">Summary</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="vt-metrics">
            <div class="vt-metric matched">
                <div class="vt-metric-label">Matched</div>
                <div class="vt-metric-value">{n_matched}</div>
                <div class="vt-metric-sub">{rate}% match rate</div>
            </div>
            <div class="vt-metric mismatch">
                <div class="vt-metric-label">Amount Mismatch</div>
                <div class="vt-metric-value">{n_mismatch}</div>
                <div class="vt-metric-sub">Doc found, amount differs</div>
            </div>
            <div class="vt-metric missing">
                <div class="vt-metric-label">Missing Documents</div>
                <div class="vt-metric-value">{n_missing}</div>
                <div class="vt-metric-sub">No receipt found</div>
            </div>
            <div class="vt-metric duplicate">
                <div class="vt-metric-label">Duplicates</div>
                <div class="vt-metric-value">{n_dup}</div>
                <div class="vt-metric-sub">Reused receipts</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
 
        if n_mismatch > 0:
            rows = rdf[rdf["Status"] == "AMOUNT_MISMATCH"]
            table_rows = ""
            for _, r in rows.iterrows():
                table_rows += f"""
                <tr>
                    <td>{r["Report ID"]}</td>
                    <td class="val-reg">₹{r["Register Amount"]:,.2f}</td>
                    <td class="val-doc">{r["Document Amount"]}</td>
                    <td class="val-diff">{r["Difference"]}</td>
                    <td style="color:#6b3a00">{r["Matched File"]}</td>
                </tr>"""
            st.markdown(f"""
            <div class="vt-alert warn">
                <span style="font-size:18px;flex-shrink:0">⚠</span>
                <div style="width:100%">
                    <strong>{n_mismatch} amount mismatch{"es" if n_mismatch>1 else ""} require manual review</strong>
                    — a supporting document was identified, but the amounts do not agree.
                    <table class="mismatch-table">
                        <thead><tr>
                            <th>Report ID</th><th>Register Amount</th>
                            <th>Document Amount</th><th>Difference</th><th>Matched File</th>
                        </tr></thead>
                        <tbody>{table_rows}</tbody>
                    </table>
                </div>
            </div>
            """, unsafe_allow_html=True)
 
        st.markdown('<div class="vt-section">Detailed Results</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="vt-result-head">Vouching Report <span class="count-tag">{total} entries</span></div>', unsafe_allow_html=True)
 
        def style_status(val):
            return {
                "MATCHED":           "color:#5ba632;font-weight:800",
                "MISSING_DOC":       "color:#d13030;font-weight:800",
                "AMOUNT_MISMATCH":   "color:#e07b00;font-weight:800",
                "DUPLICATE_RECEIPT": "color:#b8860b;font-weight:800",
            }.get(val, "")
 
        st.dataframe(
            rdf.style.applymap(style_status, subset=["Status"]),
            use_container_width=True, hide_index=True
        )
 
        st.markdown('<div class="vt-divider"></div>', unsafe_allow_html=True)
        dl_col, ts_col = st.columns([1, 3])
        with dl_col:
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                label="⬇  Download Audit Report (.xlsx)",
                data=export_excel(rdf),
                file_name=f"KKC_vouching_report_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with ts_col:
            st.markdown(
                f'<div class="vt-timestamp">KKC & Associates LLP &nbsp;·&nbsp; '
                f'Generated {datetime.now().strftime("%d %b %Y · %H:%M")} &nbsp;·&nbsp; '
                f'{total} entries &nbsp;·&nbsp; 2 sheets: Results + Summary</div>',
                unsafe_allow_html=True
            )
 
else:
    st.markdown("""
    <div class="vt-empty">
        <div class="vt-empty-icon">🔍</div>
        <div class="vt-empty-title">Ready to begin vouching</div>
        <div class="vt-empty-sub">Upload your expense register and receipts above to start</div>
    </div>
    """, unsafe_allow_html=True)
 
st.markdown("""
<div class="vt-footer">
    <strong>kkc &amp; associates llp</strong> &nbsp;·&nbsp; Chartered Accountants &nbsp;·&nbsp;
    (Formerly Khimji Kunverji &amp; Co LLP) &nbsp;·&nbsp; Internal Audit Suite v4.0
</div>
""", unsafe_allow_html=True)
 
