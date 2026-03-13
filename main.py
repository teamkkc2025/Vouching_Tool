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
# PAGE CONFIG  — must be first Streamlit call
# ---------------------------------------------------------

st.set_page_config(
    page_title="KKC Vouching Tool",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# CSS  — single block, <link> for fonts (avoids CSP issues)
# ---------------------------------------------------------

st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700;800;900&family=Libre+Baskerville:wght@400;700&display=swap" rel="stylesheet">

<style>
:root {
    --green:      #5ba632;
    --green-mid:  #6dbc3c;
    --green-bar:  #76b82a;
    --green-lt:   #e8f5e0;
    --white:      #ffffff;
    --gray-50:    #f7f9f7;
    --gray-100:   #eef2ee;
    --gray-200:   #dde5dd;
    --gray-400:   #9aaa9a;
    --gray-600:   #5a6b5a;
    --text:       #1c2b1c;
    --text-mid:   #445544;
    --text-light: #7a8e7a;
}

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    background-color: var(--gray-50) !important;
    font-family: 'Nunito Sans', sans-serif !important;
    color: var(--text) !important;
}
[data-testid="stHeader"],
[data-testid="stToolbar"],
footer { display: none !important; }

.block-container {
    padding: 0 2rem 3rem !important;
    max-width: 1300px !important;
}

[data-testid="stSidebar"] {
    background-color: var(--white) !important;
    border-right: 2px solid var(--green-lt) !important;
}
[data-testid="stSidebar"] > div:first-child { padding-top: 0 !important; }

[data-testid="stButton"] > button {
    background-color: var(--green) !important;
    color: #fff !important;
    font-family: 'Nunito Sans', sans-serif !important;
    font-weight: 800 !important; font-size: 14px !important;
    border: none !important; border-radius: 8px !important;
    padding: 12px 32px !important;
    box-shadow: 0 2px 8px rgba(91,166,50,0.3) !important;
    transition: all 0.18s !important;
}
[data-testid="stButton"] > button:hover {
    background-color: #4e9429 !important;
    transform: translateY(-1px) !important;
}
[data-testid="stDownloadButton"] > button {
    background-color: var(--white) !important;
    color: var(--green) !important;
    font-family: 'Nunito Sans', sans-serif !important;
    font-weight: 800 !important; font-size: 13px !important;
    border: 2px solid var(--green) !important;
    border-radius: 8px !important; padding: 10px 26px !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background-color: var(--green-lt) !important;
}
[data-testid="stProgressBar"] > div {
    background-color: var(--gray-100) !important;
    border-radius: 99px !important; height: 5px !important;
}
[data-testid="stProgressBar"] > div > div {
    background-color: var(--green) !important; border-radius: 99px !important;
}
[data-testid="stDataFrame"] {
    border: 1px solid var(--gray-200) !important;
    border-radius: 10px !important; overflow: hidden !important;
    box-shadow: 0 1px 6px rgba(0,0,0,0.04) !important;
}
[data-testid="stFileUploader"] {
    background-color: var(--white) !important;
    border: 2px dashed var(--gray-200) !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--green) !important;
    background-color: var(--green-lt) !important;
}
label, [data-testid="stWidgetLabel"] p {
    font-family: 'Nunito Sans', sans-serif !important;
    font-size: 13px !important; font-weight: 700 !important;
    color: var(--text-mid) !important;
}
[data-testid="stSpinner"] { color: var(--green) !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--gray-50); }
::-webkit-scrollbar-thumb { background: var(--gray-200); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--green); }

/* ── Custom components ─── */
.kkc-topstrip {
    background: var(--green-bar);
    padding: 7px 16px;
    display: flex; align-items: center; gap: 24px;
    font-size: 11.5px; color: #fff; font-weight: 600;
    margin: 0 -2rem; letter-spacing: 0.1px;
}
.kkc-navbar {
    background: var(--white);
    border-bottom: 3px solid var(--green-bar);
    padding: 14px 2rem 12px;
    margin: 0 -2rem;
    display: flex; align-items: center; gap: 20px;
}
.kkc-brand-name {
    font-family: 'Nunito Sans', sans-serif;
    font-size: 22px; font-weight: 900;
    color: var(--green); letter-spacing: -0.5px; line-height: 1;
}
.kkc-brand-tag  { font-size: 10.5px; color: var(--text-light); margin-top: 2px; }
.kkc-brand-fmr  { font-size: 9.5px; color: var(--gray-400); font-style: italic; margin-top: 1px; }
.kkc-divider-v  { width: 1px; height: 38px; background: var(--gray-200); flex-shrink: 0; }
.kkc-tool-badge {
    background: var(--green-lt); border: 1.5px solid var(--green);
    border-radius: 8px; padding: 7px 16px;
    font-size: 13px; font-weight: 800; color: var(--green);
}
.kkc-tool-sub   { font-size: 10px; color: var(--text-light); margin-top: 3px; }
.kkc-pill {
    margin-left: auto; background: var(--green-lt);
    border: 1px solid var(--green); border-radius: 20px;
    padding: 5px 16px; font-size: 10.5px; font-weight: 700; color: var(--green);
    display: flex; align-items: center; gap: 6px;
}
.kkc-dot {
    width: 7px; height: 7px; border-radius: 50%; background: var(--green);
    display: inline-block; animation: blink 1.8s ease-in-out infinite;
}
@keyframes blink { 0%,100%{opacity:1} 50%{opacity:0.3} }

.kkc-section {
    font-size: 9.5px; font-weight: 700; letter-spacing: 2px;
    text-transform: uppercase; color: var(--green);
    margin: 28px 0 14px;
    display: flex; align-items: center; gap: 12px;
}
.kkc-section::after { content: ''; flex: 1; height: 1.5px; background: var(--green-lt); }

.kkc-infobox {
    background: var(--green-lt); border: 1px solid var(--green);
    border-left: 4px solid var(--green);
    border-radius: 0 9px 9px 0;
    padding: 14px 18px; font-size: 12.5px; color: var(--text);
    display: flex; gap: 12px; align-items: flex-start;
    margin-bottom: 22px; line-height: 1.65;
}

.kkc-colhint {
    background: var(--white); border: 1px solid var(--gray-200);
    border-radius: 10px; padding: 16px 20px; margin-bottom: 20px;
}
.kkc-colhint-title {
    font-size: 10px; font-weight: 700; letter-spacing: 1.5px;
    text-transform: uppercase; color: var(--green); margin-bottom: 12px;
}
.kkc-col-grid {
    display: grid; grid-template-columns: repeat(5, 1fr); gap: 8px;
}
.kkc-col-item {
    background: var(--gray-50); border: 1px solid var(--gray-100);
    border-radius: 7px; padding: 10px 12px;
}
.kkc-col-num  { font-size: 9px; font-weight: 700; color: var(--green); margin-bottom: 3px; }
.kkc-col-name { font-size: 11px; font-weight: 800; color: var(--text); font-family: monospace; }
.kkc-col-desc { font-size: 10px; color: var(--text-light); margin-top: 3px; }
.kkc-col-req  { font-size: 8.5px; color: var(--green); font-weight: 700; margin-top: 4px; }
.kkc-col-opt  { font-size: 8.5px; color: var(--gray-400); font-weight: 700; margin-top: 4px; }

.kkc-file-ok {
    font-size: 10.5px; font-weight: 700; color: var(--green);
    background: var(--green-lt); border: 1px solid var(--green);
    padding: 7px 14px; border-radius: 7px; margin-top: 8px;
}
.kkc-file-error {
    font-size: 10.5px; font-weight: 700; color: #d13030;
    background: #fff5f5; border: 1px solid #fca5a5;
    padding: 7px 14px; border-radius: 7px; margin-top: 8px;
}
.kkc-upload-label {
    font-size: 12px; font-weight: 700; color: var(--text-mid);
    margin-bottom: 8px; display: flex; align-items: center; gap: 8px;
}
.kkc-badge {
    font-size: 9px; background: var(--gray-100); color: var(--text-light);
    padding: 2px 8px; border-radius: 4px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.5px;
}

.kkc-metrics {
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 14px; margin: 6px 0 28px;
}
.kkc-metric {
    background: var(--white); border: 1px solid var(--gray-200);
    border-top: 4px solid transparent;
    border-radius: 10px; padding: 22px 20px 18px;
    transition: box-shadow 0.2s, transform 0.2s;
}
.kkc-metric:hover { box-shadow: 0 6px 24px rgba(91,166,50,0.1); transform: translateY(-2px); }
.kkc-metric.matched   { border-top-color: var(--green); }
.kkc-metric.mismatch  { border-top-color: #e07b00; }
.kkc-metric.missing   { border-top-color: #d13030; }
.kkc-metric.duplicate { border-top-color: #b8860b; }
.kkc-metric-label {
    font-size: 9.5px; font-weight: 700; letter-spacing: 1.5px;
    text-transform: uppercase; color: var(--text-light); margin-bottom: 10px;
}
.kkc-metric-value {
    font-family: 'Libre Baskerville', serif;
    font-size: 48px; font-weight: 700; line-height: 1; letter-spacing: -2px;
}
.kkc-metric.matched   .kkc-metric-value { color: var(--green); }
.kkc-metric.mismatch  .kkc-metric-value { color: #e07b00; }
.kkc-metric.missing   .kkc-metric-value { color: #d13030; }
.kkc-metric.duplicate .kkc-metric-value { color: #b8860b; }
.kkc-metric-sub { font-size: 11px; color: var(--text-light); margin-top: 7px; }

.kkc-alert-warn {
    background: #fff8f0; border: 1px solid #f5c57a;
    border-left: 4px solid #e07b00;
    border-radius: 0 9px 9px 0;
    padding: 16px 20px; font-size: 12.5px;
    color: #6b3a00; margin-bottom: 24px; line-height: 1.6;
}
.kkc-mismatch-table {
    width: 100%; margin-top: 12px; border-collapse: collapse; font-size: 11.5px;
}
.kkc-mismatch-table th {
    text-align: left; font-size: 9px; font-weight: 700;
    letter-spacing: 1.3px; text-transform: uppercase;
    color: #a06020; padding: 4px 10px 8px 0;
}
.kkc-mismatch-table td { padding: 6px 10px 6px 0; border-top: 1px solid #fde8c8; }
.kkc-mismatch-table tr:first-child td { border-top: none; }

.kkc-result-head {
    font-size: 16px; font-weight: 800; color: var(--text);
    letter-spacing: -0.3px; margin: 4px 0 16px;
    display: flex; align-items: center; gap: 12px;
}
.kkc-count-tag {
    font-size: 10px; font-weight: 700;
    background: var(--green-lt); color: var(--green);
    border: 1px solid var(--green); padding: 3px 12px; border-radius: 20px;
}
.kkc-divider  { height: 1px; background: var(--gray-100); margin: 24px 0; }
.kkc-preview-meta { font-size: 10.5px; color: var(--text-light); margin-top: 6px; }
.kkc-proc-label   { font-size: 11px; color: var(--text-light); margin-bottom: 8px; }
.kkc-proc-done    { font-size: 11px; color: var(--green); font-weight: 700; }
.kkc-timestamp    { font-size: 10.5px; color: var(--text-light); padding-top: 11px; }
.kkc-footer {
    background: var(--white); border-top: 1px solid var(--gray-100);
    padding: 14px 0; text-align: center;
    font-size: 11px; color: var(--text-light); margin-top: 48px;
}
.kkc-empty { text-align: center; padding: 72px 24px; }
.kkc-empty-icon  { font-size: 52px; margin-bottom: 16px; }
.kkc-empty-title { font-size: 17px; font-weight: 800; color: var(--gray-400); margin-bottom: 8px; }
.kkc-empty-sub   { font-size: 11.5px; color: var(--gray-200); }

/* Sidebar */
.sb-topbar { background: var(--green-bar); height: 5px; margin: 0; }
.sb-brand  { padding: 20px 18px 16px; border-bottom: 1px solid var(--gray-100); }
.sb-name   { font-family:'Nunito Sans',sans-serif; font-size:17px; font-weight:900; color:var(--green); letter-spacing:-0.3px; line-height:1; }
.sb-sub    { font-size:10px; color:var(--text-light); margin-top:3px; }
.sb-fmr    { font-size:9.5px; color:var(--gray-400); font-style:italic; margin-top:1px; }
.sb-sec    { font-size:9px; font-weight:700; letter-spacing:1.8px; text-transform:uppercase; color:var(--gray-400); padding:14px 14px 6px; }
.sb-nav    { padding:6px 10px; }
.sb-item   { display:flex; align-items:center; gap:10px; padding:9px 12px; border-radius:7px; font-size:13px; color:var(--text-mid); font-weight:600; margin-bottom:2px; }
.sb-item.active { background:var(--green-lt); color:var(--green); border-left:3px solid var(--green); padding-left:9px; font-weight:800; }
.sb-hr     { height:1px; background:var(--gray-100); margin:8px 14px; }
.sb-block  { margin:8px 14px; background:var(--gray-50); border:1px solid var(--gray-100); border-radius:10px; padding:14px; }
.sb-block-title { font-size:9px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; color:var(--gray-400); margin-bottom:10px; }
.sb-row    { display:flex; justify-content:space-between; align-items:center; padding:5px 0; border-bottom:1px solid var(--gray-100); font-size:11.5px; }
.sb-row:last-child { border-bottom:none; }
.sb-key    { color:var(--text-light); }
.sb-val    { font-weight:700; font-size:11px; color:var(--text-mid); }
.sb-val.ok { color:var(--green); }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------

with st.sidebar:
    st.markdown("""
    <div class="sb-topbar"></div>
    <div class="sb-brand">
        <div class="sb-name">kkc &amp; associates llp</div>
        <div class="sb-sub">Chartered Accountants</div>
        <div class="sb-fmr">(Formerly Khimji Kunverji &amp; Co LLP)</div>
    </div>
    <div class="sb-sec">Navigation</div>
    <div class="sb-nav">
        <div class="sb-item active">🔍&nbsp; Vouching Engine</div>
        <div class="sb-item">📊&nbsp; Analytics</div>
        <div class="sb-item">🗂️&nbsp; Report History</div>
        <div class="sb-item">⚙️&nbsp; Settings</div>
    </div>
    <div class="sb-hr"></div>
    <div class="sb-sec">Engine Status</div>
    <div class="sb-block">
        <div class="sb-row"><span class="sb-key">OCR Engine</span><span class="sb-val ok">● Active</span></div>
        <div class="sb-row"><span class="sb-key">PDF Parser</span><span class="sb-val ok">● Active</span></div>
        <div class="sb-row"><span class="sb-key">Algorithm</span><span class="sb-val">v5.0</span></div>
        <div class="sb-row"><span class="sb-key">Tolerance</span><span class="sb-val">± ₹0.50</span></div>
        <div class="sb-row"><span class="sb-key">ID Matching</span><span class="sb-val ok">Enabled</span></div>
        <div class="sb-row"><span class="sb-key">Gate Filter</span><span class="sb-val ok">Strict</span></div>
    </div>
    <div style="padding:20px 14px 0;font-size:10px;color:#ccc;text-align:center;">
        © 2025 KKC &amp; Associates LLP
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------
# TESSERACT
# ---------------------------------------------------------

try:
    if os.name == "nt":
        p = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
    pytesseract.get_tesseract_version()
    OCR_OK = True
except Exception:
    OCR_OK = False


# ---------------------------------------------------------
# OCR helpers
# ---------------------------------------------------------

def ocr_image(img):
    try:
        return pytesseract.image_to_string(img, config="--psm 6 --oem 3")
    except Exception:
        return ""

def preprocess(img):
    gray = img.convert("L")
    gray = ImageEnhance.Contrast(gray).enhance(2)
    gray = gray.resize((gray.width * 2, gray.height * 2))
    return gray


# ---------------------------------------------------------
# Amount extraction — grand total priority
#
# Priority 1 : Gross Amount / Balance Due / Grand Total / Total INR / Amount Due
# Priority 2 : Plain "Total" lines (not tax/sub-total)
# Priority 3 : Currency-prefixed amounts
# Priority 4 : Bare numbers (fallback only)
#
# KEY FIXES:
#   1. Regex \d+ prefix — "40000.00" not mis-parsed as "400"+"00"
#   2. Multi-column Total rows → LAST (rightmost) number = payable col
#      e.g. "Total  40000.00  41000.00" → 41000.00 (Amount incl. GST)
#   3. Total label with no inline number → look-ahead up to 3 lines
#   4. Cap raised to 100 Crore for large invoices
# ---------------------------------------------------------

_EXCLUDE_PAT = re.compile(
    r"(?:"
    r"sub[\s\-]?total"
    r"|taxable[\s\-]?(?:value|amount)"
    r"|(?:add\s*:?\s*)?(?:output\s+)?(?:cgst|sgst|igst|ugst|utgst)"
    r"|cess|tax\s+amount|tds|rate\s+per"
    r"|(?:per\s+)?(?:kg|unit|pc|pcs|litre|mtr)"
    r"|hsn|qty|quantity"
    r")",
    re.IGNORECASE
)

_GRAND_TOTAL_PAT = re.compile(
    r"(?:"
    r"gross\s+amount|balance\s+due|grand\s+total"
    r"|total\s+(?:inr|rs\.?|rupees?|amount|payable|due)"
    r"|amount\s+(?:due|payable|chargeable)"
    r"|net\s+(?:payable|amount|total)|invoice\s+total"
    r")",
    re.IGNORECASE
)

_TOTAL_PAT   = re.compile(r"\btotal\b",                    re.IGNORECASE)
_CURR_PAT    = re.compile(r"(?:₹|inr|rs\.?|rupees?)",     re.IGNORECASE)
_IND_NUM_PAT = re.compile(r"\d+(?:,\d{2,3})*(?:\.\d{1,2})?")


def _is_excluded(line: str) -> bool:
    return bool(_EXCLUDE_PAT.search(line))


def _parse_indian(s: str):
    try:
        v = float(s.replace(",", ""))
        return round(v, 2) if 1.0 <= v <= 1_000_000_000 else None
    except Exception:
        return None


def _nums_on_line(line: str) -> list:
    return [v for m in _IND_NUM_PAT.findall(line)
            if (v := _parse_indian(m)) is not None]


def _pick(nums: list) -> float:
    """Rightmost number on a Total row = final payable column."""
    return nums[-1] if nums else 0.0


def extract_amounts_with_context(text: str) -> list:
    results = []
    lines   = text.splitlines()

    for i, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line or _is_excluded(line):
            continue
        nums = _nums_on_line(line)

        if _GRAND_TOTAL_PAT.search(line):
            if nums:
                results.append({"amount": _pick(nums), "priority": 1})
                results.append({"amount": max(nums),   "priority": 1})
            else:
                for j in range(i + 1, min(i + 4, len(lines))):
                    nxt = lines[j].strip()
                    if not nxt or _is_excluded(nxt):
                        continue
                    nxt_nums = _nums_on_line(nxt)
                    if nxt_nums:
                        results.append({"amount": _pick(nxt_nums), "priority": 1})
                        break

        elif _TOTAL_PAT.search(line):
            if nums:
                results.append({"amount": _pick(nums), "priority": 2})
                results.append({"amount": max(nums),   "priority": 2})
            else:
                for j in range(i + 1, min(i + 4, len(lines))):
                    nxt = lines[j].strip()
                    if not nxt or _is_excluded(nxt):
                        continue
                    nxt_nums = _nums_on_line(nxt)
                    if nxt_nums:
                        results.append({"amount": _pick(nxt_nums), "priority": 2})
                        break

        elif _CURR_PAT.search(line) and nums:
            results.append({"amount": _pick(nums), "priority": 3})

    # Priority 4 fallback — only if nothing better found
    if not any(r["priority"] <= 2 for r in results):
        for raw_line in lines:
            line = raw_line.strip()
            if not line or _is_excluded(line):
                continue
            for v in _nums_on_line(line):
                if v >= 100:
                    results.append({"amount": v, "priority": 4})

    if not results:
        return []

    seen: dict = {}
    for r in results:
        k = r["amount"]
        if k not in seen or r["priority"] < seen[k]["priority"]:
            seen[k] = r

    best   = min(seen.values(), key=lambda x: (x["priority"], -x["amount"]))
    others = [r for r in seen.values() if r["amount"] != best["amount"]]
    return [best] + sorted(others, key=lambda x: (x["priority"], -x["amount"]))


# ---------------------------------------------------------
# Invoice ID extraction
# ---------------------------------------------------------

def extract_invoice_ids(text: str) -> set:
    ids = set()
    for pat in [
        r"invoice\s*(?:no\.?|#|number)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9/\-]{3,30})",
        r"(?:bill|ref|receipt|voucher)\s*(?:no\.?|#)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9/\-]{3,30})",
        r"\b([A-Z]{2,6}/\d{2}-\d{2}/\d{4,6})\b",
        r"\b([A-Z]{2,6}-\d{4,8})\b",
    ]:
        for m in re.findall(pat, text, re.IGNORECASE):
            ids.add(m.strip().upper())
    return ids


# ---------------------------------------------------------
# Vendor detection
# ---------------------------------------------------------

def detect_vendor(text: str):
    for kw in ["uber","ola","rapido","zomato","swiggy","restaurant","cafe",
               "amazon","flipkart","airtel","jio","bsnl","vodafone","hotel","lodge"]:
        if kw in text.lower():
            return kw
    for line in [l.strip() for l in text.split("\n") if len(l.strip()) > 4][:5]:
        if any(w in line.lower() for w in ["ltd","llp","pvt","inc","co.","corp"]):
            return line[:60]
    return None


# ---------------------------------------------------------
# File processors
# ---------------------------------------------------------

def process_image(fb: bytes, name: str) -> dict:
    img  = preprocess(Image.open(io.BytesIO(fb)).convert("RGB"))
    text = ocr_image(img)
    amts = extract_amounts_with_context(text)
    return {"name": name, "amounts_detail": amts,
            "amounts": [a["amount"] for a in amts],
            "invoice_ids": extract_invoice_ids(text),
            "vendor": detect_vendor(text), "text": text}

def process_pdf(fb: bytes, name: str) -> dict:
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(fb)) as pdf:
            for pg in pdf.pages[:5]:
                text += (pg.extract_text() or "") + "\n"
    except Exception:
        pass
    amts = extract_amounts_with_context(text)
    return {"name": name, "amounts_detail": amts,
            "amounts": [a["amount"] for a in amts],
            "invoice_ids": extract_invoice_ids(text),
            "vendor": detect_vendor(text), "text": text}

def process_file(fb: bytes, name: str) -> dict:
    return process_pdf(fb, name) if name.lower().endswith(".pdf") else process_image(fb, name)


# ---------------------------------------------------------
# Column normalisation
#
# 5 canonical columns:
#   REQUIRED : ExpenseReport ID, Expense Amount, Vendor
#   OPTIONAL : Invoice Date, Category
#
# Accepts common legacy / alias names so old files still work.
# ---------------------------------------------------------

COL_ALIASES = {
    "ExpenseReport ID": [
        "expensereport id", "expense report id", "report id",
        "invoice id", "invoice no", "invoice number", "ref",
        "reference", "voucher id", "bill no", "bill number",
    ],
    "Invoice Date": [
        "invoice date", "date", "bill date", "voucher date",
        "transaction date", "exp date", "expense date",
        "start date",          # legacy alias
    ],
    "Vendor": [
        "vendor", "vendor name", "supplier", "supplier name",
        "party", "party name", "payee",
    ],
    "Category": [
        "category", "expense category", "expense type",
        "type", "head", "ledger head",
    ],
    "Expense Amount": [
        "expense amount", "amount", "invoice amount",
        "total amount", "net amount", "payable amount", "value",
    ],
}

REQUIRED_COLS = ["ExpenseReport ID", "Expense Amount", "Vendor"]
OPTIONAL_COLS = ["Invoice Date", "Category"]
ALL_5_COLS    = ["ExpenseReport ID", "Invoice Date", "Vendor", "Category", "Expense Amount"]


def normalise_columns(df: pd.DataFrame):
    """
    Rename uploaded Excel columns to canonical names.
    Returns (normalised_df, missing_required_list, warning_list).
    """
    rename_map = {}
    col_lower  = {c.lower().strip(): c for c in df.columns}
    warnings   = []

    for canonical, aliases in COL_ALIASES.items():
        if canonical in df.columns:
            continue
        for alias in aliases:
            if alias.lower() in col_lower:
                original = col_lower[alias.lower()]
                rename_map[original] = canonical
                if original != canonical:
                    warnings.append(f'Column "{original}" recognised as "{canonical}"')
                break

    df = df.rename(columns=rename_map)

    if "Invoice Date" not in df.columns:
        df["Invoice Date"] = ""
        warnings.append('"Invoice Date" not found — left blank.')
    if "Category" not in df.columns:
        df["Category"] = "Uncategorised"
        warnings.append('"Category" not found — defaulted to "Uncategorised".')

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    return df, missing, warnings


# ---------------------------------------------------------
# Matching engine v5 — strict confidence gates
#
# Gate A (ID match)      : Report ID found verbatim in doc  → always accept
# Gate B (Amt + Vendor)  : Exact amount AND vendor score ≥3 → score ≥ 13
# Gate C (Amt only)      : Exact amount alone               → score ≥ 10
#
# Vendor-text-only matches (score < 10) → REJECTED → MISSING_DOC + ₹0.00
# ---------------------------------------------------------

TOLERANCE        = 0.50
_GATE_AMT_VENDOR = 13
_GATE_AMT_ONLY   = 10


def best_amount(doc: dict):
    if not doc["amounts_detail"]:
        return None
    ranked = sorted(doc["amounts_detail"], key=lambda x: (x["priority"], -x["amount"]))
    return ranked[0]["amount"]


def amount_result(reg_amt, doc: dict):
    try:
        ra = float(reg_amt)
        if np.isnan(ra) or np.isinf(ra):
            return "not_found", None, None
    except Exception:
        return "not_found", None, None
    da   = best_amount(doc)
    if da is None:
        return "not_found", None, None
    diff = abs(ra - da)
    return ("exact" if diff <= TOLERANCE else "mismatch"), da, diff


def id_match(report_id, doc: dict) -> bool:
    rid = str(report_id).strip().upper()
    if not rid:
        return False
    if rid in doc["text"].upper():
        return True
    return any(rid == d or rid in d or d in rid for d in doc["invoice_ids"])


def vendor_score(reg_vendor, doc: dict) -> int:
    rv   = str(reg_vendor).lower().strip()
    dv   = str(doc["vendor"] or "").lower()
    text = doc["text"].lower()
    if not rv:
        return 0
    tokens = set(rv.replace("&", " ").replace(".", " ").split())
    score  = sum(1 for t in tokens if len(t) > 2 and t in text)
    if dv and (rv in dv or dv in rv):
        score += 3
    return score


def _passes_gate(c: dict) -> bool:
    sc, idm, ast = c["score"], c["idm"], c["ast"]
    if idm:
        return True
    if ast == "exact" and sc >= _GATE_AMT_VENDOR:
        return True
    if ast == "exact" and sc >= _GATE_AMT_ONLY:
        return True
    return False


def run_vouching(df: pd.DataFrame, docs: list) -> list:
    results, used = [], set()

    for _, row in df.iterrows():
        report_id = row["ExpenseReport ID"]
        amount    = row["Expense Amount"]
        vendor    = row["Vendor"]
        category  = row.get("Category", "Uncategorised")
        inv_date  = row.get("Invoice Date", "")

        # Format invoice date
        if pd.notna(inv_date) and str(inv_date).strip() not in ("", "nan", "NaT"):
            try:
                inv_date_str = pd.to_datetime(inv_date).strftime("%d %b %Y")
            except Exception:
                inv_date_str = str(inv_date).strip()
        else:
            inv_date_str = "-"

        # Score every uploaded document
        candidates = []
        for d in docs:
            idm           = id_match(report_id, d)
            vs            = vendor_score(vendor, d)
            ast, da, diff = amount_result(amount, d)
            sc = (20 if idm else 0) + min(vs, 5) + (10 if ast == "exact" else 3 if ast == "mismatch" else 0)
            candidates.append({"doc": d, "score": sc,
                                "idm": idm, "ast": ast, "da": da, "diff": diff})

        candidates.sort(key=lambda x: -x["score"])
        qualified = [c for c in candidates if _passes_gate(c)]
        best      = qualified[0] if qualified else None

        try:
            av = float(amount)
            av = 0.0 if (np.isnan(av) or np.isinf(av)) else av
        except Exception:
            av = 0.0

        if best is None:
            status   = "MISSING_DOC"
            mf       = "-"
            da_str   = "₹0.00"
            diff_str = f"₹{av:,.2f}"
            conf     = "0/35"
        else:
            mf       = best["doc"]["name"]
            da_str   = f"₹{best['da']:,.2f}" if best["da"]   is not None else "₹0.00"
            diff_str = f"₹{best['diff']:,.2f}" if best["diff"] is not None else "-"
            conf     = f"{best['score']}/35"

            if mf in used:
                status = "DUPLICATE_RECEIPT"
            elif best["ast"] == "exact":
                status = "MATCHED"
            elif best["ast"] == "mismatch":
                status = "AMOUNT_MISMATCH"
            else:
                status = "MATCHED" if best["idm"] else "MISSING_DOC"

            if status != "DUPLICATE_RECEIPT":
                used.add(mf)

        results.append({
            "Report ID":       str(report_id),
            "Invoice Date":    inv_date_str,
            "Vendor":          str(vendor),
            "Category":        str(category),
            "Register Amount": av,
            "Document Amount": da_str,
            "Difference":      diff_str,
            "Matched File":    mf,
            "ID Match":        "✓" if (best and best["idm"]) else "✗",
            "Confidence":      conf,
            "Status":          status,
        })

    return results


# ---------------------------------------------------------
# Excel export
# ---------------------------------------------------------

def export_excel(df: pd.DataFrame) -> bytes:
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_numeric(df[col], errors="coerce") \
                        .replace([np.inf, -np.inf], np.nan).fillna(0)
        else:
            df[col] = df[col].fillna("").astype(str)

    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"nan_inf_to_errors": True, "in_memory": True})
    G   = "#5ba632"

    hdr  = wb.add_format({"bold":True,"bg_color":G,"font_color":"#fff",
                           "font_name":"Calibri","font_size":11,
                           "valign":"vcenter","align":"center"})
    base = wb.add_format({"font_name":"Calibri","font_size":10,"valign":"vcenter"})
    numf = wb.add_format({"font_name":"Calibri","font_size":10,"valign":"vcenter",
                           "num_format":"₹#,##0.00"})
    ok_f = wb.add_format({"font_name":"Calibri","font_size":10,
                           "font_color":G,        "bold":True,"valign":"vcenter"})
    mi_f = wb.add_format({"font_name":"Calibri","font_size":10,
                           "font_color":"#d13030","bold":True,"valign":"vcenter"})
    mm_f = wb.add_format({"font_name":"Calibri","font_size":10,
                           "font_color":"#e07b00","bold":True,"valign":"vcenter"})
    du_f = wb.add_format({"font_name":"Calibri","font_size":10,
                           "font_color":"#b8860b","bold":True,"valign":"vcenter"})

    ws = wb.add_worksheet("Vouching Results")
    ws.set_tab_color(G); ws.set_row(0, 24); ws.hide_gridlines(2)

    col_widths = {
        "Report ID": 22, "Invoice Date": 14, "Vendor": 28,
        "Category": 22, "Register Amount": 18, "Document Amount": 18,
        "Difference": 14, "Matched File": 36,
        "ID Match": 10, "Confidence": 12, "Status": 22,
    }
    for ci, col in enumerate(df.columns):
        ws.set_column(ci, ci, col_widths.get(col, 16))
        ws.write(0, ci, col, hdr)

    sc_i = list(df.columns).index("Status")          if "Status"          in df.columns else -1
    ra_i = list(df.columns).index("Register Amount") if "Register Amount" in df.columns else -1
    sfmt_map = {"MATCHED": ok_f, "MISSING_DOC": mi_f,
                "AMOUNT_MISMATCH": mm_f, "DUPLICATE_RECEIPT": du_f}

    for ri in range(len(df)):
        sfmt = sfmt_map.get(str(df.iloc[ri].get("Status", "")), base)
        for ci in range(len(df.columns)):
            raw = df.iloc[ri, ci]
            fmt = sfmt if ci == sc_i else (numf if ci == ra_i else base)
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
    sw.set_tab_color("#a8d870")
    sw.set_column(0, 0, 30); sw.set_column(1, 1, 16); sw.hide_gridlines(2)

    tf = wb.add_format({"font_name":"Calibri","font_size":15,
                         "bold":True,"font_color":"#1c2b1c"})
    lf = wb.add_format({"font_name":"Calibri","font_size":11,"font_color":"#7a8e7a"})
    vf = wb.add_format({"font_name":"Calibri","font_size":11,
                         "bold":True,"font_color":"#1c2b1c"})
    gf = wb.add_format({"font_name":"Calibri","font_size":11,
                         "bold":True,"font_color":G})

    sw.write(0, 0, "KKC Vouching Report — Summary", tf)
    sw.write(1, 0, f"KKC & Associates LLP  ·  {datetime.now().strftime('%d %b %Y  %H:%M')}", lf)

    total    = len(df)
    matched  = int((df["Status"] == "MATCHED").sum())
    missing  = int((df["Status"] == "MISSING_DOC").sum())
    mismatch = int((df["Status"] == "AMOUNT_MISMATCH").sum())
    dup      = int((df["Status"] == "DUPLICATE_RECEIPT").sum())
    rate     = round(matched / total * 100, 1) if total else 0.0

    for i, (lbl, val, fmt) in enumerate([
        ("Total Entries",      total,    vf),
        ("Matched",            matched,  gf),
        ("Amount Mismatches",  mismatch, vf),
        ("Missing Documents",  missing,  vf),
        ("Duplicate Receipts", dup,      vf),
        ("Match Rate (%)",     rate,     gf),
    ]):
        sw.write(i + 3, 0, lbl, lf)
        sw.write_number(i + 3, 1, float(val), fmt)

    wb.close()
    buf.seek(0)
    return buf.read()


# =========================================================
# UI LAYOUT
# =========================================================

# ── Top header ───────────────────────────────────────────
st.markdown("""
<div class="kkc-topstrip">
    <span>📞 +91 22 6143 7333</span>
    <span style="opacity:0.4">|</span>
    <span>✉ info@kkcllp.in</span>
    <span style="opacity:0.4">|</span>
    <span>Vouching Tool — Internal Audit Suite</span>
</div>
<div class="kkc-navbar">
    <div>
        <div class="kkc-brand-name">kkc &amp; associates llp</div>
        <div class="kkc-brand-tag">Chartered Accountants</div>
        <div class="kkc-brand-fmr">(Formerly Khimji Kunverji &amp; Co LLP)</div>
    </div>
    <div class="kkc-divider-v"></div>
    <div>
        <div class="kkc-tool-badge">🔍 Vouching Engine</div>
        <div class="kkc-tool-sub">Expense Audit &amp; Receipt Matching</div>
    </div>
    <div class="kkc-pill"><span class="kkc-dot"></span>Session Active</div>
</div>
""", unsafe_allow_html=True)

# ── Info box ─────────────────────────────────────────────
st.markdown("""
<div class="kkc-infobox">
    <span style="font-size:16px;flex-shrink:0">ℹ</span>
    <span>Upload an <strong>Expense Register</strong> (.xlsx) with the 5 standard columns
    and one or more <strong>Receipts / Invoices</strong> (PDF, PNG, JPG).
    Matching priority: <strong>Invoice ID → Vendor → Amount</strong>.
    If no invoice is uploaded for a vendor, the entry is marked
    <strong>MISSING_DOC</strong> with Document Amount = ₹0.00.</span>
</div>
""", unsafe_allow_html=True)

# ── Required columns reference card ──────────────────────
st.markdown("""
<div class="kkc-colhint">
    <div class="kkc-colhint-title">📋 Required Excel Column Format (5 Columns)</div>
    <div class="kkc-col-grid">
        <div class="kkc-col-item">
            <div class="kkc-col-num">COL 1</div>
            <div class="kkc-col-name">ExpenseReport ID</div>
            <div class="kkc-col-desc">Invoice / report number<br>e.g. JJCO/25-26/0005</div>
            <div class="kkc-col-req">● Required</div>
        </div>
        <div class="kkc-col-item">
            <div class="kkc-col-num">COL 2</div>
            <div class="kkc-col-name">Invoice Date</div>
            <div class="kkc-col-desc">Date of the invoice<br>e.g. 15 Sep 2025</div>
            <div class="kkc-col-opt">○ Optional</div>
        </div>
        <div class="kkc-col-item">
            <div class="kkc-col-num">COL 3</div>
            <div class="kkc-col-name">Vendor</div>
            <div class="kkc-col-desc">Supplier / party name<br>e.g. KKC & Associates</div>
            <div class="kkc-col-req">● Required</div>
        </div>
        <div class="kkc-col-item">
            <div class="kkc-col-num">COL 4</div>
            <div class="kkc-col-name">Category</div>
            <div class="kkc-col-desc">Expense type / head<br>e.g. Professional Fees</div>
            <div class="kkc-col-opt">○ Optional</div>
        </div>
        <div class="kkc-col-item">
            <div class="kkc-col-num">COL 5</div>
            <div class="kkc-col-name">Expense Amount</div>
            <div class="kkc-col-desc">Amount to be vouched<br>e.g. 17700</div>
            <div class="kkc-col-req">● Required</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────
st.markdown('<div class="kkc-section">Document Intake</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2, gap="medium")

with col1:
    st.markdown(
        '<div class="kkc-upload-label">Expense Register '
        '<span class="kkc-badge">XLSX</span></div>',
        unsafe_allow_html=True)
    register = st.file_uploader(
        "Expense Register", type=["xlsx"],
        label_visibility="collapsed", key="reg")
    if register:
        st.markdown(
            f'<div class="kkc-file-ok">✓ &nbsp;{register.name}</div>',
            unsafe_allow_html=True)

with col2:
    st.markdown(
        '<div class="kkc-upload-label">Receipts &amp; Invoices '
        '<span class="kkc-badge">PDF / PNG / JPG</span></div>',
        unsafe_allow_html=True)
    docs = st.file_uploader(
        "Receipts", accept_multiple_files=True,
        label_visibility="collapsed", key="docs")
    if docs:
        st.markdown(
            f'<div class="kkc-file-ok">✓ &nbsp;{len(docs)} '
            f'file{"s" if len(docs) != 1 else ""} ready</div>',
            unsafe_allow_html=True)


# ── Main flow ─────────────────────────────────────────────
if register and docs:

    # Read and normalise to 5 columns
    raw_df = pd.read_excel(register)
    df, missing_cols, col_warnings = normalise_columns(raw_df)

    # Show column mapping info
    for w in col_warnings:
        st.markdown(
            f'<div style="font-size:11px;color:#5ba632;background:#e8f5e0;'
            f'border:1px solid #5ba632;border-radius:6px;padding:5px 12px;'
            f'margin-bottom:4px">ℹ {w}</div>',
            unsafe_allow_html=True)

    # Block on missing required columns
    if missing_cols:
        st.markdown(
            f'<div class="kkc-file-error">✗ Missing required columns: '
            f'<strong>{", ".join(missing_cols)}</strong>. '
            f'Please check the column format guide above.</div>',
            unsafe_allow_html=True)
        st.stop()

    # Register preview — show only the 5 working columns
    st.markdown('<div class="kkc-section">Register Preview</div>', unsafe_allow_html=True)
    preview_cols = [c for c in ALL_5_COLS if c in df.columns]
    st.dataframe(df[preview_cols].head(8), use_container_width=True, hide_index=True)
    st.markdown(
        f'<div class="kkc-preview-meta">'
        f'{len(df)} entries · {len(preview_cols)} columns</div>',
        unsafe_allow_html=True)

    # Process uploaded documents
    st.markdown('<div class="kkc-section">Processing Documents</div>', unsafe_allow_html=True)
    files_data, prog_text, prog_bar = [], st.empty(), st.progress(0)

    for i, f in enumerate(docs):
        prog_text.markdown(
            f'<div class="kkc-proc-label">Extracting → '
            f'<strong>{f.name}</strong> [{i+1}/{len(docs)}]</div>',
            unsafe_allow_html=True)
        files_data.append(process_file(f.read(), f.name))
        prog_bar.progress((i + 1) / len(docs))

    prog_text.markdown(
        '<div class="kkc-proc-done">✓ All documents processed successfully</div>',
        unsafe_allow_html=True)
    st.markdown('<div class="kkc-divider"></div>', unsafe_allow_html=True)

    bc, _ = st.columns([1, 4])
    with bc:
        run = st.button("🔍  Run Vouching Analysis")

    if run:
        with st.spinner("Matching expense entries to documents…"):
            results = run_vouching(df, files_data)
            rdf     = pd.DataFrame(results)

        nm  = int((rdf.Status == "MATCHED").sum())
        nmi = int((rdf.Status == "MISSING_DOC").sum())
        nmm = int((rdf.Status == "AMOUNT_MISMATCH").sum())
        nd  = int((rdf.Status == "DUPLICATE_RECEIPT").sum())
        tot = len(rdf)
        rt  = round(nm / tot * 100) if tot else 0

        # Summary cards
        st.markdown('<div class="kkc-section">Summary</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="kkc-metrics">
            <div class="kkc-metric matched">
                <div class="kkc-metric-label">Matched</div>
                <div class="kkc-metric-value">{nm}</div>
                <div class="kkc-metric-sub">{rt}% match rate</div>
            </div>
            <div class="kkc-metric mismatch">
                <div class="kkc-metric-label">Amount Mismatch</div>
                <div class="kkc-metric-value">{nmm}</div>
                <div class="kkc-metric-sub">Doc found, amount differs</div>
            </div>
            <div class="kkc-metric missing">
                <div class="kkc-metric-label">Missing Documents</div>
                <div class="kkc-metric-value">{nmi}</div>
                <div class="kkc-metric-sub">Invoice not uploaded</div>
            </div>
            <div class="kkc-metric duplicate">
                <div class="kkc-metric-label">Duplicates</div>
                <div class="kkc-metric-value">{nd}</div>
                <div class="kkc-metric-sub">Reused receipts</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Amount mismatch alert
        if nmm > 0:
            rows  = rdf[rdf["Status"] == "AMOUNT_MISMATCH"]
            trows = "".join(
                f"<tr>"
                f"<td>{r['Report ID']}</td>"
                f"<td>{r['Invoice Date']}</td>"
                f"<td>{r['Vendor']}</td>"
                f"<td style='color:#d13030;font-weight:700'>"
                f"₹{r['Register Amount']:,.2f}</td>"
                f"<td style='color:#5ba632;font-weight:700'>"
                f"{r['Document Amount']}</td>"
                f"<td style='color:#e07b00;font-weight:700'>"
                f"{r['Difference']}</td>"
                f"<td>{r['Matched File']}</td>"
                f"</tr>"
                for _, r in rows.iterrows()
            )
            st.markdown(f"""
            <div class="kkc-alert-warn">
                <strong>⚠ {nmm} amount mismatch{"es" if nmm > 1 else ""}
                require manual review</strong> — a supporting document was
                identified but amounts do not agree.
                <table class="kkc-mismatch-table">
                    <thead><tr>
                        <th>Report ID</th><th>Invoice Date</th><th>Vendor</th>
                        <th>Register Amount</th><th>Document Amount</th>
                        <th>Difference</th><th>Matched File</th>
                    </tr></thead>
                    <tbody>{trows}</tbody>
                </table>
            </div>
            """, unsafe_allow_html=True)

        # Detailed results
        st.markdown('<div class="kkc-section">Detailed Results</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="kkc-result-head">Vouching Report '
            f'<span class="kkc-count-tag">{tot} entries</span></div>',
            unsafe_allow_html=True)

        def style_status(val):
            return {
                "MATCHED":           "color:#5ba632;font-weight:800",
                "MISSING_DOC":       "color:#d13030;font-weight:800",
                "AMOUNT_MISMATCH":   "color:#e07b00;font-weight:800",
                "DUPLICATE_RECEIPT": "color:#b8860b;font-weight:800",
            }.get(val, "")

        st.dataframe(
            rdf.style.applymap(style_status, subset=["Status"]),
            use_container_width=True, hide_index=True)

        st.markdown('<div class="kkc-divider"></div>', unsafe_allow_html=True)
        dc, tc = st.columns([1, 3])
        with dc:
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                label="⬇  Download Audit Report (.xlsx)",
                data=export_excel(rdf),
                file_name=f"KKC_vouching_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with tc:
            st.markdown(
                f'<div class="kkc-timestamp">KKC &amp; Associates LLP &nbsp;·&nbsp; '
                f'Generated {datetime.now().strftime("%d %b %Y · %H:%M")} '
                f'&nbsp;·&nbsp; {tot} entries &nbsp;·&nbsp; '
                f'2 sheets: Results + Summary</div>',
                unsafe_allow_html=True)

else:
    st.markdown("""
    <div class="kkc-empty">
        <div class="kkc-empty-icon">🔍</div>
        <div class="kkc-empty-title">Ready to begin vouching</div>
        <div class="kkc-empty-sub">Upload your expense register and receipts above to start</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
<div class="kkc-footer">
    <strong style="color:#5ba632">kkc &amp; associates llp</strong> &nbsp;·&nbsp;
    Chartered Accountants &nbsp;·&nbsp; (Formerly Khimji Kunverji &amp; Co LLP)
    &nbsp;·&nbsp; Internal Audit Suite v5.0
</div>
""", unsafe_allow_html=True)
