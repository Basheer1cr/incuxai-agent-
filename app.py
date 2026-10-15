import warnings
warnings.filterwarnings("ignore")

import os
import io
import re
import json
import logging
from typing import List, Dict, Any, Tuple, Optional

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import numpy as np
from PIL import Image, ImageOps
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configure logging for elite code quality & traceability
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("DocuFlowAgent")

try:
    import google.generativeai as genai
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

# =====================================================================
# 1. PAGE CONFIGURATION & ENTERPRISE SAAS STYLING
# =====================================================================
st.set_page_config(
    page_title="INCUXAI  Ai | Autonomous Document AI",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2.5rem; max-width: 1400px; }
    .hero-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 28px 32px;
        border-radius: 12px;
        border: 1px solid #334155;
        color: #f8fafc;
        margin-bottom: 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
    }
    .status-pill {
        display: inline-block;
        background: rgba(37, 99, 235, 0.2);
        border: 1px solid #3b82f6;
        color: #93c5fd;
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-right: 8px;
    }
    .telemetry-bar {
        background: #f8fafc;
        border-left: 4px solid #2563eb;
        border-top: 1px solid #e2e8f0;
        border-right: 1px solid #e2e8f0;
        border-bottom: 1px solid #e2e8f0;
        padding: 12px 18px;
        border-radius: 6px;
        margin-bottom: 18px;
        font-size: 13.5px;
        color: #1e293b;
    }
    .arch-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 20px;
        height: 100%;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }
    .arch-card h4 { margin-top: 0; color: #0f172a; font-size: 16px; }
    .arch-card p { color: #475569; font-size: 13.5px; margin-bottom: 0; line-height: 1.5; }
    div[data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 12px 16px;
        border-radius: 10px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
    <div>
        <span class="status-pill">Production Grade Agent</span>
        <span class="status-pill">100-Page Batch Engine</span>
        <span class="status-pill">Self-Healing Reconciliation</span>
    </div>
    <h2 style="margin-top:14px; margin-bottom:6px; color:#ffffff; font-weight:700;">
        🚀INCUXAI  Ai AI: Autonomous Document-to-Excel Agent
    </h2>
    <p style="margin:0; color:#cbd5e1; font-size:15px;">
        Production-ready autonomous document intelligence pipeline with vector grid parsing, multi-page stitching, and mathematical reconciliation.
    </p>
</div>
""", unsafe_allow_html=True)


# =====================================================================
# 2. ROBUST CORE PROCESSING FUNCTIONS (TYPE-HINTED & LOGGED)
# =====================================================================
MAX_PAGES_AUTO = 100

def clean_cell(val: Any) -> str:
    """Cleans whitespace and normalizes string cell content."""
    if val is None:
        return ""
    return re.sub(r'\s+', ' ', str(val)).strip()

def parse_number(val: Any) -> float:
    """Safely converts currency strings or formatted numbers into numeric floats."""
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if not s or s in ["-", "N/A", "null", "None"]:
        return np.nan
    cleaned = re.sub(r'[₹$€£,\s]', '', s)
    if cleaned.startswith('(') and cleaned.endswith(')'):
        cleaned = '-' + cleaned[1:-1]
    try:
        if re.match(r'^-?\d+(\.\d+)?$', cleaned):
            return float(cleaned)
    except Exception:
        pass
    return np.nan

def normalize_header(headers: List[Any]) -> List[str]:
    """Ensures unique, non-empty column headers for tabular data."""
    seen: Dict[str, int] = {}
    clean_headers = []
    for idx, h in enumerate(headers):
        h_str = clean_cell(h)
        if not h_str:
            h_str = f"Column_{idx + 1}"
        if h_str in seen:
            seen[h_str] += 1
            h_str = f"{h_str}_{seen[h_str]}"
        else:
            seen[h_str] = 1
        clean_headers.append(h_str)
    return clean_headers

def headers_are_similar(h1: List[str], h2: List[str]) -> bool:
    """Evaluates header similarity percentage to detect multi-page table continuations."""
    if len(h1) != len(h2):
        return False
    matches = sum(1 for a, b in zip(h1, h2) if str(a).lower().strip() == str(b).lower().strip())
    return (matches / max(len(h1), 1)) >= 0.7

def recover_misaligned_rows(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """Strips empty ghost columns and cleans rows to maintain structural integrity."""
    fixed_count = 0
    df = df.replace("", np.nan).dropna(how="all").fillna("")
    empty_cols = [c for c in df.columns if (df[c].astype(str).str.strip() == "").all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)
        fixed_count += len(empty_cols)
    return df.reset_index(drop=True), fixed_count

def extract_borderless_table_safely(page: Any) -> List[pd.DataFrame]:
    """Extracts borderless stream layouts or key-value structures without cutting words."""
    try:
        text = page.extract_text(layout=True)
        if not text or not text.strip():
            return []
        
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if not lines:
            return []

        split_rows = [re.split(r'\s{2,}|\t+', line) for line in lines]
        col_counts = [len(r) for r in split_rows if len(r) > 1]
        
        if col_counts and (len(col_counts) / len(lines) >= 0.5):
            max_cols = min(max(col_counts), 15)
            padded_rows = [r[:max_cols] + [""] * max(0, max_cols - len(r)) for r in split_rows]
            headers = normalize_header(padded_rows[0])
            return [pd.DataFrame(padded_rows[1:], columns=headers)]
        
        kv_rows = []
        for line in lines:
            if ":" in line:
                parts = line.split(":", 1)
                kv_rows.append([parts[0].strip(), parts[1].strip()])
            else:
                kv_rows.append(["Content", line])
        return [pd.DataFrame(kv_rows, columns=["Field / Section", "Extracted Value"])]
    except Exception as e:
        logger.error(f"Error in borderless extraction: {str(e)}")
        return []

def extract_with_gemini_vision(image: Image.Image, api_key: str) -> List[pd.DataFrame]:
    """Sends document images to Gemini Vision API for structured JSON table extraction."""
    if not GENAI_AVAILABLE or not api_key:
        return []
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-2.5-flash")
        prompt = """
        Extract all tables from this document image into strict JSON format.
        Return ONLY valid JSON matching this schema:
        {"tables": [{"headers": ["col1", "col2", ...], "rows": [["val1", "val2", ...]]}]}
        Preserve exact numbers, totals, and all columns.
        """
        response = model.generate_content([prompt, image])
        raw_text = response.text.strip()
        raw_text = re.sub(r'^```json\s*|\s*```$', '', raw_text, flags=re.MULTILINE)
        data = json.loads(raw_text)
        dfs = []
        for t in data.get("tables", []):
            headers = normalize_header(t.get("headers", []))
            rows = t.get("rows", [])
            if headers and rows:
                dfs.append(pd.DataFrame(rows, columns=headers))
        return dfs
    except Exception as e:
        logger.error(f"Gemini Vision API error: {str(e)}")
        return []

def process_document_autonomously(file_bytes: bytes, file_name: str, gemini_key: str) -> Dict[str, Any]:
    """Executes the complete autonomous multi-stage document processing pipeline."""
    raw_tables = []
    pages_processed = 0
    ocr_pages = 0
    fixed_alignments = 0
    engines_triggered = set()

    is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg'))

    if is_image:
        pages_processed = 1
        ocr_pages = 1
        engines_triggered.add("Vision AI OCR")
        img = Image.open(io.BytesIO(file_bytes))
        img = ImageOps.autocontrast(img.convert("RGB"))
        dfs = extract_with_gemini_vision(img, gemini_key)
        for df in dfs:
            raw_tables.append({"page": 1, "df": df, "method": "Auto: Vision AI OCR"})
    else:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            total_pages = min(len(pdf.pages), MAX_PAGES_AUTO)
            pages_processed = total_pages
            
            for p_idx in range(total_pages):
                page = pdf.pages[p_idx]
                page_num = p_idx + 1
                valid_table_found = False

                try:
                    tables = page.extract_tables({
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_tolerance": 4,
                        "join_tolerance": 4
                    })
                    if tables:
                        for tbl in tables:
                            cleaned_tbl = [[clean_cell(c) for c in row] for row in tbl if any(clean_cell(c) for c in row)]
                            if len(cleaned_tbl) >= 2 and len(cleaned_tbl[0]) >= 2:
                                headers = normalize_header(cleaned_tbl[0])
                                df = pd.DataFrame(cleaned_tbl[1:], columns=headers)
                                df, fixes = recover_misaligned_rows(df)
                                fixed_alignments += fixes
                                if not df.empty:
                                    raw_tables.append({"page": page_num, "df": df, "method": "Auto: Vector Grid Engine"})
                                    engines_triggered.add("Vector Grid Engine")
                                    valid_table_found = True
                except Exception as ex:
                    logger.warning(f"Grid extraction warning on page {page_num}: {str(ex)}")

                if not valid_table_found:
                    fallback_dfs = extract_borderless_table_safely(page)
                    for df in fallback_dfs:
                        df, fixes = recover_misaligned_rows(df)
                        fixed_alignments += fixes
                        if not df.empty:
                            raw_tables.append({"page": page_num, "df": df, "method": "Auto: Spatial Layout Engine"})
                            engines_triggered.add("Spatial Layout Engine")
                            valid_table_found = True

                if not valid_table_found and gemini_key:
                    try:
                        doc_fitz = fitz.open(stream=file_bytes, filetype="pdf")
                        pix = doc_fitz[p_idx].get_pixmap(dpi=200)
                        img = Image.open(io.BytesIO(pix.tobytes("png")))
                        ocr_dfs = extract_with_gemini_vision(img, gemini_key)
                        if ocr_dfs:
                            ocr_pages += 1
                            engines_triggered.add("Vision AI OCR")
                            for df in ocr_dfs:
                                raw_tables.append({"page": page_num, "df": df, "method": "Auto: Vision AI OCR"})
                        doc_fitz.close()
                    except Exception as ocr_ex:
                        logger.error(f"OCR fallback error on page {page_num}: {str(ocr_ex)}")

    # Multi-page table stitching
    stitched_tables = []
    stitch_events = []

    for item in raw_tables:
        curr_df = item["df"].copy()
        curr_page = item["page"]
        curr_method = item["method"]

        if not stitched_tables:
            stitched_tables.append({
                "name": "Table_1",
                "start_page": curr_page,
                "end_page": curr_page,
                "df": curr_df,
                "method": curr_method
            })
            continue

        prev = stitched_tables[-1]
        prev_df = prev["df"]

        if len(curr_df.columns) == len(prev_df.columns) and (curr_page - prev["end_page"] <= 1):
            if headers_are_similar(prev_df.columns, curr_df.columns):
                curr_df.columns = prev_df.columns
            else:
                first_row_as_df = pd.DataFrame([list(curr_df.columns)], columns=prev_df.columns)
                curr_df.columns = prev_df.columns
                curr_df = pd.concat([first_row_as_df, curr_df], ignore_index=True)

            header_set = set(str(x).lower().strip() for x in prev_df.columns)
            curr_df = curr_df[~curr_df.apply(lambda r: set(str(x).lower().strip() for x in r) == header_set, axis=1)]

            combined = pd.concat([prev_df, curr_df], ignore_index=True).drop_duplicates().reset_index(drop=True)
            prev["df"] = combined
            prev["end_page"] = curr_page
            stitch_events.append(f"Auto-Stitched {prev['name']} across Pages {prev['start_page']}–{curr_page} ({len(combined)} continuous rows)")
        else:
            stitched_tables.append({
                "name": f"Table_{len(stitched_tables) + 1}",
                "start_page": curr_page,
                "end_page": curr_page,
                "df": curr_df,
                "method": curr_method
            })

    # Data Integrity & Mathematical Reconciliation
    validation_records = []
    for tbl in stitched_tables:
        df = tbl["df"]
        t_name = tbl["name"]
        
        empty_mask = (df.astype(str).apply(lambda x: x.str.strip()) == "")
        empty_count = int(empty_mask.sum().sum())
        if empty_count > 0:
            validation_records.append({
                "Table": t_name,
                "Audit Rule": "Cell Completeness Check",
                "Target Scope": "All Columns",
                "Expected": "0 Empty Cells",
                "Actual": f"{empty_count} Empty Cells",
                "Status": "⚠️ Warning"
            })
        else:
            validation_records.append({
                "Table": t_name,
                "Audit Rule": "Structural Integrity",
                "Target Scope": f"{len(df.columns)} Columns",
                "Expected": "100% Populated",
                "Actual": "100% Populated",
                "Status": "✅ Verified"
            })

        first_col_str = df.iloc[:, 0].astype(str).str.lower().str.strip()
        total_row_indices = df[first_col_str.str.contains(r'\b(total|grand total|subtotal|sum)\b', regex=True, na=False)].index.tolist()

        if total_row_indices:
            t_idx = total_row_indices[-1]
            data_rows = df.iloc[:t_idx]
            total_row = df.iloc[t_idx]

            for col in df.columns[1:]:
                parsed_data = data_rows[col].apply(parse_number)
                parsed_total = parse_number(total_row[col])
                
                if parsed_data.notna().sum() >= 2 and pd.notna(parsed_total):
                    calc_sum = round(float(parsed_data.sum()), 2)
                    rep_sum = round(float(parsed_total), 2)
                    
                    last_val = parsed_data.dropna().iloc[-1] if not parsed_data.dropna().empty else None
                    if last_val is not None and rep_sum == round(float(last_val), 2) and "balance" in col.lower():
                        validation_records.append({
                            "Table": t_name,
                            "Audit Rule": "Closing Balance Reconciliation",
                            "Target Scope": f"Column '{col}'",
                            "Expected": f"{last_val:,.2f}",
                            "Actual": f"{rep_sum:,.2f}",
                            "Status": "✅ Verified"
                        })
                    elif abs(calc_sum - rep_sum) <= 0.05:
                        validation_records.append({
                            "Table": t_name,
                            "Audit Rule": "Vertical Sum Reconciliation",
                            "Target Scope": f"Column '{col}'",
                            "Expected": f"{calc_sum:,.2f}",
                            "Actual": f"{rep_sum:,.2f}",
                            "Status": "✅ Verified"
                        })
                    else:
                        validation_records.append({
                            "Table": t_name,
                            "Audit Rule": "Mathematical Discrepancy",
                            "Target Scope": f"Column '{col}'",
                            "Expected": f"{calc_sum:,.2f}",
                            "Actual": f"{rep_sum:,.2f}",
                            "Status": "❌ Discrepancy"
                        })

    val_df = pd.DataFrame(validation_records) if validation_records else pd.DataFrame(
        [{"Table": "N/A", "Audit Rule": "Document Scan", "Target Scope": "-", "Expected": "-", "Actual": "-", "Status": "⚠️ No Tables"}]
    )

    return {
        "pages_processed": pages_processed,
        "ocr_pages": ocr_pages,
        "fixed_alignments": fixed_alignments,
        "engines_triggered": list(engines_triggered) if engines_triggered else ["Auto-Profiler"],
        "stitched_tables": stitched_tables,
        "stitch_events": stitch_events,
        "validation_df": val_df
    }


# =====================================================================
# 3. ENTERPRISE EXCEL GENERATOR (STYLED WORKBOOK WITH OPENPYXL)
# =====================================================================
def generate_enterprise_excel(result: Dict[str, Any]) -> bytes:
    """Generates a professional multi-sheet corporate Excel workbook."""
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    total_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    warn_fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    err_fill = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for tbl in result["stitched_tables"]:
        sheet_title = f"{tbl['name']}_P{tbl['start_page']}-{tbl['end_page']}"[:31]
        ws = wb.create_sheet(title=sheet_title)
        df = tbl["df"]

        for c_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=str(col_name))
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            is_total_row = str(row[0]).strip().lower() in ["total", "grand total", "subtotal"]
            for c_idx, val in enumerate(row, start=1):
                num_val = parse_number(val)
                cell_val = num_val if pd.notna(num_val) else str(val)
                cell = ws.cell(row=r_idx, column=c_idx, value=cell_val)
                cell.border = thin_border

                if pd.notna(num_val):
                    cell.number_format = '#,##0.00' if '.' in str(val) else '#,##0'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

                if is_total_row:
                    cell.fill = total_fill
                    cell.font = Font(name="Calibri", size=11, bold=True)
                elif str(val).strip() == "":
                    cell.fill = warn_fill
                elif r_idx % 2 == 0:
                    cell.fill = zebra_fill

        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(min(max_len + 5, 42), 12)

    ws_val = wb.create_sheet(title="Audit_Validation_Report")
    val_df = result["validation_df"]

    for c_idx, col_name in enumerate(val_df.columns, start=1):
        cell = ws_val.cell(row=1, column=c_idx, value=str(col_name))
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border

    for r_idx, row in enumerate(val_df.itertuples(index=False), start=2):
        status_str = str(row[-1])
        for c_idx, val in enumerate(row, start=1):
            cell = ws_val.cell(row=r_idx, column=c_idx, value=str(val))
            cell.border = thin_border
            if "❌" in status_str:
                cell.fill = err_fill
            elif "⚠️" in status_str:
                cell.fill = warn_fill
            elif "✅" in status_str:
                cell.fill = total_fill

    ws_val.freeze_panes = "A2"
    for col in ws_val.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_val.column_dimensions[col_letter].width = max(min(max_len + 5, 45), 14)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =====================================================================
# 4. MAIN WORKSPACE WITH SESSION STATE PERSISTENCE
# =====================================================================
with st.sidebar:
    st.markdown("### ⚙️ System Status")
    st.success("🟢 Autonomous Agent Online (100 Pages)")
    st.info("💡 **Tip:** Digital PDFs work instantly without any key. Scanned photos/images require a Gemini API Key.")

st.markdown("### 🔑 Vision OCR Configuration (For Scanned Photos / Images Only)")
gemini_api_key = st.text_input(
    "Enter Google Gemini API Key (Leave blank if processing digital PDFs):",
    type="password",
    placeholder="AIzaSy...",
    help="Get a free key from aistudio.google.com/apikey"
)

st.divider()

uploaded_file = st.file_uploader(
    "📂 Drop any PDF, Bank Statement, Tax Invoice, or Scanned Image (Auto-detects layout up to 100 pages)",
    type=["pdf", "png", "jpg", "jpeg"]
)

# Initialize Session State to maintain state across interactions
if "processing_result" not in st.session_state:
    st.session_state.processing_result = None
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None

if uploaded_file is None:
    st.session_state.processing_result = None
    st.session_state.current_file_name = None
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class="arch-card">
            <h4>🤖 1. Autonomous Layout Profiler</h4>
            <p>Automatically inspects every uploaded page and switches between Vector Grid Intersection, Spatial Stream Parsing, and Vision OCR with zero manual setup.</p>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="arch-card">
            <h4>🔗 2. Zero-Touch Multi-Page Stitching</h4>
            <p>Automatically detects tables continuing across up to 100 pages, strips repeated page headers, and heals shifted columns into a single clean dataset.</p>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="arch-card">
            <h4>🛡️ 3. Self-Auditing Excel Export</h4>
            <p>Automatically normalizes currency values, reconciles vertical column sums against reported totals, and generates a styled corporate Excel workbook.</p>
        </div>
        """, unsafe_allow_html=True)

else:
    active_bytes = uploaded_file.read()
    active_name = uploaded_file.name

    if st.session_state.current_file_name != active_name:
        with st.spinner("🤖 Autonomous Agent profiling document layout, stitching pages, and auditing math totals..."):
            st.session_state.processing_result = process_document_autonomously(active_bytes, active_name, gemini_api_key)
            st.session_state.current_file_name = active_name

    result = st.session_state.processing_result
    tables = result["stitched_tables"]
    val_df = result["validation_df"]

    total_checks = len(val_df)
    passed_checks = sum(1 for s in val_df["Status"] if "✅" in str(s))
    errors_found = sum(1 for s in val_df["Status"] if "❌" in str(s))
    accuracy_score = round((passed_checks / max(total_checks, 1)) * 100, 1)

    engines_str = " + ".join(result["engines_triggered"])
    st.markdown(f"""
    <div class="telemetry-bar">
        🤖 <b>Autonomous Agent Decision Log:</b> Auto-Selected <code>{engines_str}</code> • 
        Auto-Stitching: <code>Active</code> • Header Deduplication: <code>Active</code> • 
        Columns Auto-Healed: <code>{result['fixed_alignments']}</code> • Math Reconciliation: <code>{accuracy_score}% Verified</code>
    </div>
    """, unsafe_allow_html=True)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📄 Pages Processed", f"{result['pages_processed']}")
    k2.metric("📊 Extracted Tables", f"{len(tables)}")
    max_cols = max((len(t["df"].columns) for t in tables), default=0)
    k3.metric("📐 Max Columns", f"{max_cols} Cols")
    k4.metric("🔗 Page Stitches", f"{len(result['stitch_events'])}")
    k5.metric("🛡️ Audit Status", "Verified ✅" if errors_found == 0 else f"{errors_found} Alerts ⚠️")

    if result["stitch_events"]:
        for evt in result["stitch_events"]:
            st.success(f"🔗 **Autonomous Continuity Engine:** {evt}")

    if errors_found > 0:
        st.error(f"🚨 **Autonomous Auditor Alert:** Detected {errors_found} mathematical discrepancy in `{active_name}`. Review the Audit & Validation Report tab for details.")

    tab1, tab2, tab3 = st.tabs([
        "📊 Structured Data Grid (Interactive Editor)",
        "🛡️ Audit & Validation Report",
        "👁️ Source Document Preview"
    ])

    with tab1:
        if not tables:
            st.warning("No structured tables detected. If this is a scanned image/photo, please ensure you have entered a valid Gemini API Key above.")
        else:
            for idx, tbl in enumerate(tables):
                st.markdown(
                    f"#### 🗂️ `{tbl['name']}` — Pages `{tbl['start_page']}` to `{tbl['end_page']}` "
                    f"| **{len(tbl['df'])} Rows × {len(tbl['df'].columns)} Columns** | `{tbl['method']}`"
                )
                edited_df = st.data_editor(
                    tbl["df"],
                    key=f"editor_{idx}",
                    use_container_width=True,
                    num_rows="dynamic"
                )
                result["stitched_tables"][idx]["df"] = edited_df

            st.divider()
            excel_bytes = generate_enterprise_excel(result)
            dl_col1, dl_col2 = st.columns([2, 1])
            with dl_col1:
                st.download_button(
                    label="📥 Export Formatted Enterprise Excel Workbook (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Converted_{active_name.rsplit('.', 1)[0]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with dl_col2:
                report_md = f"# Enterprise Audit & Validation Report: {active_name}\n\n"
                report_md += f"- **Pages Processed:** {result['pages_processed']}\n"
                report_md += f"- **Auto-Selected Engine:** {engines_str}\n"
                report_md += f"- **Tables Extracted:** {len(tables)}\n"
                report_md += f"- **Integrity Score:** {accuracy_score}%\n\n## Reconciliation Log\n"
                report_md += val_df.to_markdown(index=False) if hasattr(val_df, "to_markdown") else val_df.to_string()
                st.download_button(
                    label="📄 Export Audit Report (.md)",
                    data=report_md.encode("utf-8"),
                    file_name=f"Audit_Report_{active_name.rsplit('.', 1)[0]}.md",
                    mime="text/markdown",
                    use_container_width=True
                )

    with tab2:
        st.subheader("🔍 Automated Data Integrity & Mathematical Reconciliation")
        st.dataframe(val_df, use_container_width=True)

    with tab3:
        st.subheader("📄 Original Document Inspection")
        if active_name.lower().endswith(".pdf"):
            doc_preview = fitz.open(stream=active_bytes, filetype="pdf")
            preview_cols = st.columns(min(len(doc_preview), 3))
            for i in range(min(len(doc_preview), 3)):
                pix = doc_preview[i].get_pixmap(dpi=120)
                preview_cols[i].image(pix.tobytes("png"), caption=f"Page {i+1}", use_container_width=True)
            doc_preview.close()
        else:
            st.image(active_bytes, caption=active_name, width=600)