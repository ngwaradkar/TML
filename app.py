"""
TCF PPC Dashboard — Production Planning & Shortage Report
═════════════════════════════════════════════════════════════
Streamlit application for Tata Motors TCF line tracking.
Reads raw data uploads, computes shortage reports against
paint float pipeline, and presents 4 interactive summary tabs.
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import json
import re
from pathlib import Path
from io import BytesIO
from datetime import datetime

# ═══════════════════════════════════════════════════════════════
# 1. CONSTANTS & CONFIGURATION
# ═══════════════════════════════════════════════════════════════

DATA_DIR = Path("data")
ENGINE_FILE = DATA_DIR / "engine_stock.json"

WORKBOOK_NAMES = [
    "TCF VIN  & Paint Float mapping data.xlsx",
    "TCF VIN  & Paint Float mapping data-1.xlsx",
]

RAW_FILE_TYPES = [
    "Paint Float",
    "TCF1 Wiring File",
    "TCF2 Wiring File",
    "TCF1 Cockpit",
    "TCF2 Cockpit",
    "Nova Cockpit",
    "TCF1 DPT Plan",
    "TCF2 DPT Plan",
]

MODEL_MAP = {
    "5497": ("PUNCH", "TCF1"),
    "5468": ("PUNCH.EV / NOVA", "TCF1"),
    "5473": ("HARRIER.EV", "TCF2"),
    "5466": ("SAFARI", "TCF2"),
    "5479": ("SAFARI", "TCF2"),
    "5605": ("SAFARI", "TCF2"),
    "5465": ("HARRIER", "TCF2"),
    "5478": ("HARRIER", "TCF2"),
    "5604": ("HARRIER", "TCF2"),
    "5483": ("SAFARI.EV", "TCF2"),
}

ENGINE_MASTER = [
    {"LINE": "TCF1", "Engine Part No": "54850000PTP001", "Model": "Punch MT SA", "TA Code": "3302"},
    {"LINE": "TCF1", "Engine Part No": "54850000PTP002", "Model": "Punch AMT SA", "TA Code": "3404"},
    {"LINE": "TCF1", "Engine Part No": "54970000PTP002", "Model": "Punch TC MCE", "TA Code": "7349"},
    {"LINE": "TCF1", "Engine Part No": "54970000PTP003", "Model": "Punch MCE MT", "TA Code": "3641"},
    {"LINE": "TCF1", "Engine Part No": "54970000PTP004", "Model": "Punch MCE AMT", "TA Code": "3406"},
    {"LINE": "TCF1", "Engine Part No": "54970000PTP005", "Model": "Punch MCE CNG MT", "TA Code": "3627"},
    {"LINE": "TCF1", "Engine Part No": "54970000PTP031", "Model": "Punch MCE CNG AMT", "TA Code": "3403"},
    {"LINE": "TCF1", "Engine Part No": "546816111212", "Model": "Nova", "TA Code": "5468"},
    {"LINE": "TCF2", "Engine Part No": "572900000118", "Model": "Harrier / Safari Diesel AT", "TA Code": ""},
    {"LINE": "TCF2", "Engine Part No": "572900000120", "Model": "Harrier / Safari Diesel MT", "TA Code": ""},
    {"LINE": "TCF2", "Engine Part No": "54780000PTP001", "Model": "Harrier / Safari Petrol TGDI MT", "TA Code": ""},
    {"LINE": "TCF2", "Engine Part No": "54780000PTP002", "Model": "Harrier / Safari Petrol TGDI AT", "TA Code": ""},
    {"LINE": "TCF2", "Engine Part No": "547380400103", "Model": "Harrier EV", "TA Code": "5473"},
]

FLOAT_COLS = [
    "TOTAL_FLOAT",
    "PBS_FLOAT",
    "PBS_TO_POLISHING",
    "POLISHING_TO_TOPCOAT",
    "TOPCOAT_TO_WETSANDING_ROOFBLACK",
    "TOPCOAT_TO_WETSANDING_FRESH",
    "WETSANDING_TO_SEALANT",
    "TOTAL_UPTO_SEALANT",
    "PT_ENTRY_TO_SEALENT",
    "BIW_LIFTING_TO_PT",
    "PT_BYPASS",
]

PF_HEADERS = [
    "MARKET", "PRODUCT_FAMILY", "SALES_DESCRIPTION", "SHORT_VC", "PACK",
    "TOTAL_FLOAT", "PBS_FLOAT", "PBS_TO_POLISHING", "POLISHING_TO_TOPCOAT",
    "TOPCOAT_TO_WETSANDING_ROOFBLACK", "TOPCOAT_TO_WETSANDING_FRESH",
    "WETSANDING_TO_SEALANT", "TOTAL_UPTO_SEALANT", "PT_ENTRY_TO_SEALENT",
    "BIW_LIFTING_TO_PT", "PT_BYPASS", "WIRING_PART_NUMBER", "COCKPIT",
    "ENGINE", "FOR_MODEL_FLOAT",
]


# ═══════════════════════════════════════════════════════════════
# 2. CSS THEME
# ═══════════════════════════════════════════════════════════════

CUSTOM_CSS = """
<style>
/* Hide default streamlit elements */
header[data-testid="stHeader"] { display: none; }
footer { display: none; }

.dash-header {
    background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
    color: white;
    padding: 24px 32px;
    border-radius: 16px;
    margin-bottom: 24px;
    border: 1px solid #bfdbfe;
}
.dash-header h1 {
    font-size: 2rem !important;
    color: #ffffff !important;
    margin: 0 0 8px 0 !important;
}
.dash-header p {
    color: #e0f2fe !important;
    margin: 0 !important;
    font-size: 1.1rem;
}
.section-title {
    font-size: 1.25rem;
    font-weight: 700;
    margin: 24px 0 12px 0;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 8px;
}
</style>
"""



# ═══════════════════════════════════════════════════════════════
# 3. UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def _safe(ft: str) -> str:
    """Convert file type name to a safe filename slug."""
    return ft.replace(" ", "_").lower()


def find_workbook() -> Path | None:
    """Find the main Excel workbook in the project folder."""
    for name in WORKBOOK_NAMES:
        p = Path(name)
        if p.exists():
            return p
    return None


def strip_df(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from all string columns and headers."""
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": np.nan, "None": np.nan, "": np.nan})
    return df


def save_upload(uploaded_file, file_type: str) -> Path:
    """Persist an uploaded file in the local data directory."""
    DATA_DIR.mkdir(exist_ok=True)
    ext = Path(uploaded_file.name).suffix
    target = DATA_DIR / f"{_safe(file_type)}{ext}"
    with open(target, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return target


def get_source(file_type: str, wb_path: Path | None):
    """Return (source_type, path) — 'uploaded' or 'workbook' or 'scanned'."""
    for ext in [".xlsb", ".xlsx", ".xls"]:
        up = DATA_DIR / f"{_safe(file_type)}{ext}"
        if up.exists():
            return "uploaded", up

    # Scan root directory for matching prefix
    prefix_map = {
        "Paint Float": "PPC_Float_Report",
        "TCF1 DPT Plan": "DPT_PLAN-VIN_GENERATION_REPORT",
        "TCF2 DPT Plan": "TCF2_DPT-PLAN_VIN_GENERATION_REPORT",
        "TCF1 Cockpit": "TCF1_Cockpit",
        "TCF2 Cockpit": "Harrier safari cockpit",
        "Nova Cockpit": "Nova_Cockpit",
        "TCF1 Wiring File": "Wiring Harness report NEW",
        "TCF2 Wiring File": "TCF2_Wiring"
    }
    prefix = prefix_map.get(file_type)
    if prefix:
        root_dir = Path("d:/TML PPC Dashboard")
        matches = [f for f in root_dir.glob("*.*") if f.name.upper().startswith(prefix.upper()) and f.suffix.lower() in [".xls", ".xlsx", ".xlsb"]]
        if matches:
            latest = max(matches, key=os.path.getmtime)
            return "scanned", latest

    if wb_path and wb_path.exists():
        return "workbook", wb_path
    return None, None


def read_sheet(src_type, src_path, sheet_name, **kw):
    """Read a sheet from an uploaded file (sheet 0) or the workbook."""
    engine = "pyxlsb" if str(src_path).lower().endswith(".xlsb") else None
    if src_type in ["uploaded", "scanned"]:
        try:
            return pd.read_excel(src_path, sheet_name=sheet_name, engine=engine, **kw)
        except Exception:
            return pd.read_excel(src_path, engine=engine, **kw)
    return pd.read_excel(src_path, sheet_name=sheet_name, engine=engine, **kw)


def to_excel(df_or_styler, sheet="Summary", table_type="generic") -> bytes:
    """Convert a DataFrame or Styler to downloadable xlsx bytes, adding openpyxl formatting."""
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    buf = BytesIO()
    is_styler = hasattr(df_or_styler, "to_excel") and hasattr(df_or_styler, "data")
    
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df = df_or_styler.data if is_styler else df_or_styler
        
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        if table_type == "engine":
            # Custom Engine & Battery multi-level formatting
            cols_ordered = ["LINE", "Engine Part No", "Model", "TA Code", "Clearance After 6:30AM", "Today VIN", "Bal", "PBS FLOAT", "Float UPTO SEALANT", "TOTAL FLOAT", "With respect to PBS FLOAT", "With respect to Sealant FLOAT", "With respect to Total FLOAT"]
            df = df[[c for c in cols_ordered if c in df.columns]]
            df.to_excel(w, index=False, startrow=2, header=False, sheet_name=sheet)
            ws = w.sheets[sheet]
            
            peach = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
            pink = PatternFill(start_color="E6B8B7", end_color="E6B8B7", fill_type="solid")
            cyan = PatternFill(start_color="92CDDC", end_color="92CDDC", fill_type="solid")
            yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            
            headers = [
                ("LINE", "LINE", peach),
                ("Manual data", "Engine Part No", peach),
                ("Manual data", "Model", peach),
                ("Manual data", "TA Code", peach),
                ("Manual data", "Clearance After 6:30AM", green),
                ("Manual data", "Today VIN", peach),
                ("Manual data", "Bal", pink),
                ("Paint Float", "PBS FLOAT", peach),
                ("Paint Float", "Float UPTO SEALANT", peach),
                ("Paint Float", "TOTAL FLOAT", peach),
                ("Engine & Battery requirement", "With respect to PBS FLOAT", peach),
                ("Engine & Battery requirement", "With respect to Sealant FLOAT", peach),
                ("Engine & Battery requirement", "With respect to Total FLOAT", peach),
            ]
            
            for i, (top, bottom, fill) in enumerate(headers):
                col_idx = i + 1
                c1 = ws.cell(row=1, column=col_idx, value=top)
                c2 = ws.cell(row=2, column=col_idx, value=bottom)
                c1.border = border; c1.alignment = center_align; c1.font = Font(bold=True)
                c2.border = border; c2.alignment = center_align; c2.font = Font(bold=True)
                c1.fill = fill; c2.fill = fill
                
            ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=7)
            ws.merge_cells(start_row=1, start_column=8, end_row=1, end_column=10)
            ws.merge_cells(start_row=1, start_column=11, end_row=1, end_column=13)
            ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
            
            for i, col_name in enumerate(df.columns):
                col_idx = i + 1
                col_letter = get_column_letter(col_idx)
                max_len = len(str(col_name))
                for row in range(3, len(df) + 3):
                    cell = ws.cell(row=row, column=col_idx)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    val = str(cell.value or "")
                    if len(val) > max_len: max_len = len(val)
                    
                    if col_name == "Clearance After 6:30AM": cell.fill = green
                    elif col_name == "Bal": cell.fill = pink
                    elif "With respect to" in col_name and isinstance(cell.value, (int, float)) and cell.value < 0:
                        cell.fill = red_fill
                        cell.font = Font(color="8B0000", bold=True)
                        
                    model_val = str(ws.cell(row=row, column=3).value or "")
                    if "Total" in model_val: cell.fill = cyan
                    elif model_val in ["TCF1", "TCF2"]: cell.fill = yellow
                        
                ws.column_dimensions[col_letter].width = min(max_len + 3, 30)

            start_tcf1 = end_tcf1 = start_tcf2 = end_tcf2 = None
            for row in range(3, len(df) + 3):
                line_val = ws.cell(row=row, column=1).value
                if line_val == "TCF1":
                    if not start_tcf1: start_tcf1 = row
                    end_tcf1 = row
                elif line_val == "TCF2":
                    if not start_tcf2: start_tcf2 = row
                    end_tcf2 = row
            if start_tcf1 and end_tcf1 > start_tcf1: ws.merge_cells(start_row=start_tcf1, start_column=1, end_row=end_tcf1, end_column=1)
            if start_tcf2 and end_tcf2 > start_tcf2: ws.merge_cells(start_row=start_tcf2, start_column=1, end_row=end_tcf2, end_column=1)

        else:
            # Generic / Wiring formatting
            df_or_styler.to_excel(w, index=False, sheet_name=sheet)
            ws = w.sheets[sheet]
            orange_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
            blue_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
            generic_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

            for i, col_name in enumerate(df.columns):
                col_idx = i + 1
                col_letter = get_column_letter(col_idx)
                cell = ws.cell(row=1, column=col_idx)
                
                cell.border = border
                cell.alignment = center_align
                cell.font = Font(bold=True)
                
                if table_type in ["wiring", "cockpit"]:
                    if i < 3: cell.fill = orange_fill
                    else: cell.fill = blue_fill
                else:
                    cell.fill = generic_fill
                    
                max_len = len(str(col_name))
                for row in range(2, len(df) + 2):
                    val = str(ws.cell(row=row, column=col_idx).value or "")
                    if len(val) > max_len: max_len = len(val)
                ws.column_dimensions[col_letter].width = min(max_len + 3, 30)

            for row in range(2, len(df) + 2):
                for col in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="center", vertical="center")

    return buf.getvalue()


def load_engine_json() -> dict:
    """Load persisted engine stock data."""
    if ENGINE_FILE.exists():
        with open(ENGINE_FILE) as f:
            return json.load(f)
    return {}


def save_engine_json(data: dict):
    """Persist engine stock data to JSON."""
    DATA_DIR.mkdir(exist_ok=True)
    with open(ENGINE_FILE, "w") as f:
        json.dump(data, f, indent=2)


def is_valid_vc(val) -> bool:
    """Check if a value looks like a valid Short Vehicle Code."""
    s = str(val).strip()
    return bool(re.match(r"^\d{7,9}[A-Za-z]?$", s))


# ═══════════════════════════════════════════════════════════════
# 4. PARSING FUNCTIONS
# ═══════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Loading Part Number Master…")
def load_part_master(_wb_path_str: str) -> pd.DataFrame:
    """Load the Part Number Master sheet from the main workbook."""
    df = pd.read_excel(_wb_path_str, sheet_name="Part Number Master")
    df = strip_df(df)
    # Normalise column names
    rename = {}
    for c in df.columns:
        cu = c.upper()
        if "SR" in cu and "NO" in cu:
            rename[c] = "SR_NO"
        elif "SHORT" in cu and "VEH" in cu:
            rename[c] = "SHORT_VC"
        elif "WIRING" in cu or "FRONT" in cu:
            rename[c] = "FRONT_WIRING"
        elif "COCKPIT" in cu:
            rename[c] = "COCKPIT"
        elif "ENGINE" in cu:
            rename[c] = "ENGINE"
    df.rename(columns=rename, inplace=True)
    needed = ["SHORT_VC", "FRONT_WIRING", "COCKPIT", "ENGINE"]
    for n in needed:
        if n not in df.columns:
            df[n] = np.nan
    return df[needed].drop_duplicates(subset="SHORT_VC")


def _find_header_row(df_raw, keywords, max_scan=15):
    """Find the row index containing at least one keyword."""
    for i in range(min(max_scan, len(df_raw))):
        vals = [str(v).upper().strip() for v in df_raw.iloc[i] if pd.notna(v)]
        for kw in keywords:
            if any(kw in v for v in vals):
                return i
    return None


def parse_paint_float(src_type, src_path, pm_df=None) -> pd.DataFrame:
    """Parse Paint Float into a clean DataFrame with model/line assignments."""
    df = read_sheet(src_type, src_path, "Paint Float", header=None)

    # Locate the header row (contains MARKET, PRODUCT FAMILY, or SHORT VC)
    h = _find_header_row(df, ["MARKET", "PRODUCT FAMILY", "SHORT VC"])
    if h is None:
        h = 2  # fallback

    data = df.iloc[h + 1 :].copy().reset_index(drop=True)

    # Pad / trim to 16 columns
    while len(data.columns) < 16:
        data[len(data.columns)] = np.nan
    data = data.iloc[:, :16]
    data.columns = PF_HEADERS[:16]
    
    # Initialize extra columns
    data["WIRING_PART_NUMBER"] = ""
    data["COCKPIT"] = ""
    data["ENGINE"] = ""
    data["FOR_MODEL_FLOAT"] = ""

    # Strip strings
    for col in data.columns:
        if data[col].dtype == "object":
            data[col] = data[col].astype(str).str.strip()
            data[col] = data[col].replace({"nan": np.nan, "None": np.nan, "": np.nan, "0": np.nan, "0.0": np.nan})

    # Keep only valid data rows
    data = data[data["SHORT_VC"].notna()]
    data = data[~data.get("PRODUCT_FAMILY", pd.Series(dtype=str)).str.contains("Total", na=False, case=False)]
    data = data[~data["SHORT_VC"].str.contains("Total|Grand|TCF", na=False, case=False)]

    # Merge PM Data
    if pm_df is not None and not pm_df.empty:
        pm_subset = pm_df[["SHORT_VC", "FRONT_WIRING", "COCKPIT", "ENGINE"]].drop_duplicates("SHORT_VC")
        data = data.drop(columns=["WIRING_PART_NUMBER", "COCKPIT", "ENGINE"], errors="ignore")
        data = data.merge(pm_subset, on="SHORT_VC", how="left")
        data.rename(columns={"FRONT_WIRING": "WIRING_PART_NUMBER"}, inplace=True)
    
    # Fill nan after merge
    for c in ["WIRING_PART_NUMBER", "COCKPIT", "ENGINE"]:
        if c not in data: data[c] = ""
        data[c] = data[c].fillna("")

    # Numeric conversions
    for col in FLOAT_COLS:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0).astype(int)

    # Model / Line classification from SHORT_VC column via MODEL_MAP
    def classify_by_vc(vc):
        vc_str = str(vc).strip()[:4]
        for k, (mod, ln) in MODEL_MAP.items():
            if vc_str.startswith(k[:4]):
                return pd.Series([mod, ln])
        return pd.Series(["OTHER", "UNKNOWN"])

    data[["MODEL", "LINE"]] = data["SHORT_VC"].apply(classify_by_vc)
    data = data[data["LINE"] != "UNKNOWN"].reset_index(drop=True)
    return data


def parse_dpt_plan(src_type, src_path, sheet_name, line_label) -> pd.DataFrame:
    """Parse a DPT Plan sheet (clean header row 0).
    Returns: VC, PLAN, TODAY_VIN, WIRING, COCKPIT, ENGINE, FOR_MODEL, LINE
    """
    df = read_sheet(src_type, src_path, sheet_name, header=None)

    # DPT Plan has headers at row 0: MARKET, ProductFamily, VC, SALES DESC,
    # TCF/-Plan, TCF/-VIN, WIRING, Cockpit, Engine, For Model
    h = _find_header_row(df, ["MARKET", "PRODUCTFAMILY", "VC", "SALES"])
    if h is None:
        h = 0

    data = df.iloc[h + 1 :].copy().reset_index(drop=True)
    while len(data.columns) < 10:
        data[len(data.columns)] = np.nan
    data = data.iloc[:, :10]
    data.columns = [
        "MARKET", "PRODUCT_FAMILY", "SHORT_VC", "SALES_DESC",
        "DPT_PLAN", "DPT_VIN", "C6", "C7", "C8", "C9"
    ]
    
    # We will rename C6..C9 back to generic empty if we don't need them
    # But for compatibility, let's ensure we return WIRING, COCKPIT, ENGINE, FOR_MODEL as empty
    data["WIRING"] = ""
    data["COCKPIT"] = ""
    data["ENGINE"] = ""
    data["FOR_MODEL"] = ""
    
    for col in data.columns:
        if data[col].dtype == "object":
            data[col] = data[col].astype(str).str.strip()
            data[col] = data[col].replace({"nan": np.nan, "None": np.nan, "": np.nan})

    data["DPT_PLAN"] = pd.to_numeric(data["DPT_PLAN"], errors="coerce").fillna(0).astype(int)
    data["DPT_VIN"] = pd.to_numeric(data["DPT_VIN"], errors="coerce").fillna(0).astype(int)
    data = data[data["SHORT_VC"].notna() & (data["SHORT_VC"] != "")].reset_index(drop=True)
    data["LINE"] = line_label
    return data


def parse_wiring_tcf1(src_type, src_path) -> pd.DataFrame:
    """Parse TCF1 Wiring File → part, stock, plan per wiring harness."""
    sheet_name = "TCF1 Wiring File"
    if src_type in ["uploaded", "scanned"]:
        engine = "pyxlsb" if str(src_path).lower().endswith(".xlsb") else None
        try:
            with pd.ExcelFile(src_path, engine=engine) as xls:
                sheets = xls.sheet_names
                if sheet_name not in sheets:
                    for s in sheets:
                        if "Coverage" in s or "coverage" in s:
                            sheet_name = s
                            break
                    if sheet_name not in sheets:
                        sheet_name = sheets[0]
        except Exception:
            pass

    df = read_sheet(src_type, src_path, sheet_name, header=None)

    h = _find_header_row(df, ["VC NUMBER", "TOTAL PLAN"])
    if h is None:
        h = 1  # fallback
    header_row = df.iloc[h]
    data = df.iloc[h + 1 :].copy().reset_index(drop=True)

    col_map = {}
    for j, v in enumerate(header_row):
        if pd.isna(v): continue
        vu = str(v).upper().strip()
        if "VC NUMBER" in vu or "VC NO" in vu: col_map.setdefault("VC", j)
        if "WIRING HARNESS" in vu: col_map.setdefault("WH", j)
        if "TOTAL PLAN" in vu: col_map.setdefault("PLAN", j)
        if vu == "STOCK": col_map.setdefault("STOCK", j)

    vc_col = col_map.get("VC", 2)
    wh_col = col_map.get("WH", 3)
    plan_col = col_map.get("PLAN", 5)
    stock_col = col_map.get("STOCK", 8)

    result_rows = []
    for _, row in data.iterrows():
        vc_val = str(row.iloc[vc_col]).strip() if vc_col < len(row) and pd.notna(row.iloc[vc_col]) else ""
        wh_val = str(row.iloc[wh_col]).strip() if wh_col < len(row) and pd.notna(row.iloc[wh_col]) else ""
        stock_val = row.iloc[stock_col] if stock_col < len(row) else np.nan
        plan_val = row.iloc[plan_col] if plan_col < len(row) else np.nan

        # Only keep rows with a valid Short VC and wiring part
        if is_valid_vc(vc_val) and len(wh_val) >= 10 and wh_val[0].isdigit():
            result_rows.append({
                "SHORT_VC": vc_val,
                "WIRING_PART": wh_val,
                "STOCK": pd.to_numeric(stock_val, errors="coerce") or 0,
                "PLAN": pd.to_numeric(plan_val, errors="coerce") or 0,
                "LINE": "TCF1",
            })

    return pd.DataFrame(result_rows) if result_rows else pd.DataFrame(
        columns=["SHORT_VC", "WIRING_PART", "STOCK", "PLAN", "LINE"]
    )


def parse_wiring_tcf2(src_type, src_path) -> pd.DataFrame:
    """Parse TCF2 Wiring File → Short VC, stock (needs PM mapping)."""
    df = read_sheet(src_type, src_path, "TCF2 Wiring File", header=None)

    # Structure: Row 0 = title, Row 1 = sub-headers (Plan, STOCK, …)
    # Data from row 2+. col 1 = Short VC, col 8 = STOCK
    h = _find_header_row(df, ["STOCK", "PLAN", "REMARK"])
    if h is None:
        h = 1
    data = df.iloc[h + 1 :].copy().reset_index(drop=True)

    result_rows = []
    for _, row in data.iterrows():
        vc_val = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        stock_val = row.iloc[8] if len(row) > 8 else np.nan
        plan_val = row.iloc[4] if len(row) > 4 else np.nan

        if is_valid_vc(vc_val):
            s = pd.to_numeric(stock_val, errors="coerce")
            p = pd.to_numeric(plan_val, errors="coerce")
            result_rows.append({
                "SHORT_VC": vc_val,
                "STOCK": s if pd.notna(s) else 0,
                "PLAN": p if pd.notna(p) else 0,
                "LINE": "TCF2",
            })

    return pd.DataFrame(result_rows) if result_rows else pd.DataFrame(
        columns=["SHORT_VC", "STOCK", "PLAN", "LINE"]
    )


def parse_cockpit_file(src_type, src_path, sheet_name, line_label) -> pd.DataFrame:
    """Parse a cockpit / DPT raw file.
    Extracts: Cockpit WH Number, Short VC, Stock/Coverage, Today's O/P.
    """
    if src_type in ["uploaded", "scanned"]:
        engine = "pyxlsb" if str(src_path).lower().endswith(".xlsb") else None
        try:
            with pd.ExcelFile(src_path, engine=engine) as xls:
                sheets = xls.sheet_names
                if sheet_name not in sheets:
                    if "CONSUMPTION" in sheets:
                        sheet_name = "CONSUMPTION"
                    elif "Cockpit" in sheets:
                        sheet_name = "Cockpit"
                    else:
                        sheet_name = sheets[0]
        except Exception:
            pass

    df = read_sheet(src_type, src_path, sheet_name, header=None)

    # Find header row containing "Part No" or "Cockpit WH" or "VC No"
    h = _find_header_row(df, ["PART NO", "COCKPIT WH", "VC NO", "W/H REQUIREMENT", "MODEL", "VEHICLE CODE"])
    if h is None:
        h = 3  # fallback for TCF1 Cockpit format

    header_row = df.iloc[h]
    data = df.iloc[h + 1 :].copy().reset_index(drop=True)
    if data.empty:
        return pd.DataFrame(columns=["COCKPIT_WH", "SHORT_VC", "COVERAGE", "TODAY_OP", "LINE"])

    # Detect column indices from header
    col_map = {}
    for j, v in enumerate(header_row):
        if pd.isna(v):
            continue
        vu = str(v).upper().strip()
        if "PART NO" in vu and "AB12" not in vu and "COCKPIT" not in vu:
            col_map.setdefault("COCKPIT_WH", j)
        if "COCKPIT" in vu and ("WH" in vu or "PART" in vu or "NUMBER" in vu):
            col_map["COCKPIT_WH"] = j
        if vu == "MODEL":
            col_map.setdefault("COCKPIT_WH", j)
            
        if ("VC" in vu or "VEHICLE" in vu) and ("NO" in vu or "CODE" in vu):
            col_map.setdefault("SHORT_VC", j)
        if "SHORT VC" in vu:
            col_map["SHORT_VC"] = j
            
        if "COVERAGE" in vu and "6:30" in vu:
            col_map["COVERAGE"] = j
        if "FLOAT COVERAGE" in vu and "6:30" in vu:
            col_map["COVERAGE"] = j
        if vu == "TOTAL":
            col_map.setdefault("COVERAGE", j)
            
        if "TODAY" in vu and "O/P" in vu:
            col_map.setdefault("TODAY_OP", j)

    # Fallbacks based on known structure
    cockpit_col = col_map.get("COCKPIT_WH", 0)
    vc_col = col_map.get("SHORT_VC", 4)
    coverage_col = col_map.get("COVERAGE", 20)
    today_col = col_map.get("TODAY_OP", 12)

    result_rows = []
    for _, row in data.iterrows():
        cwh = str(row.iloc[cockpit_col]).strip() if pd.notna(row.iloc[cockpit_col]) else ""
        vc = str(row.iloc[vc_col]).strip() if pd.notna(row.iloc[vc_col]) else ""

        if not (len(cwh) >= 6 and cwh[0].isdigit()) and not is_valid_vc(vc):
            continue

        cov = 0
        if coverage_col is not None and coverage_col < len(row):
            cov = pd.to_numeric(row.iloc[coverage_col], errors="coerce")
            cov = cov if pd.notna(cov) else 0

        top = 0
        if today_col < len(row):
            top = pd.to_numeric(row.iloc[today_col], errors="coerce")
            top = top if pd.notna(top) else 0

        result_rows.append({
            "COCKPIT_WH": cwh if len(cwh) >= 6 else np.nan,
            "SHORT_VC": vc,
            "COVERAGE": cov,
            "TODAY_OP": top,
            "LINE": line_label,
        })

    return pd.DataFrame(result_rows) if result_rows else pd.DataFrame(
        columns=["COCKPIT_WH", "SHORT_VC", "COVERAGE", "TODAY_OP", "LINE"]
    )


# ═══════════════════════════════════════════════════════════════
# 5. COMPUTATION FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def compute_model_wise_float(paint_df: pd.DataFrame, dpt_plans: pd.DataFrame | None, line_filter: str) -> pd.DataFrame:
    """Aggregate Paint Float by model and line, append Today VIN."""
    df = paint_df.copy()
    if line_filter != "All Lines":
        df = df[df["LINE"] == line_filter]

    if df.empty:
        return pd.DataFrame()

    agg = {c: "sum" for c in FLOAT_COLS}
    summary = df.groupby(["LINE", "MODEL"], as_index=False).agg(agg)

    if dpt_plans is not None and not dpt_plans.empty:
        dpt = dpt_plans.copy()
        def map_dpt_model(desc):
            desc = str(desc).strip().upper()
            if "NEXON" in desc: return pd.Series(["NEXON", "TCF1"])
            if "PUNCH" in desc: return pd.Series(["PUNCH", "TCF1"])
            if "HARRIER" in desc: return pd.Series(["HARRIER", "TCF2"])
            if "SAFARI" in desc: return pd.Series(["SAFARI", "TCF2"])
            if "TIAGO" in desc: return pd.Series(["TIAGO", "TCF2"])
            if "TIGOR" in desc: return pd.Series(["TIGOR", "TCF2"])
            if "ALTROZ" in desc: return pd.Series(["ALTROZ", "TCF2"])
            return pd.Series(["OTHER", "UNKNOWN"])
        
        # Guard against missing column
        if "SALES_DESC" not in dpt.columns:
            dpt["SALES_DESC"] = ""
            
        dpt[["MODEL_DPT", "LINE_DPT"]] = dpt["SALES_DESC"].apply(map_dpt_model)
        today_vin_df = dpt.groupby(["LINE_DPT", "MODEL_DPT"], as_index=False)["DPT_VIN"].sum().rename(columns={"LINE_DPT": "LINE", "MODEL_DPT": "MODEL", "DPT_VIN": "TODAY_VIN"})
        summary = summary.merge(today_vin_df, on=["LINE", "MODEL"], how="left")
    else:
        summary["TODAY_VIN"] = 0

    summary["TODAY_VIN"] = summary["TODAY_VIN"].fillna(0).astype(int)

    rows_extra = []
    for line in summary["LINE"].unique():
        sub = summary[summary["LINE"] == line]
        total = sub[FLOAT_COLS + ["TODAY_VIN"]].sum()
        total["LINE"] = line
        total["MODEL"] = f"▸ {line} TOTAL"
        rows_extra.append(total)

    grand = summary[FLOAT_COLS + ["TODAY_VIN"]].sum()
    grand["LINE"] = ""
    grand["MODEL"] = "▸▸ GRAND TOTAL"
    rows_extra.append(grand)

    summary = pd.concat([summary, pd.DataFrame(rows_extra)], ignore_index=True)

    display_cols = ["LINE", "MODEL"] + FLOAT_COLS + ["TODAY_VIN"]
    return summary[display_cols].rename(columns={
        "LINE": "Paint Float",
        "MODEL": "MODEL",
        "TOTAL_FLOAT": "TOTAL FLOAT",
        "PBS_FLOAT": "PBS FLOAT",
        "PBS_TO_POLISHING": "PBS TO POLISHING",
        "POLISHING_TO_TOPCOAT": "POLISHING TO TOPCOAT",
        "TOPCOAT_TO_WETSANDING_ROOFBLACK": "TOPCOAT TO WETSANDING ROOFBLACK",
        "TOPCOAT_TO_WETSANDING_FRESH": "TOPCOAT TO WETSANDING FRESH",
        "WETSANDING_TO_SEALANT": "WETSANDING TO SEALANT",
        "TOTAL_UPTO_SEALANT": "TOTAL UPTO SEALANT",
        "PT_ENTRY_TO_SEALENT": "PT ENTRY TO SEALENT",
        "BIW_LIFTING_TO_PT": "BIW LIFTING TO PT",
        "PT_BYPASS": "PT BYPASS",
        "TODAY_VIN": "Today VIN",
    })


def compute_wiring_summary(
    paint_df: pd.DataFrame,
    wiring_tcf1: pd.DataFrame | None,
    wiring_tcf2: pd.DataFrame | None,
    part_master: pd.DataFrame,
    dpt_plans: pd.DataFrame | None,
    line_filter: str,
) -> pd.DataFrame:
    """Compute wiring harness shortage report.

    Float requirements come from Paint Float (grouped by WIRING_PART_NUMBER).
    Stock/clearance come from wiring files or Part Number Master mapping.
    DPT plans provide Today VIN by variant.
    """
    # ── 1. Float requirements from Paint Float ──
    pf = paint_df.copy()
    pf["WIRING_PART_NUMBER"] = pf["WIRING_PART_NUMBER"].astype(str).str.strip()
    pf = pf[pf["WIRING_PART_NUMBER"].notna() & (pf["WIRING_PART_NUMBER"].str.len() >= 10)]

    req = (
        pf.groupby(["WIRING_PART_NUMBER", "LINE"], as_index=False)
        .agg(
            PAINT_TOTAL_FLOAT=("TOTAL_FLOAT", "sum"),
            PBS_FLOAT=("PBS_FLOAT", "sum"),
            CABS_FLOAT_UPTO_SEALANT=("TOTAL_UPTO_SEALANT", "sum"),
        )
    )

    # Also build from Part Number Master for parts not currently in float
    pm_wiring = part_master[["SHORT_VC", "FRONT_WIRING"]].dropna(subset=["FRONT_WIRING"]).drop_duplicates()
    pm_wiring["FRONT_WIRING"] = pm_wiring["FRONT_WIRING"].astype(str).str.strip()
    pm_parts = set(pm_wiring["FRONT_WIRING"].unique())
    existing_parts = set(req["WIRING_PART_NUMBER"].unique())
    missing_parts = pm_parts - existing_parts
    if missing_parts:
        # Determine line for missing parts via PM → Short VC → FOR_MODEL_FLOAT mapping
        extra_rows = []
        for wp in missing_parts:
            vcs = pm_wiring[pm_wiring["FRONT_WIRING"] == wp]["SHORT_VC"].values
            line = "TCF1"  # default
            for vc in vcs:
                prefix = str(vc)[:4]
                for k, (_, ln) in MODEL_MAP.items():
                    if prefix.startswith(k[:4]):
                        line = ln
                        break
            extra_rows.append({
                "WIRING_PART_NUMBER": wp,
                "LINE": line,
                "PAINT_TOTAL_FLOAT": 0,
                "PBS_FLOAT": 0,
                "CABS_FLOAT_UPTO_SEALANT": 0,
            })
        if extra_rows:
            req = pd.concat([req, pd.DataFrame(extra_rows)], ignore_index=True)

    # ── 2. Stock/clearance from wiring files ──
    stock_records = {}  # wiring_part → {STOCK, PLAN}

    # TCF1: has wiring part directly
    if wiring_tcf1 is not None and not wiring_tcf1.empty:
        for _, r in wiring_tcf1.iterrows():
            wp = str(r.get("WIRING_PART", "")).strip()
            if len(wp) >= 10:
                if wp not in stock_records or r["STOCK"] > stock_records[wp]["STOCK"]:
                    stock_records[wp] = {"STOCK": r["STOCK"], "PLAN": r["PLAN"]}

    # TCF2: needs PM mapping Short VC → Front Wiring
    if wiring_tcf2 is not None and not wiring_tcf2.empty:
        tcf2_merged = wiring_tcf2.merge(pm_wiring, on="SHORT_VC", how="left")
        for _, r in tcf2_merged.iterrows():
            wp = str(r.get("FRONT_WIRING", "")).strip()
            if len(wp) >= 10 and wp != "nan":
                if wp not in stock_records or r["STOCK"] > stock_records[wp]["STOCK"]:
                    stock_records[wp] = {"STOCK": r["STOCK"], "PLAN": r["PLAN"]}

    # ── 3. Today VIN from DPT Plans (sum plan per wiring part) ──
    today_vin_map = {}
    if dpt_plans is not None and not dpt_plans.empty:
        dpt = dpt_plans.copy()
        dpt["WIRING"] = dpt["WIRING"].astype(str).str.strip()
        dpt_agg = dpt[dpt["WIRING"].str.len() >= 10].groupby("WIRING", as_index=False)["DPT_VIN"].sum()
        today_vin_map = dict(zip(dpt_agg["WIRING"], dpt_agg["DPT_VIN"]))

    # ── 4. Merge everything ──
    stock_df = pd.DataFrame([
        {"WIRING_PART_NUMBER": k, "CLEARANCE": v["STOCK"], "TODAY_VIN_W": v["PLAN"]}
        for k, v in stock_records.items()
    ])

    if not stock_df.empty:
        summary = req.merge(stock_df, on="WIRING_PART_NUMBER", how="left")
    else:
        summary = req.copy()
        summary["CLEARANCE"] = 0
        summary["TODAY_VIN_W"] = 0

    summary["CLEARANCE"] = summary["CLEARANCE"].fillna(0).astype(int)
    summary["TODAY_VIN_W"] = summary["TODAY_VIN_W"].fillna(0).astype(int)

    # Override Today VIN with DPT plan data if available
    summary["TODAY_VIN"] = summary["WIRING_PART_NUMBER"].map(today_vin_map).fillna(summary["TODAY_VIN_W"]).astype(int)
    summary.drop(columns=["TODAY_VIN_W"], inplace=True, errors="ignore")

    # ── 5. Compute shortage ──
    summary["SHORTAGE_PBS"] = summary["CLEARANCE"] - summary["PBS_FLOAT"]
    summary["SHORTAGE_SEALANT"] = summary["CLEARANCE"] - summary["CABS_FLOAT_UPTO_SEALANT"]
    summary["SHORTAGE_TOTAL"] = summary["CLEARANCE"] - summary["PAINT_TOTAL_FLOAT"]

    # Filter by line
    if line_filter != "All Lines":
        summary = summary[summary["LINE"] == line_filter]

    # Sort: worst shortages first
    summary = summary.sort_values("SHORTAGE_TOTAL", ascending=True).reset_index(drop=True)

    return summary.rename(columns={
        "WIRING_PART_NUMBER": "Wiring Part Number",
        "LINE": "Model/Line",
        "CLEARANCE": "Clearance After 6:30AM",
        "PAINT_TOTAL_FLOAT": "Paint TOTAL FLOAT",
        "PBS_FLOAT": "PBS FLOAT",
        "CABS_FLOAT_UPTO_SEALANT": "Cabs FloatUPTO SEALANT",
        "SHORTAGE_PBS": "Shortage PBS FLOAT",
        "SHORTAGE_SEALANT": "Shortage Upto Sealant",
        "SHORTAGE_TOTAL": "Shortage for TOTAL FLOAT",
    })[["Wiring Part Number", "Model/Line", "Clearance After 6:30AM", "Paint TOTAL FLOAT", "PBS FLOAT", "Cabs FloatUPTO SEALANT", "Shortage PBS FLOAT", "Shortage Upto Sealant", "Shortage for TOTAL FLOAT"]]


def compute_cockpit_summary(
    paint_df: pd.DataFrame,
    cockpit_dfs: list[pd.DataFrame],
    part_master: pd.DataFrame,
    dpt_plans: pd.DataFrame | None,
    line_filter: str,
) -> pd.DataFrame:
    """Compute cockpit assembly shortage report.

    Float requirements come from Paint Float (grouped by COCKPIT column).
    Stock/coverage comes from cockpit raw files mapped via Part Number Master.
    """
    # ── 1. Float requirements from Paint Float ──
    pf = paint_df.copy()
    pf["COCKPIT"] = pf["COCKPIT"].astype(str).str.strip()
    pf = pf[pf["COCKPIT"].notna() & (pf["COCKPIT"].str.len() >= 10)]

    req = (
        pf.groupby(["COCKPIT", "LINE"], as_index=False)
        .agg(
            PAINT_TOTAL_FLOAT=("TOTAL_FLOAT", "sum"),
            PBS_FLOAT=("PBS_FLOAT", "sum"),
            CABS_FLOAT_UPTO_SEALANT=("TOTAL_UPTO_SEALANT", "sum"),
        )
    )

    # Add PM-only parts not currently in float
    pm_cockpit = part_master[["SHORT_VC", "COCKPIT"]].dropna(subset=["COCKPIT"]).drop_duplicates()
    pm_cockpit["COCKPIT"] = pm_cockpit["COCKPIT"].astype(str).str.strip()
    pm_parts = set(pm_cockpit["COCKPIT"].unique())
    existing = set(req["COCKPIT"].unique())
    for cp in pm_parts - existing:
        vcs = pm_cockpit[pm_cockpit["COCKPIT"] == cp]["SHORT_VC"].values
        line = "TCF1"
        for vc in vcs:
            prefix = str(vc)[:4]
            for k, (_, ln) in MODEL_MAP.items():
                if prefix.startswith(k[:4]):
                    line = ln
                    break
        req = pd.concat([req, pd.DataFrame([{
            "COCKPIT": cp, "LINE": line,
            "PAINT_TOTAL_FLOAT": 0, "PBS_FLOAT": 0, "CABS_FLOAT_UPTO_SEALANT": 0,
        }])], ignore_index=True)

    # ── 2. Stock from cockpit files (mapped via Short VC → PM COCKPIT) ──
    stock_map = {}  # cockpit_module_part → {COVERAGE, TODAY_OP}

    combined_cockpit = pd.concat([d for d in cockpit_dfs if d is not None and not d.empty], ignore_index=True)
    if not combined_cockpit.empty:
        # Map Short VC → COCKPIT module via PM
        merged = combined_cockpit.merge(pm_cockpit, on="SHORT_VC", how="left", suffixes=("_raw", "_pm"))
        cockpit_col = "COCKPIT_pm" if "COCKPIT_pm" in merged.columns else "COCKPIT"
        merged["COCKPIT_MODULE"] = merged.get(cockpit_col, pd.Series(dtype=str)).astype(str).str.strip()
        valid = merged[merged["COCKPIT_MODULE"].str.len() >= 10]

        if not valid.empty:
            agg = valid.groupby("COCKPIT_MODULE", as_index=False).agg(
                COVERAGE=("COVERAGE", "sum"),
                TODAY_OP=("TODAY_OP", "sum"),
            )
            stock_map = dict(zip(agg["COCKPIT_MODULE"], agg[["COVERAGE", "TODAY_OP"]].to_dict("records")))

    # ── 3. Today VIN from DPT plans ──
    today_vin_map = {}
    if dpt_plans is not None and not dpt_plans.empty:
        dpt = dpt_plans.copy()
        dpt["COCKPIT"] = dpt["COCKPIT"].astype(str).str.strip()
        dpt_agg = dpt[dpt["COCKPIT"].str.len() >= 10].groupby("COCKPIT", as_index=False)["DPT_VIN"].sum()
        today_vin_map = dict(zip(dpt_agg["COCKPIT"], dpt_agg["DPT_VIN"]))

    # ── 4. Merge ──
    req["CLEARANCE"] = req["COCKPIT"].map(lambda x: stock_map.get(x, {}).get("COVERAGE", 0)).fillna(0).astype(int)
    cov_today = req["COCKPIT"].map(lambda x: stock_map.get(x, {}).get("TODAY_OP", 0)).fillna(0).astype(int)
    req["TODAY_VIN"] = req["COCKPIT"].map(today_vin_map).fillna(cov_today).astype(int)

    # ── 5. Shortage ──
    req["SHORTAGE_PBS"] = req["CLEARANCE"] - req["PBS_FLOAT"]
    req["SHORTAGE_SEALANT"] = req["CLEARANCE"] - req["CABS_FLOAT_UPTO_SEALANT"]
    req["SHORTAGE_TOTAL"] = req["CLEARANCE"] - req["PAINT_TOTAL_FLOAT"]

    if line_filter != "All Lines":
        req = req[req["LINE"] == line_filter]

    req = req.sort_values("SHORTAGE_TOTAL", ascending=True).reset_index(drop=True)

    return req.rename(columns={
        "COCKPIT": "Cockpit Part Number",
        "LINE": "Model/Line",
        "CLEARANCE": "Clearance After 6:30AM",
        "PAINT_TOTAL_FLOAT": "Paint TOTAL FLOAT",
        "PBS_FLOAT": "PBS FLOAT",
        "CABS_FLOAT_UPTO_SEALANT": "Cabs FloatUPTO SEALANT",
        "SHORTAGE_PBS": "Shortage PBS FLOAT",
        "SHORTAGE_SEALANT": "Shortage Upto Sealant",
        "SHORTAGE_TOTAL": "Shortage for TOTAL FLOAT",
    })[["Cockpit Part Number", "Model/Line", "Clearance After 6:30AM", "Paint TOTAL FLOAT", "PBS FLOAT", "Cabs FloatUPTO SEALANT", "Shortage PBS FLOAT", "Shortage Upto Sealant", "Shortage for TOTAL FLOAT"]]


def build_engine_table(paint_df: pd.DataFrame, line_filter: str) -> pd.DataFrame:
    """Build engine requirements table combining Paint Float logic with predefined Master list."""
    # 1. Initialize master df
    req = pd.DataFrame(ENGINE_MASTER)
    
    # 2. Group current paint float
    pf = paint_df.copy()
    pf["ENGINE"] = pf["ENGINE"].astype(str).str.strip()
    pf = pf[pf["ENGINE"].notna() & (pf["ENGINE"].str.len() >= 5)]
    
    pf_agg = (
        pf.groupby(["ENGINE", "LINE", "MODEL"], as_index=False)
        .agg(
            TOTAL_FLOAT=("TOTAL_FLOAT", "sum"),
            PBS_FLOAT=("PBS_FLOAT", "sum"),
            UPTO_SEALANT=("TOTAL_UPTO_SEALANT", "sum"),
        )
    )
    
    # 3. Outer merge to capture all master engines + any new ones
    req = pd.merge(req, pf_agg, left_on=["Engine Part No", "LINE"], right_on=["ENGINE", "LINE"], how="outer", suffixes=("", "_pf"))
    
    # 4. Handle any missing data for newly discovered engines
    missing_mask = req["Engine Part No"].isna()
    req.loc[missing_mask, "Engine Part No"] = req.loc[missing_mask, "ENGINE"]
    
    # For model description, fallback to Paint Float generic MODEL if it's missing (a new engine)
    if "MODEL" in req.columns:
        req["Model"] = req["Model"].fillna(req["MODEL"]).fillna("NEW_ENGINE")
    else:
        req["Model"] = req["Model"].fillna("NEW_ENGINE")
    req["TA Code"] = req["TA Code"].fillna("")
    
    req["TOTAL_FLOAT"] = req["TOTAL_FLOAT"].fillna(0).astype(int)
    req["PBS_FLOAT"] = req["PBS_FLOAT"].fillna(0).astype(int)
    req["UPTO_SEALANT"] = req["UPTO_SEALANT"].fillna(0).astype(int)
    
    req = req.drop(columns=["ENGINE", "MODEL"], errors="ignore")
    
    if line_filter != "All Lines":
        req = req[req["LINE"] == line_filter]

    return req

def add_engine_subtotals(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    num_cols = ["Clearance After 6:30AM", "Today VIN", "Bal", "PBS FLOAT", "Float UPTO SEALANT", "TOTAL FLOAT", "With respect to PBS FLOAT", "With respect to Sealant FLOAT", "With respect to Total FLOAT"]

    tcf1_order = ["Punch MT SA", "Punch AMT SA", "Punch TC MCE", "Punch MCE MT", "Punch MCE AMT", "Punch MCE CNG MT", "Punch MCE CNG AMT"]
    tcf2_order = ["Harrier / Safari Diesel AT", "Harrier / Safari Diesel MT", "Harrier / Safari Petrol TGDI MT", "Harrier / Safari Petrol TGDI AT"]

    tcf1_df = df[df["LINE"] == "TCF1"].copy()
    if not tcf1_df.empty:
        is_punch = tcf1_df["Model"].str.contains("PUNCH", case=False, na=False)
        punch_df = tcf1_df[is_punch].copy()
        other_df = tcf1_df[~is_punch].copy()
        
        punch_df['SortKey'] = pd.Categorical(punch_df['Model'], categories=tcf1_order, ordered=True)
        punch_df = punch_df.sort_values('SortKey').drop(columns=['SortKey'])

        rows.extend(punch_df.to_dict('records'))
        if not punch_df.empty:
            sub = {c: punch_df[c].sum() if c in num_cols else "" for c in df.columns}
            sub["Model"] = "1.2 Lit Total"
            sub["LINE"] = "TCF1"
            rows.append(sub)

        rows.extend(other_df.to_dict('records'))

        tcf1_sub = {c: tcf1_df[c].sum() if c in num_cols else "" for c in df.columns}
        tcf1_sub["Model"] = "TCF1"
        tcf1_sub["LINE"] = "TCF1"
        rows.append(tcf1_sub)

    tcf2_df = df[df["LINE"] == "TCF2"].copy()
    if not tcf2_df.empty:
        is_2l = tcf2_df["Model"].str.contains("HARRIER", case=False, na=False) | tcf2_df["Model"].str.contains("SAFARI", case=False, na=False)
        is_ev = tcf2_df["Model"].str.contains("EV", case=False, na=False)
        is_2l = is_2l & ~is_ev

        l2_df = tcf2_df[is_2l].copy()
        ev_df = tcf2_df[~is_2l].copy()
        
        l2_df['SortKey'] = pd.Categorical(l2_df['Model'], categories=tcf2_order, ordered=True)
        l2_df = l2_df.sort_values('SortKey').drop(columns=['SortKey'])

        rows.extend(l2_df.to_dict('records'))
        if not l2_df.empty:
            sub = {c: l2_df[c].sum() if c in num_cols else "" for c in df.columns}
            sub["Model"] = "2 Lit Total"
            sub["LINE"] = "TCF2"
            rows.append(sub)
            
        rows.extend(ev_df.to_dict('records'))
        
        tcf2_sub = {c: tcf2_df[c].sum() if c in num_cols else "" for c in df.columns}
        tcf2_sub["Model"] = "TCF2"
        tcf2_sub["LINE"] = "TCF2"
        rows.append(tcf2_sub)
        
    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════
# 6. STYLING FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def highlight_shortage(val):
    """Traffic-light styling for shortage columns."""
    try:
        v = float(val)
    except (ValueError, TypeError):
        return ""
    if v < 0:
        return "background-color: #ffcccc; color: #8b0000; font-weight: 700"
    if v == 0:
        return "background-color: #fff3cd; color: #856404; font-weight: 600"
    return ""


def style_shortage_df(df: pd.DataFrame, table_type="wiring"):
    """Apply conditional formatting to shortage columns in a dataframe."""
    shortage_cols = [c for c in df.columns if "Shortage" in str(c) or "With respect to" in str(c)]
    
    styler = df.style
    
    styles = {}
    for i, col in enumerate(df.columns):
        if table_type in ["wiring", "cockpit", "engine"]:
            color = '#F4B084' if i < 3 else '#9BC2E6'
        else:
            color = '#D9E1F2'
        styles[col] = [{'selector': 'th', 'props': [('background-color', color), ('color', 'black'), ('border', '1px solid #94a3b8')]}]
    
    styler = styler.set_table_styles(styles, overwrite=False)
    
    if shortage_cols:
        styler = styler.map(highlight_shortage, subset=shortage_cols)
        
    return styler


# ═══════════════════════════════════════════════════════════════
# 7. MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="TCF PPC Dashboard",
        page_icon="🏭",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Detect workbook
    wb_path = find_workbook()

    # ═══════════════════════════════════════════
    # SIDEBAR
    # ═══════════════════════════════════════════
    with st.sidebar:
        st.markdown("## 🏭 TCF PPC Dashboard")
        st.caption("Production Planning & Control")

        st.divider()

        # ── Workbook status ──
        if wb_path:
            st.success(f"✅ Workbook found: `{wb_path.name}`")
        else:
            st.error("❌ Excel workbook not found in project folder.")
            st.info(
                "Place **'TCF VIN  & Paint Float mapping data.xlsx'** "
                "(or the `-1` variant) in the same folder as this script."
            )
            st.stop()

        st.divider()

        # ── Filters ──
        st.markdown("### 📊 Filters")
        line_filter = st.selectbox(
            "Main Line",
            ["All Lines", "TCF1", "TCF2"],
            help="Filter the entire dashboard by production line.",
        )
        lookup = st.text_input(
            "🔍 Global VIN / Part Lookup",
            placeholder="e.g. 54680124A or 546854600108",
        )

        st.divider()

        # ── File Upload ──
        st.markdown("### 📁 Upload Raw Data")
        file_type = st.selectbox("Select file type", RAW_FILE_TYPES)
        uploaded = st.file_uploader(
            f"Upload **{file_type}**",
            type=["xlsb", "xlsx", "xls"],
            key=f"upload_{file_type}",
        )
        if uploaded is not None:
            save_upload(uploaded, file_type)
            st.success(f"✅ **{file_type}** saved successfully!")
            st.cache_data.clear()
            st.rerun()

        st.divider()

        # ── File status ──
        st.markdown("### 📋 File Status")
        DATA_DIR.mkdir(exist_ok=True)
        for ft in RAW_FILE_TYPES:
            fpath = DATA_DIR / f"{_safe(ft)}.xlsx"
            if fpath.exists():
                mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
                age = (datetime.now() - mtime).total_seconds()
                badge_cls = "upload-badge-ok" if age < 86400 else "upload-badge-old"
                ts = mtime.strftime("%d-%b %H:%M")
                st.markdown(
                    f'<span class="{badge_cls}">● {ft}</span> <small style="color:#64748b">Updated {ts}</small>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<span class="upload-badge-old">○ {ft}</span> <small style="color:#475569">— using workbook</small>',
                    unsafe_allow_html=True,
                )

    # ═══════════════════════════════════════════
    # LOAD ALL DATA
    # ═══════════════════════════════════════════

    # ── Part Number Master (always from workbook) ──
    pm = load_part_master(str(wb_path))

    # ── Paint Float ──
    pf_src, pf_path = get_source("Paint Float", wb_path)
    if pf_src is None:
        st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
        st.markdown(
            '<div class="dash-header"><h1>TCF PPC Dashboard</h1>'
            "<p>Upload a Paint Float file to get started.</p></div>",
            unsafe_allow_html=True,
        )
        st.error("⚠️  No Paint Float data found. Upload via the sidebar to begin.")
        st.stop()

    paint_df = parse_paint_float(pf_src, str(pf_path), pm)
    if paint_df.empty:
        st.error("Paint Float parsed but contains no valid data rows.")
        st.stop()

    # ── DPT Plans ──
    dpt_frames = []
    for ft, sheet, label in [("TCF1 DPT Plan", "TCF1 DPT Plan", "TCF1"), ("TCF2 DPT Plan", "TCF2 DPT Plan", "TCF2")]:
        s, p = get_source(ft, wb_path)
        if s:
            try:
                dpt_frames.append(parse_dpt_plan(s, str(p), sheet, label))
            except Exception:
                pass
    dpt_all = pd.concat(dpt_frames, ignore_index=True) if dpt_frames else None

    # ── Wiring Files ──
    w1_src, w1_path = get_source("TCF1 Wiring File", wb_path)
    wiring_tcf1 = parse_wiring_tcf1(w1_src, str(w1_path)) if w1_src else None

    w2_src, w2_path = get_source("TCF2 Wiring File", wb_path)
    wiring_tcf2 = parse_wiring_tcf2(w2_src, str(w2_path)) if w2_src else None

    # ── Cockpit Files ──
    cockpit_dfs = []
    for ft, sheet, label in [
        ("TCF1 Cockpit", "TCF1 Cockpit", "TCF1"),
        ("TCF2 Cockpit", "TCF2 Cockpit", "TCF2"),
        ("Nova Cockpit", "Nova Cockpit", "TCF1"),
    ]:
        s, p = get_source(ft, wb_path)
        if s:
            try:
                cockpit_dfs.append(parse_cockpit_file(s, str(p), sheet, label))
            except Exception:
                pass

    # ═══════════════════════════════════════════
    # HEADER
    # ═══════════════════════════════════════════
    st.markdown(
        '<div class="dash-header">'
        "<h1>🏭 TCF Production Planning Dashboard</h1>"
        f"<p>Live shortage monitoring &nbsp;·&nbsp; Line: <b>{line_filter}</b> "
        f"&nbsp;·&nbsp; Data as of {datetime.now().strftime('%d %b %Y, %H:%M')}</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    # ── Global Lookup ──
    if lookup:
        lookup_u = lookup.strip().upper()
        mask = paint_df.apply(lambda row: any(lookup_u in str(v).upper() for v in row), axis=1)
        hits = mask.sum()
        if hits > 0:
            st.info(f"🔍 **{hits}** variant(s) matched for `{lookup}`")
            match_df = paint_df[mask][["SHORT_VC", "SALES_DESCRIPTION", "MODEL", "LINE", "WIRING_PART_NUMBER", "COCKPIT", "ENGINE"] + FLOAT_COLS]
            match_df = match_df.rename(columns={"SHORT_VC": "Short VC", "SALES_DESCRIPTION": "Description", "WIRING_PART_NUMBER": "Wiring Part", "MODEL": "Model", "LINE": "Line"})
            st.dataframe(match_df, use_container_width=True, hide_index=True)
        else:
            st.warning(f"No matches found for `{lookup}`")

    # ═══════════════════════════════════════════
    # TABS
    # ═══════════════════════════════════════════
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Model Wise Float",
        "🔋 Engine & Battery",
        "🔌 Wiring Harness",
        "🎛️ Cockpit Assembly",
    ])

    # ─────────────────────────────────────────
    # TAB 1 — MODEL WISE FLOAT SUMMARY
    # ─────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-title">Model Wise Float Summary</div>', unsafe_allow_html=True)

        # KPI row
        pf_filtered = paint_df if line_filter == "All Lines" else paint_df[paint_df["LINE"] == line_filter]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Float", f"{int(pf_filtered['TOTAL_FLOAT'].sum()):,}")
        c2.metric("PBS Float", f"{int(pf_filtered['PBS_FLOAT'].sum()):,}")
        c3.metric("Upto Sealant", f"{int(pf_filtered['TOTAL_UPTO_SEALANT'].sum()):,}")
        c4.metric("BIW→PT", f"{int(pf_filtered['BIW_LIFTING_TO_PT'].sum()):,}")
        c5.metric("Variants", f"{len(pf_filtered):,}")

        st.markdown("")

        model_df = compute_model_wise_float(paint_df, dpt_all, line_filter)
        if not model_df.empty:
            # Highlight total rows
            def style_total_row(row):
                if "TOTAL" in str(row.get("MODEL", "")):
                    return ["background-color: #e0e7ff; font-weight:700"] * len(row)
                return [""] * len(row)

            styled = model_df.style
            
            # Apply generic header styles to match other tables
            styles = {col: [{'selector': 'th', 'props': [('background-color', '#D9E1F2'), ('color', 'black'), ('border', '1px solid #94a3b8')]}] for col in model_df.columns}
            styled = styled.set_table_styles(styles, overwrite=False).apply(style_total_row, axis=1)
            
            st.dataframe(styled, use_container_width=True, hide_index=True, height=450)

            st.download_button(
                "📥  Download Model Wise Float",
                to_excel(styled, "Model Wise Float", table_type="generic"),
                file_name=f"model_wise_float_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("No data available for the selected filter.")

    # ─────────────────────────────────────────
    # TAB 2 — ENGINE & BATTERY SUMMARY
    # ─────────────────────────────────────────
    with tab2:
        st.markdown('<div class="section-title">Engine & Battery Summary — Manual Stock Entry</div>', unsafe_allow_html=True)
        st.caption("Edit the **Stock w/ Transit** and **Clearance 6:30AM** columns below. Your entries are saved automatically.")

        engine_req = build_engine_table(paint_df, line_filter)

        if engine_req.empty:
            st.info("No engine data available for the selected filter.")
        else:
            # Load saved engine stock
            saved = load_engine_json()

            # Build editable dataframe
            engine_req["Clearance After 6:30AM"] = engine_req["Engine Part No"].map(lambda e: saved.get(e, {}).get("clearance", 0))
            engine_req["Today VIN"] = engine_req["Engine Part No"].map(lambda e: saved.get(e, {}).get("today_vin", 0))

            display_engine = engine_req[[
                "Engine Part No", "Model", "TA Code", "Clearance After 6:30AM", "Today VIN", 
                "PBS_FLOAT", "UPTO_SEALANT", "TOTAL_FLOAT", "LINE"
            ]].rename(columns={
                "PBS_FLOAT": "PBS FLOAT",
                "UPTO_SEALANT": "Float UPTO SEALANT",
                "TOTAL_FLOAT": "TOTAL FLOAT"
            })

            edited = st.data_editor(
                display_engine,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["Engine Part No", "PBS FLOAT", "Float UPTO SEALANT", "TOTAL FLOAT", "LINE"],
                column_config={
                    "Clearance After 6:30AM": st.column_config.NumberColumn("Clearance After 6:30AM", min_value=0, step=1),
                    "Today VIN": st.column_config.NumberColumn("Today VIN", min_value=0, step=1),
                },
                key="engine_editor",
            )

            # Auto-save on every edit
            new_saved = {}
            for _, row in edited.iterrows():
                ep = row["Engine Part No"]
                new_saved[ep] = {
                    "model": row.get("Model", ""),
                    "ta_code": row.get("TA Code", ""),
                    "clearance": int(row.get("Clearance After 6:30AM", 0) or 0),
                    "today_vin": int(row.get("Today VIN", 0) or 0),
                }
            save_engine_json(new_saved)

            # Compute shortage
            result = edited.copy()
            result["Bal"] = result["Clearance After 6:30AM"] - result["Today VIN"]
            result["With respect to PBS FLOAT"] = result["Bal"] - result["PBS FLOAT"]
            result["With respect to Sealant FLOAT"] = result["Bal"] - result["Float UPTO SEALANT"]
            result["With respect to Total FLOAT"] = result["Bal"] - result["TOTAL FLOAT"]

            st.markdown('<div class="section-title">Computed Shortage</div>', unsafe_allow_html=True)
            
            # Rearrange final columns and inject subtotals
            result = result[[
                "Engine Part No", "Model", "TA Code", "Clearance After 6:30AM", "Today VIN", "Bal",
                "PBS FLOAT", "Float UPTO SEALANT", "TOTAL FLOAT",
                "With respect to PBS FLOAT", "With respect to Sealant FLOAT", "With respect to Total FLOAT", "LINE"
            ]]
            
            result = add_engine_subtotals(result)

            # KPI for engine
            total_short = result["With respect to Total FLOAT"].sum()
            crit = (result["With respect to Total FLOAT"] < 0).sum()
            ec1, ec2, ec3 = st.columns(3)
            ec1.metric("Engine Variants", len(result))
            ec2.metric("Critical Shortages", crit, delta=f"-{crit}" if crit else "0", delta_color="inverse")
            ec3.metric("Net Balance (Total)", int(total_short))

            styled_eng = style_shortage_df(result, table_type="generic")
            st.dataframe(styled_eng, use_container_width=True, hide_index=True, height=400)

            st.download_button(
                "📥  Download Engine Summary",
                to_excel(result, "Engine Summary", table_type="engine"),
                file_name=f"engine_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ─────────────────────────────────────────
    # TAB 3 — WIRING HARNESS TRACK
    # ─────────────────────────────────────────
    with tab3:
        st.markdown('<div class="section-title">Wiring Harness Shortage Track</div>', unsafe_allow_html=True)

        wiring_summary = compute_wiring_summary(
            paint_df, wiring_tcf1, wiring_tcf2, pm, dpt_all, line_filter
        )

        if wiring_summary.empty:
            st.info("No wiring data available for the selected filter.")
        else:
            # KPIs
            total_parts = len(wiring_summary)
            critical = (wiring_summary["Shortage for TOTAL FLOAT"] < 0).sum()
            marginal = (wiring_summary["Shortage for TOTAL FLOAT"] == 0).sum()
            healthy = (wiring_summary["Shortage for TOTAL FLOAT"] > 0).sum()

            wc1, wc2, wc3, wc4 = st.columns(4)
            wc1.metric("Total Parts", total_parts)
            wc2.metric("🔴 Critical", critical, delta=f"-{critical}" if critical else "0", delta_color="inverse")
            wc3.metric("🟡 Marginal", marginal)
            wc4.metric("🟢 Healthy", healthy)

            st.markdown("")

            styled_w = style_shortage_df(wiring_summary, table_type="wiring")
            st.dataframe(styled_w, use_container_width=True, hide_index=True, height=500)

            st.download_button(
                "📥  Download Wiring Summary",
                to_excel(styled_w, "Wiring Summary", table_type="wiring"),
                file_name=f"wiring_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ─────────────────────────────────────────
    # TAB 4 — COCKPIT ASSEMBLY SUMMARY
    # ─────────────────────────────────────────
    with tab4:
        st.markdown('<div class="section-title">Cockpit Assembly Shortage Report</div>', unsafe_allow_html=True)

        cockpit_summary = compute_cockpit_summary(
            paint_df, cockpit_dfs, pm, dpt_all, line_filter
        )

        if cockpit_summary.empty:
            st.info("No cockpit data available for the selected filter.")
        else:
            # KPIs
            total_cp = len(cockpit_summary)
            crit_cp = (cockpit_summary["Shortage for TOTAL FLOAT"] < 0).sum()
            marg_cp = (cockpit_summary["Shortage for TOTAL FLOAT"] == 0).sum()
            ok_cp = (cockpit_summary["Shortage for TOTAL FLOAT"] > 0).sum()

            cc1, cc2, cc3, cc4 = st.columns(4)
            cc1.metric("Total Parts", total_cp)
            cc2.metric("🔴 Critical", crit_cp, delta=f"-{crit_cp}" if crit_cp else "0", delta_color="inverse")
            cc3.metric("🟡 Marginal", marg_cp)
            cc4.metric("🟢 Healthy", ok_cp)

            st.markdown("")

            styled_c = style_shortage_df(cockpit_summary, table_type="cockpit")
            st.dataframe(styled_c, use_container_width=True, hide_index=True, height=500)

            st.download_button(
                "📥  Download Cockpit Summary",
                to_excel(styled_c, "Cockpit Summary", table_type="cockpit"),
                file_name=f"cockpit_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ═══════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()
