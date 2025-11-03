#!/usr/bin/env python3
"""
accrual_updater.py

Complete Accrual Updater System with OT Rate Detection
Version: 3.4.3 - Fixed month regex (only matches start of date)
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from datetime import date, datetime, timedelta
from fractions import Fraction
from typing import Any, Dict, List, Optional, Tuple

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.worksheet import Worksheet
except Exception:
    print("ERROR: Missing openpyxl. Install: pip install openpyxl")
    sys.exit(1)

try:
    import pandas as pd
except Exception:
    print("ERROR: Missing pandas. Install: pip install pandas")
    sys.exit(1)

SUPPORT_XLS = True
try:
    import xlrd
except Exception:
    SUPPORT_XLS = False

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

__version__ = "3.4.3"
__author__ = "ravitejavavilala07-source"
__date__ = "2025-10-22"


# ==============================================================================
# SECTION 1: UTILITY FUNCTIONS
# ==============================================================================

def _normalize_input_date_to_dateobj(s: Optional[str]) -> Optional[date]:
    """
    Parse date string to date object.
    Supports: mm/dd/yyyy, m/d/yy, yyyy-mm-dd, mm-dd-yy, etc.
    """
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    
    fmts = ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y/%m/%d", "%m-%d-%Y", "%m-%d-%y")
    for fmt in fmts:
        try:
            dt = datetime.strptime(s, fmt)
            return date(dt.year, dt.month, dt.day)
        except Exception:
            pass
    
    # Regex fallback
    m = re.match(r'^\s*0?(\d{1,2})[/-]0?(\d{1,2})[/-](\d{2,4})\s*$', s)
    if m:
        mm = int(m.group(1))
        dd = int(m.group(2))
        yy = int(m.group(3))
        if yy < 100:
            yy += 2000
        try:
            return date(yy, mm, dd)
        except Exception:
            return None
    return None


def parse_multiplier_input(s: str) -> float:
    """
    Parse multiplier string.
    Supports: 0.5, 1/2, 50%, .6, 1, etc.
    """
    s = (s or "").strip()
    if s == "":
        return 1.0
    s = s.replace(" ", "")
    
    try:
        if "%" in s:
            return float(s.replace("%", "")) / 100.0
        elif "/" in s:
            parts = s.split("/")
            if len(parts) == 2:
                return float(parts[0]) / float(parts[1])
            return float(Fraction(s))
        else:
            return float(s)
    except Exception:
        return 1.0


def safe_float(x: Any) -> float:
    """Convert any value to float safely"""
    try:
        if x is None:
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x).strip().replace(",", "").replace("$", "")
        s = s.replace("(", "-").replace(")", "")
        if s in ("", "-"):
            return 0.0
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group(0)) if m else 0.0
    except Exception:
        return 0.0


def read_xls_with_xlrd(path: str) -> Dict[str, pd.DataFrame]:
    """Read .xls file using xlrd"""
    if not SUPPORT_XLS:
        raise RuntimeError("xlrd not available. Install: pip install 'xlrd==1.2.0'")
    
    book = xlrd.open_workbook(path, formatting_info=False, on_demand=False)
    out: Dict[str, pd.DataFrame] = {}
    
    for sheet in book.sheets():
        rows = []
        for r in range(sheet.nrows):
            try:
                row_vals = []
                has_data = False
                for c in range(sheet.ncols):
                    try:
                        v = sheet.cell_value(r, c)
                        ctype = sheet.cell_type(r, c)
                        
                        if ctype == xlrd.XL_CELL_DATE:
                            try:
                                dt_tuple = xlrd.xldate_as_tuple(v, book.datemode)
                                row_vals.append(f"{dt_tuple[1]}/{dt_tuple[2]}/{dt_tuple[0]}")
                                has_data = True
                            except Exception:
                                row_vals.append(str(v))
                                if v != "":
                                    has_data = True
                        elif ctype != xlrd.XL_CELL_EMPTY:
                            row_vals.append(v)
                            if v != "":
                                has_data = True
                        else:
                            row_vals.append("")
                    except Exception:
                        row_vals.append("")
                
                if has_data:
                    rows.append(row_vals)
            except Exception:
                break
        
        if rows:
            df = pd.DataFrame(rows)
            if not df.empty:
                df.columns = [str(x) if x is not None else "" for x in df.iloc[0].tolist()]
                df = df.iloc[1:].reset_index(drop=True)
            out[sheet.name] = df
    
    return out


def load_admin_map_csv(path: Optional[str]) -> List[Dict[str, Any]]:
    """Load admin rates from CSV file"""
    out: List[Dict[str, Any]] = []
    if not path or not os.path.exists(path):
        return out
    
    try:
        with open(path, newline='', encoding='utf-8') as fh:
            reader = csv.DictReader(fh)
            for r in reader:
                file_num = str(r.get('file_num') or '').strip() or None
                project_id = str(r.get('project_id') or '').strip() or None
                admin_rate = safe_float(r.get('admin_rate') or 0.0)
                out.append({
                    'file_num': file_num,
                    'project_id': project_id,
                    'admin_rate': float(admin_rate),
                })
    except Exception:
        pass
    
    return out


# ==============================================================================
# SECTION 2: PAYSHEET PARSING
# ==============================================================================

def parse_paysheet(
    path: str,
    month: int,
    year: int,
    debug_log: List[str],
    bc_cols: Optional[List[int]] = None,
    bc_scan_rows: Optional[int] = None,
    bc_payment_min: float = 1.0,
) -> Tuple[float, float, Optional[float], Dict[str, Any]]:
    """
    Parse paysheet for hours and payments (B/C style scan).
    Returns: (hours_sum, payments_sum, salary_candidate, metadata)
    """
    debug_log.append(f"PARSING: {os.path.basename(path)}")
    hours_sum = 0.0
    payments_sum = 0.0
    salary_candidate = None
    meta: Dict[str, Any] = {"path": path, "bc_matches": []}
    
    try:
        ext = os.path.splitext(path)[1].lower()
        dfs: Dict[str, pd.DataFrame] = {}
        
        if ext == ".xls":
            dfs = read_xls_with_xlrd(path)
            debug_log.append("  [Read via xlrd]")
        else:
            try:
                dfs = pd.read_excel(path, sheet_name=None, engine=None)
                debug_log.append("  [Read via pandas]")
            except Exception:
                dfs = pd.read_excel(path, sheet_name=None, engine="openpyxl")
                debug_log.append("  [Read via openpyxl]")

        sheet_items = list(dfs.items())
        selected = [(n, d) for (n, d) in sheet_items if str(year) in str(n)]
        if not selected and sheet_items:
            selected = [sheet_items[0]]

        first_num_re = re.compile(r'(\d{1,2})')
        start_re = re.compile(rf'^\s*(?:{month:02d}|{month})(?:[/-]|\b)')

        def first_numeric_token_from_text(s: Any) -> Optional[int]:
            if s is None:
                return None
            txt = re.sub(r'[A-Za-z]+', '', str(s))
            m = first_num_re.search(txt)
            return int(m.group(1)) if m else None

        for sname, df in selected:
            if df is None or df.shape[1] == 0:
                continue
            
            df2 = df.copy()
            df2.columns = [str(c) if c is not None else "" for c in df2.columns]
            nrows, ncols = df2.shape

            bc_cols_use = bc_cols if bc_cols else [1, 2]
            rows_to_scan = nrows if (not bc_scan_rows or bc_scan_rows <= 0) else min(bc_scan_rows, nrows)
            
            for r in range(rows_to_scan):
                for col_idx in bc_cols_use:
                    if col_idx >= ncols:
                        continue
                    
                    raw = df2.iat[r, col_idx]
                    if raw is None:
                        continue
                    
                    txt_raw = str(raw).strip()
                    if not txt_raw:
                        continue
                    
                    first_num = first_numeric_token_from_text(txt_raw)
                    if first_num is None or first_num != month:
                        continue
                    
                    looks_like_period = ('-' in txt_raw)
                    starts_with_month = bool(start_re.search(txt_raw))
                    if not (looks_like_period or starts_with_month):
                        continue
                    
                    hours_col = col_idx + 1
                    try:
                        hours_val = safe_float(df2.iat[r, hours_col]) if hours_col < ncols else 0.0
                    except Exception:
                        hours_val = 0.0
                    
                    if not (hours_val > 0 and hours_val <= 500):
                        continue
                    
                    payment_val = 0.0
                    payment_col = hours_col + 1
                    if payment_col < ncols:
                        try:
                            pv = safe_float(df2.iat[r, payment_col])
                            if pv >= float(bc_payment_min):
                                payment_val = pv
                        except Exception:
                            pass
                    
                    hours_sum += hours_val
                    payments_sum += payment_val

        debug_log.append(f"  Hours: {hours_sum:.4f} | Payments: ${payments_sum:.2f}")
        return round(hours_sum, 4), round(payments_sum, 2), salary_candidate, meta
    
    except Exception as e:
        debug_log.append(f"  ✗ ERROR: {e}")
        raise RuntimeError(f"Failed parsing {path}: {e}")


# ==============================================================================
# SECTION 3: WORKBOOK MANAGEMENT
# ==============================================================================

def safe_load_workbook(path: str, logger: Optional[callable] = None) -> openpyxl.Workbook:
    """Load workbook (auto-convert if needed)"""
    if logger:
        logger(f"Opening: {path}")
    
    try:
        if zipfile.is_zipfile(path):
            return openpyxl.load_workbook(filename=path)
    except Exception:
        pass

    sheets = None
    for engine in ["xlrd", None, "openpyxl"]:
        try:
            sheets = pd.read_excel(path, sheet_name=None, engine=engine)
            if isinstance(sheets, dict):
                if logger:
                    logger(f"  Converted via {engine or 'default'}")
                break
        except Exception:
            pass

    if sheets is None:
        raise RuntimeError(f"Failed to read '{path}'")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.splitext(path)[0] + f"_converted_{ts}.xlsx"
    
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            if not hasattr(df, "to_excel"):
                df = pd.DataFrame(df)
            df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
    
    if logger:
        logger(f"  Saved converted: {out_path}")
    
    return openpyxl.load_workbook(filename=out_path)


def find_headers(ws: Worksheet, header_row: int, month_name: str) -> Dict[str, Optional[int]]:
    """Find column headers in master sheet"""
    headers: Dict[str, Optional[int]] = {}
    header_cells: Dict[int, str] = {}
    
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ""):
            header_cells[c] = str(v).strip()

    def find_by_keywords(keywords: List[str]) -> Optional[int]:
        for c, text in header_cells.items():
            if any(k.lower() in text.lower() for k in keywords):
                return c
        return None

    headers['file_col'] = find_by_keywords([
        'app id', 'appid', 'app_id', 'applicant', 'file', 'file #', 'file id', 'app no'
    ])
    headers['payroll_name_col'] = find_by_keywords([
        'employee', 'emp', 'name', 'payroll name', 'consultant name'
    ])
    headers['accrual_hours_col'] = find_by_keywords([
        f"{month_name} hours", f"{month_name.lower()} hours"
    ])
    headers['billed_col'] = find_by_keywords([
        f"{month_name} billed", "billed", "billed to the client"
    ])
    headers['admin_fee_col'] = find_by_keywords(['admin fee', 'adminfee', 'admin'])
    headers['wages_earned_col'] = find_by_keywords(['wages earned', 'wagesearned'])
    headers['salary_paid_col'] = find_by_keywords(['salary paid', 'salarypaid', 'salary'])
    
    gross_col = find_by_keywords(['gross salary', 'grosssalary', 'gross'])
    if gross_col is None and ws.max_column >= 22:
        gross_col = 22
    headers['gross_salary_col'] = gross_col
    
    return headers


def create_month_columns_if_missing(
    wb: openpyxl.Workbook, ws: Worksheet, header_row: int, month_name: str
) -> Tuple[int, int]:
    """Create month columns if they don't exist"""
    found = find_headers(ws, header_row, month_name)
    accr = found.get('accrual_hours_col')
    billed = found.get('billed_col')
    
    if accr and billed:
        return accr, billed
    
    last = 1
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=header_row, column=c).value not in (None, ""):
            last = c
    
    if not accr:
        last += 1
        ws.cell(row=header_row, column=last, value=f"{month_name} Hours")
        accr = last
    
    if not billed:
        last += 1
        ws.cell(row=header_row, column=last, value=f"{month_name} Billed to the Client")
        billed = last
    
    return accr, billed


def unmerge_cell_if_merged(ws: Worksheet, row: int, col: int):
    """Unmerge cell if it's in a merged range"""
    if not col:
        return
    
    coord = f"{get_column_letter(col)}{row}"
    to_unmerge = []
    for mr in ws.merged_cells.ranges:
        if coord in mr:
            to_unmerge.append(mr)
    
    for mr in to_unmerge:
        try:
            ws.unmerge_cells(str(mr))
        except Exception:
            pass


# ==============================================================================
# SECTION 4: COLUMN AB CALCULATION - FIXED v5.1 (CORRECT COLUMN LOGIC)
# ==============================================================================

def _other_cell_contains_keyword(text: Any) -> bool:
    """Check for retro/ACH keywords"""
    try:
        if text is None:
            return False
        s = str(text).lower()
        return ('retro' in s) or ('ach' in s)
    except Exception:
        return False


def find_amount_for_date_in_paysheet(path: str, target_date: date, debug_log: List[str]) -> float:
    """
    FIXED v5.1: Search for a date in paysheet and extract its amount.
    """
    debug_log.append(f"\n  Searching paysheet for date: {target_date.strftime('%m/%d/%Y')}")
    
    try:
        ext = os.path.splitext(path)[1].lower()
        preferred_date_cols = [6, 7]  # G, H (0-based)
        fallback_window = range(4, 9)  # E..I
        
        # ===== XLS PATH (xlrd) =====
        if ext == ".xls" and SUPPORT_XLS:
            book = xlrd.open_workbook(path, formatting_info=False, on_demand=True)
            for sheet in book.sheets():
                nrows = sheet.nrows
                ncols = sheet.ncols
                
                # First pass: preferred columns G, H
                for r in range(nrows):
                    for c in preferred_date_cols:
                        if c >= ncols:
                            continue
                        
                        v = sheet.cell_value(r, c)
                        cell_type = sheet.cell_type(r, c)
                        
                        found_date = None
                        if cell_type == xlrd.XL_CELL_DATE:
                            try:
                                dt_tuple = xlrd.xldate_as_tuple(v, book.datemode)
                                found_date = date(dt_tuple[0], dt_tuple[1], dt_tuple[2])
                            except Exception:
                                found_date = None
                        else:
                            found_date = _normalize_input_date_to_dateobj(str(v))
                        
                        if found_date and found_date == target_date:
                            debug_log.append(f"    ✓ Found date in column {chr(71 + c)}, row {r}")
                            
                            amt_col = c + 1
                            base_amt = 0.0
                            
                            if amt_col < ncols:
                                base_amt = safe_float(sheet.cell_value(r, amt_col))
                            
                            debug_log.append(f"    → Base amount: ${base_amt:.2f}")
                            
                            total_adjustments = 0.0
                            nr = r + 1
                            
                            while nr < nrows:
                                label_cell = None
                                if c < ncols:
                                    label_cell = sheet.cell_value(nr, c)
                                
                                if _other_cell_contains_keyword(label_cell):
                                    if amt_col < ncols:
                                        adj_amt = safe_float(sheet.cell_value(nr, amt_col))
                                        total_adjustments += adj_amt
                                        debug_log.append(f"    → Found {label_cell} in row {nr}: ${adj_amt:.2f}")
                                        nr += 1
                                        continue
                                
                                break
                            
                            if total_adjustments > 0:
                                debug_log.append(f"    → Total adjustments: ${total_adjustments:.2f}")
                            
                            combined_amt = base_amt + total_adjustments
                            debug_log.append(f"    → Combined (Base + Adjustments): ${combined_amt:.2f}")
                            
                            return combined_amt
                
                # Second pass: fallback window E..I
                for r in range(nrows):
                    for c in fallback_window:
                        if c >= ncols:
                            continue
                        
                        v = sheet.cell_value(r, c)
                        found_date = None
                        try:
                            if sheet.cell_type(r, c) == xlrd.XL_CELL_DATE:
                                dt_tuple = xlrd.xldate_as_tuple(v, book.datemode)
                                found_date = date(dt_tuple[0], dt_tuple[1], dt_tuple[2])
                            else:
                                found_date = _normalize_input_date_to_dateobj(str(v))
                        except Exception:
                            found_date = None
                        
                        if found_date and found_date == target_date:
                            debug_log.append(f"    ✓ Found date in column {chr(65 + c)}, row {r} (fallback)")
                            
                            amt_col = c + 1
                            base_amt = 0.0
                            
                            if amt_col < ncols:
                                base_amt = safe_float(sheet.cell_value(r, amt_col))
                            
                            debug_log.append(f"    → Base amount: ${base_amt:.2f}")
                            
                            total_adjustments = 0.0
                            nr = r + 1
                            
                            while nr < nrows:
                                label_cell = None
                                if c < ncols:
                                    label_cell = sheet.cell_value(nr, c)
                                
                                if _other_cell_contains_keyword(label_cell):
                                    if amt_col < ncols:
                                        adj_amt = safe_float(sheet.cell_value(nr, amt_col))
                                        total_adjustments += adj_amt
                                        debug_log.append(f"    → Found {label_cell} in row {nr}: ${adj_amt:.2f}")
                                        nr += 1
                                        continue
                                
                                break
                            
                            if total_adjustments > 0:
                                debug_log.append(f"    → Total adjustments: ${total_adjustments:.2f}")
                            
                            combined_amt = base_amt + total_adjustments
                            debug_log.append(f"    → Combined (Base + Adjustments): ${combined_amt:.2f}")
                            
                            return combined_amt
            
            debug_log.append(f"    ✗ Date not found in paysheet")
            return 0.0
        
        # ===== PANDAS PATH (.xlsx/.xlsm) =====
        else:
            dfs = pd.read_excel(path, sheet_name=None, engine=None)
            
            for sname, df in dfs.items():
                if df is None:
                    continue
                
                nrows, ncols = df.shape
                
                # First pass: preferred columns
                for r in range(nrows):
                    for c in preferred_date_cols:
                        if c >= ncols:
                            continue
                        
                        try:
                            v = df.iat[r, c]
                        except Exception:
                            v = None
                        
                        if v is None or (isinstance(v, float) and pd.isna(v)):
                            continue
                        
                        found_date = None
                        if isinstance(v, (datetime, date)):
                            found_date = date(v.year, v.month, v.day)
                        else:
                            found_date = _normalize_input_date_to_dateobj(str(v))
                        
                        if found_date and found_date == target_date:
                            debug_log.append(f"    ✓ Found date in column {chr(71 + c)}, row {r}")
                            
                            amt_col = c + 1
                            base_amt = 0.0
                            
                            if amt_col < ncols:
                                try:
                                    base_amt = safe_float(df.iat[r, amt_col])
                                except Exception:
                                    base_amt = 0.0
                            
                            debug_log.append(f"    → Base amount: ${base_amt:.2f}")
                            
                            total_adjustments = 0.0
                            nr = r + 1
                            
                            while nr < nrows:
                                label_cell = None
                                if c < ncols:
                                    try:
                                        label_cell = df.iat[nr, c]
                                    except Exception:
                                        label_cell = None
                                
                                if _other_cell_contains_keyword(label_cell):
                                    if amt_col < ncols:
                                        try:
                                            adj_amt = safe_float(df.iat[nr, amt_col])
                                            total_adjustments += adj_amt
                                            debug_log.append(f"    → Found {label_cell} in row {nr}: ${adj_amt:.2f}")
                                            nr += 1
                                            continue
                                        except Exception:
                                            pass
                                
                                break
                            
                            if total_adjustments > 0:
                                debug_log.append(f"    → Total adjustments: ${total_adjustments:.2f}")
                            
                            combined_amt = base_amt + total_adjustments
                            debug_log.append(f"    → Combined (Base + Adjustments): ${combined_amt:.2f}")
                            
                            return combined_amt
                
                # Second pass: fallback
                for r in range(nrows):
                    for c in fallback_window:
                        if c >= ncols:
                            continue
                        
                        try:
                            v = df.iat[r, c]
                        except Exception:
                            v = None
                        
                        if v is None or (isinstance(v, float) and pd.isna(v)):
                            continue
                        
                        found_date = None
                        if isinstance(v, (datetime, date)):
                            found_date = date(v.year, v.month, v.day)
                        else:
                            found_date = _normalize_input_date_to_dateobj(str(v))
                        
                        if found_date and found_date == target_date:
                            debug_log.append(f"    ✓ Found date in column {chr(65 + c)}, row {r} (fallback)")
                            
                            amt_col = c + 1
                            base_amt = 0.0
                            
                            if amt_col < ncols:
                                try:
                                    base_amt = safe_float(df.iat[r, amt_col])
                                except Exception:
                                    base_amt = 0.0
                            
                            debug_log.append(f"    → Base amount: ${base_amt:.2f}")
                            
                            total_adjustments = 0.0
                            nr = r + 1
                            
                            while nr < nrows:
                                label_cell = None
                                if c < ncols:
                                    try:
                                        label_cell = df.iat[nr, c]
                                    except Exception:
                                        label_cell = None
                                
                                if _other_cell_contains_keyword(label_cell):
                                    if amt_col < ncols:
                                        try:
                                            adj_amt = safe_float(df.iat[nr, amt_col])
                                            total_adjustments += adj_amt
                                            debug_log.append(f"    → Found {label_cell} in row {nr}: ${adj_amt:.2f}")
                                            nr += 1
                                            continue
                                        except Exception:
                                            pass
                                
                                break
                            
                            if total_adjustments > 0:
                                debug_log.append(f"    → Total adjustments: ${total_adjustments:.2f}")
                            
                            combined_amt = base_amt + total_adjustments
                            debug_log.append(f"    → Combined (Base + Adjustments): ${combined_amt:.2f}")
                            
                            return combined_amt
            
            debug_log.append(f"    ✗ Date not found in paysheet")
            return 0.0
    
    except Exception as e:
        debug_log.append(f"    ✗ ERROR: {e}")
        return 0.0


# ==============================================================================
# SECTION 5: OT RATE DETECTION (FIXED v3.4.3 - Month Regex)
# ==============================================================================

def find_total_rate_and_ot_rate_in_sheet(sheet) -> Tuple[float, float]:
    """Find Total Rate and OT Rate from paysheet header"""
    total_rate = 0.0
    ot_rate = 0.0
    
    for row in range(min(20, sheet.nrows)):
        for col in range(sheet.ncols):
            try:
                cell_value = sheet.cell_value(row, col)
                
                if not cell_value:
                    continue
                
                cell_text = str(cell_value).strip().lower()
                
                if "total rate" in cell_text:
                    if col + 1 < sheet.ncols:
                        total_rate = safe_float(sheet.cell_value(row, col + 1))
                
                if "ot rate" in cell_text:
                    if col + 1 < sheet.ncols:
                        ot_rate = safe_float(sheet.cell_value(row, col + 1))
            
            except Exception:
                continue
    
    return total_rate, ot_rate


def detect_ot_suffix(period_str: str) -> bool:
    """Detect if period contains "-OT" suffix"""
    if not period_str:
        return False
    
    period_upper = str(period_str).upper().strip()
    
    if re.search(r'-\s*OT\s*$', period_upper):
        return True
    
    if period_upper.endswith("-OT"):
        return True
    
    if re.search(r'\s+OT\s*$', period_upper):
        return True
    
    return False


def get_rate_for_period_with_ot(period_str: str, total_rate: float, ot_rate: float) -> float:
    """Determine correct rate for a period based on -OT suffix"""
    if detect_ot_suffix(period_str):
        return ot_rate
    else:
        return total_rate


def calculate_billed_with_ot(path: str, month: int, year: int, debug_log: List[str]) -> float:
    """
    ✅ FIXED v3.4.3: Calculate billed amount with OT rate detection
    NOW uses correct month regex (only matches start of date)
    Fixes issue where "04/06" was matching month=6
    """
    try:
        ext = os.path.splitext(path)[1].lower()
        if ext != ".xls" or not SUPPORT_XLS:
            return 0.0
        
        book = xlrd.open_workbook(path, formatting_info=False)
        
        # Find sheet with year in name
        sheet = None
        for s in book.sheets():
            if str(year) in s.name:
                sheet = s
                break
        
        if sheet is None:
            sheet = book.sheet_by_index(0)
        
        debug_log.append(f"Using sheet: {sheet.name}")
        
        # Get rates
        total_rate, ot_rate = find_total_rate_and_ot_rate_in_sheet(sheet)
        
        if total_rate <= 0 or ot_rate <= 0:
            debug_log.append(f"Rates not found: total={total_rate}, ot={ot_rate}")
            return 0.0
        
        debug_log.append(f"✓ OT Rates Found: Total=${total_rate:.2f}, OT=${ot_rate:.2f}")
        
        # Find Hours & Payment header
        payment_col = None
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                val = sheet.cell_value(r, c)
                if val:
                    text = str(val).lower().strip()
                    if text in ('hours & payment', 'work period'):
                        payment_col = c
                        debug_log.append(f"Header found at Row {r}, Col {c}")
                        break
            if payment_col is not None:
                break
        
        if payment_col is None:
            debug_log.append("Header not found")
            return 0.0
        
        # Calculate billed
        total_billed = 0.0
        
        for row in range(payment_col, sheet.nrows):
            try:
                period_text = sheet.cell_value(row, payment_col)
                
                if not period_text:
                    continue
                
                period_str = str(period_text).strip().lower()
                
                # Skip headers
                if any(x in period_str for x in ['total', 'hours', 'payment', 'deductions']):
                    continue
                
                # ✅ FIXED v3.4.3: Match only if month is at START of date (MM/DD format)
                # Pattern: ^0?{month}[/-]
                # Examples:
                #   "06/01-06/07/2025" matches month=6 ✓
                #   "04/06-04/12/2025" does NOT match month=6 ✓
                #   "07/06-07/12/2025" does NOT match month=6 ✓
                if not re.search(rf'^0?{month}[/-]', period_str):
                    continue
                
                # Get hours
                hours_col = payment_col + 1
                if hours_col >= sheet.ncols:
                    continue
                
                hours_val = safe_float(sheet.cell_value(row, hours_col))
                
                if hours_val <= 0:
                    continue
                
                # Get rate
                is_ot = detect_ot_suffix(period_str)
                rate = ot_rate if is_ot else total_rate
                amount = hours_val * rate
                total_billed += amount
                
                debug_log.append(f"  {period_str:30} {hours_val:5.0f} × ${rate:6.2f} = ${amount:10.2f}")
            
            except Exception:
                continue
        
        return round(total_billed, 2)
    
    except Exception as e:
        debug_log.append(f"ERROR: {e}")
        return 0.0


# ==============================================================================
# SECTION 6: MAIN ACCRUAL SYSTEM CLASS
# ==============================================================================

class AccrualUpdater:
    """Main Accrual Updater System"""
    
    def __init__(
        self,
        master_path: str,
        sheet_name: str = "Profit Sharing",
        header_row: int = 3,
        month: str = "June",
        year: int = 2025,
        paysheets_folder: str = "",
        date_multiplier_pairs: Optional[List[Tuple[str, float]]] = None,
        map_path: Optional[str] = None,
        dry_run: bool = True,
        backup: bool = False,
        bc_cols: Optional[List[int]] = None,
        bc_scan_rows: Optional[int] = None,
        bc_payment_min: float = 1.0,
        admin_map_path: Optional[str] = None,
        default_admin_rate: float = 0.0,
        prefer_paysheet_admin: bool = True,
        enable_ot_detection: bool = True,
        **kwargs,
    ):
        self.master_path = master_path
        self.sheet_name = sheet_name
        self.header_row = header_row
        self.month = month
        try:
            self.month_index = MONTHS.index(month) + 1
        except Exception:
            self.month_index = 1
        self.year = year
        self.paysheets_folder = paysheets_folder
        self.date_multiplier_pairs = date_multiplier_pairs or []
        self.map_path = map_path
        self.dry_run = dry_run
        self.backup = backup
        self.bc_cols = bc_cols
        self.bc_scan_rows = bc_scan_rows
        self.bc_payment_min = bc_payment_min
        self.admin_map_path = admin_map_path or map_path
        self.default_admin_rate = float(default_admin_rate or 0.0)
        self.prefer_paysheet_admin = prefer_paysheet_admin
        self.enable_ot_detection = enable_ot_detection
        self.log_lines: List[str] = []
        self.debug_log: List[str] = []
        self.updated_count = 0
        self.no_match_count = 0
        self.failed_count = 0
        self.start_time = None
        
        try:
            self.admin_map_rows = load_admin_map_csv(self.admin_map_path) if self.admin_map_path else []
        except Exception:
            self.admin_map_rows = []

    def log(self, line: str):
        """Log message"""
        print(line)
        self.log_lines.append(line)

    def build_master_lookup(self, ws: Worksheet) -> Dict[str, Dict[str, Any]]:
        """Build lookup of file_number -> master_row"""
        lookup: Dict[str, Dict[str, Any]] = {}
        headers = find_headers(ws, self.header_row, self.month)
        file_col = headers.get('file_col')
        name_col = headers.get('payroll_name_col')
        
        for r in range(self.header_row + 1, ws.max_row + 1):
            try:
                v = ws.cell(row=r, column=file_col).value if file_col else None
            except Exception:
                v = None
            
            fnum = None
            if v is not None:
                m = re.search(r'(\d{5,6})', str(v))
                if m:
                    fnum = m.group(1)
                else:
                    digits = re.sub(r'\D', '', str(v))
                    if len(digits) >= 5:
                        fnum = digits[:6]
            
            if fnum:
                name_val = ws.cell(row=r, column=name_col).value if name_col else None
                lookup[fnum] = {"row": r, "name": name_val}
        
        return lookup

    def process(self):
        """Main process"""
        self.start_time = time.time()
        
        self.log("\n" + "="*80)
        self.log("ACCRUAL UPDATER v" + __version__)
        self.log("="*80)
        self.log(f"Author: {__author__}")
        self.log(f"Date: {__date__} UTC")
        self.log(f"Version: {__version__} - OT RATE DETECTION (Month Regex Fixed)")
        self.log("="*80 + "\n")

        if not os.path.exists(self.master_path):
            raise RuntimeError(f"Master file not found: {self.master_path}")
        if not os.path.exists(self.paysheets_folder):
            raise RuntimeError(f"Paysheets folder not found: {self.paysheets_folder}")

        self.log(f"Master: {self.master_path}")
        self.log(f"Paysheets: {self.paysheets_folder}")
        self.log(f"Month: {self.month} {self.year}")
        self.log(f"OT Detection: {'ENABLED' if self.enable_ot_detection else 'DISABLED'}")
        self.log(f"Date-Multiplier Pairs: {len(self.date_multiplier_pairs)}\n")

        wb = safe_load_workbook(self.master_path, logger=self.log)
        if self.sheet_name not in wb.sheetnames:
            raise RuntimeError(f"Sheet '{self.sheet_name}' not found")
        ws = wb[self.sheet_name]

        headers = find_headers(ws, self.header_row, self.month)
        accrual_col, billed_col = create_month_columns_if_missing(wb, ws, self.header_row, self.month)
        file_col = headers.get('file_col')
        name_col = headers.get('payroll_name_col')
        admin_col = headers.get('admin_fee_col')
        salary_col = headers.get('salary_paid_col')
        gross_col = headers.get('gross_salary_col')

        self.log(f"File Col: {file_col} | Name Col: {name_col} | Hours Col: {accrual_col} | Billed Col: {billed_col}\n")

        if not file_col or not accrual_col or not billed_col:
            raise RuntimeError("Required columns not found!")

        master_lookup = self.build_master_lookup(ws)

        valid_exts = {'.xls', '.xlsx', '.xlsm'}
        files: List[str] = []
        for root, _, fnames in os.walk(self.paysheets_folder):
            for fn in fnames:
                if os.path.splitext(fn)[1].lower() in valid_exts:
                    files.append(os.path.join(root, fn))

        self.log(f"Found {len(files)} paysheet(s)\n")

        V_INDEX = 22   # Column V
        N_INDEX = 14   # Column N
        AB_INDEX = 28  # Column AB

        for idx, fp in enumerate(files, 1):
            fname = os.path.basename(fp)
            self.log(f"[{idx}/{len(files)}] {fname}")
            
            fnum = None
            m = re.search(r'(\d{5,6})', fname)
            if m:
                fnum = m.group(1)
            
            if not fnum:
                try:
                    ext = os.path.splitext(fp)[1].lower()
                    inner = read_xls_with_xlrd(fp) if ext == ".xls" else pd.read_excel(fp, sheet_name=None)
                    found = None
                    if isinstance(inner, dict):
                        for _, df in inner.items():
                            for r in range(min(20, df.shape[0])):
                                for c in range(min(10, df.shape[1])):
                                    try:
                                        v = df.iat[r, c]
                                    except Exception:
                                        v = None
                                    if v:
                                        mm = re.search(r'(\d{5,6})', str(v))
                                        if mm:
                                            found = mm.group(1)
                                            break
                                if found:
                                    break
                            if found:
                                break
                    if found:
                        fnum = found
                except Exception:
                    pass
            
            if not fnum:
                self.no_match_count += 1
                self.log("  ✗ NO FILE NUMBER")
                continue
            
            if fnum not in master_lookup:
                self.no_match_count += 1
                self.log("  ✗ NOT IN MASTER")
                continue
            
            rec = master_lookup[fnum]
            row = rec["row"]
            name_val = rec.get("name") or ""
            self.log(f"  ✓ {name_val} (Row {row})")

            try:
                dbg: List[str] = []
                hours, payments, salary_candidate, meta = parse_paysheet(
                    fp, self.month_index, self.year, dbg,
                    bc_cols=self.bc_cols, bc_scan_rows=self.bc_scan_rows, bc_payment_min=self.bc_payment_min
                )
                for L in dbg:
                    self.debug_log.append(L)

                # ✅ FIXED v3.4.3: Use calculate_billed_with_ot with corrected month regex
                if self.enable_ot_detection:
                    ot_billed = calculate_billed_with_ot(fp, self.month_index, self.year, dbg)
                    if ot_billed > 0:
                        payments = ot_billed

                # FIXED v5.1: Calculate AB from paysheet
                ab_total = 0.0
                if self.date_multiplier_pairs:
                    self.log(f"  Calculating AB from {len(self.date_multiplier_pairs)} date(s):")
                    for date_str, multiplier in self.date_multiplier_pairs:
                        target_date_obj = _normalize_input_date_to_dateobj(date_str)
                        if not target_date_obj:
                            self.log(f"    ✗ Could not parse date: {date_str}")
                            continue
                        
                        combined_amt = find_amount_for_date_in_paysheet(fp, target_date_obj, dbg)
                        computed = round(combined_amt * multiplier, 2)
                        ab_total += computed
                        
                        self.log(f"    {date_str}: ${combined_amt:.2f} × {multiplier} = ${computed:.2f}")
                    
                    self.log(f"  AB Total: ${ab_total:.2f}")

                # Check if hourly
                is_hourly = False
                try:
                    if name_col:
                        name_cell_val = ws.cell(row=row, column=name_col).value
                        if name_cell_val and 'hourly' in str(name_cell_val).lower():
                            is_hourly = True
                    if not is_hourly:
                        row_text = " ".join(str(ws.cell(row=row, column=c).value or "") for c in range(1, min(ws.max_column + 1, 30))).lower()
                        if re.search(r'\bhourly\b', row_text):
                            is_hourly = True
                    if not is_hourly and 'hourly' in fp.lower():
                        is_hourly = True
                except Exception:
                    pass

                # Unmerge cells
                for col in (accrual_col, billed_col, admin_col, salary_col, gross_col, AB_INDEX):
                    if col:
                        unmerge_cell_if_merged(ws, row, col)

                if not self.dry_run:
                    # Write Hours
                    if accrual_col:
                        ws.cell(row=row, column=accrual_col, value=round(hours, 4))
                    
                    # Write Billed
                    if billed_col:
                        ws.cell(row=row, column=billed_col, value=round(payments, 2))

                    # Hourly: Copy N -> V
                    if is_hourly:
                        n_val = ws.cell(row=row, column=N_INDEX).value
                        if n_val is not None and str(n_val).strip() != "":
                            ws.cell(row=row, column=V_INDEX, value=n_val)
                            self.log(f"  ✓ COPIED N→V: {n_val}")

                    # Write AB
                    if ab_total > 0:
                        ws.cell(row=row, column=AB_INDEX, value=round(ab_total, 2))
                        self.log(f"  ✓ AB: ${ab_total:.2f} → AB{row}")
                    else:
                        self.log(f"  → AB: $0.00")
                else:
                    self.log(f"  (dry-run) Hours={hours:.2f} | Billed=${payments:.2f} | AB=${ab_total:.2f}")

                self.updated_count += 1
            
            except Exception as e:
                self.failed_count += 1
                self.log(f"  ✗ ERROR: {e}")
                import traceback
                self.debug_log.append(traceback.format_exc())

        # Save
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        logs_dir = os.path.join(os.path.dirname(self.master_path), "Logs")
        os.makedirs(logs_dir, exist_ok=True)
        
        if not self.dry_run and self.backup:
            backup_name = os.path.splitext(self.master_path)[0] + f"_BACKUP_{ts}.xlsx"
            try:
                shutil.copy2(self.master_path, backup_name)
                self.log(f"✓ Backup: {backup_name}")
            except Exception as e:
                self.log(f"⚠ Backup failed: {e}")
        
        if not self.dry_run:
            try:
                wb.save(self.master_path)
                self.log(f"✓ Saved: {self.master_path}")
            except PermissionError:
                updated_name = os.path.splitext(self.master_path)[0] + f"_UPDATED_{ts}.xlsx"
                wb.save(updated_name)
                self.log(f"⚠ Saved as: {updated_name}")

        result_log = os.path.join(logs_dir, f"Results_{self.month}_{ts}.txt")
        debug_log_path = os.path.join(logs_dir, f"DEBUG_{self.month}_{ts}.txt")
        
        try:
            with open(result_log, "w", encoding="utf-8") as fh:
                for L in self.log_lines:
                    fh.write(L + "\n")
                total_time = time.time() - self.start_time
                fh.write(f"\nUpdated: {self.updated_count} | No Match: {self.no_match_count} | Failed: {self.failed_count}\n")
                fh.write(f"Duration: {int(total_time // 60)}m {int(total_time % 60)}s\n")
            
            with open(debug_log_path, "w", encoding="utf-8") as fh:
                for L in self.debug_log:
                    fh.write(L + "\n")
        except Exception:
            pass

        self.log("\n" + "="*80)
        self.log("SUMMARY")
        self.log("="*80)
        self.log(f"✓ Updated: {self.updated_count}")
        self.log(f"⚠ No Match: {self.no_match_count}")
        self.log(f"✗ Failed: {self.failed_count}")
        self.log("="*80 + "\n")
        
        return {
            "updated": self.updated_count,
            "no_match": self.no_match_count,
            "failed": self.failed_count,
            "log_path": result_log,
            "debug_log_path": debug_log_path
        }


# ==============================================================================
# SECTION 7: CLI
# ==============================================================================

def run_cli():
    """Command-line interface"""
    parser = argparse.ArgumentParser(
        description="Accrual Updater - Update master file from paysheets",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
  python3 accrual_updater.py --master "master.xlsx" --paysheets "paysheets/" --month June --year 2025 --dry-run
  python3 accrual_updater.py --master "master.xlsx" --paysheets "paysheets/" --month June --year 2025 --backup

Version: {__version__}
Author: {__author__}
        """
    )
    
    parser.add_argument("--master", required=True, help="Path to master .xlsx file")
    parser.add_argument("--sheet", default="Profit Sharing", help="Sheet name in master")
    parser.add_argument("--header-row", type=int, default=3, help="Header row number")
    parser.add_argument("--month", required=True, choices=MONTHS, help="Month name")
    parser.add_argument("--year", type=int, required=True, help="Year")
    parser.add_argument("--paysheets", required=True, help="Paysheets folder path")
    parser.add_argument("--dates", nargs="+", help="Dates for AB calculation (mm/dd/yy format)")
    parser.add_argument("--multipliers", nargs="+", type=float, help="Multipliers for dates")
    parser.add_argument("--dry-run", action="store_true", help="Preview only")
    parser.add_argument("--backup", action="store_true", help="Create backup before saving")
    parser.add_argument("--no-ot-detection", action="store_true", help="Disable OT rate detection")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    
    args = parser.parse_args()

    date_multiplier_pairs = []
    if args.dates and args.multipliers:
        if len(args.dates) == len(args.multipliers):
            date_multiplier_pairs = list(zip(args.dates, args.multipliers))
        else:
            print("ERROR: Number of dates must match number of multipliers")
            sys.exit(1)

    updater = AccrualUpdater(
        master_path=args.master,
        sheet_name=args.sheet,
        header_row=args.header_row,
        month=args.month,
        year=args.year,
        paysheets_folder=args.paysheets,
        date_multiplier_pairs=date_multiplier_pairs,
        dry_run=args.dry_run,
        backup=args.backup,
        enable_ot_detection=not args.no_ot_detection,
    )
    
    result = updater.process()
    print("\nResult:", result)


if __name__ == "__main__":
    run_cli()
