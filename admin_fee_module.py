#!/usr/bin/env python3
"""
admin_fee_module.py - UNIVERSAL v17b
OPTION C: Effective dates are ONGOING
- Searches BOTH 10 rows ABOVE AND BELOW each header
"""
import os, re, sys, argparse
from typing import Any, Optional, Tuple

try:
    import xlrd
except ImportError:
    print("ERROR: xlrd not available")
    sys.exit(1)

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

def safe_float(x: Any) -> float:
    try:
        if x is None: return 0.0
        if isinstance(x, (int, float)): return float(x)
        s = str(x).strip().replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
        if s in ("", "-"): return 0.0
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group(0)) if m else 0.0
    except Exception:
        return 0.0

def extract_date_from_period(period_str: str) -> Optional[Tuple[int, int, int]]:
    match = re.match(r'(\d{1,2})/(\d{1,2})', period_str)
    if match:
        month, day = int(match.group(1)), int(match.group(2))
        year_match = re.search(r'(\d{4})', period_str)
        year = int(year_match.group(1)) if year_match else 0
        return (month, day, year)
    return None

def parse_admin_fee_eff_date(text: str) -> Optional[Tuple[int, int, int]]:
    match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', text)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        year = int(match.group(3))
        if year < 100:
            year += 2000
        return (month, day, year)
    return None

def find_headers(sheet) -> list:
    work_period_headers = []
    hours_payment_headers = []
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            if val:
                text = str(val).lower().strip()
                if "work period" in text:
                    work_period_headers.append((r, c, "Work Period"))
                    break
                elif "hours" in text and "payment" in text:
                    hours_payment_headers.append((r, c, "Hours & Payment"))
                    break
    return work_period_headers if work_period_headers else hours_payment_headers

def find_all_admin_fees_around_header(sheet, header_row: int, search_rows: int = 10) -> list:
    admin_fees = []
    search_start = max(0, header_row - search_rows)
    search_end = header_row
    search_end_below = min(sheet.nrows, header_row + search_rows)
    
    for search_range in [(search_start, search_end), (header_row + 1, search_end_below)]:
        for r in range(search_range[0], search_range[1]):
            for c in range(sheet.ncols):
                cell_val = sheet.cell_value(r, c)
                if cell_val:
                    text = str(cell_val).lower().strip()
                    if "admin fee eff" in text:
                        date_tuple = parse_admin_fee_eff_date(text)
                        if date_tuple:
                            if c + 1 < sheet.ncols:
                                rate_val = sheet.cell_value(r, c + 1)
                                rate = safe_float(rate_val)
                                if rate > 0:
                                    admin_fees.append((date_tuple, rate))
                    elif text == 'admin fee' or text == 'adminfee':
                        if c + 1 < sheet.ncols:
                            rate_val = sheet.cell_value(r, c + 1)
                            rate = safe_float(rate_val)
                            if rate > 0:
                                admin_fees.append(((0, 0, 0), rate))
    
    admin_fees.sort(key=lambda x: (x[0][2], x[0][0]) if x[0][2] > 0 else (0, 0))
    return admin_fees

def get_applicable_admin_fee(admin_fees: list, target_month: int, target_year: int) -> float:
    applicable_rate = 0.0
    static_rate = 0.0
    for date_tuple, rate in admin_fees:
        if date_tuple[2] == 0:
            static_rate = rate
            continue
        eff_month, eff_day, eff_year = date_tuple
        eff_comparable = (eff_year, eff_month)
        target_comparable = (target_year, target_month)
        if eff_comparable <= target_comparable:
            applicable_rate = rate
        else:
            break
    return applicable_rate if applicable_rate > 0 else static_rate

def calculate_admin_fee_for_paysheet(paysheet_path: str, month: int, year: int, debug: bool = False) -> Tuple[float, float, float]:
    total_hours = 0.0
    total_admin_fee = 0.0
    last_rate = 0.0
    
    try:
        book = xlrd.open_workbook(paysheet_path, formatting_info=False)
        sheet = None
        for s in book.sheets():
            if "2025" in s.name:
                sheet = s
                break
        if not sheet:
            return 0.0, 0.0, 0.0
        
        headers = find_headers(sheet)
        if not headers:
            return 0.0, 0.0, 0.0
        
        for proj_idx, (hp_row, hp_col, header_type) in enumerate(headers):
            if proj_idx < len(headers) - 1:
                next_header_row = headers[proj_idx + 1][0]
            else:
                next_header_row = sheet.nrows
            
            admin_fees_list = find_all_admin_fees_around_header(sheet, hp_row, search_rows=10)
            if not admin_fees_list:
                continue
            
            admin_fee_rate = get_applicable_admin_fee(admin_fees_list, month, year)
            if admin_fee_rate == 0:
                continue
            
            date_col = hp_col
            hours_col = hp_col + 1
            project_hours = 0.0
            project_fee = 0.0
            
            for r in range(hp_row + 1, next_header_row):
                period_text = sheet.cell_value(r, date_col)
                if not period_text:
                    continue
                period_str = str(period_text).strip().lower()
                if any(x in period_str for x in ['total', 'deductions', 'gross', 'taxes', 'buffer', 'net', 'ot rate', 'rate', 'employee']):
                    continue
                date_info = extract_date_from_period(period_str)
                if not date_info:
                    continue
                period_month, period_day, period_year = date_info
                if period_month != month or period_year != year:
                    continue
                if hours_col < sheet.ncols:
                    hours_val = sheet.cell_value(r, hours_col)
                    hours = safe_float(hours_val)
                    if hours > 0:
                        project_hours += hours
                        fee_for_period = hours * admin_fee_rate
                        project_fee += fee_for_period
            
            if project_hours > 0:
                total_hours += project_hours
                total_admin_fee += project_fee
                last_rate = admin_fee_rate
        
        return float(total_hours), float(last_rate), float(total_admin_fee)
    except Exception as e:
        return 0.0, 0.0, 0.0

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Universal Admin Fee Calculator")
    parser.add_argument("--paysheet", required=True, help="Path to paysheet file")
    parser.add_argument("--month", required=True, choices=MONTHS, help="Month name")
    parser.add_argument("--year", type=int, required=True, help="Year")
    parser.add_argument("--debug", action="store_true", help="Debug output")
    args = parser.parse_args()
    month_idx = MONTHS.index(args.month) + 1
    hours, rate, fee = calculate_admin_fee_for_paysheet(
        paysheet_path=args.paysheet,
        month=month_idx,
        year=args.year,
        debug=args.debug
    )
    print(f"Total Hours: {float(hours)}")
    print(f"Rate Used: ${float(rate):.2f}")
    print(f"Total Admin Fee: ${float(fee):.2f}")
