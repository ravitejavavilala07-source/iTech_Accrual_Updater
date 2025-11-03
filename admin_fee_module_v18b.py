#!/usr/bin/env python3
import os, re, sys, argparse
from typing import Any, Optional, Tuple, List

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

def extract_period_dates(period_str: str) -> Optional[Tuple[Tuple[int, int, int], Tuple[int, int, int]]]:
    match = re.match(r'(\d{1,2})/(\d{1,2})\s*[-/]\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})', period_str)
    if match:
        return ((int(match.group(1)), int(match.group(2)), int(match.group(5))), 
                (int(match.group(3)), int(match.group(4)), int(match.group(5))))
    return None

def extract_single_date(date_str: str) -> Optional[Tuple[int, int, int]]:
    """Extract single date from MM/DD/YYYY format (for Hours & Payment style)"""
    match = re.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', date_str)
    if match:
        month, day, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if year < 100: year += 2000
        return (month, day, year)
    return None

def parse_admin_fee_eff_date(text: str) -> Optional[Tuple[int, int, int]]:
    match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', text)
    if match:
        month, day, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if year < 100: year += 2000
        return (month, day, year)
    return None

def find_admin_fee_eff(sheet, header_row: int, search_rows: int = 10) -> list:
    """Find ALL 'Admin Fee Eff' entries - search ALL columns"""
    admin_fees = []
    search_start = max(0, header_row - search_rows)
    
    for r in range(search_start, header_row):
        for c in range(sheet.ncols):
            cell_val = sheet.cell_value(r, c)
            if cell_val and "admin fee eff" in str(cell_val).lower():
                date_tuple = parse_admin_fee_eff_date(str(cell_val))
                if date_tuple and c + 1 < sheet.ncols:
                    rate = safe_float(sheet.cell_value(r, c + 1))
                    if rate > 0:
                        admin_fees.append((date_tuple, rate))
    
    admin_fees.sort(key=lambda x: (x[0][2], x[0][0], x[0][1]))
    return admin_fees

def find_static_admin_fee(sheet, header_row: int, search_rows: int = 10) -> float:
    """Find static 'Admin Fee' - search ALL columns"""
    search_start = max(0, header_row - search_rows)
    
    for r in range(search_start, header_row):
        for c in range(sheet.ncols):
            cell_val = sheet.cell_value(r, c)
            if cell_val:
                text = str(cell_val).lower().strip()
                if text in ('admin fee', 'adminfee') and 'eff' not in text:
                    if c + 1 < sheet.ncols:
                        rate = safe_float(sheet.cell_value(r, c + 1))
                        if rate > 0:
                            return rate
    return 0.0

def get_rate_for_date(admin_fees: list, period_date: Tuple[int, int, int]) -> float:
    """Get applicable rate for a specific date"""
    applicable_rate = 0.0
    for date_tuple, rate in admin_fees:
        eff_comp = (date_tuple[2], date_tuple[0], date_tuple[1])
        period_comp = (period_date[2], period_date[0], period_date[1])
        if eff_comp <= period_comp:
            applicable_rate = rate
        else:
            break
    return applicable_rate

def is_full_month_period(start_date: Tuple[int, int, int], end_date: Tuple[int, int, int]) -> bool:
    """Check if period is a full month (starts on 1st, ends on 28+)"""
    if start_date[0] != end_date[0]:
        return False
    if start_date[1] != 1:
        return False
    if end_date[1] < 28:
        return False
    return True

def has_mid_month_eff(admin_fees_list: list, month: int) -> bool:
    """Check if Admin Fee Eff is mid-month (after the 1st)"""
    for date_tuple, rate in admin_fees_list:
        if date_tuple[0] == month and date_tuple[1] > 1:
            return True
    return False

def get_full_month_eff_rate(admin_fees_list: list, month: int) -> float:
    """Get the Eff rate for full month periods (latest rate in month after 1st)"""
    latest_rate = 0.0
    for date_tuple, rate in admin_fees_list:
        if date_tuple[0] == month and date_tuple[1] > 1:
            latest_rate = rate
    return latest_rate

def calculate_admin_fee_for_paysheet(paysheet_path: str, month: int, year: int, debug: bool = False) -> Tuple[float, float, float]:
    """
    v18b UPDATED - Now handles BOTH:
    1. "Work Period" header (period ranges: MM/DD-MM/DD/YYYY)
    2. "Hours & Payment" header (single dates: MM/DD/YYYY with amounts)
    """
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
        
        # ✅ UPDATED: Find BOTH "Work Period" AND "Hours & Payment" headers
        headers = []
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                val = sheet.cell_value(r, c)
                if val:
                    text = str(val).lower().strip()
                    if text == 'work period' or text == 'hours & payment':
                        headers.append((r, c, text))
                        break
        
        if not headers:
            return 0.0, 0.0, 0.0
        
        for section_idx, (hp_row, hp_col, header_type) in enumerate(headers):
            section_end = sheet.nrows
            
            for r in range(hp_row + 1, sheet.nrows):
                cell_val = sheet.cell_value(r, hp_col)
                if not cell_val:
                    continue
                cell_str = str(cell_val).strip().lower()
                
                if any(x in cell_str for x in ['total', 'admin fees', 'deductions']):
                    section_end = r
                    break
                
                if section_idx < len(headers) - 1:
                    next_hp_row = headers[section_idx + 1][0]
                    if r >= next_hp_row - 5:
                        section_end = r
                        break
            
            if debug:
                print(f"\n=== SECTION {section_idx + 1} ({header_type}) ===")
                print(f"Header at Row {hp_row}, Col {hp_col}")
                print(f"Section ends at Row {section_end}")
            
            admin_fees_list = find_admin_fee_eff(sheet, hp_row, search_rows=10)
            static_rate = find_static_admin_fee(sheet, hp_row, search_rows=10)
            
            if admin_fees_list and static_rate > 0:
                admin_fees_list = [((1, 1, 2025), static_rate)] + admin_fees_list
                if debug:
                    print(f"Found static Admin Fee ${static_rate:.2f} + Eff entries: {admin_fees_list}")
            elif admin_fees_list and debug:
                print(f"Found Admin Fee Eff entries: {admin_fees_list}")
            
            if admin_fees_list:
                section_hours = 0.0
                section_fee = 0.0
                section_rate = 0.0
                
                # Check if there's a mid-month Eff date
                has_mid_month = has_mid_month_eff(admin_fees_list, month)
                full_month_eff_rate = get_full_month_eff_rate(admin_fees_list, month)
                
                for r in range(hp_row + 1, section_end):
                    period_text = sheet.cell_value(r, hp_col)
                    if not period_text:
                        continue
                    period_str = str(period_text).strip().lower()
                    
                    # ✅ UPDATED: Handle BOTH period formats
                    period_dates = extract_period_dates(period_str)
                    single_date = None
                    
                    if not period_dates:
                        single_date = extract_single_date(period_str)
                    
                    # =========== WORK PERIOD LOGIC (existing) ===========
                    if period_dates:
                        start_date, end_date = period_dates
                        if start_date[0] != month or start_date[2] != year:
                            continue
                        
                        hours_val = safe_float(sheet.cell_value(r, hp_col + 1))
                        if hours_val <= 0:
                            continue
                        
                        # ✅ CORRECTED LOGIC:
                        if is_full_month_period(start_date, end_date) and has_mid_month:
                            # Full month with mid-month Eff: use Eff rate for entire month
                            rate = full_month_eff_rate
                            fee = hours_val * rate if rate > 0 else 0.0
                            if debug:
                                print(f"    Full month (use Eff rate ${{rate:.2f}}): {hours_val} hrs × ${rate:.2f} = ${fee:.2f}")
                        else:
                            # Weekly/partial or no mid-month Eff: use period start date
                            rate = get_rate_for_date(admin_fees_list, start_date)
                            fee = hours_val * rate if rate > 0 else 0.0
                            if debug:
                                print(f"    Period (start date): {hours_val} hrs × ${rate:.2f} = ${fee:.2f}")
                        
                        if fee > 0:
                            section_hours += hours_val
                            section_fee += fee
                            section_rate = rate
                            if debug:
                                print(f"  Row {r}: {period_str} | {hours_val} hrs | ${fee:.2f}")
                    
                    # ✅ NEW: HOURS & PAYMENT LOGIC (single dates)
                    elif single_date:
                        if single_date[0] != month or single_date[2] != year:
                            continue
                        
                        # For "Hours & Payment": column +1 is the AMOUNT (not hours count)
                        # Amount is already the billable/fee value
                        amount_val = safe_float(sheet.cell_value(r, hp_col + 1))
                        
                        if amount_val <= 0:
                            continue
                        
                        # Get rate for this date to determine if it applies
                        rate = get_rate_for_date(admin_fees_list, single_date)
                        
                        if rate > 0:
                            # Use amount as the fee directly (already calculated in paysheet)
                            fee = amount_val
                            section_hours += 1  # Count as 1 entry (not actual hours)
                            section_fee += fee
                            section_rate = rate
                            if debug:
                                print(f"  Row {r}: {period_str} | Amount ${amount_val:.2f} | Rate ${rate:.2f}")
                
                if section_hours > 0:
                    total_hours += section_hours
                    total_admin_fee += section_fee
                    last_rate = section_rate
                    if debug:
                        print(f"✓ Section Total: {section_hours:.2f} hrs | ${section_fee:.2f}")
                continue
            
            if static_rate > 0:
                section_hours = 0.0
                
                for r in range(hp_row + 1, section_end):
                    period_text = sheet.cell_value(r, hp_col)
                    if not period_text:
                        continue
                    period_str = str(period_text).strip().lower()
                    
                    period_dates = extract_period_dates(period_str)
                    if not period_dates:
                        continue
                    
                    start_date = period_dates[0]
                    if start_date[0] != month or start_date[2] != year:
                        continue
                    
                    hours_val = safe_float(sheet.cell_value(r, hp_col + 1))
                    if hours_val > 0:
                        section_hours += hours_val
                
                if section_hours > 0:
                    section_fee = section_hours * static_rate
                    total_hours += section_hours
                    total_admin_fee += section_fee
                    last_rate = static_rate
                    if debug:
                        print(f"✓ Found static 'Admin Fee' (${static_rate:.2f}/hr) | {section_hours:.2f} hrs | ${section_fee:.2f}")
        
        if debug:
            print(f"\n=== TOTAL ===")
            print(f"Hours: {total_hours:.2f} | Last Rate: ${last_rate:.2f} | Fee: ${total_admin_fee:.2f}")
        
        return float(total_hours), float(last_rate), float(total_admin_fee)
    
    except Exception as e:
        if debug:
            print(f"ERROR: {e}")
        return 0.0, 0.0, 0.0

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Admin Fee v18b - Corrected logic with Hours & Payment support")
    parser.add_argument("--paysheet", required=True, help="Paysheet file path")
    parser.add_argument("--month", required=True, choices=MONTHS, help="Month name")
    parser.add_argument("--year", type=int, required=True, help="Year")
    parser.add_argument("--debug", action="store_true", help="Debug output")
    args = parser.parse_args()
    month_idx = MONTHS.index(args.month) + 1
    hours, rate, fee = calculate_admin_fee_for_paysheet(args.paysheet, month_idx, args.year, args.debug)
    print(f"Total Hours: {float(hours)}")
    print(f"Rate Used: ${float(rate):.2f}")
    print(f"Total Admin Fee: ${float(fee):.2f}")
