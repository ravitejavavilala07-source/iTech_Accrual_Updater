#!/usr/bin/env python3
"""
Wrapper: Runs accrual_updater_system4.py + admin_fee_module.py
Without modifying the original accrual_updater_system4.py
"""

import argparse
import sys
from accrual_updater import AccrualUpdater
from admin_fee_module import calculate_admin_fee_for_paysheet

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Accrual Updater + Admin Fee")
    
    parser.add_argument("--master", required=True, help="Path to master file")
    parser.add_argument("--paysheets", required=True, help="Paysheets folder")
    parser.add_argument("--month", required=True, choices=MONTHS, help="Month")
    parser.add_argument("--year", type=int, required=True, help="Year")
    parser.add_argument("--dates", nargs="+", help="Dates for AB calc")
    parser.add_argument("--multipliers", nargs="+", type=float, help="Multipliers")
    parser.add_argument("--admin-fee", action="store_true", help="Calculate admin fee")
    parser.add_argument("--dry-run", action="store_true", help="Dry run only")
    parser.add_argument("--backup", action="store_true", help="Create backup")
    
    args = parser.parse_args()
    
    # Build date-multiplier pairs
    date_multiplier_pairs = []
    if args.dates and args.multipliers:
        if len(args.dates) == len(args.multipliers):
            date_multiplier_pairs = list(zip(args.dates, args.multipliers))
        else:
            print("ERROR: Dates and multipliers must match")
            sys.exit(1)
    
    # Run main accrual updater (unchanged)
    print("\n" + "="*80)
    print("STEP 1: Running Accrual Updater")
    print("="*80)
    
    updater = AccrualUpdater(
        master_path=args.master,
        month=args.month,
        year=args.year,
        paysheets_folder=args.paysheets,
        date_multiplier_pairs=date_multiplier_pairs,
        dry_run=args.dry_run,
        backup=args.backup,
    )
    
    result = updater.process()
    
    # Calculate admin fees (optional)
    if args.admin_fee:
        print("\n" + "="*80)
        print("STEP 2: Calculating Admin Fees")
        print("="*80)
        
        import os
        month_idx = MONTHS.index(args.month) + 1
        
        # Get list of paysheet files
        valid_exts = {'.xls', '.xlsx', '.xlsm'}
        files = []
        for root, _, fnames in os.walk(args.paysheets):
            for fn in fnames:
                if os.path.splitext(fn)[1].lower() in valid_exts:
                    files.append(os.path.join(root, fn))
        
        print(f"\nProcessing {len(files)} paysheet(s) for admin fees...\n")
        
        for idx, fp in enumerate(files, 1):
            fname = os.path.basename(fp)
            print(f"[{idx}/{len(files)}] {fname}")
            
            total_hours, total_admin_fee, breakdown = calculate_admin_fee_for_paysheet(
                paysheet_path=fp,
                month=month_idx,
                year=args.year,
                debug=True
            )
            
            if total_admin_fee > 0:
                print(f"  ✓ Admin Fee: ${total_admin_fee:.2f}\n")
            else:
                print(f"  → No admin fee calculated\n")
    
    print("\n" + "="*80)
    print("COMPLETE")
    print("="*80)
    print(f"Result: {result}")







