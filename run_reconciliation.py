"""
Ops_Auto: Daily Automated Transaction Reconciliation CLI
Usage:
    python run_reconciliation.py --input-dir <path_to_directory>
    or
    python run_reconciliation.py --cms <path> --smms <path> --cf <path> --eb <path> --airtel <path> [--airtel-settlement <path>]
"""
import os
import sys
import argparse
from typing import Optional, Dict, List
from datetime import datetime

from core.detector import ReportType, MissingColumnsException
from core.reader import read_report, RawReport
from core.matcher import ReconciliationEngine
from core.aggregator import Aggregator
from core.excel_builder import ExcelReportBuilder
from core.normalizer import extract_iso_date


def scan_input_directory(input_dir: str) -> Dict[ReportType, str]:
    """Scans directory and classifies files by header detection."""
    found_files: Dict[ReportType, str] = {}
    valid_extensions = ('.xlsx', '.xls', '.csv')

    print(f"[*] Scanning directory '{input_dir}' for source reports...")
    for entry in os.listdir(input_dir):
        full_path = os.path.join(input_dir, entry)
        if not os.path.isfile(full_path) or not entry.lower().endswith(valid_extensions):
            continue
        
        # Skip temporary Excel lock files (~$...)
        if entry.startswith("~$"):
            continue

        try:
            raw_rep = read_report(full_path)
            if raw_rep.report_type != ReportType.UNKNOWN:
                # Disambiguate if already found
                if raw_rep.report_type not in found_files:
                    found_files[raw_rep.report_type] = full_path
                    print(f"  [+] Detected {raw_rep.report_type.value:18}: {entry}")
                else:
                    # Choose newer or better match if multiple
                    print(f"  [?] Additional {raw_rep.report_type.value:18}: {entry} (using {os.path.basename(found_files[raw_rep.report_type])})")
        except MissingColumnsException as mce:
            print(f"  [-] File skipped ({mce.report_type.value}): {entry} - {mce}")
        except Exception:
            # Not a supported report, ignore
            continue

    return found_files


def run_pipeline(
    cms_path: str,
    smms_path: str,
    cf_path: Optional[str] = None,
    eb_path: Optional[str] = None,
    airtel_path: Optional[str] = None,
    settle_path: Optional[str] = None,
    output_dir: str = ".",
    output_name: Optional[str] = None,
    tolerance: float = 0.01,
) -> str:
    """Executes reconciliation workflow and exports formatted Excel report."""
    print("=" * 80)
    print("OPS_AUTO: DAILY TRANSACTION RECONCILIATION WORKFLOW")
    print("=" * 80)

    # 1. Ingest reports
    print("[1/5] Loading and validating source reports...")
    cms_rep = read_report(cms_path)
    print(f"  Loaded CMS: {len(cms_rep.records):,} records from {os.path.basename(cms_path)}")

    smms_rep = read_report(smms_path)
    print(f"  Loaded SMMS: {len(smms_rep.records):,} records from {os.path.basename(smms_path)}")

    cf_rep = read_report(cf_path) if cf_path else None
    if cf_rep:
        print(f"  Loaded Cashfree: {len(cf_rep.records):,} records from {os.path.basename(cf_path)}")

    eb_rep = read_report(eb_path) if eb_path else None
    if eb_rep:
        print(f"  Loaded Easebuzz: {len(eb_rep.records):,} records from {os.path.basename(eb_path)}")

    air_rep = read_report(airtel_path) if airtel_path else None
    if air_rep:
        print(f"  Loaded Airtel: {len(air_rep.records):,} records from {os.path.basename(airtel_path)}")

    settle_rep = read_report(settle_path) if settle_path else None
    if settle_rep:
        print(f"  Loaded Airtel Settlement: {len(settle_rep.records):,} records from {os.path.basename(settle_path)}")
    else:
        print("  Airtel Settlement Report: Not provided (settlement checks optional)")

    # 2. Reconcile
    print("\n[2/5] Performing multi-way matching...")
    engine = ReconciliationEngine(tolerance=tolerance)
    engine.set_reports(
        cms=cms_rep,
        smms=smms_rep,
        cf=cf_rep,
        eb=eb_rep,
        airtel=air_rep,
        settle=settle_rep
    )
    engine.run()

    # 3. Compute Summary & Audit Checks
    print("\n[3/5] Computing summary aggregations and difference-to-zero control checks...")
    aggregator = Aggregator(engine)
    aggregator.compute()

    # 4. Determine output filename
    if not output_name:
        # Determine dominant transaction date from SMMS report
        sample_dates = []
        for r in smms_rep.records[:50]:
            d = extract_iso_date(r.get("Transaction Date & Time"))
            if d:
                sample_dates.append(d)
        date_str = max(set(sample_dates), key=sample_dates.count) if sample_dates else datetime.now().strftime("%Y-%m-%d")
        output_name = f"Reconciliation_{date_str}.xlsx"

    os.makedirs(output_dir, exist_ok=True)
    final_output_path = os.path.join(output_dir, output_name)

    # 5. Build and export workbook
    print(f"\n[4/5] Building styled workbook '{output_name}'...")
    builder = ExcelReportBuilder(engine, aggregator)
    builder.build_and_save(final_output_path)
    print(f"  Successfully exported to: {os.path.abspath(final_output_path)}")

    # 6. Display Console Summary Report
    print("\n[5/5] RECONCILIATION SUMMARY REPORT:")
    print("-" * 80)
    print(f"Master Population (SMMS Successful): {aggregator.grand_total['success_count']:,} txns | Gross: Rs. {aggregator.grand_total['gross_amount']:,.2f} | Net: Rs. {aggregator.grand_total['net_amount']:,.2f}")
    print("-" * 80)
    print(f"{'Partner':15} | {'Success Count':13} | {'Gross Amount':15} | {'Net Amount':15} | {'Matched':10}")
    print("-" * 80)
    for p, sub in aggregator.partner_subtotals.items():
        print(f"{p:15} | {sub['success_count']:13,d} | Rs. {sub['gross_amount']:11,.2f} | Rs. {sub['net_amount']:11,.2f} | {sub['matched_count']:10,d}")
    print("-" * 80)
    print(f"{'TOTAL':15} | {aggregator.grand_total['success_count']:13,d} | Rs. {aggregator.grand_total['gross_amount']:11,.2f} | Rs. {aggregator.grand_total['net_amount']:11,.2f} | {aggregator.grand_total['matched_count']:10,d}")
    print("-" * 80)

    print("\nEXCEPTION & DISCREPANCY SUMMARY:")
    print(f"  - Failed or Reversed txns : {len(engine.failed_or_reversed):,}")
    print(f"  - CMS Not in SMMS         : {len(engine.cms_not_in_smms):,}")
    print(f"  - SMMS Not in CMS         : {len(engine.smms_not_in_cms):,}")
    print(f"  - Unmatched Cashfree      : {len(engine.unmatched_cf):,}")
    print(f"  - Unmatched Easebuzz      : {len(engine.unmatched_eb):,}")
    print(f"  - Unmatched Airtel        : {len(engine.unmatched_airtel):,}")
    print(f"  - Duplicate Keys Found    : {len(engine.duplicates):,}")
    print(f"  - Total Exceptions        : {len(engine.exceptions):,}")

    print("\nAUDIT CONTROL CHECKS (DIFFERENCE-TO-ZERO):")
    print("-" * 80)
    all_passed = True
    for chk in aggregator.control_checks:
        status_sym = "[PASS]" if chk["status"] == "PASS" else "[FAIL]"
        if chk["status"] != "PASS":
            all_passed = False
        print(f"  {status_sym:7} {chk['check_name']:50} Diff: {chk['difference']}")
    print("-" * 80)

    if all_passed:
        print("[SUCCESS] All difference-to-zero control checks PASSED perfectly!\n")
    else:
        print("[WARNING] One or more audit balance checks require attention. Review the 'Exceptions' tab in Excel.\n")

    return final_output_path


def main():
    parser = argparse.ArgumentParser(description="Automated Daily Transaction Reconciliation Tool")
    parser.add_argument("--input-dir", type=str, help="Directory containing daily reports to auto-detect")
    parser.add_argument("--cms", type=str, help="Path to CMS report (.xlsx, .xls)")
    parser.add_argument("--smms", type=str, help="Path to SMMS report (.xlsx)")
    parser.add_argument("--cf", type=str, help="Path to Cashfree report (.xlsx)")
    parser.add_argument("--eb", type=str, help="Path to Easebuzz report (.csv, .xlsx)")
    parser.add_argument("--airtel", type=str, help="Path to Airtel report (.csv, .xlsx)")
    parser.add_argument("--airtel-settlement", type=str, help="Path to optional Airtel settlement report (.csv, .xlsx)")
    parser.add_argument("--output-dir", type=str, default=".", help="Directory to save the completed reconciliation workbook")
    parser.add_argument("--output-name", type=str, help="Custom output workbook name (default: Reconciliation_YYYY-MM-DD.xlsx)")
    parser.add_argument("--tolerance", type=float, default=0.01, help="Amount matching tolerance in rupees (default: 0.01)")

    args = parser.parse_args()

    cms = args.cms
    smms = args.smms
    cf = args.cf
    eb = args.eb
    airtel = args.airtel
    settle = args.airtel_settlement

    if args.input_dir:
        detected = scan_input_directory(args.input_dir)
        cms = cms or detected.get(ReportType.CMS)
        smms = smms or detected.get(ReportType.SMMS)
        cf = cf or detected.get(ReportType.CASHFREE)
        eb = eb or detected.get(ReportType.EASEBUZZ)
        airtel = airtel or detected.get(ReportType.AIRTEL)
        settle = settle or detected.get(ReportType.AIRTEL_SETTLEMENT)

    if not cms or not smms:
        print("\n[ERROR] Both CMS and SMMS reports are required.")
        print("Please provide --cms and --smms or specify --input-dir where they can be detected.\n")
        parser.print_help()
        sys.exit(1)

    try:
        run_pipeline(
            cms_path=cms,
            smms_path=smms,
            cf_path=cf,
            eb_path=eb,
            airtel_path=airtel,
            settle_path=settle,
            output_dir=args.output_dir,
            output_name=args.output_name,
            tolerance=args.tolerance,
        )
    except MissingColumnsException as mce:
        print(f"\n[FATAL ERROR] {mce}")
        sys.exit(2)
    except Exception as e:
        print(f"\n[FATAL ERROR] Reconciliation workflow terminated: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(3)


if __name__ == "__main__":
    main()
