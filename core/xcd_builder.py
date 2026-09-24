"""
XCD Input File Builder for Cashfree, Easebuzz, and Airtel.
Generates individual .xlsx files formatted to match the sample specifications:
- Title: SALE TRANSACTIONS
- Columns: SL NO, SwinkPay Transaction ID, Amount, Settlement Date
- Sequential 1-based SL NO
- Gross CMS Transaction Amount
- Valid Settlement Date
"""
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from typing import List, Dict, Any, Optional
from datetime import datetime, date

from .normalizer import clean_key, clean_amount, extract_iso_date


def format_settlement_date_display(settle_date: Any) -> str:
    """Standardizes settlement date to DD-MM-YYYY display string matching samples."""
    if not settle_date:
        return ""
    if isinstance(settle_date, (datetime, date)):
        return settle_date.strftime("%d-%m-%Y")

    s = str(settle_date).strip()
    if not s or s == "--" or s.lower() in ("null", "none"):
        return ""

    # Try matching YYYY-MM-DD
    try:
        dt = datetime.strptime(s[:10], "%Y-%m-%d")
        return dt.strftime("%d-%m-%Y")
    except ValueError:
        pass

    # Try DD-MM-YYYY
    try:
        dt = datetime.strptime(s[:10], "%d-%m-%Y")
        return dt.strftime("%d-%m-%Y")
    except ValueError:
        pass

    # Try DD-b-YYYY e.g. 11-SEP-2026
    try:
        dt = datetime.strptime(s[:11], "%d-%b-%Y")
        return dt.strftime("%d-%m-%Y")
    except ValueError:
        pass

    return s


def extract_record_date(record: Dict[str, Any]) -> Optional[str]:
    """Extracts ISO date (YYYY-MM-DD) from a transaction record."""
    for k in (
        "Transaction Date & Time",
        "Transaction Date",
        "Txn Registered Date and Time",
        "Created At",
        "Date",
        "Payment Time",
        "Settlement Date",
    ):
        val = record.get(k)
        if val:
            iso = extract_iso_date(val)
            if iso:
                return iso
    return None


def extract_record_amounts(
    record: Dict[str, Any],
    net_lookup: Optional[Dict[str, float]] = None
) -> tuple[float, float]:
    """
    Extracts (gross_amount, net_amount) for a matched transaction record.
    Uses explicit columns (Transaction Amount, Settlement Amount, Net Amount),
    optional net_lookup dictionary, or fallback fee deduction (UPI 1.18%, Card/Rupay 3.245%).
    """
    gross = clean_amount(
        record.get("Transaction Amount") or
        record.get("Amount") or
        record.get("CMS Amount") or
        record.get("Original Input Amt")
    ) or 0.0

    net = clean_amount(
        record.get("Settlement Amount") or
        record.get("Net Amount") or
        record.get("Net Amount Payable(CR)")
    )

    if net is None and net_lookup:
        sp_id = str(record.get("SwinkPay Txn ID") or record.get("SwinkPay Transaction ID") or "").strip()
        if sp_id and sp_id in net_lookup:
            net = net_lookup[sp_id]
        else:
            rrn = str(record.get("RRN/UTR") or record.get("RRN") or record.get("UTR No.") or "").strip()
            if rrn and rrn in net_lookup:
                net = net_lookup[rrn]

    if net is None:
        netw = str(record.get("Network") or record.get("Payment Mode") or "").lower()
        if "credit" in netw or "cc" in netw or "rupay" in netw:
            net = round(gross * (1.0 - 0.03245), 2)
        else:
            net = round(gross * (1.0 - 0.0118), 2)

    return round(gross, 2), round(net, 2)


def get_distinct_dates_with_counts(
    records_cf: List[Dict[str, Any]],
    records_eb: List[Dict[str, Any]],
    records_air: List[Dict[str, Any]],
    net_lookup: Optional[Dict[str, float]] = None
) -> List[Dict[str, Any]]:
    """
    Computes distinct transaction dates across all matched partner records,
    along with day of week, transaction counts, gross transaction amount,
    and net settlement amount per partner and in total, plus default next-day settlement date.
    """
    date_map: Dict[str, Dict[str, Dict[str, float]]] = {}

    def _tally(records: List[Dict[str, Any]], partner_key: str):
        for r in records:
            d = extract_record_date(r)
            if not d:
                d = "Unknown"
            if d not in date_map:
                date_map[d] = {
                    "cf": {"count": 0, "gross": 0.0, "net": 0.0},
                    "eb": {"count": 0, "gross": 0.0, "net": 0.0},
                    "airtel": {"count": 0, "gross": 0.0, "net": 0.0},
                }
            g, n = extract_record_amounts(r, net_lookup)
            date_map[d][partner_key]["count"] += 1
            date_map[d][partner_key]["gross"] += g
            date_map[d][partner_key]["net"] += n

    _tally(records_cf, "cf")
    _tally(records_eb, "eb")
    _tally(records_air, "airtel")

    # Sort dates chronologically
    sorted_dates = sorted([d for d in date_map.keys() if d != "Unknown"])
    if "Unknown" in date_map:
        sorted_dates.append("Unknown")

    from datetime import timedelta
    result = []
    for d_str in sorted_dates:
        data = date_map[d_str]
        day_name = ""
        default_settle = ""
        if d_str != "Unknown":
            try:
                dt = datetime.strptime(d_str, "%Y-%m-%d")
                day_name = dt.strftime("%A")  # Friday, Saturday, Sunday etc.
                next_day = dt + timedelta(days=1)
                default_settle = next_day.strftime("%Y-%m-%d")
            except Exception:
                pass

        cf_c = int(data["cf"]["count"])
        cf_g = round(data["cf"]["gross"], 2)
        cf_n = round(data["cf"]["net"], 2)

        eb_c = int(data["eb"]["count"])
        eb_g = round(data["eb"]["gross"], 2)
        eb_n = round(data["eb"]["net"], 2)

        air_c = int(data["airtel"]["count"])
        air_g = round(data["airtel"]["gross"], 2)
        air_n = round(data["airtel"]["net"], 2)

        tot_c = cf_c + eb_c + air_c
        tot_g = round(cf_g + eb_g + air_g, 2)
        tot_n = round(cf_n + eb_n + air_n, 2)

        result.append({
            "date": d_str,
            "day_name": day_name,
            "cf_count": cf_c,
            "cf_gross": cf_g,
            "cf_net": cf_n,
            "eb_count": eb_c,
            "eb_gross": eb_g,
            "eb_net": eb_n,
            "airtel_count": air_c,
            "airtel_gross": air_g,
            "airtel_net": air_n,
            "total_count": tot_c,
            "total_gross": tot_g,
            "total_net": tot_n,
            "default_settlement_date": default_settle
        })

    return result


def build_xcd_workbook(
    records: List[Dict[str, Any]],
    settlement_date: str,
    output_path: str,
    refund_records: Optional[List[Dict[str, Any]]] = None,
    chargeback_records: Optional[List[Dict[str, Any]]] = None,
    selected_dates: Optional[List[str]] = None,
    date_settlement_map: Optional[Dict[str, str]] = None
) -> str:
    """
    Creates an XCD input workbook (.xlsx) matching the user's exact format:
    Row 1: SALE TRANSACTIONS (bold, Calibri 11pt, center aligned in col A)
    Row 2: Headers: SL NO (center), SwinkPay Transaction ID (left), Amount (right), Settlement Date (left)
    Row 3+: Data rows with 1-based SL NO, SwinkPay ID, Amount, and Settlement Date.
    Trailing sections:
      - 1 blank row
      - REFUND TRANSACTIONS (bold, Calibri 11pt, left-aligned in col A)
      - Headers: SL NO (center), SwinkPay Refund Transaction ID (left), Amount (right), Refund Settlement Date (left)
      - Any refund records
      - 1 blank row
      - CHARGEBACK/ADJUSTMENT TRANSACTIONS (bold, Calibri 11pt, left-aligned in col A)
      - Headers: SL NO (center), SwinkPay Transaction ID (left), Amount (right), Chargeback Settlement Date (left)
      - Any chargeback records
    There are strictly NO cell borders on any cells.
    """
    # Filter records by selected_dates if provided
    if selected_dates is not None:
        selected_set = set(selected_dates)
        filtered_records = []
        for r in records:
            d = extract_record_date(r)
            if d in selected_set or (not d and "Unknown" in selected_set):
                filtered_records.append(r)
        records = filtered_records

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Styles (NO borders anywhere)
    font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
    font_data = Font(name="Calibri", size=11, bold=False, color="000000")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Row 1: Title
    c1 = ws.cell(1, 1, "SALE TRANSACTIONS")
    c1.font = font_bold
    c1.alignment = align_center

    # Row 2: Headers
    headers = ["SL NO", "SwinkPay Transaction ID", "Amount", "Settlement Date"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(2, col_idx, h)
        cell.font = font_bold
        if col_idx == 1:
            cell.alignment = align_center
        elif col_idx == 3:
            cell.alignment = align_right
        else:
            cell.alignment = align_left

    # Formatted settlement date display string default
    settle_display_default = format_settlement_date_display(settlement_date)

    # Row 3+: Sale Data rows
    curr_row = 2
    for r in records:
        curr_row += 1
        sl_no = curr_row - 2  # 1-based serial number

        sp_id = str(r.get("SwinkPay Txn ID") or r.get("SwinkPay Transaction ID") or "").strip()
        amt_raw = clean_amount(r.get("Transaction Amount") or r.get("Amount") or r.get("CMS Amount"))

        # Determine row-specific settlement date
        row_settle = settle_display_default
        if date_settlement_map:
            d = extract_record_date(r)
            if d and d in date_settlement_map:
                row_settle = format_settlement_date_display(date_settlement_map[d])

        # SL NO
        c_sl = ws.cell(curr_row, 1, sl_no)
        c_sl.font = font_data
        c_sl.alignment = align_center
        c_sl.number_format = "0"

        # SwinkPay Transaction ID
        c_sp = ws.cell(curr_row, 2, sp_id)
        c_sp.font = font_data
        c_sp.alignment = align_left
        c_sp.number_format = "@"

        # Amount (Gross CMS amount)
        if amt_raw is not None:
            c_amt = ws.cell(curr_row, 3, amt_raw)
            c_amt.number_format = "#,##0.00" if not float(amt_raw).is_integer() else "0"
        else:
            c_amt = ws.cell(curr_row, 3, "")
        c_amt.font = font_data
        c_amt.alignment = align_right

        # Settlement Date
        c_dt = ws.cell(curr_row, 4, row_settle)
        c_dt.font = font_data
        c_dt.alignment = align_right
        c_dt.number_format = "@"

    # Trailing Section 1: REFUND TRANSACTIONS
    curr_row += 1  # 1 Blank row

    curr_row += 1
    c_ref_title = ws.cell(curr_row, 1, "REFUND TRANSACTIONS")
    c_ref_title.font = font_bold
    c_ref_title.alignment = align_left

    curr_row += 1
    ref_headers = ["SL NO", "SwinkPay Refund Transaction ID", "Amount", "Refund Settlement Date"]
    for col_idx, h in enumerate(ref_headers, start=1):
        cell = ws.cell(curr_row, col_idx, h)
        cell.font = font_bold
        if col_idx == 1:
            cell.alignment = align_center
        elif col_idx == 3:
            cell.alignment = align_right
        else:
            cell.alignment = align_left

    if refund_records:
        for ref_sl, r in enumerate(refund_records, start=1):
            curr_row += 1
            sp_id = str(r.get("SwinkPay Refund Transaction ID") or r.get("SwinkPay Txn ID") or "").strip()
            amt_raw = clean_amount(r.get("Amount") or r.get("Transaction Amount"))
            s_dt = format_settlement_date_display(r.get("Refund Settlement Date") or settlement_date)

            c_sl = ws.cell(curr_row, 1, ref_sl)
            c_sl.font = font_data
            c_sl.alignment = align_center
            c_sl.number_format = "0"

            c_sp = ws.cell(curr_row, 2, sp_id)
            c_sp.font = font_data
            c_sp.alignment = align_left
            c_sp.number_format = "@"

            if amt_raw is not None:
                c_amt = ws.cell(curr_row, 3, amt_raw)
                c_amt.number_format = "#,##0.00" if not float(amt_raw).is_integer() else "0"
            else:
                c_amt = ws.cell(curr_row, 3, "")
            c_amt.font = font_data
            c_amt.alignment = align_right

            c_dt = ws.cell(curr_row, 4, s_dt)
            c_dt.font = font_data
            c_dt.alignment = align_right
            c_dt.number_format = "@"

    # Trailing Section 2: CHARGEBACK/ADJUSTMENT TRANSACTIONS
    curr_row += 1  # 1 Blank row

    curr_row += 1
    c_cb_title = ws.cell(curr_row, 1, "CHARGEBACK/ADJUSTMENT TRANSACTIONS")
    c_cb_title.font = font_bold
    c_cb_title.alignment = align_left

    curr_row += 1
    cb_headers = ["SL NO", "SwinkPay Transaction ID", "Amount", "Chargeback Settlement Date"]
    for col_idx, h in enumerate(cb_headers, start=1):
        cell = ws.cell(curr_row, col_idx, h)
        cell.font = font_bold
        if col_idx == 1:
            cell.alignment = align_center
        elif col_idx == 3:
            cell.alignment = align_right
        else:
            cell.alignment = align_left

    if chargeback_records:
        for cb_sl, r in enumerate(chargeback_records, start=1):
            curr_row += 1
            sp_id = str(r.get("SwinkPay Transaction ID") or r.get("SwinkPay Txn ID") or "").strip()
            amt_raw = clean_amount(r.get("Amount") or r.get("Transaction Amount"))
            s_dt = format_settlement_date_display(r.get("Chargeback Settlement Date") or settlement_date)

            c_sl = ws.cell(curr_row, 1, cb_sl)
            c_sl.font = font_data
            c_sl.alignment = align_center
            c_sl.number_format = "0"

            c_sp = ws.cell(curr_row, 2, sp_id)
            c_sp.font = font_data
            c_sp.alignment = align_left
            c_sp.number_format = "@"

            if amt_raw is not None:
                c_amt = ws.cell(curr_row, 3, amt_raw)
                c_amt.number_format = "#,##0.00" if not float(amt_raw).is_integer() else "0"
            else:
                c_amt = ws.cell(curr_row, 3, "")
            c_amt.font = font_data
            c_amt.alignment = align_right

            c_dt = ws.cell(curr_row, 4, s_dt)
            c_dt.font = font_data
            c_dt.alignment = align_right
            c_dt.number_format = "@"

    # Column dimensions
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 28

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path


def resolve_settlement_date(engine, user_settlement_date: Optional[str] = None) -> str:
    """
    Resolves the settlement date strictly:
    1. If user provided an explicit date control from dashboard, use it.
    2. Extract from uploaded Airtel Settlement report.
    3. Extract from SMMS Settlement Date column if populated.
    4. If none available, returns empty string (never guesses or hardcodes).
    """
    if user_settlement_date and str(user_settlement_date).strip():
        return str(user_settlement_date).strip()

    # Try Airtel Settlement report
    if engine.settle_report and engine.settle_report.records:
        for r in engine.settle_report.records:
            sd = r.get("Settlement Date") or r.get("SETTLEMENT_DATE")
            if sd and str(sd).strip() and str(sd).strip() != "--":
                return str(sd).strip()

    # Try SMMS Settlement Date
    if engine.smms_report and engine.smms_report.records:
        for r in engine.smms_report.records:
            sd = r.get("Settlement Date & Time") or r.get("Settlement Date")
            if sd and str(sd).strip() and str(sd).strip() != "--":
                iso_d = extract_iso_date(sd)
                if iso_d:
                    return iso_d

    # Try CMS transaction dates
    if engine.cms_report and engine.cms_report.records:
        for r in engine.cms_report.records:
            td = r.get("Transaction Date & Time") or r.get("Transaction Date")
            if td and str(td).strip() and str(td).strip() != "--":
                iso_d = extract_iso_date(td)
                if iso_d:
                    return iso_d

    # Try any matched transaction dates
    for tab_recs in [getattr(engine, "cms_cf_matched", []), getattr(engine, "cms_eb_matched", []), getattr(engine, "cms_air_matched", [])]:
        for r in tab_recs:
            td = r.get("Transaction Date & Time") or r.get("Transaction Date") or r.get("CMS Date") or r.get("Partner Date")
            if td and str(td).strip() and str(td).strip() != "--":
                iso_d = extract_iso_date(td)
                if iso_d:
                    return iso_d

    # Fallback to current date
    return datetime.now().strftime("%Y-%m-%d")


def generate_partner_xcd_files(
    engine,
    output_dir: str,
    recon_date_str: str,
    settlement_date: str
) -> Dict[str, Dict[str, Any]]:
    """
    Generates the three partner XCD files from matched tabs:
    - Cashfree: from cms_cf_matched -> XCD Input file as on YYYY-MM-DD (Cf).xlsx
    - Easebuzz: from cms_eb_matched -> XCD Input file as on YYYY-MM-DD (EB).xlsx
    - Airtel: from cms_air_matched   -> XCD Input file as on YYYY-MM-DD (Airtel).xlsx
    """
    cf_filename = f"XCD Input file as on {recon_date_str} (Cf).xlsx"
    eb_filename = f"XCD Input file as on {recon_date_str} (EB).xlsx"
    air_filename = f"XCD Input file as on {recon_date_str} (Airtel).xlsx"

    cf_path = os.path.join(output_dir, cf_filename)
    eb_path = os.path.join(output_dir, eb_filename)
    air_path = os.path.join(output_dir, air_filename)

    # Build workbooks
    build_xcd_workbook(engine.cms_cf_matched, settlement_date, cf_path)
    build_xcd_workbook(engine.cms_eb_matched, settlement_date, eb_path)
    build_xcd_workbook(engine.cms_air_matched, settlement_date, air_path)

    # Compute totals for verification
    def calc_stats(recs):
        cnt = len(recs)
        gross_amt = 0.0
        net_amt = 0.0
        for r in recs:
            g, n = extract_record_amounts(r)
            gross_amt += g
            net_amt += n
        return cnt, round(gross_amt, 2), round(net_amt, 2)

    cf_cnt, cf_gross, cf_net = calc_stats(engine.cms_cf_matched)
    eb_cnt, eb_gross, eb_net = calc_stats(engine.cms_eb_matched)
    air_cnt, air_gross, air_net = calc_stats(engine.cms_air_matched)

    return {
        "cashfree": {
            "partner": "CashFree",
            "filename": cf_filename,
            "filepath": cf_path,
            "count": cf_cnt,
            "gross_amount": cf_gross,
            "net_amount": cf_net,
            "settlement_date": format_settlement_date_display(settlement_date)
        },
        "easebuzz": {
            "partner": "EaseBuzz",
            "filename": eb_filename,
            "filepath": eb_path,
            "count": eb_cnt,
            "gross_amount": eb_gross,
            "net_amount": eb_net,
            "settlement_date": format_settlement_date_display(settlement_date)
        },
        "airtel": {
            "partner": "Airtel Bank",
            "filename": air_filename,
            "filepath": air_path,
            "count": air_cnt,
            "gross_amount": air_gross,
            "net_amount": air_net,
            "settlement_date": format_settlement_date_display(settlement_date)
        }
    }
