"""
Payment Gateway Missing Transaction Payload Builder.
Constructs Postman-ready API payloads for transactions present in
PG reports (Cashfree, Easebuzz, Airtel) but missing in CMS.

Target Endpoint (Postman):
POST https://merchants.swinkpay-fintech.com/api/v2/decision/updated
Payload Schema:
{
  "amount": "20.00",
  "terminalID": "X7VJND",
  "utr": "129761680150",
  "dateAndTime": "2026-09-17 15:59:29"
}
"""
import re
import json
from datetime import datetime
from typing import List, Dict, Any, Optional

from .normalizer import clean_amount, clean_key
from .terminal_mapper import (
    resolve_mms_terminal_id,
    resolve_terminal_details,
    extract_cf_middle_number,
    load_terminal_mappings,
    load_merchant_terminal_mappings
)


def format_payload_datetime(dt_val: Any) -> str:
    """Normalizes any date/time format into 'YYYY-MM-DD HH:MM:SS'."""
    if not dt_val:
        return ""
    if isinstance(dt_val, datetime):
        return dt_val.strftime("%Y-%m-%d %H:%M:%S")

    s = str(dt_val).strip()
    if not s or s.lower() in ("--", "null", "none"):
        return ""

    # Supported incoming formats
    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%d/%m/%y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d",
        "%d-%m-%Y",
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(s[:19], fmt)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except (ValueError, TypeError):
            continue

    return s


def format_payload_amount(amt_val: Any) -> str:
    """Formats amount string to 2 decimal places e.g. '20.00'."""
    amt = clean_amount(amt_val)
    if amt is None:
        return "0.00"
    return f"{float(amt):.2f}"


def format_payload_utr(utr_val: Any) -> str:
    """
    Normalizes UTR / RRN for SwinkPay Decision API payload.
    If the UTR is not of 12 digits, adds leading zeros to make it 12 digits.
    Example:
      '6789876543' -> '006789876543'
      'utr-6789876543' -> '006789876543'
      '129761680150' -> '129761680150'
    """
    if utr_val is None:
        return ""
    s = clean_key(utr_val)
    if not s or s == "--" or s.lower() in ("null", "none", "nan"):
        return ""

    # Remove common prefixes like 'utr-', 'utr:', 'utr ', 'rrn-', 'rrn:', 'rrn '
    s_clean = re.sub(r'^(utr|rrn)[-:\s_]*', '', s, flags=re.IGNORECASE).strip()
    if s_clean:
        s = s_clean

    # Remove any internal spaces or hyphens if numeric
    s_compact = re.sub(r'[\s\-]', '', s)
    if s_compact.isdigit():
        s = s_compact

    # If numeric and length < 12, pad with leading zeros to make it exactly 12 digits
    if s.isdigit() and len(s) < 12:
        s = s.zfill(12)

    return s


def build_cms_prefix_terminal_map(cms_report) -> Dict[str, str]:
    """
    Builds a lookup index from Cashfree Order ID prefix (e.g. '330595-4873')
    to the Merchant MMS Terminal ID (e.g. 'X1JYYO') by searching CMS records.
    """
    prefix_map: Dict[str, str] = {}
    if not cms_report or not getattr(cms_report, "records", None):
        return prefix_map

    for r in cms_report.records:
        mms_term = str(r.get("Merchant MMS Terminal ID") or "").strip()
        if not mms_term or mms_term == "--":
            continue

        # Check Auth Response JSON for order_id
        auth = str(r.get("Auth Response") or "")
        if "order_id" in auth:
            m = re.search(r'"order_id"\s*:\s*"([^"]+)"', auth)
            if m:
                oid = m.group(1).strip()
                parts = oid.split("-")
                if len(parts) >= 2:
                    pref = f"{parts[0]}-{parts[1]}"
                    if pref not in prefix_map:
                        prefix_map[pref] = mms_term

        # Check Invoice Number
        inv = str(r.get("Invoice Number") or "").strip()
        if inv and "-" in inv:
            parts = inv.split("-")
            if len(parts) >= 2:
                pref = f"{parts[0]}-{parts[1]}"
                if pref not in prefix_map:
                    prefix_map[pref] = mms_term

        # Check Partner Unique Txn ID
        p_tid = str(r.get("Partner Unique Txn ID") or "").strip()
        if p_tid and "-" in p_tid:
            parts = p_tid.split("-")
            if len(parts) >= 2:
                pref = f"{parts[0]}-{parts[1]}"
                if pref not in prefix_map:
                    prefix_map[pref] = mms_term

    return prefix_map


def build_missing_pg_payloads(
    engine,
    merchant_key: Optional[str] = None,
    terminal_mappings: Optional[Dict[str, Any]] = None
) -> List[Dict[str, Any]]:
    """
    Generates Postman/Pull-ready payload items for all unmatched PG transactions.
    Supports Cashfree, Easebuzz, and Airtel.
    Resolves Cashfree Order ID middle numbers (e.g. '4860' from '330595-4860-...')
    to MMS Terminal ID via merchant-specific terminal mappings or CMS report lookup.
    """
    if terminal_mappings is None:
        terminal_mappings = load_merchant_terminal_mappings(merchant_key)

    results: List[Dict[str, Any]] = []

    # Build fast lookup of all RRNs/UTRs/Txn IDs present in CMS report
    # If a transaction exists in CMS (regardless of SMMS Sync Status), it is already in CMS and MUST NOT be pulled!
    cms_keys = set()
    if getattr(engine, "cms_report", None) and getattr(engine.cms_report, "records", None):
        for cr in engine.cms_report.records:
            cr_rrn = clean_key(cr.get("RRN/UTR") or cr.get("RRN") or cr.get("UTR"))
            cr_sp = clean_key(cr.get("SwinkPay Txn ID") or cr.get("SwinkPay Transaction ID"))
            if cr_rrn:
                cms_keys.add(cr_rrn)
                if cr_rrn.isdigit():
                    cms_keys.add(cr_rrn.lstrip("0"))
                    cms_keys.add(cr_rrn.zfill(12))
            if cr_sp:
                cms_keys.add(cr_sp)

    def _is_in_cms(key: str) -> bool:
        if not key:
            return False
        k = clean_key(key)
        if not k:
            return False
        if k in cms_keys:
            return True
        if k.isdigit():
            if k.lstrip("0") in cms_keys or k.zfill(12) in cms_keys:
                return True
        return False

    # 1. Cashfree unmatched
    unmatched_cf = getattr(engine, "unmatched_cf", [])
    if unmatched_cf:
        cf_prefix_map = build_cms_prefix_terminal_map(getattr(engine, "cms_report", None))

        for cf_r in unmatched_cf:
            order_id = str(cf_r.get("Order Id") or cf_r.get("Matched Transaction ID") or "").strip()
            amount_str = format_payload_amount(cf_r.get("Amount"))
            raw_utr = str(cf_r.get("Bank Reference No.") or cf_r.get("UTR No.") or cf_r.get("Reference Id") or "").strip()
            utr_str = format_payload_utr(raw_utr)
            dt_str = format_payload_datetime(cf_r.get("Transaction Time") or cf_r.get("Transaction Date"))

            # If transaction is already in CMS (even if SMMS Sync Status is False), no need to pull!
            if _is_in_cms(raw_utr) or _is_in_cms(utr_str) or _is_in_cms(order_id):
                continue

            middle_number = extract_cf_middle_number(order_id)
            parts = order_id.split("-")
            prefix = f"{parts[0]}-{parts[1]}" if len(parts) >= 2 else order_id

            # Priority 1: Check merchant's Terminal Report mapping (Partner Ref ID / VPA middle number)
            term_details = resolve_terminal_details(order_id, merchant_key=merchant_key, mappings=terminal_mappings)
            term_id = term_details.get("mms_terminal_id") or term_details.get("terminal_id") if term_details else ""
            branch_name = term_details.get("branch_name") or "" if term_details else ""

            # Priority 2: Check CMS prefix map (e.g. '330595-4873')
            if not term_id and prefix in cf_prefix_map:
                term_id = cf_prefix_map[prefix]

            # Priority 3: Direct search in CMS report records
            if not term_id and getattr(engine, "cms_report", None):
                for cr in engine.cms_report.records:
                    cr_vals = str(list(cr.values()))
                    if (prefix and prefix in cr_vals) or (middle_number and middle_number in cr_vals):
                        cand = str(cr.get("Merchant MMS Terminal ID") or "").strip()
                        if cand and cand != "--":
                            term_id = cand
                            cf_prefix_map[prefix] = cand
                            break

            payload_dict = {
                "amount": amount_str,
                "terminalID": term_id,
                "utr": utr_str,
                "dateAndTime": dt_str
            }

            results.append({
                "gateway": "Cashfree",
                "order_id": order_id,
                "middle_number": middle_number,
                "terminal_id": term_id,
                "branch_name": branch_name,
                "utr": utr_str,
                "amount": amount_str,
                "date_and_time": dt_str,
                "payload": payload_dict,
                "payload_json": json.dumps(payload_dict, indent=2)
            })

    # 2. Easebuzz unmatched
    unmatched_eb = getattr(engine, "unmatched_eb", [])
    if unmatched_eb:
        for eb_r in unmatched_eb:
            order_id = str(eb_r.get("ID") or eb_r.get("UPI tid") or eb_r.get("Matched Transaction ID") or "").strip()
            amount_str = format_payload_amount(eb_r.get("Amount"))
            raw_utr = str(eb_r.get("UTR") or eb_r.get("UPI tid") or "").strip()
            utr_str = format_payload_utr(raw_utr)
            dt_str = format_payload_datetime(eb_r.get("Transaction Date") or eb_r.get("Date"))
            term_id = str(eb_r.get("Virtual Account Label") or "").strip()

            # If transaction is already in CMS, no need to pull!
            if _is_in_cms(raw_utr) or _is_in_cms(utr_str) or _is_in_cms(order_id):
                continue

            # If terminal ID is numeric, check if it maps to an MMS Terminal ID in terminal_mappings
            resolved_tid = resolve_mms_terminal_id(term_id, terminal_mappings)
            if resolved_tid:
                term_id = resolved_tid

            payload_dict = {
                "amount": amount_str,
                "terminalID": term_id,
                "utr": utr_str,
                "dateAndTime": dt_str
            }

            results.append({
                "gateway": "Easebuzz",
                "order_id": order_id,
                "middle_number": term_id,
                "terminal_id": term_id,
                "utr": utr_str,
                "amount": amount_str,
                "date_and_time": dt_str,
                "payload": payload_dict,
                "payload_json": json.dumps(payload_dict, indent=2)
            })

    # 3. Airtel unmatched
    unmatched_air = getattr(engine, "unmatched_airtel", [])
    if unmatched_air:
        for air_r in unmatched_air:
            order_id = str(air_r.get("Transaction Id") or air_r.get("Matched Transaction ID") or "").strip()
            amount_str = format_payload_amount(air_r.get("Original Input Amt") or air_r.get("Amount"))
            raw_utr = str(air_r.get("PARTNER_TXN_ID") or air_r.get("REF_TXN_NO_ORG") or order_id).strip()
            utr_str = format_payload_utr(raw_utr)
            dt_str = format_payload_datetime(air_r.get("Date and Time") or air_r.get("Transaction Date"))

            # If transaction is already in CMS, no need to pull!
            if _is_in_cms(raw_utr) or _is_in_cms(utr_str) or _is_in_cms(order_id):
                continue

            # Extract terminal from Transaction To e.g. spf-2045-xh2vuy@mairtel -> XH2VUY
            txn_to = str(air_r.get("Transaction To") or "").strip()
            term_id = ""
            if txn_to:
                clean_to = txn_to.split("@")[0].strip()
                term_parts = clean_to.split("-")
                term_id = term_parts[-1].strip().upper()

            resolved_tid = resolve_mms_terminal_id(term_id, terminal_mappings)
            if resolved_tid:
                term_id = resolved_tid

            payload_dict = {
                "amount": amount_str,
                "terminalID": term_id,
                "utr": utr_str,
                "dateAndTime": dt_str
            }

            results.append({
                "gateway": "Airtel",
                "order_id": order_id,
                "middle_number": term_id,
                "terminal_id": term_id,
                "utr": utr_str,
                "amount": amount_str,
                "date_and_time": dt_str,
                "payload": payload_dict,
                "payload_json": json.dumps(payload_dict, indent=2)
            })

    return results


def parse_pg_report_records_for_pull(
    raw_report: Any,
    merchant_key: Optional[str] = None,
    terminal_mappings: Optional[Dict[str, Any]] = None
) -> List[Dict[str, Any]]:
    """
    Parses any uploaded Payment Gateway report (Cashfree, Easebuzz, Airtel, or generic CSV/Excel)
    and constructs SwinkPay Decision API payloads for all transactions in one shot:
    {
      "amount": "20.00",
      "terminalID": "X7VJND",
      "utr": "129761680150",
      "dateAndTime": "2026-09-17 15:59:29"
    }
    """
    if terminal_mappings is None:
        terminal_mappings = load_merchant_terminal_mappings(merchant_key)

    records = getattr(raw_report, "records", []) or []
    rep_type = getattr(raw_report, "report_type", None)
    from .detector import ReportType

    results: List[Dict[str, Any]] = []

    for idx, r in enumerate(records):
        gateway = "Payment Gateway"
        order_id = ""
        utr_str = ""
        amount_str = "0.00"
        dt_str = ""
        term_id = ""
        branch_name = ""
        status_val = ""
        middle_number = ""

        # 1. CASHFREE
        if rep_type == ReportType.CASHFREE:
            gateway = "Cashfree"
            order_id = str(r.get("Order Id") or r.get("Reference Id") or "").strip()
            amount_str = format_payload_amount(r.get("Amount"))
            utr_str = format_payload_utr(r.get("Bank Reference No.") or r.get("UTR No.") or r.get("Reference Id") or "")
            dt_str = format_payload_datetime(r.get("Transaction Time") or r.get("Transaction Date"))
            status_val = str(r.get("Transaction Status") or r.get("Status") or "SUCCESS").strip()

            middle_number = extract_cf_middle_number(order_id)
            term_details = resolve_terminal_details(order_id, merchant_key=merchant_key, mappings=terminal_mappings)
            if not term_details and middle_number:
                term_details = resolve_terminal_details(middle_number, merchant_key=merchant_key, mappings=terminal_mappings)
            if term_details:
                term_id = term_details.get("mms_terminal_id") or term_details.get("terminal_id") or ""
                branch_name = term_details.get("branch_name") or ""
            if not term_id and middle_number:
                term_id = middle_number

        # 2. EASEBUZZ
        elif rep_type == ReportType.EASEBUZZ:
            gateway = "Easebuzz"
            order_id = str(r.get("ID") or r.get("UPI tid") or "").strip()
            amount_str = format_payload_amount(r.get("Amount"))
            utr_str = format_payload_utr(r.get("UTR") or r.get("UPI tid") or "")
            dt_str = format_payload_datetime(r.get("Transaction Date") or r.get("Date"))
            status_val = str(r.get("Status") or "Payment Received").strip()

            label_term = str(r.get("Virtual Account Label") or "").strip()
            term_details = resolve_terminal_details(label_term, merchant_key=merchant_key, mappings=terminal_mappings) if label_term else None
            if not term_details and order_id:
                term_details = resolve_terminal_details(order_id, merchant_key=merchant_key, mappings=terminal_mappings)
            if term_details:
                term_id = term_details.get("mms_terminal_id") or term_details.get("terminal_id") or ""
                branch_name = term_details.get("branch_name") or ""
            elif label_term:
                resolved_tid = resolve_mms_terminal_id(label_term, terminal_mappings)
                term_id = resolved_tid or label_term
            middle_number = term_id

        # 3. AIRTEL
        elif rep_type in (ReportType.AIRTEL, ReportType.AIRTEL_SETTLEMENT):
            gateway = "Airtel"
            order_id = str(r.get("Transaction Id") or "").strip()
            amount_str = format_payload_amount(r.get("Original Input Amt") or r.get("Amount") or r.get("Net Amount Payable(CR)"))
            utr_str = format_payload_utr(r.get("PARTNER_TXN_ID") or r.get("REF_TXN_NO_ORG") or r.get("UTR Num") or order_id)
            dt_str = format_payload_datetime(r.get("Date and Time") or r.get("Transaction Date") or r.get("TXN_DATE"))
            status_val = str(r.get("Status") or "SUCCESS").strip()

            txn_to = str(r.get("Transaction To") or "").strip()
            cand_tid = ""
            if txn_to:
                clean_to = txn_to.split("@")[0].strip()
                term_parts = clean_to.split("-")
                cand_tid = term_parts[-1].strip().upper()
            term_details = resolve_terminal_details(cand_tid, merchant_key=merchant_key, mappings=terminal_mappings) if cand_tid else None
            if term_details:
                term_id = term_details.get("mms_terminal_id") or term_details.get("terminal_id") or ""
                branch_name = term_details.get("branch_name") or ""
            else:
                resolved_tid = resolve_mms_terminal_id(cand_tid, terminal_mappings) if cand_tid else ""
                term_id = resolved_tid or cand_tid
            middle_number = term_id

        # 4. GENERIC / FALLBACK
        else:
            for k, v in r.items():
                k_low = str(k).lower().strip()
                v_str = str(v or "").strip()
                if not v_str:
                    continue
                if k_low in ("amount", "transaction amount", "amt", "net amount") and amount_str == "0.00":
                    amount_str = format_payload_amount(v)
                elif k_low in ("utr", "rrn", "bank reference no.", "bank ref", "reference id", "txn id", "rrn/utr") and not utr_str:
                    utr_str = format_payload_utr(v_str)
                elif k_low in ("order id", "order_id", "id", "transaction id") and not order_id:
                    order_id = v_str
                elif k_low in ("terminal id", "terminalid", "mms terminal id", "merchant mms terminal id", "tid") and not term_id:
                    term_id = v_str
                elif k_low in ("date", "transaction date", "transaction time", "date and time", "date & time", "transaction date & time") and not dt_str:
                    dt_str = format_payload_datetime(v_str)
                elif k_low in ("status", "transaction status") and not status_val:
                    status_val = v_str

            if term_id:
                term_details = resolve_terminal_details(term_id, merchant_key=merchant_key, mappings=terminal_mappings)
                if term_details:
                    branch_name = term_details.get("branch_name") or ""
                    term_id = term_details.get("mms_terminal_id") or term_id
            middle_number = term_id

        utr_str = format_payload_utr(utr_str)

        payload_dict = {
            "amount": amount_str,
            "terminalID": term_id,
            "utr": utr_str,
            "dateAndTime": dt_str
        }

        results.append({
            "index": idx,
            "gateway": gateway,
            "order_id": order_id,
            "middle_number": middle_number,
            "terminal_id": term_id,
            "branch_name": branch_name,
            "utr": utr_str,
            "amount": amount_str,
            "date_and_time": dt_str,
            "status": status_val or "SUCCESS",
            "payload": payload_dict,
            "payload_json": json.dumps(payload_dict, indent=2)
        })

    return results