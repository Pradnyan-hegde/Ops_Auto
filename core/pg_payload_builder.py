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

from .normalizer import clean_amount
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

    # 1. Cashfree unmatched
    unmatched_cf = getattr(engine, "unmatched_cf", [])
    if unmatched_cf:
        cf_prefix_map = build_cms_prefix_terminal_map(getattr(engine, "cms_report", None))

        for cf_r in unmatched_cf:
            order_id = str(cf_r.get("Order Id") or cf_r.get("Matched Transaction ID") or "").strip()
            amount_str = format_payload_amount(cf_r.get("Amount"))
            utr_str = str(cf_r.get("Bank Reference No.") or cf_r.get("UTR No.") or cf_r.get("Reference Id") or "").strip()
            dt_str = format_payload_datetime(cf_r.get("Transaction Time") or cf_r.get("Transaction Date"))

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
            utr_str = str(eb_r.get("UTR") or eb_r.get("UPI tid") or "").strip()
            dt_str = format_payload_datetime(eb_r.get("Transaction Date") or eb_r.get("Date"))
            term_id = str(eb_r.get("Virtual Account Label") or "").strip()

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
            utr_str = str(air_r.get("PARTNER_TXN_ID") or air_r.get("REF_TXN_NO_ORG") or order_id).strip()
            dt_str = format_payload_datetime(air_r.get("Date and Time") or air_r.get("Transaction Date"))

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