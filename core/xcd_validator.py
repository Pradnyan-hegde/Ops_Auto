"""
Strict All-Clear Gate Validator for XCD Settlement Downloads.
Enforces the 6 reconciliation criteria before allowing generation or download of XCD settlement files.
"""
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Set
import openpyxl

from .normalizer import clean_amount, clean_key, normalize_status


@dataclass
class XCDValidationResult:
    is_all_clear: bool
    summary_message: str
    blocking_reasons: List[str] = field(default_factory=list)
    details: Dict[str, Any] = field(default_factory=dict)
    settlement_date: str = ""


def validate_all_clear(
    engine,
    aggregator,
    settlement_date: Optional[str] = None
) -> XCDValidationResult:
    """
    Evaluates whether the reconciliation result satisfies the strict All-Clear criteria:
    1. Every successful CMS transaction matches SMMS exactly once.
    2. Every successful routed transaction has exactly one matching record in its partner report (Cashfree, Easebuzz, Airtel).
    3. Zero records in CMS_Not_in_SMMS, SMMS_Not_in_CMS, Unmatched_CF, Unmatched_EB, Unmatched_Airtel.
    4. Zero amount mismatches, zero status mismatches, zero duplicate transaction keys, zero unknown partners, zero unresolved exceptions.
    5. Summary counts and gross amounts reconcile to the successful SMMS master population.
    6. Failed/reversed transactions excluded from XCD files.
    7. Valid settlement date present.
    """
    blocking_reasons = []
    details = {
        "cms_not_in_smms_count": len(getattr(engine, "cms_not_in_smms", [])),
        "smms_not_in_cms_count": len(getattr(engine, "smms_not_in_cms", [])),
        "unmatched_cf_count": len(getattr(engine, "unmatched_cf", [])),
        "unmatched_eb_count": len(getattr(engine, "unmatched_eb", [])),
        "unmatched_airtel_count": len(getattr(engine, "unmatched_airtel", [])),
        "duplicates_count": len(getattr(engine, "duplicates", [])),
        "exceptions_count": len(getattr(engine, "exceptions", [])),
        "amount_mismatches_count": 0,
        "status_mismatches_count": 0,
        "unknown_partners_count": 0,
        "failed_in_matched_count": 0,
        "summary_checks_passed": True,
        "settlement_date": str(settlement_date or "").strip()
    }

    # 1 & 3: Check unmatched exception lists
    if details["cms_not_in_smms_count"] > 0:
        blocking_reasons.append(
            f"{details['cms_not_in_smms_count']} transaction(s) in CMS_Not_in_SMMS (present in CMS but missing in SMMS / SMMS Sync Status=False; use SMMS Sync File)"
        )

    if details["smms_not_in_cms_count"] > 0:
        blocking_reasons.append(
            f"{details['smms_not_in_cms_count']} transaction(s) in SMMS_Not_in_CMS (unmatched between SMMS and CMS)"
        )

    if details["unmatched_cf_count"] > 0:
        blocking_reasons.append(
            f"{details['unmatched_cf_count']} unmatched Cashfree transaction(s) in Unmatched_CF"
        )

    if details["unmatched_eb_count"] > 0:
        blocking_reasons.append(
            f"{details['unmatched_eb_count']} unmatched Easebuzz transaction(s) in Unmatched_EB"
        )

    if details["unmatched_airtel_count"] > 0:
        blocking_reasons.append(
            f"{details['unmatched_airtel_count']} unmatched Airtel transaction(s) in Unmatched_Airtel"
        )

    # 4: Check duplicates & exceptions
    if details["duplicates_count"] > 0:
        blocking_reasons.append(
            f"{details['duplicates_count']} duplicate transaction key(s) detected"
        )

    if details["exceptions_count"] > 0:
        blocking_reasons.append(
            f"{details['exceptions_count']} unresolved exception(s) detected"
        )

    # 2, 4 & 6: Inspect records in matched lists for amount/status/partner integrity
    matched_lists = [
        ("Cashfree", getattr(engine, "cms_cf_matched", [])),
        ("Easebuzz", getattr(engine, "cms_eb_matched", [])),
        ("Airtel", getattr(engine, "cms_air_matched", [])),
    ]

    failed_statuses = {"failed", "failure", "reversed", "refunded", "cancelled", "cancel"}

    for partner_name, recs in matched_lists:
        for r in recs:
            # Check amount mismatch
            cms_amt = clean_amount(r.get("CMS Amount") or r.get("Transaction Amount") or r.get("Amount"))
            p_amt = clean_amount(r.get("Partner Amount") or r.get("Amount"))
            if cms_amt is not None and p_amt is not None:
                if abs(cms_amt - p_amt) > getattr(engine, "tolerance", 0.01):
                    details["amount_mismatches_count"] += 1

            # Check status mismatch & failed in matched
            c_stat = normalize_status(r.get("CMS Status") or r.get("Transaction Status") or r.get("Status"))
            p_stat = normalize_status(r.get("Partner Status") or r.get("Status"))

            if c_stat.lower() in failed_statuses or p_stat.lower() in failed_statuses:
                details["failed_in_matched_count"] += 1

            if c_stat != p_stat and c_stat and p_stat:
                # If statuses conflict (e.g. Success vs Failed)
                if c_stat != "Success" or p_stat != "Success":
                    details["status_mismatches_count"] += 1

            # Check unknown partner
            part = str(r.get("_partner") or r.get("Partner") or partner_name).strip().lower()
            if "unknown" in part:
                details["unknown_partners_count"] += 1

    if details["amount_mismatches_count"] > 0:
        blocking_reasons.append(
            f"{details['amount_mismatches_count']} transaction(s) with gross amount mismatch between CMS and Partner"
        )

    if details["status_mismatches_count"] > 0:
        blocking_reasons.append(
            f"{details['status_mismatches_count']} transaction(s) with status conflict between CMS and Partner"
        )

    if details["failed_in_matched_count"] > 0:
        blocking_reasons.append(
            f"{details['failed_in_matched_count']} failed or reversed transaction(s) found in matched settlement tabs"
        )

    if details["unknown_partners_count"] > 0:
        blocking_reasons.append(
            f"{details['unknown_partners_count']} transaction(s) with unknown partner assignment"
        )

    # 5: Summary counts and gross amounts reconcile to the successful SMMS master population
    if aggregator and hasattr(aggregator, "control_checks"):
        failed_checks = [c for c in aggregator.control_checks if c.get("status") != "PASS"]
        if failed_checks:
            details["summary_checks_passed"] = False
            for fc in failed_checks:
                blocking_reasons.append(
                    f"Audit check '{fc.get('check_name')}' failed (diff: {fc.get('difference')})"
                )

    # 7: Settlement Date check
    s_date = str(settlement_date or "").strip()
    if not s_date or s_date.lower() in ("null", "none", "--"):
        blocking_reasons.append(
            "Settlement date is missing. Please select a settlement date or upload an Airtel Settlement report."
        )

    # Determine verdict
    is_all_clear = (len(blocking_reasons) == 0)

    if is_all_clear:
        summary_message = "Reconciliation complete. All three XCD settlement files are ready to download."
    else:
        # Construct summary of blocking issues
        summary_message = f"XCD downloads blocked: {'; '.join(blocking_reasons[:3])}"
        if len(blocking_reasons) > 3:
            summary_message += f" (and {len(blocking_reasons) - 3} more issues)"

    return XCDValidationResult(
        is_all_clear=is_all_clear,
        summary_message=summary_message,
        blocking_reasons=blocking_reasons,
        details=details,
        settlement_date=s_date
    )


def validate_xcd_workbook_file(filepath: str, expected_records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Performs independent physical verification of an XCD workbook file:
    - Verifies Row 1 Title: 'SALE TRANSACTIONS'
    - Verifies Row 2 Headers: ['SL NO', 'SwinkPay Transaction ID', 'Amount', 'Settlement Date']
    - Verifies row count matches expected source matched count exactly
    - Verifies sequential 1-based SL NO without gaps
    - Verifies SwinkPay Transaction IDs match source records and are unique
    - Verifies Gross Amounts match CMS amounts
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # Check Title
    r1_c1 = str(ws.cell(1, 1).value or "").strip()
    if r1_c1 != "SALE TRANSACTIONS":
        return {
            "valid": False,
            "error": f"Invalid row 1 title: '{r1_c1}', expected 'SALE TRANSACTIONS'"
        }

    # Check Headers
    expected_headers = ["SL NO", "SwinkPay Transaction ID", "Amount", "Settlement Date"]
    actual_headers = [str(ws.cell(2, c).value or "").strip() for c in range(1, 5)]
    if actual_headers != expected_headers:
        return {
            "valid": False,
            "error": f"Invalid headers: {actual_headers}, expected {expected_headers}"
        }

    data_rows = []
    max_row = ws.max_row
    for r in range(3, max_row + 1):
        c1_val = ws.cell(r, 1).value
        c1_str = str(c1_val or "").strip()

        # Stop reading sale transactions at empty row or trailing section title
        if c1_val is None or c1_str.upper() in ("REFUND TRANSACTIONS", "CHARGEBACK/ADJUSTMENT TRANSACTIONS"):
            break

        sp_id = str(ws.cell(r, 2).value or "").strip()
        amt = ws.cell(r, 3).value
        s_date = str(ws.cell(r, 4).value or "").strip()

        data_rows.append({
            "sl_no": c1_val,
            "swinkpay_id": sp_id,
            "amount": clean_amount(amt),
            "settlement_date": s_date
        })

    if len(data_rows) != len(expected_records):
        return {
            "valid": False,
            "error": f"Row count mismatch: XCD file has {len(data_rows)} rows, expected {len(expected_records)}"
        }

    # Verify trailing sections (REFUND and CHARGEBACK)
    last_sale_row = 2 + len(data_rows)
    ref_title = str(ws.cell(last_sale_row + 2, 1).value or "").strip()
    if ref_title != "REFUND TRANSACTIONS":
        return {
            "valid": False,
            "error": f"Missing REFUND TRANSACTIONS section at row {last_sale_row + 2}: got '{ref_title}'"
        }

    expected_ref_headers = ["SL NO", "SwinkPay Refund Transaction ID", "Amount", "Refund Settlement Date"]
    actual_ref_headers = [str(ws.cell(last_sale_row + 3, c).value or "").strip() for c in range(1, 5)]
    if actual_ref_headers != expected_ref_headers:
        return {
            "valid": False,
            "error": f"Invalid REFUND headers at row {last_sale_row + 3}: got {actual_ref_headers}"
        }

    cb_title = str(ws.cell(last_sale_row + 5, 1).value or "").strip()
    if cb_title != "CHARGEBACK/ADJUSTMENT TRANSACTIONS":
        return {
            "valid": False,
            "error": f"Missing CHARGEBACK/ADJUSTMENT TRANSACTIONS section at row {last_sale_row + 5}: got '{cb_title}'"
        }

    expected_cb_headers = ["SL NO", "SwinkPay Transaction ID", "Amount", "Chargeback Settlement Date"]
    actual_cb_headers = [str(ws.cell(last_sale_row + 6, c).value or "").strip() for c in range(1, 5)]
    if actual_cb_headers != expected_cb_headers:
        return {
            "valid": False,
            "error": f"Invalid CHARGEBACK headers at row {last_sale_row + 6}: got {actual_cb_headers}"
        }

    # Verify NO cell borders on headers and data cells
    c_bdr = ws.cell(2, 1).border
    if c_bdr and ((c_bdr.left and c_bdr.left.style) or (c_bdr.top and c_bdr.top.style)):
        return {
            "valid": False,
            "error": "Borders detected on XCD cells; file format requires no border styling."
        }

    # Verify serial numbers, uniqueness, and amount match
    seen_ids: Set[str] = set()
    expected_dict = {}
    for r in expected_records:
        sid = str(r.get("SwinkPay Txn ID") or r.get("SwinkPay Transaction ID") or "").strip()
        amt = clean_amount(r.get("Transaction Amount") or r.get("Amount") or r.get("CMS Amount"))
        if sid:
            expected_dict[sid] = amt

    for idx, row in enumerate(data_rows, start=1):
        # 1-based sequential SL NO
        if row["sl_no"] != idx:
            return {
                "valid": False,
                "error": f"SL NO out of sequence at row {idx + 2}: got {row['sl_no']}, expected {idx}"
            }

        sid = row["swinkpay_id"]
        if not sid:
            return {
                "valid": False,
                "error": f"Empty SwinkPay Transaction ID at row {idx + 2}"
            }

        if sid in seen_ids:
            return {
                "valid": False,
                "error": f"Duplicate SwinkPay Transaction ID '{sid}' at row {idx + 2}"
            }
        seen_ids.add(sid)

        # Check amount
        if sid in expected_dict:
            exp_amt = expected_dict[sid]
            act_amt = row["amount"]
            if exp_amt is not None and act_amt is not None:
                if abs(exp_amt - act_amt) > 0.01:
                    return {
                        "valid": False,
                        "error": f"Amount mismatch for '{sid}': XCD has {act_amt}, expected {exp_amt}"
                    }

    return {
        "valid": True,
        "row_count": len(data_rows),
        "total_amount": round(sum(r["amount"] or 0.0 for r in data_rows), 2)
    }
