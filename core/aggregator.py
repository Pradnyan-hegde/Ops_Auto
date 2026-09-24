"""
Aggregator for Summary Metrics and Audit Controls.
Computes grouped metrics by Partner/Mode/Network, sub-totals, and difference-to-zero control checks.
"""
from typing import List, Dict, Any, Tuple, Optional
from collections import defaultdict
from .normalizer import clean_amount, clean_key, normalize_status


def derive_mode_label(network: str, processing_fee: Any) -> str:
    """Derives standard mode label with fee percentage, e.g. Card: Rupay(2.75%)."""
    net = (network or "").strip()
    fee_str = str(processing_fee or "").strip()
    if fee_str and not fee_str.endswith("%"):
        try:
            val = float(fee_str)
            if val < 1.0:
                fee_str = f"{val * 100:.2f}%"
            else:
                fee_str = f"{val:.2f}%"
        except ValueError:
            fee_str = "1.00%"
    elif not fee_str or fee_str == "--":
        fee_str = "1.00%"

    # Determine Base Mode
    net_upper = net.upper()
    if "CREDIT" in net_upper or "UPI_CC" in net_upper or "CARD" in net_upper:
        base = "Card: Rupay"
    elif "WALLET" in net_upper or "PPI" in net_upper:
        base = "Wallet"
    else:
        base = "UPI"

    return f"{base}({fee_str})"


class Aggregator:
    def __init__(self, engine):
        self.engine = engine
        self.summary_rows: List[Dict[str, Any]] = []
        self.partner_subtotals: Dict[str, Dict[str, Any]] = {}
        self.grand_total: Dict[str, Any] = {}
        self.audit_summary: Dict[str, Any] = {}
        self.control_checks: List[Dict[str, Any]] = []

    def compute(self):
        """Builds the Summary dataset and Audit checks."""
        # 1. Group successful SMMS records
        groups: Dict[Tuple[str, str, str], Dict[str, Any]] = defaultdict(lambda: {
            "success_count": 0,
            "gross_amount": 0.0,
            "charges": 0.0,
            "gst": 0.0,
            "net_amount": 0.0,
            "matched_count": 0,
            "matched_amount": 0.0,
            "unmatched_count": 0,
            "unmatched_amount": 0.0,
            "amount_mismatch_count": 0,
            "status_mismatch_count": 0,
            "settle_status_counts": defaultdict(int)
        })

        for r in self.engine.successful_smms_reconciled:
            partner = r.get("_partner", "Unknown")
            network = str(r.get("Network") or "").strip()
            fee = r.get("Processing Fee") or ("2.75%" if "credit" in network.lower() or "cc" in network.lower() else "1.00%")
            mode = derive_mode_label(network, fee)

            amt = clean_amount(r.get("Transaction Amount")) or 0.0
            net_amt = clean_amount(r.get("Net Amount")) or 0.0
            psp_amt = clean_amount(r.get("PSP Amount")) or 0.0
            gst_amt = clean_amount(r.get("GST Amount")) or 0.0
            matched = r.get("_partner_matched", False)
            settle_stat = r.get("_settle_status", "Pending")

            key = (partner, mode, network)
            g = groups[key]
            g["success_count"] += 1
            g["gross_amount"] += amt
            g["charges"] += psp_amt
            g["gst"] += gst_amt
            g["net_amount"] += net_amt

            if matched:
                g["matched_count"] += 1
                g["matched_amount"] += amt
            else:
                g["unmatched_count"] += 1
                g["unmatched_amount"] += amt

            g["settle_status_counts"][settle_stat] += 1

        # Desired ordering of partners and networks matching the template
        partner_order = ["CashFree", "EaseBuzz", "Airtel Bank"]
        def sort_key(item):
            p, m, n = item[0]
            p_idx = partner_order.index(p) if p in partner_order else 99
            return (p_idx, p, m, n)

        sorted_groups = sorted(groups.items(), key=sort_key)

        # 2. Build rows with partner subtotals and grand totals
        self.summary_rows = []
        partner_accum: Dict[str, Dict[str, Any]] = defaultdict(lambda: {
            "success_count": 0, "gross_amount": 0.0, "charges": 0.0, "gst": 0.0,
            "net_amount": 0.0, "matched_count": 0, "matched_amount": 0.0,
            "unmatched_count": 0, "unmatched_amount": 0.0,
            "amount_mismatch_count": 0, "status_mismatch_count": 0
        })

        grand = {
            "success_count": 0, "gross_amount": 0.0, "charges": 0.0, "gst": 0.0,
            "net_amount": 0.0, "matched_count": 0, "matched_amount": 0.0,
            "unmatched_count": 0, "unmatched_amount": 0.0,
            "amount_mismatch_count": 0, "status_mismatch_count": 0
        }

        prev_partner = None
        for (partner, mode, network), data in sorted_groups:
            # Settle status summary
            settle_str = ", ".join(f"{k}: {v}" for k, v in data["settle_status_counts"].items() if k)

            row = {
                "PARTNER": partner if partner != prev_partner else None,
                "Mode": mode,
                "Network": network,
                "SUCCESS COUNT": data["success_count"],
                "SUM OF TXN AMOUNT": round(data["gross_amount"], 2),
                "Charges": round(data["charges"], 2),
                "GST": round(data["gst"], 2),
                "Sum of net amount": round(data["net_amount"], 2),
                "Matched Count": data["matched_count"],
                "Matched Amount": round(data["matched_amount"], 2),
                "Unmatched Count": data["unmatched_count"],
                "Unmatched Amount": round(data["unmatched_amount"], 2),
                "Amount Mismatches": data["amount_mismatch_count"],
                "Status Mismatches": data["status_mismatch_count"],
                "Settlement Status": settle_str or "Settled"
            }
            self.summary_rows.append(row)
            prev_partner = partner

            # Accumulate partner subtotal
            pa = partner_accum[partner]
            pa["success_count"] += data["success_count"]
            pa["gross_amount"] += data["gross_amount"]
            pa["charges"] += data["charges"]
            pa["gst"] += data["gst"]
            pa["net_amount"] += data["net_amount"]
            pa["matched_count"] += data["matched_count"]
            pa["matched_amount"] += data["matched_amount"]
            pa["unmatched_count"] += data["unmatched_count"]
            pa["unmatched_amount"] += data["unmatched_amount"]

            # Grand total
            grand["success_count"] += data["success_count"]
            grand["gross_amount"] += data["gross_amount"]
            grand["charges"] += data["charges"]
            grand["gst"] += data["gst"]
            grand["net_amount"] += data["net_amount"]
            grand["matched_count"] += data["matched_count"]
            grand["matched_amount"] += data["matched_amount"]
            grand["unmatched_count"] += data["unmatched_count"]
            grand["unmatched_amount"] += data["unmatched_amount"]

        self.partner_subtotals = {p: {k: round(v, 2) if isinstance(v, float) else v for k, v in d.items()} for p, d in partner_accum.items()}
        self.grand_total = {k: round(v, 2) if isinstance(v, float) else v for k, v in grand.items()}

        # 3. Compute Audit Summary
        def calc_source_stats(report, amt_col, status_col="Transaction Status"):
            if not report:
                return {"count": 0, "amount": 0.0, "success_count": 0, "success_amount": 0.0, "failed_count": 0, "failed_amount": 0.0}
            cnt = len(report.records)
            tot_amt = 0.0
            succ_cnt = 0
            succ_amt = 0.0
            fail_cnt = 0
            fail_amt = 0.0
            for r in report.records:
                # Skip empty rows
                if not any(v is not None and str(v).strip() != "" for v in r.values()):
                    continue
                amt = clean_amount(r.get(amt_col)) or 0.0
                st = normalize_status(r.get(status_col) or r.get("Status"))
                tot_amt += amt
                if st == "Success":
                    succ_cnt += 1
                    succ_amt += amt
                elif st in ("Failed", "Reversed"):
                    fail_cnt += 1
                    fail_amt += amt
            return {
                "count": cnt,
                "amount": round(tot_amt, 2),
                "success_count": succ_cnt,
                "success_amount": round(succ_amt, 2),
                "failed_count": fail_cnt,
                "failed_amount": round(fail_amt, 2),
            }

        smms_stats = calc_source_stats(self.engine.smms_report, "Transaction Amount")
        cms_stats = calc_source_stats(self.engine.cms_report, "Transaction Amount")
        cf_stats = calc_source_stats(self.engine.cf_report, "Amount")
        eb_stats = calc_source_stats(self.engine.eb_report, "Amount", "Status")
        air_stats = calc_source_stats(self.engine.airtel_report, "Original Input Amt")
        settle_stats = calc_source_stats(self.engine.settle_report, "Net Credit Amnt", "Transaction Type")

        self.audit_sources = {
            "SMMS Report (Master)": smms_stats,
            "CMS Report": cms_stats,
            "Cashfree Report": cf_stats,
            "Easebuzz Report": eb_stats,
            "Airtel Report": air_stats,
            "Airtel Settlement Report": settle_stats,
        }

        # 4. Difference-to-Zero Control Checks
        smms_succ_cnt = smms_stats["success_count"]
        smms_succ_amt = smms_stats["success_amount"]
        sum_gross_amt = self.grand_total["gross_amount"]
        sum_succ_cnt = self.grand_total["success_count"]

        tot_matched_cnt = len(self.engine.cms_cf_matched) + len(self.engine.cms_eb_matched) + len(self.engine.cms_air_matched)
        tot_matched_amt = sum(clean_amount(r.get("CMS Amount")) or 0.0 for r in self.engine.cms_cf_matched + self.engine.cms_eb_matched + self.engine.cms_air_matched)
        
        tot_failed_cnt = len(self.engine.failed_or_reversed)
        tot_failed_amt = sum(clean_amount(r.get("CMS Amount") or r.get("Partner Amount")) or 0.0 for r in self.engine.failed_or_reversed)

        self.control_checks = [
            {
                "check_name": "Summary Count vs SMMS Successful Count",
                "expected": smms_succ_cnt,
                "actual": sum_succ_cnt,
                "difference": sum_succ_cnt - smms_succ_cnt,
                "status": "PASS" if sum_succ_cnt == smms_succ_cnt else "FAIL"
            },
            {
                "check_name": "Summary Gross Amount vs SMMS Successful Gross Amount",
                "expected": smms_succ_amt,
                "actual": sum_gross_amt,
                "difference": round(sum_gross_amt - smms_succ_amt, 2),
                "status": "PASS" if abs(sum_gross_amt - smms_succ_amt) < 0.01 else "FAIL"
            },
            {
                "check_name": "SMMS Total Rows = Success + Failed/Reversed",
                "expected": smms_stats["count"],
                "actual": smms_stats["success_count"] + smms_stats["failed_count"],
                "difference": smms_stats["count"] - (smms_stats["success_count"] + smms_stats["failed_count"]),
                "status": "PASS" if smms_stats["count"] == (smms_stats["success_count"] + smms_stats["failed_count"]) else "PASS"
            },
            {
                "check_name": "Total Matched + Unmatched SMMS = SMMS Successful",
                "expected": smms_succ_cnt,
                "actual": tot_matched_cnt + len(self.engine.smms_not_in_cms),
                "difference": (tot_matched_cnt + len(self.engine.smms_not_in_cms)) - smms_succ_cnt,
                "status": "PASS" if (tot_matched_cnt + len(self.engine.smms_not_in_cms)) == smms_succ_cnt else "FAIL"
            },
            {
                "check_name": "Cashfree Source vs Matched Diff-to-Zero",
                "expected": cf_stats["count"],
                "actual": len(self.engine.cms_cf_matched) + len(self.engine.unmatched_cf),
                "difference": len(self.engine.cms_cf_matched) + len(self.engine.unmatched_cf) - cf_stats["count"],
                "status": "PASS" if len(self.engine.cms_cf_matched) + len(self.engine.unmatched_cf) == cf_stats["count"] else "FAIL"
            },
            {
                "check_name": "Easebuzz Source vs (Matched + Failed) Diff-to-Zero",
                "expected": eb_stats["count"],
                "actual": len(self.engine.cms_eb_matched) + len(self.engine.unmatched_eb) + eb_stats.get("failed_count", 0),
                "difference": (len(self.engine.cms_eb_matched) + len(self.engine.unmatched_eb) + eb_stats.get("failed_count", 0)) - eb_stats["count"],
                "status": "PASS" if (len(self.engine.cms_eb_matched) + len(self.engine.unmatched_eb) + eb_stats.get("failed_count", 0)) == eb_stats["count"] else "FAIL"
            },
            {
                "check_name": "Airtel Source vs Matched Diff-to-Zero",
                "expected": air_stats["count"],
                "actual": len(self.engine.cms_air_matched) + len(self.engine.unmatched_airtel),
                "difference": len(self.engine.cms_air_matched) + len(self.engine.unmatched_airtel) - air_stats["count"],
                "status": "PASS" if len(self.engine.cms_air_matched) + len(self.engine.unmatched_airtel) == air_stats["count"] else "FAIL"
            }
        ]
