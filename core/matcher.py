"""
Core Reconciliation Engine.
Performs multi-way matching across CMS, SMMS, Cashfree, Easebuzz, Airtel, and Airtel Settlement.
"""
from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional, Tuple, Set
from collections import defaultdict

from .detector import ReportType
from .reader import RawReport
from .normalizer import clean_key, clean_amount, is_amount_equal, normalize_status


@dataclass
class ReconRecord:
    source_system: str
    match_key: str
    matched_txn_id: str
    cms_amount: Optional[float]
    partner_amount: Optional[float]
    amount_difference: Optional[float]
    cms_status: str
    partner_status: str
    recon_status: str
    exception_reason: str
    settlement_utr: Optional[str] = None
    settlement_date: Optional[str] = None
    settlement_amount: Optional[float] = None
    settlement_status: Optional[str] = None
    raw_record: Dict[str, Any] = field(default_factory=dict)
    partner_name: str = ""
    mode: str = ""
    network: str = ""


def _index_record(index: Dict[str, Any], key: str, record: Any):
    if not key:
        return
    index[key] = record
    if key.isdigit():
        index[key.lstrip("0")] = record
        index[key.zfill(12)] = record


def _find_record(index: Dict[str, Any], key: str) -> Optional[Any]:
    if not key:
        return None
    if key in index:
        return index[key]
    if key.isdigit():
        k_strip = key.lstrip("0")
        if k_strip in index:
            return index[k_strip]
        k_zfill = key.zfill(12)
        if k_zfill in index:
            return index[k_zfill]
    return None


def _record_matched_key(matched_set: Set[str], key: str):
    if not key:
        return
    matched_set.add(key)
    if key.isdigit():
        matched_set.add(key.lstrip("0"))
        matched_set.add(key.zfill(12))


def _is_key_matched(matched_set: Set[str], key: str) -> bool:
    if not key:
        return False
    if key in matched_set:
        return True
    if key.isdigit():
        if key.lstrip("0") in matched_set or key.zfill(12) in matched_set:
            return True
    return False


def _is_adjustment_record(row: Dict[str, Any]) -> Tuple[bool, str]:
    """Detects if a transaction is an adjustment, refund, chargeback, or dispute (Status 3, 4, 5, 6)."""
    st_raw = str(row.get("Transaction Status") or "").strip()
    mode_raw = str(row.get("Mode") or "").strip().lower()
    net_raw = str(row.get("Network") or "").strip().lower()
    desc_raw = str(row.get("Transaction Description") or "").strip().lower()

    if st_raw in ("3", "3.0") or "refund" in mode_raw or "refund" in desc_raw:
        return True, "Refund"
    elif st_raw in ("4", "4.0") or "chargeback" in mode_raw or "chargeback" in desc_raw:
        return True, "Chargeback"
    elif st_raw in ("5", "5.0") or "dispute" in mode_raw or "dispute" in desc_raw:
        return True, "Dispute"
    elif st_raw in ("6", "6.0") or "adjustment" in mode_raw or "adjustment" in net_raw or "adjustment" in desc_raw:
        return True, "Adjustment"

    return False, ""


class ReconciliationEngine:
    def __init__(self, tolerance: float = 0.01):
        self.tolerance = tolerance

        # Raw reports
        self.cms_report: Optional[RawReport] = None
        self.smms_report: Optional[RawReport] = None
        self.cf_report: Optional[RawReport] = None
        self.eb_report: Optional[RawReport] = None
        self.airtel_report: Optional[RawReport] = None
        self.settle_report: Optional[RawReport] = None

        # Result sets
        self.cms_smms_matched: List[Dict[str, Any]] = []
        self.cms_cf_matched: List[Dict[str, Any]] = []
        self.cms_eb_matched: List[Dict[str, Any]] = []
        self.cms_air_matched: List[Dict[str, Any]] = []

        self.cms_not_in_smms: List[Dict[str, Any]] = []
        self.smms_not_in_cms: List[Dict[str, Any]] = []

        self.unmatched_cf: List[Dict[str, Any]] = []
        self.unmatched_eb: List[Dict[str, Any]] = []
        self.unmatched_airtel: List[Dict[str, Any]] = []

        self.failed_or_reversed: List[Dict[str, Any]] = []
        self.adjustments: List[Dict[str, Any]] = []
        self.duplicates: List[Dict[str, Any]] = []
        self.exceptions: List[Dict[str, Any]] = []
        self.network_changes: List[Dict[str, str]] = []

        # Master successful reconciled SMMS records for Summary
        self.successful_smms_reconciled: List[Dict[str, Any]] = []

    def set_reports(
        self,
        cms: RawReport,
        smms: RawReport,
        cf: Optional[RawReport] = None,
        eb: Optional[RawReport] = None,
        airtel: Optional[RawReport] = None,
        settle: Optional[RawReport] = None,
    ):
        self.cms_report = cms
        self.smms_report = smms
        self.cf_report = cf
        self.eb_report = eb
        self.airtel_report = airtel
        self.settle_report = settle

    def _classify_partner(self, smms_row: Dict[str, Any], cms_row: Optional[Dict[str, Any]] = None) -> str:
        """Determines partner: CashFree, EaseBuzz, or Airtel Bank."""
        pg = str(smms_row.get("PG/Bank") or "").strip().lower()
        if "cashfree" in pg or "cf" in pg:
            return "CashFree"
        elif "easebuzz" in pg or "eb" in pg:
            return "EaseBuzz"
        elif "airtel" in pg or "air" in pg:
            return "Airtel Bank"

        # Check CMS Payment Gateway if available
        if cms_row:
            cpg = str(cms_row.get("Payment Gateway") or "").strip().lower()
            if "cashfree" in cpg:
                return "CashFree"
            elif "easebuzz" in cpg:
                return "EaseBuzz"
            elif "airtel" in cpg:
                return "Airtel Bank"

        # Check network keywords
        netw = str(smms_row.get("Network") or "").strip().upper()
        if netw in ("CREDIT", "CURRENT", "NRO", "SAVINGS", "PPIWALLET"):
            return "Airtel Bank"
        elif netw in ("UPI_CC", "UPI"):
            return "EaseBuzz"
        elif "offline_static" in netw.lower():
            return "CashFree"

        return "Unknown"

    def _check_duplicates(self, report: RawReport, key_col: str, file_label: str):
        """Identifies duplicate keys in a report and adds to duplicates list."""
        key_counts = defaultdict(list)
        for idx, row in enumerate(report.records):
            k = clean_key(row.get(key_col))
            if k:
                key_counts[k].append((idx, row))

        for k, occurrences in key_counts.items():
            if len(occurrences) > 1:
                for idx, row in occurrences:
                    dup_entry = dict(row)
                    dup_entry.update({
                        "Source System": file_label,
                        "Match Key": k,
                        "Matched Transaction ID": row.get("SwinkPay Txn ID") or row.get("Transaction Id") or row.get("Order Id") or "",
                        "CMS Amount": clean_amount(row.get("Transaction Amount") or row.get("Amount") or row.get("Original Input Amt")),
                        "Partner Amount": None,
                        "Amount Difference": None,
                        "CMS Status": normalize_status(row.get("Transaction Status") or row.get("Status")),
                        "Partner Status": "",
                        "Reconciliation Status": "Duplicate Key",
                        "Exception Reason": f"Duplicate key '{k}' found {len(occurrences)} times in {file_label}",
                        "Settlement Details": ""
                    })
                    self.duplicates.append(dup_entry)

    def run(self):
        """Executes full multi-way reconciliation workflow."""
        if not self.cms_report or not self.smms_report:
            raise ValueError("CMS and SMMS reports are mandatory for reconciliation.")

        # 1. Duplicate checks across source reports
        self._check_duplicates(self.cms_report, "RRN/UTR", "CMS")
        self._check_duplicates(self.smms_report, "RRN/UTR", "SMMS")
        if self.cf_report:
            self._check_duplicates(self.cf_report, "Bank Reference No.", "Cashfree")
        if self.eb_report:
            self._check_duplicates(self.eb_report, "UTR", "Easebuzz")
        if self.airtel_report:
            self._check_duplicates(self.airtel_report, "PARTNER_TXN_ID", "Airtel")

        # 2. Build index of CMS records: by RRN/UTR and SwinkPay Txn ID
        cms_by_rrn: Dict[str, Dict[str, Any]] = {}
        cms_by_sp_id: Dict[str, Dict[str, Any]] = {}
        for r in self.cms_report.records:
            rrn = clean_key(r.get("RRN/UTR"))
            sp_id = clean_key(r.get("SwinkPay Txn ID"))
            if rrn:
                _index_record(cms_by_rrn, rrn, r)
            if sp_id:
                _index_record(cms_by_sp_id, sp_id, r)

        # Build index of Airtel Settlement records if available
        settle_by_ref: Dict[str, Dict[str, Any]] = {}
        if self.settle_report:
            for r in self.settle_report.records:
                pref = clean_key(r.get("Partner Txn Ref No"))
                bref = clean_key(r.get("Bank Txn Ref No"))
                if pref:
                    _index_record(settle_by_ref, pref, r)
                if bref:
                    _index_record(settle_by_ref, bref, r)

        # Index partner records
        cf_by_ref: Dict[str, Dict[str, Any]] = {}
        if self.cf_report:
            for r in self.cf_report.records:
                bref = clean_key(r.get("Bank Reference No."))
                if bref:
                    _index_record(cf_by_ref, bref, r)

        eb_by_utr: Dict[str, Dict[str, Any]] = {}
        if self.eb_report:
            for r in self.eb_report.records:
                utr = clean_key(r.get("UTR"))
                if utr:
                    _index_record(eb_by_utr, utr, r)

        air_by_pid: Dict[str, Dict[str, Any]] = {}
        air_by_tid: Dict[str, Dict[str, Any]] = {}
        if self.airtel_report:
            for r in self.airtel_report.records:
                pid = clean_key(r.get("PARTNER_TXN_ID"))
                tid = clean_key(r.get("Transaction Id"))
                if pid:
                    _index_record(air_by_pid, pid, r)
                if tid:
                    _index_record(air_by_tid, tid, r)

        # Track which CMS, CF, EB, Airtel records and failures were matched
        matched_cms_keys: Set[str] = set()
        matched_cf_keys: Set[str] = set()
        matched_eb_keys: Set[str] = set()
        matched_air_keys: Set[str] = set()
        handled_failed_keys: Set[str] = set()

        # 3. Master Reconciliation: Iterate through SMMS records
        for smms_r in self.smms_report.records:
            smms_rrn = clean_key(smms_r.get("RRN/UTR"))
            smms_sp_id = clean_key(smms_r.get("SwinkPay Txn ID"))
            smms_amt = clean_amount(smms_r.get("Transaction Amount"))
            smms_status = normalize_status(smms_r.get("Transaction Status"))

            # Skip empty blank rows
            if not smms_rrn and not smms_sp_id and smms_amt is None and not smms_status:
                continue

            partner = self._classify_partner(smms_r)

            # Match CMS: Try RRN first, then SwinkPay Txn ID
            cms_r = _find_record(cms_by_rrn, smms_rrn)
            match_key = smms_rrn
            if not cms_r and smms_sp_id:
                cms_r = _find_record(cms_by_sp_id, smms_sp_id)
                match_key = smms_sp_id

            if not cms_r:
                # Missing in CMS
                entry = dict(smms_r)
                entry.update({
                    "Source System": "SMMS",
                    "Match Key": match_key,
                    "Matched Transaction ID": smms_sp_id,
                    "CMS Amount": None,
                    "Partner Amount": smms_amt,
                    "Amount Difference": None,
                    "CMS Status": "",
                    "Partner Status": smms_status,
                    "Reconciliation Status": "Missing in CMS",
                    "Exception Reason": f"Transaction exists in SMMS but missing in CMS report (Key: {match_key})",
                    "Settlement UTR": smms_r.get("Settlement UTR"),
                    "Settlement Date": smms_r.get("Settlement Date & Time"),
                    "Settlement Amount": clean_amount(smms_r.get("Net Amount")),
                    "Settlement Status": smms_r.get("Settlement Status")
                })
                self.smms_not_in_cms.append(entry)
                self.exceptions.append(entry)
                continue

            # Record CMS as matched to SMMS
            if smms_rrn:
                _record_matched_key(matched_cms_keys, smms_rrn)
            if smms_sp_id:
                _record_matched_key(matched_cms_keys, smms_sp_id)

            cms_amt = clean_amount(cms_r.get("Transaction Amount"))
            cms_status = normalize_status(cms_r.get("Transaction Status"))

            # Check if Failed, Reversed, or other non-success status (3, 4, 5, 6)
            if cms_status != "Success" or smms_status != "Success":
                is_adj, adj_cat = _is_adjustment_record(cms_r)
                if not is_adj:
                    is_adj, adj_cat = _is_adjustment_record(smms_r)

                non_success_st = adj_cat if is_adj else (cms_status if cms_status != "Success" else smms_status)
                entry = dict(cms_r)
                entry.update({
                    "Source System": "CMS/SMMS",
                    "Match Key": match_key,
                    "Matched Transaction ID": smms_sp_id,
                    "CMS Amount": cms_amt,
                    "Partner Amount": smms_amt,
                    "Amount Difference": round(cms_amt - smms_amt, 2) if cms_amt is not None and smms_amt is not None else None,
                    "CMS Status": cms_status,
                    "Partner Status": smms_status,
                    "Reconciliation Status": non_success_st,
                    "Adjustment Category": adj_cat if is_adj else "",
                    "Exception Reason": (f"{adj_cat}: {cms_r.get('Transaction Description') or 'Adjustment/Deduction'}" if is_adj else f"Transaction not successful in source (CMS: {cms_status}, SMMS: {smms_status})"),
                    "Settlement UTR": smms_r.get("Settlement UTR"),
                    "Settlement Date": smms_r.get("Settlement Date & Time"),
                    "Settlement Amount": clean_amount(smms_r.get("Net Amount")),
                    "Settlement Status": smms_r.get("Settlement Status")
                })
                if match_key:
                    _record_matched_key(handled_failed_keys, match_key)
                if smms_rrn:
                    _record_matched_key(handled_failed_keys, smms_rrn)
                if smms_sp_id:
                    _record_matched_key(handled_failed_keys, smms_sp_id)
                if is_adj:
                    self.adjustments.append(entry)
                self.failed_or_reversed.append(entry)
                continue

            # Amount Mismatch between CMS and SMMS
            if not is_amount_equal(cms_amt, smms_amt, self.tolerance):
                diff = round((cms_amt or 0.0) - (smms_amt or 0.0), 2)
                entry = dict(cms_r)
                entry.update({
                    "Source System": "CMS/SMMS",
                    "Match Key": match_key,
                    "Matched Transaction ID": smms_sp_id,
                    "CMS Amount": cms_amt,
                    "Partner Amount": smms_amt,
                    "Amount Difference": diff,
                    "CMS Status": cms_status,
                    "Partner Status": smms_status,
                    "Reconciliation Status": "Amount Mismatch",
                    "Exception Reason": f"Amount mismatch between CMS ({cms_amt}) and SMMS ({smms_amt}), diff={diff}",
                    "Settlement UTR": smms_r.get("Settlement UTR"),
                    "Settlement Date": smms_r.get("Settlement Date & Time"),
                    "Settlement Amount": clean_amount(smms_r.get("Net Amount")),
                    "Settlement Status": smms_r.get("Settlement Status")
                })
                self.exceptions.append(entry)
                continue

            # Status Mismatch between CMS and SMMS
            if cms_status != smms_status:
                entry = dict(cms_r)
                entry.update({
                    "Source System": "CMS/SMMS",
                    "Match Key": match_key,
                    "Matched Transaction ID": smms_sp_id,
                    "CMS Amount": cms_amt,
                    "Partner Amount": smms_amt,
                    "Amount Difference": 0.0,
                    "CMS Status": cms_status,
                    "Partner Status": smms_status,
                    "Reconciliation Status": "Status Mismatch",
                    "Exception Reason": f"Status mismatch between CMS ({cms_status}) and SMMS ({smms_status})",
                    "Settlement UTR": smms_r.get("Settlement UTR"),
                    "Settlement Date": smms_r.get("Settlement Date & Time"),
                    "Settlement Amount": clean_amount(smms_r.get("Net Amount")),
                    "Settlement Status": smms_r.get("Settlement Status")
                })
                self.exceptions.append(entry)
                continue

            # Both CMS and SMMS matched and Successful!
            cms_smms_entry = dict(cms_r)
            cms_smms_entry.update({
                "Source System": "CMS & SMMS",
                "Match Key": match_key,
                "Matched Transaction ID": smms_sp_id,
                "CMS Amount": cms_amt,
                "Partner Amount": smms_amt,
                "Amount Difference": 0.0,
                "CMS Status": cms_status,
                "Partner Status": smms_status,
                "Reconciliation Status": "Matched",
                "Exception Reason": "",
                "Settlement UTR": smms_r.get("Settlement UTR"),
                "Settlement Date": smms_r.get("Settlement Date & Time"),
                "Settlement Amount": clean_amount(smms_r.get("Net Amount")),
                "Settlement Status": smms_r.get("Settlement Status")
            })
            self.cms_smms_matched.append(cms_smms_entry)

            # 4. Match against Partner Report (Cashfree / Easebuzz / Airtel)
            partner_matched = False
            settle_details: Dict[str, Any] = {
                "Settlement UTR": smms_r.get("Settlement UTR"),
                "Settlement Date": smms_r.get("Settlement Date & Time"),
                "Settlement Amount": clean_amount(smms_r.get("Net Amount")),
                "Settlement Status": smms_r.get("Settlement Status")
            }

            if partner == "CashFree":
                cf_r = _find_record(cf_by_ref, smms_rrn)
                if cf_r:
                    cf_amt = clean_amount(cf_r.get("Amount"))
                    cf_status = normalize_status(cf_r.get("Transaction Status"))
                    _record_matched_key(matched_cf_keys, smms_rrn)
                    
                    if is_amount_equal(cms_amt, cf_amt, self.tolerance):
                        partner_matched = True
                        entry = dict(cms_r)
                        entry.update({
                            "Source System": "Cashfree",
                            "Match Key": smms_rrn,
                            "Matched Transaction ID": cf_r.get("Reference Id") or cf_r.get("Order Id") or "",
                            "CMS Amount": cms_amt,
                            "Partner Amount": cf_amt,
                            "Amount Difference": 0.0,
                            "CMS Status": cms_status,
                            "Partner Status": cf_status,
                            "Reconciliation Status": "Matched",
                            "Exception Reason": "",
                            "Cashfree Payment Mode": cf_r.get("Payment Mode") or "",
                            **settle_details
                        })
                        self.cms_cf_matched.append(entry)
                    else:
                        diff = round((cms_amt or 0.0) - (cf_amt or 0.0), 2)
                        entry = dict(cms_r)
                        entry.update({
                            "Source System": "Cashfree",
                            "Match Key": smms_rrn,
                            "Matched Transaction ID": cf_r.get("Reference Id") or cf_r.get("Order Id") or "",
                            "CMS Amount": cms_amt,
                            "Partner Amount": cf_amt,
                            "Amount Difference": diff,
                            "CMS Status": cms_status,
                            "Partner Status": cf_status,
                            "Reconciliation Status": "Amount Mismatch",
                            "Exception Reason": f"Amount mismatch between CMS ({cms_amt}) and Cashfree ({cf_amt})",
                            **settle_details
                        })
                        self.exceptions.append(entry)
                else:
                    # Missing in Cashfree report
                    entry = dict(cms_r)
                    entry.update({
                        "Source System": "Cashfree",
                        "Match Key": smms_rrn,
                        "Matched Transaction ID": smms_sp_id,
                        "CMS Amount": cms_amt,
                        "Partner Amount": None,
                        "Amount Difference": None,
                        "CMS Status": cms_status,
                        "Partner Status": "",
                        "Reconciliation Status": "Missing in Partner Report",
                        "Exception Reason": f"Successful CMS transaction missing in Cashfree report (RRN: {smms_rrn})",
                        **settle_details
                    })
                    self.exceptions.append(entry)

            elif partner == "EaseBuzz":
                eb_r = _find_record(eb_by_utr, smms_rrn)
                if eb_r:
                    eb_amt = clean_amount(eb_r.get("Amount"))
                    eb_status = normalize_status(eb_r.get("Status"))
                    _record_matched_key(matched_eb_keys, smms_rrn)

                    if is_amount_equal(cms_amt, eb_amt, self.tolerance):
                        partner_matched = True
                        entry = dict(cms_r)
                        entry.update({
                            "Source System": "Easebuzz",
                            "Match Key": smms_rrn,
                            "Matched Transaction ID": eb_r.get("ID") or eb_r.get("UPI tid") or "",
                            "CMS Amount": cms_amt,
                            "Partner Amount": eb_amt,
                            "Amount Difference": 0.0,
                            "CMS Status": cms_status,
                            "Partner Status": eb_status,
                            "Reconciliation Status": "Matched",
                            "Exception Reason": "",
                            **settle_details
                        })
                        self.cms_eb_matched.append(entry)
                    else:
                        diff = round((cms_amt or 0.0) - (eb_amt or 0.0), 2)
                        entry = dict(cms_r)
                        entry.update({
                            "Source System": "Easebuzz",
                            "Match Key": smms_rrn,
                            "Matched Transaction ID": eb_r.get("ID") or eb_r.get("UPI tid") or "",
                            "CMS Amount": cms_amt,
                            "Partner Amount": eb_amt,
                            "Amount Difference": diff,
                            "CMS Status": cms_status,
                            "Partner Status": eb_status,
                            "Reconciliation Status": "Amount Mismatch",
                            "Exception Reason": f"Amount mismatch between CMS ({cms_amt}) and Easebuzz ({eb_amt})",
                            **settle_details
                        })
                        self.exceptions.append(entry)
                else:
                    # Missing in Easebuzz report
                    entry = dict(cms_r)
                    entry.update({
                        "Source System": "Easebuzz",
                        "Match Key": smms_rrn,
                        "Matched Transaction ID": smms_sp_id,
                        "CMS Amount": cms_amt,
                        "Partner Amount": None,
                        "Amount Difference": None,
                        "CMS Status": cms_status,
                        "Partner Status": "",
                        "Reconciliation Status": "Missing in Partner Report",
                        "Exception Reason": f"Successful CMS transaction missing in Easebuzz report (RRN: {smms_rrn})",
                        **settle_details
                    })
                    self.exceptions.append(entry)

            elif partner == "Airtel Bank":
                air_r = _find_record(air_by_pid, smms_rrn)
                if not air_r and smms_sp_id:
                    air_r = _find_record(air_by_tid, smms_sp_id)

                # Check settlement report details
                s_r = _find_record(settle_by_ref, smms_rrn)
                if s_r:
                    settle_details = {
                        "Settlement UTR": s_r.get("UTR Num") or smms_r.get("Settlement UTR"),
                        "Settlement Date": s_r.get("Settlement Date") or smms_r.get("Settlement Date & Time"),
                        "Settlement Amount": clean_amount(s_r.get("Net Credit Amnt") or s_r.get("ORIG_AMNT")),
                        "Settlement Status": "Settled" if s_r.get("Transaction Type") == "C" else (s_r.get("Transaction Type") or "Settled")
                    }

                if air_r:
                    air_amt = clean_amount(air_r.get("Original Input Amt"))
                    air_status = normalize_status(air_r.get("Transaction Status"))
                    _record_matched_key(matched_air_keys, smms_rrn)

                    if is_amount_equal(cms_amt, air_amt, self.tolerance):
                        partner_matched = True
                        entry = dict(cms_r)
                        entry.update({
                            "Source System": "Airtel",
                            "Match Key": smms_rrn,
                            "Matched Transaction ID": air_r.get("Transaction Id") or "",
                            "CMS Amount": cms_amt,
                            "Partner Amount": air_amt,
                            "Amount Difference": 0.0,
                            "CMS Status": cms_status,
                            "Partner Status": air_status,
                            "Reconciliation Status": "Matched",
                            "Exception Reason": "",
                            **settle_details
                        })
                        self.cms_air_matched.append(entry)
                    else:
                        diff = round((cms_amt or 0.0) - (air_amt or 0.0), 2)
                        entry = dict(cms_r)
                        entry.update({
                            "Source System": "Airtel",
                            "Match Key": smms_rrn,
                            "Matched Transaction ID": air_r.get("Transaction Id") or "",
                            "CMS Amount": cms_amt,
                            "Partner Amount": air_amt,
                            "Amount Difference": diff,
                            "CMS Status": cms_status,
                            "Partner Status": air_status,
                            "Reconciliation Status": "Amount Mismatch",
                            "Exception Reason": f"Amount mismatch between CMS ({cms_amt}) and Airtel ({air_amt})",
                            **settle_details
                        })
                        self.exceptions.append(entry)
                else:
                    # Missing in Airtel report
                    entry = dict(cms_r)
                    entry.update({
                        "Source System": "Airtel",
                        "Match Key": smms_rrn,
                        "Matched Transaction ID": smms_sp_id,
                        "CMS Amount": cms_amt,
                        "Partner Amount": None,
                        "Amount Difference": None,
                        "CMS Status": cms_status,
                        "Partner Status": "",
                        "Reconciliation Status": "Missing in Partner Report",
                        "Exception Reason": f"Successful CMS transaction missing in Airtel report (RRN: {smms_rrn})",
                        **settle_details
                    })
                    self.exceptions.append(entry)

            # Record successful SMMS record with full reconciliation state for Summary calculation
            smms_recon_state = dict(smms_r)
            smms_recon_state["_partner"] = partner
            smms_recon_state["_partner_matched"] = partner_matched
            smms_recon_state["_settle_status"] = settle_details.get("Settlement Status") or smms_r.get("Settlement Status") or "Pending"
            self.successful_smms_reconciled.append(smms_recon_state)

        # 5. Check for CMS records that were never matched to SMMS
        for cms_r in self.cms_report.records:
            crrn = clean_key(cms_r.get("RRN/UTR"))
            csp = clean_key(cms_r.get("SwinkPay Txn ID"))
            c_amt = clean_amount(cms_r.get("Transaction Amount"))
            c_status = normalize_status(cms_r.get("Transaction Status"))

            # Skip empty blank rows
            if not crrn and not csp and c_amt is None and not c_status:
                continue

            if not _is_key_matched(matched_cms_keys, crrn) and not _is_key_matched(matched_cms_keys, csp):
                if c_status != "Success":
                    if not _is_key_matched(handled_failed_keys, crrn) and not _is_key_matched(handled_failed_keys, csp):
                        if crrn:
                            _record_matched_key(handled_failed_keys, crrn)
                        if csp:
                            _record_matched_key(handled_failed_keys, csp)
                        is_adj, adj_cat = _is_adjustment_record(cms_r)
                        entry = dict(cms_r)
                        entry.update({
                            "Source System": "CMS",
                            "Match Key": crrn or csp,
                            "Matched Transaction ID": csp,
                            "CMS Amount": c_amt,
                            "Partner Amount": None,
                            "Amount Difference": None,
                            "CMS Status": c_status,
                            "Partner Status": "",
                            "Reconciliation Status": adj_cat if is_adj else c_status,
                            "Adjustment Category": adj_cat if is_adj else "",
                            "Exception Reason": (f"{adj_cat}: {cms_r.get('Transaction Description') or 'CMS Adjustment'}" if is_adj else "Failed or reversed CMS transaction without SMMS record"),
                            "Settlement Details": ""
                        })
                        if is_adj:
                            self.adjustments.append(entry)
                        self.failed_or_reversed.append(entry)
                else:
                    entry = dict(cms_r)
                    entry.update({
                        "Source System": "CMS",
                        "Match Key": crrn or csp,
                        "Matched Transaction ID": csp,
                        "CMS Amount": c_amt,
                        "Partner Amount": None,
                        "Amount Difference": None,
                        "CMS Status": c_status,
                        "Partner Status": "",
                        "Reconciliation Status": "Missing in SMMS",
                        "Exception Reason": f"Transaction in CMS but not found in SMMS (RRN: {crrn})",
                        "Settlement Details": ""
                    })
                    self.cms_not_in_smms.append(entry)
                    self.exceptions.append(entry)

        # 6. Check for Unmatched Partner records
        # Cashfree unmatched
        if self.cf_report:
            for cf_r in self.cf_report.records:
                bref = clean_key(cf_r.get("Bank Reference No."))
                if bref and not _is_key_matched(matched_cf_keys, bref):
                    cf_status = normalize_status(cf_r.get("Transaction Status"))
                    cf_amt = clean_amount(cf_r.get("Amount"))
                    entry = dict(cf_r)
                    entry.update({
                        "Source System": "Cashfree",
                        "Match Key": bref,
                        "Matched Transaction ID": cf_r.get("Reference Id") or cf_r.get("Order Id") or "",
                        "CMS Amount": None,
                        "Partner Amount": cf_amt,
                        "Amount Difference": None,
                        "CMS Status": "",
                        "Partner Status": cf_status,
                        "Reconciliation Status": "Missing in CMS",
                        "Exception Reason": f"Cashfree transaction not found in CMS/SMMS (RRN: {bref})",
                        "Settlement Details": ""
                    })
                    if cf_status in ("Failed", "Reversed"):
                        if not _is_key_matched(handled_failed_keys, bref):
                            _record_matched_key(handled_failed_keys, bref)
                            self.failed_or_reversed.append(entry)
                    else:
                        self.unmatched_cf.append(entry)
                        self.exceptions.append(entry)

        # Easebuzz unmatched
        if self.eb_report:
            for eb_r in self.eb_report.records:
                utr = clean_key(eb_r.get("UTR"))
                if utr and not _is_key_matched(matched_eb_keys, utr):
                    eb_status = normalize_status(eb_r.get("Status"))
                    eb_amt = clean_amount(eb_r.get("Amount"))
                    entry = dict(eb_r)
                    entry.update({
                        "Source System": "Easebuzz",
                        "Match Key": utr,
                        "Matched Transaction ID": eb_r.get("ID") or eb_r.get("UPI tid") or "",
                        "CMS Amount": None,
                        "Partner Amount": eb_amt,
                        "Amount Difference": None,
                        "CMS Status": "",
                        "Partner Status": eb_status,
                        "Reconciliation Status": "Missing in CMS" if eb_status == "Success" else eb_status,
                        "Exception Reason": f"Easebuzz transaction not in matched CMS (UTR: {utr})",
                        "Settlement Details": ""
                    })
                    if eb_status in ("Failed", "Reversed"):
                        if not _is_key_matched(handled_failed_keys, utr):
                            _record_matched_key(handled_failed_keys, utr)
                            self.failed_or_reversed.append(entry)
                    else:
                        self.unmatched_eb.append(entry)
                        self.exceptions.append(entry)

        # Airtel unmatched
        if self.airtel_report:
            for air_r in self.airtel_report.records:
                pid = clean_key(air_r.get("PARTNER_TXN_ID"))
                if pid and not _is_key_matched(matched_air_keys, pid):
                    air_status = normalize_status(air_r.get("Transaction Status"))
                    air_amt = clean_amount(air_r.get("Original Input Amt"))
                    entry = dict(air_r)
                    entry.update({
                        "Source System": "Airtel",
                        "Match Key": pid,
                        "Matched Transaction ID": air_r.get("Transaction Id") or "",
                        "CMS Amount": None,
                        "Partner Amount": air_amt,
                        "Amount Difference": None,
                        "CMS Status": "",
                        "Partner Status": air_status,
                        "Reconciliation Status": "Missing in CMS" if air_status == "Success" else air_status,
                        "Exception Reason": f"Airtel transaction not found in CMS/SMMS (RRN: {pid})",
                        "Settlement Details": ""
                    })
                    if air_status in ("Failed", "Reversed"):
                        if not _is_key_matched(handled_failed_keys, pid):
                            _record_matched_key(handled_failed_keys, pid)
                            self.failed_or_reversed.append(entry)
                    else:
                        self.unmatched_airtel.append(entry)
                        self.exceptions.append(entry)

        # 8. Detect Network Changes (CMS Network vs Cashfree Payment Mode)
        from .network_builder import detect_network_changes
        self.network_changes = detect_network_changes(
            self.cms_cf_matched,
            cms_records=self.cms_report.records if self.cms_report else None,
            cf_records=self.cf_report.records if self.cf_report else None,
            cms_not_in_smms=self.cms_not_in_smms
        )
