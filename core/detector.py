"""
Report Type Detector and Header Validator.
Identifies reports based on header signatures across the first few rows.
"""
from enum import Enum
from typing import List, Tuple, Optional, Set, Dict


class ReportType(str, Enum):
    CMS = "CMS"
    SMMS = "SMMS"
    CASHFREE = "Cashfree"
    EASEBUZZ = "Easebuzz"
    AIRTEL = "Airtel"
    AIRTEL_SETTLEMENT = "Airtel Settlement"
    UNKNOWN = "Unknown"


class MissingColumnsException(Exception):
    """Raised when a detected report is missing mandatory columns."""
    def __init__(self, report_type: ReportType, missing_columns: List[str], file_path: str = ""):
        self.report_type = report_type
        self.missing_columns = missing_columns
        self.file_path = file_path
        msg = (f"Report layout validation failed for {report_type.value}: "
               f"Missing required column(s): {', '.join(missing_columns)}. "
               f"Matching halted for this report. Check column names in {file_path or 'uploaded file'}.")
        super().__init__(msg)


# Canonical required columns for each report type
REQUIRED_COLUMNS: Dict[ReportType, List[str]] = {
    ReportType.CMS: [
        "RRN/UTR",
        "SwinkPay Txn ID",
        "Transaction Amount",
        "Transaction Status",
        "Network",
        "Transaction Date & Time",
    ],
    ReportType.SMMS: [
        "RRN/UTR",
        "SwinkPay Txn ID",
        "Transaction Amount",
        "Transaction Status",
        "Network",
        "Transaction Date & Time",
        "PG/Bank",
    ],
    ReportType.CASHFREE: [
        "Bank Reference No.",
        "Amount",
        "Settlement Amount",
        "Payment Mode",
        "Transaction Status",
        "Transaction Time",
    ],
    ReportType.EASEBUZZ: [
        "UTR",
        "Amount",
        "Platform Charges",
        "GST Amount",
        "Payment Mode",
        "Transaction Date",
    ],
    ReportType.AIRTEL: [
        "PARTNER_TXN_ID",
        "Original Input Amt",
        "Commission(DR)",
        "Transaction Status",
        "Date and Time",
    ],
    ReportType.AIRTEL_SETTLEMENT: [
        "Partner Txn Ref No",
        "UTR Num",
        "Settlement Date",
        "Net Credit Amnt",
    ],
}

# Distinctive signature header keywords used for scoring and identification
SIGNATURE_KEYWORDS: Dict[ReportType, Set[str]] = {
    ReportType.SMMS: {
        "pg/bank", "psp charges", "psp amount", "sp charges", "sp amount", 
        "deduction", "net amount", "swinkpay txn id", "mcc description"
    },
    ReportType.CMS: {
        "smms sync status", "merchant mms id", "merchant mms terminal id",
        "merchant details[mobile/email]", "customer info", "payment gateway"
    },
    ReportType.CASHFREE: {
        "bank reference no.", "order id", "order note", "settlement amount",
        "card scheme", "service charge", "st/gst"
    },
    ReportType.EASEBUZZ: {
        "virtual account label", "virtual account number", "virtual upi handle",
        "platform charges", "platform charges with gst", "remitter name", "remitter upi handle"
    },
    ReportType.AIRTEL_SETTLEMENT: {
        "partner txn ref no", "bank txn ref no", "utr num", "settlement date",
        "net credit amnt", "merchant settlement type"
    },
    ReportType.AIRTEL: {
        "partner_txn_id", "original input amt", "commission(dr)", "net amount payable(cr)",
        "net amount payable(dr)", "ref_txn_no_org", "shop display name"
    },
}


def _normalize_header(h: Optional[str]) -> str:
    if h is None:
        return ""
    return str(h).strip().lower()


def detect_report_type(raw_rows: List[List[Optional[str]]]) -> Tuple[ReportType, int, List[str]]:
    """
    Scans the first several rows of a table to find the header row and determine the report type.
    Returns: (ReportType, header_row_index, list_of_headers)
    """
    if not raw_rows:
        return ReportType.UNKNOWN, -1, []

    max_scan = min(10, len(raw_rows))
    best_type = ReportType.UNKNOWN
    best_row_idx = -1
    best_headers: List[str] = []
    highest_score = 0

    for r_idx in range(max_scan):
        row = raw_rows[r_idx]
        if not row:
            continue
        
        # Normalized headers for this candidate row
        headers_norm = [_normalize_header(cell) for cell in row]
        headers_set = set(h for h in headers_norm if h)
        if len(headers_set) < 3:
            continue

        # Score each candidate report type
        for rep_type, sig_set in SIGNATURE_KEYWORDS.items():
            matched_keywords = headers_set.intersection(sig_set)
            score = len(matched_keywords)

            # Extra specificity rules to disambiguate CMS vs SMMS
            if rep_type == ReportType.SMMS:
                if "pg/bank" in headers_set:
                    score += 5
                if "psp charges" in headers_set or "deduction" in headers_set:
                    score += 3
            elif rep_type == ReportType.CMS:
                if "smms sync status" in headers_set:
                    score += 5
                if "customer info" in headers_set:
                    score += 2
                # If PG/Bank is present, it's SMMS, not CMS
                if "pg/bank" in headers_set:
                    score = 0
            elif rep_type == ReportType.AIRTEL_SETTLEMENT:
                if "utr num" in headers_set and "settlement date" in headers_set:
                    score += 5
            elif rep_type == ReportType.AIRTEL:
                if "partner_txn_id" in headers_set and "original input amt" in headers_set:
                    score += 5

            if score > highest_score and score >= 2:
                highest_score = score
                best_type = rep_type
                best_row_idx = r_idx
                best_headers = [str(c).strip() if c is not None else f"Col_{i}" for i, c in enumerate(row)]

    return best_type, best_row_idx, best_headers


def validate_report_headers(report_type: ReportType, headers: List[str], file_path: str = "") -> List[str]:
    """
    Validates that all mandatory columns are present in the headers.
    Returns list of missing columns (empty if valid).
    Raises MissingColumnsException if any mandatory columns are missing.
    """
    if report_type not in REQUIRED_COLUMNS:
        return []

    req_cols = REQUIRED_COLUMNS[report_type]
    headers_norm_map = {_normalize_header(h): h for h in headers}
    
    missing: List[str] = []
    for req in req_cols:
        req_norm = _normalize_header(req)
        if req_norm not in headers_norm_map:
            # Check for close variations if necessary, else mark missing
            missing.append(req)

    if missing:
        raise MissingColumnsException(report_type, missing, file_path)

    return []
