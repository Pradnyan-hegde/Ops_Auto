"""
Normalization utilities for Keys, Amounts, Statuses, and Dates.
Strictly distinguishes missing/blank from 0.00 and preserves leading zeros.
"""
import re
from typing import Optional, Any
from datetime import datetime


def clean_key(val: Any) -> str:
    """
    Cleans RRN, UTR, and Txn IDs:
    - Strips whitespace
    - Strips leading apostrophe/quotes
    - Strips trailing '.0' from float conversions
    - Preserves leading zeros (e.g. '090984560431')
    """
    if val is None:
        return ""
    
    # If already an integer or float
    if isinstance(val, int):
        return str(val)
    elif isinstance(val, float):
        # check if it's an integer float like 12345.0
        if val.is_integer():
            return str(int(val))
        return f"{val:.0f}"

    s = str(val).strip()
    if not s or s == "--" or s.lower() == "null" or s.lower() == "none":
        return ""

    # Remove leading apostrophe or quotes often used in Excel for text formatting
    if s.startswith("'") or s.startswith('"'):
        s = s[1:].strip()
    if s.endswith('"'):
        s = s[:-1].strip()

    # If it ends with .0 and remainder is all digits
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]

    # Handle scientific notation like 4.72696E+11
    if "e" in s.lower() and ("+" in s or "-" in s):
        try:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
        except ValueError:
            pass

    return s


def clean_amount(val: Any) -> Optional[float]:
    """
    Parses and standardizes amounts.
    CRITICAL: Returns None for missing/blank values; never converts None to 0.0.
    """
    if val is None:
        return None

    if isinstance(val, (int, float)):
        return round(float(val), 2)

    s = str(val).strip()
    if not s or s == "--" or s.lower() == "null" or s.lower() == "none":
        return None

    # Remove currency symbols (Rs, INR, ₹, $, etc.), commas, spaces
    cleaned = re.sub(r"(?i)\b(rs\.?|inr)\b|[₹$,\s]", "", s)
    try:
        return round(float(cleaned), 2)
    except ValueError:
        m = re.search(r"[-+]?\d+(?:\.\d+)?", cleaned)
        if m:
            try:
                return round(float(m.group(0)), 2)
            except ValueError:
                return None
        return None


def is_amount_equal(amt1: Optional[float], amt2: Optional[float], tolerance: float = 0.01) -> bool:
    """
    Checks amount equality within a configurable tolerance.
    If either is None, they are only equal if both are None.
    """
    if amt1 is None and amt2 is None:
        return True
    if amt1 is None or amt2 is None:
        return False
    return abs(amt1 - amt2) <= tolerance


def normalize_status(val: Any) -> str:
    """
    Normalizes status strings/codes to standard forms:
    Success, Failed, Reversed, Pending, Needs Review.
    """
    if val is None:
        return ""

    if isinstance(val, bool):
        return "Success" if val else "Failed"
    if isinstance(val, (int, float)):
        ival = int(val)
        if ival == 1:
            return "Success"
        elif ival in (0, 2):
            return "Failed"
        elif ival in (3, 4, 5, 6):
            return f"Status {ival}"

    s = str(val).strip()
    if not s or s == "--" or s.lower() in ("none", "null"):
        return ""

    if s in ("1", "1.0"):
        return "Success"
    elif s in ("2", "2.0", "0", "0.0"):
        return "Failed"
    elif s in ("3", "3.0"):
        return "Status 3"
    elif s in ("4", "4.0"):
        return "Status 4"
    elif s in ("5", "5.0"):
        return "Status 5"
    elif s in ("6", "6.0"):
        return "Status 6"

    sl = s.lower()
    if sl in ("success", "successful", "captured", "completed", "paid", "received", "settled", "c", "misc cr", "true"):
        return "Success"
    elif sl in ("failed", "failure", "fail", "declined", "f", "rejected", "false"):
        return "Failed"
    elif sl in ("reversed", "refunded", "refund", "r"):
        return "Reversed"
    elif sl in ("pending", "in_process", "initiated"):
        return "Pending"

    return s


def extract_iso_date(val: Any) -> Optional[str]:
    """
    Extracts YYYY-MM-DD from various date formats:
    - datetime objects
    - '2026-09-10 00:07:46'
    - '10-09-2026' or '10/09/26'
    - '10-SEP-2026'
    """
    if val is None:
        return None

    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")

    s = str(val).strip()
    if not s:
        return None

    # Try ISO YYYY-MM-DD
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # Try DD-MM-YYYY or DD/MM/YYYY
    m = re.search(r"(\d{2})[-/](\d{2})[-/](\d{4})", s)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"

    # Try DD/MM/YY
    m = re.search(r"(\d{2})[-/](\d{2})[-/](\d{2})", s)
    if m:
        year = int(m.group(3))
        full_year = 2000 + year if year < 50 else 1900 + year
        return f"{full_year}-{m.group(2)}-{m.group(1)}"

    # Try DD-MON-YYYY (e.g. 10-SEP-2026)
    month_names = {
        "jan": "01", "feb": "02", "mar": "03", "apr": "04",
        "may": "05", "jun": "06", "jul": "07", "aug": "08",
        "sep": "09", "oct": "10", "nov": "11", "dec": "12"
    }
    m = re.search(r"(\d{1,2})[-/\s]([A-Za-z]{3})[-/\s](\d{4})", s)
    if m:
        day = f"{int(m.group(1)):02d}"
        mon = month_names.get(m.group(2).lower(), "01")
        year = m.group(3)
        return f"{year}-{mon}-{day}"

    return None
