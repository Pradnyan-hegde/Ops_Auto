"""
Terminal Reference Mapping Module.
Parses Terminal Report Excel files (TID_FILE) containing:
- TERMINAL ID
- MMS TERMINAL ID
- Partner Ref ID
- VPA

Maps Cashfree Order ID middle numbers (e.g. '4860' from '330595-4860-...')
to their authoritative MMS Terminal ID (e.g. 'X47VVH' / 'X1JYYO').
Persists mappings to data/terminal_mappings.json for instant lookups across sessions.
"""
import os
import re
import json
from typing import Dict, Any, Optional, Tuple, List
from datetime import datetime
import openpyxl

from .normalizer import clean_key


DATA_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data"))
MAPPINGS_FILE = os.path.join(DATA_DIR, "terminal_mappings.json")

# In-memory cached mappings
_CACHE: Optional[Dict[str, Any]] = None


def _normalize_header(val: Any) -> str:
    """Normalizes header text for robust comparison."""
    if val is None:
        return ""
    s = str(val).strip().lower()
    s = re.sub(r'[\s_\-/]+', ' ', s)
    return s.strip()


def parse_terminal_report(file_path: str) -> Dict[str, Any]:
    """
    Parses a Terminal Report workbook (.xlsx).
    Finds the header row containing 'TERMINAL ID', 'MMS TERMINAL ID', and 'Partner Ref ID'.
    Returns structured mappings indexed by Partner Ref ID, VPA, and Terminal ID.
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    
    # Prefer TID_FILE sheet if present, else active sheet
    sheet_name = None
    for name in wb.sheetnames:
        if "tid" in name.lower() or "terminal" in name.lower():
            sheet_name = name
            break
    ws = wb[sheet_name] if sheet_name else wb.active

    header_row_idx = None
    col_map = {}

    # Scan rows to locate header row (up to 2000 rows)
    row_count = 0
    for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
        row_count += 1
        if row_count > 2000:
            break

        norm_cells = [_normalize_header(c) for c in row]
        norm_dict = {norm_cells[i]: i for i in range(len(norm_cells)) if norm_cells[i]}

        # Check for key columns
        has_mms = any("mms terminal id" in k or "mms terminal" in k or "merchant mms terminal" in k for k in norm_dict)
        has_tid = any(k == "terminal id" or k == "terminalid" for k in norm_dict)
        has_pref = any("partner ref id" in k or "partner ref" in k or "ref id" in k for k in norm_dict)
        has_vpa = any("vpa" in k for k in norm_dict)

        if has_mms and (has_tid or has_pref or has_vpa):
            header_row_idx = r_idx
            # Map canonical column names to indices
            for k, idx in norm_dict.items():
                if "mms terminal id" in k or "merchant mms terminal id" in k:
                    col_map["mms_terminal_id"] = idx
                elif k == "terminal id" or k == "terminalid":
                    col_map["terminal_id"] = idx
                elif "partner ref id" in k or "partner ref" in k:
                    col_map["partner_ref_id"] = idx
                elif k == "vpa" or "vpa" in k:
                    col_map["vpa"] = idx
                elif "merchant id" in k:
                    col_map["merchant_id"] = idx
                elif "merchant name" in k:
                    col_map["merchant_name"] = idx
            break

    if header_row_idx is None or "mms_terminal_id" not in col_map:
        wb.close()
        raise ValueError("Could not find required columns (MMS TERMINAL ID, Partner Ref ID / TERMINAL ID) in terminal file.")

    mappings_by_ref_id: Dict[str, Dict[str, str]] = {}
    mappings_by_vpa: Dict[str, str] = {}
    mappings_by_terminal_id: Dict[str, Dict[str, str]] = {}
    total_records = 0

    # Parse data rows
    for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if r_idx <= header_row_idx:
            continue

        mms_tid = clean_key(row[col_map["mms_terminal_id"]]) if "mms_terminal_id" in col_map and col_map["mms_terminal_id"] < len(row) else ""
        tid = clean_key(row[col_map["terminal_id"]]) if "terminal_id" in col_map and col_map["terminal_id"] < len(row) else ""
        pref_id = clean_key(row[col_map["partner_ref_id"]]) if "partner_ref_id" in col_map and col_map["partner_ref_id"] < len(row) else ""
        vpa = str(row[col_map["vpa"]]).strip() if "vpa" in col_map and col_map["vpa"] < len(row) and row[col_map["vpa"]] is not None else ""

        if not mms_tid and not tid:
            continue
        if mms_tid.lower() in ("mms terminal id", "terminal id", "mms_terminal_id", "merchant mms terminal id", "slno", "sr no"):
            continue

        effective_mms = mms_tid or tid
        effective_tid = tid or mms_tid

        record_payload = {
            "mms_terminal_id": effective_mms,
            "terminal_id": effective_tid,
            "partner_ref_id": pref_id,
            "vpa": vpa
        }
        total_records += 1

        # Index by Partner Ref ID (e.g. '4860', '4030', '4803', '138960')
        if pref_id and pref_id != "--":
            mappings_by_ref_id[pref_id] = record_payload

        # Index by VPA text or numbers within VPA (e.g. 'tn= 4860' -> '4860')
        if vpa and vpa != "--":
            clean_vpa = clean_key(vpa)
            if clean_vpa:
                mappings_by_vpa[clean_vpa] = effective_mms
            # Extract standalone numbers (e.g. 4860) from VPA string like 'tn= 4860'
            nums = re.findall(r'\b\d{3,8}\b', vpa)
            for n in nums:
                clean_n = clean_key(n)
                if clean_n and clean_n not in mappings_by_ref_id:
                    mappings_by_ref_id[clean_n] = record_payload
                if clean_n:
                    mappings_by_vpa[clean_n] = effective_mms

        # Index by Terminal ID
        if tid and tid != "--":
            mappings_by_terminal_id[tid] = record_payload

    wb.close()

    result = {
        "updated_at": datetime.now().isoformat(),
        "source_filename": os.path.basename(file_path),
        "total_rows_scanned": total_records,
        "ref_id_count": len(mappings_by_ref_id),
        "terminal_id_count": len(mappings_by_terminal_id),
        "mappings_by_ref_id": mappings_by_ref_id,
        "mappings_by_vpa": mappings_by_vpa,
        "mappings_by_terminal_id": mappings_by_terminal_id
    }
    return result


def save_terminal_mappings(mappings_data: Dict[str, Any]) -> str:
    """Saves parsed terminal mappings to disk and updates in-memory cache."""
    global _CACHE
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(MAPPINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(mappings_data, f, indent=2)
    _CACHE = mappings_data
    return MAPPINGS_FILE


def load_terminal_mappings(force_reload: bool = False) -> Dict[str, Any]:
    """Loads terminal mappings from disk cache or memory."""
    global _CACHE
    if _CACHE is not None and not force_reload:
        return _CACHE

    if os.path.exists(MAPPINGS_FILE):
        try:
            with open(MAPPINGS_FILE, "r", encoding="utf-8") as f:
                _CACHE = json.load(f)
                return _CACHE
        except Exception:
            pass

    _CACHE = {
        "updated_at": None,
        "source_filename": None,
        "mappings_by_ref_id": {},
        "mappings_by_vpa": {},
        "mappings_by_terminal_id": {}
    }
    return _CACHE


def load_and_save_terminal_file(file_path: str) -> Dict[str, Any]:
    """Convenience function: parses an uploaded file and saves it immediately."""
    parsed = parse_terminal_report(file_path)
    save_terminal_mappings(parsed)
    return {
        "success": True,
        "source_filename": parsed.get("source_filename"),
        "total_mappings": len(parsed.get("mappings_by_ref_id", {})),
        "terminal_count": len(parsed.get("mappings_by_terminal_id", {})),
        "updated_at": parsed.get("updated_at")
    }


def extract_cf_middle_number(order_id: str) -> str:
    """
    Extracts the middle number from Cashfree Order IDs:
    e.g. '330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline' -> '4860'
    e.g. '330595-4873-AXLa3e2d...' -> '4873'
    """
    if not order_id:
        return ""
    parts = str(order_id).strip().split("-")
    if len(parts) >= 2:
        return clean_key(parts[1])
    return ""


def resolve_mms_terminal_id(identifier: str, mappings: Optional[Dict[str, Any]] = None) -> Optional[str]:
    """
    Resolves an identifier (Order ID, middle number, Partner Ref ID, or Terminal ID)
    to its corresponding MMS Terminal ID.
    """
    if not identifier:
        return None

    if mappings is None:
        mappings = load_terminal_mappings()

    by_ref = mappings.get("mappings_by_ref_id", {})
    by_vpa = mappings.get("mappings_by_vpa", {})
    by_tid = mappings.get("mappings_by_terminal_id", {})

    raw_str = str(identifier).strip()

    # 1. Try extracting middle number if hyphenated Order ID
    if "-" in raw_str:
        mid = extract_cf_middle_number(raw_str)
        if mid:
            if mid in by_ref:
                rec = by_ref[mid]
                return rec.get("mms_terminal_id") or rec.get("terminal_id")
            if mid in by_vpa:
                return by_vpa[mid]
            if mid in by_tid:
                rec = by_tid[mid]
                return rec.get("mms_terminal_id") or rec.get("terminal_id")

    # 2. Try clean key direct lookup
    clean_id = clean_key(raw_str)
    if clean_id in by_ref:
        rec = by_ref[clean_id]
        return rec.get("mms_terminal_id") or rec.get("terminal_id")

    if clean_id in by_vpa:
        return by_vpa[clean_id]

    if clean_id in by_tid:
        rec = by_tid[clean_id]
        return rec.get("mms_terminal_id") or rec.get("terminal_id")

    return None


def clear_terminal_mappings() -> bool:
    """Clears persisted mappings and resets in-memory cache."""
    global _CACHE
    if os.path.exists(MAPPINGS_FILE):
        try:
            os.remove(MAPPINGS_FILE)
        except Exception:
            pass
    _CACHE = {
        "updated_at": None,
        "source_filename": None,
        "mappings_by_ref_id": {},
        "mappings_by_vpa": {},
        "mappings_by_terminal_id": {}
    }
    return True
