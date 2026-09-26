"""
Merchant Profiles and Configuration for Ops_Auto Reconciliation.
Supports CCD Value Express, SBB Medicare, and dynamic registration of custom merchants.
Includes persistent registry storage in data/merchants.json.
"""
import os
import json
import re
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
MERCHANTS_FILE = os.path.join(DATA_DIR, "merchants.json")


@dataclass
class MerchantProfile:
    key: str
    display_name: str
    merchant_ids: List[str] = field(default_factory=list)
    merchant_keywords: List[str] = field(default_factory=list)
    gateways: List[str] = field(default_factory=lambda: ["CashFree", "EaseBuzz", "Airtel Bank"])
    has_split_settlement: bool = False
    softpos_label: str = "CashFree"
    input_file_prefix: str = ""
    terminal_file_name: str = ""
    terminal_mappings_count: int = 0
    upi_fee_rate: float = 0.0015
    cc_fee_rate: float = 0.0225
    default_gst_rate: float = 0.18

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "MerchantProfile":
        valid_keys = {
            "key", "display_name", "merchant_ids", "merchant_keywords",
            "gateways", "has_split_settlement", "softpos_label",
            "input_file_prefix", "terminal_file_name", "terminal_mappings_count",
            "upi_fee_rate", "cc_fee_rate", "default_gst_rate"
        }
        filtered = {k: v for k, v in data.items() if k in valid_keys}
        return cls(**filtered)


# Built-in baseline merchant profiles
DEFAULT_MERCHANT_PROFILES: Dict[str, MerchantProfile] = {
    "xcd": MerchantProfile(
        key="xcd",
        display_name="Cafe Value Express (XCD)",
        merchant_ids=["0000000000002045", "2045"],
        merchant_keywords=["VALUE EXPRESS", "CAFE VALUE EXPRESS", "XCD"],
        gateways=["CashFree", "EaseBuzz", "Airtel Bank"],
        has_split_settlement=False,  # Single Daily Settlement
        softpos_label="CashFree",
        input_file_prefix="XCD Input file as on",
        upi_fee_rate=0.01,    # 1.00%
        cc_fee_rate=0.0275,   # 2.75%
        default_gst_rate=0.18
    ),
    "ccd": MerchantProfile(
        key="ccd",
        display_name="Cafe Coffee Day (CCD)",
        merchant_ids=["0000000000002045", "2045"],
        merchant_keywords=["CAFE COFFEE DAY", "COFFEE DAY", "CCD"],
        gateways=["CashFree", "EaseBuzz", "Airtel Bank"],
        has_split_settlement=False,  # Single Daily Settlement
        softpos_label="CashFree",
        input_file_prefix="CCD Input file as on",
        upi_fee_rate=0.01,    # 1.00%
        cc_fee_rate=0.0275,   # 2.75%
        default_gst_rate=0.18
    ),
    "ags": MerchantProfile(
        key="ags",
        display_name="Advance Genuine Spares (AGS)",
        merchant_ids=["0000000000002095", "000000000002095", "2095"],
        merchant_keywords=["ADVANCE GENUINE SPARES", "GENUINE SPARES", "ADVANCE SPARES", "AGS"],
        gateways=["CashFree", "EaseBuzz", "Airtel Bank"],
        has_split_settlement=False,  # Single Daily Settlement
        softpos_label="CashFree",
        input_file_prefix="AGS Input file as on",
        upi_fee_rate=0.0015,  # 0.15%
        cc_fee_rate=0.0225,   # 2.25%
        default_gst_rate=0.18
    ),
    "sbb": MerchantProfile(
        key="sbb",
        display_name="SBB Medicare",
        merchant_ids=["0000000000002145", "2145"],
        merchant_keywords=["SBB", "MEDICARE", "SBB MEDICARE"],
        gateways=["CashFree", "EaseBuzz"],
        has_split_settlement=True,  # 2 settlements per day (12 AM-12 PM & 12 PM-12 AM)
        softpos_label="CF_SoftPOS",
        input_file_prefix="SBB_INPUTFILE",
        upi_fee_rate=0.0015,  # 0.15%
        cc_fee_rate=0.0225,   # 2.25%
        default_gst_rate=0.18
    )
}


def load_all_merchant_profiles() -> Dict[str, MerchantProfile]:
    """Loads default merchant profiles plus any custom merchants stored in data/merchants.json."""
    profiles = dict(DEFAULT_MERCHANT_PROFILES)
    deleted_keys = []
    saved_custom_keys = set()
    if os.path.exists(MERCHANTS_FILE):
        try:
            with open(MERCHANTS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if isinstance(saved, dict):
                    deleted_keys = [str(k).lower() for k in saved.get("__deleted__", [])]
                    for k, v in saved.items():
                        if k != "__deleted__" and isinstance(v, dict):
                            profiles[k.lower()] = MerchantProfile.from_dict(v)
                            saved_custom_keys.add(k.lower())
        except Exception as e:
            print(f"[MerchantRegistry] Warning loading {MERCHANTS_FILE}: {e}")

    for dk in deleted_keys:
        if dk in profiles and dk not in saved_custom_keys:
            del profiles[dk]

    return profiles


def delete_merchant_profile(key: str) -> bool:
    """
    Permanently deletes a merchant profile by key from persistent storage.
    Works for both custom and default merchants.
    """
    os.makedirs(DATA_DIR, exist_ok=True)
    key_lower = key.lower().strip()
    saved = {}
    deleted_keys = []

    if os.path.exists(MERCHANTS_FILE):
        try:
            with open(MERCHANTS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            deleted_keys = [str(k).lower() for k in saved.get("__deleted__", [])]
        except Exception:
            saved = {}

    found = False
    if key_lower in saved:
        del saved[key_lower]
        found = True

    if key_lower in DEFAULT_MERCHANT_PROFILES or found:
        if key_lower not in deleted_keys:
            deleted_keys.append(key_lower)
        found = True

    saved["__deleted__"] = deleted_keys

    try:
        with open(MERCHANTS_FILE, "w", encoding="utf-8") as f:
            json.dump(saved, f, indent=2)
        return found
    except Exception as e:
        print(f"[MerchantRegistry] Error deleting merchant '{key}': {e}")
        return False


def save_merchant_profile(data: Dict[str, Any]) -> MerchantProfile:
    """
    Registers or updates a merchant profile and persists it to data/merchants.json.
    """
    os.makedirs(DATA_DIR, exist_ok=True)
    profiles = load_all_merchant_profiles()

    raw_name = str(data.get("name") or data.get("display_name") or "").strip()
    raw_key = str(data.get("key") or "").strip()

    if not raw_key and raw_name:
        raw_key = re.sub(r"[^a-zA-Z0-9]+", "_", raw_name).strip("_").lower()
    elif not raw_key:
        raw_key = "custom_merchant"

    display_name = raw_name or raw_key.replace("_", " ").title()

    # Input file prefix defaults to {SHORT_NAME}_INPUTFILE
    short_code = re.sub(r"[^a-zA-Z0-9]+", "_", display_name).strip("_").upper()
    prefix = str(data.get("input_file_prefix") or "").strip()
    if not prefix:
        prefix = f"{short_code}_INPUTFILE"

    # Keywords
    keywords = data.get("merchant_keywords") or []
    if not keywords:
        keywords = [display_name.upper(), raw_key.upper()]
    else:
        keywords = [str(k).upper() for k in keywords]

    m_ids = [str(m).strip() for m in (data.get("merchant_ids") or []) if str(m).strip()]
    single_mid = str(data.get("merchant_id") or data.get("primary_merchant_id") or "").strip()
    if single_mid and single_mid not in m_ids:
        m_ids.insert(0, single_mid)

    # Prefix can come as merchant_prefix or input_file_prefix
    if (not prefix or prefix == f"{short_code}_INPUTFILE") and data.get("merchant_prefix"):
        pref_cand = str(data.get("merchant_prefix")).strip()
        if pref_cand:
            prefix = pref_cand

    term_fn = str(data.get("terminal_file_name") or "").strip()
    term_cnt = int(data.get("terminal_mappings_count") or 0)

    gateways = data.get("gateways") or ["CashFree", "EaseBuzz", "Airtel Bank"]
    has_split = bool(data.get("has_split_settlement", False))
    softpos = str(data.get("softpos_label") or ("CF_SoftPOS" if has_split else "CashFree"))

    try:
        upi_rate = float(data.get("upi_fee_rate", 0.0015))
    except (ValueError, TypeError):
        upi_rate = 0.0015

    try:
        cc_rate = float(data.get("cc_fee_rate", 0.0225))
    except (ValueError, TypeError):
        cc_rate = 0.0225

    try:
        gst_rate = float(data.get("default_gst_rate", 0.18))
    except (ValueError, TypeError):
        gst_rate = 0.18

    profile = MerchantProfile(
        key=raw_key.lower(),
        display_name=display_name,
        merchant_ids=m_ids,
        merchant_keywords=keywords,
        gateways=gateways,
        has_split_settlement=has_split,
        softpos_label=softpos,
        input_file_prefix=prefix,
        terminal_file_name=term_fn,
        terminal_mappings_count=term_cnt,
        upi_fee_rate=upi_rate,
        cc_fee_rate=cc_rate,
        default_gst_rate=gst_rate
    )

    profiles[profile.key] = profile

    # Save custom profiles
    custom_dict = {}
    for k, p in profiles.items():
        custom_dict[k] = p.to_dict()

    # Ensure un-deleted if re-saved
    deleted_keys = []
    if os.path.exists(MERCHANTS_FILE):
        try:
            with open(MERCHANTS_FILE, "r", encoding="utf-8") as f:
                prev = json.load(f)
            deleted_keys = [str(k).lower() for k in prev.get("__deleted__", [])]
            if profile.key in deleted_keys:
                deleted_keys.remove(profile.key)
        except Exception:
            pass

    custom_dict["__deleted__"] = deleted_keys

    try:
        with open(MERCHANTS_FILE, "w", encoding="utf-8") as f:
            json.dump(custom_dict, f, indent=2)
    except Exception as e:
        print(f"[MerchantRegistry] Error writing {MERCHANTS_FILE}: {e}")

    return profile


def list_all_merchants() -> List[Dict[str, Any]]:
    """Returns a list of all available merchants formatted for UI dropdowns and API responses."""
    profiles = load_all_merchant_profiles()
    results = []
    for k in sorted(profiles.keys()):
        p = profiles[k]
        results.append({
            "key": p.key,
            "display_name": p.display_name,
            "input_file_prefix": p.input_file_prefix,
            "has_split_settlement": p.has_split_settlement,
            "merchant_ids": p.merchant_ids,
            "primary_merchant_id": p.merchant_ids[0] if p.merchant_ids else "",
            "terminal_file_name": getattr(p, "terminal_file_name", ""),
            "terminal_mappings_count": getattr(p, "terminal_mappings_count", 0),
            "gateways": p.gateways,
            "softpos_label": p.softpos_label,
            "upi_fee_rate": p.upi_fee_rate,
            "cc_fee_rate": p.cc_fee_rate
        })
    return results


def detect_merchant(
    records: Optional[List[Dict[str, Any]]] = None,
    filenames: Optional[List[str]] = None,
    smms_records: Optional[List[Dict[str, Any]]] = None
) -> str:
    """
    Auto-detects merchant key from uploaded filenames, SMMS records, or CMS records
    by checking 'Merchant Name', 'Merchant ID', and filename patterns against all registered profiles.
    If an unknown merchant is discovered from standard report naming, auto-registers it dynamically.
    """
    profiles = load_all_merchant_profiles()

    # 1. Inspect filenames first (fastest and most reliable for SwinkPay exports)
    if filenames:
        for fn in filenames:
            fn_upper = os.path.basename(str(fn)).upper()

            # Priority exact checks
            if "ADVANCE GENUINE SPARES" in fn_upper or "GENUINE SPARES" in fn_upper or "000000000002095" in fn_upper:
                if "ags" in profiles or "ags" in DEFAULT_MERCHANT_PROFILES:
                    return "ags"
            if "VALUE EXPRESS" in fn_upper or "XCD" in fn_upper:
                if "xcd" in profiles or "xcd" in DEFAULT_MERCHANT_PROFILES:
                    return "xcd"
            if "COFFEE DAY" in fn_upper or "CCD" in fn_upper:
                if "ccd" in profiles or "ccd" in DEFAULT_MERCHANT_PROFILES:
                    return "ccd"
            if "SBB" in fn_upper or "MEDICARE" in fn_upper:
                if "sbb" in profiles or "sbb" in DEFAULT_MERCHANT_PROFILES:
                    return "sbb"

            # Check all registered profile keywords and merchant IDs
            for key, prof in profiles.items():
                for kw in prof.merchant_keywords:
                    if kw and kw in fn_upper:
                        return prof.key
                for mid in prof.merchant_ids:
                    if mid and mid in fn_upper:
                        return prof.key

            # Pattern check: {MERCHANT_NAME}_{MERCHANT_ID}_TransactionsReport_...
            match = re.match(r"^([A-Z0-9\s]+)_([0-9]+)_TransactionsReport", fn_upper)
            if match:
                m_name = match.group(1).strip()
                m_id = match.group(2).strip()
                # Dynamically register custom profile if not recognized
                new_key = re.sub(r"[^a-zA-Z0-9]+", "_", m_name).strip("_").lower()
                short_code = "".join([w[0] for w in m_name.split() if w]) or new_key.upper()
                new_prof = save_merchant_profile({
                    "name": m_name.title(),
                    "key": new_key,
                    "merchant_ids": [m_id],
                    "merchant_keywords": [m_name, short_code],
                    "input_file_prefix": f"{short_code} Input file as on"
                })
                return new_prof.key

    # 2. Inspect SMMS records and CMS records (SMMS contains 'Merchant Name' and 'Merchant ID')
    source_recs = (smms_records or []) + (records or [])
    for r in source_recs[:100]:
        m_name = str(r.get("Merchant Name") or "").upper().strip()
        m_id = str(r.get("Merchant ID") or "").strip()

        if not m_name and not m_id:
            continue

        if "ADVANCE GENUINE SPARES" in m_name or "GENUINE SPARES" in m_name or m_id.endswith("2095"):
            if "ags" in profiles or "ags" in DEFAULT_MERCHANT_PROFILES:
                return "ags"
        if "VALUE EXPRESS" in m_name or "XCD" in m_name:
            if "xcd" in profiles or "xcd" in DEFAULT_MERCHANT_PROFILES:
                return "xcd"
        if "COFFEE DAY" in m_name or "CCD" in m_name:
            if "ccd" in profiles or "ccd" in DEFAULT_MERCHANT_PROFILES:
                return "ccd"
        if "SBB" in m_name or "MEDICARE" in m_name:
            if "sbb" in profiles or "sbb" in DEFAULT_MERCHANT_PROFILES:
                return "sbb"

        clean_mid = m_id.lstrip("0")
        for key, prof in profiles.items():
            for mid in prof.merchant_ids:
                if mid:
                    clean_reg = str(mid).strip().lstrip("0")
                    if clean_mid and clean_reg and (clean_mid == clean_reg or m_id == str(mid).strip() or m_id.endswith(str(mid).strip()) or str(mid).strip().endswith(m_id)):
                        return prof.key

        for key, prof in profiles.items():
            for kw in prof.merchant_keywords:
                if kw and kw in m_name:
                    return prof.key

        # If a non-empty merchant name is found but unmatched, register it dynamically
        if m_name and len(m_name) > 2:
            new_key = re.sub(r"[^a-zA-Z0-9]+", "_", m_name).strip("_").lower()
            short_code = "".join([w[0] for w in m_name.split() if w]) or new_key.upper()
            new_prof = save_merchant_profile({
                "name": m_name.title(),
                "key": new_key,
                "merchant_ids": [m_id] if m_id else [],
                "merchant_keywords": [m_name, short_code],
                "input_file_prefix": f"{short_code} Input file as on"
            })
            return new_prof.key

    if profiles:
        return next(iter(profiles.keys()))

    return "xcd"


def get_merchant_profile(
    key: Optional[str] = None,
    records: Optional[List[Dict[str, Any]]] = None,
    filenames: Optional[List[str]] = None,
    smms_records: Optional[List[Dict[str, Any]]] = None
) -> MerchantProfile:
    """
    Gets MerchantProfile by key or auto-detects from filenames, SMMS, or CMS records.
    If a key is provided that does not exist, automatically creates a dynamic profile for it.
    """
    profiles = load_all_merchant_profiles()

    if key and key.lower() in profiles:
        return profiles[key.lower()]

    # If key matches a known baseline profile, return its default
    if key and key.lower() in DEFAULT_MERCHANT_PROFILES:
        return DEFAULT_MERCHANT_PROFILES[key.lower()]

    # If key is a custom string (and not "auto" or empty), register it dynamically
    if key and key.lower() not in ("auto", "none", ""):
        profile = save_merchant_profile({"name": key, "key": key.lower()})
        return profile

    detected_key = detect_merchant(records=records, filenames=filenames, smms_records=smms_records)
    if detected_key and detected_key in profiles:
        return profiles[detected_key]

    if detected_key and detected_key in DEFAULT_MERCHANT_PROFILES:
        return DEFAULT_MERCHANT_PROFILES[detected_key]

    if profiles:
        return list(profiles.values())[0]

    # Clean fallback if no merchant is configured yet
    return DEFAULT_MERCHANT_PROFILES.get("xcd", MerchantProfile(
        key="custom",
        display_name="Custom Merchant",
        input_file_prefix="RECON_INPUTFILE",
        gateways=["CashFree", "EaseBuzz", "Airtel Bank"]
    ))
