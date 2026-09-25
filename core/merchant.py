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
            "input_file_prefix", "upi_fee_rate", "cc_fee_rate", "default_gst_rate"
        }
        filtered = {k: v for k, v in data.items() if k in valid_keys}
        return cls(**filtered)


# Built-in baseline merchant profiles
DEFAULT_MERCHANT_PROFILES: Dict[str, MerchantProfile] = {
    "ccd": MerchantProfile(
        key="ccd",
        display_name="CCD Value Express (Coffee Day)",
        merchant_ids=["0000000000002045", "2045"],
        merchant_keywords=["CCD", "COFFEE DAY", "VALUE EXPRESS", "XCD"],
        gateways=["CashFree", "EaseBuzz", "Airtel Bank"],
        has_split_settlement=False,
        softpos_label="CashFree",
        input_file_prefix="XCD Input file as on",
        upi_fee_rate=0.01,    # 1.00%
        cc_fee_rate=0.0275,   # 2.75%
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
    if os.path.exists(MERCHANTS_FILE):
        try:
            with open(MERCHANTS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if isinstance(saved, dict):
                    for k, v in saved.items():
                        if isinstance(v, dict):
                            profiles[k.lower()] = MerchantProfile.from_dict(v)
        except Exception as e:
            print(f"[MerchantRegistry] Warning loading {MERCHANTS_FILE}: {e}")
    return profiles


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
    prefix = str(data.get("input_file_prefix") or "").strip()
    if not prefix:
        short_code = re.sub(r"[^a-zA-Z0-9]+", "_", display_name).strip("_").upper()
        prefix = f"{short_code}_INPUTFILE"

    # Keywords
    keywords = data.get("merchant_keywords") or []
    if not keywords:
        keywords = [display_name.upper(), raw_key.upper()]
    else:
        keywords = [str(k).upper() for k in keywords]

    m_ids = [str(m).strip() for m in (data.get("merchant_ids") or []) if str(m).strip()]
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
        upi_fee_rate=upi_rate,
        cc_fee_rate=cc_rate,
        default_gst_rate=gst_rate
    )

    profiles[profile.key] = profile

    # Save custom profiles (excluding baseline defaults unless modified)
    custom_dict = {}
    for k, p in profiles.items():
        custom_dict[k] = p.to_dict()

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
    # CCD first, SBB second, then others alphabetically
    order = ["ccd", "sbb"]
    sorted_keys = order + sorted([k for k in profiles.keys() if k not in order])

    for k in sorted_keys:
        if k in profiles:
            p = profiles[k]
            results.append({
                "key": p.key,
                "display_name": p.display_name,
                "input_file_prefix": p.input_file_prefix,
                "has_split_settlement": p.has_split_settlement,
                "merchant_ids": p.merchant_ids,
                "gateways": p.gateways,
                "softpos_label": p.softpos_label,
                "upi_fee_rate": p.upi_fee_rate,
                "cc_fee_rate": p.cc_fee_rate
            })
    return results


def detect_merchant(records: List[Dict[str, Any]]) -> str:
    """
    Auto-detects merchant key from CMS or SMMS records
    by checking 'Merchant Name' and 'Merchant ID' against all registered profiles.
    """
    if not records:
        return "ccd"

    profiles = load_all_merchant_profiles()

    for r in records[:60]:
        m_name = str(r.get("Merchant Name") or "").upper().strip()
        m_id = str(r.get("Merchant ID") or "").strip()

        # Check all profiles
        for key, prof in profiles.items():
            # Check keywords in merchant name
            for kw in prof.merchant_keywords:
                if kw and kw in m_name:
                    return prof.key

            # Check merchant IDs
            for mid in prof.merchant_ids:
                if mid and m_id.endswith(mid):
                    return prof.key

    return "ccd"


def get_merchant_profile(key: Optional[str] = None, records: Optional[List[Dict[str, Any]]] = None) -> MerchantProfile:
    """
    Gets MerchantProfile by key or auto-detects from records.
    If a key is provided that does not exist, automatically creates a dynamic profile for it.
    """
    profiles = load_all_merchant_profiles()

    if key and key.lower() in profiles:
        return profiles[key.lower()]

    # If key is a custom string (and not "auto" or empty), register it dynamically
    if key and key.lower() not in ("auto", "none", ""):
        profile = save_merchant_profile({"name": key, "key": key.lower()})
        return profile

    if records:
        detected_key = detect_merchant(records)
        if detected_key in profiles:
            return profiles[detected_key]

    return profiles.get("ccd", DEFAULT_MERCHANT_PROFILES["ccd"])
