"""
Merchant Profiles and Configuration for Ops_Auto Reconciliation.
Supports CCD Value Express (Coffee Day / XCD) and SBB Medicare.
"""
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional


@dataclass
class MerchantProfile:
    key: str
    display_name: str
    merchant_ids: List[str]
    merchant_keywords: List[str]
    gateways: List[str]
    has_split_settlement: bool
    softpos_label: str
    input_file_prefix: str
    upi_fee_rate: float
    cc_fee_rate: float
    default_gst_rate: float = 0.18


# Pre-configured merchant profiles
MERCHANT_PROFILES: Dict[str, MerchantProfile] = {
    "ccd": MerchantProfile(
        key="ccd",
        display_name="CCD Value Express (Coffee Day)",
        merchant_ids=["0000000000002045", "2045"],
        merchant_keywords=["CCD", "COFFEE DAY", "VALUE EXPRESS"],
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


def detect_merchant(records: List[Dict[str, Any]]) -> str:
    """
    Auto-detects merchant key ('sbb' or 'ccd') from CMS or SMMS records
    by checking 'Merchant Name' and 'Merchant ID'. Defaults to 'ccd' if indeterminate.
    """
    if not records:
        return "ccd"

    for r in records[:50]:
        m_name = str(r.get("Merchant Name") or "").upper()
        m_id = str(r.get("Merchant ID") or "").strip()

        # Check SBB signatures
        for kw in MERCHANT_PROFILES["sbb"].merchant_keywords:
            if kw in m_name:
                return "sbb"
        for mid in MERCHANT_PROFILES["sbb"].merchant_ids:
            if m_id.endswith(mid):
                return "sbb"

        # Check CCD signatures
        for kw in MERCHANT_PROFILES["ccd"].merchant_keywords:
            if kw in m_name:
                return "ccd"
        for mid in MERCHANT_PROFILES["ccd"].merchant_ids:
            if m_id.endswith(mid):
                return "ccd"

    return "ccd"


def get_merchant_profile(key: Optional[str] = None, records: Optional[List[Dict[str, Any]]] = None) -> MerchantProfile:
    """Gets MerchantProfile by key, or auto-detects from records if key is omitted or 'auto'."""
    if key and key.lower() in MERCHANT_PROFILES:
        return MERCHANT_PROFILES[key.lower()]

    if records:
        detected_key = detect_merchant(records)
        return MERCHANT_PROFILES[detected_key]

    return MERCHANT_PROFILES["ccd"]
