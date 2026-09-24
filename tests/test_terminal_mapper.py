"""
Unit tests for Terminal Reference Mapping and Cashfree Middle Number Resolution.
"""
import os
import unittest
import openpyxl
import tempfile
import json

from core.terminal_mapper import (
    extract_cf_middle_number,
    parse_terminal_report,
    resolve_mms_terminal_id,
    clean_key,
)
from core.pg_payload_builder import build_missing_pg_payloads


class TestTerminalMapper(unittest.TestCase):
    def test_extract_cf_middle_number(self):
        # Standard format: <merchant_id>-<middle_number>-<unique_id>
        order_id = "330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline"
        mid = extract_cf_middle_number(order_id)
        self.assertEqual(mid, "4860")

        # Another format with whitespace or dashes
        order_id2 = "330595-4873-AXI12345"
        self.assertEqual(extract_cf_middle_number(order_id2), "4873")

        # Two-part format: <pref>-<middle>
        order_id3 = "12345-6789"
        self.assertEqual(extract_cf_middle_number(order_id3), "6789")

        # Blank or invalid
        self.assertEqual(extract_cf_middle_number(""), "")
        self.assertEqual(extract_cf_middle_number("SINGLEKEY"), "")

    def test_parse_terminal_report_and_resolve(self):
        # Create a synthetic TID_FILE Excel workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TID_FILE"

        # Headers
        ws.append([
            "Sr No", "TERMINAL ID", "Merchant Name", "Merchant ID",
            "MMS TERMINAL ID", "VPA", "Partner Ref ID"
        ])
        # Data rows
        ws.append([1, "T1001", "Merchant Alpha", "M001", "X47VVH", "alpha@swinkpay", "4860"])
        ws.append([2, "T1002", "Merchant Beta", "M002", "X1JYYO", "beta-4873@axis", "4873"])
        ws.append([3, "T1003", "Merchant Gamma", "M003", "X99ZZZ", "gamma@okhdfc", ""])

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            temp_name = tf.name

        try:
            wb.save(temp_name)
            wb.close()

            mappings = parse_terminal_report(temp_name)
            self.assertIn("4860", mappings["mappings_by_ref_id"])
            self.assertEqual(mappings["mappings_by_ref_id"]["4860"]["mms_terminal_id"], "X47VVH")

            # Resolve by exact middle number / ref ID
            resolved_1 = resolve_mms_terminal_id("4860", mappings)
            self.assertEqual(resolved_1, "X47VVH")

            # Resolve from full Cashfree Order ID
            cf_order_id = "330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline"
            resolved_full = resolve_mms_terminal_id(cf_order_id, mappings)
            self.assertEqual(resolved_full, "X47VVH")

            # Resolve by terminal id
            resolved_tid = resolve_mms_terminal_id("T1003", mappings)
            self.assertEqual(resolved_tid, "X99ZZZ")
        finally:
            if os.path.exists(temp_name):
                os.remove(temp_name)


if __name__ == "__main__":
    unittest.main()
