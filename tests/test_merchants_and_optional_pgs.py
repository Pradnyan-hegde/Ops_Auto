"""
Unit tests for Merchant Profiles (XCD vs CCD, SBB), Delete Merchant API,
and Optional Payment Gateway Reconciliation.
"""
import os
import shutil
import tempfile
import unittest
from datetime import datetime
import asyncio
from fastapi import HTTPException

from core.merchant import (
    get_merchant_profile,
    save_merchant_profile,
    delete_merchant_profile,
    list_all_merchants,
    DEFAULT_MERCHANT_PROFILES,
    MERCHANTS_FILE
)
from core.matcher import ReconciliationEngine
from core.aggregator import Aggregator
from core.reader import RawReport
from core.detector import ReportType
from core.xcd_builder import generate_partner_xcd_files
from core.xcd_validator import validate_all_clear
from dashboard.server import app, delete_merchant


def _call(fn, *args, **kwargs):
    res = fn(*args, **kwargs)
    if asyncio.iscoroutine(res):
        return asyncio.run(res)
    return res


class TestMerchantProfilesAndDeletion(unittest.TestCase):
    def test_xcd_and_ccd_separation(self):
        """Verify XCD and CCD are distinct single-settlement merchant profiles."""
        xcd = get_merchant_profile("xcd")
        ccd = get_merchant_profile("ccd")

        self.assertEqual(xcd.key, "xcd")
        self.assertEqual(xcd.display_name, "Cafe Value Express (XCD)")
        self.assertFalse(xcd.has_split_settlement)
        self.assertEqual(xcd.input_file_prefix, "XCD Input file as on")

        self.assertEqual(ccd.key, "ccd")
        self.assertEqual(ccd.display_name, "Cafe Coffee Day (CCD)")
        self.assertFalse(ccd.has_split_settlement)
        self.assertEqual(ccd.input_file_prefix, "CCD Input file as on")

    def test_create_and_delete_merchant_profile(self):
        """Verify dynamic registration and deletion via function and REST API."""
        test_key = "test_brand_xyz"
        # 1. Create new merchant
        created = save_merchant_profile({
            "key": test_key,
            "name": "Test Brand XYZ",
            "has_split_settlement": False,
            "input_file_prefix": "TEST_BRAND_INPUT"
        })
        self.assertEqual(created.key, test_key)

        # Verify present in list
        all_m = list_all_merchants()
        keys = [m["key"] for m in all_m]
        self.assertIn(test_key, keys)

        # 2. Delete via REST API endpoint handler
        response = _call(delete_merchant, test_key)
        self.assertEqual(response.get("status"), "success")

        # Verify removed from list
        all_m_after = list_all_merchants()
        keys_after = [m["key"] for m in all_m_after]
        self.assertNotIn(test_key, keys_after)


class TestOptionalPGReconciliation(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_reconciliation_with_only_cms_and_cashfree(self):
        """
        Verify that when Easebuzz and Airtel reports are NOT uploaded,
        the engine reconciles Cashfree, does NOT raise false missing-partner exceptions
        for Easebuzz or Airtel, passes Diff-to-Zero, and passes All-Clear.
        """
        # Synthesize CMS report (has 2 Cashfree transactions)
        cms_records = [
            {
                "RRN/UTR": "1001",
                "SwinkPay Txn ID": "SP001",
                "Transaction Amount": "100.00",
                "Transaction Status": "SUCCESS",
                "Network": "UPI",
                "Transaction Date & Time": "2026-09-20 10:00:00",
                "Merchant MMS Terminal ID": "TID01"
            },
            {
                "RRN/UTR": "1002",
                "SwinkPay Txn ID": "SP002",
                "Transaction Amount": "200.00",
                "Transaction Status": "SUCCESS",
                "Network": "UPI",
                "Transaction Date & Time": "2026-09-20 11:00:00",
                "Merchant MMS Terminal ID": "TID02"
            }
        ]
        cms_headers = list(cms_records[0].keys())
        cms_rep = RawReport("cms_test.csv", ReportType.CMS, cms_headers, 0, cms_records, [cms_headers])

        # Synthesize Cashfree report
        cf_records = [
            {
                "Bank Reference No.": "1001",
                "Reference Id": "SP001",
                "Amount": "100.00",
                "Transaction Status": "SUCCESS",
                "Payment Mode": "UPI",
                "Payment Time": "2026-09-20 10:00:00"
            },
            {
                "Bank Reference No.": "1002",
                "Reference Id": "SP002",
                "Amount": "200.00",
                "Transaction Status": "SUCCESS",
                "Payment Mode": "UPI",
                "Payment Time": "2026-09-20 11:00:00"
            }
        ]
        cf_headers = list(cf_records[0].keys())
        cf_rep = RawReport("cf_test.csv", ReportType.CASHFREE, cf_headers, 0, cf_records, [cf_headers])

        # Synthesize SMMS report matching CMS
        smms_records = [
            {
                "RRN/UTR": "1001",
                "SwinkPay Txn ID": "SP001",
                "Transaction Amount": "100.00",
                "Transaction Status": "SUCCESS",
                "Network": "UPI",
                "Transaction Date & Time": "2026-09-20 10:00:00",
                "PG/Bank": "Cashfree",
                "Net Amount": "98.50"
            },
            {
                "RRN/UTR": "1002",
                "SwinkPay Txn ID": "SP002",
                "Transaction Amount": "200.00",
                "Transaction Status": "SUCCESS",
                "Network": "UPI",
                "Transaction Date & Time": "2026-09-20 11:00:00",
                "PG/Bank": "Cashfree",
                "Net Amount": "197.00"
            }
        ]
        smms_headers = list(smms_records[0].keys())
        smms_rep = RawReport("smms_test.csv", ReportType.SMMS, smms_headers, 0, smms_records, [smms_headers])

        # Initialize engine with ONLY CMS, SMMS, and Cashfree (Easebuzz and Airtel are None)
        engine = ReconciliationEngine(tolerance=0.01)
        engine.set_reports(
            cms=cms_rep,
            smms=smms_rep,
            cf=cf_rep,
            eb=None,
            airtel=None
        )
        engine.run()

        # Both records should match in Cashfree
        self.assertEqual(len(engine.cms_cf_matched), 2)
        self.assertEqual(len(engine.cms_eb_matched), 0)
        self.assertEqual(len(engine.cms_air_matched), 0)

        # No false "Missing in Partner Report" exceptions
        missing_in_partner = [e for e in engine.exceptions if e.get("Exception Category") == "Missing in Partner Report"]
        self.assertEqual(len(missing_in_partner), 0)

        # Aggregator control checks
        aggregator = Aggregator(engine, merchant_key="xcd")
        aggregator.compute()

        # Cashfree diff to zero passes
        cf_check = next((c for c in aggregator.control_checks if "Cashfree" in c.get("check_name", "")), None)
        self.assertIsNotNone(cf_check)
        self.assertEqual(cf_check.get("status"), "PASS")

        # Easebuzz and Airtel checks are omitted when not uploaded (no false failures)
        eb_check = next((c for c in aggregator.control_checks if "Easebuzz" in c.get("check_name", "")), None)
        air_check = next((c for c in aggregator.control_checks if "Airtel" in c.get("check_name", "")), None)
        self.assertIsNone(eb_check)
        self.assertIsNone(air_check)

        # Validate All-Clear gate passes with 0 blocking reasons
        gate = validate_all_clear(engine, aggregator, settlement_date="2026-09-21")
        self.assertTrue(gate.is_all_clear)
        self.assertEqual(len(gate.blocking_reasons), 0)

        # Test XCD vs CCD file naming
        files_xcd = generate_partner_xcd_files(engine, self.temp_dir, "2026-09-20", "2026-09-21", merchant_key="xcd")
        self.assertTrue(files_xcd["cashfree"]["filename"].startswith("XCD Input file as on"))

        files_ccd = generate_partner_xcd_files(engine, self.temp_dir, "2026-09-20", "2026-09-21", merchant_key="ccd")
        self.assertTrue(files_ccd["cashfree"]["filename"].startswith("CCD Input file as on"))


if __name__ == "__main__":
    unittest.main()
