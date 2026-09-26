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

    def test_ags_profile_and_filename_detection(self):
        """Verify AGS profile detection from standard export filenames and output file prefixing."""
        from core.merchant import detect_merchant
        ags = get_merchant_profile("ags")
        self.assertEqual(ags.key, "ags")
        self.assertEqual(ags.display_name, "Advance Genuine Spares (AGS)")
        self.assertEqual(ags.input_file_prefix, "AGS Input file as on")

        # Filename auto-detection
        ags_fn = "ADVANCE GENUINE SPARES_000000000002095_TransactionsReport_2026-09-26T01_16_58_472Z.xlsx"
        detected = detect_merchant(filenames=[ags_fn])
        self.assertEqual(detected, "ags")

        # Engine file generation
        cms_records = [{
            "RRN/UTR": "1001",
            "SwinkPay Txn ID": "SP001",
            "Transaction Amount": "100.00",
            "Transaction Status": "SUCCESS",
            "Network": "UPI",
            "Transaction Date & Time": "2026-09-25 10:00:00",
            "Merchant MMS Terminal ID": "TID01",
            "Payment Gateway": "Cashfree"
        }]
        cms_rep = RawReport("cms.csv", ReportType.CMS, list(cms_records[0].keys()), 0, cms_records, [])
        smms_records = [{
            "RRN/UTR": "1001",
            "SwinkPay Txn ID": "SP001",
            "Transaction Amount": "100.00",
            "Transaction Status": "SUCCESS",
            "Network": "UPI",
            "Transaction Date & Time": "2026-09-25 10:00:00",
            "Merchant MMS Terminal ID": "TID01",
            "PG/Bank": "Cashfree",
            "Net Amount": "98.50"
        }]
        smms_rep = RawReport("smms.csv", ReportType.SMMS, list(smms_records[0].keys()), 0, smms_records, [])
        cf_records = [{
            "Bank Reference No.": "1001",
            "Reference Id": "SP001",
            "Amount": "100.00",
            "Transaction Status": "SUCCESS",
            "Payment Mode": "UPI",
            "Payment Time": "2026-09-25 10:00:00"
        }]
        cf_rep = RawReport("cf.csv", ReportType.CASHFREE, list(cf_records[0].keys()), 0, cf_records, [])

        engine = ReconciliationEngine()
        engine.set_reports(cms=cms_rep, smms=smms_rep, cf=cf_rep)
        engine.run()

        files_ags = generate_partner_xcd_files(engine, self.temp_dir, "2026-09-25", "2026-09-26", merchant_key="ags")
        self.assertTrue(files_ags["cashfree"]["filename"].startswith("AGS Input file as on"))

    def test_mandatory_smms_check(self):
        """Verify that omitting SMMS raises HTTPException 400."""
        from dashboard.server import execute_recon_for_files
        cms_path = os.path.join(self.temp_dir, "cms_only.csv")
        with open(cms_path, "w", encoding="utf-8") as f:
            f.write("RRN/UTR,SwinkPay Txn ID,Transaction Amount,Transaction Status,Network,Transaction Date & Time,Merchant MMS Terminal ID,Payment Gateway\n"
                    "1001,SP001,100.00,SUCCESS,UPI,2026-09-20 10:00:00,TID01,Cashfree\n")

        sess_dir = os.path.join(self.temp_dir, "sess_missing_smms")
        os.makedirs(sess_dir, exist_ok=True)

        with self.assertRaises(HTTPException) as ctx:
            execute_recon_for_files([cms_path], sess_dir)
        self.assertEqual(ctx.exception.status_code, 400)
        self.assertIn("SMMS Report must be uploaded", ctx.exception.detail)

    def test_mandatory_pg_report_when_transactions_present_in_cms(self):
        """Verify that if CMS has Cashfree transactions, omitting Cashfree report raises HTTPException 400."""
        from dashboard.server import execute_recon_for_files
        cms_path = os.path.join(self.temp_dir, "cms_cf_present.csv")
        with open(cms_path, "w", encoding="utf-8") as f:
            f.write("RRN/UTR,SwinkPay Txn ID,Transaction Amount,Transaction Status,Network,Transaction Date & Time,Merchant MMS Terminal ID,Payment Gateway\n"
                    "1001,SP001,100.00,SUCCESS,UPI,2026-09-20 10:00:00,TID01,Cashfree\n")

        smms_path = os.path.join(self.temp_dir, "smms_cf_present.csv")
        with open(smms_path, "w", encoding="utf-8") as f:
            f.write("RRN/UTR,SwinkPay Txn ID,Transaction Amount,Transaction Status,Network,Transaction Date & Time,Merchant MMS Terminal ID,PG/Bank,Net Amount\n"
                    "1001,SP001,100.00,SUCCESS,UPI,2026-09-20 10:00:00,TID01,Cashfree,98.50\n")

        sess_dir = os.path.join(self.temp_dir, "sess_missing_pg")
        os.makedirs(sess_dir, exist_ok=True)

        # Uploaded only CMS and SMMS, but transactions are routed to Cashfree
        with self.assertRaises(HTTPException) as ctx:
            execute_recon_for_files([cms_path, smms_path], sess_dir)
        self.assertEqual(ctx.exception.status_code, 400)
        self.assertIn("Missing Mandatory Partner Report(s)", ctx.exception.detail)
        self.assertIn("Cashfree", ctx.exception.detail)

    def test_airtel_invoice_summary_single_settlement_for_xcd(self):
        """Verify that XCD generates Single Daily Settlement text, not 2-batch timing."""
        from dashboard.server import execute_recon_for_files

        # Create sample CMS, SMMS, and Cashfree files
        cms_path = os.path.join(self.temp_dir, "cms.csv")
        with open(cms_path, "w", encoding="utf-8") as f:
            f.write("RRN/UTR,SwinkPay Txn ID,Transaction Amount,Transaction Status,Network,Transaction Date & Time,Merchant MMS Terminal ID,Payment Gateway\n"
                    "1001,SP001,100.00,SUCCESS,UPI,2026-09-20 10:00:00,TID01,Cashfree\n")

        smms_path = os.path.join(self.temp_dir, "smms.csv")
        with open(smms_path, "w", encoding="utf-8") as f:
            f.write("RRN/UTR,SwinkPay Txn ID,Transaction Amount,Transaction Status,Network,Transaction Date & Time,Merchant MMS Terminal ID,PG/Bank,Net Amount\n"
                    "1001,SP001,100.00,SUCCESS,UPI,2026-09-20 10:00:00,TID01,Cashfree,98.50\n")

        cf_path = os.path.join(self.temp_dir, "cf.csv")
        with open(cf_path, "w", encoding="utf-8") as f:
            f.write("Bank Reference No.,Reference Id,Amount,Settlement Amount,Transaction Status,Payment Mode,Transaction Time\n"
                    "1001,SP001,100.00,98.50,SUCCESS,UPI,2026-09-20 10:00:00\n")

        sess_dir = os.path.join(self.temp_dir, "session_test")
        os.makedirs(sess_dir, exist_ok=True)

        files = [cms_path, smms_path, cf_path]

        # 1. Run for XCD
        res_xcd = execute_recon_for_files(files, sess_dir, merchant_key="xcd")
        sum_xcd = res_xcd["airtel_invoice_summary"]
        self.assertFalse(sum_xcd["has_split_settlement"])
        self.assertEqual(sum_xcd["schedule_badge"], "Daily Schedule: Single Daily Settlement")
        self.assertEqual(sum_xcd["batch_heading"], "Bank Settlement & UTR Details")
        self.assertIn("single daily settlement", sum_xcd["timing_explanation"].lower())
        self.assertNotIn("two intraday batches", sum_xcd["timing_explanation"].lower())

        # 2. Run for SBB (split settlement)
        res_sbb = execute_recon_for_files(files, sess_dir, merchant_key="sbb")
        sum_sbb = res_sbb["airtel_invoice_summary"]
        self.assertTrue(sum_sbb["has_split_settlement"])
        self.assertEqual(sum_sbb["schedule_badge"], "Daily Schedule: Batch 1 & Batch 2 (2 Settlements)")
        self.assertEqual(sum_sbb["batch_heading"], "Two-Batch Settlement Timing & UTR Breakdown")
        self.assertIn("two intraday batches", sum_sbb["timing_explanation"].lower())


if __name__ == "__main__":
    unittest.main()


