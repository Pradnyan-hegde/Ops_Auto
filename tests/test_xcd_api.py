"""
Direct Endpoint verification tests for XCD settlement download gate.
Uses asyncio to test route handlers directly without requiring httpx.
"""
import os
import sys
import unittest
import json
import asyncio
from fastapi import HTTPException

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from dashboard.server import serve_dashboard, download_xcd, SESSIONS_DIR
from core.xcd_builder import build_xcd_workbook


def _call(fn, *args, **kwargs):
    res = fn(*args, **kwargs)
    if asyncio.iscoroutine(res):
        return asyncio.run(res)
    return res


class TestXCDAPI(unittest.TestCase):
    def test_dashboard_homepage_serves_html(self):
        resp = _call(serve_dashboard)
        self.assertEqual(resp.status_code, 200)
        content = resp.body.decode("utf-8")
        self.assertIn("SwinkPay Reconciliation Portal", content)
        self.assertIn("XCD Input file (Cashfree)", content)
        self.assertIn("Settlement Date (Authoritative Control)", content)

    def test_download_xcd_blocked_when_session_invalid(self):
        with self.assertRaises(HTTPException) as ctx:
            _call(download_xcd, "non_existent_session_xyz", "cashfree")
        self.assertEqual(ctx.exception.status_code, 404)

    def test_download_xcd_allowed_even_when_metadata_not_all_clear(self):
        session_id = "test_unblocked_session_123"
        session_dir = os.path.join(SESSIONS_DIR, session_id)
        os.makedirs(session_dir, exist_ok=True)
        try:
            test_file = os.path.join(session_dir, "XCD Input file as on 2026-09-10 (Cf).xlsx")
            records = [{"SwinkPay Txn ID": "SP_CF_1", "Transaction Amount": 500.0}]
            build_xcd_workbook(records, "11-09-2026", test_file)

            meta = {
                "all_clear": False,
                "summary_message": "Audit discrepancy noted: count variance of -41.",
                "blocking_reasons": ["Count variance in Cashfree"]
            }
            with open(os.path.join(session_dir, "session_meta.json"), "w") as f:
                json.dump(meta, f)

            # Downloads are now unlocked even when discrepancies exist
            resp = _call(download_xcd, session_id, "cashfree")
            self.assertEqual(resp.media_type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        finally:
            import shutil
            shutil.rmtree(session_dir, ignore_errors=True)

    def test_download_xcd_success_when_all_clear(self):
        session_id = "test_pass_session_123"
        session_dir = os.path.join(SESSIONS_DIR, session_id)
        os.makedirs(session_dir, exist_ok=True)
        try:
            test_file = os.path.join(session_dir, "XCD Input file as on 2026-09-10 (Cf).xlsx")
            records = [{"SwinkPay Txn ID": "SP_CF_1", "Transaction Amount": 500.0}]
            build_xcd_workbook(records, "11-09-2026", test_file)

            meta = {
                "all_clear": True,
                "summary_message": "Reconciliation complete. All three XCD settlement files are ready to download.",
                "settlement_date": "11-09-2026",
                "files": {
                    "cashfree": {
                        "filename": "XCD Input file as on 2026-09-10 (Cf).xlsx"
                    }
                }
            }
            with open(os.path.join(session_dir, "session_meta.json"), "w") as f:
                json.dump(meta, f)

            resp = _call(download_xcd, session_id, "cashfree")
            self.assertEqual(resp.filename, "XCD Input file as on 2026-09-10 (Cf).xlsx")
            self.assertEqual(resp.media_type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        finally:
            import shutil
            shutil.rmtree(session_dir, ignore_errors=True)

    def test_generate_xcd_for_session_unlocks_downloads(self):
        from dashboard.server import generate_xcd_for_session
        from starlette.requests import Request
        import openpyxl

        session_id = "test_unlock_session_456"
        session_dir = os.path.join(SESSIONS_DIR, session_id)
        os.makedirs(session_dir, exist_ok=True)
        try:
            # Create dummy reconciled workbook
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            for sheet_name in ["CMS_CF_Matched", "CMS_EB_Matched", "CMS_Air_Matched"]:
                ws = wb.create_sheet(title=sheet_name)
                ws.append(["SwinkPay Txn ID", "Transaction Amount"])
                ws.append(["SP12345", 250.0])
            recon_path = os.path.join(session_dir, "Reconciliation_2026-09-17.xlsx")
            wb.save(recon_path)

            meta = {
                "all_clear": False,
                "summary_message": "XCD downloads blocked: Settlement date is missing.",
                "blocking_reasons": [
                    "Settlement date is missing. Please select a settlement date or upload an Airtel Settlement report."
                ],
                "files": {}
            }
            with open(os.path.join(session_dir, "session_meta.json"), "w") as f:
                json.dump(meta, f)

            req = Request({"type": "http", "query_string": b"", "headers": []})
            result = _call(generate_xcd_for_session, session_id, req, settlement_date="18-09-2026")

            self.assertTrue(result["success"])
            self.assertTrue(result["xcd_status"]["all_clear"])
            self.assertEqual(len(result["xcd_status"]["files"]), 3)
            self.assertIn("cashfree", result["xcd_status"]["files"])
            self.assertIn("easebuzz", result["xcd_status"]["files"])
            self.assertIn("airtel", result["xcd_status"]["files"])
        finally:
            import shutil
            shutil.rmtree(session_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
