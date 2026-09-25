"""
Unit tests for 5-slot report detection, active outlet counts, and Missing PG API push.
"""
import os
import sys
import json
import unittest
from unittest.mock import patch, MagicMock
from io import BytesIO
from fastapi import UploadFile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from dashboard.server import (
    app,
    detect_uploaded_files,
    push_missing_pg_transaction,
    _push_payload_to_swinkpay,
    SESSIONS_DIR
)

class TestEmailAndOutlets(unittest.TestCase):

    def test_outlet_count_calculation_from_terminal_ids(self):
        sample_records = [
            {"Merchant MMS Terminal ID": "X7VJND"},
            {"Merchant MMS Terminal ID": "X7VJND"},
            {"Merchant MMS Terminal ID": "X1JYYO"},
            {"Merchant MMS Terminal ID": "XH2VUY"},
            {"Merchant MMS Terminal ID": ""},
            {"Merchant MMS Terminal ID": "none"},
            {"Merchant MMS Terminal ID": "nan"},
            {"Merchant MMS Terminal ID": "--"},
            {"Merchant MMS Terminal ID": "X7VJND"},
        ]
        distinct_terminals = set()
        for r in sample_records:
            tid = str(r.get("Merchant MMS Terminal ID") or "").strip()
            if tid and tid.lower() not in ("nan", "none", "", "--"):
                distinct_terminals.add(tid)

        self.assertEqual(len(distinct_terminals), 3)
        self.assertEqual(distinct_terminals, {"X7VJND", "X1JYYO", "XH2VUY"})

    def test_detect_uploaded_files_endpoint(self):
        cms_csv = ("RRN/UTR,SwinkPay Txn ID,Transaction Amount,Transaction Status,Network,Transaction Date & Time,Merchant MMS Terminal ID,SMMS Sync Status\n" "129761680150,SP001,20.00,SUCCESS,UPI,2026-09-17 15:59:29,X7VJND,true\n").encode("utf-8")
        smms_csv = ("RRN/UTR,SwinkPay Txn ID,Transaction Amount,Transaction Status,Network,Transaction Date & Time,PG/Bank,Net Amount\n" "129761680150,SP001,20.00,SUCCESS,UPI,2026-09-17 15:59:29,Cashfree,19.65\n").encode("utf-8")

        upload_cms = UploadFile(filename="daily_cms.csvtarget.csv", file=BytesIO(cms_csv))
        upload_smms = UploadFile(filename="daily_smms.csvtarget.csv", file=BytesIO(smms_csv))

        # CMS + SMMS uploaded: CMS is present and all PGs are optional -> is_ready is True
        resp = detect_uploaded_files([upload_cms, upload_smms])
        data = json.loads(resp.body.decode("utf-8"))

        self.assertIn("CMS", data["detected"])
        self.assertIn("SMMS", data["detected"])
        self.assertTrue(data["is_ready"])

        # When CMS is missing -> is_ready is False and CMS in missing
        upload_smms_only = UploadFile(filename="daily_smms.csvtarget.csv", file=BytesIO(smms_csv))
        resp_missing = detect_uploaded_files([upload_smms_only])
        data_missing = json.loads(resp_missing.body.decode("utf-8"))
        self.assertIn("CMS", data_missing["missing"])
        self.assertFalse(data_missing["is_ready"])

    @patch("urllib.request.urlopen")
    def test_push_payload_to_swinkpay_success(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.read.return_value = b'{"status": "success", "message": "Transaction updated"}'
        mock_urlopen.return_value.__enter__.return_value = mock_resp

        payload = {
            "amount": "20.00",
            "terminalID": "X7VJND",
            "utr": "129761680150",
            "dateAndTime": "2026-09-17 15:59:29"
        }

        result = _push_payload_to_swinkpay(payload)
        self.assertTrue(result["success"])
        self.assertEqual(result["status_code"], 200)

    @patch("urllib.request.urlopen")
    def test_push_missing_pg_session_endpoint(self, mock_urlopen):
        import asyncio
        session_id = "test_push_session_999"
        session_dir = os.path.join(SESSIONS_DIR, session_id)
        os.makedirs(session_dir, exist_ok=True)

        try:
            sample_payload = {
                "amount": "20.00",
                "terminalID": "X7VJND",
                "utr": "129761680150",
                "dateAndTime": "2026-09-17 15:59:29"
            }
            meta = {
                "all_clear": True,
                "missing_pg_payloads": [
                    {
                        "gateway": "Cashfree",
                        "order_id": "330595-4873",
                        "terminal_id": "X7VJND",
                        "payload": sample_payload
                    }
                ]
            }
            with open(os.path.join(session_dir, "session_meta.json"), "w", encoding="utf-8") as f:
                json.dump(meta, f)

            mock_resp = MagicMock()
            mock_resp.status = 200
            mock_resp.read.return_value = b'{"success": true}'
            mock_urlopen.return_value.__enter__.return_value = mock_resp

            class DummyRequest:
                async def json(self):
                    return {"index": 0}

            resp = asyncio.run(push_missing_pg_transaction(session_id, DummyRequest()))
            data = json.loads(resp.body.decode("utf-8"))

            self.assertTrue(data["success"])
            self.assertEqual(data["status_code"], 200)
            self.assertEqual(data["order_id"], "330595-4873")
        finally:
            import shutil
            shutil.rmtree(session_dir, ignore_errors=True)


if __name__ == '__main__':
    unittest.main()
