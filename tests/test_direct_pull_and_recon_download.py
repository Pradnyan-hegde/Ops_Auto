"""
Unit and integration tests for:
1. Reconciled workbook download resolution bugfix (ensuring uploaded SMMS files are never downloaded as reconciliation workbooks).
2. Direct PG report upload and 1-shot pull endpoints (/api/direct-pull/upload and /api/direct-pull/push).
3. Payload construction across Cashfree, Easebuzz, Airtel, and generic reports.
"""
import os
import shutil
import tempfile
import unittest
import json
from unittest.mock import patch, MagicMock

from core.detector import ReportType
from core.reader import RawReport
from core.pg_payload_builder import (
    parse_pg_report_records_for_pull,
    format_payload_amount,
    format_payload_datetime
)
from dashboard.server import (
    find_session_recon_workbook,
    upload_pg_report_for_pull,
    push_direct_payloads,
    SESSIONS_DIR
)


def make_raw_report(report_type: ReportType, records: list, file_name: str = "report.csv", headers: list = None) -> RawReport:
    if headers is None:
        headers = list(records[0].keys()) if records else []
    raw_matrix = [headers] + [[r.get(h, "") for h in headers] for r in records]
    return RawReport(
        file_path=file_name,
        report_type=report_type,
        headers=headers,
        header_row_idx=0,
        records=records,
        raw_matrix=raw_matrix
    )


class TestDirectPullAndReconDownload(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_recon_session_")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_find_session_recon_workbook_never_returns_smms_report(self):
        """
        Verify that even when an SMMS file with name
        'CCD value express_0000000000002045_TransactionsReport_2026-09-28.xlsx'
        exists in the session, find_session_recon_workbook returns the true reconciliation file.
        """
        # Create dummy SMMS input report in session_dir
        smms_filename = "CCD value express_0000000000002045_TransactionsReport_2026-09-28T06_02_33_852Z.xlsx"
        smms_path = os.path.join(self.test_dir, smms_filename)
        with open(smms_path, "w") as f:
            f.write("dummy smms content")

        # Create actual reconciliation workbook
        recon_filename = "CCD_Reconciliation_2026-09-28.xlsx"
        recon_path = os.path.join(self.test_dir, recon_filename)
        with open(recon_path, "w") as f:
            f.write("dummy recon content")

        # Case 1: metadata with output_filename
        meta = {
            "output_filename": recon_filename,
            "recon_name": "CCD Recon"
        }
        found = find_session_recon_workbook(self.test_dir, meta)
        self.assertEqual(found, recon_path)
        self.assertNotEqual(found, smms_path)

        # Case 2: metadata is None or empty - must match recon_filename by 'reconciliation' pattern
        found_no_meta = find_session_recon_workbook(self.test_dir, None)
        self.assertEqual(found_no_meta, recon_path)
        self.assertNotEqual(found_no_meta, smms_path)

    def test_parse_pg_report_records_for_cashfree(self):
        """Verify Cashfree report parsing and payload construction."""
        cf_records = [
            {
                "Order Id": "330595-4860-ORD01",
                "Bank Reference No.": "129761680101",
                "Amount": "100.50",
                "Transaction Time": "2026-09-28 10:15:30",
                "Transaction Status": "SUCCESS"
            },
            {
                "Order Id": "330595-4860-ORD02",
                "Bank Reference No.": "129761680102",
                "Amount": "250.00",
                "Transaction Time": "2026-09-28 10:20:00",
                "Transaction Status": "SUCCESS"
            }
        ]
        rep = make_raw_report(ReportType.CASHFREE, cf_records, "cf_report.csv")

        with patch("core.pg_payload_builder.resolve_terminal_details", return_value={"mms_terminal_id": "XKD8M3", "branch_name": "CCD value express"}):
            items = parse_pg_report_records_for_pull(rep, merchant_key="ccd")

        self.assertEqual(len(items), 2)
        item1 = items[0]
        self.assertEqual(item1["gateway"], "Cashfree")
        self.assertEqual(item1["terminal_id"], "XKD8M3")
        self.assertEqual(item1["utr"], "129761680101")
        self.assertEqual(item1["amount"], "100.50")
        self.assertEqual(item1["date_and_time"], "2026-09-28 10:15:30")
        self.assertEqual(item1["payload"]["terminalID"], "XKD8M3")
        self.assertEqual(item1["payload"]["amount"], "100.50")
        self.assertEqual(item1["payload"]["utr"], "129761680101")

    def test_parse_pg_report_records_for_easebuzz(self):
        """Verify Easebuzz report parsing and payload construction."""
        eb_records = [
            {
                "ID": "EB_1001",
                "UPI tid": "129761680201",
                "UTR": "129761680201",
                "Amount": "45.00",
                "Transaction Date": "2026-09-28 11:00:00",
                "Status": "Payment Received",
                "Virtual Account Label": "XKD8M3"
            }
        ]
        rep = make_raw_report(ReportType.EASEBUZZ, eb_records, "eb_report.csv")

        items = parse_pg_report_records_for_pull(rep, merchant_key="ccd")
        self.assertEqual(len(items), 1)
        item = items[0]
        self.assertEqual(item["gateway"], "Easebuzz")
        self.assertEqual(item["terminal_id"], "XKD8M3")
        self.assertEqual(item["amount"], "45.00")
        self.assertEqual(item["utr"], "129761680201")
        self.assertEqual(item["payload"]["terminalID"], "XKD8M3")

    def test_parse_pg_report_records_for_airtel(self):
        """Verify Airtel report parsing and payload construction."""
        air_records = [
            {
                "Transaction Id": "AIR_9901",
                "PARTNER_TXN_ID": "129761680301",
                "Original Input Amt": "180.00",
                "Date and Time": "2026-09-28 12:30:00",
                "Transaction To": "spf-2045-xh2vuy@mairtel"
            }
        ]
        rep = make_raw_report(ReportType.AIRTEL, air_records, "airtel_report.csv")

        items = parse_pg_report_records_for_pull(rep, merchant_key="ccd")
        self.assertEqual(len(items), 1)
        item = items[0]
        self.assertEqual(item["gateway"], "Airtel")
        self.assertEqual(item["terminal_id"], "XH2VUY")
        self.assertEqual(item["amount"], "180.00")
        self.assertEqual(item["utr"], "129761680301")
        self.assertEqual(item["payload"]["terminalID"], "XH2VUY")

    def test_push_direct_payloads_api(self):
        """Verify push_direct_payloads endpoint pushes payloads and returns structured telemetry."""
        from fastapi import Request
        import asyncio

        scope = {
            "type": "http",
            "method": "POST",
            "headers": [(b"content-type", b"application/json")]
        }
        body_data = {
            "payloads": [
                {
                    "payload": {
                        "amount": "100.00",
                        "terminalID": "XKD8M3",
                        "utr": "129761680101",
                        "dateAndTime": "2026-09-28 10:15:30"
                    }
                }
            ],
            "auth_token": "TEST_TOKEN",
            "channel": "14"
        }

        async def receive():
            return {
                "type": "http.request",
                "body": json.dumps(body_data).encode("utf-8")
            }

        req = Request(scope, receive)

        # Mock SwinkPay endpoint response
        mock_res = {
            "success": True,
            "status_code": 200,
            "response_time_ms": 115,
            "response": '{"data":{"referenceNo":"129761680101","txnId":"SP001"},"message":"Success"}',
            "data": {"referenceNo": "129761680101", "txnId": "SP001"},
            "message": "Success"
        }

        with patch("dashboard.server._push_payload_to_swinkpay", return_value=mock_res):
            resp = asyncio.run(push_direct_payloads(req))

        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.body.decode("utf-8"))
        self.assertTrue(data["success"])
        self.assertEqual(data["total"], 1)
        self.assertEqual(data["successful_count"], 1)
        self.assertEqual(data["results"][0]["status_code"], 200)


if __name__ == "__main__":
    unittest.main()
