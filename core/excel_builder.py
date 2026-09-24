"""
Excel Workbook Generator for Ops_Auto.
Creates fully formatted workbook with exact styling, number formats, borders, formulas, and Audit controls.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional
from collections import defaultdict

from .reader import RawReport
from .normalizer import clean_key, clean_amount, normalize_status


class ExcelReportBuilder:
    def __init__(self, engine, aggregator):
        self.engine = engine
        self.agg = aggregator
        self.wb = openpyxl.Workbook()
        # Remove default sheet
        if "Sheet" in self.wb.sheetnames:
            self.wb.remove(self.wb["Sheet"])

        # Standard Styles
        self.font_header = Font(name="Calibri", size=11, bold=True, color="000000")
        self.font_title = Font(name="Calibri", size=12, bold=True, color="1F497D")
        self.font_data = Font(name="Calibri", size=11, color="000000")
        self.font_total = Font(name="Calibri", size=11, bold=True, color="000000")
        self.font_pass = Font(name="Calibri", size=11, bold=True, color="006100")
        self.font_fail = Font(name="Calibri", size=11, bold=True, color="9C0006")

        self.fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        self.fill_subtotal = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        self.fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        self.fill_audit_hdr = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
        self.font_audit_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        self.fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        self.fill_fail = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        self.thin_side = Side(border_style="thin", color="D9D9D9")
        self.thin_black = Side(border_style="thin", color="000000")
        self.double_bottom = Side(border_style="double", color="000000")

        self.border_data = Border(left=self.thin_side, right=self.thin_side, top=self.thin_side, bottom=self.thin_side)
        self.border_header = Border(left=self.thin_black, right=self.thin_black, top=self.thin_black, bottom=self.thin_black)
        self.border_total = Border(left=self.thin_black, right=self.thin_black, top=self.thin_black, bottom=self.double_bottom)

        self.align_left = Alignment(horizontal="left", vertical="center")
        self.align_center = Alignment(horizontal="center", vertical="center")
        self.align_right = Alignment(horizontal="right", vertical="center")

    def _auto_fit_columns(self, ws, min_width: int = 12, max_width: int = 50):
        """Auto fits column widths with comfortable padding."""
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                val = cell.value
                if val is not None:
                    s_val = str(val)
                    if len(s_val) > max_len:
                        max_len = len(s_val)
            ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 3, max_width))

    def _write_table_sheet(self, sheet_name: str, records: List[Dict[str, Any]], base_headers: Optional[List[str]] = None):
        """Writes a list of records into an Excel sheet with standardized styling."""
        ws = self.wb.create_sheet(title=sheet_name)
        if not records:
            headers = base_headers or ["No Data Available"]
            ws.append(headers)
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(1, col_idx)
                cell.font = self.font_header
                cell.fill = self.fill_header
                cell.border = self.border_header
            self._auto_fit_columns(ws)
            return

        # Standard reconciliation headers to place at the front
        standard_recon_cols = [
            "Source System", "Match Key", "Matched Transaction ID", "CMS Amount",
            "Partner Amount", "Amount Difference", "CMS Status", "Partner Status",
            "Reconciliation Status", "Exception Reason", "Settlement UTR",
            "Settlement Date", "Settlement Amount", "Settlement Status"
        ]

        # Get all keys present in records
        all_record_keys = []
        for r in records:
            for k in r.keys():
                if k not in all_record_keys and not k.startswith("_"):
                    all_record_keys.append(k)

        # Order headers: standard recon cols first if present, then remaining original source columns
        ordered_headers = [c for c in standard_recon_cols if c in all_record_keys]
        for k in all_record_keys:
            if k not in ordered_headers:
                ordered_headers.append(k)

        ws.append(ordered_headers)

        # Style Header Row
        for col_idx in range(1, len(ordered_headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = self.align_center
            cell.border = self.border_header

        # Write Data Rows
        numeric_keywords = ["amount", "charges", "gst", "fee", "deduction", "dr", "cr", "payable"]
        key_keywords = ["rrn", "utr", "id", "ref", "mobile", "account", "invoice", "number"]

        for row_idx, r in enumerate(records, start=2):
            for col_idx, h in enumerate(ordered_headers, start=1):
                val = r.get(h)
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = self.font_data
                cell.border = self.border_data

                h_lower = h.lower()
                is_num = any(k in h_lower for k in numeric_keywords) and not any(k in h_lower for k in key_keywords)
                is_key = any(k in h_lower for k in key_keywords)

                if val is None or val == "":
                    cell.value = ""
                    cell.alignment = self.align_left
                elif is_num:
                    try:
                        f_val = float(val)
                        cell.value = f_val
                        cell.number_format = "#,##0.00"
                        cell.alignment = self.align_right
                    except (ValueError, TypeError):
                        cell.value = str(val)
                        cell.alignment = self.align_right
                elif is_key:
                    cell.value = str(val)
                    cell.number_format = "@"
                    cell.alignment = self.align_left
                else:
                    cell.value = val
                    cell.alignment = self.align_left

        self._auto_fit_columns(ws)

    def _write_raw_source_sheet(self, sheet_name: str, report: Optional[RawReport]):
        """Writes exact preserved raw source data into dedicated raw tabs.
        For CMS and SMMS, removes the first 3 title/metadata rows so row 1 starts directly with the table headers.
        """
        ws = self.wb.create_sheet(title=sheet_name)
        if not report or not report.raw_matrix:
            ws.append(["Source File Not Provided"])
            return

        # For CMS and SMMS, user requested: "in cms and smms that first 3 lines should be removed it is noyt needed"
        start_idx = report.header_row_idx if sheet_name in ("CMS", "SMMS") else 0
        matrix_to_write = report.raw_matrix[start_idx:]

        for row_idx, row in enumerate(matrix_to_write, start=1):
            # Skip report footer metadata row (e.g. 'Report generated on Date: ...')
            first_val = str(row[0] or "").strip().lower()
            if first_val.startswith("report generated on"):
                continue

            for col_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = self.font_data
                cell.value = val
                # Style header row (which is row 1 when sliced)
                if (start_idx > 0 and row_idx == 1) or (start_idx == 0 and row_idx == (report.header_row_idx + 1)):
                    cell.font = self.font_header
                    cell.fill = self.fill_header
                    cell.border = self.border_header
                    cell.alignment = self.align_center

        self._auto_fit_columns(ws)

    def _write_summary_sheet(self):
        """Builds the Summary tab matching user's exact format:
        Columns: PARTNER, Mode, Network, SUCCESS COUNT, SUM OF TXN AMOUNT, Sum of net amount
        Partner is printed on the first row of each group and blank on subsequent rows.
        Final row is TOTAL.
        Placed at the END of the workbook.
        """
        ws = self.wb.create_sheet(title="Summary")

        headers = ["PARTNER", "Mode", "Network", "SUCCESS COUNT", "SUM OF TXN AMOUNT", "Sum of net amount"]
        ws.append(headers)

        # Style header row
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = self.align_center
            cell.border = self.border_header

        curr_row = 2
        last_partner = None
        for r_data in self.agg.summary_rows:
            partner = r_data["PARTNER"]
            partner_display = partner if partner != last_partner else None
            last_partner = partner

            ws.cell(curr_row, 1, partner_display).alignment = self.align_left
            ws.cell(curr_row, 2, r_data["Mode"]).alignment = self.align_left
            ws.cell(curr_row, 3, r_data["Network"]).alignment = self.align_left

            c_cnt = ws.cell(curr_row, 4, r_data["SUCCESS COUNT"])
            c_cnt.number_format = "#,##0"
            c_cnt.alignment = self.align_right

            c_gross = ws.cell(curr_row, 5, r_data["SUM OF TXN AMOUNT"])
            gross_val = r_data["SUM OF TXN AMOUNT"]
            if gross_val is not None:
                c_gross.number_format = "#,##0.00" if (isinstance(gross_val, float) and not gross_val.is_integer()) else "0"
            c_gross.alignment = self.align_right

            c_net = ws.cell(curr_row, 6, r_data["Sum of net amount"])
            c_net.number_format = "#,##0.00"
            c_net.alignment = self.align_right

            for c in range(1, 7):
                ws.cell(curr_row, c).font = self.font_data
                ws.cell(curr_row, c).border = self.border_header

            curr_row += 1

        # Grand Total Row
        gt = self.agg.grand_total
        ws.cell(curr_row, 1, "TOTAL").alignment = self.align_left
        ws.cell(curr_row, 2, None)
        ws.cell(curr_row, 3, None)

        c_cnt = ws.cell(curr_row, 4, gt["success_count"])
        c_cnt.number_format = "#,##0"
        c_cnt.alignment = self.align_right

        c_gross = ws.cell(curr_row, 5, gt["gross_amount"])
        gross_tot = gt["gross_amount"]
        if gross_tot is not None:
            c_gross.number_format = "#,##0.00" if (isinstance(gross_tot, float) and not gross_tot.is_integer()) else "0"
        c_gross.alignment = self.align_right

        c_net = ws.cell(curr_row, 6, gt["net_amount"])
        c_net.number_format = "#,##0.00"
        c_net.alignment = self.align_right

        for c in range(1, 7):
            cell = ws.cell(curr_row, c)
            cell.font = self.font_data
            cell.border = self.border_header

        self._auto_fit_columns(ws)

    def _write_cms_matched_sheet(self, sheet_name: str, records: List[Dict[str, Any]]):
        """Writes CMS records containing the exact original columns from the CMS source report."""
        ws = self.wb.create_sheet(title=sheet_name)
        if not self.engine.cms_report or not self.engine.cms_report.headers:
            headers = ["No Data Available"]
            ws.append(headers)
            self._auto_fit_columns(ws)
            return

        headers = self.engine.cms_report.headers
        ws.append(headers)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = self.align_center
            cell.border = self.border_header

        numeric_keywords = ["amount", "charges", "gst", "fee", "deduction", "dr", "cr", "payable"]
        key_keywords = ["rrn", "utr", "id", "ref", "mobile", "account", "invoice", "number", "terminal", "code"]

        for row_idx, r in enumerate(records, start=2):
            for col_idx, h in enumerate(headers, start=1):
                val = r.get(h)
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = self.font_data
                cell.border = self.border_data

                h_lower = h.lower()
                is_num = any(k in h_lower for k in numeric_keywords) and not any(k in h_lower for k in key_keywords)
                is_key = any(k in h_lower for k in key_keywords)

                if val is None or val == "":
                    cell.value = ""
                    cell.alignment = self.align_left
                elif is_num:
                    try:
                        f_val = float(val)
                        cell.value = f_val
                        cell.number_format = "#,##0.00" if not f_val.is_integer() else "0"
                        cell.alignment = self.align_right
                    except (ValueError, TypeError):
                        cell.value = str(val)
                        cell.alignment = self.align_right
                elif is_key:
                    cell.value = str(val)
                    cell.number_format = "@"
                    cell.alignment = self.align_left
                else:
                    cell.value = val
                    cell.alignment = self.align_left

        self._auto_fit_columns(ws)

    def build_and_save(self, output_path: str):
        """Generates sheets in exact order:
        1. CMS
        2. SMMS
        3. Easebuzz
        4. Cashfree
        5. Airtel
        (Airtel Settlement if uploaded)
        (cms_failure_txn if CMS has failed records)
        6. CMS_SMMS_Matched
        7. CMS_CF_Matched
        8. CMS_EB_Matched
        9. CMS_Air_Matched
        10. Unmatched sheets (ONLY IF UNMATCHED > 0):
            - CMS_Not_in_SMMS
            - SMMS_Not_in_CMS
            - Unmatched_CF
            - Unmatched_EB
            - Unmatched_Airtel
            - Duplicates (if > 0)
            - Exceptions (if > 0)
        11. Summary (IN THE LAST POSITION!)
        """
        # 1. Raw Preserved Source Tabs
        self._write_raw_source_sheet("CMS", self.engine.cms_report)
        self._write_raw_source_sheet("SMMS", self.engine.smms_report)
        self._write_raw_source_sheet("Easebuzz", self.engine.eb_report)
        self._write_raw_source_sheet("Cashfree", self.engine.cf_report)
        self._write_raw_source_sheet("Airtel", self.engine.airtel_report)
        if self.engine.settle_report:
            self._write_raw_source_sheet("Airtel Settlement", self.engine.settle_report)

        # 2. CMS Non-Success Transactions by Status (only created if records exist)
        # Status 2: cms_failure_txn
        # Status 3: cms_status_3_txn
        # Status 4: cms_status_4_txn
        # Status 5: cms_status_5_txn
        # Status 6: cms_status_6_txn
        cms_by_status = defaultdict(list)
        if self.engine.cms_report:
            for r in self.engine.cms_report.records:
                sp_id = clean_key(r.get("SwinkPay Txn ID") or r.get("SwinkPay Transaction ID"))
                rrn = clean_key(r.get("RRN/UTR") or r.get("RRN") or r.get("UTR"))
                if not sp_id and not rrn:
                    continue

                st_raw = r.get("Transaction Status")
                st_norm = normalize_status(st_raw)
                st_clean = clean_key(st_raw).strip()

                if st_clean == "2" or st_norm == "Failed":
                    cms_by_status["2"].append(r)
                elif st_clean == "3" or st_norm == "Status 3":
                    cms_by_status["3"].append(r)
                elif st_clean == "4" or st_norm == "Status 4":
                    cms_by_status["4"].append(r)
                elif st_clean == "5" or st_norm == "Status 5":
                    cms_by_status["5"].append(r)
                elif st_clean == "6" or st_norm == "Status 6":
                    cms_by_status["6"].append(r)

        if cms_by_status["2"]:
            self._write_cms_matched_sheet("cms_failure_txn", cms_by_status["2"])
        if cms_by_status["3"]:
            self._write_cms_matched_sheet("cms_status_3_txn", cms_by_status["3"])
        if cms_by_status["4"]:
            self._write_cms_matched_sheet("cms_status_4_txn", cms_by_status["4"])
        if cms_by_status["5"]:
            self._write_cms_matched_sheet("cms_status_5_txn", cms_by_status["5"])
        if cms_by_status["6"]:
            self._write_cms_matched_sheet("cms_status_6_txn", cms_by_status["6"])

        # 3. Matched Result Tabs (with exact CMS source columns)
        self._write_cms_matched_sheet("CMS_SMMS_Matched", self.engine.cms_smms_matched)
        self._write_cms_matched_sheet("CMS_CF_Matched", self.engine.cms_cf_matched)
        self._write_cms_matched_sheet("CMS_EB_Matched", self.engine.cms_eb_matched)
        self._write_cms_matched_sheet("CMS_Air_Matched", self.engine.cms_air_matched)

        # 4. Discrepancy & Unmatched Tabs (ONLY IF RECORDS EXIST)
        if self.engine.cms_not_in_smms:
            self._write_table_sheet("CMS_Not_in_SMMS", self.engine.cms_not_in_smms)
        if self.engine.smms_not_in_cms:
            self._write_table_sheet("SMMS_Not_in_CMS", self.engine.smms_not_in_cms)
        if self.engine.unmatched_cf:
            self._write_table_sheet("Unmatched_CF", self.engine.unmatched_cf)
        if self.engine.unmatched_eb:
            self._write_table_sheet("Unmatched_EB", self.engine.unmatched_eb)
        if self.engine.unmatched_airtel:
            self._write_table_sheet("Unmatched_Airtel", self.engine.unmatched_airtel)
        if self.engine.duplicates:
            self._write_table_sheet("Duplicates", self.engine.duplicates)
        if self.engine.exceptions:
            self._write_table_sheet("Exceptions", self.engine.exceptions)

        # 5. Summary Sheet (IN THE LAST POSITION)
        self._write_summary_sheet()

        self.wb.save(output_path)
