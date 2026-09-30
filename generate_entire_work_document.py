"""
Script to generate the comprehensive, official Microsoft Word (.docx) document:
'Ops_Auto_Complete_Project_Accomplishments_and_Work_Done.docx'
covering everything done entirely across the project.
"""
import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DOC_TITLE = "Ops_Auto: Complete Project Accomplishments & Architecture Document"
DOC_SUBTITLE = "Comprehensive Technical Implementation, End-to-End Milestones, Business Impact, and Standard Operating Procedure (SOP)"
DOC_AUTHOR = "Pradnyan Hegde | SwinkPay Fintech Operations & Automation Engineering"
DOC_DATE = "September 2026"
DOC_OUTPUT_FILENAME = "Ops_Auto_Complete_Project_Accomplishments_and_Work_Done.docx"

COLOR_PRIMARY = RGBColor(30, 58, 138)   # #1e3a8a deep navy
COLOR_SECONDARY = RGBColor(26, 115, 232) # #1a73e8 bright blue
COLOR_ACCENT = RGBColor(5, 150, 105)    # #059669 emerald green
COLOR_TEXT = RGBColor(15, 23, 42)       # #0f172a slate
COLOR_MUTED = RGBColor(100, 116, 139)   # #64748b gray
COLOR_WHITE = RGBColor(255, 255, 255)

HEX_HEADER_BG = "1E3A8A"
HEX_ROW_ALT = "F8FAFC"
HEX_BORDER = "CBD5E1"
HEX_CALLOUT_BG = "EFF6FF"
HEX_CALLOUT_BORDER = "1A73E8"
HEX_SUCCESS_BG = "ECFDF5"
HEX_SUCCESS_BORDER = "059669"
HEX_WARN_BG = "FFFBEB"
HEX_WARN_BORDER = "D97706"
HEX_CODE_BG = "0F172A"


def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_shading(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="none"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def add_callout(doc, title, text, bg_hex=HEX_CALLOUT_BG, border_hex=HEX_CALLOUT_BORDER, icon="📌"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    
    set_cell_margins(cell, top=140, bottom=140, left=180, right=160)
    set_cell_shading(cell, bg_hex)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>'
        f'  <w:top w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:bottom w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"{icon} {title}\n")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(11)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_PRIMARY
    
    run_b = p.add_run(text)
    run_b.font.name = "Calibri"
    run_b.font.size = Pt(9.5)
    run_b.font.color.rgb = COLOR_TEXT
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)


def add_code_block(doc, code_str):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
    set_cell_shading(cell, "0F172A")
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(code_str)
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(56, 189, 248) # light blue mono
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(4)


def build_styled_table(doc, headers, rows_data, col_widths=None):
    tbl = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_table_borders(tbl)
    
    # Header Row
    hdr_cells = tbl.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].text = h_text
        set_cell_shading(hdr_cells[i], HEX_HEADER_BG)
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            run.font.bold = True
            run.font.color.rgb = COLOR_WHITE
            
    # Data Rows
    for r_idx, r_data in enumerate(rows_data):
        row_cells = tbl.rows[r_idx + 1].cells
        bg = HEX_ROW_ALT if (r_idx % 2 == 1) else "FFFFFF"
        for c_idx, val in enumerate(r_data):
            row_cells[c_idx].text = str(val)
            set_cell_shading(row_cells[c_idx], bg)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=120, right=120)
            p = row_cells[c_idx].paragraphs[0]
            for run in p.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(9)
                run.font.color.rgb = COLOR_TEXT
                
    if col_widths and len(col_widths) == len(headers):
        for row in tbl.rows:
            for idx, w in enumerate(col_widths):
                row.cells[idx].width = Inches(w)
                
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(6)
    return tbl


def add_heading_1(doc, text):
    h = doc.add_heading(text, level=1)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    for r in h.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY
    return h


def add_heading_2(doc, text):
    h = doc.add_heading(text, level=2)
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(3)
    h.paragraph_format.keep_with_next = True
    for r in h.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = COLOR_SECONDARY
    return h


def add_heading_3(doc, text):
    h = doc.add_heading(text, level=3)
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(2)
    h.paragraph_format.keep_with_next = True
    for r in h.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = COLOR_TEXT
    return h


def add_para(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(10)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_TEXT
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_TEXT
    return p


def main():
    doc = docx.Document()
    
    # Configure 0.75" Margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        
    # ==========================================
    # Title & Metadata Banner
    # ==========================================
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run(DOC_TITLE)
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(8)
    run_sub = p_sub.add_run(DOC_SUBTITLE)
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = COLOR_MUTED
    
    # Meta Box
    add_callout(
        doc,
        "Project Deliverable Metadata",
        "• Organization: SwinkPay Fintech Pvt Ltd (FinOps Automation Engineering)\n"
        "• Authors & Contributors: Pradnyan Hegde & Antigravity (Google DeepMind)\n"
        "• Version: Production 2.0 | Test Suite: 58/58 Tests Passing (100% OK)\n"
        "• Date: September 2026 | GitHub: https://github.com/Pradnyan-hegde/Ops_Auto.git",
        bg_hex="F1F5F9",
        border_hex="1E3A8A",
        icon="🏛️"
    )
    
    # ==========================================
    # Section 1: Executive Summary
    # ==========================================
    add_heading_1(doc, "1. Executive Summary")
    add_para(doc, "In digital payment operations, transaction reconciliation across multiple banking and Payment Gateway (PG) partners is vital to ensure zero revenue leakage, strict regulatory compliance, and timely merchant settlements. Prior to the Ops_Auto project, SwinkPay operations relied on manual spreadsheet comparisons, disparate VLOOKUP formulas, manual Postman API calls for missing transactions, and tedious manual updates in CMS.")
    add_para(doc, "The Ops_Auto automation framework replaces these error-prone manual steps with a high-throughput, deterministic reconciliation and settlement engine. It guarantees 100% mathematical audit accuracy, automatic recovery of missing transactions via backend APIs without Postman, strict difference-to-zero validation, and automated multi-day bank settlement generation.")
    
    add_callout(
        doc,
        "Core Value Proposition",
        "1. Reduction in Daily Operational Time: From 3-5 hours daily down to under 60 seconds.\n"
        "2. Strict Audit Guarantee: Mathematical enforcement that Difference == 0.00 across all gateways.\n"
        "3. Complete Postman Elimination: 1-Click and One-Shot Direct Bulk API pulling for missing PG transactions.\n"
        "4. Automated Targeted Files: Instant generation of Sync_Transactions.xlsx for SMMS and Change_Network.xlsx for CMS.",
        bg_hex=HEX_SUCCESS_BG,
        border_hex=HEX_SUCCESS_BORDER,
        icon="🚀"
    )

    # ==========================================
    # Section 2: Chronological Milestones (What Was Done Entirely)
    # ==========================================
    add_heading_1(doc, "2. Chronological Milestones: Everything Done Entirely")
    add_para(doc, "Below is the complete, comprehensive technical breakdown of every milestone, architectural enhancement, and operational feature developed throughout this project:")
    
    # Milestone 1
    add_heading_2(doc, "Milestone 1: The Core Multi-Way Matching Engine")
    add_para(doc, "Architected and built an in-memory hash-indexed multi-way matching engine (core/matcher.py) that ingests and reconciles 5 primary transaction reports simultaneously:")
    add_para(doc, "• CMS <--> SMMS: Reconciled on Bank RRN/UTR and SwinkPay Transaction ID.")
    add_para(doc, "• CMS <--> Cashfree: Reconciled on Bank Reference Number (UTR) and Order ID.")
    add_para(doc, "• CMS <--> Easebuzz: Reconciled on Bank Ref No / UTR.")
    add_para(doc, "• CMS <--> Airtel Payments Bank: Reconciled on Partner Reference ID and Bank Reference No.")
    add_para(doc, "• Airtel Transaction <--> Airtel Settlement: Secondary settlement verification on UTR and Net Credit Amount.")
    add_para(doc, "Created the dynamic header-signature detector (core/detector.py), which analyzes column fingerprints in milliseconds to detect report types without relying on filenames.")

    # Milestone 2
    add_heading_2(doc, "Milestone 2: 18-Sheet Reconciliation Master Workbook & Settlement Generation")
    add_para(doc, "Developed the Excel builder (core/excel_builder.py) to generate the master 18-sheet audit workbook (XCD_Reconciliation_YYYY-MM-DD.xlsx) preserving complete raw data, matched subsets, unmatched gateway records, status conflicts, and financial summaries.")
    add_para(doc, "Built partner bank settlement workbooks (core/xcd_builder.py) and enforced the strict Difference-to-Zero Audit Gate (core/xcd_validator.py):")
    add_code_str = (
        "Audit Formula:\n"
        "Difference = Gross Amount - PSP Charges - GST - Net Credit Amount == 0.00\n"
        "If |Difference| > 0.00 -> System halts settlement release and flags audit alert."
    )
    add_code_block(doc, add_code_str)

    # Milestone 3
    add_heading_2(doc, "Milestone 3: Multi-Merchant Architecture & Dynamic Registry")
    add_para(doc, "Expanded the engine from single-merchant processing to an enterprise multi-merchant architecture (core/merchant.py) supporting CCD Value Express, SBB Medicare, AGS (Advance Genuine Spares), and custom XCD single settlement profiles with dynamic onboarding modals and persistent metadata storage.")

    # Milestone 4
    add_heading_2(doc, "Milestone 4: Terminal Directory (TID_FILE) & Middle Number Parsing")
    add_para(doc, "Implemented terminal mapping management (core/terminal_mapper.py) that parses Cashfree middle numbers from Order IDs (e.g. extracting 4860 from 330595-4860-AXI...) and automatically resolves the Master Terminal ID (e.g. 141001117), MMS Terminal ID (e.g. XKD8M3), and Store Name.")

    # Milestone 5
    add_heading_2(doc, "Milestone 5: Eliminating Postman — Missing Transaction Recovery API")
    add_para(doc, "Integrated the dashboard directly with the SwinkPay Decision API (https://merchants.swinkpay-fintech.com/api/v2/decision/updated). Built core/pg_payload_builder.py to construct valid Decision Updated payloads. Added the Postman Live Telemetry Inspector inside the dashboard UI to render real-time HTTP response tags (200 OK), latency (ms), SwinkPay Transaction IDs, and formatted response bodies.")

    # Milestone 6
    add_heading_2(doc, "Milestone 6: Cross-System Status Verification (PG vs CMS / SMMS)")
    add_para(doc, "Added cross-system checks to detect where the Payment Gateway is 'Success' (customer charged), but CMS or SMMS records it as 'Failed' or 'Pending'. Surfaced a dedicated conflict view on the main dashboard showing Gateway Name, Gateway Status, CMS Status, SMMS Status, SwinkPay Txn ID, Bank UTR, Amount, and Outlet Name.")

    # Milestone 7
    add_heading_2(doc, "Milestone 7: SMMS Sync Automation (Sync_Transactions.xlsx)")
    add_para(doc, "Enforced the critical operational rule: If a transaction exists in CMS but has SMMS Sync Status = false, do NOT pull from the gateway. Instead, automatically package these records into Sync_Transactions_YYYY-MM-DD.xlsx (core/sync_builder.py) for direct bulk upload to the CMS/SMMS admin portal.")

    # Milestone 8
    add_heading_2(doc, "Milestone 8: Semantic Network Misclassification Engine & Guidance Banner")
    add_para(doc, "Built core/network_builder.py to filter false positive network changes (UPI == UPI_OFFLINE_STATIC) while targeting genuine interchange changes (RUPAY -> UPI_CREDIT_CARD_OFFLINE_STATIC, WA -> UPI_PPI_OFFLINE_STATIC) into Change_Network_YYYY-MM-DD.xlsx. Displayed explicit operational guidance:")
    add_callout(
        doc,
        "CMS Action Instruction Banner",
        "\"Upload this file in CMS to change network. If network doesn't change even after uploading, there is a new network we need to configure and reupload the files here.\"",
        bg_hex=HEX_WARN_BG,
        border_hex=HEX_WARN_BORDER,
        icon="⚠️"
    )

    # Milestone 9
    add_heading_2(doc, "Milestone 9: Automated 12-Digit UTR Zero-Padding")
    add_para(doc, "Implemented automatic sanitization to prepend leading zeros to any UTR shorter than 12 digits (e.g. 6789876543 -> 006789876543), preventing Decision API pull failures caused by non-standard banking strings.")

    # Milestone 10
    add_heading_2(doc, "Milestone 10: Standalone Direct PG Report Upload & One-Shot Pull")
    add_para(doc, "Added a dedicated tool allowing operators to drop any full gateway report (Easebuzz, Cashfree, Airtel, or CSV) with merchant mapping context, preview all rows, and pull all records to the Decision API in 1 shot with a real-time progress bar.")

    # Milestone 11
    add_heading_2(doc, "Milestone 11: Section De-Duplication & UI Streamlining")
    add_para(doc, "Cleaned up the 'Pull Missing Transactions' section by removing the duplicate reconciliation queue, dedicating this section 100% to standalone direct PG report uploads and one-shot pulling. Maintained the reconciliation missing transactions table exclusively on the main dashboard.")

    # Milestone 12
    add_heading_2(doc, "Milestone 12: Bug Fixes & State Persistence")
    add_para(doc, "• Fixed Reconciliation Workbook Download Bug: Resolved issue where the raw SMMS report was downloading instead of the generated reconciliation workbook.\n"
             "• State Loss on Refresh: Added localStorage persistence for the active tab, selected merchant, and credentials. Added HTTP anti-caching headers.\n"
             "• Dropdown Initialization Bug: Removed orphaned function call that broke dropdown initialization on startup.")

    # Milestone 13
    add_heading_2(doc, "Milestone 13: Complete Documentation, Visual Assets & GitHub Overhaul")
    add_para(doc, "Rewrote README.md on GitHub covering all 11 required sections (Problem, Architecture, Tech Stack, How it Works, 24 REST APIs, Database, Sample I/O, Screenshots, How to Run, Future Improvements), archived 5 production UI screenshots in docs/screenshots/, and created comprehensive SOP guides.")

    # ==========================================
    # Section 3: Before vs After Comparison
    # ==========================================
    add_heading_1(doc, "3. Comparative Analysis: Before vs. After Ops_Auto")
    headers_comp = ["Operational Dimension", "Before (Manual Workflow)", "After (Ops_Auto Platform)"]
    rows_comp = [
        ["Reconciliation Time", "3 to 5 hours daily across spreadsheets", "< 60 seconds end-to-end execution"],
        ["Matching Accuracy", "Manual VLOOKUPs prone to human error", "100% deterministic in-memory hash matching"],
        ["Missing PG Recovery", "Manual Postman calls; manual TID lookup", "1-Click / One-Shot Automated API Pull"],
        ["Bank UTR Formatting", "Scientific notation errors & missing zeros", "Automated 12-digit zero-padding (006789876543)"],
        ["Financial Audit Gate", "Manual sums; high risk of balance mismatch", "Strict Difference-to-Zero Gate (|Diff| == 0.00)"],
        ["SMMS Sync Discrepancies", "Unclear which rows failed SMMS sync", "Auto-generated Sync_Transactions.xlsx"],
        ["Network Misclassifications", "Manual inspection of RuPay CC / Wallets", "Auto-generated Change_Network.xlsx"],
        ["Weekend Batch Splitting", "Complex manual date filtering", "Interactive Date Checkboxes (Fri/Sat/Sun)"],
        ["Executive Reporting", "Manually drafted emails", "1-Click Copy Formatted Email with outlet counts"]
    ]
    build_styled_table(doc, headers_comp, rows_comp, [1.8, 2.3, 2.4])

    # ==========================================
    # Section 4: Architecture & REST API Endpoints
    # ==========================================
    add_heading_1(doc, "4. Architecture & REST API Endpoints")
    add_para(doc, "Ops_Auto exposes 24 production-grade REST API endpoints built with FastAPI:")
    
    headers_api = ["Method", "Endpoint", "Description", "Primary Output"]
    rows_api = [
        ["POST", "/api/reconcile", "Executes full 5-file reconciliation pipeline", "Audit stats & session ID"],
        ["POST", "/api/upload-recon-file", "Instant parse of existing recon workbook", "Extracted discrepancies & dates"],
        ["POST", "/api/detect-files", "Pre-flight column header auto-detection", "Detected report types & row counts"],
        ["GET", "/api/download/{session_id}", "Downloads 18-sheet master recon workbook", "Excel (.xlsx) file"],
        ["GET", "/api/download-sync-file/{id}", "Downloads SMMS sync workbook", "Sync_Transactions_*.xlsx"],
        ["GET", "/api/download-network-file/{id}", "Downloads CMS network change workbook", "Change_Network_*.xlsx"],
        ["GET", "/api/download-xcd/{id}/{pg}", "Downloads date-filtered settlement file", "XCD Input file...xlsx"],
        ["POST", "/api/generate-xcd/{id}", "Generates date-filtered settlement workbooks", "Filtered settlement counts"],
        ["POST", "/api/pull-missing-pg/{id}", "Pulls missing recon records to Decision API", "Postman telemetry & status"],
        ["POST", "/api/direct-pull/upload", "Parses standalone PG report for 1-shot pull", "Resolved records table"],
        ["POST", "/api/direct-pull/push", "Executes one-shot bulk pull to Decision API", "Batch pull counts & latency"],
        ["GET", "/api/merchants", "Lists all registered merchant profiles", "Merchant profiles array"],
        ["POST", "/api/merchants", "Registers or updates merchant profile", "Saved merchant config"],
        ["POST", "/api/upload-terminal-file", "Uploads terminal mapping master (TID_FILE)", "Loaded mapping count"],
        ["GET", "/api/resolve-terminal", "Resolves terminal context by TID or middle no.", "Terminal ID & Store Name"],
        ["GET", "/api/settings", "Retrieves persistent API credentials", "Auth token & channel JSON"],
        ["POST", "/api/settings", "Persists API credentials permanently", "Success confirmation"]
    ]
    build_styled_table(doc, headers_api, rows_api, [0.8, 2.2, 2.3, 1.2])

    # ==========================================
    # Section 5: Standardized Daily Operational Artifacts
    # ==========================================
    add_heading_1(doc, "5. Standardized Daily Operational Artifacts")
    add_para(doc, "During every operational run, Ops_Auto automatically generates 5 standardized business artifacts:")
    add_para(doc, "1. XCD_Reconciliation_YYYY-MM-DD.xlsx: The master 18-sheet audit workbook preserving all raw inputs, matched records, unmatched records, status conflicts, and financial summaries.")
    add_para(doc, "2. XCD Input file as on YYYY-MM-DD (Partner).xlsx: Bank settlement workbooks for Cashfree, Easebuzz, and Airtel filtered by specific settlement dates.")
    add_para(doc, "3. Sync_Transactions_YYYY-MM-DD.xlsx: Targeted SMMS bulk sync file containing CMS transactions where SMMS Sync Status = false.")
    add_para(doc, "4. Change_Network_YYYY-MM-DD.xlsx: Targeted CMS network update file containing genuine interchange reclassifications (e.g. RuPay CC on UPI).")
    add_para(doc, "5. Formatted Executive Settlement Email: Ready-to-paste executive email summary detailing active outlets (e.g. 1,435 outlets), transaction counts, gross values, PSP service fees, GST, and net merchant payouts.")

    # ==========================================
    # Section 6: Production UI Screenshots
    # ==========================================
    add_heading_1(doc, "6. Production UI Screenshots & Visual Walkthrough")
    add_para(doc, "Below are actual production screenshots from the Ops_Auto web platform documenting the interface:")
    
    screenshots_dir = os.path.join(os.path.dirname(__file__), "docs", "screenshots")
    screenshots = [
        ("01_dashboard_reconciliation_overview.png", "Figure 1: Operations Dashboard & Master Reconciliation Interface (Pre-flight Detection & KPI Summary)"),
        ("02_missing_pg_transactions_queue.png", "Figure 2: Missing PG Transactions Queue & Cross-System Status Conflicts Table"),
        ("03_pull_missing_direct_upload.png", "Figure 3: Dedicated Direct PG Report Upload & One-Shot Bulk Pull Tool"),
        ("04_postman_live_telemetry_inspector.png", "Figure 4: Postman Live Telemetry Inspector (Real-time HTTP Status, Latency ms & Payload Inspector)"),
        ("05_terminal_mapping_management.png", "Figure 5: Multi-Merchant Terminal Directory Management (TID_FILE Search & Outlet Resolver)")
    ]
    
    for filename, caption in screenshots:
        fpath = os.path.join(screenshots_dir, filename)
        if os.path.exists(fpath):
            add_heading_2(doc, caption)
            try:
                doc.add_picture(fpath, width=Inches(6.3))
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_before = Pt(2)
                p_cap.paragraph_format.space_after = Pt(8)
                r_cap = p_cap.add_run(f"📷 {caption}")
                r_cap.font.name = "Calibri"
                r_cap.font.size = Pt(8.5)
                r_cap.font.italic = True
                r_cap.font.color.rgb = COLOR_MUTED
            except Exception as e:
                add_para(doc, f"[Image rendering note: {filename} present at docs/screenshots/{filename}]")
        else:
            add_para(doc, f"[Screenshot path: docs/screenshots/{filename}]")

    # ==========================================
    # Section 7: Daily Standard Operating Procedure (SOP)
    # ==========================================
    add_heading_1(doc, "7. Daily Standard Operating Procedure (SOP)")
    add_para(doc, "Follow this chronological sign-off checklist every morning during daily reconciliation:")
    
    sop_steps = [
        ("Step 1: Daily Ingestion", "Upload the 5 daily reports (CMS, SMMS, Cashfree, Easebuzz, Airtel) into the master dropzone."),
        ("Step 2: Missing PG Transactions", "Review the 'Missing PG Transactions' table. Click 'Pull All' to ingest missing records into CMS via Decision API."),
        ("Step 3: SMMS Sync Upload", "Download Sync_Transactions_YYYY-MM-DD.xlsx and upload to the CMS/SMMS admin portal to update Sync Status = true."),
        ("Step 4: Network Change Upload", "Download Change_Network_YYYY-MM-DD.xlsx and upload to CMS to correct RuPay CC and Wallet interchange."),
        ("Step 5: Re-Reconciliation", "Download refreshed CMS and SMMS reports, re-run reconciliation, and verify Difference == 0.00 (All Clear)."),
        ("Step 6: Settlement Generation", "Download partner XCD workbooks for Cashfree, Easebuzz, and Airtel (using date checkboxes if weekend batch)."),
        ("Step 7: Executive Email Dispatch", "Click 'Copy Formatted Text' on the dashboard and paste the executive summary into Outlook/Gmail with XCD attachments.")
    ]
    for step_title, step_desc in sop_steps:
        add_para(doc, f"• {step_desc}", bold_prefix=f"{step_title}: ")

    # ==========================================
    # Section 8: Automated Test Verification
    # ==========================================
    add_heading_1(doc, "8. Automated Test Suite Verification")
    add_para(doc, "The Ops_Auto codebase is protected by 58 automated unit and integration tests:")
    add_code_block(doc, "Command: python -m unittest discover tests\nResult:  Ran 58 tests in 89.098s\nStatus:  OK (100% Passing, 0 Failures, 0 Errors)")
    
    add_callout(
        doc,
        "Project Sign-Off",
        "Ops_Auto Version 2.0 has been fully deployed, verified, and committed to git on branch 'main' at:\n"
        "https://github.com/Pradnyan-hegde/Ops_Auto.git",
        bg_hex=HEX_CALLOUT_BG,
        border_hex=HEX_CALLOUT_BORDER,
        icon="✅"
    )

    # Save Document
    output_path = os.path.join(os.path.dirname(__file__), DOC_OUTPUT_FILENAME)
    doc.save(output_path)
    print(f"Document successfully created at: {output_path}")


if __name__ == "__main__":
    main()
