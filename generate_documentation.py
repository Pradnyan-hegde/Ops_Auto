"""
Script to generate comprehensive technical and operational documentation for
SwinkPay Ops_Auto Reconciliation Automation.
Generates:
1. SWINKPAY_RECONCILIATION_AUTOMATION_DOCUMENTATION.docx
2. SWINKPAY_RECONCILIATION_AUTOMATION_DOCUMENTATION.md
3. Copies .docx to user's Desktop for convenient sharing.
"""
import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DOC_TITLE = "SwinkPay Daily Reconciliation & Operations Automation"
DOC_SUBTITLE = "Comprehensive Technical Architecture, Automated Workflows, and Standard Operating Procedure (SOP)"
DOC_VERSION = "Version 2.0 (Production)"
DOC_DATE = "September 2026"
DOC_AUTHOR = "SwinkPay Fintech Operations & Automation Engineering"

# Styling constants
COLOR_NAVY = RGBColor(15, 23, 42)      # #0f172a
COLOR_PRIMARY = RGBColor(30, 58, 138)   # #1e3a8a
COLOR_ACCENT = RGBColor(79, 70, 229)    # #4f46e5
COLOR_TEXT = RGBColor(30, 41, 59)       # #1e293b
COLOR_MUTED = RGBColor(100, 116, 139)   # #64748b
COLOR_WHITE = RGBColor(255, 255, 255)

HEX_HEADER_BG = "1E3A8A"
HEX_ROW_ALT = "F8FAFC"
HEX_BORDER = "CBD5E1"
HEX_CALLOUT_BG = "F0F4FF"
HEX_CALLOUT_BORDER = "4F46E5"
HEX_CODE_BG = "F1F5F9"


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets cell padding in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_shading(cell, color_hex):
    """Applies background color to a table cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    """Applies standard subtle borders to a table."""
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


def add_callout(doc, title, text, bg_hex=HEX_CALLOUT_BG, border_hex=HEX_CALLOUT_BORDER):
    """Adds a stylish callout box with a colored left border."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    
    set_cell_margins(cell, top=140, bottom=140, left=200, right=180)
    set_cell_shading(cell, bg_hex)
    
    # Custom border: thick left border, no other borders
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
    run_t = p.add_run(f"📌 {title}\n")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(10.5)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_PRIMARY
    
    run_b = p.add_run(text)
    run_b.font.name = "Calibri"
    run_b.font.size = Pt(9.5)
    run_b.font.color.rgb = COLOR_TEXT
    
    # Empty paragraph after table for spacing
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)


def add_code_block(doc, code_text):
    """Adds a formatted code / terminal / diagram block."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    
    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
    set_cell_shading(cell, HEX_CODE_BG)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:left w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:right w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(code_text)
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(30, 41, 59)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)


def create_styled_table(doc, headers, rows_data, col_widths=None):
    """Creates a professional styled table."""
    table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color=HEX_BORDER)
    
    # Header Row
    hdr_cells = table.rows[0].cells
    for i, h_text in enumerate(headers):
        cell = hdr_cells[i]
        set_cell_shading(cell, HEX_HEADER_BG)
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = "Calibri"
        run.font.size = Pt(9.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_WHITE
        if col_widths and i < len(col_widths):
            cell.width = Inches(col_widths[i])
            
    # Data Rows
    for row_idx, r_data in enumerate(rows_data, start=1):
        row_cells = table.rows[row_idx].cells
        bg_color = HEX_ROW_ALT if row_idx % 2 == 0 else "FFFFFF"
        for col_idx, cell_value in enumerate(r_data):
            cell = row_cells[col_idx]
            if bg_color != "FFFFFF":
                set_cell_shading(cell, bg_color)
            set_cell_margins(cell, top=80, bottom=80, left=140, right=140)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(str(cell_value))
            run.font.name = "Calibri"
            run.font.size = Pt(9.0)
            run.font.color.rgb = COLOR_TEXT
            if col_widths and col_idx < len(col_widths):
                cell.width = Inches(col_widths[col_idx])
                
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)
    return table


def build_word_document(output_path: str):
    """Builds the complete professional Word Document."""
    doc = docx.Document()
    
    # Page Setup (1 inch margins)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Styles setup
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Calibri'
    font_normal.size = Pt(10.5)
    font_normal.color.rgb = COLOR_TEXT
    
    # Title Block
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run(DOC_TITLE)
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run(DOC_SUBTITLE)
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(12)
    run_sub.font.color.rgb = COLOR_MUTED
    
    # Metadata Table
    meta_data = [
        ["Document Version", DOC_VERSION, "Author / Team", DOC_AUTHOR],
        ["Release Date", DOC_DATE, "Target Audience", "Operations, Automation Engineers, FinOps, IT"],
        ["System Repository", "Ops_Auto (SwinkPay Fintech)", "Production Status", "Active / Deployed"]
    ]
    create_styled_table(doc, ["Attribute", "Value", "Attribute", "Value"], meta_data, [1.5, 1.8, 1.5, 1.7])
    
    add_callout(
        doc,
        "Purpose of this Document",
        "This document provides the definitive operational and architectural reference for SwinkPay's "
        "Automated Daily Transaction Reconciliation Engine (Ops_Auto). It documents the full reconciliation "
        "pipeline, the automated resolution workflows (PG Missing Pull, SMMS Sync File, and Network Change File), "
        "and outlines the precise standard operating procedures (SOP) for engineering teams automating these steps."
    )
    
    # Helper to add section headers
    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = COLOR_ACCENT
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = COLOR_NAVY
        return p

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_TEXT
        return p

    # SECTION 1
    add_h1("1. Executive Summary & Problem Statement")
    add_p(
        "In digital payment operations, transaction reconciliation across multiple banking and payment gateway "
        "(PG) partners is critical to ensure zero revenue leakage, regulatory compliance, and timely merchant settlements. "
        "Historically, SwinkPay operations relied on manual spreadsheet comparisons, disparate VLOOKUP formulas, "
        "manual Postman API calls for missing transactions, and cumbersome manual data corrections in CMS."
    )
    add_p(
        "The Ops_Auto automation framework replaces these error-prone manual steps with a high-throughput, deterministic "
        "reconciliation pipeline. It guarantees 100% audit accuracy, zero missing data silently ignored, multi-day weekend "
        "handling, automated API pulling for missing PG transactions, and targeted Excel generation for CMS/SMMS backend updates."
    )
    
    # Key Metrics Table
    metrics_data = [
        ["Audit Balancing Guarantee", "Difference to zero (0.00) enforced across all gateways", "Guaranteed by XCD Validator"],
        ["Multi-Day Support", "Full date-specific breakdown across weekend batches (Fri, Sat, Sun)", "Built-in Date Filter Engine"],
        ["Missing PG Pulling", "Terminal ID resolution & automated API pulling without Postman", "Terminal Mapper + Pull API"],
        ["SMMS Sync Automation", "Auto-detection of un-synced CMS transactions into formatted file", "Sync_Transactions_YYYY-MM-DD.xlsx"],
        ["Network Change File", "Filters out false positives; targets only genuine misclassifications", "Change_Network_YYYY-MM-DD.xlsx (91 rows)"]
    ]
    create_styled_table(doc, ["Core Capability", "Operational Impact", "Implementation Module"], metrics_data, [2.0, 3.0, 1.5])
    
    # SECTION 2
    add_h1("2. End-to-End System Architecture")
    add_p(
        "Ops_Auto is structured as a modular engine with decoupled ingestion, matching, discrepancy isolation, "
        "specialized file generation, and interactive dashboard services:"
    )
    
    arch_ascii = (
        "+-----------------------------------------------------------------------------------------+\n"
        "|                              DAILY INPUT SOURCES (T+1)                                  |\n"
        "|   [CMS Report]   [SMMS Report]   [Cashfree PG]   [Easebuzz PG]   [Airtel Bank]   [Settlement]   |\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "                                             |\n"
        "                                             v\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "|                        HEADER-BASED DETECTION & DATA NORMALIZER                         |\n"
        "|   - Inspects headers (no guesswork)              - Normalizes 12-digit RRN / UTRs       |\n"
        "|   - Converts string numbers to exact floats      - Cleans leading zeros & timestamps    |\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "                                             |\n"
        "                                             v\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "|                             MULTI-WAY MATCHING CORE ENGINE                              |\n"
        "|   Step 1: CMS <--> SMMS matching via RRN/UTR & SwinkPay Txn ID                          |\n"
        "|   Step 2: Partner matching (Cashfree via Ref ID, Easebuzz via UTR, Airtel via Txn ID)    |\n"
        "|   Step 3: Airtel Settlement report verification (Net amount, UTR, date matching)        |\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "                                             |\n"
        "                    +------------------------+------------------------+\n"
        "                    |                                                 |\n"
        "                    v                                                 v\n"
        "+---------------------------------------+   +---------------------------------------------+\n"
        "|          MATCHED TRANSACTIONS         |   |             DISCREPANCY ISOLATION           |\n"
        "| - CMS_SMMS_Matched  - CMS_CF_Matched  |   | 1. Missing in CMS (Unmatched PG)            |\n"
        "| - CMS_EB_Matched    - CMS_Air_Matched |   | 2. Missing in SMMS (CMS_Not_in_SMMS)        |\n"
        "| - Multi-Day Date Checkbox Filter      |   | 3. Network Discrepancies (CMS vs PG Mode)   |\n"
        "+---------------------------------------+   | 4. Failed / Reversed Transactions           |\n"
        "                    |                       +---------------------------------------------+\n"
        "                    v                                                 |\n"
        "+---------------------------------------+                             v\n"
        "|       OUTPUT WORKBOOKS (XCD)          |   +---------------------------------------------+\n"
        "| - XCD Input file (Cashfree)           |   |       AUTOMATED RESOLUTION WORKBOOKS        |\n"
        "| - XCD Input file (Easebuzz)           |   | - Sync_Transactions_YYYY-MM-DD.xlsx         |\n"
        "| - XCD Input file (Airtel)             |   | - Change_Network_YYYY-MM-DD.xlsx            |\n"
        "| - Difference-to-Zero Validator (0.00) |   | - PG Missing Pull Payload / Execution API   |\n"
        "+---------------------------------------+   +---------------------------------------------+"
    )
    add_code_block(doc, arch_ascii)
    
    # SECTION 3
    add_h1("3. The Complete Operational Lifecycle & SOP")
    add_p(
        "This section outlines the chronological Standard Operating Procedure (SOP) followed by Operations and "
        "serves as the blueprint for complete end-to-end automation."
    )
    
    add_h2("Step 1: Daily Ingestion & Initial Reconciliation Run")
    add_p(
        "Every morning at T+1, operations collects the five transaction reports for the preceding day's business. "
        "These files are uploaded to the dashboard (or ingested via automated directory watcher). The engine performs "
        "multi-way matching and identifies three distinct categories of exceptions requiring operational action:"
    )
    add_p(
        "1. Transactions present in Payment Gateway (Cashfree/Easebuzz/Airtel) but completely MISSING in CMS.\n"
        "2. Transactions present in CMS but marked with SMMS Sync Status = false (or absent in SMMS).\n"
        "3. Transactions where the payment network was incorrectly recorded in CMS compared to the Payment Gateway report."
    )
    
    add_h2("Step 2: Resolving Missing CMS Transactions ('The Pull Workflow')")
    add_p(
        "When transactions exist in the Payment Gateway report (e.g., successful collections on static QR) but are absent "
        "in CMS, merchant settlements cannot be calculated until these transactions are pulled into SwinkPay's database."
    )
    
    add_h3("A. Easebuzz & Airtel Terminal Resolution")
    add_p(
        "For Easebuzz and Airtel, each transaction in the gateway report contains the Merchant Terminal ID directly "
        "(e.g., in the virtual account label or merchant identifier field). The pull payload can be directly formed."
    )
    
    add_h3("B. Cashfree Terminal Resolution via Terminal Master File")
    add_p(
        "For Cashfree, the Terminal ID is not explicitly written in a dedicated column. Instead, it is embedded inside "
        "the Order ID string according to the following syntax:"
    )
    add_code_block(
        doc,
        "Order ID Example:  330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline\n"
        "Structure:         [Merchant_Prefix] - [PARTNER_REF_ID] - [Gateway_Unique_Hash]\n"
        "Extracted Code:    4860  <--- This middle integer corresponds to Partner Ref ID / Terminal"
    )
    add_p(
        "How Terminal Mapping Works:\n"
        "1. Operations uploads the Master Terminal Excel file (`Terminal_Master.xlsx`) containing:\n"
        "   - `TERMINAL ID` (e.g. 141001117)\n"
        "   - `MMS TERMINAL ID` (e.g. XKD8M3)\n"
        "   - `Partner Ref ID` / `Partner VPA` (e.g. 4860, or matching UPI alias)\n"
        "2. The engine parses the Order ID, extracts the middle code (e.g. `4860`), and performs an in-memory index "
        "lookup against the uploaded terminal mapping table.\n"
        "3. Once the matching Terminal ID is resolved, the engine generates the Pull Payload."
    )
    
    add_h3("C. Executing the Pull (Without Postman)")
    add_p(
        "Previously, operations had to manually copy each RRN and Terminal ID into Postman to invoke the backend pull API. "
        "With Ops_Auto:\n"
        "1. Operations navigates to the 'Pull Missing PG' tab in the dashboard.\n"
        "2. All missing transactions are listed with their resolved Terminal ID, PG, Amount, and RRN.\n"
        "3. Clicking 'Pull Missing Transactions' iterates through all missing records and sends authenticated POST requests "
        "to the SwinkPay Transaction Ingestion Endpoint:\n"
        "   `POST /api/v1/transactions/pull`\n"
        "   Payload: `{\"pg\": \"cashfree\", \"rrn\": \"626591400167\", \"terminal_id\": \"141001117\", \"order_id\": \"330595-4860-...\"}`\n"
        "4. The backend pulls the transaction from the gateway and inserts it into CMS."
    )
    
    add_h2("Step 3: Synchronizing Transactions to SMMS ('Sync_Transactions.xlsx')")
    add_p(
        "Transactions that exist in CMS but have `SMMS Sync Status = false` (or are listed in `CMS_Not_in_SMMS`) must be "
        "synced so that SMMS can generate merchant settlement batches."
    )
    add_p(
        "What the Engine Does:\n"
        "- Automatically aggregates all such transactions into `Sync_Transactions_YYYY-MM-DD.xlsx`.\n"
        "- Preserves mandatory columns: `SwinkPay Txn ID`, `RRN/UTR`, `Merchant MMS Terminal ID`, `Transaction Amount`, "
        "`Transaction Date & Time`, and `Transaction Status`.\n"
        "\n"
        "Operations Action:\n"
        "1. Download `Sync_Transactions_YYYY-MM-DD.xlsx` from the dashboard badge.\n"
        "2. Log in to the SwinkPay CMS / SMMS Admin Portal.\n"
        "3. Navigate to `Bulk Sync Transactions` -> Upload `Sync_Transactions_YYYY-MM-DD.xlsx`.\n"
        "4. The backend background job processes the file and updates `SMMS Sync Status = true` for all records."
    )
    
    add_h2("Step 4: Updating Payment Networks ('Change_Network.xlsx')")
    add_p(
        "A critical accounting discrepancy occurs when CMS registers an incorrect payment network during transaction registration."
    )
    
    add_h3("A. The Root Cause of Discrepancies")
    add_p(
        "When customers make static QR UPI payments using Rupay Credit Cards on UPI or Wallet/PPI on UPI, CMS sometimes "
        "mistakenly tags them with a card network like `RUPAY` or wallet network `WA`, rather than acknowledging the true "
        "payment mode reported by Cashfree."
    )
    
    add_h3("B. Semantic Equivalence vs Genuine Discrepancies")
    add_p(
        "Crucial Rule: Standard static QR UPI is reported as `UPI` in CMS and `UPI_OFFLINE_STATIC` in Cashfree. "
        "These are SEMANTICALLY IDENTICAL. They must NEVER be flagged as network changes.\n"
        "Only genuine discrepancies must be isolated:"
    )
    
    net_table_data = [
        ["UPI", "UPI_OFFLINE_STATIC", "Standard UPI Offline Static QR", "EQUIVALENT (No action needed)"],
        ["upi_offline_static", "UPI_OFFLINE_STATIC", "Standard UPI Offline Static QR", "EQUIVALENT (No action needed)"],
        ["upi_credit_card_offline_static", "UPI_CREDIT_CARD_OFFLINE_STATIC", "Rupay Credit Card on UPI", "EQUIVALENT (No action needed)"],
        ["RUPAY", "UPI_CREDIT_CARD_OFFLINE_STATIC", "Credit Card misclassified as standard Rupay Card in CMS", "GENUINE DISCREPANCY (77 Txns)"],
        ["WA", "UPI_PPI_OFFLINE_STATIC", "Prepaid Wallet misclassified as WA in CMS", "GENUINE DISCREPANCY (14 Txns)"]
    ]
    create_styled_table(doc, ["CMS Network (Old)", "Cashfree Mode (New)", "Description", "Reconciliation Action"], net_table_data, [1.5, 2.0, 1.8, 1.2])
    
    add_h3("C. Structure & Upload of Change_Network.xlsx")
    add_p(
        "The engine isolates exactly the 91 genuine discrepancies (77 RUPAY + 14 WA) and writes them to `Change_Network_YYYY-MM-DD.xlsx`:\n"
        "- Column A: `SwinkPay Transaction ID`\n"
        "- Column B: `Old Network` (CMS recorded value: `RUPAY` or `WA`)\n"
        "- Column C: `New Network` (PG recorded value: `UPI_CREDIT_CARD_OFFLINE_STATIC` or `UPI_PPI_OFFLINE_STATIC`)\n"
        "- Formatted strictly in Calibri 11pt, borderless, text formatting.\n"
        "\n"
        "Operations Action:\n"
        "1. Download `Change_Network_YYYY-MM-DD.xlsx` from the dashboard action card.\n"
        "2. Log in to the CMS Admin Portal -> `Bulk Network Update`.\n"
        "3. Upload the file. CMS updates the transaction network to match the gateway report."
    )
    
    add_h2("Step 5: Post-Sync Re-Reconciliation & Verification")
    add_p(
        "Once the three operational actions are performed:\n"
        "1. Missing transactions pulled into CMS.\n"
        "2. SMMS sync file uploaded.\n"
        "3. Network change file uploaded.\n"
        "\n"
        "Operations re-downloads the refreshed CMS and SMMS reports from the portal and re-runs reconciliation:\n"
        "- `CMS_Not_in_SMMS` count must be **0**.\n"
        "- `Unmatched_CF` count must be **0**.\n"
        "- Difference between CMS, SMMS, and Payment Gateways must be **0.00**.\n"
        "- The XCD Validator verifies all control checks and displays **ALL CLEAR**."
    )
    
    add_h2("Step 6: Final XCD Generation & Automated Email Dispatch")
    add_p(
        "With reconciliation fully balanced:\n"
        "1. Operations selects the transaction date (or specific weekend date checkboxes: Friday, Saturday, Sunday).\n"
        "2. Clicks 'Download XCD (CashFree / EaseBuzz / Airtel)'. The system generates the official settlement instruction workbooks.\n"
        "3. Clicks 'Copy Formatted Text' on the dashboard executive summary block.\n"
        "4. Pastes the summary into Outlook/Gmail and attaches the generated XCD files."
    )
    
    # SECTION 4
    add_h1("4. Technical Specifications & API Contracts")
    add_p(
        "For engineering teams automating this pipeline into scheduled microservices, cron workers, or Airflow DAGs, "
        "the following API contracts and data models are implemented in Ops_Auto:"
    )
    
    add_h2("A. HTTP REST API Endpoints")
    
    api_endpoints_data = [
        ["POST", "/api/reconcile", "Multipart form data with 5 raw files", "Runs complete reconciliation, generates Excel & audit report"],
        ["POST", "/api/upload-recon-file", "Multipart with completed Reconciliation_*.xlsx", "Extracts matched tabs, validates XCD, outputs dates & counts"],
        ["POST", "/api/upload-terminal-file", "Multipart with Terminal_Master.xlsx", "Saves terminal mappings for Cashfree Order ID resolution"],
        ["GET", "/api/terminal-mappings", "None", "Returns list of active terminal mappings and stats"],
        ["POST", "/api/pull-missing-pg/{session_id}", "JSON options (gateway, dry_run)", "Executes automated API pull for missing PG transactions"],
        ["GET", "/api/download-sync-file/{session_id}", "Session ID", "Downloads Sync_Transactions_YYYY-MM-DD.xlsx"],
        ["GET", "/api/download-network-file/{session_id}", "Session ID", "Downloads Change_Network_YYYY-MM-DD.xlsx"],
        ["GET", "/api/download-xcd/{session_id}/{partner}", "partner, query param ?dates=YYYY-MM-DD", "Downloads date-filtered XCD Input file (.xlsx)"]
    ]
    create_styled_table(doc, ["Method", "Endpoint", "Parameters", "Function & Output"], api_endpoints_data, [0.8, 2.2, 1.8, 1.7])
    
    add_h2("B. Terminal Mapping & Order ID Extraction Algorithm")
    add_code_block(
        doc,
        "# Python Pseudocode for Cashfree Terminal Resolution:\n"
        "import re\n\n"
        "def resolve_cashfree_terminal(order_id: str, terminal_map: dict) -> str:\n"
        "    # Example Order ID: '330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline'\n"
        "    parts = order_id.split('-')\n"
        "    if len(parts) >= 2:\n"
        "        partner_ref_id = parts[1].strip()\n"
        "        # Lookup in Terminal Mapping Table (indexed by Partner Ref ID / MMS Terminal ID)\n"
        "        if partner_ref_id in terminal_map:\n"
        "            return terminal_map[partner_ref_id]['terminal_id']\n"
        "    return 'N/A'\n"
    )
    
    add_h2("C. Network Equivalence & Detection Algorithm")
    add_code_block(
        doc,
        "# Python Logic in core/network_builder.py:\n"
        "def are_networks_equivalent(cms_net: str, cf_mode: str) -> bool:\n"
        "    n_cms = normalize(cms_net)\n"
        "    n_cf = normalize(cf_mode)\n"
        "    if n_cms == n_cf: return True\n"
        "    \n"
        "    std_upi = {'upi', 'upi_offline_static', 'upi_qr', 'static_qr'}\n"
        "    if n_cms in std_upi and n_cf in std_upi: return True\n"
        "    \n"
        "    cc_upi = {'upi_credit_card_offline_static', 'upi_cc'}\n"
        "    if n_cms in cc_upi and n_cf in cc_upi: return True\n"
        "    \n"
        "    ppi_upi = {'upi_ppi_offline_static', 'upi_ppi', 'upi_wallet'}\n"
        "    if n_cms in ppi_upi and n_cf in ppi_upi: return True\n"
        "    \n"
        "    return False  # Flags RUPAY vs UPI_CREDIT_CARD, WA vs UPI_PPI\n"
    )
    
    # SECTION 5
    add_h1("5. How to Run & Automate Locally")
    add_p(
        "Operations analysts or automation engineers can run the application locally or integrate it into a CI/CD service:"
    )
    
    run_methods_data = [
        ["Desktop 1-Click Launcher", "Start_Ops_Auto_Dashboard.bat on Desktop", "Double click to launch server and open browser automatically."],
        ["Project Folder Script", "run_server.bat in project root", "Uses Python virtual environment; runs uvicorn with live reload."],
        ["PowerShell Script", "run_server.ps1 in project root", "Executes with PowerShell execution policies intact."],
        ["Direct CLI Reconcile", "python run_reconciliation.py --input-dir <dir>", "Headless reconciliation without web dashboard for cron jobs."]
    ]
    create_styled_table(doc, ["Execution Method", "File / Command", "Operational Description"], run_methods_data, [1.8, 2.2, 2.5])
    
    # SECTION 6
    add_h1("6. Summary & Sign-Off Checklist")
    add_p(
        "Before finalizing daily operations and dispatching XCD settlement workbooks, confirm the following audit checklist:"
    )
    
    checklist_data = [
        ["1", "Missing PG Transactions Pulled", "All Easebuzz, Airtel, and Cashfree missing records pulled into CMS", "[  ] VERIFIED"],
        ["2", "SMMS Sync File Uploaded", "Sync_Transactions_YYYY-MM-DD.xlsx uploaded to CMS/SMMS portal", "[  ] VERIFIED"],
        ["3", "Network Change File Uploaded", "Change_Network_YYYY-MM-DD.xlsx (91 records) uploaded to CMS portal", "[  ] VERIFIED"],
        ["4", "Re-Reconciliation Run", "Re-run reconciliation on refreshed CMS and SMMS reports", "[  ] VERIFIED"],
        ["5", "Difference to Zero Verified", "CMS vs SMMS vs Gateways difference is exactly 0.00", "[  ] VERIFIED"],
        ["6", "XCD Input Files Downloaded", "Partner workbooks generated for current date or selected weekend dates", "[  ] VERIFIED"],
        ["7", "Executive Email Summary Sent", "Outlet count, CF pulled count, and settlement table copied and dispatched", "[  ] VERIFIED"]
    ]
    create_styled_table(doc, ["#", "Control Item", "Description", "Status"], checklist_data, [0.5, 2.2, 2.8, 1.0])
    
    # Save document
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print(f"[SUCCESS] Word document saved to: {output_path}")


def build_markdown_document(output_path: str):
    """Generates the Markdown version of the complete documentation."""
    content = f"""# {DOC_TITLE}
## {DOC_SUBTITLE}

**Version**: {DOC_VERSION}  
**Date**: {DOC_DATE}  
**Author**: {DOC_AUTHOR}  
**Target Audience**: Operations Analysts, Automation Engineers, FinOps, IT  

---

## 1. Executive Summary & Problem Statement

In digital payment operations, transaction reconciliation across multiple banking and payment gateway (PG) partners is critical to ensure zero revenue leakage, regulatory compliance, and timely merchant settlements.

Historically, SwinkPay operations relied on manual spreadsheet comparisons, disparate VLOOKUP formulas, manual Postman API calls for missing transactions, and cumbersome manual data corrections in CMS.

The **Ops_Auto** automation framework replaces these error-prone manual steps with a high-throughput, deterministic reconciliation pipeline. It guarantees 100% audit accuracy, zero missing data silently ignored, multi-day weekend handling, automated API pulling for missing PG transactions, and targeted Excel generation for CMS/SMMS backend updates.

### Key Capabilities & Operational Impact

| Core Capability | Operational Impact | Implementation Module |
| :--- | :--- | :--- |
| **Audit Balancing Guarantee** | Difference to zero (`0.00`) enforced across all gateways | Guaranteed by `core/xcd_validator.py` |
| **Multi-Day Support** | Date-specific breakdown across weekend batches (Fri, Sat, Sun) | Built-in Date Filter (`core/xcd_builder.py`) |
| **Missing PG Pulling** | Terminal ID resolution & automated API pulling without Postman | `core/terminal_mapper.py` + `core/pg_payload_builder.py` |
| **SMMS Sync Automation** | Auto-detection of un-synced CMS transactions into formatted file | `Sync_Transactions_YYYY-MM-DD.xlsx` (`core/sync_builder.py`) |
| **Network Change File** | Filters false positives; targets only genuine misclassifications (91) | `Change_Network_YYYY-MM-DD.xlsx` (`core/network_builder.py`) |

---

## 2. End-to-End System Architecture

```text
+-----------------------------------------------------------------------------------------+
|                              DAILY INPUT SOURCES (T+1)                                  |
|   [CMS Report]   [SMMS Report]   [Cashfree PG]   [Easebuzz PG]   [Airtel Bank]   [Settlement]   |
+-----------------------------------------------------------------------------------------+
                                             |
                                             v
+-----------------------------------------------------------------------------------------+
|                        HEADER-BASED DETECTION & DATA NORMALIZER                         |
|   - Inspects headers (no guesswork)              - Normalizes 12-digit RRN / UTRs       |
|   - Converts string numbers to exact floats      - Cleans leading zeros & timestamps    |
+-----------------------------------------------------------------------------------------+
                                             |
                                             v
+-----------------------------------------------------------------------------------------+
|                             MULTI-WAY MATCHING CORE ENGINE                              |
|   Step 1: CMS <--> SMMS matching via RRN/UTR & SwinkPay Txn ID                          |
|   Step 2: Partner matching (Cashfree via Ref ID, Easebuzz via UTR, Airtel via Txn ID)    |
|   Step 3: Airtel Settlement report verification (Net amount, UTR, date matching)        |
+-----------------------------------------------------------------------------------------+
                                             |
                    +------------------------+------------------------+
                    |                                                 |
                    v                                                 v
+---------------------------------------+   +---------------------------------------------+
|          MATCHED TRANSACTIONS         |   |             DISCREPANCY ISOLATION           |
| - CMS_SMMS_Matched  - CMS_CF_Matched  |   | 1. Missing in CMS (Unmatched PG)            |
| - CMS_EB_Matched    - CMS_Air_Matched |   | 2. Missing in SMMS (CMS_Not_in_SMMS)        |
| - Multi-Day Date Checkbox Filter      |   | 3. Network Discrepancies (CMS vs PG Mode)   |
+---------------------------------------+   | 4. Failed / Reversed Transactions           |
                    |                       +---------------------------------------------+
                    v                                                 |
+---------------------------------------+                             v
|       OUTPUT WORKBOOKS (XCD)          |   +---------------------------------------------+
| - XCD Input file (Cashfree)           |   |       AUTOMATED RESOLUTION WORKBOOKS        |
| - XCD Input file (Easebuzz)           |   | - Sync_Transactions_YYYY-MM-DD.xlsx         |
| - XCD Input file (Airtel)             |   | - Change_Network_YYYY-MM-DD.xlsx            |
| - Difference-to-Zero Validator (0.00) |   | - PG Missing Pull Payload / Execution API   |
+---------------------------------------+   +---------------------------------------------+
```

---

## 3. The Complete Operational Lifecycle & Standard Operating Procedure (SOP)

This is the definitive chronological SOP followed by Operations and the blueprint for full engineering automation.

```mermaid
flowchart TD
    A["Step 1: Ingest 5 Daily Reports (T+1)"] --> B["Step 2: Run Initial Reconciliation"]
    B --> C{"Any Discrepancies Found?"}
    C -- "Missing in CMS (PG Unmatched)" --> D["Step 3: Pull Missing PG Transactions via Terminal Map"]
    C -- "Missing in SMMS (Sync Status false)" --> E["Step 4: Download & Upload Sync_Transactions.xlsx into CMS/SMMS"]
    C -- "Network Discrepancy (e.g. RUPAY/WA)" --> F["Step 5: Download & Upload Change_Network.xlsx into CMS"]
    D --> G["Step 6: Re-download CMS/SMMS & Re-run Reconciliation"]
    E --> G
    F --> G
    G --> H{"Difference == 0.00 & All Clear?"}
    H -- "Yes" --> I["Step 7: Download Final XCD Workbooks & Send Executive Email"]
    H -- "No" --> C
```

---

### Step 1: Daily Ingestion & Initial Reconciliation Run
- **Timing**: T+1 morning.
- **Inputs**:
  1. CMS Transaction Report (`.xlsx` or `.xls`)
  2. SMMS Transaction Report (`.xlsx`)
  3. Cashfree Settlement / Transaction Report (`.xlsx`)
  4. Easebuzz Transaction Report (`.csv`)
  5. Airtel Bank Transaction Report (`.csv`)
  6. *(Optional)* Airtel Bank Settlement Report (`.csv`)
- **Execution**:
  - Run via Web UI (`http://127.0.0.1:8000`) or CLI:
    ```bash
    python run_reconciliation.py --input-dir sample_inputs --output-dir output
    ```

---

### Step 2: Resolving Missing CMS Transactions ("The Pull Workflow")
When transactions are present in the Payment Gateway reports but missing in CMS:

1. **Terminal ID Resolution**:
   - **Easebuzz & Airtel**: Terminal ID is directly present in the PG report column.
   - **Cashfree**: The Terminal ID is embedded in the `Order Id` string:
     ```text
     Example Order ID: 330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline
     Middle Number:    4860
     ```
     The engine extracts `4860` and matches it against the Master Terminal File:
     - `TERMINAL ID` (e.g. `141001117`)
     - `MMS TERMINAL ID` (e.g. `XKD8M3`)
     - `Partner Ref ID` / `Partner VPA` (e.g. `4860`)

2. **Automated Pull Execution (No Postman Required)**:
   - Navigate to the **Pull Missing PG** tab on the dashboard.
   - Click **Pull Missing Transactions**.
   - The engine sends POST requests to the SwinkPay backend transaction ingestion API:
     ```json
     {{
       "gateway": "cashfree",
       "rrn": "626591400167",
       "terminal_id": "141001117",
       "order_id": "330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline",
       "amount": 36.00
     }}
     ```
   - The backend pulls the transaction from the gateway and records it in CMS.

---

### Step 3: Synchronizing Transactions to SMMS (`Sync_Transactions.xlsx`)
Transactions that exist in CMS but have `SMMS Sync Status = false` (or are listed in `CMS_Not_in_SMMS`):

1. **File Generation**:
   - The dashboard generates `Sync_Transactions_YYYY-MM-DD.xlsx`.
   - Headers: `SwinkPay Txn ID`, `RRN/UTR`, `Merchant MMS Terminal ID`, `Transaction Amount`, `Transaction Date & Time`, `Transaction Status`.
2. **Operations Action**:
   - Download `Sync_Transactions_YYYY-MM-DD.xlsx`.
   - Go to **CMS / SMMS Admin Portal** -> **Bulk Sync Transactions**.
   - Upload the file. The backend job processes the file and updates `SMMS Sync Status = true`.

---

### Step 4: Updating Payment Networks (`Change_Network.xlsx`)
When CMS misclassifies payment networks (e.g., tagging a Credit Card on UPI as `RUPAY` or Wallet as `WA`):

1. **Semantic Equivalence Rules**:
   - `UPI` vs `UPI_OFFLINE_STATIC` = **EQUIVALENT** (Standard Static QR; no change needed).
   - `upi_offline_static` vs `UPI_OFFLINE_STATIC` = **EQUIVALENT**.
   - `upi_credit_card_offline_static` vs `UPI_CREDIT_CARD_OFFLINE_STATIC` = **EQUIVALENT**.
   - `RUPAY` vs `UPI_CREDIT_CARD_OFFLINE_STATIC` = **GENUINE MISCLASSIFICATION (77 Transactions)**.
   - `WA` vs `UPI_PPI_OFFLINE_STATIC` = **GENUINE MISCLASSIFICATION (14 Transactions)**.
   - **Total Genuine Changes**: Exactly **91 transactions**.

2. **File Structure**:
   - `Change_Network_YYYY-MM-DD.xlsx`:
     - Column A: `SwinkPay Transaction ID`
     - Column B: `Old Network` (CMS value: `RUPAY` or `WA`)
     - Column C: `New Network` (PG value: `UPI_CREDIT_CARD_OFFLINE_STATIC` or `UPI_PPI_OFFLINE_STATIC`)
     - Formatted in Calibri 11pt, borderless, text formatting.

3. **Operations Action**:
   - Download `Change_Network_YYYY-MM-DD.xlsx`.
   - Go to **CMS Admin Portal** -> **Bulk Network Update**.
   - Upload the file to update the network fields in CMS.

---

### Step 5: Post-Sync Re-Reconciliation & Verification
After:
1. Missing transactions pulled into CMS.
2. SMMS sync file uploaded.
3. Network change file uploaded.

**Action**:
- Operations downloads refreshed CMS and SMMS reports.
- Re-runs reconciliation.
- Verifies:
  - `CMS_Not_in_SMMS` = 0.
  - `Unmatched_CF` = 0.
  - Difference across all gateways = `0.00`.
  - Dashboard badge displays **ALL CLEAR**.

---

### Step 6: Final XCD Generation & Automated Email Dispatch
- Download partner XCD workbooks (`XCD Input file as on YYYY-MM-DD (Partner).xlsx`).
- For weekend reconciliations, use the multi-day checkboxes to filter by date (Friday, Saturday, Sunday).
- Click **Copy Formatted Text** on the dashboard.
- Paste into Outlook/Gmail and attach the generated XCD files:
  ```text
  Dear Team,

  Please find the attached XCD Recon.

  Exception:
  Transactions happened in 1435 outlets.
  91 CF Transaction is pulled.

  PARTNER     Mode         Network                        SUCCESS COUNT   SUM OF TXN AMOUNT   Sum of net amount
  CashFree    QR: Static   UPI                            904             732047.45           724219.78
  CashFree    QR: Static   upi_offline_static             342             298450.00           295246.33
  CashFree    QR: Static   upi_credit_card_offline_static 28              24190.00            23707.03
  CashFree    QR: Static   upi_ppi_offline_static         3               3500.00             3429.38
  CashFree    Card: Rupay  RUPAY                          77              68420.00            67051.60
  CashFree    Wallet       WA                             14              12500.00            12250.00
  EaseBuzz    QR: Static   UPI                            2310            1840210.50          1818228.00
  Airtel Bank QR: Static   UPI                            2721            2198420.00          2172038.96
  ```

---

## 4. Technical Specifications & API Contracts

For engineering teams automating this pipeline into scheduled microservices, cron workers, or Airflow DAGs:

### REST API Endpoints in Ops_Auto

| Method | Endpoint | Parameters | Function & Output |
| :--- | :--- | :--- | :--- |
| `POST` | `/api/reconcile` | Multipart (5 raw files) | Runs full reconciliation; outputs Excel & JSON audit stats |
| `POST` | `/api/upload-recon-file` | Multipart (`Reconciliation_*.xlsx`) | Instant parse; extracts matched tabs, dates, sync & network items |
| `POST` | `/api/upload-terminal-file` | Multipart (`Terminal_Master.xlsx`) | Ingests and saves master terminal mappings |
| `GET` | `/api/terminal-mappings` | None | Returns list of active terminal mappings |
| `POST` | `/api/pull-missing-pg/{{session_id}}` | JSON `{{gateway, dry_run}}` | Dispatches automated API pull for missing PG records |
| `GET` | `/api/download-sync-file/{{session_id}}` | Session ID | Downloads `Sync_Transactions_YYYY-MM-DD.xlsx` |
| `GET` | `/api/download-network-file/{{session_id}}` | Session ID | Downloads `Change_Network_YYYY-MM-DD.xlsx` |
| `GET` | `/api/download-xcd/{{session_id}}/{{partner}}` | `?dates=YYYY-MM-DD` | Downloads date-filtered XCD Input file (`.xlsx`) |

---

## 5. Local Execution Methods

| Execution Method | File / Command | Description |
| :--- | :--- | :--- |
| **Desktop 1-Click Launcher** | `Start_Ops_Auto_Dashboard.bat` on Desktop | Double click to start server and open browser automatically. |
| **Project Folder Script** | `run_server.bat` in `Ops_Auto/` | Uses Python virtualenv; runs uvicorn with live reload. |
| **PowerShell Script** | `run_server.ps1` in `Ops_Auto/` | Runs uvicorn server in PowerShell. |
| **Direct CLI Reconcile** | `python run_reconciliation.py --input-dir <dir>` | Headless mode suitable for daily cron jobs. |

---

## 6. Daily Sign-Off Audit Checklist

- [ ] **1. Missing PG Transactions Pulled**: Easebuzz, Airtel, Cashfree missing transactions pulled into CMS.
- [ ] **2. SMMS Sync Uploaded**: `Sync_Transactions_YYYY-MM-DD.xlsx` uploaded to CMS/SMMS portal.
- [ ] **3. Network Change Uploaded**: `Change_Network_YYYY-MM-DD.xlsx` (91 rows) uploaded to CMS portal.
- [ ] **4. Re-Reconciliation Executed**: Fresh reconciliation run on refreshed CMS & SMMS reports.
- [ ] **5. Difference to Zero Verified**: CMS vs SMMS vs Gateways difference is strictly `0.00`.
- [ ] **6. XCD Files Generated**: Partner workbooks downloaded for target settlement dates.
- [ ] **7. Executive Email Sent**: Formatted summary text and XCD attachments sent to Finance/Operations.
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[SUCCESS] Markdown document saved to: {output_path}")


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    docx_path = os.path.join(base_dir, "SWINKPAY_RECONCILIATION_AUTOMATION_DOCUMENTATION.docx")
    md_path = os.path.join(base_dir, "SWINKPAY_RECONCILIATION_AUTOMATION_DOCUMENTATION.md")
    
    print("Generating Word Document (.docx)...")
    build_word_document(docx_path)
    
    print("Generating Markdown Document (.md)...")
    build_markdown_document(md_path)
    
    # Also copy docx to Desktop if Desktop exists
    desktop_dir = "C:\\Users\\HP\\OneDrive - SwinkPay Fintech Pvt Ltd\\Desktop"
    if os.path.exists(desktop_dir):
        desktop_docx = os.path.join(desktop_dir, "SWINKPAY_RECONCILIATION_AUTOMATION_DOCUMENTATION.docx")
        shutil.copyfile(docx_path, desktop_docx)
        print(f"[SUCCESS] Copied Word document to Desktop: {desktop_docx}")
