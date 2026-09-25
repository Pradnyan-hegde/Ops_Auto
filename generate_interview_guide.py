"""
Script to generate the complete GitHub Workflow & Interview Preparation Guide
in both .docx and .md formats.
"""
import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DOC_TITLE = "Ops_Auto: Git & Technical Interview Master Guide"
DOC_SUBTITLE = "Manual GitHub Workflow, Architecture Q&A, and Practical Machine Coding Tasks"
DOC_AUTHOR = "Pradnyan Hegde | SwinkPay Fintech Project"
DOC_DATE = "September 2026"

COLOR_PRIMARY = RGBColor(30, 58, 138)   # #1e3a8a
COLOR_ACCENT = RGBColor(79, 70, 229)    # #4f46e5
COLOR_TEXT = RGBColor(30, 41, 59)       # #1e293b
COLOR_MUTED = RGBColor(100, 116, 139)   # #64748b
COLOR_WHITE = RGBColor(255, 255, 255)

HEX_HEADER_BG = "1E3A8A"
HEX_ROW_ALT = "F8FAFC"
HEX_BORDER = "CBD5E1"
HEX_CALLOUT_BG = "EEF2FF"      # soft indigo
HEX_CALLOUT_BORDER = "4F46E5"
HEX_CODE_BG = "F1F5F9"


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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


def add_callout(doc, title, text, bg_hex=HEX_CALLOUT_BG, border_hex=HEX_CALLOUT_BORDER):
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
    run_t = p.add_run(f"📌 {title}\n")
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
    p_after.paragraph_format.space_after = Pt(6)


def add_code_box(doc, text):
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
    run = p.add_run(text)
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(15, 23, 42)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)


def create_styled_table(doc, headers, rows_data, col_widths=None):
    table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color=HEX_BORDER)
    
    hdr_cells = table.rows[0].cells
    for i, h_text in enumerate(headers):
        cell = hdr_cells[i]
        set_cell_shading(cell, HEX_HEADER_BG)
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = COLOR_WHITE
        if col_widths and i < len(col_widths):
            cell.width = Inches(col_widths[i])
            
    for row_idx, r_data in enumerate(rows_data, start=1):
        row_cells = table.rows[row_idx].cells
        bg_color = HEX_ROW_ALT if row_idx % 2 == 0 else "FFFFFF"
        for col_idx, cell_value in enumerate(r_data):
            cell = row_cells[col_idx]
            if bg_color != "FFFFFF":
                set_cell_shading(cell, bg_color)
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(str(cell_value))
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            run.font.color.rgb = COLOR_TEXT
            if col_widths and col_idx < len(col_widths):
                cell.width = Inches(col_widths[col_idx])
                
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)
    return table


def build_interview_guide_docx(output_path: str):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = COLOR_TEXT
    
    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run(DOC_TITLE)
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run(DOC_SUBTITLE)
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = COLOR_MUTED
    
    add_callout(
        doc,
        "Guide Overview",
        "This master document serves two critical purposes:\n"
        "1. Complete step-by-step instructions on how to manually manage, commit, and push code to GitHub.\n"
        "2. Comprehensive Technical Interview Preparation: Exact interview questions, architecture answers, "
        "and practical live-coding tasks derived directly from your work on Ops_Auto at SwinkPay Fintech."
    )
    
    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = COLOR_ACCENT
        return p

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_TEXT
        return p

    # SECTION 1: GIT MANUAL PUSH
    add_h1("1. How to Push Code to GitHub Manually (Step-by-Step)")
    add_p(
        "Git works in four distinct stages:\n"
        "1. Working Directory (your project files)\n"
        "2. Staging Area (`git add` - placing files into a shipping container)\n"
        "3. Local Repository (`git commit` - sealing the container with a snapshot version)\n"
        "4. Remote Repository (`git push` - sending the container to GitHub over the internet)"
    )
    
    add_h2("The 6 Universal Steps to Push Any Project from Scratch:")
    
    steps_table = [
        ["Step 1", "Create .gitignore", "Define what NOT to track (passwords, .venv, .xlsx, .csv, logs)", "echo .venv/ >> .gitignore"],
        ["Step 2", "Initialize Git", "Create the local repository tracker in your folder", "git init"],
        ["Step 3", "Set Identity", "Configure your GitHub username and email (one-time)", "git config --global user.name 'Name'"],
        ["Step 4", "Stage & Commit", "Add files to staging and capture snapshot", "git add . && git commit -m 'Initial commit'"],
        ["Step 5", "Rename Branch", "Ensure default branch is set to modern standard 'main'", "git branch -M main"],
        ["Step 6", "Link & Push", "Connect your remote GitHub repo and push code", "git remote add origin <url> && git push -u origin main"]
    ]
    create_styled_table(doc, ["Step", "Action", "Purpose", "Command"], steps_table, [0.8, 1.6, 2.3, 1.8])
    
    add_h2("The Daily 3 Commands (Whenever You Make Future Changes)")
    add_p(
        "After the initial setup, you only need three commands to push updates to GitHub:\n"
        "1. `git add .` (Stage all changed files)\n"
        "2. `git commit -m \"Clear description of what changed\"` (Snapshot the changes)\n"
        "3. `git push` (Upload to GitHub)"
    )
    
    # SECTION 2: BEST PRACTICES
    add_h1("2. Golden Rules & Important Points to Remember")
    add_p(
        "• Data Security First (FinTech Compliance): NEVER push actual merchant financial reports, bank account numbers, "
        "customer VPAs, or secret API keys to a public GitHub repo. Always verify with `git status` that `.xlsx`, `.csv`, "
        "and `.env` files are ignored before committing.\n"
        "• Meaningful Commit Messages: Use imperative action verbs (e.g., 'Add terminal parser', 'Fix 91 network changes', "
        "'Update README') instead of vague words like 'update' or 'done'.\n"
        "• Maintain requirements.txt: Keep dependencies clean with `pip freeze > requirements.txt` so anyone cloning "
        "the project can run `pip install -r requirements.txt` immediately."
    )
    
    # SECTION 3: INTERVIEW QUESTIONS & ANSWERS
    add_h1("3. Technical Interview Questions & Answers (Ops_Auto)")
    add_p(
        "These are the exact high-probability questions interviewers will ask about this project, along with "
        "structured, high-impact answers you can speak with confidence."
    )
    
    add_h2("Q1: 'Can you explain the high-level architecture of your reconciliation project?'")
    add_p(
        "Answer to speak:\n"
        "\"At SwinkPay, merchants accept customer UPI payments on static QR codes through three Payment Gateways: "
        "Cashfree, Easebuzz, and Airtel Bank. Daily reconciliation involves a 3-tier matching process:\n"
        "1. First, matching SwinkPay's database (CMS) with the internal merchant management system (SMMS) on 12-digit RRN "
        "and SwinkPay Transaction ID.\n"
        "2. Second, matching CMS against external Payment Gateway reports (Cashfree, Easebuzz, Airtel) on reference IDs and transaction amounts.\n"
        "3. Third, enforcing a difference-to-zero rule: Gross Amount collected minus Gateway Fees minus GST must strictly equal Net Settlement.\n"
        "Once balanced, the system generates partner XCD settlement workbooks for bank transfers and tax invoices for merchant fees.\""
    )
    
    add_h2("Q2: 'How did you handle transactions that were present in the Gateway report but missing in CMS?'")
    add_p(
        "Answer to speak:\n"
        "\"When a transaction succeeded on the static QR at the gateway but failed to record in CMS due to a webhook network issue, "
        "we built an automated 'Pull' workflow.\n"
        "• For Easebuzz and Airtel, the Merchant Terminal ID is directly present in the report.\n"
        "• For Cashfree, the Terminal ID is embedded in the Order ID string (e.g. '330595-4860-AXI...'). My code extracts "
        "the middle integer ('4860') and cross-references it with an in-memory Terminal Master File to resolve the exact Terminal ID.\n"
        "• The system then dispatches an automated REST API call to ingest the transaction into CMS, completely eliminating manual Postman calls.\""
    )
    
    add_h2("Q3: 'What was the Network Mismatch issue and how did you resolve it?'")
    add_p(
        "Answer to speak:\n"
        "\"In Cashfree, standard static QR payments are called 'UPI_OFFLINE_STATIC', whereas in CMS they were labeled 'UPI'. "
        "A naive string inequality check was falsely flagging 904 normal UPI transactions as discrepancies.\n"
        "I implemented a semantic equivalence function that recognized standard UPI variations as identical. "
        "This isolated only the 91 genuine misclassifications:\n"
        "• 77 transactions where Rupay Credit Cards on UPI were tagged in CMS as regular card 'RUPAY'.\n"
        "• 14 transactions where Wallets on UPI were tagged as 'WA'.\n"
        "The engine generated 'Change_Network.xlsx' containing strictly these 91 records for bulk CMS updates.\""
    )
    
    add_h2("Q4: 'Why did you choose FastAPI over Flask or Django for this project?'")
    add_p(
        "Answer to speak:\n"
        "\"FastAPI was chosen because:\n"
        "1. Performance: It is built on Starlette and Uvicorn, offering high-throughput asynchronous request handling for large file uploads.\n"
        "2. Data Validation: Pydantic schemas handle automatic request validation and type safety.\n"
        "3. Developer Experience: It automatically generates interactive Swagger / OpenAPI documentation at `/docs`.\""
    )
    
    add_h2("Q5: 'How did you handle multi-day weekend reconciliations?'")
    add_p(
        "Answer to speak:\n"
        "\"On Mondays, the reconciliation covers transactions from Friday, Saturday, and Sunday in a single combined batch. "
        "I built a date-detection parser that extracts timestamps, groups transactions by calendar date, and exposes interactive "
        "date checkboxes on the dashboard. Operations can filter any specific date to generate date-filtered XCD settlement workbooks.\""
    )

    # SECTION 4: MACHINE CODING TASKS
    add_h1("4. Practical Coding / Machine Test Tasks")
    add_p("Interviewers frequently ask you to code parts of this logic live on screen:")
    
    add_h2("Task 1: Two-Way Matching Algorithm (O(1) Hash Map)")
    add_p("Prompt: 'Write a function to reconcile CMS records against Payment Gateway records by RRN.'")
    code_t1 = (
        "def reconcile_transactions(cms_records, pg_records):\n"
        "    # Index CMS records by RRN for O(1) instantaneous lookup\n"
        "    cms_by_rrn = {str(r['rrn']).strip(): r for r in cms_records if r.get('rrn')}\n"
        "    \n"
        "    matched = []\n"
        "    missing_in_cms = []\n"
        "    amount_mismatch = []\n"
        "    \n"
        "    for pg in pg_records:\n"
        "        rrn = str(pg.get('rrn', '')).strip()\n"
        "        if rrn in cms_by_rrn:\n"
        "            cms = cms_by_rrn[rrn]\n"
        "            if round(float(cms['amount']), 2) == round(float(pg['amount']), 2):\n"
        "                matched.append({'rrn': rrn, 'cms': cms, 'pg': pg, 'status': 'MATCHED'})\n"
        "            else:\n"
        "                amount_mismatch.append({'rrn': rrn, 'diff': float(cms['amount']) - float(pg['amount'])})\n"
        "        else:\n"
        "            missing_in_cms.append(pg)\n"
        "            \n"
        "    return matched, missing_in_cms, amount_mismatch\n"
    )
    add_code_box(doc, code_t1)
    
    add_h2("Task 2: Order ID Token Extraction (Terminal ID Resolver)")
    add_p("Prompt: 'Given Cashfree Order ID format 330595-4860-AXI..., extract the terminal code.'")
    code_t2 = (
        "def extract_terminal_code(order_id: str) -> str:\n"
        "    if not order_id or not isinstance(order_id, str):\n"
        "        return ''\n"
        "    parts = order_id.split('-')\n"
        "    if len(parts) >= 2:\n"
        "        return parts[1].strip()  # Returns '4860'\n"
        "    return ''\n"
    )
    add_code_box(doc, code_t2)
    
    add_h2("Task 3: Difference-to-Zero Verification (Audit Control)")
    add_p("Prompt: 'Verify Gross minus Commission minus GST equals Net Amount with 1 paisa tolerance.'")
    code_t3 = (
        "def verify_audit_balance(gross: float, commission: float, gst: float, net: float) -> bool:\n"
        "    calculated_net = round(gross - commission - gst, 2)\n"
        "    # 0.01 tolerance handles IEEE-754 floating point rounding\n"
        "    return abs(calculated_net - round(net, 2)) <= 0.01\n"
    )
    add_code_box(doc, code_t3)

    # SECTION 5: KEY METRICS CHEAT SHEET
    add_h1("5. Key Numbers & Metrics to Remember for Interviews")
    
    metrics_table = [
        ["Daily Transaction Volume", "~6,500 transactions per day across all outlets", "Proves production scale"],
        ["Gateways Reconciled", "3 Payment Gateways: Cashfree, Easebuzz, Airtel Bank", "Multi-gateway complexity"],
        ["Audit Balancing", "Difference-to-Zero (0.00 difference guaranteed)", "Financial integrity"],
        ["Network Discrepancies", "Exactly 91 genuine changes (77 Rupay CC + 14 Wallet)", "Data cleansing accuracy"],
        ["Operational Impact", "Reduced daily recon from 3+ hours manual to under 10 seconds", "Tangible business value"]
    ]
    create_styled_table(doc, ["Metric / Aspect", "Real Production Value", "What It Demonstrates"], metrics_table, [1.8, 2.5, 2.2])
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print(f"[SUCCESS] Saved Word doc to: {output_path}")


def build_interview_guide_md(output_path: str):
    content = f"""# {DOC_TITLE}
## {DOC_SUBTITLE}

**Author**: {DOC_AUTHOR}  
**Date**: {DOC_DATE}  

---

## 1. How to Push Code to GitHub Manually (Step-by-Step)

Git works in four distinct stages:
1. **Working Directory**: Your local files.
2. **Staging Area (`git add`)**: Placing files in the shipping container.
3. **Local Commit (`git commit`)**: Sealing the container with a version snapshot.
4. **Remote Push (`git push`)**: Uploading the container to GitHub.

### The 6 Universal Steps:

| Step | Action | Command | Purpose |
| :---: | :--- | :--- | :--- |
| **1** | Create `.gitignore` | `echo .venv/ >> .gitignore` | Exclude sensitive data, sheets, and virtual environments |
| **2** | Initialize Git | `git init` | Starts tracking the current folder |
| **3** | Set Identity | `git config --global user.name "Your Name"`<br>`git config --global user.email "your@email.com"` | Tags your commits with your name |
| **4** | Stage & Commit | `git add .`<br>`git commit -m "Initial commit: Automated Reconciliation Engine"` | Creates your first snapshot |
| **5** | Set Main Branch | `git branch -M main` | Renames default branch to `main` |
| **6** | Link & Push | `git remote add origin https://github.com/Username/Repo.git`<br>`git push -u origin main` | Pushes your code live to GitHub |

### The Daily 3 Commands (Whenever you change code):
```bash
git add .
git commit -m "Describe what changed"
git push
```

---

## 2. Golden Rules & Important Points to Remember

* **Data Security First (FinTech Compliance)**: NEVER push actual merchant financial reports, bank account numbers, customer VPAs, or secret API keys to a public GitHub repo. Always verify with `git status` that `.xlsx`, `.csv`, and `.env` files are ignored before committing.
* **Meaningful Commit Messages**: Use imperative action verbs (e.g., `"Add terminal parser"`, `"Fix 91 network changes"`, `"Update README"`) instead of vague words like `"update"` or `"done"`.
* **Maintain `requirements.txt`**: Keep dependencies clean with `pip freeze > requirements.txt` so anyone cloning the project can run `pip install -r requirements.txt` immediately.

---

## 3. Technical Interview Questions & Answers (Ops_Auto)

### Q1: "Can you explain the high-level architecture of your reconciliation project?"
> **Answer**:  
> *"At SwinkPay, merchants accept customer UPI payments on static QR codes through three Payment Gateways: Cashfree, Easebuzz, and Airtel Bank. Daily reconciliation involves a 3-tier matching process:  
> 1. Matching SwinkPay's database (**CMS**) with the internal merchant management system (**SMMS**) on 12-digit RRN and SwinkPay Transaction ID.  
> 2. Matching CMS against external Payment Gateway reports (**Cashfree, Easebuzz, Airtel**) on reference IDs and transaction amounts.  
> 3. Enforcing a **difference-to-zero** rule: Gross Amount collected minus Gateway Fees minus GST must strictly equal Net Settlement.  
> Once balanced, the system generates partner XCD settlement workbooks for bank transfers and tax invoices for merchant fees."*

---

### Q2: "How did you handle transactions that were present in the Gateway report but missing in CMS?"
> **Answer**:  
> *"When a transaction succeeded on static QR at the gateway but failed to record in CMS due to a webhook network issue, we built an automated 'Pull' workflow:  
> • For Easebuzz and Airtel, the Merchant Terminal ID is directly present in the report.  
> • For Cashfree, the Terminal ID is embedded in the Order ID string (e.g. `330595-4860-AXI...`). My code extracts the middle integer (`4860`) and cross-references it with an in-memory Terminal Master File to resolve the exact Terminal ID.  
> • The system then dispatches an automated REST API call to ingest the transaction into CMS, completely eliminating manual Postman calls."*

---

### Q3: "What was the Network Mismatch issue and how did your code detect only the 91 genuine changes?"
> **Answer**:  
> *"In Cashfree, standard static QR payments are called `UPI_OFFLINE_STATIC`, whereas in CMS they were labeled `UPI`. A naive string inequality check was falsely flagging 904 normal UPI transactions as discrepancies.  
> I implemented a semantic equivalence function that recognized standard UPI variations as identical. This isolated only the **91 genuine misclassifications**:  
> • 77 transactions where Rupay Credit Cards on UPI were tagged in CMS as regular card `RUPAY`.  
> • 14 transactions where Wallets on UPI were tagged as `WA`.  
> The engine generated `Change_Network.xlsx` containing strictly these 91 records for bulk CMS updates."*

---

### Q4: "Why did you choose FastAPI over Flask or Django for this project?"
> **Answer**:  
> *"FastAPI was chosen because:  
> 1. **Performance**: Built on Starlette and Uvicorn, offering high-throughput asynchronous request handling for large file uploads.  
> 2. **Data Validation**: Pydantic schemas handle automatic request validation and type safety.  
> 3. **Automatic Documentation**: Automatically generates interactive Swagger/OpenAPI documentation at `/docs`."*

---

### Q5: "How did you handle multi-day weekend reconciliations?"
> **Answer**:  
> *"On Mondays, the reconciliation covers transactions from Friday, Saturday, and Sunday in a single combined batch. I built a date-detection parser that extracts timestamps, groups transactions by calendar date, and exposes interactive date checkboxes on the dashboard. Operations can filter any specific date to generate date-filtered XCD settlement workbooks."*

---

## 4. Practical Machine Coding Tasks

### Task 1: Two-Way Matching Algorithm ($O(1)$ Hash Map)
```python
def reconcile_transactions(cms_records, pg_records):
    # Index CMS records by RRN for O(1) instantaneous lookup
    cms_by_rrn = {{str(r['rrn']).strip(): r for r in cms_records if r.get('rrn')}}
    
    matched = []
    missing_in_cms = []
    amount_mismatch = []
    
    for pg in pg_records:
        rrn = str(pg.get('rrn', '')).strip()
        if rrn in cms_by_rrn:
            cms = cms_by_rrn[rrn]
            if round(float(cms['amount']), 2) == round(float(pg['amount']), 2):
                matched.append({{'rrn': rrn, 'cms': cms, 'pg': pg, 'status': 'MATCHED'}})
            else:
                amount_mismatch.append({{'rrn': rrn, 'diff': float(cms['amount']) - float(pg['amount'])}})
        else:
            missing_in_cms.append(pg)
            
    return matched, missing_in_cms, amount_mismatch
```

---

### Task 2: Order ID Token Extraction (Terminal ID Resolver)
```python
def extract_terminal_code(order_id: str) -> str:
    if not order_id or not isinstance(order_id, str):
        return ''
    parts = order_id.split('-')
    if len(parts) >= 2:
        return parts[1].strip()  # Returns '4860'
    return ''
```

---

### Task 3: Difference-to-Zero Verification (Audit Control)
```python
def verify_audit_balance(gross: float, commission: float, gst: float, net: float) -> bool:
    calculated_net = round(gross - commission - gst, 2)
    # 0.01 tolerance handles IEEE-754 floating point rounding
    return abs(calculated_net - round(net, 2)) <= 0.01
```

---

## 5. Key Metrics Cheat Sheet for Interviews

| Metric / Aspect | Real Production Value | What It Demonstrates |
| :--- | :--- | :--- |
| **Daily Volume** | ~6,500 transactions per day across all outlets | Production scale |
| **Gateways Reconciled** | 3 Payment Gateways: Cashfree, Easebuzz, Airtel Bank | Multi-gateway complexity |
| **Audit Balancing** | Difference-to-Zero (`0.00` difference guaranteed) | Financial integrity |
| **Network Discrepancies** | Exactly 91 genuine changes (77 Rupay CC + 14 Wallet) | Data cleansing accuracy |
| **Operational Impact** | Reduced daily recon from 3+ hours manual to under 10 seconds | Tangible business value |
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[SUCCESS] Saved Markdown to: {output_path}")


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    docx_path = os.path.join(base_dir, "Ops_Auto_GitHub_and_Interview_Preparation_Guide.docx")
    md_path = os.path.join(base_dir, "Ops_Auto_GitHub_and_Interview_Preparation_Guide.md")
    
    print("Generating Word Document...")
    build_interview_guide_docx(docx_path)
    
    print("Generating Markdown Document...")
    build_interview_guide_md(md_path)
    
    desktop_dir = "C:\\Users\\HP\\OneDrive - SwinkPay Fintech Pvt Ltd\\Desktop"
    if os.path.exists(desktop_dir):
        desktop_docx = os.path.join(desktop_dir, "Ops_Auto_GitHub_and_Interview_Preparation_Guide.docx")
        shutil.copyfile(docx_path, desktop_docx)
        print(f"[SUCCESS] Copied to Desktop: {desktop_docx}")
