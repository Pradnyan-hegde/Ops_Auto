# Ops_Auto — Automated Payment Reconciliation Engine

[![Python](https://img.shields.io/badge/Python-3.12-blue.svg)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.110+-green.svg)](https://fastapi.tiangolo.com/)
[![openpyxl](https://img.shields.io/badge/openpyxl-3.1+-orange.svg)](https://openpyxl.readthedocs.io/)
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](LICENSE)

An automated, high-throughput financial transaction reconciliation and settlement engine. Designed to replace manual spreadsheet workflows with deterministic multi-way matching, automated missing transaction recovery via gateway APIs, strict difference-to-zero audit verification, and instant bank settlement file generation.

---

## 🚀 Key Highlights

* **Multi-Way Matching Engine**: Reconciles transactions across Core Merchant Systems (**CMS**), **SMMS**, and multiple Payment Gateways (**Cashfree**, **Easebuzz**, **Airtel Bank**), including secondary settlement verification.
* **Automated Missing Transaction Recovery**: Integrates with Payment Gateway REST APIs to dynamically extract Terminal IDs from Order IDs (e.g., parsing middle tokens from `330595-4860-...`) and pull missing payments directly into CMS without manual Postman calls.
* **Strict Difference-to-Zero Audit Control**: Automatically enforces balance verification across gross, PSP service charges, GST, and net credit amounts before releasing settlement workbooks.
* **Multi-Day Weekend Filtering**: Intelligently detects distinct transaction dates (Friday, Saturday, Sunday) within weekend batches and enables date-checkbox filtering for individual settlement calculations.
* **Operational Discrepancy Workbooks**:
  * `Sync_Transactions_YYYY-MM-DD.xlsx`: Automatically aggregates un-synced CMS transactions for bulk SMMS ingestion.
  * `Change_Network_YYYY-MM-DD.xlsx`: Isolates genuine payment network misclassifications (Rupay Credit Card on UPI & Wallets) while treating standard static QR UPI as equivalent.
* **Interactive Web Dashboard**: Single-page application built with **FastAPI, Vanilla JavaScript (ES6+), and CSS Glassmorphism** for drag-and-drop file ingestion, live status tracking, and automated executive email drafting.

---

## 🏛️ System Architecture

```text
[Daily Ingestion (T+1)]
   • CMS Transaction Report (.xlsx / .xls)
   • SMMS Transaction Report (.xlsx)
   • Cashfree Settlement Report (.xlsx)
   • Easebuzz Transaction Report (.csv)
   • Airtel Bank Transaction Report (.csv)
   • Airtel Bank Settlement Report (.csv)
                │
                ▼
[Header Signature Auto-Detection & Data Normalizer]
   • Dynamic schema identification (no filename reliance)
   • 12-digit RRN / UTR sanitization & amount float conversion
                │
                ▼
[Multi-Way Matching Engine]
   • CMS <──> SMMS (RRN / SwinkPay Txn ID)
   • CMS <──> Cashfree (Bank Reference No / Order ID)
   • CMS <──> Easebuzz (UTR)
   • CMS <──> Airtel (Partner Txn ID)
                │
         ┌──────┴──────────────────────────────┐
         ▼                                     ▼
[Matched Transactions]              [Discrepancy Isolation]
   • CMS_SMMS_Matched                  • Missing in CMS (PG Unmatched) ──► Gateway Pull API
   • CMS_CF_Matched                    • Missing in SMMS (Sync false)   ──► Sync_Transactions.xlsx
   • CMS_EB_Matched                    • Network Misclassifications     ──► Change_Network.xlsx
   • CMS_Air_Matched                   • Failed / Reversed Quarantine
         │
         ▼
[Audit Balancing & Validation (Difference == 0.00)]
         │
         ├─────────────────────────────────────┐
         ▼                                     ▼
[Bank Settlement Workbooks (XCD)]    [Executive Reporting]
   • XCD Input file (Cashfree)          • Executive Email Summary
   • XCD Input file (Easebuzz)          • Active Outlet Counter
   • XCD Input file (Airtel)            • Partner Settlement Table
```

---

## 🛠️ Tech Stack

* **Backend**: Python 3.12, FastAPI, Uvicorn (ASGI)
* **Data Processing**: openpyxl, Regular Expressions (`re`), In-Memory Hash Indexing
* **Frontend**: Vanilla JavaScript (ES6+), HTML5, CSS3 (Glassmorphism & CSS Variables)
* **Testing & Tools**: Python `unittest`, REST APIs, PowerShell & Batch Scripting

---

## 📂 Project Structure

```text
Ops_Auto/
├── core/
│   ├── aggregator.py          # Summary statistics & audit balance checks
│   ├── detector.py            # Header-signature report auto-detection
│   ├── excel_builder.py       # 18-sheet reconciliation workbook generator
│   ├── matcher.py             # Multi-way matching engine across CMS, SMMS & PGs
│   ├── network_builder.py     # Semantic network comparison & Change_Network.xlsx
│   ├── normalizer.py          # RRN, date, amount, and string normalizers
│   ├── pg_payload_builder.py  # Gateway API pull payload constructor
│   ├── reader.py              # Robust Excel & CSV streaming file reader
│   ├── recon_parser.py        # Instant parser for pre-reconciled workbooks
│   ├── sync_builder.py        # Sync_Transactions.xlsx generator for SMMS
│   ├── terminal_mapper.py     # Terminal master mapper & Order ID parser
│   ├── xcd_builder.py         # Partner XCD workbooks & multi-day date filter
│   └── xcd_validator.py       # Strict difference-to-zero audit validator
├── dashboard/
│   ├── server.py              # FastAPI application & REST API routes
│   └── static/
│       ├── index.html         # Single-page operations dashboard UI
│       └── logo.jpg           # Branding asset
├── tests/
│   ├── test_multi_day_and_network.py # Unit tests for date filtering & network rules
│   └── test_reconciliation.py        # Verification against reference datasets
├── run_reconciliation.py      # Standalone CLI reconciliation runner
├── run_server.bat             # 1-click Windows batch launcher
├── run_server.ps1             # Windows PowerShell launcher
├── requirements.txt           # Python package dependencies
└── README.md                  # Project documentation
```

---

## ⚡ Quick Start

### 1. Installation
Clone the repository and install dependencies:
```bash
git clone https://github.com/Pradnyan-hegde/Ops_Auto.git
cd Ops_Auto
pip install -r requirements.txt
```

### 2. Run the Dashboard
Start the local server:
```bash
# Windows 1-Click:
run_server.bat

# Or via Command Line:
uvicorn dashboard.server:app --host 127.0.0.1 --port 8000 --reload
```
Open your browser at `http://127.0.0.1:8000`.

### 3. Headless CLI Mode (Cron / Batch)
Run directly from the terminal without the web interface:
```bash
python run_reconciliation.py --input-dir /path/to/daily_reports --output-dir /path/to/output
```

---

## 🧪 Running Automated Tests

Run the test suite:
```bash
python -m unittest discover tests
```

---

## 📄 License
This project is licensed under the MIT License.
