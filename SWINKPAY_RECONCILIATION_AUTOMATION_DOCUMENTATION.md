# SwinkPay Daily Reconciliation & Operations Automation
## Comprehensive Technical Architecture, Automated Workflows, and Standard Operating Procedure (SOP)

**Version**: Version 2.0 (Production)  
**Date**: September 2026  
**Author**: SwinkPay Fintech Operations & Automation Engineering  
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
    B --> CAny Discrepancies Found?
    C -- "Missing in CMS (PG Unmatched)" --> D["Step 3: Pull Missing PG Transactions via Terminal Map"]
    C -- "Missing in SMMS (Sync Status false)" --> E["Step 4: Download & Upload Sync_Transactions.xlsx into CMS/SMMS"]
    C -- "Network Discrepancy (e.g. RUPAY/WA)" --> F["Step 5: Download & Upload Change_Network.xlsx into CMS"]
    D --> G["Step 6: Re-download CMS/SMMS & Re-run Reconciliation"]
    E --> G
    F --> G
    G --> HDifference == 0.00 & All Clear?
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
     {
       "gateway": "cashfree",
       "rrn": "626591400167",
       "terminal_id": "141001117",
       "order_id": "330595-4860-AXIdbfc2143bcf04bd897057bc8f80e4eb7axisupioffline",
       "amount": 36.00
     }
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
| `POST` | `/api/pull-missing-pg/{session_id}` | JSON `{gateway, dry_run}` | Dispatches automated API pull for missing PG records |
| `GET` | `/api/download-sync-file/{session_id}` | Session ID | Downloads `Sync_Transactions_YYYY-MM-DD.xlsx` |
| `GET` | `/api/download-network-file/{session_id}` | Session ID | Downloads `Change_Network_YYYY-MM-DD.xlsx` |
| `GET` | `/api/download-xcd/{session_id}/{partner}` | `?dates=YYYY-MM-DD` | Downloads date-filtered XCD Input file (`.xlsx`) |

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
