# SwinkPay Daily Reconciliation & Operations Guide
## Simple End-to-End Guide: Recon, Automated Fixes, XCD Input Files & Invoice Generation

**Version**: Version 2.0 (Simplified)  
**Date**: September 2026  
**Author**: SwinkPay Fintech Operations & Automation Team  

---

## 1. The Big Picture: What Are We Doing & Why?

Every day, thousands of customers scan SwinkPay static QR codes at merchant outlets (like Coffee Day, etc.) and make UPI payments. The money flows through three Payment Gateways (**Cashfree, Easebuzz, Airtel Bank**).

At the end of every day, SwinkPay has three main goals:
1. **Audit & Match**: Ensure every rupee collected by the Payment Gateways is accounted for in SwinkPay's system (Zero Loss).
2. **Merchant Settlement**: Pay each merchant their money (Gross amount collected minus SwinkPay's fees).
3. **Invoicing**: Generate tax invoices for SwinkPay's commission/MDR fees + GST charged to the merchant.

To make this happen, every morning at **T+1**, the Operations team runs **Daily Reconciliation**.

---

## 2. Flow 1: When Everything is "ALL FINE" (The Happy Path)

When all files match perfectly on the morning run (or after fixing any discrepancies), the difference is exactly **0.00**.

```text
[Step 1: Upload Reports] ──> [Step 2: Run Recon] ──> [Difference is 0.00 (ALL CLEAR)]
                                                                │
            ┌───────────────────────────────────────────────────┼─────────────────────────────────┐
            ▼                                                   ▼                                 ▼
[Step 3: Download XCD Input Files]                    [Step 4: Invoice Generation]      [Step 5: Daily Recon Email]
  • XCD Cashfree (.xlsx)                                • CMS calculates exact MDR fees   • Copy summary from dashboard
  • XCD Easebuzz (.xlsx)                                • Generates GST Tax Invoices      • Send to Finance with XCD
  • XCD Airtel (.xlsx)                                    for each merchant outlet          attachments
  └──> Sent to Bank to credit merchants' accounts
```

### What are XCD Input Files?
XCD Input files are the official settlement instruction workbooks. There is one file per payment gateway:
- `XCD Input file as on YYYY-MM-DD (Cf).xlsx`
- `XCD Input file as on YYYY-MM-DD (EB).xlsx`
- `XCD Input file as on YYYY-MM-DD (Airtel).xlsx`

**What happens with them?**  
These files are sent to the nodal bank / settlement account. The bank uses them to credit the **Net Settlement Amount** into each merchant's bank account.

### What is Invoice Generation?
Because reconciliation has verified every transaction, SwinkPay can legally and accurately bill the merchant:
- **Gross Amount Collected** (e.g. ₹100.00)
- **Service Charge / MDR Fee** (e.g. ₹1.65)
- **GST @ 18%** (e.g. ₹0.30)
- **Net Paid to Merchant** (₹98.05)

CMS generates a **Tax Invoice** for the merchant covering the MDR fee and GST.

### Daily Recon Email
Operations clicks **"Copy Formatted Text"** on the Ops_Auto dashboard. It generates a clean executive email with outlet count, Cashfree pulled count, and partner settlement table. Operations pastes this into email, attaches the XCD workbooks, and sends it to Finance.

---

## 3. Flow 2: What If There Are Issues? (The 3 Simple Fixes)

If the initial run does NOT balance to 0.00, it is ALWAYS due to one of three common operational issues. Ops_Auto makes fixing all three simple and automated:

| Issue | What Happened? | How We Fix It in Ops_Auto | Result / Next Step |
| :--- | :--- | :--- | :--- |
| **1. Missing in CMS** | Customer paid on QR, Gateway captured money, but CMS missed it. | Click **"Pull"** on dashboard. Ops_Auto looks up Terminal ID & pulls txn via API. | Backend adds transaction into CMS. |
| **2. Missing in SMMS** | Transaction is in CMS, but SMMS sync status is `false` (un-synced). | Download **`Sync_Transactions.xlsx`** from dashboard badge. | Upload file to CMS/SMMS portal under Bulk Sync. |
| **3. Network Mismatch** | CMS recorded wrong network (e.g. RUPAY/WA instead of UPI Credit/PPI). | Download **`Change_Network.xlsx`** (exactly 91 genuine changes). | Upload file to CMS portal under Bulk Network Update. |

### Deep-Dive on Fix 1: Pulling Missing Transactions (No Postman Required)
When money is in the gateway report but missing in CMS, we must "Pull" it so CMS creates a record:
- **Easebuzz & Airtel**: The Terminal ID is already written in the PG report.
- **Cashfree**: The Terminal ID is embedded in the Order ID (e.g. `330595-4860-AXI...`). The engine copies the middle code (`4860`), searches your uploaded Terminal File, finds the matching Terminal ID, and prepares the pull request.
- Operations clicks **"Pull Missing Transactions"** directly on the dashboard. The system pulls all missing records directly into CMS without opening Postman!

### Deep-Dive on Fix 2: Syncing Transactions to SMMS
When transactions are in CMS but missing in SMMS, SMMS cannot calculate merchant settlements. Ops_Auto generates `Sync_Transactions_YYYY-MM-DD.xlsx`. Operations downloads this file and uploads it into the CMS/SMMS Admin Portal. The backend runs a sync job and marks them synced (`status = true`).

### Deep-Dive on Fix 3: Changing Network (The 91 Transactions)
CMS sometimes mislabels Rupay Credit Cards on UPI as standard card `RUPAY`, or Wallets as `WA`.
- Standard UPI (`UPI` vs `UPI_OFFLINE_STATIC`) is normal static QR and is **NOT** a change.
- The engine isolates **ONLY the 91 real misclassifications**: 77 RUPAY (Credit Card UPI) + 14 WA (Wallet UPI).
- Ops_Auto generates `Change_Network_YYYY-MM-DD.xlsx` (Columns: `SwinkPay Transaction ID`, `Old Network`, `New Network`).
- Operations uploads this file to the CMS Bulk Network Update screen. CMS updates the network field.

---

## 4. The Complete Everyday Cycle (In 1 Picture)

```text
                    Start Day (Morning T+1)
                              │
                              ▼
                    Run Initial Recon
                              │
               Are there any discrepancies?
                              │
           ┌──────────────────┴──────────────────┐
          YES                                   NO
           │                                     │
Fix the 3 things:                                │
1. Click 'Pull' for missing PG transactions      │
2. Upload Sync file to CMS/SMMS                  │
3. Upload Network Change file to CMS             │
           │                                     │
           ▼                                     │
Re-download fresh CMS & SMMS                     │
           │                                     │
           ▼                                     │
   Re-run Recon ─────────────────────────────────┘
                              │
                              ▼
                    Difference is 0.00!
                     (ALL IS CLEAR)
                              │
    ┌─────────────────────────┼─────────────────────────┐
    ▼                         ▼                         ▼
Download XCD            Generate Invoices          Send Email
Input Files             for Merchant Fees          to Finance
    │                         │                         │
    ▼                         ▼                         ▼
Money settled into       Invoices recorded         Daily Recon Closed
merchant bank accounts    in billing system           Successfully!
```

---

## 5. Everyday Operations Checklist

| Step | Action | Description | Status |
| :---: | :--- | :--- | :---: |
| **1** | **Ingest Reports** | Upload CMS, SMMS, Cashfree, Easebuzz, Airtel reports | `[ ]` |
| **2** | **Pull Missing PG** | If any txns in PG but not CMS, click 'Pull' on dashboard | `[ ]` |
| **3** | **Upload Sync File** | Download `Sync_Transactions.xlsx` and upload to CMS/SMMS portal | `[ ]` |
| **4** | **Upload Network File** | Download `Change_Network.xlsx` (91 rows) and upload to CMS portal | `[ ]` |
| **5** | **Re-run Recon** | Re-run recon on refreshed reports -> Verify difference is 0.00 | `[ ]` |
| **6** | **Download XCD Files** | Download partner XCD workbooks for settlement banking | `[ ]` |
| **7** | **Invoice & Email** | Invoices generated in CMS; copy email summary and send to Finance | `[ ]` |
