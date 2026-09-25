# Ops_Auto: Git & Technical Interview Master Guide
## Manual GitHub Workflow, Architecture Q&A, and Practical Machine Coding Tasks

**Author**: Pradnyan Hegde | SwinkPay Fintech Project  
**Date**: September 2026  

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
    cms_by_rrn = {str(r['rrn']).strip(): r for r in cms_records if r.get('rrn')}
    
    matched = []
    missing_in_cms = []
    amount_mismatch = []
    
    for pg in pg_records:
        rrn = str(pg.get('rrn', '')).strip()
        if rrn in cms_by_rrn:
            cms = cms_by_rrn[rrn]
            if round(float(cms['amount']), 2) == round(float(pg['amount']), 2):
                matched.append({'rrn': rrn, 'cms': cms, 'pg': pg, 'status': 'MATCHED'})
            else:
                amount_mismatch.append({'rrn': rrn, 'diff': float(cms['amount']) - float(pg['amount'])})
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
