# STEP-BY-STEP: How to Run the EUF Monitoring Code

## What you need first
- A laptop with Python 3 installed (Windows/Mac/Linux all fine).
- These 4 files saved together in ONE folder (e.g. `C:\EUF_Tool\`):
  - `end_use_engine.py`   (required by all the others — don't skip this one)
  - `batch_report.py`     (for the whole folder of 700 files)
  - `build_onepager.py`   (for a single loan's one-page summary)
  - `build_excel.py`      (for a single loan's full 10-sheet workbook)

---

## STEP 1 — Install the two Python packages (one-time only)
Open Command Prompt (Windows) or Terminal (Mac), then run:
```
pip install openpyxl pandas --break-system-packages
```
If that errors, try without the flag: `pip install openpyxl pandas`

## STEP 2 — Move into the tool folder
```
cd C:\EUF_Tool
```
(replace with wherever you saved the 4 `.py` files)

## STEP 3 — Run the batch report on your folder of 700 statements
```
python batch_report.py --folder "C:\Users\YourName\Documents\LoanStatements" --out portfolio_report.xlsx
```
- Replace the path in quotes with the actual folder holding your 700 Excel files.
- `--out portfolio_report.xlsx` is just the name of the report it will create — change it if
  you like.
- Wait for it to finish — it prints progress every 25 files, e.g. `processed 50/700`.
- When done, it prints a summary: how many processed OK, how many had errors, how many were
  flagged.

## STEP 4 — Open `portfolio_report.xlsx`
- **Dashboard** tab — overall picture (how many clean, how many flagged, totals).
- **Portfolio Summary** tab — one row per loan; the **Status** column says CLEAN or
  REVIEW REQUIRED (red).
- **All Red Flags** tab — every flagged transaction, across all 700, with the filename it
  came from.
- **Errors** tab — any files it couldn't read, and why (so you know which ones to check
  manually).

## STEP 5 (optional) — More accurate results with a loan master list
If you have the real loan amount/type for each file, make a CSV like this (Excel → Save As →
CSV) and call it `loan_master.csv`:

| filename | loan_amount | loan_type | lender_hint | borrower_name | loan_account_no |
|---|---|---|---|---|---|
| statement_001.xlsx | 1500000 | HL | HDFC | Ramesh Kumar | LAN0001 |
| statement_002.xlsx | 892242 | LAP | SBFC | Sunita Rao | LAN0002 |

`filename` must match your statement file names exactly. Then run:
```
python batch_report.py --folder "C:\Users\YourName\Documents\LoanStatements" --mapping loan_master.csv --out portfolio_report.xlsx
```
You don't need every file listed — files not in the CSV still get auto-detected, and the
report tells you which method was used for each ("Amount Source" column).

## STEP 6 (optional) — Drill into one flagged loan
For any loan marked REVIEW REQUIRED that you want to look at closely, run either:

**One-page summary** (quick look, with full transactions one click away inside the same file):
```
python build_onepager.py --input "C:\...\statement_001.xlsx" --loan-amount 1500000 --loan-type HL --lender-hint HDFC --borrower "Ramesh Kumar" --loan-account-no "LAN0001" --purpose "Home purchase" --out statement_001_summary.xlsx
```

**Full 10-sheet workbook** (deepest detail — Loan Master, Disbursement Register, full Txn
Register, Utilisation Certificate, etc.):
```
python -c "
from build_excel import build_workbook
build_workbook(
    statement_path='C:/.../statement_001.xlsx',
    loan_amount=1500000,
    loan_type='HL',
    lender_hint='HDFC',
    borrower_name='Ramesh Kumar',
    loan_account_no='LAN0001',
    sanction_date=None,
    purpose_declared='Home purchase',
    property_address='Flat 302, ABC Apartments',
    out_path='statement_001_full.xlsx',
    window_days=90,
)
"
```

---

## If something goes wrong
- **"python is not recognized"** → Python isn't installed or not on PATH. Install from
  python.org and re-open Command Prompt.
- **A specific file shows up in the "Errors" tab** → usually means that file doesn't have the
  expected columns (sDate, sAmount, sCategory, etc.) or is corrupted/password-protected. Open
  it manually to check.
- **Numbers look off for a particular loan** → check whether it was "Auto-detected" (Amount
  Source column) — add that file to your `loan_master.csv` mapping for an accurate figure.
