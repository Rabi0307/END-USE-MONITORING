"""
end_use_engine.py
------------------
Reusable engine for Post-Disbursement End-Use of Funds (EUF) monitoring
for secured retail loans -- Loan Against Property (LAP) and Home Loan (HL).

WHAT IT DOES
1. Reads a bank-statement extract (xlsx/csv) of the borrower's operative
   account -- the same format exported by most Account Aggregator / bank
   statement analysis tools (columns: sMode, sAmount, transactionTypes,
   sBalance, sNarration, counter_party, sDate, sCreditOrDebit, sCategory).
2. Auto-detects the disbursement credit (or accepts one you specify).
3. Classifies every post-disbursement transaction into an End-Use Bucket
   and a Risk Level, using rules aligned to RBI's KYC (Fair Practices
   Code) end-use-of-funds monitoring expectations for HL/LAP:
      - Funds must flow to the declared purpose (property purchase /
        construction / business use), not be diverted to speculation,
        capital markets, or unrelated personal spends.
      - Cash withdrawals and unexplained third-party transfers close to
        the disbursement date are classic diversion red flags.
4. Produces a structured summary (category-wise utilisation, % of loan
   utilised, red-flag register) that feeds the Excel workbook.

This module has NO hardcoded numbers for any file other than Book8.xlsx.
Import it and call `run()` on any statement + loan parameters.
"""

import pandas as pd
import numpy as np
from datetime import timedelta

# ---------------------------------------------------------------------------
# 1. END-USE BUCKET MAP
#    Maps the raw bank-statement category to:
#      - Bucket        : the RBI/lender end-use classification
#      - Permitted for LAP / HL : whether this bucket is normally an
#        acceptable end use for that loan type ('Y', 'N', 'Review')
#      - Risk Level     : Low / Medium / High -- drives the red-flag log
# Lenders should tailor this map to their own sanctioned End-Use Policy;
# it is deliberately kept as data (not code) so it can be edited without
# touching the logic below.
# ---------------------------------------------------------------------------
END_USE_MAP = {
    # --- Incoming funds / not applicable to end-use tracking ---
    "Direct Deposit":            {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Cheque Deposit":            {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Cash Deposit":               {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Refunds":                   {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Interest Received":         {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Dividends":                  {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Investment Income":          {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Rent Received":              {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Salary Received":            {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},
    "Subsidy":                    {"Bucket": "Incoming Funds / Not Applicable", "LAP": "N/A",    "HL": "N/A",    "Risk": "Low"},

    # --- New / other borrowing landing on this account (over-leverage signal) ---
    "Loan Disbursed":             {"Bucket": "Other Loan Disbursement (New Debt)", "LAP": "Review", "HL": "Review", "Risk": "High"},

    # --- Debt servicing (existing obligations, separate from the tracked loan) ---
    "EMI and Loans":              {"Bucket": "Debt Servicing (Existing Loans)", "LAP": "Review",  "HL": "Review", "Risk": "Medium"},
    "Emi_Paid":                    {"Bucket": "Debt Servicing (Existing Loans)", "LAP": "Review",  "HL": "Review", "Risk": "Medium"},
    "Home Loan":                   {"Bucket": "Debt Servicing (Existing Loans)", "LAP": "Review",  "HL": "Review", "Risk": "Medium"},
    "Card Payment":                {"Bucket": "Debt Servicing (Existing Loans)", "LAP": "Review",  "HL": "Review", "Risk": "Medium"},

    # --- Bounced / dishonoured payments — always a red flag regardless of amount ---
    "Bounced I/W ECS":             {"Bucket": "Bounced Payment / Dishonour", "LAP": "N", "HL": "N", "Risk": "High"},
    "Bounced I/W ECS Charges":     {"Bucket": "Bounced Payment / Dishonour", "LAP": "N", "HL": "N", "Risk": "High"},
    "Bounced O/W Cheque":          {"Bucket": "Bounced Payment / Dishonour", "LAP": "N", "HL": "N", "Risk": "High"},

    # --- Cash / untraceable outflows ---
    "Cash Withdrawals":          {"Bucket": "Cash Withdrawal",                "LAP": "N",       "HL": "N",       "Risk": "High"},
    "Cheque Withdrawal":           {"Bucket": "Cash Withdrawal",                "LAP": "N",       "HL": "N",       "Risk": "High"},

    # --- Speculative / investment outflows ---
    "Gold":                      {"Bucket": "Investment / Gold Purchase",     "LAP": "N",       "HL": "N",       "Risk": "High"},
    "Investment Expense":          {"Bucket": "Investment / Gold Purchase",     "LAP": "N",       "HL": "N",       "Risk": "High"},
    "Fixed Deposit":               {"Bucket": "Investment / Gold Purchase",     "LAP": "Review",  "HL": "Review",  "Risk": "Medium"},

    # --- Household / personal spend (not a permitted LAP/HL end-use) ---
    "Groceries and Shopping":    {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Medium"},
    "Shopping":                   {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Medium"},
    "Utilities and Bills":       {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Medium"},
    "Dining Restaurant and Entertainment": {"Bucket": "Household / Personal Spend", "LAP": "N", "HL": "N", "Risk": "Medium"},
    "Entertainment":               {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Medium"},
    "Hotel and Travel":            {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Medium"},
    "Medical":                     {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Low"},
    "Insurance Premium":           {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Low"},
    "Rent Paid":                   {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Medium"},
    "Fuel":                      {"Bucket": "Household / Personal Spend",     "LAP": "N",       "HL": "N",       "Risk": "Low"},

    # --- Bank-levied charges ---
    "Bank Fees and Charges":     {"Bucket": "Bank Charges",                   "LAP": "N/A",     "HL": "N/A",     "Risk": "Low"},
    "Bank_fees_and_Charges":     {"Bucket": "Bank Charges",                   "LAP": "N/A",     "HL": "N/A",     "Risk": "Low"},

    # --- Third-party movement, needs a purpose narrative to clear ---
    "External Transfers":        {"Bucket": "Third-Party Transfer (Unverified)", "LAP": "Review", "HL": "Review", "Risk": "Medium"},
    "Transfer to UPI":           {"Bucket": "Third-Party Transfer (Unverified)", "LAP": "Review", "HL": "Review", "Risk": "Medium"},
}
DEFAULT_BUCKET = {"Bucket": "Uncategorised", "LAP": "Review", "HL": "Review", "Risk": "Medium"}

# Bounce/dishonour categories are always red-flagged regardless of amount thresholds --
# a single cheque/ECS bounce is itself a materially adverse credit signal.
BOUNCE_CATEGORIES = {"Bounced I/W ECS", "Bounced I/W ECS Charges", "Bounced O/W Cheque"}


def resolve_category(cat):
    """
    Look up a category in END_USE_MAP. Handles multi-label values some BSA
    exports produce (e.g. 'EMI and Loans,Investment Expense') by matching
    on the first recognised comma-separated component; falls back to
    DEFAULT_BUCKET (flagged 'Review') if nothing matches, so unrecognised
    categories are never silently treated as safe.
    """
    if cat in END_USE_MAP:
        return END_USE_MAP[cat]
    if isinstance(cat, str) and "," in cat:
        for part in cat.split(","):
            part = part.strip()
            if part in END_USE_MAP:
                return END_USE_MAP[part]
    return DEFAULT_BUCKET

# Thresholds used purely for automated flagging -- tune to policy.
CASH_WITHDRAWAL_FLAG_PCT   = 0.05   # single cash withdrawal > 5% of loan amount
LARGE_TXN_FLAG_PCT         = 0.10   # any single debit > 10% of loan amount
SAME_DAY_OUTFLOW_FLAG_PCT  = 0.15   # outflow within 24h of disbursement > 15% of loan
MONITORING_WINDOW_DAYS     = 90     # RBI/lender EUF monitoring window post-disbursement
MIN_FLAG_AMOUNT            = 1000   # ignore trivially small txns (e.g. recurring gold SIPs) when flagging


def load_statement(path: str) -> pd.DataFrame:
    """
    Load a bank-statement export (xlsx or csv) into a standard DataFrame.
    Also works directly on a full multi-sheet Bank Statement Analyzer (BSA)
    report -- e.g. one that also has 'Customer & Account Information',
    'Overall Analysis', 'Fixed Obligations' etc. -- by picking out its
    transaction-level sheet (prefers 'AllAccountXns', falls back to
    'AccountXns', then any sheet with 'transaction' in the name).
    """
    if str(path).lower().endswith(".csv"):
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path, sheet_name=0)
        xl = pd.ExcelFile(path)
        sheet_names_lower = {s.lower(): s for s in xl.sheet_names}
        # Preference order: fullest transaction-level sheet first.
        for candidate in ("allaccountxns", "accountxns"):
            if candidate in sheet_names_lower:
                df = pd.read_excel(path, sheet_name=sheet_names_lower[candidate])
                break
        else:
            for sheet in xl.sheet_names:
                if "bank account" in sheet.lower() or "transaction" in sheet.lower():
                    df = pd.read_excel(path, sheet_name=sheet)
                    break

    df.columns = [c.strip() for c in df.columns]
    # Full BSA reports mark reversed/duplicate rows via is_excluded -- drop them if present.
    if "is_excluded" in df.columns:
        df = df[df["is_excluded"] != True]  # noqa: E712
    df["sDate"] = pd.to_datetime(df["sDate"])
    df["sAmount"] = pd.to_numeric(df["sAmount"], errors="coerce")
    df = df.sort_values("sDate").reset_index(drop=True)
    return df


def detect_disbursement(df: pd.DataFrame, loan_amount: float = None,
                         tolerance: float = 0.02, lender_hint: str = None):
    """
    Find the disbursement credit.
    - If loan_amount given, match the CREDIT closest to that amount (within
      `tolerance` fraction), optionally narrowing by a lender name hint
      (e.g. 'SBFC', 'IDFC') found in sNarration/counter_party.
    - Else, return the single largest CREDIT in the statement.
    Returns the matching row (pandas Series) or None.
    """
    credits = df[df["sCreditOrDebit"].str.upper() == "CREDIT"].copy()
    if lender_hint:
        mask = credits["sNarration"].str.contains(lender_hint, case=False, na=False)
        if mask.any():
            credits = credits[mask]
    if loan_amount:
        credits["_diff"] = (credits["sAmount"] - loan_amount).abs() / loan_amount
        candidate = credits[credits["_diff"] <= tolerance].sort_values("_diff")
        if not candidate.empty:
            return candidate.iloc[0]
    if credits.empty:
        return None
    return credits.sort_values("sAmount", ascending=False).iloc[0]


def classify(df: pd.DataFrame, loan_type: str = "LAP") -> pd.DataFrame:
    """Add Bucket / Permitted / Risk columns based on END_USE_MAP."""
    loan_type = loan_type.upper()
    out = df.copy()

    out["End_Use_Bucket"] = out["sCategory"].apply(lambda c: resolve_category(c)["Bucket"])
    out["Permitted_For_Loan_Type"] = out["sCategory"].apply(lambda c: resolve_category(c).get(loan_type, "Review"))
    out["Base_Risk_Level"] = out["sCategory"].apply(lambda c: resolve_category(c)["Risk"])
    return out


def flag_transactions(df: pd.DataFrame, disb_row, loan_amount: float) -> pd.DataFrame:
    """
    Apply rule-based red flags on top of the base classification.
    Adds: Days_From_Disbursement, Red_Flag (bool), Flag_Reason (str)
    """
    out = df.copy()
    disb_date = disb_row["sDate"] if disb_row is not None else out["sDate"].min()
    out["Days_From_Disbursement"] = (out["sDate"] - disb_date).dt.days

    reasons = []
    flags = []
    for _, r in out.iterrows():
        reason = []
        cat = r["sCategory"]
        # Bounce/dishonour events are flagged unconditionally -- amount and
        # direction (debit/credit) don't matter; the event itself is the signal.
        is_bounce = cat in BOUNCE_CATEGORIES or (
            isinstance(cat, str) and any(b in cat for b in BOUNCE_CATEGORIES))
        if is_bounce:
            reason.append(f"Cheque/ECS bounce or dishonour event ('{cat}') — review immediately")
        if r["sCreditOrDebit"].upper() == "DEBIT" and r["Days_From_Disbursement"] >= 0:
            pct = r["sAmount"] / loan_amount if loan_amount else 0
            if cat == "Cash Withdrawals" and pct >= CASH_WITHDRAWAL_FLAG_PCT:
                reason.append(f"Cash withdrawal = {pct:.1%} of loan amount")
            if pct >= LARGE_TXN_FLAG_PCT:
                reason.append(f"Single debit = {pct:.1%} of loan amount (>{LARGE_TXN_FLAG_PCT:.0%} threshold)")
            if 0 <= r["Days_From_Disbursement"] <= 1 and pct >= SAME_DAY_OUTFLOW_FLAG_PCT:
                reason.append("Large outflow within 24 hrs of disbursement")
            if cat == "Gold" and r["sAmount"] >= MIN_FLAG_AMOUNT:
                reason.append("Funds routed to gold/investment purchase")
        flags.append(bool(reason))
        reasons.append("; ".join(reason))
    out["Red_Flag"] = flags
    out["Flag_Reason"] = reasons
    return out


def summarize(df: pd.DataFrame, disb_row, loan_amount: float,
              window_days: int = MONITORING_WINDOW_DAYS) -> dict:
    """Build the category-wise utilisation summary for the monitoring window."""
    disb_date = disb_row["sDate"] if disb_row is not None else df["sDate"].min()
    window_end = disb_date + timedelta(days=window_days)
    post = df[(df["sDate"] >= disb_date) & (df["sDate"] <= window_end) &
              (df["sCreditOrDebit"].str.upper() == "DEBIT")]

    by_bucket = (post.groupby("End_Use_Bucket")["sAmount"]
                 .agg(["sum", "count"]).reset_index()
                 .rename(columns={"sum": "Total_Amount", "count": "Txn_Count"}))
    if loan_amount:
        by_bucket["Pct_of_Loan_Amount"] = by_bucket["Total_Amount"] / loan_amount
    else:
        by_bucket["Pct_of_Loan_Amount"] = np.nan
    by_bucket = by_bucket.sort_values("Total_Amount", ascending=False)

    total_utilised = post["sAmount"].sum()
    total_flagged = post.loc[post.index.intersection(df[df["Red_Flag"]].index), "sAmount"].sum() \
        if "Red_Flag" in df.columns else np.nan

    return {
        "disbursement_date": disb_date,
        "window_end": window_end,
        "loan_amount": loan_amount,
        "total_utilised_in_window": total_utilised,
        "pct_utilised": (total_utilised / loan_amount) if loan_amount else np.nan,
        "by_bucket": by_bucket,
        "post_disbursement_txns": post,
    }


def run(statement_path: str, loan_amount: float = None, loan_type: str = "LAP",
        lender_hint: str = None, window_days: int = MONITORING_WINDOW_DAYS):
    """
    End-to-end pipeline: load -> detect disbursement -> classify -> flag -> summarize.
    Returns (full_df, disb_row, summary_dict)
    """
    df = load_statement(statement_path)
    disb_row = detect_disbursement(df, loan_amount=loan_amount, lender_hint=lender_hint)
    effective_loan_amount = loan_amount or (disb_row["sAmount"] if disb_row is not None else None)
    df = classify(df, loan_type=loan_type)
    df = flag_transactions(df, disb_row, effective_loan_amount)
    summary = summarize(df, disb_row, effective_loan_amount, window_days=window_days)
    return df, disb_row, summary


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="RBI End-Use-of-Funds analysis for LAP/HL")
    p.add_argument("--input", required=True, help="Path to bank statement xlsx/csv")
    p.add_argument("--loan-amount", type=float, default=None)
    p.add_argument("--loan-type", default="LAP", choices=["LAP", "HL"])
    p.add_argument("--lender-hint", default=None, help="Substring to find disbursement credit, e.g. 'SBFC'")
    p.add_argument("--window-days", type=int, default=MONITORING_WINDOW_DAYS)
    args = p.parse_args()

    df, disb_row, summary = run(args.input, args.loan_amount, args.loan_type,
                                 args.lender_hint, args.window_days)
    print("Disbursement detected:")
    print(disb_row)
    print("\nUtilisation summary (post-disbursement window):")
    print(summary["by_bucket"])
    print(f"\nTotal utilised: {summary['total_utilised_in_window']:.2f} "
          f"({summary['pct_utilised']:.1%} of loan amount)")
    print(f"\nRed-flagged transactions: {df['Red_Flag'].sum()}")
