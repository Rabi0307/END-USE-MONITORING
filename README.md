# EUF Monitoring — Post-Disbursement End-Use of Funds Tracker

Tools for tracking whether Loan-Against-Property (LAP) / Home Loan (HL) proceeds were used
for the sanctioned purpose, aligned to RBI Fair Practices Code / KYC Master Direction
expectations on end-use monitoring for secured retail loans.

## 🔗 Live tool (GitHub Pages)
Once Pages is enabled on this repo (Settings → Pages → Deploy from branch → `main` / `root`),
it will be live at:

```
https://<your-username>.github.io/<repo-name>/
```

That page (`index.html`) is a **fully client-side** portfolio scanner — folder upload,
in-browser classification, Excel/Word/PDF export. Nothing is uploaded anywhere; all
processing happens in the visitor's own browser.

## Repo structure
```
index.html          ← the browser tool (this is what GitHub Pages serves)
python/              ← standalone Python scripts (same logic, for local/bulk runs)
  end_use_engine.py         core classification + red-flag engine (imported by the rest)
  bsa_report_reader.py      reads full multi-sheet Bank Statement Analyzer (BSA) reports
  build_excel.py            full 10-sheet workbook for ONE loan
  build_onepager.py         condensed one-page summary for ONE loan
  build_from_bsa_report.py  one-pager auto-filled from a full BSA report
  batch_report.py           scans a WHOLE FOLDER (recursive) → one portfolio report
  run_euf_scan.py           standalone, no-dependencies-on-other-files version of batch_report.py
  flag_categories.py        flags a plain list of bank-statement categories
  requirements.txt
samples/             ← example output files, so you can see what each script produces
STEP_BY_STEP_GUIDE.md
```

## Quick start — Python (bulk/local use)
```bash
cd python
pip install -r requirements.txt

# Scan a whole folder (recursive — handles one-subfolder-per-loan layouts)
python batch_report.py --folder "/path/to/statements" --out portfolio_report.xlsx

# Or: drop run_euf_scan.py directly into the folder of statements and run it there —
# no arguments needed, output is named after the folder itself.
```
Full walkthrough: see `STEP_BY_STEP_GUIDE.md`.

## Quick start — browser tool
Open `index.html` locally (double-click, or via the Pages link above), choose a folder of
statements, click "Process folder", then export to Excel / Word / PDF.

## How the classification works
See the `END_USE_MAP` dictionary at the top of `python/end_use_engine.py` — it's kept as
data, not buried in logic, specifically so you can edit it to match your own sanctioned
end-use policy without touching anything else. The same rules are mirrored in `index.html`
(vanilla JS) and `run_euf_scan.py` (standalone copy) — if you change the rules, update all
three, or just work off `python/end_use_engine.py` and have everything else import it.

**Red-flag rules** (constants at the top of `end_use_engine.py`):
- Cash withdrawal ≥ 5% of loan amount
- Any single debit ≥ 10% of loan amount
- Large outflow within 24 hrs of disbursement (≥15%)
- Gold/investment purchase ≥ ₹1,000
- Any cheque/ECS bounce or dishonour event — flagged unconditionally, regardless of amount

## License
Add a LICENSE file of your choice before making this repo public (MIT is a common default
for internal tooling like this).
