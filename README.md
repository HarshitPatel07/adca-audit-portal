# ADCA Audit Portal

Upload a borrower's raw **Working** folder (bank statements, GSTR-1/3B, insurance policies, stock statements,
Tally debtor/creditor/stock files, MCA master) and download one Excel with three sections:

| Section | Rule |
|---|---|
| **Summation** | Every bank credit classified: debtor receipt / excluded (self, related party, bank, refund, tax) / needs review. Monthly totals tie to the statement. |
| **GST** | Month-wise outward supply, taxable value only (GSTR-1 net of credit notes vs GSTR-3B 3.1(a)), filing dates vs due dates. |
| **Insurance** | Stock-only sum insured per policy, matched to every stock date by policy period; Fire and Burglary separate; first-loss, gaps, under-insurance flagged. |

**Rules only - no AI.** Every figure comes from a fixed rule on the document. A file no rule can read is listed on the
review page as "not read"; a scanned policy's stock SI is typed in the policy table. New bank / insurer formats are added
as new rules in `app.py`.

## Use it (no install)
Open **https://harshitpatel07.github.io/adca-audit-portal/** in Chrome or Edge, pick the Working folder, check the
review page, click Generate Excel. Everything runs inside the browser (Python via Pyodide) - files are never uploaded.
The reading rules are in `docs/core.py`; the page is `docs/index.html` + `docs/portal.js`.

## Run on your PC
```
pip install -r requirements.txt
python app.py
```
Opens http://127.0.0.1:5055 . You can paste a folder path or upload.

## Live on Render
1. Render dashboard -> **New + -> Blueprint** -> select this repo.
2. Set `ADMIN_PASSWORD` when asked.
3. Open the `.onrender.com` URL from any browser and log in.

Uploaded files and the Excel are deleted from the server as soon as the Excel is downloaded (and after 1 hour otherwise).
Free plan sleeps when idle - the first open takes about a minute.
