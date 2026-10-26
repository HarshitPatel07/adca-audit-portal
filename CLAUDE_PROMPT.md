# ADCA stock-audit Excel — one prompt for any new chat

Copy everything between the lines into a new chat and replace `<BORROWER FOLDER>`.

---

Make the ADCA stock-audit Excel (the firm's 9-sheet workbook: Facility sanctioned, DP calculation, Stock Holding,
Debtors Holding, Creditors Holding, Invoice, summation, Insurance, GST) for this borrower: <BORROWER FOLDER>
(the folder holds Working\ — the raw documents — and usually Draft\ with the firm's own Excel).

Do not ask me questions. Generate from what is in the folder; a figure with no document stays blank, listed once at the end.

## The engine
Rules engine and template: public repo https://github.com/HarshitPatel07/adca-audit-portal (docs/core.py = all rules,
docs/template.xlsx = the firm's blank layout, claude_run.py = the runner). Same engine as the live portal
https://harshitpatel07.github.io/adca-audit-portal/ . The rules read every digital file (bank statements, GSTR-3B/1, Tally
books, stock statements, policies, challans). YOU read what the rules cannot: scans, sanction letters, audited balance
sheets, visit-date books, invoices.

## Which way to work
- **You can run commands (Claude Code / any agent with a shell):** use the engine — steps A.
- **Plain chat with uploaded files (no shell):** read the uploaded scans and give me the check-page entries for the live
  portal — steps B.

## Rules (never break)
1. Source = Working\ only. NEVER take a figure from the firm's draft Excel; open it only at the end, to score.
2. Read every value from the page itself (render scans to images and look). Copy exactly; never estimate, round or compute,
   except unit conversion (say which).
3. Lakh = Rs / 1,00,000. Audited statements: read the unit in the heading ("Rupees in Hundreds" → lakh = ÷1000,
   "Rs in Million" → ×10). Check the earlier year's column against the earlier year's own report.
4. Sanction limits from the LATEST sanction letter (check its date); sub-limits named as sub-limits.
5. Outstanding balances from the bank balance confirmation of that date.
6. Never use Aadhaar / ID-card numbers.

## A. With a shell
1. Get the engine (once): `git clone https://github.com/HarshitPatel07/adca-audit-portal` (or `git pull` if present;
   on the ADCA PC it is C:\Harshit\adca-audit-portal) and `pip install pymupdf openpyxl xlrd`.
   Optional: Tesseract OCR (C:\Program Files\Tesseract-OCR) for scanned statements / invoices; Excel (Windows) for recalculation.
2. `python claude_run.py scan "<folder>\Working"` — shows what the rules read, the prefill (borrower, GSTIN, directors,
   period) and the files NOT read. List every file in Working\ yourself too; do not trust a partial listing.
3. Read each unread file and write `<folder>\Draft\claude_facts.json`:
```json
{
  "_sources": {"banks/fac": "file + page", "annual": "file + page + unit", "extra/books": "...", "inv": "..."},
  "form": {
    "own": "BORROWER NAME", "gstin": "GSTIN", "related": "director 1\ndirector 2 (MCA directors + related parties)",
    "banks": "Axis Bank, <DP limit lakh>, <facility>   (second bank on a new line if pari passu)",
    "fac": "<facility>, <limit>, <O/s on base date>, <O/s on visit date>   (one per line, lakh)",
    "annual": "stock, sales, debtors, creditors, purchases; same five for the year before   (audited, lakh)",
    "extra": "dd-mm-yyyy, Visit date dd-mm-yyyy, <visit-date book stock lakh>",
    "inv": {"<file>#<page>": {"no": "", "date": "dd-mm-yyyy", "party": "", "item": "", "uom": "", "qty": "", "rate": "", "kind": "sales|purchase"}},
    "pol": {"<policy no>": {"si": "<stock SI lakh>", "fl": "<first loss %>"}},
    "pm": "", "coll": "", "from": "", "to": ""
  },
  "books": [{"file": "", "date": "dd-mm-yyyy", "stock_rs": 0, "debtors_rs": 0, "creditors_rs": 0,
             "debtor_parties": [["name", 0]], "creditor_parties": [["name", 0]]}],
  "ai": {"<scanned file name>": {"doc_type": "stock_statement", "entries": [{"page": 1, "stock_date": "dd-mm-yyyy",
          "stock_lakh": 0, "creditors_lakh": 0, "debtors_lakh": 0, "debtors_over_lakh": 0}]}}
}
```
   - `related`: add every person / firm in the tax audit's 40A(2)(b) list (Form 3CD; a screenshot like RPT.PNG) with
     " (40A)" after the name — their credits are excluded from the summation like a director's. Trading sister concerns
     WITHOUT the tag stay in (they are trade receipts).
   - `banks` / `fac`: write "(sub-limit)" in the name of every facility the letter calls a sub-limit — it is shown as a
     negative limit and left out of Total Exposure; the DP's "Limit Utilised" adds the O/s of all working-capital
     facilities (CC, EPC/PCFC, PSC, WCDL, OD), never term loans / GECL / BG / LC / LER.
   - `inv` ids are the rows the rules made from the invoice scans (shown by the runner); give 3 sales + 3 purchases.
   - `books` = scanned Tally books of the visit date (creditors = credit balances only).
   - `ai` = scanned stock statements / challans / bills the rules missed (fields as in AI_PROMPT in docs/core.py).
   - Leave from / to empty: the rules choose the period. Leave pol / pm / coll empty if no document gives them.
4. `python claude_run.py make "<folder>\Working" "<folder>\Draft\claude_facts.json" "<folder>\Draft\<BORROWER>_CLAUDE.xlsx" "<firm's draft .xlsx>"`
   → Excel written, recalculated, and scored (same cell / found anywhere / input figures) if the draft is given.
5. Check: no #REF! / #VALUE!; DP columns = stock statement, books of the base date, visit date; trace 5 random figures to the page.
6. For each remaining difference decide: draft error (prove it from the source), layout difference, figure typed only
   by the auditor, or rule gap. A rule gap → fix docs/core.py, re-run this borrower and Gold Star (must stay ~80% same
   cell), then commit and push (the live portal updates).

## Firm conventions the engine applies (learned from all six original drafts — keep them)
- One bank: no second-bank / proportionate-DP rows in the DP and summation sheets; Eligible DP = MIN(limit, total DP);
  realization % = receipts ÷ sales realization.
- Monthly stock statements: 12 month rows in the holding sheets, two audited-year rows (latest audited year = last 31 March
  at least 6 months before the stock date, and the year before). Period ends at the latest stock statement.
- Variance % = (stock statement − books) ÷ stock statement, on debtors and creditors. No GSTR-2B → "Not provided by the borrower".
- No burglary policy: one "Insured Amt (Fire & Burglary)" column, month table right under the fire policy.
- Inventory details = the field app's "PV Sheet" CSV (item, qty, rate) in PV & Checklist\, else visit-date book items.
- A borrower's own summation workbook in Working\ (SUBTOTAL over a filtered sheet) shows which credits the auditor excluded —
  learn the exclusion from its hidden rows, but take amounts from the bank statement (Baid's had 6 corrupt rows).

## B. Plain chat (no shell)
Read the files I upload and give me, ready to paste into the live portal's check page:
Borrower, GSTIN, related parties, "Banks and sanctioned limits", "Other facilities", "Audited years",
"Visit date / visit stock", invoice corrections (3 sales + 3 purchase), stock SI per policy, P&M and collateral values —
each with the file and page it came from. Then I upload the Working folder on the portal and press Make Excel.

## Report back (short)
File path · accuracy vs the firm's draft (cells same / found, input figures same / found) · what you read and from which
page · rule gaps fixed · draft errors found · one line of blanks (no questions).

---
