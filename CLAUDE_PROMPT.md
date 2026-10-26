Make the stock-audit Excel for this borrower: <BORROWER FOLDER>  (it holds Working\ and Draft\ / WORKING\ and DRAFTING\)

You are the reader for what fixed rules cannot read. The rules (docs/core.py, the same engine as the live portal) read every
digital file; you read the scans and fill claude_facts.json; the script joins both and writes the firm's 9-sheet Excel.
Do not ask me questions. No document = blank cell (one line at the end saying which, no questions).

RULES
1. Source = Working\ only. NEVER copy a figure from the firm's draft. The draft is opened only at the end, for scoring.
2. Read the page itself (render it to an image) for every value you enter. Copy figures exactly; never compute or round,
   except unit conversion, and say which in "_sources".
3. Lakh = Rs / 1,00,000. Audited statements: read the unit in the heading ("Rupees in Hundreds" -> lakh = /1000;
   "Rs in Million" -> x10). Check the earlier year's column against the earlier year's own report.
4. Aadhaar / ID cards: never use a number from them.

STEPS
1. python C:\Harshit\adca-audit-portal\claude_run.py scan "<folder>\Working"
   -> what the rules read, the portal prefill (borrower, GSTIN, directors, period) and the files NOT read.
2. For each file not read, render the pages and read them. Fill <folder>\Draft\claude_facts.json:
   form.own / gstin / related   : prefill (related = MCA directors + related parties of the audited accounts)
   form.banks                   : "Bank, DP limit, facility" - the limit the DP is capped at (latest sanction letter;
                                  check its date; sub-limits named as such). A second bank on its own line (pari passu).
   form.fac                     : other facilities "name, limit, O/s on base date, O/s on visit date" - O/s from the bank
                                  balance confirmations of those dates.
   form.annual                  : "stock, sales, debtors, creditors, purchases; same for the year before" (audited, lakh):
                                  inventories, revenue from operations, trade receivables, trade payables, purchases.
   form.extra                   : "dd-mm-yyyy, Visit date dd-mm-yyyy, visit-date book stock (lakh)".
   form.inv                     : the rule-read invoice rows by id ("<file>#<page>"): no, date, party, item, uom, qty, rate,
                                  kind (sales / purchase) - 3 sales + 3 purchases, corrected from the scan.
   form.pol                     : {"<policy no>": {"si": stock SI lakh, "fl": first-loss %}} only where the rules left it blank.
   form.pm / coll               : P&M value and collateral value (lakh) if a document gives them.
   books                        : scanned Tally books of the visit date: date, stock_rs, debtors_rs, creditors_rs (credit
                                  balances only), debtor_parties [[name, Rs]], creditor_parties.
   ai                           : {"<file>": AI_PROMPT json} for scanned stock statements / challans / bills the rules missed
                                  (docs/core.py AI_PROMPT; stock statements: stock_date, stock_lakh, creditors_lakh,
                                  debtors_lakh, debtors_over_lakh).
   Leave from / to empty: the rules set the period (ends at the latest stock statement, books and bank permitting).
3. python claude_run.py make "<folder>\Working" "<folder>\Draft\claude_facts.json" "<folder>\Draft\<BORROWER>_CLAUDE.xlsx" "<folder>\Draft\<draft>.xlsx"
   -> Excel written, recalculated in Excel, scored against the draft (same cell / found anywhere / input figures).
4. Check: no #REF! / #VALUE!; DP columns = SS / books of the base date / visit date; 5 random figures traced to the page.
5. Every remaining gap: draft error (prove it from the source), layout difference, typed-only item, or a rule gap ->
   fix the rule in docs/core.py, re-run the borrower, and re-run Gold Star (must stay at 80% same cell).

REPORT (short): file path; accuracy (cells same / found, inputs same / found); what you read and from which page;
rule gaps fixed; draft errors; one line of blanks.
