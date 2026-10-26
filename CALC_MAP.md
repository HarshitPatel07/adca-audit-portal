# Calculation map - every cell group of the 9-sheet workbook

Type **A** = calculated from a digital document by a fixed rule (target 100%, checked by a control).
Type **B** = typed once on the check page (scanned / outside the folder) - never guessed.
Type **C** = firm convention (fixed rule agreed with the firm).

Units: Rs lakh = Rs / 1,00,000 unless stated. "Base date" = DP date; "visit" = visit date.

## 1. Facility sanctioned
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| Facility names, limits | B (A when the sanction letter has a text layer) | Sanction letter | names + amounts in printed order; amounts add up to the printed Total | sum = Total |
| CC outstanding on base / visit date | A | CC bank statement | balance after the last entry on or before the date | running balance ties |
| Other facilities' outstanding | B | loan / BG / LC statements | typed | - |
| DP columns | A | DP calculation sheet | formula link | - |

## 2. DP calculation
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| Stock, creditors, debtors (SS column) | A if the SS is digital, B if scanned | Stock statement | printed SUB-TOTAL(A) / unpaid creditors / book debts | SS total = sum of its categories |
| < 90 / > 90 (SS) | A / B | SS Annexure (ageing) | printed 0-90 / 90+ | < 90 + > 90 = total |
| Stock (books) | A | Tally Stock Summary on the date | Grand Total value | items add up to the Grand Total |
| Creditors (books) | A | Tally Sundry Creditors on the date | credit balances only | parties add up to the Grand Total |
| Debtors (books), < 90 / > 90 | A | Tally Sundry Debtors + ageing file | Dr - Cr; ageing buckets | buckets add up to the total |
| Sister-concern deduction | B (names) + A (amounts) | related-party note (scanned) + ageing | < 90 balance of the named parties | - |
| Margin, paid stock, DP, eligible DP, proportionate DP, excess | A | formulas | margin 25% (C: per sanction), MIN(limit, share) | formula only |
| Remark | C | formula | Regular / Over-utilised within / beyond limit | - |

## 3. Stock / Debtors / Creditors Holding
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| Period rows | C | stock dates | 12 months to the summation end, not before start - 1 month; quarterly if the SS are quarterly | - |
| Stock, debtors, creditors per row | A | Tally per month-end | as in DP | Tally totals |
| Sales per month | A | GSTR-3B 3.1(a)+(b)+(c) | per return period | 3B vs GSTR-1 |
| Purchases per month | A | Tally Purchase Accounts (monthly) | Debit - Credit | - |
| Holding days, averages | A | formulas | stock / cumulative sales x days | formula only |
| Audited-year figures | B | audited balance sheet | typed | - |
| Visit-date stock items | A | Tally Stock Summary on the visit date | six largest items, qty x rate | qty x rate = value |
| Physical verification | B | auditor's count | typed | - |
| Debtors / creditors >= 5% | A | ageing / Tally on the visit date | parties >= 5% of the total | shares add up |
| Key buyers / suppliers | A | invoices + GSTIN | party, PAN = GSTIN 3-12, amount | - |

## 4. Invoice
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| 3 sales + 3 purchase rows | A for digital invoices, B for handwritten | invoice PDFs / GSTR-1 invoice-level | first invoice of each new party in file order | qty x rate = printed amount |
| Taxable value | A | formula | qty x rate | - |

## 5. Summation
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| Debit / credit summation per account | A | bank statements (TXT / Excel / PDF) | sum over the period; credits less own-account and director credits | running balance ties |
| Opening / closing debtors | A | Tally debtors at month-ends | Dr - Cr | Tally totals |
| Sales | A | GSTR-3B 3.1(a)+(b)+(c) | sum over the period | - |
| Limits | B / A | sanction letter | as Facility | - |
| Realisation, proportionate DP | A | formulas | - | formula only |

## 6. Insurance
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| Policy blocks (insurer, period, bank clause, items, location) | A for digital policies, B for scanned SI | policy schedules | fixed label rules per insurer | items add up to the printed total SI |
| Month rows: stock vs SI | A | holding stock + policy in force on the date | newer policy on overlap; side-by-side policies added | - |
| Remark | C | formula | adequately / inadequately insured | - |
| P&M / collateral asset value | B | audited balance sheet / valuation | typed | - |

## 7. GST
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| GSTR-3B list | A | 3B PDFs | due = 20th of next month (or the notified extension); filed = ARN date | - |
| GSTR-1 list | A | GSTR-1 PDFs | due = 11th; filed = ARN date | - |
| GST challans | A | PMT-06 receipts, else bank GST debits | per tax period, summed | ties to 3B cash paid |
| PF / ESIC | A | EPFO / ESIC receipts | payment confirmation date vs 15th | - |
| Electricity | A | bills | bill date + net payable / next bill's payment | - |
| Advance tax | A | challan receipts | vs 15 Jun / Sep / Dec / Mar | - |
| Directors | A | MCA master data | name, DIN | - |

## 8. GST Taxable Summary (master prompt)
| Cells | Type | Source | Rule | Control |
|---|---|---|---|---|
| GSTIN x month taxable value | A | GSTR-1 invoice-level Excel, else filed GSTR-1, else 3B | taxable only; month = document date; CN -, DN +; duplicates once; cancelled out | ties to the filed GSTR-1 per month |

## What 100% means
- Every **A** cell equals the document value (verified answer key), or is red with the failed control.
- Every **B** cell is blank until typed, and is listed on the check page.
- Every **C** cell follows the agreed rule.
