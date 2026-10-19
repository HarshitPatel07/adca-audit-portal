r"""AUDIT PORTAL - GST / Insurance / Summation (debtor receipts) Excel from raw documents.

RUN:   local  -> python app.py   (opens http://127.0.0.1:5055, folder path allowed, no login)
       live   -> Render runs "gunicorn app:app" (upload only, ADMIN_PASSWORD login, files deleted after the Excel)
       RULES ONLY - no AI. A file no rule can read is listed as "not read" and its figures are typed on the review page.
Upload every raw file of a borrower (or point to its folder). Files are recognised by CONTENT:
  bank statement  : Axis .txt (Finacle), Axis PDF "Account Statement Report", any .xlsx/.csv with Date/Narration/Debit/Credit
  GSTR-1 / GSTR-3B: GST portal PDFs
  insurance       : policy schedule PDFs (New India Bharat Laghu / Sookshma / Burglary / generic "Sum Insured")
  stock statement : "STOCK STATEMENT AS ON dd-mm-yyyy" PDFs (SUB-TOTAL(A), unpaid creditors)
  book stock      : Tally "Stock Summary" xlsx (Grand Total, period end)
  debtor ledger   : Tally "Sundry Debtors" xlsx -> customer names ; creditors -> supplier names
  MCA master      : MCA.XLSX -> director names
RULES = the three master prompts (debtor receipts / outward supply taxable value / stock vs insured amount).
Nothing is guessed: unknown credits go to "Needs review", missing SI -> "Policy not provided".
"""
import os, re, io, sys, json, glob, math, datetime as dt, collections
import pymupdf, openpyxl

MON = {m: i for i, m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split(), 1)}


def num(s):
    s = str(s).replace(",", "").strip()
    return float(s) if s not in ("", "-") else 0.0


def pdf_text(p):
    try:
        with pymupdf.open(p) as doc:
            return "".join(pg.get_text() for pg in doc)
    except Exception:
        return ""


def xlsx_rows(p, max_rows=None):
    out = []
    try:
        wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        for ws in wb.worksheets:
            for r in ws.iter_rows(max_row=max_rows, values_only=True):
                if any(c is not None for c in r):
                    out.append(list(r))
        wb.close()
    except Exception:
        pass
    return out


def d(s):
    for f in ("%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d-%b-%y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return dt.datetime.strptime(s.strip(), f).date()
        except Exception:
            pass
    return None


# ---------------- bank statements ----------------
def bank_txt(t):
    """Finacle text statement (Axis 'STATEMENT FOR A/C'). Amount = balance difference."""
    acc = re.search(r"A/C\s*:\s*(\d+)", t)
    tx, prev = [], None
    for l in t.splitlines():
        if "OPENING" in l and "BALANCE" in l:
            prev = num(l.split()[-1]); continue
        m = re.match(r"(\d{2}-\d{2}-\d{4})\s+(.*)", l)
        if not m or prev is None:
            continue
        nums = re.findall(r"-?\d+\.\d{2}", l)
        if not nums:
            continue
        bal = num(nums[-1]); amt = round(bal - prev, 2); prev = bal
        tx.append(dict(date=d(m.group(1)), narr=l[13:66].strip() or m.group(2)[:60], amt=amt, bal=bal))
    return (acc.group(1) if acc else "bank"), tx


def bank_pdf(t):
    """Axis 'Account Statement Report' PDF: S.NO|date|value date|particulars...|amount|DR/CR|balance|"""
    acc = re.search(r"Account No\s*:\s*(\d+)", t)
    f = t.replace("\n", "|")
    # balance may print as '-,19,23,01,895.63' or with the minus on its own line: '-|,19,23,01,895.63'
    pat = re.compile(r"\|(\d+)\|(\d{2}[/-]\d{2}[/-]\d{4})\|(\d{2}[/-]\d{2}[/-]\d{4})\|(.*?)\|([\d,]+\.\d\d)\|(DR|CR)\|(-?\|?[\d,]+\.\d\d)\|")
    tx = []
    for m in pat.finditer("|" + f):
        a = num(m.group(5)) * (1 if m.group(6) == "CR" else -1)
        tx.append(dict(date=d(m.group(2)), narr=m.group(4).replace("|", ""), amt=a, bal=num(m.group(7).replace("|", ""))))
    return (acc.group(1) if acc else "bank"), tx


def bal_msg(tx):
    chk, bad = balance_check(tx)
    if not chk:
        return "no running balance to check"
    return "running balance ties on all %d rows" % chk if not bad else "RUNNING BALANCE BREAKS on %d of %d rows - CHECK this file" % (bad, chk)


def balance_check(tx):
    """Running-balance test of one file's rows: previous balance + amount must equal the printed balance.
    Returns (rows checked, rows that do not tie). A misread or missing row always shows up here."""
    chk = bad = 0
    for a, b in zip(tx, tx[1:]):
        if a.get("bal") is None or b.get("bal") is None:
            continue
        chk += 1
        # statements print an overdrawn CC balance either as negative or as a plain number with Dr
        if abs(a["bal"] + b["amt"] - b["bal"]) > 0.5 and abs(a["bal"] - b["amt"] - b["bal"]) > 0.5:
            bad += 1
    return chk, bad


def _cellnum(v):
    """Excel cell -> number: 1,234.00 / '1,234.00 Dr' / '-' / None. Text that is not a number -> None."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = re.sub(r"(?i)\s*(dr|cr)\.?$", "", str(v or "").replace(",", "").strip())
    try:
        return float(s) if s not in ("", "-") else None
    except ValueError:
        return None


def bank_table(rows):
    """Any spreadsheet statement: Date + Narration + either separate Debit / Credit columns, or one Amount column with a
    Dr/Cr column. Returns (account no or None, transactions)."""
    top = " ".join(str(c) for r in rows[:10] for c in r if c)
    acc = re.search(r"(?:Account|A/c)\s*(?:No|Number)\.?\s*:?\s*(\d{9,})", top, re.I)
    for i, r in enumerate(rows[:40]):
        low = [str(c or "").lower().strip() for c in r]
        di = next((j for j, c in enumerate(low) if "date" in c), None)
        ni = next((j for j, c in enumerate(low) if any(k in c for k in ("narration", "particular", "description", "remarks", "details"))), None)
        sg = next((j for j, c in enumerate(low) if re.fullmatch(r"(debit\s*/\s*credit|dr\s*/\s*cr|cr\s*/\s*dr|type|dr/cr)", c)), None)
        am = next((j for j, c in enumerate(low) if c.startswith("amount")), None)
        dr = next((j for j, c in enumerate(low) if j != sg and any(k in c for k in ("debit", "withdrawal"))), None)
        cr = next((j for j, c in enumerate(low) if j != sg and any(k in c for k in ("credit", "deposit"))), None)
        bl = next((j for j, c in enumerate(low) if "balance" in c), None)
        if di is None or ni is None or not ((dr is not None and cr is not None) or (am is not None and sg is not None)):
            continue
        tx = []
        for rr in rows[i + 1:]:
            if len(rr) <= max(x for x in (di, ni, dr, cr, am, sg, bl) if x is not None):
                continue
            dd = rr[di]
            dd = dd.date() if isinstance(dd, dt.datetime) else d(str(dd or ""))
            if not dd:
                continue
            if dr is not None and cr is not None:
                a = (_cellnum(rr[cr]) or 0) - (_cellnum(rr[dr]) or 0)
            else:
                v = _cellnum(rr[am]) or 0
                a = v if str(rr[sg] or "").strip().upper().startswith("C") else -v
            if a:
                tx.append(dict(date=dd, narr=str(rr[ni] or ""), amt=round(a, 2), bal=_cellnum(rr[bl]) if bl is not None else None))
        return (acc.group(1) if acc else None), tx
    return None, []


# ---------------- GST returns ----------------
def gstr3b(t):
    f = t.replace("\n", "|")
    if "GSTR-3B" not in f:
        return None
    g = re.search(r"GSTIN of the supplier\|(\w{15})", f)
    yr = re.search(r"Year\|(\d{4})-(\d{2})\|Period\|(\w+)", f)
    a = re.search(r"\(a\) Outward taxable supplies.*?exempted\)\|([\d.,-]+)", f)
    arn = re.search(r"Date of ARN\|([\d/]+)", f)
    if not (g and yr and a):
        return None
    m = MON[yr.group(3)[:3].lower()]
    y = int(yr.group(1)) + (1 if m < 4 else 0)
    return dict(kind="3B", gstin=g.group(1), ym=(y, m), taxable=num(a.group(1)), arn=arn.group(1) if arn else "")


def gstr1(t):
    f = t.replace("\n", "|")
    if "GSTR-1" not in f:
        return None
    g = re.search(r"GSTIN\|(\w{15})", f)
    yr = re.search(r"Financial year\|(\d{4})-(\d{2})\|Tax period\|(\w+)", f)
    if not (g and yr):
        return None
    def grab(p):
        m = re.search(p, f)
        return num(m.group(1)) if m else 0.0
    b2b = grab(r"4A - .*?B2B Regular\|Total\|\d+\|Invoice\|([\d.,-]+)")
    b2cl = grab(r"5 - Taxable outward inter-state.*?B2CL \(Large\)\|Total\|\d+\|Invoice\|([\d.,-]+)")
    b2cs = grab(r"7- Taxable supplies.*?Total\|\d+\|Net Value\|([\d.,-]+)")
    exp = grab(r"6A . Exports.*?\|Total\|\d+\|Invoice\|([\d.,-]+)")
    sez = grab(r"6B - Supplies made to SEZ.*?\|Total\|\d+\|Invoice\|([\d.,-]+)")
    cdnr = grab(r"9B - Credit/Debit Notes \(Registered\).*?Debit notes - Credit notes\)\|\d+\|Note\|([\d.,-]+)")
    cdnur = grab(r"9B - Credit/Debit Notes \(Unregistered\).*?Debit notes - Credit notes\)\|\d+\|Note\|([\d.,-]+)")
    amd = sum(num(x) for x in re.findall(r"Net differential amount \(Amended - Original\)(?: - Total)?\|([\d.,-]+)", f))
    arn = re.search(r"ARN date\|([\d/]+)", f)
    m = MON[yr.group(3)[:3].lower()]
    y = int(yr.group(1)) + (1 if m < 4 else 0)
    return dict(kind="1", gstin=g.group(1), ym=(y, m), b2b=b2b, b2cl=b2cl, b2cs=b2cs, exp=exp, sez=sez, cdn=cdnr + cdnur, amd=amd,
                taxable=round(b2b + b2cl + b2cs + exp + sez + cdnr + cdnur + amd, 2), arn=arn.group(1) if arn else "")


# ---------------- stock statement / books ----------------
def stock_statement(t):
    f = t.replace("\n", "|")
    m = re.search(r"STOCK STATEMENT AS ON (\d{2}-\d{2}-\d{4})", f, re.I)
    s = re.search(r"SUB-TOTAL\(A\)\|([\d.,]+)", f)
    if not (m and s):
        return None
    c = re.search(r"Unpaid Stocks \(Trade Creditors\)\|([\d.,]+)", f)
    b = re.search(r"Book-debt as per Annexure II\|([\d.,]+)", f)
    def part(lbl):
        x = re.search(lbl + r"\|([\d.,]+)", f)
        return num(x.group(1)) if x else None
    # category split as printed (Rs lakh) - compared with the fire policy's RM / WIP / FG sub-limits
    split = dict(RM=part(r"1\. Total Raw Material"), WIP=part(r"2\. Stock in process"),
                 Stores=part(r"3\. Stores & Spares"), FG=part(r"4\. Finished Goods"))
    return dict(date=d(m.group(1)), stock=num(s.group(1)), creditors=num(c.group(1)) if c else None,
                debtors=num(b.group(1)) if b else None, split=split if any(v is not None for v in split.values()) else None)


def tally_kind(rows):
    head = " ".join(str(c) for r in rows[:15] for c in r if c)
    for k in ("Stock Summary", "Sundry Debtors", "Sundry Creditors"):
        if k in head:
            return k, head
    return None, head


def tally_sales(rows, head):
    """Tally 'Monthly Summary' of a SALES ledger/group (not purchase): month -> credit - debit (sales net of returns/credit notes)."""
    if "Monthly Summary" not in head or not re.search(r"\bSales\b", head) or re.search(r"Purchase", head, re.I):
        return None
    st = re.search(r"1-([A-Za-z]{3})-(\d{2,4}) to", head)
    if not st:
        return None
    y = int(st.group(2)); y = y + 2000 if y < 100 else y; m = MON[st.group(1).lower()]
    out, seen = {}, set()
    names = [dt.date(2000, i, 1).strftime("%B") for i in range(1, 13)]
    for r in rows:
        n = str(r[0] or "").strip()
        if n in names:
            mm = names.index(n) + 1
            while mm != m:                      # months run in order from the period start
                m = m % 12 + 1; y += (m == 1)
            if (y, m) in seen:
                continue
            seen.add((y, m))
            nums = [c for c in r[1:] if isinstance(c, (int, float))]
            dr = num(r[1] or 0) if len(r) > 1 and isinstance(r[1], (int, float)) else 0.0
            cr = num(r[2] or 0) if len(r) > 2 and isinstance(r[2], (int, float)) else 0.0
            if nums:
                out[(y, m)] = round(cr - dr, 2)
    return out or None


def tally_end(head):
    m = re.search(r"to (\d{1,2}-[A-Za-z]{3}-\d{2,4})", head)
    return d(m.group(1)) if m else None


def tally_stock(rows, head):
    for r in rows:
        if r and str(r[0]).strip() == "Grand Total":
            v = [c for c in r[1:] if isinstance(c, (int, float))]
            if v:
                return dict(date=tally_end(head), stock=round(v[-1] / 1e5, 5))
    return None


SKIP = ("particulars", "grand total", "closing", "opening", "debit", "balance", "sundry", "group",
        "cin:", "monthly", "transactions")


def tally_names(rows):
    out = []
    for r in rows:
        n = str(r[0] or "").strip()
        if n and not any(n.lower().startswith(s) for s in SKIP) and any(isinstance(c, (int, float)) for c in r[1:]) \
                and not re.match(r"^[\d.,-]+$", n):
            out.append(n)
    return out


def mca_directors(rows):
    out = []
    for i, r in enumerate(rows):
        low = [str(c or "").lower() for c in r]
        if "name" in low and "designation" in low:
            ni, di = low.index("name"), low.index("designation")
            for rr in rows[i + 1:]:
                if len(rr) > di and rr[di] and "director" in str(rr[di]).lower():
                    out.append(str(rr[ni]).strip())
            break
    return out


# ---------------- insurance ----------------
def policy(t, fname):
    p = _policy(t, fname)
    if p:
        _policy_details(t.replace("\n", "|"), p)
    return p


BANK_NAMES = (r"\b(?:AXIS BANK|STATE BANK OF INDIA|UNION BANK OF INDIA|BANK OF INDIA|BANK OF BARODA|BANK OF MAHARASHTRA|"
              r"CENTRAL BANK OF INDIA|PUNJAB NATIONAL BANK|PUNJAB & SIND BANK|CANARA BANK|INDIAN OVERSEAS BANK|INDIAN BANK|UCO BANK|"
              r"HDFC BANK|ICICI BANK|KOTAK MAHINDRA BANK|IDBI BANK|YES BANK|INDUSIND BANK|FEDERAL BANK|IDFC FIRST BANK|RBL BANK|"
              r"SARASWAT BANK|COSMOS BANK|KARUR VYSYA BANK|SOUTH INDIAN BANK|CITY UNION BANK|DCB BANK|BANDHAN BANK|AU SMALL FINANCE BANK)"
              r"(?:\s*(?:LTD|LIMITED)\.?)?(?:\s*,\s*[A-Z]{3,15})?")


def _policy_details(f, p):
    """Nature of the policy, financier (bank) and insured address - shown in the Insurance sheet blocks."""
    m = re.search(r"Undertaking\)\|([^|]{6,90})", f) or \
        re.search(r"(INDUSTRIAL ALL RISK[^|]{0,40}|BURGLARY[^|]{0,40}INSURANCE[^|]{0,20}|STANDARD FIRE[^|]{0,40}|BHARAT [^|]{0,40})", f, re.I)
    p["nature"] = m.group(1).strip() if m else p["cls"]
    # financier: known bank names after the financier / hypothecation label (several banks are joined)
    banks = []
    for mk in re.finditer(r"Financier|Hypothecation|Bank Details|HP-|Hypo\b|Agreed Bank|Bank Clause|Interest of", f, re.I):
        seg = f[mk.start():mk.start() + 400]
        for bm in re.finditer(BANK_NAMES, seg, re.I):
            nm = re.sub(r"\s+", " ", bm.group(0)).upper().strip(" .,")
            if nm not in banks:
                banks.append(nm)
        if banks:
            break
    p["bank"] = " & ".join(banks)
    loc = None
    for rx in (r"Location (?:Details|address)\s*\|?\s*(?:with Pin Code\s*\|)?\s*:?\s*((?:[^|]{3,120}\|){1,3}?)(?=First Loss|Occupancy|Risk|Ris\||Building|Type|Location|Sum|\d{6}\d{6})",
               r"(?:Insured Address|Risk Location Address|Mailing Address of the Insured\s*:?)\s*\|?\s*:?\s*((?:[^|]{3,120}\|){1,4}?)(?=Policy|Phone|PAN|GSTIN|Bank|Mobile|E-?mail|Are you|Occupancy|Product)",
               r"\|Address\s*\|?\s*:?\s*\|?\s*:?\s*((?:[^|]{2,80}\|){1,5}?)(?=GSTIN|Phone|PAN|Mobile|E-?mail|Address)"):
        for m2 in re.finditer(rx, f, re.I):
            cand = re.sub(r"\s+", " ", m2.group(1).replace("|", " ")).strip(" ,:.")
            if len(cand) >= 15 and re.search(r"\d", cand) and "@" not in cand:
                loc = cand; break
        if loc:
            break
    p["location"] = loc or ""
    if "items" not in p:
        p["items"] = [("Stock", p["stock"])] if p["stock"] is not None else []


def _policy(t, fname):
    f = t.replace("\n", "|")
    if not re.search(r"NEW INDIA ASSURANCE", f[:20000], re.I):      # other insurers: label-value rules below
        return policy_generic(f, fname) if re.search(r"Policy\s*(No|Number)", f, re.I) else None
    if not re.search(r"Policy (No|Number)", f) or "Sum Insured" not in f and "SI" not in f:
        return None
    no = re.search(r"Policy (?:Number|No\.?)\s*\|?:?\s*\|?\s*(\d{10,})", f)
    per = re.search(r"From:?\s*(\d{2}/\d{2}/\d{4}).*?To:?\s*\|?\s*(\d{2}/\d{2}/\d{4})", f)
    if not (no and per):
        return None
    ins = re.search(r"Insured(?:'s)? Name\|?:?\s*\|?:?\s*([^|]+)", f)
    gst = re.search(r"GSTIN/UIN\.?\|?:?\|?\s*:?\s*(\w{15})", f)
    prev = re.search(r"Prev\. Policy no\.?\|?:?\s*\|?:?\s*(\d{10,})?", f)
    insurer = re.search(r"THE ([A-Z ]+ASSURANCE CO\. LTD\.|[A-Z ]+INSURANCE CO(?:MPANY)?\.? LTD\.?)", f)
    p = dict(file=fname, no=no.group(1), frm=per.group(1), to=per.group(2), insured=(ins.group(1).strip(" :") if ins else ""),
             gstin=gst.group(1) if gst else "", prev=(prev.group(1) or "") if prev else "",
             insurer=("THE " + insurer.group(1)) if insurer else "", first_loss=None, stock=None, build="", excluded="", sub=None)
    if "BURGLARY" in f.upper():
        p["cls"] = "Burglary"
        fl = re.search(r"First Loss Percentage\|?:?\s*\|?:?\s*(\d+)\s*%", f)
        p["first_loss"] = int(fl.group(1)) if fl else None
        s = re.search(r"STOCK DETAILS\|Sum Insured\|(\d+)\|([^|]+)\|(\d+)", f)
        if s:
            p["stock"] = int(s.group(3)) / 1e5
            p["build"] = "Stocks in Trade: %s = %s" % (s.group(2).strip(), s.group(3))
        others = re.findall(r"(FURNITURE/FIXTURE/FITTINGS|OFFICE EQUIPMENT|COINS/CURRENCY/CURIOS|GOODS HELD|OTHER ITEM) DETAILS\|Sum Insured\|\d+\|[^|]*\|(\d+)", f)
        p["excluded"] = "; ".join("%s %s" % (a.title(), b) for a, b in others if int(b))
        if p["stock"] is None:
            p["build"] = "Stock SI not in text layer (scanned) - enter on review page"
    else:
        p["cls"] = "Fire"
        h = re.search(r"Finished\|Stocks SI\|(.*?)(?:Ris\|k|Type of Construction|7\.\s)", f)
        if h:
            nums = [int(x) for x in h.group(1).split("|") if re.fullmatch(r"\d+", x.strip() or "x") and int(x) >= 1000]
            blocks = [nums[i:i + 7] for i in range(0, len(nums) - 6, 7)]
            if blocks:
                rm = sum(b[4] for b in blocks); wip = sum(b[5] for b in blocks); fg = sum(b[6] for b in blocks)
                p["stock"] = (rm + wip + fg) / 1e5
                p["sub"] = dict(RM=rm / 1e5, WIP=wip / 1e5, FG=fg / 1e5)   # separate sums insured per stock category
                p["build"] = " + ".join("loc%d: RM %g + WIP %g + FG %g" % (i + 1, b[4] / 1e5, b[5] / 1e5, b[6] / 1e5) for i, b in enumerate(blocks)) \
                    + " = %g lakh (Block Details)" % p["stock"]
                p["excluded"] = "Building %g; FFF %g; P&M %g; Other contents %g (lakh)" % tuple(
                    sum(b[k] for b in blocks) / 1e5 for k in (0, 1, 2, 3))
                tot = lambda k: sum(b[k] for b in blocks) / 1e5
                p["items"] = [("Building", tot(0)), ("Furniture & Fittings", tot(1)), ("Plant and Machinery", tot(2)),
                              ("Stock", p["stock"]), ("Other Contents", tot(3))]
        if p["stock"] is None:
            s = re.search(r"\|Stocks?[^|]{0,60}\|(?:[^|]{0,40}\|)?([\d,]{6,})\|", f)
            if s:
                p["stock"] = num(s.group(1)) / 1e5
                p["build"] = "Line 'Stock' %s (check)" % s.group(1)
            else:
                p["build"] = "Stock SI not found - enter on review page"
    return p


# ---------- other insurers (IFFCO Tokio, ICICI Lombard, Bajaj, generic): label-value rules on the '|' text ----------
INSURERS = [(r"IFFCO[- ]TOKIO", "IFFCO-TOKIO GENERAL INSURANCE CO. LTD"), (r"ICICI Lombard", "ICICI LOMBARD GENERAL INSURANCE CO. LTD"),
            (r"Bajaj (?:General|Allianz)", "BAJAJ GENERAL INSURANCE LTD"), (r"NEW INDIA ASSURANCE", "THE NEW INDIA ASSURANCE CO. LTD."),
            (r"United India Insurance", "UNITED INDIA INSURANCE CO. LTD"), (r"National Insurance", "NATIONAL INSURANCE CO. LTD"),
            (r"Oriental Insurance", "THE ORIENTAL INSURANCE CO. LTD"), (r"HDFC ERGO", "HDFC ERGO GENERAL INSURANCE CO. LTD"),
            (r"Tata AIG", "TATA AIG GENERAL INSURANCE CO. LTD"), (r"SBI General", "SBI GENERAL INSURANCE CO. LTD"),
            (r"Reliance General", "RELIANCE GENERAL INSURANCE CO. LTD"), (r"Future Generali", "FUTURE GENERALI INDIA INSURANCE CO. LTD"),
            (r"Cholamandalam", "CHOLAMANDALAM MS GENERAL INSURANCE CO. LTD"), (r"Royal Sundaram", "ROYAL SUNDARAM GENERAL INSURANCE CO. LTD"),
            (r"Universal Sompo", "UNIVERSAL SOMPO GENERAL INSURANCE CO. LTD"), (r"Liberty General", "LIBERTY GENERAL INSURANCE LTD"),
            (r"Kotak (?:Mahindra )?General", "KOTAK GENERAL INSURANCE CO. LTD"), (r"Go Digit|Digit General", "GO DIGIT GENERAL INSURANCE LTD")]
AMT = r"[`₹]?\s*(?:Rs\.?\s*)?([\d,]{5,}(?:\.\d+)?)"


def _gdate(s):
    """'28/01/2026', '05-JUN-2026', '07-JUN-26', 'Feb 02, 2025' -> 'dd/mm/yyyy'"""
    s = s.strip()
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%d-%b-%Y", "%d-%b-%y", "%b %d, %Y", "%d %b %Y", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(s.title() if "%b" in f else s, f).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return None


DATE = r"(\d{2}[/.-](?:\d{2}|[A-Za-z]{3})[/.-]\d{2,4}|[A-Z][a-z]{2} \d{1,2}, \d{4})"


def _period(f):
    i = f.find("Period of Insurance")
    seg = f[i:i + 400] if i >= 0 else f
    m = re.search(r"From\W{0,6}(?:\d{2}:\d{2}(?::\d{2})?\s*(?:Hours of\s*)?)?" + DATE + r".{0,80}?To\W{0,6}(?:Midnight of\s*)?" + DATE, seg, re.S)
    if not m and i >= 0:
        m = re.search(r"From\W{0,6}(?:\d{2}:\d{2}(?::\d{2})?\s*(?:Hours of\s*)?)?" + DATE + r".{0,80}?To\W{0,6}(?:Midnight of\s*)?" + DATE, f, re.S)
    if not m:
        m = re.search(DATE + r"\s*To\s*" + DATE, f)
    if not m:
        return None, None
    return _gdate(m.group(1)), _gdate(m.group(2))


def _label(f, labels, val=r"([^|]{3,80})"):
    for lb in labels:
        m = re.search(lb + r"\s*\|?\s*:?\s*\|?\s*:?\s*" + val, f, re.I)
        if m:
            return m.group(1).strip(" :|")
    return ""


def policy_generic(f, fname):
    no = _label(f, [r"(?<!Fire )Policy\s*(?:Number|No\.?)"], r"([A-Z0-9][A-Z0-9/-]{5,}\d)")
    if not no or not re.search(r"\d", no):
        m = re.search(r"policy number\s+([A-Z0-9][A-Z0-9/-]{5,}\d)", f, re.I)
        no = m.group(1) if m else ""
    frm, to = _period(f)
    if not (no and frm and to):
        return None
    head = f[:6000].upper()
    if re.search(r"BOILER|MACHINERY BREAKDOWN POLICY|MARINE|MOTOR|HEALTH|GROUP PERSONAL|LIABILITY POLICY", head) and "FIRE" not in head[:3000]:
        cls = "Other"
    elif not re.search(r"stock", f, re.I):
        cls = "Other"
    elif "BURGLARY" in head[:4000]:
        cls = "Burglary"
    else:
        cls = "Fire"
    insurer = next((nm for rx, nm in INSURERS if re.search(rx, f, re.I)), "")
    insured = _label(f, [r"Name of the Insured", r"Insured['’]?s? Name", r"\|Insured(?!'s)"])
    gstin = _label(f, [r"GSTIN Number", r"GSTIN\s*/\s*UIN", r"GSTIN ID"], r"(\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z])")
    for m in ([] if gstin else re.finditer(r"(GSTIN(?: Number| ID|\s*/\s*UIN|/UIN)?)\W{0,8}(\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z])", f)):
        pre = f[max(0, m.start() - 300):m.start()]
        if not re.search(r"Company|Issuing Office|Insurer", pre, re.I):
            gstin = m.group(2); break
    prev = _label(f, [r"Previous Policy (?:No\.?|Number)", r"Prev\. Policy no\.?"], r"([A-Z0-9][A-Z0-9/-]{5,}\d)")
    p = dict(file=fname, no=no, frm=frm, to=to, insured=insured, gstin=gstin, prev=prev, insurer=insurer,
             first_loss=None, stock=None, build="", excluded="", sub=None, cls=cls, iar="INDUSTRIAL ALL RISK" in head)
    fl = re.search(r"First Loss (?:Percentage|%|Basis)\W{0,6}(\d{1,3})\s*%", f, re.I)
    if fl:
        p["first_loss"] = int(fl.group(1))
    if cls == "Other":
        p["build"] = "Not a fire / burglary policy - stock not covered"
        return p
    # 1) burglary item table: Item SI | Item Premium ... description | SI | premium ...
    found = []
    if cls == "Burglary":
        i = f.find("Item SI")
        j = f.find("Base Premium", i) if i >= 0 else -1
        if i >= 0:
            seg = f[i:j if j > 0 else i + 3000]
            amts = [num(x) for x in re.findall(r"\|([\d,]{4,}\.\d\d)\|", seg)]
            sis = amts[0::2]
            if sis:
                found.append(("Item SI (burglary schedule)", sum(sis)))
    # 2) sum-insured table: 'Stocks' label, optional description lines, amount
    if not found:
        for m in re.finditer(r"\|(Stocks?(?:\s*[-:][^|]{0,100})?)\|((?:[^|]{0,130}\|){0,6}?)" + AMT + r"\|", f, re.I):
            lab, desc, v = m.group(1), m.group(2), num(m.group(3))
            # skip a 'stock in process' line and the header row of a column table (row number / 'Total sum insured' before the amount)
            if re.match(r"Stocks? in Process", lab, re.I) or re.search(r"\d{1,3}%|Upto|maximum|Total", desc, re.I) \
                    or re.search(r"\|\d{1,2}\|", "|" + desc):
                continue
            found.append(((lab + " " + desc.replace("|", " ")).strip()[:90], v)); break
        wip = re.search(r"\|Stocks? in Process\|" + AMT + r"\|", f, re.I)
        if found and wip and num(wip.group(1)) > 0:
            found.append(("Stock in Process", num(wip.group(1))))
    if found:
        tot = sum(v for _, v in found)
        p["stock"] = tot / 1e5
        p["build"] = " + ".join("%s %s" % (lab, "{:,.2f}".format(v)) for lab, v in found) + " = %.2f lakh" % (tot / 1e5)
        clause = re.search(r"STOCK SI\s*-\s*(\d{5,})", f)
        if clause and abs(num(clause.group(1)) - tot) > 1:
            p["build"] += " | POLICY WORDING SAYS 'STOCK SI-%s' (%.2f lakh) - CHECK which is right" % (clause.group(1), num(clause.group(1)) / 1e5)
    else:
        p["build"] = "Stock SI not found in this policy text - type it on the review page"
    ex = []
    for lab in ("Building", "Plant (?:&|and) Machinery", "Furniture[^|]{0,40}", "Contents", "Electrical Installations?"):
        m = re.search(r"\|(" + lab + r")[^|]{0,40}\|" + AMT + r"\|", f, re.I)
        if m and num(m.group(2)) > 0:
            ex.append("%s %.2f" % (m.group(1)[:30], num(m.group(2)) / 1e5))
            p.setdefault("items", []).append((m.group(1).strip()[:30], num(m.group(2)) / 1e5))
    p["excluded"] = "; ".join(ex) + (" (lakh)" if ex else "")
    p["items"] = p.get("items", []) + ([("Stock", p["stock"])] if p["stock"] is not None else [])
    return p


# ---------------- book data saved as PDF (Tally reports, company stock tables, company stock statements) ----------------
PDF_AMT = r"-?[\d,]+(?:\.\d+)?"
HEADWORDS = {"particulars", "debit", "credit", "closing balance", "opening balance", "balance", "transactions", "closing", "opening",
             "grand total", "total", "page", "group summary", "sundry debtors", "sundry creditors", "debtors", "creditors", "value",
             "rate", "quantity", "nett", "dr", "cr"}


def _pdf_amount(tok):
    t = tok.strip().replace(" Dr", "").replace(" Cr", "").replace(" dr", "").replace(" cr", "")
    return bool(re.fullmatch(PDF_AMT, t)) and re.search(r"\d", t) is not None


def tally_pdf_kind(t):
    """'Sundry Debtors' / 'Sundry Creditors' / 'Stock' / 'Sales' report saved from Tally as PDF; None otherwise."""
    f = t.replace("\n", "|"); ends = f[:1500] + "|" + f[-1500:]
    if re.search(r"Group Monthly Summary|Monthly Summary", ends) and re.search(r"\|Sales(?: Accounts| Local| A/c)?\|", f) \
            and not re.search(r"Purchase", ends, re.I):
        return "Sales"
    if re.search(r"Stock (?:Group )?Summary", ends):
        return "Stock"
    if re.search(r"\|(?:Sundry )?Debtors\b", ends, re.I) and re.search(r"Group Summary|Closing Balance|Closing", ends, re.I):
        return "Sundry Debtors"
    if re.search(r"\|(?:Sundry )?Creditors\b", ends, re.I) and re.search(r"Group Summary|Closing Balance|Closing", ends, re.I):
        return "Sundry Creditors"
    return None


def tally_pdf_names(t):
    """Party names: a text token directly followed by an amount token."""
    toks = [x.strip() for x in t.split("\n")]
    out = []
    for a, b in zip(toks, toks[1:]):
        if a and re.search(r"[A-Za-z]{2}", a) and not _pdf_amount(a) and _pdf_amount(b) and a.lower() not in HEADWORDS \
                and not re.match(r"(page|cin|pan|e-mail|udyam|gstin|\d+-[a-z]{3}-\d+)", a.lower()):
            out.append(a)
    return out


def tally_pdf_stock(t):
    """Tally Stock (Group) Summary PDF: closing value of the Grand Total row, dated by the period end."""
    f = t.replace("\n", "|")
    # Tally prints 'G r a n d T o t a l'; PDF readers differ on the spacing ('G r a n d  T o t a l')
    m = re.search(r"G\s*r\s*a\s*n\s*d\s*T\s*o\s*t\s*a\s*l\|(" + PDF_AMT + r")\|", f)
    per = re.search(r"\d{1,2}-[A-Za-z]{3}-\d{2,4} to (\d{1,2}-[A-Za-z]{3}-\d{2,4})", f)
    if m and per and d(per.group(1)):
        return dict(date=d(per.group(1)), stock=round(num(m.group(1)) / 1e5, 5))
    return None


def stock_month_table(t):
    """Company table 'Month | ... | Total' with one row per month (e.g. 'Sep-25 | ... | 94,552,504') -> book stock per month-end."""
    f = t.replace("\n", "|")
    if not re.search(r"Month\s*\|", f) or "Total" not in f:
        return []
    out = []
    for m in re.finditer(r"(?<=\|)([A-Z][a-z]{2})-(\d{2})\|((?:[^|]*\|){1,30}?)(?=[A-Z][a-z]{2}-\d{2}\||$)", "|" + f + "|"):
        nums = [num(x) for x in re.findall(PDF_AMT, m.group(3)) if re.search(r"\d", x)]
        if nums and m.group(1).lower() in MON:
            y, mo = 2000 + int(m.group(2)), MON[m.group(1).lower()]
            out.append(dict(date=month_end(y, mo), stock=round(nums[-1] / 1e5, 5)))
    return out


def tally_pdf_sales(t):
    """Tally Group Monthly Summary of Sales Accounts (PDF): month -> net transactions of that month."""
    f = t.replace("\n", "|")
    st = re.search(r"1-([A-Za-z]{3})-(\d{2,4}) to", f)
    if not st:
        return None
    y = int(st.group(2)); y = y + 2000 if y < 100 else y; m = MON[st.group(1).lower()]
    names = [dt.date(2000, i, 1).strftime("%B") for i in range(1, 13)]
    out, prev = {}, None
    for mm in re.finditer(r"(?<=\|)(" + "|".join(names) + r")\|(" + PDF_AMT + r")\s*(Cr|Dr)?(?:\|(" + PDF_AMT + r")\s*(Cr|Dr)?)?(?=\|)", f):
        mi = names.index(mm.group(1)) + 1
        while mi != m:
            m = m % 12 + 1; y += (m == 1)
        closing = num(mm.group(2)) * (1 if (mm.group(3) or "Cr") == "Cr" else -1)
        out[(y, m)] = round(closing - prev, 2) if prev is not None else round(num(mm.group(4)) if mm.group(4) else closing, 2)
        prev = closing
    return out or None


def company_stock_statement(t):
    """Company-format stock statement ('Stock Statement as on dd-mm-yyyy', 'Rs. in Lacs'): stock, unpaid creditors, book debts.
    Text is often OCR'd - '1,,026.87' and similar are cleaned, and the result is marked for checking."""
    f = t.replace("\n", "|")
    m = re.search(r"Stock\s*Statement\s*as\s*on\s*(\d{2}[-/.]\d{2}[-/.]\d{4})", f, re.I)
    if not m:
        return None
    def amt(lbl):
        x = re.search(lbl + r".{0,90}?\|\(?(\d[\d, ]*\.\d{1,2})\)?(?=\|)", f, re.I)
        if not x:
            return None
        v = re.sub(r"[ ,]", "", x.group(1)).strip(".")
        try:
            return float(v)
        except ValueError:
            return None
    stock = amt(r"\|\d?\|?Stock(?! Statement)")
    if stock is None:
        return None
    lakh = bool(re.search(r"L[aA]c|L[aA]kh", f[:3000]))
    return dict(date=d(m.group(1).replace(".", "-").replace("/", "-")), stock=stock if lakh else stock / 1e5,
                creditors=amt(r"Unpaid Stock\s*/?\s*Creditor"), debtors=amt(r"Total Book Debts"), split=None, check=True)


# ---------------- statutory dues: PF (EPFO), ESIC, electricity bills ----------------
def pf_receipt(t):
    """EPFO payment confirmation / TRRN details: wage month, total amount, date of payment."""
    f = t.replace("\n", "|")
    if "Wage Month" not in f or "TRRN" not in f:
        return None
    wm = re.search(r"\|([A-Z]{3}-\d{2,4})\|Wage Month", f) or re.search(r"Wage Month\s*:?\|?\s*([A-Z]{3}-\d{2,4})", f)
    am = re.search(r"\|([\d,]+)\|Total Amount \(Rs\)", f) or re.search(r"Total Amount \(Rs\)\s*:?\|?\s*([\d,]+)", f)
    pd_ = re.search(r"(?:Payment Confirmation Date|Date of Credit|Realization Date)\s*:?\|?:?\|?\s*(\d{2}-[A-Z]{3}-\d{4})", f)
    if not (wm and am):
        return None
    m = re.match(r"([A-Z]{3})-(\d{2,4})", wm.group(1)); y = int(m.group(2)); y = y + 2000 if y < 100 else y
    return dict(kind="PF", ym=(y, MON[m.group(1).lower()]), amount=num(am.group(1)), paid=d(pd_.group(1).title()) if pd_ else None)


def esic_receipt(t):
    """ESIC online challan status: challan period, amount paid, submitted (paid) date."""
    f = t.replace("\n", "|")
    if "ESIC" not in f or "Challan Period" not in f:
        return None
    per = re.search(r"Challan Period:?\|?\s*([A-Za-z]{3})-(\d{4})", f)
    am = re.search(r"Amount Paid:?\|?\s*([\d,]+(?:\.\d+)?)", f)
    sub = re.search(r"Challan Submitted Date\|?\s*(\d{2}-\d{2}-\d{4})", f) or re.search(r"Challan Created Date\|?\s*(\d{2}-\d{2}-\d{4})", f)
    if not (per and am):
        return None
    return dict(kind="ESIC", ym=(int(per.group(2)), MON[per.group(1).lower()]), amount=num(am.group(1)), paid=d(sub.group(1)) if sub else None)


def electricity_bill(t):
    """Electricity (HT) bill: bill month, net payable, due date, and the previous month's payment (date, amount)."""
    f = t.replace("\n", "|")
    bm = re.search(r"Bill month\s*:?\|?\s*([A-Za-z]+)-(\d{4})", f)
    if not bm or not re.search(r"VIDYUT|ELECTRICITY|POWER|Energy Charges", f, re.I):
        return None
    names = [dt.date(2000, i, 1).strftime("%B").lower() for i in range(1, 13)]
    if bm.group(1).lower() not in names:
        return None
    net = re.search(r"NET BILL PAYABLE\|([\d,]+\.\d\d)", f)
    due = re.search(r"Last Dates? Of Payment\|(?:1\. By Online\|)?(\d{2}-[A-Z]{3}-\d{4})", f)
    pay = re.search(r"Last Month Payment Details.*?Amt \(Rs\.\)\|(\d{2}-[A-Z]{3}-\d{4})\|[^|]*\|[^|]*\|([\d,]+(?:\.\d+)?)", f)
    return dict(ym=(int(bm.group(2)), names.index(bm.group(1).lower()) + 1), net=num(net.group(1)) if net else None,
                due=d(due.group(1).title()) if due else None,
                prev_paid=d(pay.group(1).title()) if pay else None, prev_amount=num(pay.group(2)) if pay else None)


def mca_dins(rows):
    """[(DIN/PAN, name)] of directors from the MCA master data sheet."""
    out = []
    for i, r in enumerate(rows):
        low = [str(c or "").lower().strip() for c in r]
        if "name" in low and "designation" in low:
            ni, di = low.index("name"), low.index("designation")
            ki = next((j for j, c in enumerate(low) if c in ("din/pan", "din", "din / pan")), None)
            for rr in rows[i + 1:]:
                if len(rr) > di and rr[di] and "director" in str(rr[di]).lower():
                    out.append((str(rr[ki]).strip() if ki is not None and rr[ki] is not None else "", str(rr[ni]).strip()))
            break
    return out


def tally_group_total(rows, head):
    """Tally Sundry Debtors / Creditors group summary (Excel): Grand Total net (debit - credit), dated by the period end.
    Creditors: credit balances only (debit-balance suppliers are advances, not creditors)."""
    if re.search(r"Transactions", head, re.I):
        for r in rows:
            if r and str(r[0] or "").strip() == "Grand Total":
                v = [c for c in r[1:] if isinstance(c, (int, float))]
                return dict(date=tally_end(head), dr=v[-1], cr=0.0, closing_only=True) if v else None
        return None
    di = ci = None
    for r in rows[:20]:
        low = [str(c or "").strip().lower() for c in r]
        if "debit" in low and "credit" in low:
            di, ci = low.index("debit"), low.index("credit")
    if di is None:
        return None
    for r in rows:
        if r and str(r[0] or "").strip() == "Grand Total":
            val = lambda j: float(r[j]) if j < len(r) and isinstance(r[j], (int, float)) else 0.0
            return dict(date=tally_end(head), dr=val(di), cr=val(ci))
    return None


# ---------------- scan everything ----------------
def scan(paths, log):
    F = dict(bank={}, g1=[], g3=[], ss=[], book=[], debtors=set(), creditors=set(), directors=[], policies=[], names=set(),
             pf=[], esic=[], elec=[], dins=[], deb_tot=[], cred_tot=[])
    for p in paths:
        fn = os.path.basename(p); ext = fn.lower().rsplit(".", 1)[-1] if "." in fn else ""
        try:
            if ext == "pdf":
                t = pdf_text(p)
                for fx, key in ((gstr3b, "g3"), (gstr1, "g1")):
                    r = fx(t)
                    if r:
                        F[key].append(r); log.append(("GSTR-" + r["kind"], fn, "%d-%02d taxable %s" % (r["ym"] + (r["taxable"],)))); break
                else:
                    if "Account Statement Report" in t or "Statement of Axis Bank" in t:
                        a, tx = bank_pdf(t); [x.update(src=fn) for x in tx]; F["bank"].setdefault(a, []).extend(tx); log.append(("Bank PDF", fn, "%d txns; %s" % (len(tx), bal_msg(tx))))
                    elif stock_statement(t):
                        r = stock_statement(t); F["ss"].append(r); log.append(("Stock statement", fn, "%s stock %s" % (r["date"], r["stock"])))
                    elif company_stock_statement(t):
                        r = company_stock_statement(t); F["ss"].append(r)
                        log.append(("Stock statement", fn, "%s stock %s lakh (company format, text looks OCR'd - check)" % (r["date"], r["stock"])))
                    elif tally_pdf_kind(t):
                        k = tally_pdf_kind(t)
                        if k == "Sundry Debtors":
                            n = tally_pdf_names(t); F["debtors"].update(n); log.append(("Debtor ledger", fn, "%d names (PDF)" % len(n)))
                        elif k == "Sundry Creditors":
                            n = tally_pdf_names(t); F["creditors"].update(n); log.append(("Creditor ledger", fn, "%d names (PDF)" % len(n)))
                        elif k == "Stock":
                            r = tally_pdf_stock(t)
                            if r:
                                F["book"].append(r); log.append(("Book stock", fn, "%s %s lakh (PDF)" % (r["date"], r["stock"])))
                            else:
                                log.append(("skipped", fn, "stock summary without a grand total"))
                        else:
                            r = tally_pdf_sales(t) or {}; F.setdefault("sales", {}).update(r)
                            log.append(("Books sales", fn, "%d months (PDF)" % len(r)))
                    elif stock_month_table(t):
                        r = stock_month_table(t); F["book"].extend(r)
                        log.append(("Book stock", fn, "%d month-ends from a monthly stock table (PDF)" % len(r)))
                    elif re.search(r"1ST CHG ON IMMOVABLE PROP", t, re.I):
                        v = [num(x) for x in re.findall(r"1ST CHG ON IMMOVABLE PROP[^\n|]*?(\d{5,})", t, re.I)]
                        F["collateral"] = dict(value=sum(v), count=len(v))
                        log.append(("Collateral", fn, "%d properties, insured value %s (bank's list)" % (len(v), sum(v))))
                    elif pf_receipt(t):
                        r = pf_receipt(t); F["pf"].append(r); log.append(("PF", fn, "%d-%02d Rs %s paid %s" % (r["ym"] + (r["amount"], r["paid"]))))
                    elif esic_receipt(t):
                        r = esic_receipt(t); F["esic"].append(r); log.append(("ESIC", fn, "%d-%02d Rs %s paid %s" % (r["ym"] + (r["amount"], r["paid"]))))
                    elif electricity_bill(t):
                        r = electricity_bill(t); F["elec"].append(r)
                        log.append(("Electricity bill", fn, "%d-%02d net payable %s; previous month paid %s on %s" % (r["ym"] + (r["net"], r["prev_amount"], r["prev_paid"]))))
                    elif policy(t, fn):
                        r = policy(t, fn); F["policies"].append(r); F["names"].add(r["insured"])
                        log.append(("Policy " + r["cls"], fn, "%s %s-%s stock SI %s" % (r["no"], r["frm"], r["to"], r["stock"])))
                    else:
                        log.append(("skipped", fn, "not recognised" if t.strip() else "no text layer (scanned)"))
            elif ext == "txt":
                t = open(p, encoding="utf-8", errors="replace").read()
                if "STATEMENT" in t and "BALANCE" in t:
                    a, tx = bank_txt(t); [x.update(src=fn) for x in tx]; F["bank"].setdefault(a, []).extend(tx); log.append(("Bank TXT", fn, "%d txns; %s" % (len(tx), bal_msg(tx))))
            elif ext in ("xlsx", "xlsm", "csv"):
                if ext == "csv":
                    rows = [l.rstrip("\n").split(",") for l in open(p, encoding="utf-8", errors="replace")]
                else:
                    # classify from the first 200 rows; read the whole workbook only for a type that is used
                    # (big sales registers / GST Excel downloads took 40-85 s each to read in full)
                    top = xlsx_rows(p, 200)
                    k0, h0 = tally_kind(top)
                    if not (k0 or tally_sales(top, h0) or mca_directors(top) or bank_table(top)[1]):
                        log.append(("skipped", fn, "not recognised")); continue
                    rows = xlsx_rows(p)
                k, head = tally_kind(rows)
                if k == "Stock Summary":
                    r = tally_stock(rows, head)
                    if r: F["book"].append(r); log.append(("Book stock", fn, "%s %s lakh" % (r["date"], r["stock"])))
                elif k == "Sundry Debtors":
                    F["debtors"].update(tally_names(rows)); log.append(("Debtor ledger", fn, "names read"))
                    tt = tally_group_total(rows, head)
                    if tt and tt["date"]: F["deb_tot"].append(tt)
                elif k == "Sundry Creditors":
                    F["creditors"].update(tally_names(rows)); log.append(("Creditor ledger", fn, "names read"))
                    tt = tally_group_total(rows, head)
                    if tt and tt["date"]: F["cred_tot"].append(tt)
                elif tally_sales(rows, head):
                    r = tally_sales(rows, head); F.setdefault("sales", {}).update(r)
                    log.append(("Books sales", fn, "%d months (sales net of returns)" % len(r)))
                elif mca_directors(rows):
                    F["directors"] = mca_directors(rows); F["dins"] = mca_dins(rows); log.append(("MCA", fn, ", ".join(F["directors"])))
                else:
                    acc, tx = bank_table(rows)
                    if tx:
                        [x.update(src=fn) for x in tx]
                        F["bank"].setdefault(acc or fn, []).extend(tx); log.append(("Bank sheet", fn, "%d txns; %s" % (len(tx), bal_msg(tx))))
                    else:
                        log.append(("skipped", fn, "not recognised"))
            else:
                log.append(("skipped", fn, "type not used"))
        except Exception as e:
            log.append(("ERROR", fn, repr(e)[:120]))
    # a sheet without an account number whose rows are the same as a real account's (e.g. a working copy such as
    # Summation.xlsx) is merged into that account - otherwise its credits would be counted twice
    accts = [k for k in F["bank"] if re.fullmatch(r"\d{9,}", k)]
    for k in [k for k in list(F["bank"]) if k not in accts]:
        mine = {(t["date"], round(t["amt"], 2)) for t in F["bank"][k]}
        for a in accts:
            theirs = {(t["date"], round(t["amt"], 2)) for t in F["bank"][a]}
            if mine and len(mine & theirs) >= 0.9 * len(mine):
                F["bank"][a].extend(F["bank"].pop(k))
                log.append(("Bank copy", k, "same rows as account %s - treated as a copy, not a second account" % a))
                break
    return F


# ---------------- debtor receipts classification (prompt 1) ----------------
STOP = {"PVT", "PRIVATE", "LTD", "LIMITED", "THE", "M/S", "MS", "AND", "&", "CO", "COMPANY", "SALES", "PURCHASE", "UNIT", "LLP", "INDIA", "URP"}


def toks(name):
    n = re.sub(r"^\s*M\s*/\s*S\.?\s+", "", name.upper())
    return [w for w in re.split(r"[^A-Z0-9]+", n) if len(w) > 1 and w not in STOP and not w.isdigit()]


def nospace(s):
    return re.sub(r"[^A-Z0-9]", "", s.upper())


import functools


@functools.lru_cache(maxsize=None)
def _name_key(name):
    """Words of a party name and its two partial-match patterns, built once per name (not once per bank row)."""
    t = toks(name)
    part = re.compile(r"(^|/)\s*" + re.escape(t[0]) + r"\s*(?:(?:AND|&)\s*)?" + re.escape(t[1][0])) if len(t) >= 2 and len(t[0]) >= 5 else None
    cut = re.compile(r"(^|/)\s*" + re.escape(t[0]) + r"\s*$") if t and len(t[0]) >= 6 else None
    return t, part, cut


def name_hit(name, narr, n=None, U=None):
    """'full' = first two words found; 'part' = distinctive first word (>=6 chars) only."""
    t, part, cut = _name_key(name)
    if n is None:
        n = nospace(narr)
    if not t or t[0] not in n:          # every match below needs the first word in the narration - cheap test first
        return None
    if len(t) >= 2 and len(t[0] + t[1]) >= 6 and (t[0] + t[1]) in n:
        return "full"
    U = U if U is not None else narr.upper()
    # partial matches only where the payer name STARTS (after '/' or at the start) - never inside "X S/O <name>"
    if part is not None and part.search(U):
        return "part"
    if len(t) == 1 and len(t[0]) >= 5:
        return "full"
    if cut is not None and cut.search(U):  # name cut right after first word, e.g. 'Hdfc Bank /AGRAWAL'
        return "part"
    return None


KW = [  # (category, keywords, reason)
    ("Bank-originated credit", ("INT.COLL", "INTEREST", "INT CR", "TXN REV", "REVERSAL", "LOAN DISB", "DISBURS", "FD CLOS", "SWEEP", "MATURITY", "CHQ RTN", "CHEQUE RETURN"), "Bank entry: interest / reversal / disbursement / FD"),
    ("Refunds", ("RETURN/", "NEFT/RETURN", "REFUND", "RTN/"), "Own payment returned or refund"),
    ("Tax-related", ("GST REFUND", "CBDT", "ITD", "INCOME TAX", "TDS REFUND", "PMVBRY", "SUBSIDY"), "Tax refund / government scheme credit"),
    ("Investment/other income", ("DIVIDEND", "MUTUAL FUND", "REDEMPTION", " MF "), "Investment income"),
    ("Loans/borrowings (non-bank)", ("UNSECURED LOAN", "LOAN FROM", "USL"), "Loan received"),
]


def classify(tx, F, related, own):
    out = []
    debtors = sorted(F["debtors"], key=len, reverse=True)
    creditors = sorted(F["creditors"])
    deb_first = {toks(dn)[0] for dn in debtors if toks(dn)}
    own_keys = [("".join(toks(nm)), nm) for nm in own if nm and toks(nm)]
    for t in tx:
        if t["amt"] <= 0:
            continue
        u = t["narr"].upper(); r = dict(t); n = nospace(u)
        hit = None
        for key, nm in own_keys:
            if key in n:
                hit = ("EXC", "Self-transfer / contra", "Own name (%s) - transfer from own account" % nm)
        if not hit:
            for nm in related:
                if name_hit(nm, u, n, u):
                    extra = " - also in debtor ledger" if toks(nm) and toks(nm)[0] in deb_first else ""
                    hit = ("EXC", "Director/Partner/Owner/Related party", "Related party / director: %s%s" % (nm, extra)); break
        if not hit:
            for cat, kws, why in KW:
                k = next((k for k in kws if k in u), None)
                if k:
                    hit = ("EXC", cat, "%s (narration has '%s')" % (why, k.strip())); break
        best = None
        if not hit:
            # debtor match once per row (used both to block the supplier-refund rule and as the debtor receipt)
            for nm in debtors:
                h = name_hit(nm, u, n, u)
                if h == "full":
                    best = (nm, ""); break
                if h == "part" and not best:
                    best = (nm, "first-word match only")
            for nm in creditors:
                if not best and name_hit(nm, u, n, u):
                    hit = ("EXC", "Refunds", "Payer is a supplier in creditor ledger (%s) - refund" % nm); break
        if not hit and best:
            hit = ("INC", best[0], best[1])
        if not hit:
            if re.search(r"TREASU|GOVT|GOVERNMENT|COLLECTORATE|SECRETARIAT|SECRETRIATE", u):
                hit = ("REV", "Unidentified / needs review", "Government / treasury payment - a government customer paying for "
                       "supplies (debtor receipt) or a refund / subsidy (exclude); confirm from the debtor ledger")
            else:
                hit = ("REV", "Unidentified / needs review", "Payer not in debtor ledger and no rule matched - confirm nature")
        r["b"], r["cat"], r["why"] = hit
        out.append(r)
    return out


def dedupe(txs):
    """Same account read from several files (e.g. full-period TXT + monthly PDFs): for each month keep the ONE source
    file - never mix two files in one month. Preference: running balance ties for that month, then most transactions."""
    by = collections.defaultdict(lambda: collections.defaultdict(list))
    for t in txs:
        by[(t["date"].year, t["date"].month)][t.get("src", "")].append(t)
    out = []
    for k in sorted(by):
        src = max(by[k].items(), key=lambda kv: (balance_check(kv[1])[1] == 0, len(kv[1])))
        out += [dict(t, src=src[0]) for t in src[1]]
    return sorted(out, key=lambda t: t["date"])


# ---------------- insurance date mapping (prompt 3) ----------------
def pdate(s):
    return dt.datetime.strptime(s, "%d/%m/%Y").date()


def cover(pols, cls, day):
    """Policy (or several) of one class in force on day. A renewal overlap of a few days -> the later policy (rule 9).
    Policies running side by side for months (separate locations / blocks) -> their stock SIs are added (rule 3)."""
    live = sorted((p for p in pols if p["cls"] == cls and pdate(p["frm"]) <= day <= pdate(p["to"])), key=lambda p: pdate(p["frm"]))
    if not live:
        return None, ""
    keep = []
    for p in live:
        keep = [k for k in keep if (min(pdate(k["to"]), pdate(p["to"])) - max(pdate(k["frm"]), pdate(p["frm"]))).days > 60]
        keep.append(p)
    note = "overlap - later policy used" if len(live) > 1 and len(keep) == 1 else ""
    if len(keep) == 1:
        return keep[0], note
    sis = [k["stock"] for k in keep]
    fls = {k["first_loss"] for k in keep}
    agg = dict(keep[-1], no=" + ".join(k["no"] for k in keep), stock=None if None in sis else round(sum(sis), 2),
               first_loss=fls.pop() if len(fls) == 1 else None, sub=None, iar=any(k.get("iar") for k in keep))
    return agg, "%d policies in force together (separate locations) - stock SIs added: %s" % (
        len(keep), " + ".join("%s" % k["stock"] for k in keep))


def month_end(y, m):
    return (dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1))


# ---------------- Excel ----------------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
BOLD = Font(bold=True); HEAD = PatternFill("solid", fgColor="DDEBF7"); WARN = PatternFill("solid", fgColor="FCE4D6")
NUMF = "#,##0.00"


def head(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row, i, c); x.font = BOLD; x.fill = HEAD; x.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(widths or []):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w


def mlabel(y, m):
    return dt.date(y, m, 1).strftime("%b-%Y")


def build_excel(F, cfg, out):
    wb = Workbook(); wb.remove(wb.active)
    borrower = cfg.get("borrower") or next(iter(F["names"]), "")
    # ===== SUMMATION (debtor receipts) =====
    allc = []
    for acc, tx in F["bank"].items():
        for r in classify(dedupe(tx), F, cfg["related"], cfg["own"]):
            r["acc"] = acc; allc.append(r)
    allc.sort(key=lambda r: r["date"])
    ws = wb.create_sheet("Summation - Included")
    ws["A1"] = "%s - debtor receipts (credits linked to customers in the debtor ledger)" % borrower; ws["A1"].font = BOLD
    head(ws, 2, ["Date", "Account", "Narration", "Amount (Rs)", "Matched Customer", "Note"], [12, 18, 60, 16, 40, 24])
    for r in allc:
        if r["b"] == "INC":
            ws.append([r["date"], r["acc"], r["narr"], r["amt"], r["cat"], r["why"]])
    ws = wb.create_sheet("Summation - Excluded")
    head(ws, 1, ["Date", "Account", "Narration", "Amount (Rs)", "Exclusion Category", "Reason / basis"], [12, 18, 60, 16, 34, 70])
    for r in allc:
        if r["b"] == "EXC":
            ws.append([r["date"], r["acc"], r["narr"], r["amt"], r["cat"], r["why"]])
    ws = wb.create_sheet("Summation - Needs Review")
    head(ws, 1, ["Date", "Account", "Narration", "Amount (Rs)", "Why it could not be classified"], [12, 18, 60, 16, 70])
    for r in allc:
        if r["b"] == "REV":
            ws.append([r["date"], r["acc"], r["narr"], r["amt"], r["why"]])
    ws = wb.create_sheet("Summation - Monthly", 0)
    ws["A1"] = "%s - Monthly debtor receipts from bank statement(s): %s" % (borrower, ", ".join(F["bank"]) or "NO BANK STATEMENT FOUND"); ws["A1"].font = BOLD
    head(ws, 2, ["Month", "Total Credits", "Debtor Receipts (Included)", "Total Excluded", "Needs Review", "Count Included",
                 "Count Excluded", "Count Review", "Total Debits", "Check (Incl+Excl+Review = Total)"], [12, 18, 20, 16, 16, 10, 10, 10, 18, 22])
    agg = collections.OrderedDict()
    for acc, tx in F["bank"].items():
        for t in dedupe(tx):
            a = agg.setdefault((t["date"].year, t["date"].month), collections.Counter())
            if t["amt"] < 0: a["DR"] += -t["amt"]
    for r in allc:
        a = agg.setdefault((r["date"].year, r["date"].month), collections.Counter())
        a["T"] += r["amt"]; a[r["b"]] += r["amt"]; a["n" + r["b"]] += 1
    for k in sorted(agg):
        a = agg[k]
        ok = abs(a["T"] - a["INC"] - a["EXC"] - a["REV"]) < 0.01
        ws.append([mlabel(*k), round(a["T"], 2), round(a["INC"], 2), round(a["EXC"], 2), round(a["REV"], 2), a["nINC"], a["nEXC"], a["nREV"], round(a["DR"], 2), "OK" if ok else "MISMATCH"])
        if not ok:
            ws.cell(ws.max_row, 10).fill = WARN
    n = ws.max_row
    if n > 2:
        ws.append(["Total"] + ["=SUM(%s3:%s%d)" % (c, c, n) for c in "BCDEFGHI"])
        for c in ws[ws.max_row]: c.font = BOLD
    ws.append([]); ws.append(["Needs-review items are NOT in debtor receipts. Move them after confirming with the borrower."])
    ws.append(["Related parties used: " + ("; ".join(cfg["related"]) or "none given")])
    ws.append(["Own names used: " + "; ".join(n for n in cfg["own"] if n)])
    # ===== GST =====
    ws = wb.create_sheet("GST")
    g1 = {}; g3 = {}
    for r in F["g1"]: g1[(r["gstin"], r["ym"])] = r
    for r in F["g3"]: g3[(r["gstin"], r["ym"])] = r
    gst_ins = sorted({k[0] for k in list(g1) + list(g3)})
    months = sorted({k[1] for k in list(g1) + list(g3)})
    ws["A1"] = "Month-wise outward supply - TAXABLE VALUE ONLY (Rs). GSTR-1 = B2B + B2CL + B2CS + exports + SEZ + net credit/debit notes + amendments"; ws["A1"].font = BOLD
    head(ws, 3, ["GSTIN"] + [mlabel(*m) for m in months] + ["Total"], [34] + [14] * (len(months) + 1))
    L = openpyxl.utils.get_column_letter
    for gi in gst_ins:
        for lab, src in (("GSTR-1", g1), ("GSTR-3B 3.1(a)", g3)):
            row = [gi + "  " + lab] + [round(src[(gi, m)]["taxable"], 2) if (gi, m) in src else 0 for m in months]
            ws.append(row + ["=SUM(B%d:%s%d)" % (ws.max_row + 1, L(len(months) + 1), ws.max_row + 1)])
        r1 = ws.max_row - 1
        ws.append(["Difference (GSTR-1 - 3B)"] + ["=%s%d-%s%d" % (L(i + 2), r1, L(i + 2), r1 + 1) for i in range(len(months) + 1)])
    gr = ws.max_row + 1
    if gst_ins:
        ws.append(["Grand Total (GSTR-1)"] + ["=SUMIF($A$4:$A$%d,\"*GSTR-1\",%s4:%s%d)" % (gr - 1, L(i + 2), L(i + 2), gr - 1) for i in range(len(months) + 1)])
        for c in ws[ws.max_row]: c.font = BOLD
        grand = ws.max_row; books_row = None
        sales = F.get("sales") or {}
        if sales:                                # books cover every GSTIN of the company -> compared with the grand total
            ws.append(["Books - sales register (Tally, net of returns)"] + [sales.get(m, "not in books") for m in months]
                      + ["=SUM(B%d:%s%d)" % (ws.max_row + 1, L(len(months) + 1), ws.max_row + 1)])
            books_row = ws.max_row
            # total = sum of monthly differences, so months missing from the books are not counted as a gap
            ws.append(["Difference (Books - GSTR-1, months in books)"] + ["=IF(ISNUMBER(%s%d),%s%d-%s%d,\"-\")" % (L(i + 2), books_row, L(i + 2), books_row, L(i + 2), grand)
                                                         for i in range(len(months))] + ["=SUM(B%d:%s%d)" % (ws.max_row + 1, L(len(months) + 1), ws.max_row + 1)])
            diff_row = ws.max_row
            for i, m in enumerate(months):
                g = sum(g1[(gi, m)]["taxable"] for gi in gst_ins if (gi, m) in g1)
                if m in sales and abs(sales[m] - g) > 1:
                    ws.cell(ws.max_row, i + 2).fill = WARN
        # financial-year totals (April-March), as SUM formulas over the month columns above
        fys = collections.OrderedDict()
        for i, (y, m) in enumerate(months):
            fys.setdefault(y if m >= 4 else y - 1, []).append(i + 2)
        ws.append([])
        hr = ws.max_row + 1
        head(ws, hr, ["Financial-year totals (Rs)"] + ["FY %d-%02d%s" % (fy, (fy + 1) % 100, "" if len(c) == 12 else " (%d months)" % len(c))
                                                     for fy, c in fys.items()])
        for src_row in range(4, gr, 3):          # each GSTIN has 3 rows above: GSTR-1, GSTR-3B, Difference
            for k in (0, 1):
                ws.append([ws.cell(src_row + k, 1).value] + ["=SUM(%s%d:%s%d)" % (L(c[0]), src_row + k, L(c[-1]), src_row + k) for c in fys.values()])
            ws.append(["Difference (GSTR-1 - 3B)"] + ["=%s%d-%s%d" % (L(j + 2), ws.max_row - 1, L(j + 2), ws.max_row) for j in range(len(fys))])
        if books_row:
            ws.append(["Books - sales register (months in books)"] + ["=SUM(%s%d:%s%d)" % (L(c[0]), books_row, L(c[-1]), books_row) for c in fys.values()])
            ws.append(["Difference (Books - GSTR-1, months in books)"] + ["=SUM(%s%d:%s%d)" % (L(c[0]), diff_row, L(c[-1]), diff_row) for c in fys.values()])
            miss = [mlabel(*m) for m in months if m not in sales]
            if miss:
                ws.append(["Months not in the books file: " + ", ".join(miss)])
    bad = [r for r in F["g1"] + F["g3"] if not re.fullmatch(r"\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]", r["gstin"])]
    ws.append([]); ws.append(["Rows skipped: 0 blank taxable value; invalid GSTIN: %d. Period: %s to %s." % (
        len(bad), mlabel(*months[0]) if months else "-", mlabel(*months[-1]) if months else "-")])
    ws.append([]); r0 = ws.max_row + 1
    head(ws, r0, ["Return filing", "Tax period", "Due date", "Filed (ARN date)", "Remarks"])
    for r in sorted(F["g3"] + F["g1"], key=lambda r: (r["kind"], r["ym"]), reverse=True):
        y, m = r["ym"]; ny, nm = (y + (m == 12), m % 12 + 1)
        due = dt.date(ny, nm, 20 if r["kind"] == "3B" else 11)
        fd = d(r["arn"]) if r["arn"] else None
        ws.append(["GSTR-" + r["kind"] + "  " + r["gstin"], mlabel(y, m), due, fd, "Timely Filed" if fd and fd <= due else ("Late Filed" if fd else "ARN date not found")])
    build_insurance(wb, F, cfg)
    for s in wb.worksheets:
        for row in s.iter_rows(min_row=2):
            for c in row:
                # money: every float; on the GST sheet also whole-rupee ints and formulas (no counts live there)
                if isinstance(c.value, float) or (s.title == "GST" and c.column > 1 and not isinstance(c.value, bool)
                                                  and (isinstance(c.value, int) or str(c.value).startswith("="))):
                    c.number_format = NUMF
                if isinstance(c.value, dt.date): c.number_format = "DD-MM-YYYY"
    wb.save(out)
    return allc


def build_insurance(wb, F, cfg):
    ws = wb.create_sheet("Insurance")
    # one row per policy number; of several copies keep a stock policy that has its stock SI (not an endorsement / letter)
    best = collections.OrderedDict()
    for p in F["policies"]:
        q = best.get(p["no"])
        rank = lambda x: (x["cls"] != "Other", x["stock"] is not None)
        if q is None or rank(p) > rank(q):
            best[p["no"]] = p
    pols = list(best.values())
    dups = len(F["policies"]) - len(pols)
    own = [n for n in cfg["own"] if n]
    notes = []
    for p in pols:
        if own and not any(name_hit(n, p["insured"]) for n in own):
            notes.append("POLICY %s IS IN THE NAME OF '%s' - NOT THE BORROWER; not merged." % (p["no"], p["insured"]))
            p["foreign"] = True
        if cfg.get("gstin") and p["gstin"] and p["gstin"] != cfg["gstin"]:
            notes.append("Policy %s GSTIN %s differs from borrower GSTIN %s." % (p["no"], p["gstin"], cfg["gstin"]))
    other = [p for p in pols if p["cls"] == "Other"]
    pols = [p for p in pols if p["cls"] != "Other"]
    use = [p for p in pols if not p.get("foreign")]
    r = 1
    for n in notes:
        ws.cell(r, 1, n).font = Font(bold=True, color="C00000"); r += 1
    if pols and not notes:
        names = sorted({p["insured"] for p in pols if p["insured"]}); gsts = sorted({p["gstin"] for p in pols if p["gstin"]})
        nog = [p["no"] for p in pols if not p["gstin"]]
        ws.cell(r, 1, "Insured check: all %d policies are in the name of %s%s - same as the borrower%s.%s" % (
            len(pols), " / ".join(names) or "(name not printed)", (", GSTIN " + " / ".join(gsts)) if gsts else "",
            (" (" + "; ".join(n for n in own) + (", GSTIN " + cfg["gstin"] if cfg.get("gstin") else "") + ")") if own else "",
            (" GSTIN not printed on: " + ", ".join(nog) + ".") if nog else "")).font = Font(bold=True, color="1A7F37"); r += 1
    if dups:
        ws.cell(r, 1, "%d duplicate upload(s) of the same policy number ignored." % dups); r += 1
    if other:
        ws.cell(r, 1, "Not stock policies (left out): " + "; ".join("%s - %s (%s)" % (p["no"], p["build"], p["file"]) for p in other)); r += 1
    r += 1; ws.cell(r, 1, "Table 1 - Policy inventory (Rs lakh)").font = BOLD; r += 1
    head(ws, r, ["Policy No.", "Insurer", "Class", "Period of Insurance", "Qualifying SI (stock only)", "Build-up / source line",
                 "Items excluded and value", "First loss %", "Insured / GSTIN", "Prev. Policy No.", "File"],
         [24, 30, 12, 26, 14, 60, 50, 10, 40, 24, 36])
    for p in sorted(use, key=lambda p: (p["cls"], pdate(p["frm"]))):
        r += 1
        ws.append([p["no"], p["insurer"], p["cls"], "%s to %s" % (p["frm"], p["to"]),
                   p["stock"] if p["stock"] is not None else "SI not stated", p["build"], p["excluded"],
                   ("%d%%" % p["first_loss"]) if p["first_loss"] else "", "%s / %s" % (p["insured"], p["gstin"]), p["prev"], p["file"]])
    # dates: stock statements + book stock + extra dates typed on the portal
    rows = collections.OrderedDict()
    for s in F["ss"]:
        rows.setdefault(s["date"], {})["ss"] = s["stock"]
        rows[s["date"]]["split"] = s.get("split")
        rows[s["date"]]["check"] = rows[s["date"]].get("check") or bool(s.get("check"))
    for b in F["book"]:
        if b["date"]:
            rows.setdefault(b["date"], {})["book"] = b["stock"]
    for x in cfg.get("extra", []):
        rows.setdefault(x["date"], {}).update({"ss": x.get("stock")} if x.get("stock") is not None else {}, label=x.get("label", "Visit date"))
    r = ws.max_row + 2
    ws.cell(r, 1, "Table 2 - Result (Rs lakh)").font = BOLD; r += 1
    head(ws, r, ["Month / Date", "Total Stock - Stock Statement", "Total Stock - Books", "Fire - Insured Amt (lakh)",
                 "Burglary - Insured Amt (lakh)", "Fire policy", "Burglary policy", "Excess / (Shortfall) Fire",
                 "Fire sub-limit check (stock statement category vs policy SI)", "Remark"])
    ws.column_dimensions["I"].width = 46; ws.column_dimensions["J"].width = 60
    gaps, prev_stock = [], None
    for day in sorted(rows):
        v = rows[day]
        lab = v.get("label") or (day.strftime("%b'%y") + " (" + day.strftime("%d-%m-%Y") + ")")
        cells = []
        rem = []
        for cls in ("Fire", "Burglary"):
            p, ov = cover(use, cls, day)
            if p is None:
                cells.append(("Policy not provided", "")); gaps.append((day, cls))
            else:
                si = p["stock"] if p["stock"] is not None else "SI not stated"
                cells.append((si, p["no"]))
                if ov: rem.append("%s: %s" % (cls, ov))
                if cls == "Burglary" and p["first_loss"]:
                    rem.append("Burglary first loss %d%% = max claim %.2f" % (p["first_loss"], (p["stock"] or 0) * p["first_loss"] / 100))
        stock = max([x for x in (v.get("ss"), v.get("book")) if x is not None] or [0])
        fire = cells[0][0]
        if isinstance(fire, (int, float)) and stock > fire:
            rem.insert(0, "UNDER-INSURED (Fire)")
        if isinstance(cells[1][0], (int, float)) and stock > cells[1][0]:
            rem.insert(0, "UNDER-INSURED (Burglary)")
        if isinstance(fire, (int, float)) and fire == 0:
            rem.append("Stock SI nil")
        if prev_stock and stock and (stock / prev_stock >= 2 or prev_stock / stock >= 2):
            rem.append("Stock %.1fx previous row - check" % (stock / prev_stock))
        prev_stock = stock or prev_stock
        if v.get("ss") is None and v.get("book") is None:
            rem.insert(0, "No stock figure for this date - enter stock on the portal")
        if v.get("check"):
            rem.append("stock statement figure read from a scanned / OCR'd page - CHECK against the statement")
        fp, _ = cover(use, "Fire", day); sub = fp.get("sub") if fp else None; sp = v.get("split")
        if sub and sp:
            parts, short = [], []
            for k in ("RM", "WIP", "FG"):
                st, lim = sp.get(k) or 0, sub.get(k) or 0
                if st > lim:
                    parts.append("%s %.2f > SI %.2f - SHORT %.2f" % (k, st, lim, st - lim)); short.append(k)
                else:
                    parts.append("%s %.2f / SI %.2f OK" % (k, st, lim))
            if sp.get("Stores"):
                parts.append("Stores %.2f - not in stock SI" % sp["Stores"])
            subtxt = "; ".join(parts)
            if short:
                rem.insert(0, "UNDER-INSURED (Fire %s sub-limit)" % "/".join(short))
        elif sub:
            subtxt = "No category split for this date (books / typed stock) - total compared only"
        else:
            subtxt = "Policy has no category sub-limits" if fp else ""
        ws.append([lab, v.get("ss"), v.get("book"), cells[0][0], cells[1][0], cells[0][1], cells[1][1],
                   (fire - stock) if isinstance(fire, (int, float)) and stock else None, subtxt, "; ".join(rem) or "Adequately insured"])
        if any("UNDER" in x for x in rem):
            for c in ws[ws.max_row]: c.fill = WARN
    r = ws.max_row + 2
    ws.cell(r, 1, "Gaps / missing policies").font = BOLD
    if not gaps:
        ws.cell(r + 1, 1, "None - every date is inside a supplied Fire and Burglary policy period (inclusive).")
    for i, (day, cls) in enumerate(gaps):
        prevs = [p["prev"] for p in use if p["cls"] == cls and p["prev"] and pdate(p["frm"]) > day]
        fp, _ = cover(use, "Fire", day)
        iar = " (fire policy %s is Industrial All Risk - check whether it covers burglary / theft)" % fp["no"] if cls == "Burglary" and fp and fp.get("iar") else ""
        ws.cell(r + 1 + i, 1, "%s: no %s policy supplied%s%s" % (day.strftime("%d-%m-%Y"), cls,
                ("; missing policy probably %s (Prev. Policy No. on later schedule)" % prevs[0]) if prevs else "", iar))
    ws.cell(ws.max_row + 2, 1, "Fire policies with separate RM / WIP / FG sums insured are checked category by category against the stock statement split - a total within the SI can still be under-insured in one category.")
    ws.cell(ws.max_row + 1, 1, "Only stock in trade (RM + WIP + FG / Stocks in trade) counted. Building, P&M, FFF, other contents, cash, LOP excluded. Units: Rs lakh as in the stock statement.")
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")





# ================= Excel in the user's own layout (docs/template.xlsx = blank copy of the human draft) =================
TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.xlsx")


def _ym_add(ym, k):
    y, m = ym; m += k
    while m < 1: m += 12; y -= 1
    while m > 12: m -= 12; y += 1
    return (y, m)


def _mlab(ym):
    return dt.date(ym[0], ym[1], 1).strftime("%b'%y")


def _lakh(v):
    return None if v is None else round(v / 1e5, 5)


def _is_cc(tx):
    bals = sorted(t["bal"] for t in tx if t.get("bal") is not None)
    return bool(bals) and bals[len(bals) // 2] < 0          # overdrawn most of the time -> cash credit


def _book_total(rows, ym, creditors=False):
    """Books debtors (debit - credit) or creditors (credit balances) at a month end, Rs."""
    for r in rows:
        if r["date"] == month_end(*ym):
            return r["cr"] if creditors else r["dr"] - r["cr"]
    return None


def _short_bank(name):
    """'Axis Bank Ltd' / 'AXIS BANK' / 'Axis' -> 'Axis';  'Union Bank of India' -> 'Union'"""
    w = re.sub(r"(?i)\b(bank|ltd\.?|limited|of india)\b", " ", name or "").split()
    return w[0].title() if w else (name or "")


def draft_summation(ws, F, cfg, allc):
    base = cfg["sum_to"]; start = cfg["sum_from"]
    months = [];  ym = start
    while ym <= base:
        months.append(ym); ym = _ym_add(ym, 1)
    per = "%s To %s" % (_mlab(start), _mlab(base))
    banks = cfg.get("banks") or []
    bank1 = banks[0][0] if banks else "Bank"
    accs = sorted(F["bank"].items(), key=lambda kv: -len(kv[1]))
    cc = next((a for a, tx in accs if _is_cc(dedupe(tx))), accs[0][0] if accs else "")
    ca = next((a for a, tx in accs if a != cc), None)
    inp = lambda t: start <= (t["date"].year, t["date"].month) <= base

    def sums(acc):
        if acc is None:
            return None, None
        tx = [t for t in dedupe(F["bank"][acc]) if inp(t)]
        dr = sum(-t["amt"] for t in tx if t["amt"] < 0)
        rec = sum(r["amt"] for r in allc if r["acc"] == acc and r["b"] == "INC" and inp(r))   # debtor receipts only
        return _lakh(dr), _lakh(rec)

    ws["B2"] = " Cash Credit A/C %s Bank: %s   " % (_short_bank(bank1), cc)
    ws["F2"] = " Current A/C %s Bank: %s" % (_short_bank(bank1), ca or "")
    for c, v in (("B3", "Period"), ("C3", "Debit Summation"), ("D3", "Credit Summation"),
                 ("F3", "Period"), ("G3", "Debit Summation"), ("H3", "Credit Summation")):
        ws[c] = v
    ws["B4"] = per; ws["C4"], ws["D4"] = sums(cc)
    if ca:
        ws["F4"] = per; ws["G4"], ws["H4"] = sums(ca)
    ws["B5"] = "Total"; ws["C5"] = "=SUM(C4:C4)"; ws["D5"] = "=SUM(D4:D4)"
    ws["F5"] = "Total"; ws["G5"] = "=SUM(G4:G4)"; ws["H5"] = "=SUM(H4:H4)"
    ws["B9"] = "Particulars"; ws["C9"] = "Amount (In lakhs)"
    opening = _book_total(F.get("deb_tot", []), _ym_add(start, -1))
    closing = _book_total(F.get("deb_tot", []), base)
    g3 = {r["ym"]: r["taxable"] for r in F["g3"]} or {r["ym"]: r["taxable"] for r in F["g1"]}
    sales = sum(g3.get(m, 0) for m in months) if all(m in g3 for m in months) else None
    ws["B10"] = "Opening Debtors as on %s" % dt.date(start[0], start[1], 1).strftime("%d-%m-%Y"); ws["C10"] = _lakh(opening)
    ws["B11"] = "Add: Sales during %s" % per; ws["C11"] = _lakh(sales)
    ws["B12"] = "Less: Closing Debtors as on %s" % month_end(*base).strftime("%d-%m-%Y"); ws["C12"] = _lakh(closing)
    ws["B13"] = "Sales Realization during %s" % per; ws["C13"] = "=C10+C11-C12"
    b2 = banks[1] if len(banks) > 1 else ("", None)
    sh1 = _short_bank(bank1); sh2 = _short_bank(b2[0]) if b2[0] else "other"
    ws["B14"] = "Limit Sanctioned (%s bank)" % sh1; ws["C14"] = banks[0][1] if banks else None
    ws["B15"] = "Limit Sanctioned (%s bank)" % sh2; ws["C15"] = b2[1] if b2[1] is not None else 0
    ws["B16"] = "Total Limit Sanctioned"; ws["C16"] = "=C14+C15"
    ws["B17"] = "Proportionate DP for %s Bank Ltd." % sh1; ws["C17"] = "=C14*C13/C16"
    ws["B18"] = "Proportionate DP for %s Bank Ltd." % sh2; ws["C18"] = "=C13*C15/C16"
    ws["B19"] = "Routed In %s – CC" % sh1; ws["C19"] = "=D5"
    ws["B20"] = "Routed In %s – CA" % sh1; ws["C20"] = "=H5" if ca else 0
    ws["B21"] = "Total Receipt In %s" % sh1; ws["C21"] = "=C19+C20"
    ws["B22"] = "Sales Realization in %"; ws["C22"] = "=C21/C17%"


def _pol_block(ws, col, top, p, items):
    """One policy block of the draft layout (5 columns starting at col, starting row top)."""
    L = openpyxl.utils.get_column_letter
    c = lambda k: L(openpyxl.utils.column_index_from_string(col) + k)
    if p is None:
        return
    bank = re.split(r"\s*[,&]\s*", p.get("bank") or "")[0].strip(" .")
    ws["%s%d" % (c(0), top)] = "Issued by: %s" % (p.get("insurer") or "")
    ws["%s%d" % (c(3), top)] = "Nature:  %s" % re.sub(r"\s*\([^)]*\)", "", p.get("nature") or p["cls"]).strip()
    ws["%s%d" % (c(0), top + 1)] = "Bank Clause: %s" % ("Included in the policy" if bank else "Not stated in the policy")
    ws["%s%d" % (c(3), top + 1)] = "Period:  From: %s To: %s" % (p["frm"], p["to"])
    ws["%s%d" % (c(0), top + 2)] = ("Interest of %s Bank: Included & Endorsed in favour of  %s " % (_short_bank(bank), bank)
                                    if bank else "Interest of Bank: Not stated in the policy")
    ws["%s%d" % (c(0), top + 3)] = "Policy No"; ws["%s%d" % (c(1), top + 3)] = "Risk Covered"
    ws["%s%d" % (c(2), top + 3)] = "SI (Rs in lakh)"; ws["%s%d" % (c(4), top + 3)] = "Location covered:"
    ws["%s%d" % (c(0), top + 4)] = p["no"]; ws["%s%d" % (c(4), top + 4)] = p.get("location") or ""
    for i, (lab, v) in enumerate(items):
        ws["%s%d" % (c(1), top + 4 + i)] = lab
        ws["%s%d" % (c(2), top + 4 + i)] = round(v, 2) if isinstance(v, (int, float)) else "SI not stated"


def draft_insurance(ws, F, cfg):
    pols = {}
    for p in F["policies"]:
        if p["cls"] == "Other":
            continue
        q = pols.get(p["no"])
        if q is None or (q["stock"] is None and p["stock"] is not None):
            pols[p["no"]] = p
    fire = sorted((p for p in pols.values() if p["cls"] == "Fire"), key=lambda p: pdate(p["frm"]), reverse=True)
    burg = sorted((p for p in pols.values() if p["cls"] == "Burglary"), key=lambda p: pdate(p["frm"]), reverse=True)
    fitems = lambda p: [(k, p["stock"] if k == "Stock" else v) for k, v in (p.get("items") or [("Stock", p["stock"])])][:5]
    _pol_block(ws, "B", 1, fire[0] if fire else None, fitems(fire[0]) if fire else [])
    _pol_block(ws, "H", 1, fire[1] if len(fire) > 1 else None, fitems(fire[1]) if len(fire) > 1 else [])
    _pol_block(ws, "B", 12, burg[0] if burg else None, [("Stock", burg[0]["stock"])] if burg else [])
    _pol_block(ws, "H", 12, burg[1] if len(burg) > 1 else None, [("Stock", burg[1]["stock"])] if len(burg) > 1 else [])
    # stock vs insured: the last three stock-statement dates up to the summation end month + the visit date
    base = month_end(*cfg["sum_to"])
    books = {b["date"]: b["stock"] for b in F["book"] if b.get("date")}
    ss = {s["date"]: s for s in F["ss"] if s.get("date")}
    dates = sorted(x for x in ss if x <= base)[-3:]
    visit = cfg.get("visit") or max([x for x in books if x > base] or [None], key=lambda x: x or dt.date.min)
    for c, v in (("B20", "Month"), ("C20", "Total Stock"), ("D20", "Insured Amt (Fire)"), ("E20", "Insured Amt (Burglary)"), ("F20", "Remarks")):
        ws[c] = v
    use = list(pols.values())
    rows = [(x, _mlab((x.year, x.month))) for x in dates] + ([(visit, "Visit date %s" % visit.strftime("%d-%m-%Y"))] if visit else [])
    for i, (day, lab) in enumerate(rows[:4]):
        r = 21 + i
        stock = books.get(day)
        if stock is None and day in ss:
            stock = ss[day]["stock"]
        if stock is None and cfg.get("visit_stock") is not None and day == visit:
            stock = cfg["visit_stock"]
        fp, _ = cover(use, "Fire", day); bp, _ = cover(use, "Burglary", day)
        fsi = fp["stock"] if fp and fp["stock"] is not None else ("SI not stated" if fp else "Policy not provided")
        bsi = bp["stock"] if bp and bp["stock"] is not None else ("SI not stated" if bp else "Policy not provided")
        ws["B%d" % r] = lab; ws["C%d" % r] = round(stock, 2) if stock is not None else None
        ws["D%d" % r] = fsi; ws["E%d" % r] = bsi
        short = []
        sub, sp = (fp or {}).get("sub"), (ss.get(day) or {}).get("split")
        if sub and sp:
            short = [k for k in ("RM", "WIP", "FG") if (sp.get(k) or 0) > (sub.get(k) or 0)]
        if stock is None or not isinstance(fsi, (int, float)):
            rem = "Unable to comment"
        elif stock > fsi or (isinstance(bsi, (int, float)) and stock > bsi) or short:
            rem = "Inadequately Insured" + (" (%s sub-limit)" % "/".join(short) if short and stock <= fsi else "")
        else:
            rem = "Adequately Insured"
        ws["F%d" % r] = rem
    # fixed assets and collateral as in the draft: insured value from the current fire policy; asset values typed on the portal
    ws["B27"] = "Fixed asset schedule"
    for c, v in (("B28", "Particulars"), ("C28", "Assets value (Rs. Lakhs)"), ("D28", "Insured Value (Rs. In Lakhs)"), ("E28", "Remarks")):
        ws[c] = v
    it = dict((k.lower(), v) for k, v in (fire[0].get("items") or [])) if fire else {}
    pm = next((v for k, v in it.items() if "plant" in k), None)
    bld = next((v for k, v in it.items() if "building" in k), None)
    ws["B29"] = "Plant & Machinery"; ws["C29"] = cfg.get("pm_value"); ws["D29"] = round(pm, 2) if pm is not None else None
    if cfg.get("pm_value") is not None and pm is not None:
        ws["E29"] = "Adequately Insured" if pm >= cfg["pm_value"] else "Inadequately Insured"
    for c, v in (("B31", "Address of collateral"), ("C31", "Assets value (Rs. Lakhs)"), ("D31", "Insured Value (Rs. In Lakhs)"), ("E31", "Remarks")):
        ws[c] = v
    ws["B32"] = cfg.get("collateral_address") or (fire[0].get("location") if fire else "")
    cv = cfg.get("collateral_value")
    if cv is None and F.get("collateral"):
        cv = round(F["collateral"]["value"] / 1e5, 5)
    ws["C32"] = cv; ws["D32"] = round(bld, 2) if bld is not None else None
    cfg = dict(cfg, collateral_value=cv)
    if cfg.get("collateral_value") is not None and bld is not None:
        ws["E32"] = "Adequately Insured" if bld >= cfg["collateral_value"] else "Inadequately Insured"


GST_STATES = {"01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh", "05": "Uttarakhand",
              "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
              "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura", "17": "Meghalaya",
              "18": "Assam", "19": "West Bengal", "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh",
              "24": "Gujarat", "26": "Dadra and Nagar Haveli and Daman and Diu", "27": "Maharashtra", "29": "Karnataka", "30": "Goa",
              "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry", "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh"}


def _fy(ym):
    y = ym[0] if ym[1] >= 4 else ym[0] - 1
    return "%d-%d" % (y, y + 1)


def _mname(ym):
    return dt.date(ym[0], ym[1], 1).strftime("%B")


def _ret_rows(ws, rows, col, top, kind, n):
    L = openpyxl.utils.get_column_letter; c0 = openpyxl.utils.column_index_from_string(col)
    for i, r in enumerate(sorted(rows, key=lambda r: r["ym"], reverse=True)[:n]):
        y, m = r["ym"]; due = dt.date(y + (m == 12), m % 12 + 1, 20 if kind == "3B" else 11)
        fd = d(r["arn"]) if r.get("arn") else None
        vals = [_fy(r["ym"]), _mname(r["ym"]), due, fd, ("Timely Filed" if fd <= due else "Late Filed") if fd else "Filing date not found"]
        for k, v in enumerate(vals):
            ws["%s%d" % (L(c0 + k), top + i)] = v


def _due_rows(ws, rows, col, top, label, n):
    L = openpyxl.utils.get_column_letter; c0 = openpyxl.utils.column_index_from_string(col)
    for i, r in enumerate(sorted(rows, key=lambda r: r["ym"], reverse=True)[:n]):
        y, m = r["ym"]; due = dt.date(y + (m == 12), m % 12 + 1, 15)
        rem = ("Timely Paid" if r["paid"] <= due else "Lately Paid") if r.get("paid") else "Payment date not found"
        for k, v in enumerate([label, _mname(r["ym"]), r["amount"], r.get("paid"), rem]):
            ws["%s%d" % (L(c0 + k), top + i)] = v


def draft_gst(ws, F, cfg):
    gst = cfg.get("gstin") or next(iter(sorted({r["gstin"] for r in F["g1"] + F["g3"]})), "")
    st = GST_STATES.get(gst[:2], "")
    ws["B1"] = "GSTIN:  %s(%s)" % (gst, st)
    for c, v in zip("BCDEF", ("Financial Year", "Tax Period", "Due Date", "Date of filing", "Remarks")):
        ws["%s2" % c] = v
    _ret_rows(ws, [r for r in F["g3"] if not gst or r["gstin"] == gst], "B", 3, "3B", 6)
    for c, v in zip("JKLMN", ("Particulars", "Month", "Amount (Rs.)", "Date Paid", "Remarks")):
        ws["%s2" % c] = v; ws["%s10" % c] = v
    # GST challans: no rule reads them yet - the month rows are left for the auditor
    for c, v in zip("BCDEF", ("Particulars", "Month", "Amount (Rs.)", "Date Paid", "Remarks")):
        ws["%s10" % c] = v
    _due_rows(ws, F.get("pf", []), "B", 11, "PF", 6)
    _due_rows(ws, F.get("esic", []), "J", 11, "ESIC", 6)
    for c, v in zip("BCDEF", ("Particulars", "Tax Period", "Amount (Rs.)", "Date Paid", "Remarks")):
        ws["%s18" % c] = v
    # electricity: the payment of month M is printed on the bill of month M+1 ("Last Month Payment Details")
    ws["B24"] = "Payment Details of Electricity Bill (For last 12 month):"
    bills = {b["ym"]: b for b in F.get("elec", [])}
    if bills:
        last = max(bills)
        for blk, top in ((0, 25), (1, 29)):
            ws["B%d" % top] = "Month"; ws["B%d" % (top + 1)] = "Date"; ws["B%d" % (top + 2)] = "Amount"
            for i in range(6):
                ym = _ym_add(last, -(blk * 6 + i)); col = "CDEFGH"[i]
                ws["%s%d" % (col, top)] = _mlab(ym)
                nxt = bills.get(_ym_add(ym, 1))
                if nxt and nxt.get("prev_paid"):
                    ws["%s%d" % (col, top + 1)] = nxt["prev_paid"]; ws["%s%d" % (col, top + 2)] = nxt["prev_amount"]
    ws["J19"] = "GSTIN:  %s(%s) - GSTR-1" % (gst, st)
    for c, v in zip("JKLMN", ("Financial Year", "Tax Period", "Due Date", "Date of filing", "Remarks")):
        ws["%s20" % c] = v
    _ret_rows(ws, [r for r in F["g1"] if (not gst or r["gstin"] == gst) and r["ym"] <= cfg["sum_to"]], "J", 21, "1", 7)
    ws["J31"] = "Sr. No"; ws["K31"] = "DIN/PAN"; ws["L31"] = "Name of Directors"
    dins = F.get("dins") or [("", n) for n in F.get("directors", [])]
    for i, (din, nm) in enumerate(dins):
        ws["J%d" % (32 + i)] = str(i + 1); ws["K%d" % (32 + i)] = din; ws["L%d" % (32 + i)] = nm


def build_excel_draft(F, cfg, out):
    """Workbook in the user's own layout: summation, Insurance, GST (template = blank copy of their draft)."""
    allc = []
    for acc, tx in F["bank"].items():
        for r in classify(dedupe(tx), F, cfg["related"], cfg["own"]):
            r["acc"] = acc; allc.append(r)
    wb = openpyxl.load_workbook(TEMPLATE)
    draft_summation(wb["summation"], F, cfg, allc)
    draft_insurance(wb["Insurance"], F, cfg)
    draft_gst(wb["GST"], F, cfg)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, dt.date):
                    c.number_format = "DD-MM-YYYY"
    wb.save(out)
    return allc


def default_period(F):
    """Summation period: six months ending at the latest stock-statement month (the drawing-power base)."""
    ss = [s["date"] for s in F["ss"] if s.get("date")]
    end = max(ss) if ss else max((t["date"] for tx in F["bank"].values() for t in tx), default=dt.date.today())
    to = (end.year, end.month)
    return _ym_add(to, -5), to


# ---------------- browser (GitHub Pages / Pyodide) entry points ----------------
_B = {}


def browser_scan(root):
    """Read every file under root (files written into the browser's memory). Returns a JSON summary for the review page."""
    paths = sorted(os.path.join(r, f) for r, _, fs in os.walk(root) for f in fs)
    log = []
    F = scan(paths, log)
    _B.clear(); _B.update(F=F, log=log, paths=paths)
    return browser_summary()


def browser_summary():
    F, log = _B["F"], _B["log"]
    ocr_done = {b for a, b, c in log if a.startswith("OCR")}
    unk = [(b, c) for a, b, c in log if a == "skipped" and b not in ocr_done and b.lower().rsplit(".", 1)[-1] in ("pdf", "xlsx", "xlsm", "xls", "csv")]
    unk += [(p["file"], "policy read, but stock SI not in the text layer (scanned part) - type it in the policy table")
            for p in F["policies"] if p["stock"] is None and p["cls"] != "Other"]
    gst = sorted({r["gstin"] for r in F["g1"] + F["g3"]})
    return json.dumps(dict(
        banks=[[k, len(v)] for k, v in F["bank"].items()], g1=len(F["g1"]), g3=len(F["g3"]), gstins=gst,
        ss=len(F["ss"]), book=len(F["book"]), npol=len(F["policies"]), debtors=len(F["debtors"]), creditors=len(F["creditors"]),
        directors=F["directors"], borrower=next(iter(sorted(n for n in F["names"] if n)), ""), books_sales=len(F.get("sales", {})),
        period=[dt.date(y, m, 1).strftime("%b-%Y") for y, m in default_period(F)],
        collateral=round(F["collateral"]["value"] / 1e5, 5) if F.get("collateral") else None,
        policies=[dict(no=p["no"], cls=p["cls"], frm=p["frm"], to=p["to"], stock=p["stock"], fl=p["first_loss"], build=p["build"],
                       file=p["file"], ocr=bool(p.get("ocr"))) for p in F["policies"]],
        unread=unk, log=log))


# ---------------- OCR of scanned pages (done in the browser by tesseract.js; parsed here by fixed rules) ----------------
OCR_MAX_PAGES = 3
OCR_FIELDS = ["First Loss Percentage"]        # boxed cells read separately (digits only) - JS crops right of the label


def ocr_candidates():
    """Scanned PDFs (no text layer, <= OCR_MAX_PAGES pages) and policies whose stock SI is in a scanned part."""
    F, log = _B["F"], _B["log"]
    by_name = {}
    for pth in _B["paths"]:
        by_name.setdefault(os.path.basename(pth), pth)
    want = [b for a, b, c in log if a == "skipped" and "no text layer" in c]
    want += [p["file"] for p in F["policies"] if p["cls"] != "Other" and (p["stock"] is None or (p["cls"] == "Burglary" and p["first_loss"] is None))]
    out = []
    for fn in dict.fromkeys(want):
        pth = by_name.get(fn)
        if not pth:
            continue
        try:
            with pymupdf.open(pth) as doc:
                n = len(doc)
        except Exception:
            continue
        if n <= OCR_MAX_PAGES:
            out.append(dict(path=pth, file=fn, pages=n))
    return json.dumps(dict(files=out, fields=OCR_FIELDS))


def ocr_render(path, page, out):
    """Render one page as a 300 dpi grey PNG for OCR."""
    with pymupdf.open(path) as doc:
        open(out, "wb").write(doc[page].get_pixmap(dpi=300, colorspace=pymupdf.csGRAY).tobytes("png"))


def _ocr_num(s):
    s = s.replace("O", "0").replace("o", "0").replace(",", "")
    return float(s)


def ocr_merge(path, text, fields_json):
    """Apply fixed rules to OCR text of one file. Every value found is marked 'OCR - check'."""
    F, log = _B["F"], _B["log"]; fn = os.path.basename(path)
    fields = json.loads(fields_json or "{}")
    T = re.sub(r"[ \t]+", " ", text)
    got = []
    # 1) the policy already read from this same file's text layer, 2) an existing policy whose number appears in the OCR text,
    # 3) only then a new policy from the OCR'd number (OCR can add stray digits from table borders)
    p = next((x for x in F["policies"] if x["file"] == fn), None) or         next((x for x in F["policies"] if x["no"] in re.sub(r"\D", "", T)), None)
    no = re.search(r"Policy\s*(?:Number|No\.?)\s*[\[\]|:;!]*\s*(\d{15,})", T)
    if p is not None or no:
        if p is None:
            per = re.search(r"From:?\s*(\d{2}/\d{2}/\d{4}).{0,120}?(\d{2}/\d{2}/\d{4})", T, re.S)
            if per:
                p = dict(file=fn, no=no.group(1), frm=per.group(1), to=per.group(2), insured="", gstin="", prev="", insurer="",
                         first_loss=None, stock=None, build="", excluded="", sub=None,
                         cls="Burglary" if "BURGLARY" in T.upper() else "Fire")
                ins = re.search(r"Insured'?s?\s*Name\W*([A-Z/][A-Z0-9 ./&-]{4,})", T)
                g = re.search(r"\b(\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z])\b", T)
                p["insured"] = ins.group(1).strip() if ins else ""; p["gstin"] = g.group(1) if g else ""
                F["policies"].append(p); F["names"].add(p["insured"]); got.append("policy %s" % p["no"])
        if p is not None:
            if p["stock"] is None:
                m = re.search(r"STOCK\s*DETAILS.*?\n(.*?)\b(\d[\d,O]{5,})\b", T, re.S | re.I) if p["cls"] == "Burglary" else \
                    re.search(r"Stocks?[^\n]{0,60}?\b(\d[\d,O]{5,})\b", T, re.I)
                if m:
                    v = _ocr_num(m.groups()[-1])
                    p["stock"] = v / 1e5; p["ocr"] = True
                    p["build"] = "OCR-read - CHECK against the policy: '%s %s'" % ((m.group(1).strip()[:60] + " ") if p["cls"] == "Burglary" else "", m.groups()[-1])
                    got.append("stock SI %.2f lakh" % p["stock"])
            if p["cls"] == "Burglary" and p["first_loss"] is None:
                fl = re.search(r"(\d{1,3})\s*%", fields.get("First Loss Percentage", ""))
                if fl:
                    p["first_loss"] = int(fl.group(1)); p["ocr"] = True; p["build"] += " | first loss %s%% by OCR - check" % fl.group(1)
                    got.append("first loss %s%%" % fl.group(1))
    if not got:
        ss = stock_statement(text) or company_stock_statement(text)
        if ss:
            ss["check"] = True; F["ss"].append(ss)
            got.append("stock statement %s: stock %s lakh" % (ss["date"], ss["stock"]))
    log.append(("OCR", fn, ("read: " + ", ".join(got) + " - CHECK these against the scan") if got else "OCR done - no policy / statement figures found"))
    return browser_summary()


def browser_make(form_json, out):
    """Apply the review-page entries and write the Excel to out."""
    f = json.loads(form_json); F = _B["F"]
    lines = lambda s: [x.strip() for x in (s or "").splitlines() if x.strip()]
    for p in F["policies"]:
        e = f.get("pol", {}).get(p["no"])
        if not e:
            continue
        if str(e.get("si", "")).strip():
            v = num(e["si"])
            if p["stock"] != v: p["build"] += " | SI entered on portal: %s" % e["si"]
            p["stock"] = v
        p["first_loss"] = int(num(e["fl"])) if str(e.get("fl", "")).strip() else None
    extra = []
    for l in lines(f.get("extra")):
        parts = [x.strip() for x in l.split(",")]
        if d(parts[0]):
            extra.append(dict(date=d(parts[0]), label=parts[1] if len(parts) > 1 and parts[1] else None,
                              stock=num(parts[2]) if len(parts) > 2 and parts[2] else None))
    own = lines(f.get("own"))
    banks = []
    for l in lines(f.get("banks")):
        parts = [x.strip() for x in re.split(r"[,:]", l)]
        if parts and parts[0]:
            banks.append((parts[0], num(parts[1]) if len(parts) > 1 and parts[1] else None))
    fr, to = default_period(F)
    pf = lambda k, dflt: (lambda m: (int(m.group(2)) + (2000 if int(m.group(2)) < 100 else 0), MON[m.group(1)[:3].lower()]) if m else dflt)(
        re.match(r"\s*([A-Za-z]{3,9})[\s'-]*(\d{2,4})", f.get(k) or ""))
    fl = lambda k: num(f[k]) if str(f.get(k) or "").strip() else None
    cfg = dict(own=own, borrower=own[0] if own else "", gstin=(f.get("gstin") or "").strip(), related=lines(f.get("related")), extra=extra,
               banks=banks, sum_from=pf("from", fr), sum_to=pf("to", to), pm_value=fl("pm"), collateral_value=fl("coll"),
               visit=extra[0]["date"] if extra else None, visit_stock=extra[0]["stock"] if extra else None)
    build_excel_draft(F, cfg, out)
    return re.sub(r"[^\w .-]", "", cfg["borrower"] or "Borrower")[:60].strip()


# Month-wise GST taxable-value summary added after the existing user-format workbook is built.
def _add_gst_taxable_summary(wb, F):
    rows_by_gstin = collections.defaultdict(dict)
    invalid = []
    duplicates = 0
    for r in F.get("g1", []):
        gstin = str(r.get("gstin") or "").strip().upper()
        if not re.fullmatch(r"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]Z[0-9A-Z]", gstin):
            invalid.append((gstin or "(blank)", r.get("ym"), r.get("arn", "")))
            continue
        ym = r.get("ym")
        if not ym or not isinstance(r.get("taxable"), (int, float)):
            invalid.append((gstin, ym, "Missing period or taxable value"))
            continue
        previous = rows_by_gstin[gstin].get(ym)
        if previous is None:
            rows_by_gstin[gstin][ym] = r
        else:
            # A duplicate file for this return period counts once; keep the latest ARN copy.
            duplicates += 1
            old_arn = d(previous.get("arn", "")) if previous.get("arn") else None
            new_arn = d(r.get("arn", "")) if r.get("arn") else None
            if new_arn and (not old_arn or new_arn > old_arn):
                rows_by_gstin[gstin][ym] = r
    months = sorted({ym for by_month in rows_by_gstin.values() for ym in by_month},
                    key=lambda ym: (ym[0] if ym[1] >= 4 else ym[0] - 1, (ym[1] - 4) % 12))
    ws = wb.create_sheet("GST Taxable Summary")
    headers = ["GSTIN"] + [_mname(ym) + "-" + str(ym[0]) for ym in months] + ["Total"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="173B38")
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
    for gstin in sorted(rows_by_gstin):
        values = [round(float(rows_by_gstin[gstin][ym]["taxable"]), 2) if ym in rows_by_gstin[gstin] else 0.0 for ym in months]
        ws.append([gstin] + values + [round(sum(values), 2)])
    grand = [round(sum(float(by_month[ym]["taxable"]) for by_month in rows_by_gstin.values() if ym in by_month), 2) for ym in months]
    ws.append(["Grand Total"] + grand + [round(sum(grand), 2)])
    for cell in ws[ws.max_row]:
        cell.font = openpyxl.styles.Font(bold=True, color="173B38")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="EAF4F1")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        width = min(22, max(14, max(len(str(cell.value or "")) for cell in col) + 2))
        ws.column_dimensions[col[0].column_letter].width = width
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.number_format = "#,##0.00"
    ws.append(["Basis: filed GSTR-1 taxable value by return tax period. Reported debit/credit notes and amendments are included in each filed-period net amount; tax columns are excluded."])
    ws.cell(ws.max_row, 1).font = openpyxl.styles.Font(italic=True, color="66736D")
    ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=max(1, len(headers)))
    review = wb.create_sheet("GST Summary Review")
    review.append(["Check", "Count", "Details"])
    review.append(["Duplicate GSTIN-period return skipped", duplicates, "Each return period is included once; latest ARN copy retained when available."])
    review.append(["Invalid or incomplete GSTR-1 return", len(invalid), "; ".join("%s %s %s" % item for item in invalid)[:30000]])
    for cell in review[1]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="173B38")
    review.column_dimensions["A"].width = 38
    review.column_dimensions["B"].width = 14
    review.column_dimensions["C"].width = 100
    review.freeze_panes = "A2"


_build_excel_draft_without_taxable_summary = build_excel_draft
def build_excel_draft(F, cfg, out):
    allc = _build_excel_draft_without_taxable_summary(F, cfg, out)
    wb = openpyxl.load_workbook(out)
    _add_gst_taxable_summary(wb, F)
    wb.save(out)
    return allc
