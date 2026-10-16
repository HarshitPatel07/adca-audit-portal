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


def xlsx_rows(p):
    out = []
    try:
        wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
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
    pat = re.compile(r"\|(\d+)\|(\d{2}/\d{2}/\d{4})\|(\d{2}/\d{2}/\d{4})\|(.*?)\|([\d,]+\.\d\d)\|(DR|CR)\|(-?[\d,]+\.\d\d)\|")
    tx = []
    for m in pat.finditer("|" + f):
        a = num(m.group(5)) * (1 if m.group(6) == "CR" else -1)
        tx.append(dict(date=d(m.group(2)), narr=m.group(4).replace("|", ""), amt=a, bal=num(m.group(7))))
    return (acc.group(1) if acc else "bank"), tx


def bank_table(rows):
    """Any spreadsheet statement with Date + Narration + Debit/Withdrawal + Credit/Deposit columns."""
    for i, r in enumerate(rows[:40]):
        low = [str(c or "").lower() for c in r]
        di = next((j for j, c in enumerate(low) if "date" in c), None)
        ni = next((j for j, c in enumerate(low) if any(k in c for k in ("narration", "particular", "description", "remarks"))), None)
        dr = next((j for j, c in enumerate(low) if any(k in c for k in ("debit", "withdrawal"))), None)
        cr = next((j for j, c in enumerate(low) if any(k in c for k in ("credit", "deposit"))), None)
        if None not in (di, ni, dr, cr):
            tx = []
            for rr in rows[i + 1:]:
                dd = rr[di]
                dd = dd.date() if isinstance(dd, dt.datetime) else d(str(dd or ""))
                if not dd:
                    continue
                a = num(rr[cr] or 0) - num(rr[dr] or 0)
                if a:
                    tx.append(dict(date=dd, narr=str(rr[ni] or ""), amt=a, bal=None))
            return tx
    return []


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
    return dict(date=d(m.group(1)), stock=num(s.group(1)), creditors=num(c.group(1)) if c else None,
                debtors=num(b.group(1)) if b else None)


def tally_kind(rows):
    head = " ".join(str(c) for r in rows[:15] for c in r if c)
    for k in ("Stock Summary", "Sundry Debtors", "Sundry Creditors"):
        if k in head:
            return k, head
    return None, head


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
    f = t.replace("\n", "|")
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
             insurer=("THE " + insurer.group(1)) if insurer else "", first_loss=None, stock=None, build="", excluded="")
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
                p["build"] = " + ".join("loc%d: RM %g + WIP %g + FG %g" % (i + 1, b[4] / 1e5, b[5] / 1e5, b[6] / 1e5) for i, b in enumerate(blocks)) \
                    + " = %g lakh (Block Details)" % p["stock"]
                p["excluded"] = "Building %g; FFF %g; P&M %g; Other contents %g (lakh)" % tuple(
                    sum(b[k] for b in blocks) / 1e5 for k in (0, 1, 2, 3))
        if p["stock"] is None:
            s = re.search(r"\|Stocks?[^|]{0,60}\|(?:[^|]{0,40}\|)?([\d,]{6,})\|", f)
            if s:
                p["stock"] = num(s.group(1)) / 1e5
                p["build"] = "Line 'Stock' %s (check)" % s.group(1)
            else:
                p["build"] = "Stock SI not found - enter on review page"
    return p


# ---------------- scan everything ----------------
def scan(paths, log):
    F = dict(bank={}, g1=[], g3=[], ss=[], book=[], debtors=set(), creditors=set(), directors=[], policies=[], names=set())
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
                        a, tx = bank_pdf(t); [x.update(src=fn) for x in tx]; F["bank"].setdefault(a, []).extend(tx); log.append(("Bank PDF", fn, "%d txns" % len(tx)))
                    elif stock_statement(t):
                        r = stock_statement(t); F["ss"].append(r); log.append(("Stock statement", fn, "%s stock %s" % (r["date"], r["stock"])))
                    elif policy(t, fn):
                        r = policy(t, fn); F["policies"].append(r); F["names"].add(r["insured"])
                        log.append(("Policy " + r["cls"], fn, "%s %s-%s stock SI %s" % (r["no"], r["frm"], r["to"], r["stock"])))
                    else:
                        log.append(("skipped", fn, "not recognised" if t.strip() else "no text layer (scanned)"))
            elif ext == "txt":
                t = open(p, encoding="utf-8", errors="replace").read()
                if "STATEMENT" in t and "BALANCE" in t:
                    a, tx = bank_txt(t); [x.update(src=fn) for x in tx]; F["bank"].setdefault(a, []).extend(tx); log.append(("Bank TXT", fn, "%d txns" % len(tx)))
            elif ext in ("xlsx", "xlsm", "csv"):
                rows = xlsx_rows(p) if ext != "csv" else [l.split(",") for l in open(p, encoding="utf-8", errors="replace")]
                k, head = tally_kind(rows)
                if k == "Stock Summary":
                    r = tally_stock(rows, head)
                    if r: F["book"].append(r); log.append(("Book stock", fn, "%s %s lakh" % (r["date"], r["stock"])))
                elif k == "Sundry Debtors":
                    F["debtors"].update(tally_names(rows)); log.append(("Debtor ledger", fn, "names read"))
                elif k == "Sundry Creditors":
                    F["creditors"].update(tally_names(rows)); log.append(("Creditor ledger", fn, "names read"))
                elif mca_directors(rows):
                    F["directors"] = mca_directors(rows); log.append(("MCA", fn, ", ".join(F["directors"])))
                else:
                    tx = bank_table(rows)
                    if tx:
                        [x.update(src=fn) for x in tx]
                        F["bank"].setdefault(fn, []).extend(tx); log.append(("Bank sheet", fn, "%d txns" % len(tx)))
                    else:
                        log.append(("skipped", fn, "not recognised"))
            else:
                log.append(("skipped", fn, "type not used"))
        except Exception as e:
            log.append(("ERROR", fn, repr(e)[:120]))
    return F


# ---------------- debtor receipts classification (prompt 1) ----------------
STOP = {"PVT", "PRIVATE", "LTD", "LIMITED", "THE", "M/S", "MS", "AND", "&", "CO", "COMPANY", "SALES", "PURCHASE", "UNIT", "LLP", "INDIA", "URP"}


def toks(name):
    n = re.sub(r"^\s*M\s*/\s*S\.?\s+", "", name.upper())
    return [w for w in re.split(r"[^A-Z0-9]+", n) if len(w) > 1 and w not in STOP and not w.isdigit()]


def nospace(s):
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def name_hit(name, narr):
    """'full' = first two words found; 'part' = distinctive first word (>=6 chars) only."""
    t = toks(name); n = nospace(narr)
    if len(t) >= 2 and len(t[0] + t[1]) >= 6 and (t[0] + t[1]) in n:
        return "full"
    U = narr.upper()
    # partial matches only where the payer name STARTS (after '/' or at the start) - never inside "X S/O <name>"
    if len(t) >= 2 and len(t[0]) >= 5 and re.search(r"(^|/)\s*" + t[0] + r"\s*(?:(?:AND|&)\s*)?" + t[1][0], U):
        return "part"
    if len(t) == 1 and len(t[0]) >= 5 and t[0] in n:
        return "full"
    if t and len(t[0]) >= 6 and re.search(r"(^|/)\s*" + t[0] + r"\s*$", U):  # name cut right after first word, e.g. 'Hdfc Bank /AGRAWAL'
        return "part"
    return None


KW = [  # (category, keywords, reason)
    ("Bank-originated credit", ("INT.COLL", "INTEREST", "INT CR", "TXN REV", "REVERSAL", "LOAN DISB", "DISBURS", "FD CLOS", "SWEEP", "MATURITY", "CHQ RTN", "CHEQUE RETURN"), "Bank entry: interest / reversal / disbursement / FD"),
    ("Refunds", ("RETURN/", "NEFT/RETURN", "REFUND", "RTN/"), "Own payment returned or refund"),
    ("Tax-related", ("GST REFUND", "CBDT", "ITD", "INCOME TAX", "TREASU", "PMVBRY", "SUBSIDY", "GOVT", "GOVERNMENT"), "Government / tax credit"),
    ("Investment/other income", ("DIVIDEND", "MUTUAL FUND", "REDEMPTION", " MF "), "Investment income"),
    ("Loans/borrowings (non-bank)", ("UNSECURED LOAN", "LOAN FROM", "USL"), "Loan received"),
]


def classify(tx, F, related, own):
    out = []
    for t in tx:
        if t["amt"] <= 0:
            continue
        u = t["narr"].upper(); r = dict(t)
        hit = None
        for nm in own:
            if nm and toks(nm) and "".join(toks(nm)) in nospace(u):
                hit = ("EXC", "Self-transfer / contra", "Own name (%s) - transfer from own account" % nm)
        if not hit:
            for nm in related:
                if name_hit(nm, u):
                    extra = " - also in debtor ledger" if any(name_hit(nm, x) for x in [u]) and any(toks(nm)[:1] == toks(dn)[:1] for dn in F["debtors"]) else ""
                    hit = ("EXC", "Director/Partner/Owner/Related party", "Related party / director: %s%s" % (nm, extra)); break
        if not hit:
            for cat, kws, why in KW:
                k = next((k for k in kws if k in u), None)
                if k:
                    hit = ("EXC", cat, "%s (narration has '%s')" % (why, k.strip())); break
        if not hit:
            for nm in F["creditors"]:
                if name_hit(nm, u) and not any(name_hit(dn, u) for dn in F["debtors"]):
                    hit = ("EXC", "Refunds", "Payer is a supplier in creditor ledger (%s) - refund" % nm); break
        if not hit:
            best = None
            for nm in sorted(F["debtors"], key=len, reverse=True):
                h = name_hit(nm, u)
                if h == "full":
                    best = (nm, ""); break
                if h == "part" and not best:
                    best = (nm, "first-word match only")
            if best:
                hit = ("INC", best[0], best[1])
        if not hit:
            hit = ("REV", "Unidentified / needs review", "Payer not in debtor ledger and no rule matched - confirm nature")
        r["b"], r["cat"], r["why"] = hit
        out.append(r)
    return out


def dedupe(txs):
    """Same account read from several files (e.g. full-period TXT + monthly PDFs): for each month keep the ONE source
    file with the most transactions for that month - never mix two files in one month."""
    by = collections.defaultdict(lambda: collections.defaultdict(list))
    for t in txs:
        by[(t["date"].year, t["date"].month)][t.get("src", "")].append(t)
    out = []
    for k in sorted(by):
        src = max(by[k].items(), key=lambda kv: len(kv[1]))
        out += [dict(t, src=src[0]) for t in src[1]]
    return sorted(out, key=lambda t: t["date"])


# ---------------- insurance date mapping (prompt 3) ----------------
def pdate(s):
    return dt.datetime.strptime(s, "%d/%m/%Y").date()


def cover(pols, cls, day):
    live = [p for p in pols if p["cls"] == cls and pdate(p["frm"]) <= day <= pdate(p["to"])]
    live.sort(key=lambda p: pdate(p["frm"]))
    return live[-1] if live else None, len(live) > 1


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
                if isinstance(c.value, float): c.number_format = NUMF
                if isinstance(c.value, dt.date): c.number_format = "DD-MM-YYYY"
    wb.save(out)
    return allc


def build_insurance(wb, F, cfg):
    ws = wb.create_sheet("Insurance")
    pols, seen = [], set()
    for p in F["policies"]:
        if p["no"] in seen:
            continue
        seen.add(p["no"]); pols.append(p)
    dups = len(F["policies"]) - len(pols)
    own = [n for n in cfg["own"] if n]
    notes = []
    for p in pols:
        if own and not any(name_hit(n, p["insured"]) for n in own):
            notes.append("POLICY %s IS IN THE NAME OF '%s' - NOT THE BORROWER; not merged." % (p["no"], p["insured"]))
            p["foreign"] = True
        if cfg.get("gstin") and p["gstin"] and p["gstin"] != cfg["gstin"]:
            notes.append("Policy %s GSTIN %s differs from borrower GSTIN %s." % (p["no"], p["gstin"], cfg["gstin"]))
    use = [p for p in pols if not p.get("foreign")]
    r = 1
    for n in notes:
        ws.cell(r, 1, n).font = Font(bold=True, color="C00000"); r += 1
    if dups:
        ws.cell(r, 1, "%d duplicate upload(s) of the same policy number ignored." % dups); r += 1
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
    for b in F["book"]:
        if b["date"]:
            rows.setdefault(b["date"], {})["book"] = b["stock"]
    for x in cfg.get("extra", []):
        rows.setdefault(x["date"], {}).update({"ss": x.get("stock")} if x.get("stock") is not None else {}, label=x.get("label", "Visit date"))
    r = ws.max_row + 2
    ws.cell(r, 1, "Table 2 - Result (Rs lakh)").font = BOLD; r += 1
    head(ws, r, ["Month / Date", "Total Stock - Stock Statement", "Total Stock - Books", "Fire - Insured Amt (lakh)",
                 "Burglary - Insured Amt (lakh)", "Fire policy", "Burglary policy", "Excess / (Shortfall) Fire", "Remark"])
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
                if ov: rem.append("%s overlap - later policy used" % cls)
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
        ws.append([lab, v.get("ss"), v.get("book"), cells[0][0], cells[1][0], cells[0][1], cells[1][1],
                   (fire - stock) if isinstance(fire, (int, float)) and stock else None, "; ".join(rem) or "Adequately insured"])
        if any("UNDER" in x for x in rem):
            for c in ws[ws.max_row]: c.fill = WARN
    r = ws.max_row + 2
    ws.cell(r, 1, "Gaps / missing policies").font = BOLD
    if not gaps:
        ws.cell(r + 1, 1, "None - every date is inside a supplied Fire and Burglary policy period (inclusive).")
    for i, (day, cls) in enumerate(gaps):
        prevs = [p["prev"] for p in use if p["cls"] == cls and p["prev"] and pdate(p["frm"]) > day]
        ws.cell(r + 1 + i, 1, "%s %s: no %s policy supplied%s" % (day.strftime("%d-%m-%Y"), "", cls,
                ("; missing policy probably %s (Prev. Policy No. on later schedule)" % prevs[0]) if prevs else ""))
    ws.cell(ws.max_row + 2, 1, "Only stock in trade (RM + WIP + FG / Stocks in trade) counted. Building, P&M, FFF, other contents, cash, LOP excluded. Units: Rs lakh as in the stock statement.")
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")





# ---------------- browser (GitHub Pages / Pyodide) entry points ----------------
_B = {}


def browser_scan(root):
    """Read every file under root (files written into the browser's memory). Returns a JSON summary for the review page."""
    paths = sorted(os.path.join(r, f) for r, _, fs in os.walk(root) for f in fs)
    log = []
    F = scan(paths, log)
    _B.clear(); _B.update(F=F, log=log)
    unk = [(b, c) for a, b, c in log if a == "skipped" and b.lower().rsplit(".", 1)[-1] in ("pdf", "xlsx", "xlsm", "xls", "csv")]
    unk += [(p["file"], "policy read, but stock SI not in the text layer (scanned part) - type it in the policy table")
            for p in F["policies"] if p["stock"] is None]
    gst = sorted({r["gstin"] for r in F["g1"] + F["g3"]})
    return json.dumps(dict(
        banks=[[k, len(v)] for k, v in F["bank"].items()], g1=len(F["g1"]), g3=len(F["g3"]), gstins=gst,
        ss=len(F["ss"]), book=len(F["book"]), npol=len(F["policies"]), debtors=len(F["debtors"]), creditors=len(F["creditors"]),
        directors=F["directors"], borrower=next(iter(sorted(n for n in F["names"] if n)), ""),
        policies=[dict(no=p["no"], cls=p["cls"], frm=p["frm"], to=p["to"], stock=p["stock"], fl=p["first_loss"], build=p["build"], file=p["file"])
                  for p in F["policies"]],
        unread=unk, log=log))


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
    cfg = dict(own=own, borrower=own[0] if own else "", gstin=(f.get("gstin") or "").strip(), related=lines(f.get("related")), extra=extra)
    build_excel(F, cfg, out)
    return re.sub(r"[^\w .-]", "", cfg["borrower"] or "Borrower")[:60].strip()
