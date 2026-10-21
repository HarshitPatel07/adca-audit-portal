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


PDF_FIX = str.maketrans({"\xa0": " ", " ": " ", " ": " ", "\xad": "-", ";": ";"})


def pdf_text(p):
    """Text layer of a PDF. Non-breaking spaces / soft hyphens are made plain (ICICI Lombard prints every space as
    U+00A0, so 'Policy Number' would never match)."""
    try:
        with pymupdf.open(p) as doc:
            return "".join(pg.get_text() for pg in doc).translate(PDF_FIX)
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
BANK_HEAD = [(r"A\s*X\s*I\s*S\s+B\s*A\s*N\s*K|Axis Bank", "Axis Bank"), (r"State Bank of India|\bSBI\b", "State Bank of India"),
             (r"HDFC Bank", "HDFC Bank"), (r"ICICI Bank", "ICICI Bank"), (r"Union Bank of India", "Union Bank of India"),
             (r"Bank of Baroda", "Bank of Baroda"), (r"Punjab National Bank", "Punjab National Bank"), (r"Canara Bank", "Canara Bank"),
             (r"Kotak Mahindra Bank", "Kotak Mahindra Bank"), (r"IndusInd Bank", "IndusInd Bank"), (r"Yes Bank", "Yes Bank"),
             (r"IDBI Bank", "IDBI Bank"), (r"Central Bank of India", "Central Bank of India"), (r"Bank of India", "Bank of India"),
             (r"Indian Overseas Bank", "Indian Overseas Bank"), (r"UCO Bank", "UCO Bank"), (r"Federal Bank", "Federal Bank")]


def bank_of(t):
    """Bank that issued the statement, from its header only (narrations name other banks: 'ACH ICICI BANK-...')."""
    head = t[:600] if "STATEMENT" in t[:1500] else t[:1500]
    head = head.split("STATEMENT  FOR")[0] if "STATEMENT  FOR" in head else head
    return next((nm for rx, nm in BANK_HEAD if re.search(rx, head, re.I)), "")


def sheet_head(rows):
    """Rows above a spreadsheet statement's column headings (the bank's own header; narrations below name other banks)."""
    heads = ("narration", "particulars", "description", "transaction details", "remarks")
    return next((i for i, r in enumerate(rows[:40]) if any(str(c or "").lower().strip() in heads for c in r)), 0)


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
    # 3.1(b) zero rated (exports / SEZ) and 3.1(c) nil rated / exempted: the drafts' summation sales = (a) + (b) + (c)
    # (Baid Agro, an exporter: May'25-Feb'26 = 5087.75 lakh; (a) alone is 877 lakh)
    zb = re.search(r"\(b\) Outward taxable supplies \(zero rated\)\|([\d.,-]+)", f)
    nc = re.search(r"\(c ?\) Other outward supplies.*?exempted\)\|([\d.,-]+)", f)
    ta, tb, tc = num(a.group(1)), num(zb.group(1)) if zb else 0.0, num(nc.group(1)) if nc else 0.0
    return dict(kind="3B", gstin=g.group(1), ym=(y, m), taxable=round(ta + tb, 2), sales=round(ta + tb + tc, 2),
                arn=arn.group(1) if arn else "")


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


# ---------------- outward supply, taxable value only (the GST master prompt) ----------------
GSTIN_RX = r"\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]"
OUT_YES = r"outward|\bsales?\b|gstr[-\s_]*1(?!\d)|\bR1\b|\bb2b\b|\bb2c[ls]?\b|\bexp(?:ort)?s?\b|sr\d|dr[-\s/]*cr|credit\s*note|debit\s*note"
OUT_NO = r"purchase|inward|gstr[-\s_]*2|gstr[-\s_]*3|\b2a\b|\b2b\b|\bitc\b|rcm|reverse charge|hsn|advance|documents issued"
NOT_TAXABLE = r"nil\s*rated|exempt|non[-\s]*gst"


def sheet_tables(p):
    """[(sheet name, rows)] of an Excel / CSV file: .xlsx/.xlsm (openpyxl), .xls (xlrd), .csv. Trailing blank rows of a
    sheet formatted to a million rows are cut after 300 empty rows."""
    ext = p.lower().rsplit(".", 1)[-1]
    out = []
    try:
        if ext == "csv":
            import csv
            with open(p, encoding="utf-8", errors="replace", newline="") as fh:
                out.append((os.path.basename(p), [r for r in csv.reader(fh)]))
        elif ext == "xls":
            import xlrd
            wb = xlrd.open_workbook(p)
            for sh in wb.sheets():
                rows = []
                for i in range(sh.nrows):
                    r = []
                    for c in sh.row(i):
                        v = c.value
                        if c.ctype == xlrd.XL_CELL_DATE:
                            try:
                                v = xlrd.xldate.xldate_as_datetime(v, wb.datemode)
                            except Exception:
                                pass
                        r.append(None if v == "" else v)
                    rows.append(r)
                out.append((sh.name, rows))
        else:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            for ws in wb.worksheets:
                rows, blank = [], 0
                for r in ws.iter_rows(values_only=True):
                    if any(c is not None and str(c).strip() for c in r):
                        rows.append(list(r)); blank = 0
                    else:
                        blank += 1
                        if blank > 300:
                            break
                out.append((ws.title, rows))
            wb.close()
    except Exception:
        return []
    return out


def outward_file(F, p, log):
    """Read one outward-supply workbook into F['out'] (rows) and F['out_notes']; True if it held any."""
    rows, notes = outward_rows(p, sheet_tables(p))
    if not rows and not any("cancelled" in n for n in notes):
        return False
    F.setdefault("out", []).extend(rows); F.setdefault("out_notes", []).extend(notes)
    tax = sum(r["value"] for r in rows if not r["nontax"])
    log.append(("GST outward", os.path.basename(p), "%d rows, taxable %s (%s)" % (len(rows), "{:,.2f}".format(tax),
                                                                                "; ".join(n.split(": ", 1)[-1] for n in notes)[:200])))
    return True


def _hdr(c):
    return " ".join(str(c or "").replace("\n", " ").split()).lower()


def _cell_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v or "").strip()
    for f in ("%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d-%b-%y", "%d.%m.%Y", "%Y-%m-%d", "%d-%m-%y", "%d/%m/%y", "%d %b %Y"):
        try:
            return dt.datetime.strptime(s[:11].strip(), f).date()
        except ValueError:
            pass
    return None


def outward_rows(p, sheets, own_gstin=""):
    """Rows of outward-supply data (invoice / line / note level) found in a workbook, per the master prompt:
    taxable value only (tax and invoice-value columns never read), month from the document date (a note's own date),
    credit note minus / debit note plus, nil / exempt / non-GST tables kept apart. Returns (rows, notes)."""
    fn = os.path.basename(p)
    if re.search(OUT_NO, fn, re.I) or not re.search(OUT_YES, fn, re.I):
        name_ok = False
    else:
        name_ok = True
    fgst = re.search(GSTIN_RX, fn)
    out, notes = [], []
    for sname, rows in sheets:
        top = " ".join(_hdr(c) for r in rows[:4] for c in r if c)
        title = sname.lower() + " " + top
        if re.search(OUT_NO, sname, re.I) or re.search(r"\binward\b|purchase|gstr[-\s_]*2[ab]\b|\b2[ab]\b", top):
            continue
        if not (name_ok or re.search(OUT_YES, title, re.I)):
            continue
        for hi, r in enumerate(rows[:20]):
            h = [_hdr(c) for c in r]
            tv = next((j for j, c in enumerate(h) if re.fullmatch(r"(taxable|assessable)\s*(value|amount|amt)( \(.*\))?|value \(before tax\)|taxable value \(₹\)", c)), None)
            if tv is None:
                continue
            dates = [j for j, c in enumerate(h) if "date" in c and "shipping" not in c and "port" not in c]
            note_d = [j for j in dates if re.search(r"note|voucher|document|dr/ cr|dr/cr|refund", h[j])]
            di = (note_d or dates or [None])[0]
            per = next((j for j, c in enumerate(h) if re.search(r"return period|tax period|\bperiod\b|\bmonth\b", c)), None)
            # the note's OWN number - never 'Original Invoice / Debit Note No.' (an older invoice, often last year's)
            no = next((j for j, c in enumerate(h) if re.search(r"(note|voucher)\s*(no|number)|dr/ ?cr note no", c)
                       and not c.startswith("original")), None) if note_d else None
            no = no if no is not None else next((j for j, c in enumerate(h) if re.search(r"invoice\s*(no|number)|inv\.?\s*no|voucher\s*no|bill\s*no|document\s*(no|number)", c)), None)
            line = next((j for j, c in enumerate(h) if "line" in c and ("sr" in c or "no" in c)), None)
            dtyp = next((j for j, c in enumerate(h) if re.search(r"document\s*type|voucher\s*type|note\s*type|doc\.?\s*type|^type of note", c)), None)
            gi = next((j for j, c in enumerate(h) if re.search(r"gstin|gst\s*no|registration", c) and not re.search(r"recip|receip|buyer|party|customer|counter", c)), None)
            nontax = bool(re.search(NOT_TAXABLE, title, re.I))
            kind = "note" if note_d or re.search(r"credit|debit|cdn|dr[-\s/]*cr", title, re.I) else "invoice"
            used = skip_blank = 0
            cancel_from, cancelled = None, []
            for k, rr in enumerate(rows[hi + 1:], hi + 2):
                g = lambda j: rr[j] if j is not None and j < len(rr) else None
                txt = " ".join(str(c) for c in rr if isinstance(c, str))
                # 'following Invoice Cancel' (a text-only line): every row below it is a cancelled invoice (Denis Feb-26:
                # 7 invoices, 15,43,541.50 - exactly what the filed GSTR-1 leaves out); '(cancelled )' rows likewise
                filled = [c for c in rr if c is not None and str(c).strip()]
                if cancel_from is None and re.search(r"cancel", txt, re.I) and not re.search(r"\(\s*cancel", txt, re.I) \
                        and len(filled) <= 3 and all(isinstance(c, str) for c in filled):
                    cancel_from = k; continue
                if cancel_from is not None or re.search(r"\(\s*cancel", txt, re.I):
                    cv = g(tv)
                    cn = str(g(no) or "").strip()
                    if not re.search(r"\d", cn) or re.match(r"\d{2}-[A-Za-z]", cn):      # columns shift in this section
                        cn = next((str(c).strip() for c in rr if isinstance(c, str) and re.search(r"\d", c) and re.search(r"[/-]|^[A-Z]{2}\d", c.strip())
                                   and not re.fullmatch(GSTIN_RX, c.strip()) and not re.match(r"\d{2}-[A-Za-z]", c.strip())), "")
                    if isinstance(cv, (int, float)) and cv and cn:          # a document, not the section's sub-total line
                        cancelled.append((cn, cv))
                    continue
                v = g(tv)
                if isinstance(v, str):
                    try:
                        v = float(v.replace(",", "").strip())
                    except ValueError:
                        v = None
                inv = str(g(no) or "").strip()
                day = _cell_date(g(di)) if di is not None else None
                if day is None and per is not None:
                    pv = g(per); pd_ = _cell_date(pv) or (lambda m: dt.date(int(m.group(2)) + (2000 if len(m.group(2)) == 2 else 0), MON[m.group(1)[:3].lower()], 1) if m else None)(
                        re.match(r"\s*([A-Za-z]{3,9})[\s'-]*(\d{2,4})", str(pv or "")))
                    day = pd_
                if not inv and day is None:
                    continue                                  # a total / sub-total line, not a document
                fixed = None
                fym = re.search(r"(?<!\d)(?:20)?(\d{2})-(\d{2})(?!\d)", inv) or re.search(r"\b[A-Z]{1,4}[-/](\d{2})(\d{2})[-/]", inv)
                if day and fym and int(fym.group(2)) == (int(fym.group(1)) + 1) % 100:
                    fy0 = 2000 + int(fym.group(1))            # 'EXP061/2025-26' dated 2025-02-xx: year typed wrong
                    if not dt.date(fy0, 4, 1) <= day <= dt.date(fy0 + 1, 3, 31):
                        try:
                            nd = day.replace(year=fy0 if day.month >= 4 else fy0 + 1)
                        except ValueError:
                            nd = None
                        if nd:
                            fixed, day = day, nd
                if v is None or not isinstance(v, (int, float)):
                    skip_blank += 1; continue
                if day is None:
                    skip_blank += 1; continue
                t = str(g(dtyp) or "").strip().lower()
                doc = "invoice"
                if re.match(r"c\b|credit|cr\b|cn\b", t):
                    doc = "credit"
                elif re.match(r"d\b|debit|dr\b|dn\b", t):
                    doc = "debit"
                elif re.match(r"r\b|refund", t):
                    notes.append("%s / %s row %d: refund voucher - not a supply, left out" % (fn, sname, k)); continue
                elif kind == "note":
                    has_c, has_d = bool(re.search(r"credit", title)), bool(re.search(r"debit", title))
                    if has_c != has_d:                # a credit-note-only or debit-note-only table
                        doc = "credit" if has_c else "debit"
                    else:                             # a notes table with no D / C column: not guessed
                        notes.append("%s / %s row %d: note without a D / C type - left out, CHECK" % (fn, sname, k)); continue
                gst = str(g(gi) or "").strip().upper() if gi is not None else ""
                gst = gst or (fgst.group(0) if fgst else "") or own_gstin
                # credit note always reduces, debit note always adds, whatever sign the sheet prints (Denis Oct-25 prints
                # debit note DN/033 as -4,000 inside the notes table; the filed GSTR-1 counts it +4,000)
                val = -abs(v) if doc == "credit" else (abs(v) if doc == "debit" else v)
                out.append(dict(gstin=gst, ym=(day.year, day.month), value=round(val, 2), no=inv,
                                day=day, line=str(g(line) or ""), kind=kind, nontax=nontax, src="%s / %s" % (fn, sname.strip()), row=k,
                                fixed=fixed, doc=doc))
                used += 1
            if used or skip_blank:
                notes.append("%s / %s: %d row(s) read%s%s" % (fn, sname.strip(), used, ", %d skipped (blank / non-numeric taxable value or no date)" % skip_blank if skip_blank else "",
                                                              " - nil / exempt / non-GST, NOT taxable value" if nontax else ""))
            if cancelled:
                notes.append("%s / %s: %d cancelled invoice line(s) left out, taxable %s (%s)" % (
                    fn, sname.strip(), len(cancelled), "{:,.2f}".format(sum(v for _, v in cancelled)), ", ".join(n for n, _ in cancelled[:12])))
            break
    return out, notes


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


MONTH_WORDS = {dt.date(2000, i, 1).strftime(f).lower() for i in range(1, 13) for f in ("%B", "%b")} | {"sept"}


def tally_names(rows):
    out = []
    for r in rows:
        n = str(r[0] or "").strip()
        if n and not any(n.lower().startswith(s) for s in SKIP) and any(isinstance(c, (int, float)) for c in r[1:]) \
                and not re.match(r"^[\d.,-]+$", n) and n.lower() not in MONTH_WORDS:
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
    # product name as the schedule prints it ('ICICI LOMBARD MSME SURAKSHA KAVACH', 'FLEXI COMMERCIAL PROPERTY GUARD')
    m = re.search(r"Undertaking\)\|([^|]{6,90})", f) or \
        re.search(r"(FLEXI COMMERCIAL PROPERTY GUARD)", f) or \
        re.search(r"\|((?:ICICI LOMBARD )?MSME SURAKSHA KAVACH|Business Guard Sookshma[^|]{0,40}"
                  r"|[A-Z][A-Za-z ]{0,30}Sookshma Udyam[^|]{0,30}|[A-Z][A-Za-z ]{0,30}Laghu Udyam[^|]{0,30})\s*(?:\(|\|)", f) or \
        re.search(r"(INDUSTRIAL ALL RISK[^|]{0,40}|BURGLARY[^|]{0,40}INSURANCE[^|]{0,20}|STANDARD FIRE[^|]{0,40}|BHARAT [^|]{0,40})", f, re.I)
    p["nature"] = re.sub(r"\s*(?:\(UIN\)|:\s*IRDA|UIN\s*:).*$", "", m.group(1)).strip() if m else p["cls"]   # drop '(UIN) : IRDAN...'
    # financier: known bank names after the financier / hypothecation label (several banks are joined)
    banks = []
    for mk in re.finditer(r"Financier|Financial Institute Name|Name of the Pledgee|Hypothecation|Bank Details|HP-|Hypo\b|Agreed Bank|Bank Clause|Interest of", f, re.I):
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
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%d-%b-%Y", "%d-%b-%y", "%b %d, %Y", "%B %d, %Y", "%d %b %Y", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(s.title() if "%b" in f else s, f).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return None


DATE = r"(\d{2}[/.-](?:\d{2}|[A-Za-z]{3})[/.-]\d{2,4}|[A-Z][a-z]{2,8} \d{1,2}, \d{4})"


def _period(f):
    i = f.find("Period of Insurance")
    seg = f[i:i + 400] if i >= 0 else f
    # 'From : 00:00 Hours of 21/04/2025 To : 23:59 Hours of 20/04/2026' (ICICI) / 'To Midnight of ...'
    TM = r"(?:\d{2}:\d{2}(?::\d{2})?\s*(?:Hours?\s*of\s*)?|Midnight\s*of\s*)?"
    rx = r"From\W{0,6}" + TM + DATE + r".{0,80}?To\W{0,6}" + TM + DATE
    m = re.search(rx, seg, re.S)
    if not m and i >= 0:
        m = re.search(rx, f, re.S)
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
    # 1b) first-loss table (ICICI burglary): 'S No.|Sum Insured Component|First Loss Sum Insured|Balance|Total Sum Insured'
    #     each row 'Stocks|Rs. 60,000,000.0|0|Rs. 90,000,000.0|0|Rs. 150,000,000.|00' -> the TOTAL column is the SI
    i = f.find("Sum Insured Component")
    if not found and i >= 0 and re.search(r"First Loss Sum", f[i:i + 200]):
        seg = re.sub(r"(\.\d?)\|(\d{1,2})(?=\|)", r"\1\2", f[i:i + 2000])      # '150,000,000.|00|' -> '150,000,000.00|'
        rows = re.findall(r"(?<=\|)\d+\.\|([A-Za-z][^|]{1,60})\|((?:Rs\.?\s*[\d,]+(?:\.\d+)?\|?){2,3})", seg)
        for lab, amts in rows:
            vals = [num(x) for x in re.findall(r"Rs\.?\s*([\d,]+(?:\.\d+)?)", amts)]
            p.setdefault("items", []).append((lab.strip(), vals[-1] / 1e5))
            if re.match(r"Stocks?\b", lab.strip(), re.I):
                found.append(("%s (total SI; first loss %s)" % (lab.strip(), "{:,.2f}".format(vals[0])), vals[-1]))
                if len(vals) >= 2 and vals[-1]:
                    p["first_loss"] = int(round(100 * vals[0] / vals[-1]))
    # 2) sum-insured table: 'Stocks' / 'Finished Stocks' label, optional description lines, amount
    if not found:
        # description lines are text only: a figure right after 'Stocks' is ITS sum insured, even 0.00 (Bajaj
        # 'Stocks|0.00|Stock in Process|0.00|...|Plant and Machinery|1,50,00,000.00' is NOT 150 lakh of stock)
        for m in re.finditer(r"\|((?:Finished |Raw Material |Total )?Stocks?(?:\s*[-:][^|]{0,100})?)\|((?:(?!\s*[`₹]?\s*(?:Rs\.?\s*)?[\d,]+(?:\.\d+)?\s*\|)[^|]{0,130}\|){0,6}?)"
                             r"[`₹]?\s*(?:Rs\.?\s*)?([\d,]{5,}(?:\.\d+)?|[\d,]+\.\d\d)\|", f, re.I):
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
    for lab in ("(?:Super Structure )?Building", r"Plant[ ,]*(?:&|and)?\s*Machinery", "Furniture[^|]{0,40}", "Contents", "Electrical Installations?"):
        m = re.search(r"\|(" + lab + r")[^|]{0,70}\|" + AMT + r"\|", f, re.I)
        if m and num(m.group(2)) > 0:
            ex.append("%s %.2f" % (m.group(1)[:30], num(m.group(2)) / 1e5))
            p.setdefault("items", []).append((m.group(1).strip()[:30], num(m.group(2)) / 1e5))
    p["excluded"] = "; ".join(ex) + (" (lakh)" if ex else "")
    has_stock = any(re.match(r"Stocks?\b", k, re.I) for k, _ in p.get("items", []))
    slab = "Stock"
    if len(found) == 1 and re.match(r"(?:Finished |Raw Material )Stocks?$", found[0][0].strip(), re.I):
        slab = found[0][0].strip()                    # the policy's own words ('Finished Stocks'), as the drafts copy them
    p["items"] = p.get("items", []) + ([(slab, p["stock"])] if p["stock"] is not None and not has_stock else [])
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
    if re.search(r"Stock\s*(?:Group|Item|Category)?\s*Monthly\s*Summary", ends, re.I):
        return "StockMonthly"
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
                and a.lower() not in MONTH_WORDS and not re.match(r"(page|cin|pan|e-mail|udyam|gstin|\d+-[a-z]{3}-\d+)", a.lower()):
            out.append(a)
    return out


def tally_pdf_totals(t, creditors=False):
    """Book debtors / creditors from a Tally Sundry Debtors / Creditors PDF, as [dict(date, dr, cr)] (Rs):
    'Group Monthly Summary' -> closing balance of every month (Baid: April 7,66,06,526.42 Dr = opening 01-05-2025);
    'Group Summary' -> Grand Total at the period end, credit and debit columns in the printed order
    (Clan Jul-25: Credit 2,78,89,020.19 | Debit 14,15,52,857.90 -> debtors 1136.64 lakh)."""
    f = t.replace("\n", "|")
    out = []
    if re.search(r"Monthly Summary", f):
        st = re.search(r"1-([A-Za-z]{3})-(\d{2,4}) to", f)
        if not st:
            return out
        y = int(st.group(2)); y = y + 2000 if y < 100 else y; m = MON[st.group(1).lower()]
        names = [dt.date(2000, i, 1).strftime("%B") for i in range(1, 13)]
        for mm in re.finditer(r"(?<=\|)(" + "|".join(names) + r")\|(" + PDF_AMT + r")\s*(Cr|Dr)(?=\|)", f):
            mi = names.index(mm.group(1)) + 1
            while mi != m:
                m = m % 12 + 1; y += (m == 1)
            v = num(mm.group(2)); dr = mm.group(3) == "Dr"
            out.append(dict(date=month_end(y, m), dr=v if dr else 0.0, cr=0.0 if dr else v))
        return out
    per = re.search(r"\d{1,2}-[A-Za-z]{3}-\d{2,4} to (\d{1,2}-[A-Za-z]{3}-\d{2,4})", f)
    gt = re.search(r"G\s*r\s*a\s*n\s*d\s*T\s*o\s*t\s*a\s*l((?:\|" + PDF_AMT + r"(?:\s*(?:Dr|Cr))?){1,4})\|", f)
    if not (per and gt and d(per.group(1))):
        return out
    day = d(per.group(1))
    vals = re.findall(r"(" + PDF_AMT + r")(?:\s*(Dr|Cr))?", gt.group(1))
    tagged = [(num(v), s) for v, s in vals if s]
    if tagged:          # Opening | Debit | Credit | Closing: the closing is the Dr / Cr figure on the side the header
        # prints 'Closing' (Denis text layer runs right to left: 'Closing|Transactions|Opening' -> first figure)
        hc, ho = f.find("|Closing"), f.find("|Opening")
        v, s = tagged[0] if hc >= 0 and (ho < 0 or hc < ho) else tagged[-1]
        dr, cr = (v, 0.0) if s == "Dr" else (0.0, v)
    elif len(vals) == 2:                          # two columns: Debit and Credit in the order the header prints them
        hd = re.search(r"\|(Debit|Credit)\|(?:[^|]*\|)?(Debit|Credit)\|", f)
        a, b = num(vals[0][0]), num(vals[1][0])
        dr, cr = (b, a) if hd and hd.group(1) == "Credit" else (a, b)
    elif len(vals) == 1:
        v = num(vals[0][0])
        dr, cr = (0.0, v) if creditors else (v, 0.0)
    else:
        return out
    out.append(dict(date=day, dr=dr, cr=cr))
    return out


def tally_pdf_stock(t):
    """Tally Stock (Group) Summary PDF: closing value of the Grand Total row, dated by the period end."""
    r = tally_pdf_stocks(t)
    return r[0] if r else None


def tally_pdf_stock_monthly(t):
    """Tally 'Stock Group MonthlySummary': closing value of every month up to the report end (Bhawani: April
    13,19,52,881.54 ...). Months after the end date repeat the last closing and are skipped; the end month is dated
    by the report end ('1-Apr-25 to 24-Dec-25' -> 24-12-2025)."""
    f = t.replace("\n", "|")
    per = re.search(r"1-([A-Za-z]{3})-(\d{2,4}) to (\d{1,2}-[A-Za-z]{3}-\d{2,4})", f)
    if not per or not d(per.group(3)):
        return []
    # one stock group only ('Page 1 | RAW MATERIAL | P a r t i c u l a r s') is a part of the stock, not the total
    grp = re.search(r"Page\s*\d+\|([^|]{2,60})\|P\s*a\s*r\s*t\s*i\s*c\s*u\s*l\s*a\s*r\s*s", f)
    if grp:
        return []
    y = int(per.group(2)); y = y + 2000 if y < 100 else y; m = MON[per.group(1).lower()]; end = d(per.group(3))
    names = [dt.date(2000, i, 1).strftime("%B") for i in range(1, 13)]
    out = []
    for mm in re.finditer(r"(?<=\|)(" + "|".join(names) + r")\|(" + PDF_AMT + r")(?=\|)", f):
        mi = names.index(mm.group(1)) + 1
        while mi != m:
            m = m % 12 + 1; y += (m == 1)
        day = min(month_end(y, m), end)
        if dt.date(y, m, 1) > end:
            break
        out.append(dict(date=day, stock=round(num(mm.group(2)) / 1e5, 5)))
    return out


def tally_pdf_stocks(t):
    """Every Stock Summary in the PDF: one file can hold a whole year, one report per month (Baid STOCK_25-26.pdf).
    Each Grand Total is dated by the period printed above it."""
    f = t.replace("\n", "|")
    out = {}
    # Tally prints 'G r a n d T o t a l'; PDF readers differ on the spacing ('G r a n d  T o t a l')
    for m in re.finditer(r"G\s*r\s*a\s*n\s*d\s*T\s*o\s*t\s*a\s*l\|(" + PDF_AMT + r")\|", f):
        per = list(re.finditer(r"\d{1,2}-[A-Za-z]{3}-\d{2,4} to (\d{1,2}-[A-Za-z]{3}-\d{2,4})", f[:m.start()]))
        if per and d(per[-1].group(1)):
            out[d(per[-1].group(1))] = round(num(m.group(1)) / 1e5, 5)
    return [dict(date=k, stock=v) for k, v in sorted(out.items())]


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


def gst_challan(t):
    """GST portal 'PAYMENT RECEIPT' (PMT-06 challan): CPIN, GSTIN, deposit date, total amount."""
    f = t.replace("\n", "|")
    if "PAYMENT RECEIPT" not in f or "CPIN" not in f:
        return None
    cp = re.search(r"CPIN\s*:?\|?\s*(\d{10,})", f)
    dd = re.search(r"Deposit Date\s*:?\|?\s*(\d{2}/\d{2}/\d{4})", f)
    am = re.search(r"Total Amount\s*\|\s*([\d,]+(?:\.\d+)?)", f)
    g = re.search(r"GSTIN\s*:?\|?\s*(\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z])", f)
    if not (cp and dd and am):
        return None
    return dict(cpin=cp.group(1), gstin=g.group(1) if g else "", paid=d(dd.group(1)), amount=num(am.group(1)))


def elec_bill_issue(t):
    """Electricity bill with a 'Billing Month' yyyymm table (JVVNL / Rajasthan discoms): billing month, bill issue date,
    due date and the Net Payable Amount - written as the drafts do (Bhawani Oct'25: 05-10-2025, 10,00,303)."""
    f = t.replace("\n", "|")
    i = f.find("Billing Month")
    if i < 0 or not re.search(r"VIDYUT|ELECTRICITY|POWER|DISCOM|Energy", f[:4000], re.I):
        return None
    seg = f[i:i + 700]
    bm = re.search(r"\|(20\d{2})(0[1-9]|1[0-2])\|", seg)
    if not bm:
        return None
    ds = [d(x) for x in re.findall(r"\d{2}-(?:\d{2}|[A-Za-z]{3})-\d{4}", seg[bm.end():])]
    ds = [x for x in ds if x]
    hd = seg[:bm.start()]                         # order of the date columns as the header prints them
    cols = [(k, hd.find(w)) for k, w in (("read", "Reading Date"), ("issue", "Bill Issue"), ("due", "Due Date")) if hd.find(w) >= 0]
    order = [k for k, _ in sorted(cols, key=lambda kv: kv[1])]
    dd = dict(zip(order, ds))
    net = re.search(r"Net Payable Amount\|(?:[\d.,-]+\|){0,8}?(\d[\d,]*(?:\.\d+)?)\|[A-Z][a-z]+(?: [A-Za-z]+)* (?:Crore|Lakh|Thousand|Hundred|Rupees)", f)
    if not (dd.get("issue") and net):
        return None
    return dict(ym=(int(bm.group(1)), int(bm.group(2))), bill_date=dd["issue"], due=dd.get("due"), amount=num(net.group(1)))


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


# ---------------- invoices (Invoice sheet) ----------------
INV_UNITS = r"KGS?|NOS?\.?|PCS|PES|MT|MTS|LTRS?|LITRES?|BAGS?|BOXES|BOX|SETS?|UNITS?|QTL|TONS?|MTRS?|SQM|SQFT|ROLLS?|DOZ"
INV_DATE = r"(\d{1,2}-[A-Za-z]{3}-\d{2,4}|\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})"
CO_WORDS = r"\b(?:Pvt|Private|Ltd|Limited|LLP|Industries|Enterprises|Traders|Trading|Corporation|Company|Plastics?|Polymers?|Packaging)\b"


def pdf_pages(p):
    try:
        with pymupdf.open(p) as doc:
            return [pg.get_text().translate(PDF_FIX) for pg in doc]
    except Exception:
        return []


def _inv_nums(s):
    return [num(x) for x in re.findall(r"\d[\d,]*(?:\.\d+)?", s)]


def _uom(u):
    u = u.rstrip(".").title()
    return "Pcs" if u == "Pes" else u          # OCR reads 'pcs' as 'pes'; the qty x rate proof still applies


def invoice_page(t, own_names=(), kind_hint=""):
    """One tax-invoice page -> dict(no, date, party, item, uom, qty, rate) or None.
    Text-layer Tally PDFs print the item row reversed ('1 Item | Amount | per | Rate | Qty UNIT'); OCR keeps reading order.
    Quantity x rate must equal an amount printed on the page (to the rupee), so a misread figure is never used.
    A page that is an invoice but whose item row cannot be proved returns dict(unread=True)."""
    if not re.search(r"Tax\s*Invoice|Invoice\s*No", t, re.I):
        return None
    f = re.sub(r"[ \t]+", " ", t)
    amounts = [round(v, 2) for v in _inv_nums(f) if v >= 100]
    # first item row: from the '1' serial after the goods header to the tax lines / total
    hdr = re.search(r"Description\s*of\s*Goods", f, re.I)
    body = f[hdr.end():] if hdr else f
    row = re.search(r"(?:^|[\n|])\s*1\s*[|)\]]?\s*([A-Za-z][^\n]*?(?:\n[^\n]*?){0,8}?)(?=\n\s*(?:2\s*[|)\]]?\s*[A-Za-z]|OUTPUT|Output|CGST|SGST|IGST|Total|continued))", body)
    seg = row.group(1) if row else ""
    tot = re.search(r"\bTotal\b[^\n]{0,60}?(\d[\d,]*\.?\d*)\s*(" + INV_UNITS + r")\b", body, re.I)
    best = None
    qtys = [(num(m.group(1)), m.group(2)) for m in re.finditer(r"(\d[\d,]*\.?\d*)\s*(" + INV_UNITS + r")\b", seg, re.I)]
    if tot:
        qtys.append((num(tot.group(1)), tot.group(2)))
    for q, u in qtys:
        if q <= 0:
            continue
        for r in _inv_nums(seg):
            if r <= 0 or r == q:
                continue
            hit = next((a for a in amounts if abs(a - q * r) <= 1.0 and a != q and a != r), None)
            if hit:
                best = (q, _uom(u), r); break
        if best:
            break
    if not best:
        # OCR slipped ONE digit of the rate (Gold Star 432: '162.51' printed 152.51): amount / qty is the rate when it
        # differs from a rate read on the row in exactly one digit, and qty x that rate equals the printed amount
        for q, u in qtys:
            for r in _inv_nums(seg):
                for a in amounts:
                    if q <= 0 or r <= 0 or a in (q, r):
                        continue
                    rr = round(a / q, 2)
                    s1, s2 = "%.2f" % r, "%.2f" % rr
                    if abs(rr * q - a) <= 1.0 and len(s1) == len(s2) and sum(c1 != c2 for c1, c2 in zip(s1, s2)) == 1:
                        best = (q, _uom(u), rr); break
                if best: break
            if best: break
    item = ""
    if seg:
        item = re.split(r"\s*[|]\s*|\s+\d{4,8}\b|\s+HSN\b|\s*\(", seg.split("\n")[0].strip())[0].strip(" :-")
    # invoice number + date: Tally prints 'Invoice No. | e-Way Bill No. | Dated' with the values below
    no = dte = None
    INV_NO = r"In[vw][o0]?[il1]?[cec]e\W{0,3}N[o0]"          # OCR: 'Invoice No', "Involce'No:", 'lnvoice N0'
    m = re.search(INV_NO + r"\.?.{0,40}?\n(?:.*?\b)?([A-Z]{0,4}/?[0-9][A-Z0-9/\-]{0,18})[\s|]+(?:\d{12}[\s|]+)?" + INV_DATE, f, re.I | re.S)
    if m and len(m.group(0)) < 260:
        no, dte = m.group(1), m.group(2)
    if not no:
        m = re.search(INV_NO + r"\.?\s*\n\s*(?:e-?Way\s*Bill\s*No\.?\s*\n\s*)?([A-Z0-9][A-Z0-9/\-]{0,20})\s*\n", f, re.I)
        no = m.group(1) if m else None
        md = re.search(r"Dated\s*\n\s*" + INV_DATE, f, re.I) or re.search(r"Ack\s*Date\s*:?\s*" + INV_DATE, f, re.I)
        dte = md.group(1) if md else None
    if dte and not d(dte):
        # OCR misread the date cell ('65-Oct-26'): the reference '<no> dt. <date>' or the e-invoice Ack Date are the same day
        md = (re.search(re.escape(no) + r"\s*dt\.?\s*" + INV_DATE, f, re.I) if no else None) or \
            re.search(r"Ack\s*Date\s*:?\s*" + INV_DATE, f, re.I)
        dte = md.group(1) if md and d(md.group(1)) else dte
    # names: the company lines at the top (seller first) and the 'Buyer (Bill to)' name; sales / purchase is decided at
    # build time (folder name, GSTIN, own name), when the borrower's names and GSTIN are known
    lines = [x.strip(" |") for x in f.split("\n") if x.strip(" |")]
    cos = [re.split(r"\s+(?:Invoice|\d{4}-\d{2}\b|\()", x)[0].strip(" -") for x in lines
           if re.search(CO_WORDS, x, re.I) and not re.search(r"IRN|Ack|Invoice\s*$", x)]
    buyer = re.search(r"Buyer\s*\(Bill\s*to\)\s*\n\s*([^\n]+)", f, re.I)
    gstins = re.findall(r"\b(\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z])\b", f)
    x = dict(unread=not best, hint=kind_hint, cos=cos[:6], buyer=buyer.group(1).strip(" |") if buyer else "",
             seller_gstin=gstins[0] if gstins else "", no=no or "", date=d(dte) if dte else None, item=item,
             uom=best[1] if best else "", qty=best[0] if best else None, rate=best[2] if best else None)
    x["party"] = x["buyer"] or (cos[0] if cos else "")          # provisional (page grouping); final in inv_resolve
    return x


def _clean_party(s):
    s = re.sub(r"^\s*(?:Name|M/s\.?)\s*:?\s*", "", re.sub(r"\s+", " ", s or ""), flags=re.I).strip(" |-")
    # OCR joins the header cell to the name: "Magnum Polysure Pvt Ltd Involce'No:, eWay Bil No.| Dated"
    s = re.split(r"\s+(?:In[vw]\w{0,3}[cec]e\W*N[o0]|Invoice|e-?\s*Way|Dated\b|Ack\b|IRN\b)", s, flags=re.I)[0].strip(" |-,:'")
    return s[:-1] if s.endswith(".") and s.count(".") == 1 else s     # 'Sevenstar Polytex Pvt Ltd.' -> no stray dot


def _looks_garbled(s):
    """OCR noise in a name ('et"A$'$ vtrRlptgs PVT. LTD.', 'Ci.AFI \\IITRNfIED'): characters a firm name never has,
    or too few letters. Such a row is shown on the review page to be typed, never written as read."""
    if not s or re.search(r"[\"$\\\\{}<>~^*=_;!?@#%]", s):
        return True
    letters = len(re.findall(r"[A-Za-z]", s))
    return letters < 4 or letters / max(1, len(s.replace(" ", ""))) < 0.75


def inv_resolve(F, own_names, own_gstin):
    """Final kind and party of every invoice: folder name ('Sales bill' / 'purchases bill') first, then the seller GSTIN,
    then whether the first company line is the borrower. Sales -> buyer; purchase -> seller (first company line not ours)."""
    own = [nospace(n.upper())[:12] for n in own_names if n]
    is_own = lambda s: any(o and o in nospace((s or "").upper()) for o in own)
    for x in F.get("inv", []):
        kind = x["hint"] or ("sales" if own_gstin and x["seller_gstin"] == own_gstin else "") \
            or ("sales" if x["cos"] and is_own(x["cos"][0]) else "purchase")
        seller = next((c for c in x["cos"] if not is_own(c)), "")
        x["kind"] = kind
        x["party"] = x.get("typed_party") or _clean_party(x["buyer"] if kind == "sales" and x["buyer"] else seller)
        # anything not proven goes to the review page as 'type it': garbled names, impossible dates, no invoice number
        if not x.get("typed"):
            bad_date = x["date"] is not None and not (dt.date.today().year - 3 <= x["date"].year <= dt.date.today().year + 1)
            if bad_date:
                x["date"] = None
            if _looks_garbled(x["party"]) or _looks_garbled(x["item"]) or not x["no"] or x["date"] is None:
                x["unread"] = True


def invoice_kind_hint(path):
    p = path.replace("\\", "/").lower()
    return "purchase" if "purchase" in p else ("sales" if re.search(r"sale", p) else "")


def add_invoices(F, path, pages, ocr=False):
    """Every invoice page of a file, in page order; a page that only continues the previous invoice is skipped."""
    hint = invoice_kind_hint(path); fn = os.path.basename(path); got = []
    for i, t in enumerate(pages):
        r = invoice_page(t, sorted(F["names"]), hint)
        if not r or any(re.search(r"INSURANCE|ASSURANCE", c, re.I) for c in r["cos"][:2]):
            continue                          # an insurer's premium tax invoice is not a trade invoice
        if not r["unread"] and got and got[-1]["no"] == r["no"] and r["no"]:
            continue
        # page 2 / 3 of the same invoice: Tally 'Tax Invoice(Page 2)' and the e-Way Bill page ('Doc No. : Tax Invoice - 857')
        if r["unread"] and got and (r["no"] == got[-1]["no"] or
                                    re.search(r"\(\s*Page\s*\d|^\s*e-?Way\s*Bill|Doc\s*No\.?\s*:?\s*Tax\s*Invoice", t[:300], re.I)):
            continue
        r.update(file=fn, page=i + 1, ocr=ocr, id="%s#%d" % (fn, i + 1)); got.append(r)
    F.setdefault("inv", []).extend(got)
    return got


def pick_invoices(F, kind, n=3):
    """As in the drafts: in file / page order, the first invoice of each new party, up to three.
    An unread page (handwritten / scanned) stays in its place so the auditor can type it on the review page."""
    out, seen = [], set()
    for x in sorted([x for x in F.get("inv", []) if x["kind"] == kind], key=lambda x: (x["file"].lower(), x["page"])):
        key = nospace((x.get("party") or "").upper())[:10]
        if key and _looks_garbled(x.get("party")):
            key = ""
        # an unread page keeps its place AND its party: a later invoice of the same party is not taken instead
        # (Gold Star: OCR missed Sevenstar 432 on page 1 -> Sevenstar 379 on page 2 was wrongly used)
        if key in seen and key:
            continue
        if key:
            seen.add(key)
        out.append(x)
        if len(out) == n:
            break
    return out


# ---------------- scan everything ----------------
def scan(paths, log):
    F = dict(bank={}, bank_name={}, inv=[], g1=[], g3=[], ss=[], book=[], debtors=set(), creditors=set(), directors=[], policies=[], names=set(),
             pf=[], esic=[], elec=[], dins=[], deb_tot=[], cred_tot=[], challan=[], advtax=[], elec_ai=[])
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
                        a, tx = bank_pdf(t); [x.update(src=fn) for x in tx]; F["bank"].setdefault(a, []).extend(tx); F["bank_name"][a] = F["bank_name"].get(a) or bank_of(t); log.append(("Bank PDF", fn, "%d txns; %s" % (len(tx), bal_msg(tx))))
                    elif stock_statement(t):
                        r = stock_statement(t); F["ss"].append(r); log.append(("Stock statement", fn, "%s stock %s" % (r["date"], r["stock"])))
                    elif company_stock_statement(t):
                        r = company_stock_statement(t); F["ss"].append(r)
                        log.append(("Stock statement", fn, "%s stock %s lakh (company format, text looks OCR'd - check)" % (r["date"], r["stock"])))
                    elif tally_pdf_kind(t):
                        k = tally_pdf_kind(t)
                        if k in ("Sundry Debtors", "Sundry Creditors"):
                            cred = k == "Sundry Creditors"
                            n = tally_pdf_names(t); F["creditors" if cred else "debtors"].update(n)
                            tt = tally_pdf_totals(t, cred); F["cred_tot" if cred else "deb_tot"].extend(tt)
                            log.append(("Creditor ledger" if cred else "Debtor ledger", fn, "%d names (PDF)%s" % (len(n), (
                                "; closing %s" % ", ".join("%s %.2f lakh" % (x["date"].strftime("%d-%m-%Y"), (x["cr"] if cred else x["dr"] - x["cr"]) / 1e5)
                                                           for x in tt[-3:])) if tt else "")))
                        elif k == "StockMonthly":
                            rs = tally_pdf_stock_monthly(t)
                            F["book"].extend(rs)
                            log.append(("Book stock" if rs else "skipped", fn, "; ".join("%s %s lakh" % (r["date"], r["stock"]) for r in rs) + " (monthly summary PDF)" if rs else "monthly stock summary without figures"))
                        elif k == "Stock":
                            rs = tally_pdf_stocks(t)
                            if rs:
                                F["book"].extend(rs); log.append(("Book stock", fn, "; ".join("%s %s lakh" % (r["date"], r["stock"]) for r in rs) + " (PDF)"))
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
                    elif gst_challan(t):
                        r = gst_challan(t)
                        if r["cpin"] in {x["cpin"] for x in F["challan"]}:
                            log.append(("GST challan", fn, "CPIN %s already read from another file - counted once" % r["cpin"]))
                        else:
                            F["challan"].append(r); log.append(("GST challan", fn, "Rs %s deposited %s" % (r["amount"], r["paid"])))
                    elif elec_bill_issue(t):
                        r = elec_bill_issue(t)
                        if not any(x["ym"] == r["ym"] for x in F["elec_ai"]):
                            F["elec_ai"].append(r)
                        log.append(("Electricity bill", fn, "%d-%02d issued %s net payable %s" % (r["ym"] + (r["bill_date"], r["amount"]))))
                    elif electricity_bill(t):
                        r = electricity_bill(t); F["elec"].append(r)
                        log.append(("Electricity bill", fn, "%d-%02d net payable %s; previous month paid %s on %s" % (r["ym"] + (r["net"], r["prev_amount"], r["prev_paid"]))))
                    elif policy(t, fn):
                        r = policy(t, fn); F["policies"].append(r); F["names"].add(r["insured"])
                        log.append(("Policy " + r["cls"], fn, "%s %s-%s stock SI %s" % (r["no"], r["frm"], r["to"], r["stock"])))
                    elif re.search(r"Tax\s*Invoice|Invoice\s*No", t, re.I) and add_invoices(F, p, pdf_pages(p)):
                        r = [x for x in F["inv"] if x["file"] == fn]
                        log.append(("Invoice", fn, "%d invoice(s): %s" % (len(r), ", ".join(x["no"] or "page %d not readable" % x["page"] for x in r))))
                    else:
                        log.append(("skipped", fn, "not recognised" if t.strip() else "no text layer (scanned)"))
            elif ext == "txt":
                t = open(p, encoding="utf-8", errors="replace").read()
                if "STATEMENT" in t and "BALANCE" in t:
                    a, tx = bank_txt(t); [x.update(src=fn) for x in tx]; F["bank"].setdefault(a, []).extend(tx); F["bank_name"][a] = F["bank_name"].get(a) or bank_of(t); log.append(("Bank TXT", fn, "%d txns; %s" % (len(tx), bal_msg(tx))))
            elif ext in ("xlsx", "xlsm", "csv", "xls") and re.search(OUT_YES, fn, re.I) and not re.search(OUT_NO, fn, re.I) \
                    and outward_file(F, p, log):
                pass                                          # outward-supply data read for the GST summary
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
                        F["bank"].setdefault(acc or fn, []).extend(tx)
                        F["bank_name"][acc or fn] = F["bank_name"].get(acc or fn) or bank_of(" ".join(str(c) for r in rows[:sheet_head(rows)] for c in r if c))
                        log.append(("Bank sheet", fn, "%d txns; %s" % (len(tx), bal_msg(tx))))
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


def remitter(u):
    """Remitter of a structured credit 'NEFT/<UTR>/<REMITTER>/<BANK>/<remarks>' (also RTGS, IMPS, INB/RTGS); '' otherwise.
    Customers often put the borrower's own name in the remarks (Clan: 'NEFT/.../RDENTERPRISE/ICICI BANK/CLANVITRIFIED'),
    so the own-name test must look at the remitter only."""
    seg = [x.strip() for x in u.split("/")]
    i = 1 if seg and seg[0] == "INB" else 0
    if len(seg) >= i + 3 and re.match(r"^(NEFT|RTGS|IMPS|NEFT-?IN)", seg[i]):
        return seg[i + 2]
    if seg and seg[0] == "TRF":               # 'TRF/TAPADIYA TRADERS/clan vitrified pvt ltd', 'TRF/568/ADVANCE COOLING/tran'
        return next((x for x in seg[1:] if len(re.sub(r"[^A-Z]", "", x)) >= 3), "")
    return ""


def director_keys(names):
    """Narrations cut a name at ~22 characters ('MAHESHWAR SINGH RANAWA'), so the full name is not enough."""
    keys = set()
    for nm in names:
        w = (nm or "").upper().split()
        if not w:
            continue
        if len(w[-1]) >= 5:
            keys.add(w[-1])
        if len(w) >= 3 and len(" ".join(w[:2])) >= 9:
            keys.add(" ".join(w[:2]))
        if len(w) >= 2:
            keys.add(w[0] + " " + w[-1][:5])
    return sorted(keys, key=len, reverse=True)


def classify(tx, F, related, own):
    out = []
    debtors = sorted(F["debtors"], key=len, reverse=True)
    creditors = sorted(F["creditors"])
    deb_first = {toks(dn)[0] for dn in debtors if toks(dn)}
    own_keys = [("".join(toks(nm)), nm) for nm in own if nm and toks(nm)]
    dkeys = director_keys([nm for nm in related if nm in F.get("directors", [])])
    for t in tx:
        if t["amt"] <= 0:
            continue
        u = t["narr"].upper(); r = dict(t); n = nospace(u)
        hit = None
        rem = remitter(u)
        hay = nospace(rem) if rem else n
        for key, nm in own_keys:
            if key in hay:
                hit = ("EXC", "Self-transfer / contra", "Own name (%s) - transfer from own account" % nm)
        if not hit:
            k = next((k for k in dkeys if k in u), None)
            if k:
                hit = ("EXC", "Director/Partner/Owner/Related party", "Promoter-director (narration has '%s')" % k)
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
    accs = sorted(F["bank"].items(), key=lambda kv: -len(kv[1]))
    cc = next((a for a, tx in accs if _is_cc(dedupe(tx))), accs[0][0] if accs else "")
    # check page left the bank blank -> the bank named in the CC statement's header (Gold Star: 'A X I S  B A N K')
    bank1 = banks[0][0] if banks and banks[0][0] else ((F.get("bank_name") or {}).get(cc) or "Bank")
    ca = next((a for a, tx in accs if a != cc), None)
    inp = lambda t: start <= (t["date"].year, t["date"].month) <= base

    def sums(acc):
        if not acc or acc not in F["bank"]:
            return None, None
        tx = [t for t in dedupe(F["bank"][acc]) if inp(t)]
        dr = sum(-t["amt"] for t in tx if t["amt"] < 0)
        # credits as the drafts total them: every credit, less transfers from own accounts and from directors / related parties
        ex = sum(r["amt"] for r in allc if r["acc"] == acc and inp(r) and r["cat"].startswith(("Self-transfer", "Director")))
        cr = sum(t["amt"] for t in tx if t["amt"] > 0) - ex
        return _lakh(dr), _lakh(cr)

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
    # sales = GSTR-3B 3.1 (a)+(b)+(c) per month; a month without a 3B -> GSTR-1, then the books sales register
    g3 = {r["ym"]: r.get("sales", r["taxable"]) for r in F["g3"]}
    g1 = {r["ym"]: r["taxable"] for r in F["g1"]}
    bk = F.get("sales") or {}
    vals = [g3.get(m, g1.get(m, bk.get(m))) for m in months]
    miss = [_mlab(m) for m, v in zip(months, vals) if v is None]
    sales = sum(v for v in vals if v is not None) if len(miss) < len(months) else None
    ws["B10"] = "Opening Debtors as on %s" % dt.date(start[0], start[1], 1).strftime("%d-%m-%Y"); ws["C10"] = _lakh(opening)
    ws["B11"] = "Add: Sales during %s" % per + (" (no GST return / books sales for %s)" % ", ".join(miss) if miss and sales is not None else "")
    ws["C11"] = _lakh(sales)
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


def _addr(a):
    """Policy address as the drafts write it: '..., SANWER ROAD, INDORE 452015' - no repeated town, no state name,
    town and PIN joined by a space (policies print 'INDORE INDORE ,MADHYA PRADESH, 452015' / 'INDORE-452015')."""
    a = re.sub(r"\s+", " ", a or "").strip(" ,.")
    a = re.sub(r"\s*,\s*", ", ", a)
    a = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", a, flags=re.I)                    # INDORE INDORE -> INDORE
    states = sorted({v for v in GST_STATES.values() if v}, key=len, reverse=True)
    a = re.sub(r",\s*(?:%s)\s*(?=,|\s*-?\s*\d{6}\b|$)" % "|".join(re.escape(x) for x in states), "", a, flags=re.I)
    a = re.sub(r"\s*[,-]\s*(\d{6})\b", r" \1", a)                                # INDORE, 452015 / INDORE-452015
    return a.strip(" ,")


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
    ws["%s%d" % (c(0), top + 4)] = p["no"]; ws["%s%d" % (c(4), top + 4)] = _addr(p.get("location"))
    for i, (lab, v) in enumerate(items):
        ws["%s%d" % (c(1), top + 4 + i)] = lab
        ws["%s%d" % (c(2), top + 4 + i)] = round(v, 2) if isinstance(v, (int, float)) else "SI not stated"


def draft_insurance(ws, F, cfg):
    pols = {}
    for p in F["policies"]:
        if p["cls"] == "Other":
            continue
        q = pols.get(p["no"])
        # copies of one policy number with different stock SI (Baid: schedule 0 / endorsement 1500 lakh) -> the higher
        rank = lambda x: (x["stock"] is not None, x["stock"] or 0)
        if q is None or rank(p) > rank(q):
            pols[p["no"]] = p
    fire = sorted((p for p in pols.values() if p["cls"] == "Fire"), key=lambda p: pdate(p["frm"]), reverse=True)
    burg = sorted((p for p in pols.values() if p["cls"] == "Burglary"), key=lambda p: pdate(p["frm"]), reverse=True)
    fitems = lambda p: [(k, p["stock"] if k == "Stock" else v) for k, v in (p.get("items") or [("Stock", p["stock"])])][:5]
    _pol_block(ws, "B", 1, fire[0] if fire else None, fitems(fire[0]) if fire else [])
    _pol_block(ws, "H", 1, fire[1] if len(fire) > 1 else None, fitems(fire[1]) if len(fire) > 1 else [])
    bitems = lambda p: [(k, v) for k, v in p["items"]][:5] if len(p.get("items") or []) > 1 else [("Stock", p["stock"])]
    _pol_block(ws, "B", 12, burg[0] if burg else None, bitems(burg[0]) if burg else [])
    _pol_block(ws, "H", 12, burg[1] if len(burg) > 1 else None, bitems(burg[1]) if len(burg) > 1 else [])
    # stock vs insured, as the drafts list it (Gold Star, Clan, Denis, Baid): every stock date from the month before the
    # summation period to its end - book stock where the books give it, else the stock statements - not before the first
    # stock policy starts (Baid: policy from 14-11-2025 -> Nov'25 on) - then the visit date
    base = month_end(*cfg["sum_to"])
    books = {}
    for b in F["book"]:
        if b.get("date"):
            books[b["date"]] = max(books.get(b["date"], 0), b["stock"])     # a group report never beats the total
    ss = {s["date"]: s for s in F["ss"] if s.get("date")}
    lo = _ym_add(cfg["sum_from"], -1)
    starts = [pdate(p["frm"]) for p in pols.values() if p["cls"] in ("Fire", "Burglary") and p.get("stock")]
    if starts:
        lo = max(lo, (min(starts).year, min(starts).month))
    inwin = lambda xs: sorted(x for x in xs if lo <= (x.year, x.month) and x <= base)
    dates = inwin(books) or inwin(ss)
    visit = cfg.get("visit") or max([x for x in books if x > base] or [None], key=lambda x: x or dt.date.min)
    for c, v in (("B20", "Month"), ("C20", "Total Stock"), ("D20", "Insured Amt (Fire)"), ("E20", "Insured Amt (Burglary)"), ("F20", "Remarks")):
        ws[c] = v
    use = list(pols.values())
    rows = [(x, _mlab((x.year, x.month))) for x in dates] + ([(visit, "Visit date %s" % visit.strftime("%d-%m-%Y"))] if visit else [])
    extra = max(0, len(rows) - 4)                     # the template has 4 rows (21-24): insert more, styled like row 24
    if extra:
        ws.insert_rows(25, extra)
        from copy import copy
        for k in range(extra):
            for col in range(1, 8):
                src, dst = ws.cell(24, col), ws.cell(25 + k, col)
                if src.has_style:
                    dst._style = copy(src._style)
                dst.number_format = src.number_format
    o = extra
    for i, (day, lab) in enumerate(rows):
        r = 21 + i
        stock = books.get(day)
        if stock is None and day in ss:
            stock = ss[day]["stock"]
        if stock is None and cfg.get("visit_stock") is not None and day == visit:
            stock = cfg["visit_stock"]
        fp, _ = cover(use, "Fire", day); bp, _ = cover(use, "Burglary", day)
        fsi = fp["stock"] if fp and fp["stock"] is not None else ("SI not stated" if fp else "Policy not provided")
        bsi = bp["stock"] if bp and bp["stock"] is not None else ("SI not stated" if bp else "Policy not provided")
        ws["B%d" % r] = lab; ws["C%d" % r] = round(stock, 5) if stock is not None else None
        ws["D%d" % r] = fsi; ws["E%d" % r] = bsi
        if stock is None or not isinstance(fsi, (int, float)):
            rem = "Unable to comment"
        elif stock > fsi or (isinstance(bsi, (int, float)) and stock > bsi):
            rem = "Inadequately Insured"
        else:
            rem = "Adequately Insured"
        ws["F%d" % r] = rem
    # fixed assets and collateral as in the draft: insured value from the current fire policy; asset values typed on the portal
    ws["B%d" % (27 + o)] = "Fixed asset schedule"
    for c, v in (("B", "Particulars"), ("C", "Assets value (Rs. Lakhs)"), ("D", "Insured Value (Rs. In Lakhs)"), ("E", "Remarks")):
        ws["%s%d" % (c, 28 + o)] = v
    it = dict((k.lower(), v) for k, v in (fire[0].get("items") or [])) if fire else {}
    pm = next((v for k, v in it.items() if "plant" in k), None)
    bld = next((v for k, v in it.items() if "building" in k), None)
    ws["B%d" % (29 + o)] = "Plant & Machinery"; ws["C%d" % (29 + o)] = cfg.get("pm_value")
    ws["D%d" % (29 + o)] = round(pm, 2) if pm is not None else None
    if cfg.get("pm_value") is not None and pm is not None:
        ws["E%d" % (29 + o)] = "Adequately Insured" if pm >= cfg["pm_value"] else "Inadequately Insured"
    for c, v in (("B", "Address of collateral"), ("C", "Assets value (Rs. Lakhs)"), ("D", "Insured Value (Rs. In Lakhs)"), ("E", "Remarks")):
        ws["%s%d" % (c, 31 + o)] = v
    ws["B%d" % (32 + o)] = cfg.get("collateral_address") or (_addr(fire[0].get("location")) if fire else "")
    cv = cfg.get("collateral_value")
    if cv is None and F.get("collateral"):
        cv = round(F["collateral"]["value"] / 1e5, 5)
    ws["C%d" % (32 + o)] = cv; ws["D%d" % (32 + o)] = round(bld, 2) if bld is not None else None
    cfg = dict(cfg, collateral_value=cv)
    if cfg.get("collateral_value") is not None and bld is not None:
        ws["E%d" % (32 + o)] = "Adequately Insured" if bld >= cfg["collateral_value"] else "Inadequately Insured"


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


# due dates extended by the government (add a line when CBIC extends one): (return, tax period) -> new due date
GST_DUE_EXT = {("3B", (2025, 9)): dt.date(2025, 10, 25)}       # GSTR-3B Sep-2025 extended to 25-10-2025


def gst_due(kind, ym):
    y, m = ym
    return GST_DUE_EXT.get((kind, ym)) or dt.date(y + (m == 12), m % 12 + 1, 20 if kind == "3B" else 11)


def _ret_rows(ws, rows, col, top, kind, n):
    L = openpyxl.utils.get_column_letter; c0 = openpyxl.utils.column_index_from_string(col)
    one = {}
    for r in rows:                                    # the same return uploaded twice ('..._012025 (1).pdf') -> once
        one.setdefault(r["ym"], r)
    for i, r in enumerate(sorted(one.values(), key=lambda r: r["ym"], reverse=True)[:n]):
        due = gst_due(kind, r["ym"])
        fd = d(r["arn"]) if r.get("arn") else None
        vals = [_fy(r["ym"]), _mname(r["ym"]), due, fd, ("Timely Filed" if fd <= due else "Late Filed") if fd else "Filing date not found"]
        for k, v in enumerate(vals):
            ws["%s%d" % (L(c0 + k), top + i)] = v


def gst_from_bank(F):
    """GST cash payments read from the bank statements' debits: {return period: dict(amount, paid)}.
    A payment made up to the 30th of month M is the return of month M-1 (GSTR-3B due on the 20th).
    Gold Star: every period ties to the GSTR-3B 'Tax paid in cash' (incl. RCM), e.g. May'26 43,03,092."""
    res = {}
    for tx in F["bank"].values():
        for t in dedupe(tx):
            u = t["narr"].upper()
            if t["amt"] >= 0 or not re.search(r"GST\s*(TAX)?\s*PAYMENT|GSTN|GST PMT|GST CHALLAN|CBIC GST", u) \
                    or re.search(r"ICEGATE|@|CHRG|CHARGE", u):
                continue
            ym = (t["date"].year, t["date"].month)
            ym = _ym_add(ym, -1) if t["date"].day <= 30 else ym
            g = res.setdefault(ym, dict(amount=0.0, paid=t["date"]))
            g["amount"] = round(g["amount"] - t["amt"], 2); g["paid"] = max(g["paid"], t["date"])
    # GST challan receipts are the documents themselves: a period that has them is taken from them, not the bank
    rec = {}
    for c in F.get("challan", []):
        if not c["paid"]:
            continue
        ym = (c["paid"].year, c["paid"].month)
        ym = _ym_add(ym, -1) if c["paid"].day <= 30 else ym
        g = rec.setdefault(ym, dict(amount=0.0, paid=c["paid"]))
        g["amount"] = round(g["amount"] + c["amount"], 2); g["paid"] = max(g["paid"], c["paid"])
    res.update(rec)
    return res


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
    # GST challans: one row per GSTR-3B period shown on the left (same months), amount and date from the bank debits
    paid = gst_from_bank(F)
    g3 = sorted({r["ym"] for r in F["g3"] if not gst or r["gstin"] == gst}, reverse=True)[:6]
    for i, ym in enumerate(g3):
        v = paid.get(ym)
        if not v:
            continue
        due = gst_due("3B", ym)
        for c, x in zip("JKLMN", ("GST Challan", _mname(ym), v["amount"], v["paid"], "Timely Filed" if v["paid"] <= due else "Late Filed")):
            ws["%s%d" % (c, 3 + i)] = x
    for c, v in zip("BCDEF", ("Particulars", "Month", "Amount (Rs.)", "Date Paid", "Remarks")):
        ws["%s10" % c] = v
    _due_rows(ws, F.get("pf", []), "B", 11, "PF", 6)
    _due_rows(ws, F.get("esic", []), "J", 11, "ESIC", 6)
    for c, v in zip("BCDEF", ("Particulars", "Tax Period", "Amount (Rs.)", "Date Paid", "Remarks")):
        ws["%s18" % c] = v
    # advance tax challans (latest four): due 15 Jun / 15 Sep / 15 Dec / 15 Mar of the financial year
    for i, a in enumerate(sorted(F.get("advtax", []), key=lambda a: a["paid"], reverse=True)[:4]):
        p = a["paid"]; q = dt.date(p.year, 6, 15) if p.month in (4, 5, 6) else dt.date(p.year, 9, 15) if p.month in (7, 8, 9) \
            else dt.date(p.year, 12, 15) if p.month in (10, 11, 12) else dt.date(p.year, 3, 15)
        fy = p.year if p.month >= 4 else p.year - 1
        for c, v in zip("BCDEF", ("Advance Tax", "FY %d-%02d" % (fy, (fy + 1) % 100), a["amount"], p, "Timely Paid" if p <= q else "Lately Paid")):
            ws["%s%d" % (c, 19 + i)] = v
    # electricity: the payment of month M is printed on the bill of month M+1 ("Last Month Payment Details")
    ws["B24"] = "Payment Details of Electricity Bill (For last 12 month):"
    bills = {b["ym"]: b for b in F.get("elec", [])}
    ai_bills = {b["ym"]: b for b in F.get("elec_ai", [])}
    if ai_bills and not bills:              # bills read by AI: bill date and net amount, as the drafts write them
        last = max(ai_bills)
        for blk, top in ((0, 25), (1, 29)):
            ws["B%d" % top] = "Month"; ws["B%d" % (top + 1)] = "Date"; ws["B%d" % (top + 2)] = "Amount"
            for i in range(6):
                ym = _ym_add(last, -(blk * 6 + i)); col = "CDEFGH"[i]; b = ai_bills.get(ym)
                ws["%s%d" % (col, top)] = _mlab(ym)
                if b:
                    ws["%s%d" % (col, top + 1)] = b["bill_date"]; ws["%s%d" % (col, top + 2)] = b["amount"]
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


def draft_invoice(ws, F, cfg):
    """Sales and purchase invoices as in the drafts: in file order, the first invoice of each new party, three of each.
    Taxable value = quantity x rate (formula), as the auditor writes it."""
    inv_resolve(F, list(cfg["own"]) + sorted(F["names"]), cfg.get("gstin") or "")
    for k, v in {"A": 9.6, "B": 14, "C": 12, "D": 32, "E": 28, "F": 8, "G": 12, "H": 10, "I": 15}.items():
        ws.column_dimensions[k].width = v
    bold = openpyxl.styles.Font(bold=True)
    for top, kind, title, who in ((2, "sales", "Sales Invoices", "Name of Buyer"), (8, "purchase", "Purchase invoices", "Name of Supplier")):
        ws["B%d" % top] = title; ws["B%d" % top].font = bold
        for c, v in zip("BCDEFGHI", ("Invoice No.", "Invoice Date", who, "Item", "UOM", "Quantity", "Rate", "Taxable Value ")):
            ws["%s%d" % (c, top + 1)] = v; ws["%s%d" % (c, top + 1)].font = bold
        sel = pick_invoices(F, kind)
        for i, x in enumerate(sel):
            r = top + 2 + i
            if x["unread"]:                 # not proven and not typed on the check page: row kept blank in its place
                continue
            no = x["no"]; no = int(no) if str(no).isdigit() else no
            for c, v in zip("BCDEFGH", (no, x["date"], x["party"], x["item"], x["uom"], x["qty"], x["rate"])):
                ws["%s%d" % (c, r)] = v
            ws["I%d" % r] = "=G%d*H%d" % (r, r) if x["qty"] is not None and x["rate"] is not None else None
        if not sel:
            ws["B%d" % (top + 2)] = ("%s invoices are scanned / handwritten - type them on the check page" % kind.title()
                                     if pick_invoices(F, kind) else "No %s invoice found in the documents" % kind)


def build_excel_draft(F, cfg, out):
    """Workbook in the user's own layout: summation, Insurance, GST (template = blank copy of their draft)."""
    allc = []
    for acc, tx in F["bank"].items():
        for r in classify(dedupe(tx), F, cfg["related"], cfg["own"]):
            r["acc"] = acc; allc.append(r)
    wb = openpyxl.load_workbook(TEMPLATE)
    draft_invoice(wb.create_sheet("Invoice", 0), F, cfg)
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
    """Summation period: six months ending at the latest stock month (stock statement or books - the drawing-power base)
    that the bank statements cover. A stock date outside the statements (Denis: only Mar'25 read, bank Sep'25-Aug'26)
    would give a period with no bank rows, so the bank's last month is used then."""
    bank = [t["date"] for tx in F["bank"].values() for t in tx]
    ss = [s["date"] for s in F["ss"] if s.get("date")]
    book = [s["date"] for s in F["book"] if s.get("date")]       # books only when no stock statement fits (the
    if bank:                                                      # latest book stock is often the visit date)
        lo, hi = min(bank), max(bank)
        fit = lambda xs: [x for x in xs if (lo.year, lo.month) <= (x.year, x.month) <= (hi.year, hi.month)]
        end = max(fit(ss) or fit(book) or [hi])
    else:
        end = max(ss or book or [dt.date.today()])
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
    ocr_done = {b for a, b, c in log if a.startswith("OCR") or (a == "AI" and "used" in c and "nothing" not in c)}
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
        invoices=_review_invoices(F), unread=unk, log=log, ai=F.get("ai_rows", [])), default=str)


def _review_invoices(F):
    inv_resolve(F, sorted(F["names"]), next(iter(sorted({r["gstin"] for r in F["g1"] + F["g3"]})), ""))
    out = []
    for kind in ("sales", "purchase"):
        for x in pick_invoices(F, kind):
            out.append(dict(id=x["id"], kind=kind, file=x["file"], page=x["page"], unread=x["unread"], ocr=x["ocr"], no=x["no"],
                            date=x["date"].strftime("%d-%m-%Y") if x["date"] else "", party=x["party"], item=x["item"],
                            uom=x["uom"], qty=x["qty"], rate=x["rate"]))
    return out


# ---------------- OCR of scanned pages (done in the browser by tesseract.js; parsed here by fixed rules) ----------------
OCR_MAX_PAGES = 10
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
    if not got and re.search(r"Tax\s*Invoice|Invoice\s*No", text, re.I):
        r = add_invoices(F, path, text.split("\f"), ocr=True)
        if r:
            got.append("%d invoice(s): %s" % (len(r), ", ".join(x["no"] or "page %d not readable - type it" % x["page"] for x in r)))
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
    for x in F.get("inv", []):                  # invoice rows typed / corrected on the review page
        e = (f.get("inv") or {}).get(x["id"])
        if not e:
            continue
        shown = dict(no=x["no"], date=x["date"].strftime("%d-%m-%Y") if x["date"] else "", party=x.get("party", ""),
                     item=x["item"], uom=x["uom"], qty="" if x["qty"] is None else str(x["qty"]),
                     rate="" if x["rate"] is None else str(x["rate"]))
        if all(str(e.get(k, "")).strip() == str(v).strip() for k, v in shown.items()):
            continue                          # sent back unchanged: the rules' verdict (read / type it) stands
        for k in ("no", "item", "uom"):
            if str(e.get(k, "")).strip():
                x[k] = str(e[k]).strip()
        if str(e.get("party", "")).strip():
            x["typed_party"] = str(e["party"]).strip()
        if str(e.get("date", "")).strip() and d(e["date"].strip()):
            x["date"] = d(e["date"].strip())
        for k in ("qty", "rate"):
            if str(e.get(k, "")).strip():
                x[k] = num(str(e[k]))
        x["typed"] = True
        if x["qty"] is not None and x["rate"] is not None:
            x["unread"] = False
        x["hint"] = x["hint"] or e.get("kind", "")
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


# ---------------- AI fallback (Gemini, called from the browser with the user's own key) ----------------
# Only files that NO rule could read are sent. Every value is checked by fixed rules below (dates, year, period vs
# payment date, totals) and shown on the check page as 'AI-read - check'; a value that fails a check is not used.
AI_MAX_PAGES = 40
AI_SKIP = r"balance\s*sheet|audit|3ca|3cd|itr|annual report|sanction|\bsl\b|licen[cs]e|consent|valuation|cte|cca|gpcb|pollution|checklist"
AI_PROMPT = """You read Indian business documents for a bank stock audit. Return JSON only, no prose.
{"doc_type": one of "gst_challan","pf_challan","esic_challan","electricity_bill","advance_tax_challan","stock_statement",
 "insurance_policy","gstr1","gstr3b","other",
 "entries": [one object per challan / bill / return / statement / policy in the file - a file may hold many months]}
Entry fields (null when not printed; dates "DD-MM-YYYY"; money in rupees as plain numbers unless said otherwise):
 page (1-based), ref (TRRN / CPIN / challan no / bill no / policy no),
 period "YYYY-MM": PF/ESIC wage or contribution month, electricity bill month, GST return tax period,
 amount: total paid (challans) or net bill payable (electricity),
 paid_date: payment confirmation / deposit / realisation date (NOT the challan generation date, NOT the due date),
 bill_date, due_date, filed_date (GST return ARN / filing date),
 gst_a, gst_b, gst_c: GSTR-3B table 3.1 rows (a), (b), (c) taxable value; gst_taxable: GSTR-1 total taxable value,
 stock_date and stock_lakh (total inventory / stock as on that date, Rs lakh), debtors_lakh, creditors_lakh,
 insurer, policy_class ("Fire" / "Burglary" / "Other"), from_date, to_date, stock_si (sum insured of stocks only, rupees).
Read every page. Copy figures and dates exactly as printed - never compute, round or guess; use null if unsure.
Read years carefully from the full date printed (e.g. 15-09-2026), never from a reference number."""


def ai_candidates():
    """PDFs no rule (and no OCR rule) could read: the browser sends these to Gemini."""
    F, log = _B["F"], _B["log"]
    read = {b for a, b, c in log if a not in ("skipped",) and not (a == "OCR" and "no policy" in c)}
    by_name = {}
    for pth in _B["paths"]:
        by_name.setdefault(os.path.basename(pth), pth)
    want = [b for a, b, c in log if a == "skipped" and b not in read and b.lower().endswith(".pdf")
            and not re.search(AI_SKIP, b, re.I)]
    want += [p["file"] for p in F["policies"] if p["cls"] != "Other" and p["stock"] is None]
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
        if n <= AI_MAX_PAGES and os.path.getsize(pth) < 15e6:
            out.append(dict(path=pth, file=fn, pages=n))
    return json.dumps(dict(files=out, prompt=AI_PROMPT))


def _ai_fy(fn):
    """Financial year printed in a file name ('PF_25-26.pdf', 'ESIC 2026-27') -> start year, else None."""
    m = re.search(r"(?:20)?(\d{2})\s*[-_]\s*(\d{2})(?!\d)", fn)
    return 2000 + int(m.group(1)) if m and int(m.group(2)) == (int(m.group(1)) + 1) % 100 else None


def _ai_date(s, fy=None):
    """'DD-MM-YYYY' -> date. A year outside today-4..today+1 is moved into the file's financial year, else None."""
    x = d(str(s or "").replace("/", "-").replace(".", "-"))
    if not x:
        return None
    if not (dt.date.today().year - 4 <= x.year <= dt.date.today().year + 1):
        if not fy:
            return None
        x = x.replace(year=fy if x.month >= 4 else fy + 1)
    return x


def _ai_ym(s, fy=None):
    m = re.match(r"\s*(\d{4})-(\d{1,2})", str(s or ""))
    if not m or not 1 <= int(m.group(2)) <= 12:
        return None
    y, mo = int(m.group(1)), int(m.group(2))
    if not (dt.date.today().year - 4 <= y <= dt.date.today().year + 1):
        if not fy:
            return None
        y = fy if mo >= 4 else fy + 1
    return (y, mo)


def _ai_money(v):
    try:
        v = float(str(v).replace(",", "").replace("Rs", "").strip())
        return v if v > 0 else None
    except (TypeError, ValueError):
        return None


def ai_merge(path, text):
    """Apply one Gemini answer to the data, entry by entry, after the rule checks. Returns the review summary."""
    F, log = _B["F"], _B["log"]; fn = os.path.basename(path); fy = _ai_fy(fn)
    rows = F.setdefault("ai_rows", [])
    try:
        j = json.loads(re.sub(r"^```(?:json)?|```$", "", text.strip()))
    except Exception:
        log.append(("AI", fn, "answer was not readable JSON - file not used")); return browser_summary()
    kind = j.get("doc_type") or "other"; used = 0
    def row(e, what, ok, why=""):
        rows.append(dict(file=fn, page=e.get("page"), kind=kind, what=what, ok=ok, why=why))
    gst = next(iter(sorted({r["gstin"] for r in F["g1"] + F["g3"]})), "")
    for e in j.get("entries") or []:
        amt = _ai_money(e.get("amount")); paid = _ai_date(e.get("paid_date"), fy); ym = _ai_ym(e.get("period"), fy)
        if kind in ("pf_challan", "esic_challan"):
            if ym is None and paid:
                ym = _ym_add((paid.year, paid.month), -1)
            # a wage month is paid after it ends, normally by the 15th of the next month (late up to ~3 months)
            gap = (paid - month_end(*ym)).days if paid and ym else None
            what = "%s %s Rs %s paid %s" % (kind[:-8].upper(), _mlab(ym) if ym else "?", amt, paid)
            if not (amt and paid and ym) or not 0 < gap <= 105:
                row(e, what, False, "missing figure" if gap is None else "payment date %s is not after the wage month %s" % (paid, _mlab(ym))); continue
            lst = F["pf" if kind == "pf_challan" else "esic"]
            if any(x["ym"] == ym and not x.get("ai") for x in lst):
                row(e, what, False, "month already read by a rule"); continue
            same = next((x for x in lst if x["ym"] == ym and x.get("ref") != e.get("ref")), None)
            if same:                       # second challan of one month (arrears): added together
                same["amount"] += amt; same["paid"] = max(same["paid"], paid); row(e, what, True, "added to the other challan of this month")
            elif not any(x.get("ref") and x.get("ref") == e.get("ref") for x in lst):
                lst.append(dict(kind=kind[:-8].upper(), ym=ym, amount=amt, paid=paid, ai=True, ref=e.get("ref"))); row(e, what, True)
            used += 1
        elif kind == "electricity_bill":
            bd = _ai_date(e.get("bill_date"), fy)
            ym = ym or ((bd.year, bd.month) if bd else None)
            what = "Electricity %s bill %s Rs %s" % (_mlab(ym) if ym else "?", bd, amt)
            if not (ym and bd and amt) or abs((bd.year * 12 + bd.month) - (ym[0] * 12 + ym[1])) > 1:
                row(e, what, False, "bill month / date missing or do not agree"); continue
            if not any(x["ym"] == ym for x in F["elec_ai"]):
                F["elec_ai"].append(dict(ym=ym, bill_date=bd, amount=amt)); used += 1
            row(e, what, True)
        elif kind == "advance_tax_challan":
            what = "Advance tax Rs %s paid %s" % (amt, paid)
            if not (amt and paid):
                row(e, what, False, "missing figure"); continue
            if not any(x["paid"] == paid and x["amount"] == amt for x in F["advtax"]):
                F["advtax"].append(dict(paid=paid, amount=amt)); used += 1
            row(e, what, True)
        elif kind == "gst_challan":
            what = "GST challan Rs %s deposited %s" % (amt, paid)
            ref = str(e.get("ref") or "%s-%s" % (paid, amt))
            if not (amt and paid):
                row(e, what, False, "missing figure"); continue
            if ref not in {x["cpin"] for x in F["challan"]}:
                F["challan"].append(dict(cpin=ref, gstin=gst, paid=paid, amount=amt, ai=True)); used += 1
            row(e, what, True)
        elif kind == "stock_statement":
            sd = _ai_date(e.get("stock_date"), fy); st = _ai_money(e.get("stock_lakh"))
            what = "Stock statement %s stock %s lakh" % (sd, st)
            if not (sd and st):
                row(e, what, False, "date or stock missing"); continue
            if sd in {s["date"] for s in F["ss"]}:
                row(e, what, False, "date already read by a rule"); continue
            F["ss"].append(dict(date=sd, stock=st, creditors=_ai_money(e.get("creditors_lakh")),
                                debtors=_ai_money(e.get("debtors_lakh")), split=None, check=True, ai=True)); used += 1
            row(e, what, True)
        elif kind == "insurance_policy":
            no = re.sub(r"\s+", "", str(e.get("ref") or "")); si = _ai_money(e.get("stock_si"))
            fr, to = _ai_date(e.get("from_date")), _ai_date(e.get("to_date"))
            what = "Policy %s %s to %s stock SI %s" % (no, fr, to, si)
            p = next((x for x in F["policies"] if x["file"] == fn or (no and x["no"] == no)), None)
            if p is not None:
                if p["stock"] is None and si is not None:
                    p["stock"] = si / 1e5; p["ocr"] = True; p["build"] += " | AI-read stock SI %s - CHECK against the policy" % si
                    used += 1; row(e, what, True)
                else:
                    row(e, what, False, "policy already read by a rule")
                continue
            cls = e.get("policy_class") if e.get("policy_class") in ("Fire", "Burglary", "Other") else None
            if not (no and fr and to and cls and fr < to):
                row(e, what, False, "policy no / period / class missing"); continue
            F["policies"].append(dict(file=fn, no=no, frm=fr.strftime("%d/%m/%Y"), to=to.strftime("%d/%m/%Y"), insured="", gstin="",
                                      prev="", insurer=e.get("insurer") or "", first_loss=None, stock=si / 1e5 if si else None,
                                      build="AI-read - CHECK against the policy", excluded="", sub=None, cls=cls, ocr=True,
                                      items=[("Stock", si / 1e5)] if si else []))
            used += 1; row(e, what, True)
        elif kind in ("gstr3b", "gstr1"):
            k = "g3" if kind == "gstr3b" else "g1"
            vals = [_ai_money(e.get(x)) or 0.0 for x in (("gst_a", "gst_b", "gst_c") if k == "g3" else ("gst_taxable",))]
            fd = _ai_date(e.get("filed_date"), fy)
            what = "%s %s taxable %s filed %s" % (kind.upper(), _mlab(ym) if ym else "?", sum(vals), fd)
            if not ym or not any(vals):
                row(e, what, False, "period or value missing"); continue
            if any(r["ym"] == ym for r in F[k]):
                row(e, what, False, "period already read by a rule"); continue
            r = dict(kind="3B" if k == "g3" else "1", gstin=gst, ym=ym, arn=fd.strftime("%d/%m/%Y") if fd else "", ai=True)
            r.update(taxable=round(vals[0] + vals[1], 2), sales=round(sum(vals), 2)) if k == "g3" else r.update(taxable=vals[0])
            F[k].append(r); used += 1; row(e, what, True)
        else:
            row(e, "-", False, "not a document the Excel uses")
    log.append(("AI", fn, "%s: %d value(s) used - CHECK these on the page" % (kind, used) if used else "%s: nothing used" % kind))
    return browser_summary()


# ================= Month-wise outward supply, TAXABLE VALUE ONLY, all GSTINs (the GST master prompt) =================
def gst_outward_summary(F):
    """Steps of the master prompt, as fixed rules:
    1 combine every outward table, GSTIN tagged (file / column, else the company's own GSTIN from its returns)
    2 taxable value only   3 month = document date (MMM-YYYY)   4 credit note -, debit note +
    5 GSTIN x month sums   6 invalid / blank GSTIN flagged      7 blank / non-numeric taxable skipped and counted
    Duplicate uploads count once. A GSTIN-month with no invoice-level data is taken from the filed GSTR-1 (tax period).
    Every month is also reconciled with the filed GSTR-1 of that period."""
    own = sorted({r["gstin"] for r in F["g1"] + F["g3"] if re.fullmatch(GSTIN_RX, r["gstin"] or "")})
    rows = F.get("out", [])
    # duplicates: the same document line in two sheets / files (an 'SR Combine' sheet, a re-uploaded month) counts once
    per_src = collections.defaultdict(collections.Counter)
    first = {}
    for r in rows:
        g = r["gstin"] if re.fullmatch(GSTIN_RX, r["gstin"] or "") else (own[0] if len(own) == 1 and not r["gstin"] else r["gstin"])
        r["g"] = g
        key = (g, r["doc"], r["no"], r["day"], r["line"], r["value"], r["nontax"])
        per_src[key][r["src"]] += 1
        first.setdefault(key, r)
    data, dups = [], 0
    for key, by in per_src.items():
        n = max(by.values()); dups += sum(by.values()) - n
        data += [first[key]] * n
    invalid = [r for r in data if not re.fullmatch(GSTIN_RX, r["g"] or "")]
    good = [r for r in data if re.fullmatch(GSTIN_RX, r["g"] or "")]
    inv = collections.defaultdict(float); nontax = collections.defaultdict(float)
    for r in good:
        (nontax if r["nontax"] else inv)[(r["g"], r["ym"])] += r["value"]
    ret = {}
    for r in F["g1"]:                                   # filed GSTR-1, one per GSTIN + tax period (latest ARN kept)
        k = (r["gstin"], r["ym"]); q = ret.get(k)
        if q is None or (d(r.get("arn") or "") or dt.date.min) > (d(q.get("arn") or "") or dt.date.min):
            ret[k] = r
    cells, src = {}, {}
    for k, v in inv.items():
        cells[k] = round(v, 2); src[k] = "invoice-level data"
    for k, r in ret.items():
        if k not in cells:
            cells[k] = round(r["taxable"], 2); src[k] = "filed GSTR-1 (tax period)"
    if not F["g1"]:                                     # no GSTR-1 at all (Baid): the filed GSTR-3B 3.1(a)+(b), labelled so
        for r in F["g3"]:
            k = (r["gstin"], r["ym"])
            if k not in cells:
                cells[k] = round(r["taxable"], 2); src[k] = "filed GSTR-3B 3.1(a)+(b) - no GSTR-1 in the folder"
    gstins = sorted({k[0] for k in cells})
    allm = sorted({k[1] for k in cells})
    months = []
    if allm:
        ym = allm[0]
        while ym <= allm[-1]:
            months.append(ym); ym = _ym_add(ym, 1)
    skipped = sum(int(m.group(1)) for n in F.get("out_notes", []) for m in [re.search(r"(\d+) skipped", n)] if m)
    recon = []
    for g in gstins:
        for m in months:
            a = inv.get((g, m)); b = ret.get((g, m))
            if a is None and b is None and (g, m) not in cells:
                continue
            recon.append(dict(gstin=g, ym=m, source=src.get((g, m), "-"), inv=None if a is None else round(a, 2),
                              ret=None if b is None else round(b["taxable"], 2),
                              diff=None if a is None or b is None else round(a - b["taxable"], 2), nontax=round(nontax.get((g, m), 0), 2)))
    # a difference with the filed return: name the document(s) of that month whose taxable value equals it
    docs = collections.defaultdict(float)
    for r in good:
        if not r["nontax"]:
            docs[(r["g"], r["ym"], r["doc"], r["no"])] += r["value"]
    for x in recon:
        if x["diff"] is not None and abs(x["diff"]) > 1:
            hit = ["%s %s (%s)" % (dc, no, "{:,.2f}".format(v)) for (g, m, dc, no), v in docs.items()
                   if g == x["gstin"] and m == x["ym"] and abs(abs(v) - abs(x["diff"])) <= 1]
            x["why"] = ("the difference equals %s - check whether it was filed in this return" % "; ".join(hit[:5])) if hit else \
                "no single document equals the difference - compare this month's documents with the return"
    fixed = [r for r in good if r.get("fixed")]
    return dict(gstins=gstins, months=months, cells=cells, recon=recon, dups=dups, invalid=invalid, skipped=skipped,
                fixed=fixed, notes=F.get("out_notes", []), rows=good)


def _add_gst_taxable_summary(wb, F):
    S = gst_outward_summary(F)
    Font, Fill, Al = openpyxl.styles.Font, openpyxl.styles.PatternFill, openpyxl.styles.Alignment
    head = lambda c: (setattr(c, "font", Font(bold=True, color="FFFFFF")), setattr(c, "fill", Fill("solid", fgColor="173B38")),
                      setattr(c, "alignment", Al(wrap_text=True, vertical="top")))
    ws = wb.create_sheet("GST Taxable Summary")
    ms = S["months"]
    ws.append(["GSTIN"] + [dt.date(y, m, 1).strftime("%b-%Y") for y, m in ms] + ["Total"])
    for c in ws[1]: head(c)
    for g in S["gstins"]:
        vals = [S["cells"].get((g, m), 0.0) for m in ms]
        ws.append([g] + vals + [round(sum(vals), 2)])
    grand = [round(sum(S["cells"].get((g, m), 0.0) for g in S["gstins"]), 2) for m in ms]
    ws.append(["Grand Total"] + grand + [round(sum(grand), 2)])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 20
    for i in range(len(ms) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 2)].width = 15
    ws.freeze_panes = "B2"
    ws.append([])
    ws.append(["Rows skipped: %d (invalid / blank GSTIN: %d; blank or non-numeric taxable value: %d). Duplicate rows counted once: %d. "
               "Date range: %s to %s." % (len(S["invalid"]) + S["skipped"], len(S["invalid"]), S["skipped"], S["dups"],
                                          dt.date(*ms[0], 1).strftime("%b-%Y") if ms else "-", dt.date(*ms[-1], 1).strftime("%b-%Y") if ms else "-")])
    ws.append(["Taxable value only (tax and invoice-value columns not used). Month = invoice / note date; credit notes minus, debit notes plus. "
               "A month with no invoice-level file is taken from the filed GSTR-1 of that tax period - see 'GST Summary Review'."])
    # ---- review: source of every month, reconciliation with the filed GSTR-1, every correction and exclusion ----
    rv = wb.create_sheet("GST Summary Review")
    rv.append(["GSTIN", "Month", "Figure taken from", "Invoice-level total", "Filed GSTR-1 (same tax period)", "Difference", "Check",
               "Nil / exempt / non-GST (not included)"])
    for c in rv[1]: head(c)
    for r in S["recon"]:
        chk = "Ties with the filed GSTR-1" if r["diff"] is not None and abs(r["diff"]) <= 1 else \
            ("DIFFERS from the filed GSTR-1: " + r.get("why", "") if r["diff"] is not None else
             ("No invoice-level file - figure is the " + r["source"] if r["inv"] is None else "No GSTR-1 for this period to compare"))
        rv.append([r["gstin"], dt.date(*r["ym"], 1).strftime("%b-%Y"), r["source"], r["inv"], r["ret"], r["diff"], chk, r["nontax"] or None])
        if r["diff"] is not None and abs(r["diff"]) > 1:
            for c in rv[rv.max_row]:
                c.fill = Fill("solid", fgColor="FCE4D6")
    for row in rv.iter_rows(min_row=2, min_col=4, max_col=8):
        for c in row:
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0.00"
    rv.append([])
    rv.append(["Files and sheets read"]); rv.cell(rv.max_row, 1).font = Font(bold=True)
    for n in S["notes"]:
        rv.append([n])
    if S["fixed"]:
        rv.append([]); rv.append(["Dates corrected (date outside the financial year printed in the document's own number)"])
        rv.cell(rv.max_row, 1).font = Font(bold=True)
        for r in S["fixed"]:
            rv.append(["%s  %s: printed %s -> %s  (%s row %d), taxable %s" % (r["doc"], r["no"], r["fixed"].strftime("%d-%m-%Y"),
                                                                          r["day"].strftime("%d-%m-%Y"), r["src"], r["row"], "{:,.2f}".format(r["value"]))])
    if S["invalid"]:
        rv.append([]); rv.append(["Rows with an invalid / blank GSTIN (NOT in the summary)"]); rv.cell(rv.max_row, 1).font = Font(bold=True)
        for r in S["invalid"][:200]:
            rv.append(["%s | %s | %s | %s row %d" % (r["g"] or "(blank)", r["no"], "{:,.2f}".format(r["value"]), r["src"], r["row"])])
    rv.column_dimensions["A"].width = 20; rv.column_dimensions["B"].width = 10; rv.column_dimensions["C"].width = 26
    for c in "DEF": rv.column_dimensions[c].width = 17
    rv.column_dimensions["G"].width = 44; rv.column_dimensions["H"].width = 18
    rv.freeze_panes = "A2"
    return S


_build_excel_draft_without_taxable_summary = build_excel_draft


def build_excel_draft(F, cfg, out):
    allc = _build_excel_draft_without_taxable_summary(F, cfg, out)
    wb = openpyxl.load_workbook(out)
    _add_gst_taxable_summary(wb, F)
    wb.save(out)
    return allc
