r"""Stock-audit Excel with Claude as the reader (prompt: CLAUDE_PROMPT.md).
The portal's rules (docs/core.py) read every machine-readable file; Claude reads what they cannot (scans, sanction
letters, balance sheets) and writes those facts into claude_facts.json; this script joins both and writes the Excel.

USE:
  python claude_run.py scan "<borrower Working folder>"                  -> what the rules read / could not read
  python claude_run.py make "<Working folder>" claude_facts.json out.xlsx [draft.xlsx]
       claude_facts.json = {"form": {check-page entries}, "ai": {"<file name>": {AI_PROMPT json}}}
       with a draft: the Excel is recalculated and scored against it (figures found / same cell / inputs)."""
import sys, os, json, subprocess
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs"))
import core, openpyxl
core.TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "template.xlsx")


def scan(W):
    s = json.loads(core.browser_scan(W))
    for a, b, c in core._B["log"]:
        print("%-18s %-55s %s" % (a[:18], b[:55], c[:110]))
    print("\nprefill:", json.dumps({k: s.get(k) for k in ("borrower", "gstins", "directors", "period", "sanction", "collateral")}, default=str)[:1500])
    print("\nnot read by rules (Claude reads these):")
    for f in json.loads(core.ai_candidates())["files"]:
        print("  ", f["file"], f["pages"], "pages")
    return s


def recalc(p):
    """Full recalculation in Excel (Windows + Excel only); elsewhere the formulas calculate when the file is opened."""
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        "$x=New-Object -ComObject Excel.Application; $x.DisplayAlerts=$false; $w=$x.Workbooks.Open('%s'); "
                        "$x.CalculateFull(); $w.Save(); $w.Close(); $x.Quit()" % os.path.abspath(p).replace("'", "''")],
                       capture_output=True, timeout=300)
    except Exception as e:
        print("(no Excel here - recalculation skipped: %s; the score counts typed figures only)" % e)


def score(P, R):
    """Every non-empty draft cell: same value in the same cell / anywhere on the sheet; plain-number inputs separately."""
    pv, rv, rf = (openpyxl.load_workbook(P, data_only=True), openpyxl.load_workbook(R, data_only=True), openpyxl.load_workbook(R))
    norm = lambda v: round(v, 2) if isinstance(v, (int, float)) and not isinstance(v, bool) else str(v).strip().lower()
    T = [0] * 6
    for sh in rv.sheetnames:
        if sh not in pv.sheetnames or sh in ("MCA", "Sheet1", "Sheet2", "Sheet3"):
            continue
        vals = {norm(c.value) for row in pv[sh].iter_rows() for c in row if c.value is not None}
        t = [0] * 6
        for row in rv[sh].iter_rows():
            for c in row:
                if c.value is None or str(c.value).strip() == "":
                    continue
                n, p = norm(c.value), pv[sh][c.coordinate].value
                same = p is not None and norm(p) == n
                t[0] += 1; t[1] += same; t[2] += same or n in vals
                fv = rf[sh][c.coordinate].value
                if isinstance(fv, (int, float)) and not isinstance(fv, bool):        # typed number = an input figure
                    t[3] += 1; t[4] += same; t[5] += same or n in vals
        T = [a + b for a, b in zip(T, t)]
        print("  %-20s cells %4d | same cell %3.0f%% | found %3.0f%% || inputs %3d | same %3.0f%% | found %3.0f%%"
              % (sh, t[0], 100 * t[1] / max(t[0], 1), 100 * t[2] / max(t[0], 1), t[3], 100 * t[4] / max(t[3], 1), 100 * t[5] / max(t[3], 1)))
    print("  %-20s cells %4d | same cell %3.0f%% | found %3.0f%% || inputs %3d | same %3.0f%% | found %3.0f%%"
          % ("TOTAL", T[0], 100 * T[1] / T[0], 100 * T[2] / T[0], T[3], 100 * T[4] / T[3], 100 * T[5] / T[3]))


TESS = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def ocr():
    """The website's OCR step (tesseract.js there, desktop Tesseract here) on the scanned files the rules list."""
    if not os.path.exists(TESS):
        return
    png = os.path.join(os.environ.get("TEMP", "."), "_claude_ocr.png")
    for f in json.loads(core.ocr_candidates())["files"]:
        text = []
        for pg in range(f["pages"]):
            core.ocr_render(f["path"], pg, png)
            text.append(subprocess.run([TESS, png, "-", "--psm", "3"], capture_output=True, encoding="utf-8", errors="replace").stdout)
        core.ocr_merge(f["path"], "\n\f\n".join(text), "{}")


def make(W, facts, out, draft=None):
    core.browser_scan(W)
    ocr()
    cf = json.load(open(facts, encoding="utf-8"))
    paths = {os.path.basename(p): p for p in core._B["paths"]}
    for fn, js in (cf.get("ai") or {}).items():            # Claude's reading of the scanned files, checked by the rules
        core.ai_merge(paths[fn], json.dumps(js))
    F = core._B["F"]
    for b in cf.get("books") or []:                         # scanned Tally books Claude read (visit-date stock / debtors / creditors)
        day = core.d(b["date"])
        if b.get("stock_rs") is not None:
            F["book"].append(dict(date=day, stock=b["stock_rs"] / 1e5))
        if b.get("debtors_rs") is not None:
            F["deb_tot"].append(dict(date=day, dr=b["debtors_rs"], cr=0.0))
        if b.get("creditors_rs") is not None:
            F["cred_tot"].append(dict(date=day, dr=0.0, cr=b["creditors_rs"]))
        if b.get("debtor_parties"):                          # no ageing printed -> all < 90 days, as the drafts take it
            F.setdefault("ageing", []).append(dict(date=day, lt90=sum(v for _, v in b["debtor_parties"]), gt90=0.0,
                                                   parties=[(n, v, 0.0) for n, v in b["debtor_parties"]]))
        if b.get("creditor_parties"):
            F.setdefault("cred_parties", {})[day] = [(n, v) for n, v in b["creditor_parties"]]
        core._B["log"].append(("Claude-read books", b.get("file", ""), "%s stock %s debtors %s creditors %s" % (day, b.get("stock_rs"), b.get("debtors_rs"), b.get("creditors_rs"))))
    for a, b, c in core._B["log"]:
        if a.startswith(("AI", "Claude")):
            print("%-18s %-45s %s" % (a, b[:45], c[:120]))
    core.browser_make(json.dumps(cf["form"]), out)
    recalc(out)
    print("saved", out)
    if draft:
        score(out, draft)


if __name__ == "__main__":
    if sys.argv[1] == "scan":
        scan(sys.argv[2])
    else:
        make(*sys.argv[2:6])
