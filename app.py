r"""AUDIT PORTAL server (local PC or Render). All reading rules live in docs/core.py - the same file the GitHub Pages
version runs inside the browser, so both always give the same Excel.
RUN:   python app.py   -> http://127.0.0.1:5055
"""
import os, sys, io, re, json, datetime as dt
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs"))
from core import *  # noqa: F401,F403 - scan, build_excel, num, d ...

# ---------------- web portal ----------------
import uuid, tempfile, shutil, html, secrets, functools, time, webbrowser, threading
from flask import Flask, request, send_file, session, redirect
APP = Flask(__name__)
APP.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
APP.config["MAX_CONTENT_LENGTH"] = 400 * 1024 * 1024
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")          # set on the live server; local runs need no login
HOSTED = bool(os.environ.get("RENDER") or os.environ.get("HOSTED"))
STORE = {}                                                  # tok -> dict(F, log, paths, tmp, t)
OUTDIR = os.path.join(tempfile.gettempdir(), "audit_portal_out") if HOSTED else \
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "Output")
CSS = """<style>body{font-family:Segoe UI,Arial;margin:0;background:#f4f6f9;color:#1d2733}main{max-width:1100px;margin:auto;padding:24px 16px}
h1{font-size:22px}h2{font-size:17px;margin-top:28px}.card{background:#fff;border:1px solid #dde3ea;border-radius:8px;padding:18px;margin:14px 0}
label{display:block;font-weight:600;margin:12px 0 4px}input[type=text],input[type=password],textarea{width:100%;box-sizing:border-box;padding:8px;border:1px solid #c8d0da;border-radius:6px;font:inherit}
textarea{min-height:70px}button{background:#1f5fbf;color:#fff;border:0;border-radius:6px;padding:10px 18px;font-size:15px;cursor:pointer;margin-top:14px}
button.alt{background:#5b6573}table{border-collapse:collapse;width:100%;font-size:13px}td,th{border:1px solid #dde3ea;padding:5px 7px;text-align:left;vertical-align:top}
th{background:#e8eef7}.ok{color:#1a7f37}.bad{color:#b42318}.small{font-size:12px;color:#5b6573}td input{width:100%;box-sizing:border-box;border:1px solid #c8d0da;padding:4px}
.wrap{overflow-x:auto}#busy{display:none;color:#1f5fbf;font-weight:600;margin-top:10px}</style>"""
BUSY_JS = "<script>document.querySelectorAll('form').forEach(f=>f.addEventListener('submit',()=>{const b=document.getElementById('busy');if(b)b.style.display='block'}))</script>"


def page(body):
    return ("<!doctype html><html><head><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'>"
            "<title>Audit Portal</title>%s</head><body><main>%s<p id=busy>Working... large folders and AI reading can take a few minutes.</p>"
            "</main>%s</body></html>") % (CSS, body, BUSY_JS)


def esc(x):
    return html.escape(str(x if x is not None else ""))


def cleanup(old=3600):
    for k, v in list(STORE.items()):
        if time.time() - v["t"] > old:
            if v.get("tmp"):
                shutil.rmtree(v["tmp"], ignore_errors=True)
            STORE.pop(k, None)


@APP.before_request
def guard():
    if ADMIN_PASSWORD and not session.get("authed") and request.path not in ("/login", "/health"):
        return redirect("/login")


@APP.route("/login", methods=["GET", "POST"])
def login():
    err = ""
    if request.method == "POST":
        if ADMIN_PASSWORD and secrets.compare_digest(request.form.get("password", ""), ADMIN_PASSWORD):
            session["authed"] = True
            return redirect("/")
        err = "<p class=bad>Wrong password.</p>"
    return page("<h1>Audit Portal</h1><div class=card><form method=post><label>Password</label>"
                "<input type=password name=password autofocus>%s<button>Log in</button></form></div>" % err)


@APP.get("/logout")
def logout():
    session.clear()
    return redirect("/login")


@APP.get("/health")
def health():
    return "ok"


@APP.get("/")
def home():
    folder = "" if HOSTED else """<label>Borrower WORKING folder path on this PC (reads every file in every sub-folder)</label>
<input type=text name=folder placeholder="C:\\Users\\DELL\\Downloads\\BORROWER\\Working"><label>- or upload</label>"""
    return page("""<h1>Stock Audit Portal - GST, Insurance, Summation</h1>
<div class=card><form method=post action=/read enctype=multipart/form-data id=up>%s
<label>Upload the whole Working folder</label><input type=file name=files multiple webkitdirectory>
<label>- or pick individual files</label><input type=file name=files2 multiple>
<p class=small>Used: PDF, TXT, XLSX, XLS, CSV (photos are skipped automatically). Nothing is kept after the Excel is made.</p>
<button>Read documents</button></form></div>
<p class=small>Rules only - every figure comes from a fixed rule on the document. %s</p>
<script>
document.getElementById('up').addEventListener('submit', async e => {
  const f = e.target, ok = /\\.(pdf|txt|xlsx|xlsm|xls|csv)$/i, fd = new FormData();
  if (f.folder) fd.append('folder', f.folder.value);
  for (const n of ['files','files2']) for (const x of f[n].files) if (ok.test(x.name)) fd.append(n, x, x.webkitRelativePath || x.name);
  e.preventDefault(); document.getElementById('busy').style.display='block';
  const r = await fetch('/read', {method:'POST', body:fd}); document.open(); document.write(await r.text()); document.close();
});
</script>""" % (folder, "<a href=/logout>log out</a>" if ADMIN_PASSWORD else ""))


@APP.post("/read")
def read():
    cleanup()
    paths, tmp = [], None
    folder = "" if HOSTED else request.form.get("folder", "").strip().strip('"')
    if folder and os.path.isdir(folder):
        for root, _, fs in os.walk(folder):
            paths += [os.path.join(root, f) for f in fs]
    ups = [f for f in request.files.getlist("files") + request.files.getlist("files2") if f.filename]
    if ups:
        tmp = tempfile.mkdtemp(prefix="portal_")
        for i, f in enumerate(ups):
            p = os.path.join(tmp, "%03d_%s" % (i, os.path.basename(f.filename.replace("\\", "/")))); f.save(p); paths.append(p)
    if not paths:
        return page("<h1>No files found</h1><p>Check the folder path or pick files.</p><a href=/>Back</a>")
    log = []
    F = scan(sorted(paths), log)
    tok = uuid.uuid4().hex[:12]
    STORE[tok] = dict(F=F, log=log, paths={os.path.basename(p): p for p in paths}, tmp=tmp, t=time.time())
    return review(tok)


def review(tok):
    S = STORE[tok]; F = S["F"]; log = S["log"]
    gst = sorted({r["gstin"] for r in F["g1"] + F["g3"]})
    borrower = next(iter(sorted(n for n in F["names"] if n)), "")
    unk = [(b, c) for a, b, c in log if a == "skipped" and b.lower().rsplit(".", 1)[-1] in ("pdf", "xlsx", "xlsm", "xls", "csv")]
    unk += [(p["file"], "policy read, but stock SI not in the text layer (scanned part) - type it in the policy table below")
            for p in F["policies"] if p["stock"] is None]
    lg = "".join("<tr><td>%s</td><td>%s</td><td class=%s>%s</td></tr>" % (esc(a), esc(b), "bad" if a in ("skipped", "ERROR") else "ok", esc(c)) for a, b, c in log)
    ai = ""
    if unk:
        rows = "".join("<tr><td>%s</td><td class=small>%s</td></tr>" % (esc(b), esc(c)) for b, c in unk)
        ai = """<div class=card><h2>Not read by any rule (%d)</h2>
<p class=small>If any of these is a bank statement, GST return, policy or stock statement, its figures are NOT in the Excel.
Send that file format to be added as a rule, or type the policy SI below.</p>
<div class=wrap><table><tr><th>File</th><th>Why</th></tr>%s</table></div></div>""" % (len(unk), rows)
    pol = "".join("""<tr><td>%s<input type=hidden name=pno value="%s"></td><td>%s</td><td>%s to %s</td><td><input name=psi value="%s"></td>
<td><input name=pfl value="%s"></td><td class=small>%s</td><td class=small>%s</td></tr>""" % (
        esc(p["no"]), esc(p["no"]), esc(p["cls"]), esc(p["frm"]), esc(p["to"]), "" if p["stock"] is None else p["stock"],
        p["first_loss"] or "", esc(p["build"]), esc(p["file"])) for p in F["policies"])
    banks = ", ".join("%s (%d txns)" % (esc(k), len(v)) for k, v in F["bank"].items()) or "<span class=bad>none found</span>"
    return page("""<h1>Documents read</h1>
<div class=card>Bank: %s<br>GSTR-1: %d &nbsp; GSTR-3B: %d &nbsp; GSTIN: %s<br>Stock statements: %d &nbsp; Book stock: %d &nbsp; Policies: %d<br>
Debtor names: %d &nbsp; Creditor names: %d &nbsp; Directors (MCA): %s</div>%s
<form method=post action=/make><input type=hidden name=tok value=%s>
<div class=card><h2>Check before generating</h2>
<label>Borrower name(s) - own-name credits = self transfer (one per line)</label><textarea name=own>%s</textarea>
<label>Borrower GSTIN</label><input type=text name=gstin value="%s">
<label>Directors / partners / related parties - their credits are excluded (one per line; add the related-party note of the audited balance sheet)</label>
<textarea name=related>%s</textarea>
<label>Extra dates for insurance (visit date etc.): one per line as  dd-mm-yyyy, label, stock in lakh (stock optional)</label>
<textarea name=extra placeholder="30-09-2026, Visit date 30-09-2026, 117.13"></textarea>
<h2>Policies - correct the stock SI (lakh) or first-loss %% if blank or wrong</h2>
<div class=wrap><table><tr><th>Policy</th><th>Class</th><th>Period</th><th>Stock SI (lakh)</th><th>First loss %%</th><th>Build-up</th><th>File</th></tr>%s</table></div>
<button>Generate Excel</button></div></form>
<h2>Every file</h2><div class=wrap><table><tr><th>Read as</th><th>File</th><th>Result</th></tr>%s</table></div><p><a href=/>Start again</a></p>""" % (
        banks, len(F["g1"]), len(F["g3"]), esc(", ".join(gst)), len(F["ss"]), len(F["book"]), len(F["policies"]),
        len(F["debtors"]), len(F["creditors"]), esc(", ".join(F["directors"]) or "none"), ai, tok, esc(borrower),
        esc(gst[0] if gst else ""), esc("\n".join(F["directors"])), pol, lg))


@APP.post("/make")
def make():
    tok = request.form.get("tok", ""); S = STORE.get(tok)
    if S is None:
        return page("<h1>Session expired</h1><a href=/>Start again</a>")
    F = S["F"]
    lines = lambda k: [x.strip() for x in request.form.get(k, "").splitlines() if x.strip()]
    for no, si, fl in zip(request.form.getlist("pno"), request.form.getlist("psi"), request.form.getlist("pfl")):
        for p in F["policies"]:
            if p["no"] == no:
                if si.strip():
                    if p["stock"] != num(si): p["build"] += " | SI entered on portal: %s" % si
                    p["stock"] = num(si)
                p["first_loss"] = int(num(fl)) if fl.strip() else None
    extra = []
    for l in lines("extra"):
        parts = [x.strip() for x in l.split(",")]
        if d(parts[0]):
            extra.append(dict(date=d(parts[0]), label=parts[1] if len(parts) > 1 and parts[1] else None,
                              stock=num(parts[2]) if len(parts) > 2 and parts[2] else None))
    own = lines("own")
    cfg = dict(own=own, borrower=own[0] if own else "", gstin=request.form.get("gstin", "").strip(), related=lines("related"), extra=extra)
    os.makedirs(OUTDIR, exist_ok=True)
    out = os.path.join(OUTDIR, "%s - GST Insurance Summation %s.xlsx" % (re.sub(r"[^\w .-]", "", cfg["borrower"] or "Borrower")[:60].strip(), dt.datetime.now().strftime("%Y%m%d-%H%M%S")))
    build_excel(F, cfg, out)
    data = io.BytesIO(open(out, "rb").read())
    if HOSTED:                                   # nothing stays on the server
        os.remove(out)
        if S.get("tmp"): shutil.rmtree(S["tmp"], ignore_errors=True)
        if S.get("tmp") and os.path.isdir(S["tmp"]): APP.logger.warning("temp folder not fully deleted: %s", S["tmp"])
        STORE.pop(tok, None)
    return send_file(data, as_attachment=True, download_name=os.path.basename(out),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


app = APP  # gunicorn app:app

if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--cli":  # python app.py --cli "<Working folder>" "Related;Names"
        paths = [os.path.join(r, f) for r, _, fs in os.walk(sys.argv[2]) for f in fs]
        log = []; F = scan(sorted(paths), log)
        for x in log: print(" | ".join(x))
        own = sorted(n for n in F["names"] if n); gst = sorted({r["gstin"] for r in F["g1"] + F["g3"]})
        rel = F["directors"] + (sys.argv[3].split(";") if len(sys.argv) > 3 else [])
        os.makedirs(OUTDIR, exist_ok=True); out = os.path.join(OUTDIR, "cli_test.xlsx")
        build_excel(F, dict(own=own, borrower=own[0] if own else "", gstin=gst[0] if gst else "", related=rel, extra=[]), out)
        print("WROTE", out)
    else:
        port = int(os.environ.get("PORT", 5055))
        if not HOSTED:
            threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:%d" % port)).start()
        APP.run(host="0.0.0.0" if HOSTED else "127.0.0.1", port=port)
