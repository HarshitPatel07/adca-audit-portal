// Stock Audit Portal - browser version. Python (Pyodide) runs docs/core.py on files kept in browser memory only.
const $ = id => document.getElementById(id);
const OK = /\.(pdf|txt|xlsx|xlsm|xls|csv)$/i;
let py = null;

function status(msg, err) { const s = $("status"); s.textContent = msg; s.className = "status" + (err ? " err" : ""); }
function esc(x) { return String(x ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c])); }

async function boot() {
  try {
    py = await loadPyodide({indexURL: "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/"});
    await py.loadPackage(["pymupdf", "micropip"]);
    await py.pyimport("micropip").install(["openpyxl"]);
    const code = await (await fetch("core.py", {cache: "no-cache"})).text();
    py.FS.writeFile("/core.py", code);
    py.FS.writeFile("/template.xlsx", new Uint8Array(await (await fetch("template.xlsx", {cache: "no-cache"})).arrayBuffer()));
    py.runPython("import sys; sys.path.insert(0, '/'); import core; core.TEMPLATE = '/template.xlsx'");
    status("Ready. Pick the Working folder.");
    $("readBtn").disabled = false;
  } catch (e) { status("Could not load the reader: " + e, true); console.error(e); }
}

function rmTree(p) {
  if (!py.FS.analyzePath(p).exists) return;
  for (const n of py.FS.readdir(p)) {
    if (n === "." || n === "..") continue;
    const q = p + "/" + n;
    py.FS.isDir(py.FS.stat(q).mode) ? rmTree(q) : py.FS.unlink(q);
  }
  py.FS.rmdir(p);
}

async function readDocs() {
  const files = [...$("folder").files, ...$("files").files].filter(f => OK.test(f.name));
  if (!files.length) { status("No PDF / TXT / Excel files picked.", true); return; }
  $("readBtn").disabled = true;
  try {
    rmTree("/work"); py.FS.mkdir("/work");
    for (let i = 0; i < files.length; i++) {
      status(`Copying file ${i + 1} of ${files.length} into the reader...`);
      // keep the file's folders (Invoices/Sales bill, Invoices/purchases ...) - the invoice rules read the folder name
      const rel = (files[i].webkitRelativePath || files[i].name).split("/").slice(1, -1)
        .map(x => x.replace(/[^\w .&'()-]/g, "_")).filter(Boolean);
      let dir = "/work/" + i; py.FS.mkdir(dir);
      for (const part of rel) { dir += "/" + part; if (!py.FS.analyzePath(dir).exists) py.FS.mkdir(dir); }
      py.FS.writeFile(`${dir}/${files[i].name}`, new Uint8Array(await files[i].arrayBuffer()));
    }
    status(`Reading ${files.length} files...`);
    await new Promise(r => setTimeout(r, 30));
    let s = JSON.parse(py.globals.get("core").browser_scan("/work"));
    s = await runOCR(s);
    s = await runAI(s);
    showReview(s);
    status("Read. Check the details below, then Generate Excel.");
  } catch (e) { status("Error while reading: " + e, true); console.error(e); }
  $("readBtn").disabled = false;
}

// ---- OCR for scanned pages: tesseract.js runs in this browser (no upload, no key); core.py applies the fixed rules ----
let ocrLib = null;
async function loadOCR() {
  if (!ocrLib) ocrLib = new Promise((res, rej) => {
    const sc = document.createElement("script");
    sc.src = "https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js";
    sc.onload = res; sc.onerror = () => rej(new Error("OCR library did not load")); document.head.appendChild(sc);
  });
  return ocrLib;
}

async function cropRight(blob, line, label) {
  // crop the boxed cell to the right of a label (e.g. "First Loss Percentage"), scaled x2 for small print
  const img = await createImageBitmap(blob), c = document.createElement("canvas");
  const lw = line.words.filter(w => label.split(" ").some(k => w.text.toLowerCase().includes(k.toLowerCase().slice(0, 5))));
  const x0 = lw.length ? Math.max(...lw.map(w => w.bbox.x1)) : line.bbox.x0;
  const y0 = Math.max(0, line.bbox.y0 - 8), h = line.bbox.y1 - line.bbox.y0 + 16, wd = Math.min(img.width - x0, 900);
  c.width = wd * 2; c.height = h * 2; c.getContext("2d").drawImage(img, x0, y0, wd, h, 0, 0, wd * 2, h * 2);
  return new Promise(r => c.toBlob(r));
}

async function runOCR(s) {
  const core = py.globals.get("core");
  const cand = JSON.parse(core.ocr_candidates());
  if (!cand.files.length) return s;
  try { await loadOCR(); } catch (e) { status("OCR not available (" + e.message + ") - scanned pages skipped.", true); return s; }
  const total = cand.files.reduce((a, f) => a + f.pages, 0); let done = 0;
  const worker = await Tesseract.createWorker("eng");
  for (const f of cand.files) {
    let text = ""; const fields = {};
    for (let pg = 0; pg < f.pages; pg++) {
      status(`Reading scanned page ${++done} of ${total} with OCR (${f.file})...`);
      core.ocr_render(f.path, pg, "/ocr.png");
      const blob = new Blob([py.FS.readFile("/ocr.png")], {type: "image/png"});
      await worker.setParameters({tessedit_pageseg_mode: "3", tessedit_char_whitelist: ""});
      const r = await worker.recognize(blob, {}, {blocks: true, text: true});
      text += (pg ? "\n\f\n" : "") + r.data.text;   // pages separated by a form feed (one invoice per page)
      const lines = (r.data.blocks || []).flatMap(b => b.paragraphs.flatMap(p => p.lines));
      for (const label of cand.fields) {
        if (fields[label]) continue;
        const ln = lines.find(l => new RegExp(label.replace(/ /g, "\\s*"), "i").test(l.text));
        if (!ln) continue;
        await worker.setParameters({tessedit_pageseg_mode: "7", tessedit_char_whitelist: "0123456789.,%:"});
        fields[label] = (await worker.recognize(await cropRight(blob, ln, label))).data.text.trim();
      }
    }
    s = JSON.parse(core.ocr_merge(f.path, text, JSON.stringify(fields)));
  }
  await worker.terminate();
  return s;
}

// ---- AI fallback: only files no rule / OCR could read go to Gemini, with the user's own key; core.py checks every value ----
function b64(bytes) {
  let s = ""; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

async function askGemini(key, model, prompt, data) {
  const body = {contents: [{parts: [{inline_data: {mime_type: "application/pdf", data}}, {text: prompt}]}],
                generationConfig: {responseMimeType: "application/json", temperature: 0}};
  for (let attempt = 0; attempt < 5; attempt++) {
    const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`,
      {method: "POST", headers: {"Content-Type": "application/json", "x-goog-api-key": key}, body: JSON.stringify(body)});
    if (r.status === 429 || r.status >= 500) { await new Promise(res => setTimeout(res, 4000 * (attempt + 1))); continue; }
    const j = await r.json();
    if (!r.ok) throw new Error((j.error && j.error.message) || ("HTTP " + r.status));
    return ((j.candidates || [])[0]?.content?.parts || []).map(p => p.text || "").join("");
  }
  throw new Error("Gemini busy (rate limit) - try again in a minute");
}

async function runAI(s) {
  const key = $("aiKey").value.trim(), model = $("aiModel").value;
  try { localStorage.setItem("aiKey", key); localStorage.setItem("aiModel", model); } catch (e) {}
  if (!key) return s;
  const core = py.globals.get("core");
  const cand = JSON.parse(core.ai_candidates());
  let done = 0, failed = [];
  const queue = cand.files.slice();
  async function worker() {
    while (queue.length) {
      const f = queue.shift();
      status(`AI reading ${++done} of ${cand.files.length}: ${f.file} ...`);
      try {
        const text = await askGemini(key, model, cand.prompt, b64(py.FS.readFile(f.path)));
        s = JSON.parse(core.ai_merge(f.path, text));
      } catch (e) { failed.push(f.file + " (" + e.message + ")"); }
    }
  }
  await Promise.all([worker(), worker(), worker()]);
  if (failed.length) status("AI could not read: " + failed.join("; "), true);
  return s;
}

function showAI(s) {
  const rows = s.ai || [];
  $("aiCard").hidden = !rows.length; $("nAi").textContent = rows.length;
  $("ai").innerHTML = "<tr><th>File (page)</th><th>Read as</th><th>Value read</th><th>Check</th></tr>" + rows.map(r =>
    `<tr class="${r.ok ? "ocr" : "missing"}"><td>${esc(r.file)}${r.page ? " (p." + esc(r.page) + ")" : ""}</td><td>${esc(r.kind)}</td>
     <td>${esc(r.what)}</td><td class="small">${r.ok ? "used - check against the document" + (r.why ? "; " + esc(r.why) : "") : "NOT used: " + esc(r.why)}</td></tr>`).join("");
}

function showReview(s) {
  showAI(s);
  const banks = s.banks.map(b => `${esc(b[0])} (${b[1]} txns)`).join(", ") || '<span class="bad">none found</span>';
  $("summary").innerHTML = `Bank: ${banks}<br>GSTR-1: ${s.g1} &nbsp; GSTR-3B: ${s.g3} &nbsp; GSTIN: ${esc(s.gstins.join(", "))}<br>
    Stock statements: ${s.ss} &nbsp; Book stock: ${s.book} &nbsp; Policies: ${s.npol}<br>
    Debtor names: ${s.debtors} &nbsp; Creditor names: ${s.creditors} &nbsp; Directors (MCA): ${esc(s.directors.join(", ") || "none")}`;
  $("own").value = s.borrower; $("gstin").value = s.gstins[0] || ""; $("related").value = s.directors.join("\n");
  $("from").value = (s.period || [])[0] || ""; $("to").value = (s.period || [])[1] || "";
  $("coll").value = s.collateral ?? "";
  $("banks").addEventListener("input", checkWarnings);
  $("unreadCard").hidden = !s.unread.length; $("nUnread").textContent = s.unread.length;
  $("unread").innerHTML = "<tr><th>File</th><th>Why</th></tr>" + s.unread.map(u => `<tr><td>${esc(u[0])}</td><td class="small">${esc(u[1])}</td></tr>`).join("");
  $("pol").innerHTML = "<tr><th>Policy</th><th>Class</th><th>Period</th><th>Stock SI (lakh)</th><th>First loss %</th><th>Build-up</th><th>File</th></tr>" +
    s.policies.map(p => `<tr data-no="${esc(p.no)}" class="${p.ocr ? "ocr" : ""}"><td>${esc(p.no)}</td><td>${esc(p.cls)}</td><td>${esc(p.frm)} to ${esc(p.to)}</td>
      <td><input class="si" value="${p.stock ?? ""}"></td><td><input class="fl" value="${p.fl ?? ""}"></td>
      <td class="small">${esc(p.build)}</td><td class="small">${esc(p.file)}</td></tr>`).join("");
  $("log").innerHTML = "<tr><th>Read as</th><th>File</th><th>Result</th></tr>" + s.log.map(l =>
    `<tr><td>${esc(l[0])}</td><td>${esc(l[1])}</td><td class="${["skipped", "ERROR"].includes(l[0]) || /BREAKS|CHECK/.test(l[2]) ? "bad" : "ok"}">${esc(l[2])}</td></tr>`).join("");
  const iv = (x, k) => `<input class="i-${k}" value="${esc(x[k] ?? "")}">`;
  $("inv").innerHTML = "<tr><th>Type</th><th>Invoice No.</th><th>Date</th><th>Party</th><th>Item</th><th>UOM</th><th>Qty</th><th>Rate</th><th>File (page)</th></tr>" +
    (s.invoices || []).map(x => `<tr data-id="${esc(x.id)}" data-kind="${esc(x.kind)}" class="${x.unread ? "missing" : (x.ocr ? "ocr" : "")}">
      <td>${esc(x.kind)}</td><td>${iv(x, "no")}</td><td>${iv(x, "date")}</td><td>${iv(x, "party")}</td><td>${iv(x, "item")}</td>
      <td>${iv(x, "uom")}</td><td>${iv(x, "qty")}</td><td>${iv(x, "rate")}</td>
      <td class="small">${esc(x.file)} (p.${x.page})${x.unread ? " - not readable, type it from the scan" : (x.ocr ? " - OCR, check" : "")}</td></tr>`).join("");
  const unreadInv = (s.invoices || []).filter(x => x.unread);
  $("invWarn").hidden = !unreadInv.length;
  $("invWarn").textContent = unreadInv.length ? `${unreadInv.length} invoice page(s) could not be read (handwritten / scanned): ` +
    unreadInv.map(x => `${x.file} p.${x.page}`).join(", ") + ". Type Invoice No., Date, Party, Item, UOM, Qty and Rate from the scan." : "";
  const ocrp = s.policies.filter(p => p.ocr).map(p => p.no);
  $("ocrWarn").hidden = !ocrp.length;
  $("ocrWarn").textContent = ocrp.length ? `Read from a scanned page by OCR - check against the policy before generating: ${ocrp.join(", ")} (amber rows).` : "";
  directors = s.directors.map(x => x.trim().toUpperCase());
  document.querySelectorAll("#pol input.si").forEach(i => i.addEventListener("input", checkWarnings));
  $("related").addEventListener("input", checkWarnings);
  checkWarnings();
  $("review").hidden = false; $("step1").hidden = true;
}

let directors = [];

function blankPolicies() {
  const out = [];
  document.querySelectorAll("#pol tr[data-no]").forEach(r => {
    const blank = !r.querySelector(".si").value.trim() && r.children[1].textContent !== "Other";
    r.classList.toggle("missing", blank);
    if (blank) out.push(r.dataset.no + " (" + r.children[1].textContent + ", " + r.children[2].textContent + ")");
  });
  return out;
}

function onlyDirectors() {
  const rel = $("related").value.split("\n").map(x => x.trim().toUpperCase()).filter(Boolean);
  return rel.every(x => directors.includes(x));
}

function checkWarnings() {
  const b = blankPolicies();
  $("polWarn").hidden = !b.length;
  $("polWarn").textContent = b.length ? `${b.length} polic${b.length > 1 ? "ies have" : "y has"} no stock SI: ${b.join("; ")}. ` +
    "Every date in that period will show 'SI not stated'. Type the stock SI (lakh) from the policy schedule." : "";
  $("relWarn").hidden = !onlyDirectors();
  $("bankWarn").hidden = !!$("banks").value.trim();
}

function confirmGenerate() {
  const msgs = [], b = blankPolicies();
  if (b.length) msgs.push(`- ${b.length} policy stock SI blank: ${b.join("; ")}`);
  if (!$("banks").value.trim()) msgs.push("- Banks and sanctioned limits are blank (summation sheet needs them).");
  if (onlyDirectors()) msgs.push("- Related parties: only the MCA directors are listed. Have you added the related parties from the balance sheet note?");
  return !msgs.length || confirm("Before generating, please check:\n\n" + msgs.join("\n\n") + "\n\nOK = generate anyway, Cancel = go back and fill them.");
}

function makeExcel(skipCheck) {
  if (skipCheck !== true && !confirmGenerate()) return;
  const pol = {};
  document.querySelectorAll("#pol tr[data-no]").forEach(r => pol[r.dataset.no] = {si: r.querySelector(".si").value, fl: r.querySelector(".fl").value});
  const inv = {};
  document.querySelectorAll("#inv tr[data-id]").forEach(r => {
    const e = {kind: r.dataset.kind};
    for (const k of ["no", "date", "party", "item", "uom", "qty", "rate"]) e[k] = r.querySelector(".i-" + k).value;
    inv[r.dataset.id] = e;
  });
  const form = {own: $("own").value, gstin: $("gstin").value, related: $("related").value, extra: $("extra").value, pol, inv,
                banks: $("banks").value, from: $("from").value, to: $("to").value, pm: $("pm").value, coll: $("coll").value};
  try {
    status("Making the Excel...");
    const name = py.globals.get("core").browser_make(JSON.stringify(form), "/out.xlsx");
    const bytes = py.FS.readFile("/out.xlsx");
    const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([bytes], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
    a.download = `${name} - GST Insurance Summation ${stamp}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    status("Excel downloaded. You can change entries and generate again.");
  } catch (e) { status("Error while making the Excel: " + e, true); console.error(e); }
}

try { $("aiKey").value = localStorage.getItem("aiKey") || ""; $("aiModel").value = localStorage.getItem("aiModel") || "gemini-flash-latest"; } catch (e) {}
$("readBtn").onclick = readDocs;
$("makeBtn").onclick = makeExcel;
$("againBtn").onclick = () => { $("review").hidden = true; $("step1").hidden = false; $("folder").value = ""; $("files").value = ""; status("Ready. Pick the Working folder."); };
boot();
