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
    py.runPython("import sys; sys.path.insert(0, '/'); import core");
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
      py.FS.mkdir("/work/" + i);
      py.FS.writeFile(`/work/${i}/${files[i].name}`, new Uint8Array(await files[i].arrayBuffer()));
    }
    status(`Reading ${files.length} files...`);
    await new Promise(r => setTimeout(r, 30));
    const s = JSON.parse(py.globals.get("core").browser_scan("/work"));
    showReview(s);
    status("Read. Check the details below, then Generate Excel.");
  } catch (e) { status("Error while reading: " + e, true); console.error(e); }
  $("readBtn").disabled = false;
}

function showReview(s) {
  const banks = s.banks.map(b => `${esc(b[0])} (${b[1]} txns)`).join(", ") || '<span class="bad">none found</span>';
  $("summary").innerHTML = `Bank: ${banks}<br>GSTR-1: ${s.g1} &nbsp; GSTR-3B: ${s.g3} &nbsp; GSTIN: ${esc(s.gstins.join(", "))}<br>
    Stock statements: ${s.ss} &nbsp; Book stock: ${s.book} &nbsp; Policies: ${s.npol}<br>
    Debtor names: ${s.debtors} &nbsp; Creditor names: ${s.creditors} &nbsp; Directors (MCA): ${esc(s.directors.join(", ") || "none")}`;
  $("own").value = s.borrower; $("gstin").value = s.gstins[0] || ""; $("related").value = s.directors.join("\n");
  $("unreadCard").hidden = !s.unread.length; $("nUnread").textContent = s.unread.length;
  $("unread").innerHTML = "<tr><th>File</th><th>Why</th></tr>" + s.unread.map(u => `<tr><td>${esc(u[0])}</td><td class="small">${esc(u[1])}</td></tr>`).join("");
  $("pol").innerHTML = "<tr><th>Policy</th><th>Class</th><th>Period</th><th>Stock SI (lakh)</th><th>First loss %</th><th>Build-up</th><th>File</th></tr>" +
    s.policies.map(p => `<tr data-no="${esc(p.no)}"><td>${esc(p.no)}</td><td>${esc(p.cls)}</td><td>${esc(p.frm)} to ${esc(p.to)}</td>
      <td><input class="si" value="${p.stock ?? ""}"></td><td><input class="fl" value="${p.fl ?? ""}"></td>
      <td class="small">${esc(p.build)}</td><td class="small">${esc(p.file)}</td></tr>`).join("");
  $("log").innerHTML = "<tr><th>Read as</th><th>File</th><th>Result</th></tr>" + s.log.map(l =>
    `<tr><td>${esc(l[0])}</td><td>${esc(l[1])}</td><td class="${["skipped", "ERROR"].includes(l[0]) ? "bad" : "ok"}">${esc(l[2])}</td></tr>`).join("");
  $("review").hidden = false; $("step1").hidden = true;
}

function makeExcel() {
  const pol = {};
  document.querySelectorAll("#pol tr[data-no]").forEach(r => pol[r.dataset.no] = {si: r.querySelector(".si").value, fl: r.querySelector(".fl").value});
  const form = {own: $("own").value, gstin: $("gstin").value, related: $("related").value, extra: $("extra").value, pol};
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

$("readBtn").onclick = readDocs;
$("makeBtn").onclick = makeExcel;
$("againBtn").onclick = () => { $("review").hidden = true; $("step1").hidden = false; $("folder").value = ""; $("files").value = ""; status("Ready. Pick the Working folder."); };
boot();
