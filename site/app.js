/* Houzz → Grist Converter — client-side logic.
 * Faithful port of houzz_to_grist.py (pandas) using SheetJS.
 */

// Column layout from ORDERS.csv (the Grist import template).
const TARGET_COLUMNS = [
  "Proposal", "Ordered", "Item", "Vendor", "QTY", "Unit COST",
  "Markup%", "Markup", "Subtotal", "Pre-Tax", "Shipping", "Total",
  "Prepaid Tax?", "Project Tax Rate", "Tax", "Prepaid Tax Amt", "Created",
  "Notes", "Project_Project Tax Rate", "Received", "URL",
];

// ---- Python helpers, ported 1:1 ---------------------------------------

function cleanCurrency(value) {
  if (value === null || value === undefined || value === "") return 0.0;
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/[\$,]/g, "");
  const m = cleaned.match(/[\d,]+\.?\d*/);
  return m ? parseFloat(m[0].replace(/,/g, "")) : 0.0;
}

function parseMarkupPercent(value) {
  if (value === null || value === undefined || value === "") return 0.0;
  const m = String(value).match(/\((\d+(?:\.\d+)?)%\)/);
  return m ? parseFloat(m[1]) / 100 : 0.0;
}

function titleCase(s) {
  return s.replace(/\w\S*/g, (t) => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
}

function extractVendorFromUrl(url) {
  if (url === null || url === undefined || url === "") return "";
  try {
    let hostname = new URL(String(url)).hostname;
    if (hostname.startsWith("www.")) hostname = hostname.slice(4);
    const vendor = hostname.includes(".") ? hostname.split(".")[0] : hostname;
    return titleCase(vendor);
  } catch {
    return "";
  }
}

function fmtCurrency(x) {
  if (!x) return "";
  const neg = x < 0;
  const abs = Math.abs(x).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return (neg ? "-" : "") + "$" + abs;
}

const pad = (n) => String(n).padStart(2, "0");

function nowOrdered() {
  const d = new Date();
  return `${pad(d.getMonth() + 1)}/${pad(d.getDate())}/${d.getFullYear()}`;
}

function nowCreated() {
  const d = new Date();
  let h = d.getHours();
  const ampm = h >= 12 ? "pm" : "am";
  h = h % 12;
  if (h === 0) h = 12;
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ` +
         `${pad(h)}:${pad(d.getMinutes())}${ampm}`;
}

// ---- Row mapping (get_row_data) ----------------------------------------

function getRowData(proposalNumber, row) {
  const itemName = row["Item"] ?? "";
  const description = row["Description"] ?? "";
  const qty = row["Qty"];
  const cost = cleanCurrency(row["Cost"] ?? 0);
  const markupPct = parseMarkupPercent(row["Markup"] ?? "");
  const unitPrice = cleanCurrency(row["Unit Price"] ?? 0);
  const room = row["Room"] ?? "";
  const shipping = cleanCurrency(row["Shipping"] ?? 0);
  const websiteLink = row["Website Link"] ?? "";

  const markupAmt = markupPct > 0 ? cost * markupPct : unitPrice - cost;
  const qtyInt = (qty !== null && qty !== undefined && qty !== "" && !isNaN(qty))
    ? parseInt(qty, 10) : 1;

  const ordered = nowOrdered();
  const created = nowCreated();

  return {
    Proposal: proposalNumber,
    Item: itemName,
    Product: description,
    Room: room !== null && room !== undefined ? room : "",
    Vendor: extractVendorFromUrl(websiteLink),
    URL: websiteLink,
    Notes: description,
    QTY: qtyInt,
    "Unit COST": fmtCurrency(cost),
    "Unit MSRP": fmtCurrency(unitPrice),
    Cost: fmtCurrency(cost),
    "Markup%": markupPct > 0 ? (markupPct * 100).toFixed(2) + "%" : "",
    Markup: fmtCurrency(markupAmt * qtyInt),
    Shipping: fmtCurrency(shipping),
    Subtotal: "",
    "Pre-Tax": "",
    Total: "",
    Tax: "",
    "Prepaid Tax?": "",
    "Project Tax Rate": "",
    "Prepaid Tax Amt": "",
    "Owed State Tax Amt": "",
    Ordered: ordered,
    Created: created,
    "Created By": "houzz_import",
    Modified: created,
    "Modified By": "houzz_import",
    Received: "",
    "In Houzz": "true",
    Description: description,
  };
}

function standardizeId(v) {
  let s = String(v ?? "").trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s;
}

// ---- Main conversion (convert_houzz_to_grist) --------------------------

function convertHouzzToGrist(workbook, proposalNumber) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  let rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  // Remove parent rows when dotted child rows exist (mirrors '#' filtering).
  if (rows.length > 0 && Object.prototype.hasOwnProperty.call(rows[0], "#")) {
    const stdIds = rows.map((r) => standardizeId(r["#"]));
    const parentsWithChildren = new Set();
    for (const pid of stdIds) {
      if (pid.includes(".")) parentsWithChildren.add(pid.split(".")[0]);
    }
    rows = rows.filter((_r, i) => {
      const id = stdIds[i];
      return !(parentsWithChildren.has(id) && !id.includes("."));
    });
  }

  return rows.map((r) => {
    const rd = getRowData(proposalNumber, r);
    const ordered = {};
    for (const col of TARGET_COLUMNS) {
      ordered[col] = Object.prototype.hasOwnProperty.call(rd, col) ? rd[col] : "";
    }
    return ordered;
  });
}

// ---- Serialization ------------------------------------------------------

function csvEscape(v) {
  const s = String(v ?? "");
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function toCsv(rows) {
  const header = TARGET_COLUMNS.join(",");
  const body = rows
    .map((r) => TARGET_COLUMNS.map((c) => csvEscape(r[c])).join(","))
    .join("\n");
  return header + "\n" + body;
}

function tsvEscape(v) {
  const s = String(v ?? "");
  return /[\t\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function toTsvNoHeader(rows) {
  return rows
    .map((r) => TARGET_COLUMNS.map((c) => tsvEscape(r[c])).join("\t"))
    .join("\n");
}

// Expose for the test harness / browser.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    TARGET_COLUMNS, cleanCurrency, parseMarkupPercent, extractVendorFromUrl,
    fmtCurrency, standardizeId, convertHouzzToGrist, toCsv, toTsvNoHeader,
  };
}

// ---- Browser UI (runs only in the browser) -----------------------------
if (typeof document !== "undefined") {
  const $ = (id) => document.getElementById(id);
  const proposalEl = $("proposal");
  const fileEl = $("file");
  const dropzone = $("dropzone");
  const fileNameEl = $("fileName");
  const hintEl = $("hint");
  const resultsEl = $("results");
  const resultsTitle = $("resultsTitle");
  const previewEl = $("preview");
  const copyBtn = $("copyBtn");

  let currentFile = null;
  let currentRows = [];
  let currentTsv = "";

  function setHint(text, isError) {
    hintEl.textContent = text;
    hintEl.classList.toggle("error", !!isError);
  }

  async function runConversion() {
    if (!currentFile) return;
    const proposal = proposalEl.value.trim();
    if (!proposal) {
      setHint("⚠️ Enter a Proposal # first.", true);
      return;
    }
    try {
      setHint("Converting…");
      const buf = await currentFile.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      currentRows = convertHouzzToGrist(wb, proposal);
      renderResults(currentRows);
      setHint("Conversion successful — " + currentRows.length + " rows.");
    } catch (err) {
      setHint("Error during conversion: " + (err && err.message ? err.message : err), true);
      resultsEl.classList.add("hidden");
    }
  }

  function renderResults(rows) {
    resultsTitle.textContent = "Result — " + rows.length + " rows";

    // Preview table
    previewEl.innerHTML = "";
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    for (const c of TARGET_COLUMNS) {
      const th = document.createElement("th");
      th.textContent = c;
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    previewEl.appendChild(thead);

    const tbody = document.createElement("tbody");
    const maxRows = 200;
    for (const r of rows.slice(0, maxRows)) {
      const tr = document.createElement("tr");
      for (const c of TARGET_COLUMNS) {
        const td = document.createElement("td");
        td.textContent = r[c] ?? "";
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    previewEl.appendChild(tbody);

    currentTsv = toTsvNoHeader(rows);
    resultsEl.classList.remove("hidden");
  }

  fileEl.addEventListener("change", () => {
    if (fileEl.files && fileEl.files[0]) {
      currentFile = fileEl.files[0];
      fileNameEl.textContent = currentFile.name;
      runConversion();
    }
  });

  proposalEl.addEventListener("input", () => {
    if (currentFile) runConversion();
  });

  dropzone.addEventListener("click", () => fileEl.click());
  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("drag");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("drag"));
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("drag");
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) {
      currentFile = f;
      fileNameEl.textContent = f.name;
      runConversion();
    }
  });

  copyBtn.addEventListener("click", async () => {
    const old = copyBtn.textContent;
    try {
      await navigator.clipboard.writeText(currentTsv);
    } catch {
      const ta = document.createElement("textarea");
      ta.value = currentTsv;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
    }
    copyBtn.textContent = "✅ Copied!";
    setTimeout(() => (copyBtn.textContent = old), 1500);
  });
}
