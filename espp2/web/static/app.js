/* TaxLedger Pro — Frontend SPA
   Views: reports (input), processing, results.
   Communicates with the FastAPI backend via fetch(). */

"use strict";

// ── State ──────────────────────────────────────────────────────────────────
const state = {
  currentView: "reports",
  settings: { year: new Date().getFullYear() - 1, rates_loaded: false, rates_days: 0 },
  reportingPdfFile: null,
  priorPdfFile: null,
  wireRows: [],        // [{id, date, usd, nok, fromPdf}]
  jobId: null,
  pollTimer: null,
  lastResult: null,
  lastJobId: null,
  exports: [],
};

// ── Utilities ──────────────────────────────────────────────────────────────
function fmt(n, decimals = 0) {
  if (n === null || n === undefined || n === "—") return "—";
  return Number(n).toLocaleString("nb-NO", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function show(id) { document.getElementById(id).classList.remove("hidden"); }
function hide(id) { document.getElementById(id).classList.add("hidden"); }
function setHtml(id, html) { document.getElementById(id).innerHTML = html; }
function setText(id, text) { document.getElementById(id).textContent = text; }
function getEl(id) { return document.getElementById(id); }

// ── Navigation ─────────────────────────────────────────────────────────────
function showView(name) {
  ["reports", "processing", "results"].forEach(v => {
    const el = getEl(`view-${v}`);
    if (el) el.classList.add("hidden");
  });
  const target = getEl(`view-${name}`);
  if (target) target.classList.remove("hidden");

  document.querySelectorAll(".nav-tab").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.view === name);
  });
  state.currentView = name;
}

document.querySelectorAll(".nav-tab").forEach(btn => {
  btn.addEventListener("click", () => showView(btn.dataset.view));
});

// ── Year Selector & Rate Fetching ──────────────────────────────────────────
function initYearSelector(savedYear) {
  const sel = getEl("reporting-year-select");
  if (!sel) return;
  const thisYear = new Date().getFullYear();
  for (let y = thisYear - 1; y >= thisYear - 6; y--) {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
    if (y === savedYear) opt.selected = true;
    sel.appendChild(opt);
  }
  sel.addEventListener("change", async () => {
    const year = parseInt(sel.value);
    state.settings.year = year;
    updateYearLabels(year);
    await fetchRates(year);
  });
}

async function fetchRates(year) {
  const statusEl = getEl("rates-status-inline");
  if (statusEl) { statusEl.textContent = "Loading rates…"; statusEl.className = "rates-status loading"; }
  const fd = new FormData();
  fd.append("year", year);
  try {
    const res = await fetch("/api/fetch-rates", { method: "POST", body: fd });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const d = await res.json();
    state.settings.rates_loaded = true;
    state.settings.rates_days = d.days_loaded;
    if (statusEl) {
      statusEl.textContent = `✓ ${d.days_loaded} rate days loaded`;
      statusEl.className = "rates-status ok";
    }
  } catch (e) {
    console.error("fetchRates:", e);
    if (statusEl) { statusEl.textContent = "⚠ Could not load rates"; statusEl.className = "rates-status error"; }
  }
}

function updateYearLabels(year) {
  const sub = getEl("reporting-pdf-subtitle");
  if (sub) sub.textContent = `Full annual statement from StockPlanConnect (Jan 1–Dec 31 ${year})`;
  const dzRep = getEl("dz-reporting-label");
  if (dzRep && !state.reportingPdfFile) dzRep.textContent = `Drop your ${year} Morgan Stanley PDF or click to select`;
  const dzPrior = getEl("dz-prior-label");
  if (dzPrior && !state.priorPdfFile) dzPrior.textContent = `Drop your ${year - 1} December monthly statement or click to select`;
}

// ── Drop Zone helper ───────────────────────────────────────────────────────
function setupDropZone(zoneId, inputId, onFile) {
  const zone = getEl(zoneId);
  const input = getEl(inputId);
  if (!zone || !input) return;

  input.addEventListener("change", () => { if (input.files[0]) onFile(input.files[0]); });
  // Stop the input's click from bubbling to zone (would trigger input.click() a second time)
  input.addEventListener("click", e => e.stopPropagation());
  zone.addEventListener("click", () => input.click());
  zone.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") input.click(); });
  zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  });
}

// ── PDF Upload Zones ───────────────────────────────────────────────────────
setupDropZone("dz-reporting", "input-reporting-pdf", async (file) => {
  state.reportingPdfFile = file;
  const dz = getEl("dz-reporting");
  if (dz) dz.classList.add("has-file");
  const lbl = getEl("dz-reporting-label");
  if (lbl) lbl.textContent = file.name;
  const fn = getEl("dz-reporting-filename");
  if (fn) { fn.textContent = `✓ ${file.name}`; fn.classList.remove("hidden"); }

  // Auto-populate wire table from PDF WIRE entries, and auto-detect reporting year
  try {
    const fd = new FormData();
    fd.append("reporting_year_pdf", file);
    const res = await fetch("/api/parse-wires", { method: "POST", body: fd });
    if (res.ok) {
      const data = await res.json();

      // Auto-set year from PDF if detected and different from current
      if (data.year && data.year !== state.settings.year) {
        state.settings.year = data.year;
        const sel = getEl("reporting-year-select");
        if (sel) {
          // Add year option if not already present
          if (!sel.querySelector(`option[value="${data.year}"]`)) {
            const opt = document.createElement("option");
            opt.value = data.year;
            opt.textContent = data.year;
            sel.insertBefore(opt, sel.firstChild);
          }
          sel.value = data.year;
        }
        updateYearLabels(data.year);
        await fetchRates(data.year);
      }

      // Pre-populate wire rows
      if (data.wires && data.wires.length > 0) {
        wireIdCounter = 0;
        state.wireRows = data.wires.map(w => ({
          id: ++wireIdCounter,
          date: w.date,
          usd: w.usd,
          nok: "",
          fromPdf: true,
        }));
        renderWireRows();
        const hint = getEl("wire-pdf-hint");
        if (hint) {
          hint.textContent = `${data.wires.length} wire transfer${data.wires.length !== 1 ? "s" : ""} found in PDF — enter the NOK amount for each from your bank statement.`;
          hint.classList.remove("hidden");
        }
      }
    }
  } catch (e) {
    console.error("parse-wires:", e);
  }
});

setupDropZone("dz-prior", "input-prior-pdf", (file) => {
  state.priorPdfFile = file;
  const dz = getEl("dz-prior");
  if (dz) dz.classList.add("has-file");
  const lbl = getEl("dz-prior-label");
  if (lbl) lbl.textContent = file.name;
  const fn = getEl("dz-prior-filename");
  if (fn) { fn.textContent = `✓ ${file.name}`; fn.classList.remove("hidden"); }
});

// ── Wire Rows ──────────────────────────────────────────────────────────────
let wireIdCounter = 0;

function addWireRow(date = "", usd = "", nok = "") {
  const id = ++wireIdCounter;
  state.wireRows.push({ id, date, usd, nok, fromPdf: false });
  renderWireRows();
  setTimeout(() => {
    const inp = document.querySelector(`[data-wire-id="${id}"][data-wire-field="date"]`);
    if (inp) inp.focus();
  }, 50);
}

function removeWireRow(id) {
  state.wireRows = state.wireRows.filter(r => r.id !== id);
  renderWireRows();
}

function renderWireRows() {
  const tbody = getEl("wire-tbody");
  if (!tbody) return;
  tbody.innerHTML = state.wireRows.map(r => {
    const usdAttrs = r.fromPdf
      ? `readonly class="wire-usd-readonly" tabindex="-1"`
      : `data-wire-id="${r.id}" data-wire-field="usd"`;
    return `
    <tr>
      <td><input type="date" value="${r.date}" data-wire-id="${r.id}" data-wire-field="date" /></td>
      <td><input type="number" step="0.01" placeholder="0.00" value="${r.usd || ""}" ${usdAttrs} style="text-align:right" /></td>
      <td><input type="number" step="0.01" placeholder="0.00" value="${r.nok || ""}" data-wire-id="${r.id}" data-wire-field="nok" style="text-align:right" /></td>
      <td><button class="btn btn-danger btn-sm" onclick="removeWireRow(${r.id})">✕</button></td>
    </tr>`;
  }).join("");

  tbody.querySelectorAll("input[data-wire-id]").forEach(inp => {
    inp.addEventListener("input", () => {
      const id = parseInt(inp.dataset.wireId);
      const field = inp.dataset.wireField;
      const row = state.wireRows.find(r => r.id === id);
      if (row) {
        row[field] = inp.value;
        if (field === "nok") { inp.style.outline = ""; updateWireTotal(); }
      }
    });
    inp.addEventListener("blur", updateWireTotal);
  });

  updateWireTotal();
}

function updateWireTotal() {
  let total = 0;
  state.wireRows.forEach(r => { total += parseFloat(r.nok || "0") || 0; });
  setText("wire-total", fmt(total, 2) + " kr");
}

getEl("btn-add-wire")?.addEventListener("click", () => addWireRow());

// ── Submit ─────────────────────────────────────────────────────────────────
getEl("btn-submit")?.addEventListener("click", async () => {
  if (!state.reportingPdfFile) {
    showSubmitError("Please upload the reporting-year PDF statement.");
    return;
  }

  hide("submit-error");
  showView("processing");
  updateProgress(1, 2, "Uploading files…");

  // Validate wire rows: every row that has a date or USD must also have NOK > 0
  const incompleteWires = state.wireRows.filter(r => {
    const hasData = r.date || parseFloat(r.usd) > 0;
    const nokVal = parseFloat(r.nok) || 0;
    return hasData && nokVal <= 0;
  });
  if (incompleteWires.length > 0) {
    showSubmitError(
      `Please enter the NOK amount received for ${incompleteWires.length === 1 ? "the highlighted wire transfer" : "all highlighted wire transfers"} before calculating. ` +
      "You can find this in your bank statement (look for incoming international transfers on the same dates)."
    );
    // Highlight the incomplete rows
    incompleteWires.forEach(r => {
      const input = document.querySelector(`[data-wire-id="${r.id}"][data-wire-field="nok"]`);
      if (input) { input.style.outline = "2px solid #e74c3c"; input.focus(); }
    });
    return;
  }

  const wires = state.wireRows
    .filter(r => r.date && (parseFloat(r.usd) > 0 || parseFloat(r.nok) > 0))
    .map(r => ({ date: r.date, usd: parseFloat(r.usd) || 0, nok: parseFloat(r.nok) || 0 }));

  const fd = new FormData();
  fd.append("reporting_year_pdf", state.reportingPdfFile);
  if (state.priorPdfFile) fd.append("prior_year_pdf", state.priorPdfFile);
  fd.append("wires", JSON.stringify(wires));
  fd.append("year", state.settings.year);

  try {
    const res = await fetch("/api/process", { method: "POST", body: fd });
    if (!res.ok) {
      showView("reports");
      showSubmitError(`Server error: ${await res.text()}`);
      return;
    }
    const data = await res.json();
    state.jobId = data.job_id;
    startPolling(data.job_id);
  } catch (e) {
    showView("reports");
    showSubmitError(`Network error: ${e.message}`);
  }
});

function showSubmitError(msg) {
  const el = getEl("submit-error");
  const msgEl = getEl("submit-error-msg");
  if (el && msgEl) { msgEl.textContent = msg; el.classList.remove("hidden"); }
}

// ── Polling ────────────────────────────────────────────────────────────────
function startPolling(jobId) {
  if (state.pollTimer) clearInterval(state.pollTimer);
  state.pollTimer = setInterval(async () => {
    try {
      const res = await fetch(`/api/job/${jobId}`);
      if (!res.ok) return;
      const data = await res.json();
      updateProgress(data.phase, data.pct, data.message);
      if (data.status === "done") {
        clearInterval(state.pollTimer);
        await loadResults(jobId);
      } else if (data.status === "error") {
        clearInterval(state.pollTimer);
        showView("reports");
        showSubmitError(`Calculation failed: ${data.message}`);
      }
    } catch (e) { console.error("poll error:", e); }
  }, 600);
}

function updateProgress(phase, pct, message) {
  const phaseLabels = ["", "PARSING PDF STATEMENTS", "COMPUTING OPENING BALANCE", "RUNNING TAX CALCULATIONS", "GENERATING REPORT"];
  setText("progress-phase-label", phaseLabels[phase] || `PROCESSING PHASE ${phase} OF 4`);
  setText("progress-message", message);
  setText("progress-pct", `${pct}%`);
  const bar = getEl("progress-bar");
  if (bar) bar.style.width = `${pct}%`;
}

// ── Results ────────────────────────────────────────────────────────────────
async function loadResults(jobId) {
  try {
    const res = await fetch(`/api/results/${jobId}`);
    if (!res.ok) return;
    const data = await res.json();
    state.lastResult = data;
    state.lastJobId = jobId;
    renderResults(data);
    showView("results");
  } catch (e) {
    showView("reports");
    showSubmitError(`Failed to load results: ${e.message}`);
  }
}

function renderResults(data) {
  setText("results-year", data.year);

  const table = getEl("results-table");
  if (!table) return;

  let rows = "";
  const addSection = (label) => { rows += `<tr class="results-section-header"><td colspan="2">${label}</td></tr>`; };
  const addRow = (label, value, cssClass = "") => {
    const cls = cssClass ? ` class="${cssClass}"` : "";
    rows += `<tr><td>${label}</td><td${cls}>${value}</td></tr>`;
  };
  const addBlank = () => { rows += `<tr><td colspan="2" style="height:8px;background:var(--surface-container-low)"></td></tr>`; };

  (data.foreignshares || []).forEach(fs => {
    addSection(`Utenlandske aksjer — ${fs.symbol}`);
    addRow("Land", fs.country || "USA");
    addRow("Kontofører/bank", data.broker || "Morgan Stanley");
    if (data.account_id) addRow("Kontonummer", data.account_id);
    addRow("Navn på aksjeselskap", companyName(fs.symbol));
    addRow("ISIN", fs.isin || "—");
    addRow("Antall aksjer per 31. desember", fmt(fs.shares, 4));
    addBlank();

    addSection("Formue (markedsverdi per 31. desember)");
    addRow("Valuta", "USD – Amerikansk dollar");
    addRow("Beløp (USD)", fmt(fs.wealth_usd, 2) + " USD");
    addRow("Beløp i NOK", fmt(fs.wealth_nok) + " NOK");
    addRow("Valutakurs (NOK/USD)", fmt(fs.exchange_rate, 4));
    addBlank();

    addSection("Inntekt og fradrag");
    addRow("Skattepliktig utbytte (NOK)", fmt(fs.dividend_nok) + " NOK");
    const gain = fs.taxable_gain_nok;
    if (gain >= 0) {
      addRow("Skattepliktig gevinst (NOK)", fmt(gain) + " NOK", "val-positive");
      addRow("Fradragsberettiget tap (NOK)", "—", "val-empty");
    } else {
      addRow("Skattepliktig gevinst (NOK)", "—", "val-empty");
      addRow("Fradragsberettiget tap (NOK)", fmt(Math.abs(gain)) + " NOK", "val-negative");
    }
    addRow("Anvendt skjerming (NOK)", fmt(fs.tax_deduction_used_nok) + " NOK");
    addBlank();

    if (fs.credit_deduction) {
      const cd = fs.credit_deduction;
      addSection("Kreditfradrag — dobbeltbeskatning");
      addRow("Land", cd.country || "USA");
      addRow("Inntektsskatt betalt i utlandet (NOK)", fmt(cd.income_tax_nok) + " NOK");
      addRow("Brutto aksjeutbytte (NOK)", fmt(cd.gross_dividend_nok) + " NOK");
      addRow("Herav skatt på brutto aksjeutbytte (NOK)", fmt(cd.tax_on_gross_nok) + " NOK");
      addBlank();
    }
  });

  addSection("Valutakursgevinst/-tap — USD-konto");

  // Type A: aggregated FX gain already included in share taxable_gain above
  const fxAgg = data.fx_gain_aggregated_nok ?? 0;
  const fxAggLabel = fxAgg >= 0
    ? "Type A — Aggregert valutagevinst (NOK)"
    : "Type A — Aggregert valutatap (NOK)";
  const fxAggNote = "Allerede inkludert i Fradragsberettiget tap / gevinst — ikke rapporter separat";
  rows += `<tr class="result-note-row"><td colspan="2">${fxAggNote}</td></tr>`;
  addRow(fxAggLabel, fmt(Math.abs(fxAgg)) + " NOK", fxAgg >= 0 ? "val-positive" : "val-negative");

  // Type B: separate FX gain to report under Finans → Valuta
  const fx = data.fx_gain_nok ?? 0;
  if (fx === 0) {
    addRow("Type B — Separat valutagevinst/-tap (NOK)", "0 NOK",
           "");
    rows += `<tr class="result-note-row"><td colspan="2">Ingen separat rapportering nødvendig under Finans → Valuta</td></tr>`;
  } else if (fx > 0) {
    addRow("Type B — Separat valutagevinst (NOK)", fmt(fx) + " NOK", "val-positive");
    rows += `<tr class="result-note-row"><td colspan="2">Rapporter under Finans → Valuta i Skattemeldingen</td></tr>`;
  } else {
    addRow("Type B — Separat valutatap (NOK)", fmt(Math.abs(fx)) + " NOK", "val-negative");
    rows += `<tr class="result-note-row"><td colspan="2">Rapporter under Finans → Valuta i Skattemeldingen</td></tr>`;
  }
  addBlank();

  addSection(`USD-kontobeholdning per 31. desember ${data.year}`);
  addRow("Beholdning (USD)", fmt(data.usd_balance, 2) + " USD");
  addRow("Formue (NOK)", fmt(data.usd_balance_nok) + " NOK");

  table.innerHTML = rows;

  const filename = `Skattemeldingen_${data.year}.xlsx`;
  const now = new Date().toLocaleString("nb-NO", { dateStyle: "short", timeStyle: "short" });
  state.exports.unshift({ filename, timestamp: now });
  renderExportHistory();
}

function renderExportHistory() {
  const el = getEl("export-history");
  if (!el) return;
  if (state.exports.length === 0) {
    el.innerHTML = '<div class="body-md" style="padding:0.5rem 0">No reports downloaded yet.</div>';
    return;
  }
  el.innerHTML = state.exports.slice(0, 5).map((e, i) => `
    <div class="export-item">
      <div>
        <div class="export-name">${e.filename}</div>
        <div class="export-meta">${e.timestamp}</div>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="downloadReport()">⬇</button>
    </div>
    ${i < state.exports.length - 1 ? '<div class="divider" style="margin:0"></div>' : ""}
  `).join("");
}

function companyName(symbol) {
  const names = {
    GOOG: "Alphabet Inc.", GOOGL: "Alphabet Inc.",
    MSFT: "Microsoft Corporation", AAPL: "Apple Inc.",
    AMZN: "Amazon.com Inc.", NVDA: "NVIDIA Corporation",
    META: "Meta Platforms Inc.", CSCO: "Cisco Systems Inc.",
  };
  return names[symbol?.toUpperCase()] || symbol || "—";
}

// ── Download xlsx ──────────────────────────────────────────────────────────
async function downloadReport() {
  if (!state.lastJobId) return;
  try {
    const res = await fetch(`/api/download/${state.lastJobId}`);
    if (!res.ok) { alert("Report not ready. Please wait."); return; }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Skattemeldingen_${state.lastResult?.year || "report"}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (e) {
    alert("Download failed: " + e.message);
  }
}

getEl("btn-download")?.addEventListener("click", downloadReport);
getEl("btn-new-report")?.addEventListener("click", () => showView("reports"));

// ── Init ───────────────────────────────────────────────────────────────────
(async function init() {
  let savedYear = new Date().getFullYear() - 1;
  try {
    const res = await fetch("/api/settings");
    if (res.ok) {
      const s = await res.json();
      state.settings = { ...state.settings, ...s };
      savedYear = s.year || savedYear;
    }
  } catch (e) { console.error("init settings:", e); }

  initYearSelector(savedYear);
  updateYearLabels(savedYear);

  const statusEl = getEl("rates-status-inline");
  if (state.settings.rates_loaded && state.settings.rates_days > 0) {
    if (statusEl) {
      statusEl.textContent = `✓ ${state.settings.rates_days} rate days loaded`;
      statusEl.className = "rates-status ok";
    }
  } else {
    // Auto-fetch rates in the background
    fetchRates(savedYear);
  }

  showView("reports");
})();
