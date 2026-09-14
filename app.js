// FedEx Shipment Cost Tracker — phase 1 app logic.
// No shipment data is persisted anywhere (by design, this phase) — the
// session log is an in-memory array only and clears on refresh. Login
// and dark-mode preference are the only things kept in localStorage,
// as light per-viewer conveniences, not shipment data.

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const STORAGE_USER_KEY = "fedexTracker_currentUser";
const STORAGE_THEME_KEY = "fedexTracker_theme";

let currentParsed = null;
let currentUser = null;
const sessionLog = [];

// A single PDF can now yield more than one shipment record: a true
// multi-piece shipment stays one record (pieces bundled in), but a
// batch of unrelated single-piece labels dropped in as one PDF comes
// back as separate records that get reviewed one at a time here.
let parseQueue = [];
let queueIndex = 0;

// ---------- Dark mode ----------
function applyTheme(theme) {
  if (theme === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
    document.getElementById("darkToggle").textContent = "☀️ Light";
  } else {
    document.documentElement.setAttribute("data-theme", "light");
    document.getElementById("darkToggle").textContent = "🌙 Dark";
  }
}

(function initTheme() {
  const saved = localStorage.getItem(STORAGE_THEME_KEY);
  if (saved) applyTheme(saved);
})();

document.getElementById("darkToggle").addEventListener("click", () => {
  const isDark = document.documentElement.getAttribute("data-theme") === "dark";
  const next = isDark ? "light" : "dark";
  applyTheme(next);
  localStorage.setItem(STORAGE_THEME_KEY, next);
});

// ---------- Login ----------
function findCredential(login) {
  const needle = login.trim().toLowerCase();
  return CREDENTIALS.find((c) => c.login.toLowerCase() === needle) || null;
}

function showApp(cred) {
  currentUser = cred;
  document.getElementById("loginBackdrop").hidden = true;
  document.getElementById("app").hidden = false;
  document.getElementById("whoami").hidden = false;
  document.getElementById("whoamiName").textContent = `${cred.name} (${cred.role})`;
}

function showLogin() {
  currentUser = null;
  document.getElementById("loginBackdrop").hidden = false;
  document.getElementById("app").hidden = true;
  document.getElementById("whoami").hidden = true;
}

function attemptLogin() {
  const input = document.getElementById("loginInput");
  const errorEl = document.getElementById("loginError");
  const cred = findCredential(input.value);
  if (!cred) {
    errorEl.textContent = "Login not recognized. Check the Credentials tab spelling.";
    return;
  }
  errorEl.textContent = "";
  localStorage.setItem(STORAGE_USER_KEY, cred.login);
  showApp(cred);
}

document.getElementById("loginBtn").addEventListener("click", attemptLogin);
document.getElementById("loginInput").addEventListener("keydown", (e) => {
  if (e.key === "Enter") attemptLogin();
});
document.getElementById("logoutBtn").addEventListener("click", () => {
  localStorage.removeItem(STORAGE_USER_KEY);
  showLogin();
});

(function initSession() {
  const savedLogin = localStorage.getItem(STORAGE_USER_KEY);
  if (savedLogin) {
    const cred = findCredential(savedLogin);
    if (cred) {
      showApp(cred);
      return;
    }
  }
  showLogin();
})();

// ---------- Project dropdown ----------
(function populateProjects() {
  const sel = document.getElementById("f_project");
  sel.innerHTML = '<option value="" disabled selected>Choose a project…</option>';
  for (const p of PROJECTS) {
    const opt = document.createElement("option");
    opt.value = p.name;
    opt.textContent = p.archived ? `${p.name} (archived)` : p.name;
    sel.appendChild(opt);
  }
})();

// ---------- Tabs ----------
const tabBtnLog = document.getElementById("tabBtnLog");
const tabBtnExpenses = document.getElementById("tabBtnExpenses");
const tabLog = document.getElementById("tabLog");
const tabExpenses = document.getElementById("tabExpenses");

tabBtnLog.addEventListener("click", () => switchTab("log"));
tabBtnExpenses.addEventListener("click", () => switchTab("expenses"));

function switchTab(name) {
  const onLog = name === "log";
  tabBtnLog.classList.toggle("active", onLog);
  tabBtnExpenses.classList.toggle("active", !onLog);
  tabLog.hidden = !onLog;
  tabExpenses.hidden = onLog;
  if (!onLog) renderExpenses();
}

// ---------- Dropzone ----------
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const dropzoneIdle = document.getElementById("dropzoneIdle");
const dropzoneBusy = document.getElementById("dropzoneBusy");

document.getElementById("browseLink").addEventListener("click", () => fileInput.click());
dropzone.addEventListener("click", (e) => {
  if (e.target.id !== "fileInput") fileInput.click();
});

["dragenter", "dragover"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  })
);
dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) handleFile(file);
});
fileInput.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (file) handleFile(file);
});

async function handleFile(file) {
  if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
    alert("Please drop a PDF file — FedEx Ship Manager labels export as PDF.");
    return;
  }
  dropzoneIdle.hidden = true;
  dropzoneBusy.hidden = false;
  try {
    const records = await parseFedexLabelFile(file);
    if (!records.length) {
      alert("Couldn't find any FedEx label content in that PDF.");
      return;
    }
    parseQueue = records;
    queueIndex = 0;
    loadQueueItem();
  } catch (err) {
    console.error(err);
    alert("Couldn't read that PDF. It may not be a FedEx Ship Manager label, or the file is corrupted.");
  } finally {
    dropzoneIdle.hidden = false;
    dropzoneBusy.hidden = true;
    fileInput.value = "";
  }
}

function loadQueueItem() {
  currentParsed = parseQueue[queueIndex];
  populateConfirmForm(currentParsed);

  const note = document.getElementById("queuePositionNote");
  if (parseQueue.length > 1) {
    note.hidden = false;
    note.textContent = `Shipment ${queueIndex + 1} of ${parseQueue.length} found in this PDF — reviewing one at a time.`;
  } else {
    note.hidden = true;
  }
}

// After confirming or discarding one item, move to the next queued one
// (a batch of unrelated labels in one PDF), or close out if that was
// the last one.
function advanceQueue() {
  queueIndex++;
  if (queueIndex < parseQueue.length) {
    loadQueueItem();
  } else {
    parseQueue = [];
    queueIndex = 0;
    currentParsed = null;
    document.getElementById("confirmCard").hidden = true;
  }
}

// ---------- Confirm form ----------
function populateConfirmForm(parsed) {
  document.getElementById("sourceFileName").textContent = parsed.sourceFile;

  const previewImg = document.getElementById("labelPreviewImg");
  if (parsed.previewDataUrl) {
    previewImg.src = parsed.previewDataUrl;
    previewImg.hidden = false;
  } else {
    previewImg.src = "";
    previewImg.hidden = true;
  }
  document.getElementById("f_trackingNumber").value = parsed.trackingNumber || "";
  document.getElementById("f_shipDate").value = parsed.shipDate || "";
  document.getElementById("f_service").value = parsed.service || "";
  document.getElementById("f_destCountry").value = parsed.destCountry || "";

  // Multi-piece shipment: banner, bundled piece tracking numbers, and a
  // total-plus-breakdown weight instead of just one box's weight.
  const multiBanner = document.getElementById("multiPieceBanner");
  const pieceRow = document.getElementById("pieceTrackingRow");
  if (parsed.isMultiPiece) {
    multiBanner.hidden = false;
    multiBanner.textContent = `Multi-piece shipment — ${parsed.pieceCount} boxes bundled under this master tracking number.`;
    pieceRow.hidden = false;
    document.getElementById("f_pieceTrackingNumbers").value = parsed.pieces
      .map((p) => p.trackingNumber)
      .filter((t) => t && t !== parsed.trackingNumber)
      .join(", ");
    const breakdown = parsed.pieces.map((p) => `${p.weight ?? "?"} ${p.weightUnit ?? ""}`.trim()).join(", ");
    document.getElementById("f_weight").value = parsed.totalWeight
      ? `${parsed.totalWeight} ${parsed.totalWeightUnit} total (${parsed.pieceCount} pieces: ${breakdown})`
      : "";
  } else {
    multiBanner.hidden = true;
    pieceRow.hidden = true;
    document.getElementById("f_pieceTrackingNumbers").value = "";
    document.getElementById("f_weight").value = parsed.weight ? `${parsed.weight} ${parsed.weightUnit}` : "";
  }

  document.getElementById("f_dimensions").value = parsed.dimensions || "";
  document.getElementById("f_reference").value = parsed.reference || "";

  const invPoDeptParts = [];
  if (parsed.invoice) invPoDeptParts.push(`INV ${parsed.invoice}`);
  if (parsed.po) invPoDeptParts.push(`PO ${parsed.po}`);
  if (parsed.dept) invPoDeptParts.push(`DEPT ${parsed.dept}`);
  document.getElementById("f_invPoDept").value = invPoDeptParts.join(" · ");

  document.getElementById("f_senderName").value = parsed.sender.name || "";
  document.getElementById("f_senderAddress").value = parsed.sender.address || "";
  document.getElementById("f_senderPhone").value = parsed.sender.phone || "";
  document.getElementById("f_recipientName").value = parsed.recipient.name || "";
  document.getElementById("f_recipientAddress").value = parsed.recipient.address || "";
  document.getElementById("f_recipientPhone").value = parsed.recipient.phone || "";

  // Service confidence pill
  const confPill = document.getElementById("serviceConfidencePill");
  if (parsed.serviceConfidence === "labeled") {
    confPill.hidden = true;
  } else if (parsed.serviceConfidence) {
    confPill.hidden = false;
    confPill.className = "pill pill-review";
    confPill.textContent = "inferred — confirm";
  } else {
    confPill.hidden = false;
    confPill.className = "pill pill-review";
    confPill.textContent = "not found — enter manually";
  }

  // International / domestic pill
  const intlPill = document.getElementById("intlPill");
  if (parsed.isInternational === null) {
    intlPill.hidden = true;
  } else {
    intlPill.hidden = false;
    intlPill.className = parsed.isInternational ? "pill pill-intl" : "pill pill-domestic";
    intlPill.textContent = parsed.isInternational ? "International" : "Domestic";
  }

  // Warnings banner
  const banner = document.getElementById("warningsBanner");
  const list = document.getElementById("warningsList");
  list.innerHTML = "";
  if (parsed.warnings && parsed.warnings.length) {
    banner.hidden = false;
    for (const w of parsed.warnings) {
      const li = document.createElement("li");
      li.textContent = w;
      list.appendChild(li);
    }
  } else {
    banner.hidden = true;
  }

  // Reset manual fields
  document.getElementById("f_project").value = "";
  document.getElementById("f_price").value = "";
  document.getElementById("f_notes").value = "";

  document.getElementById("confirmCard").hidden = false;
  document.getElementById("confirmCard").scrollIntoView({ behavior: "smooth", block: "start" });
}

document.getElementById("discardBtn").addEventListener("click", () => {
  advanceQueue();
});

document.getElementById("addToLogBtn").addEventListener("click", () => {
  if (!currentParsed) return;
  const project = document.getElementById("f_project").value;
  if (!project) {
    alert("Pick a project before adding this to the log.");
    return;
  }
  const priceRaw = document.getElementById("f_price").value;
  const record = {
    trackingNumber: document.getElementById("f_trackingNumber").value,
    shipDate: document.getElementById("f_shipDate").value,
    service: document.getElementById("f_service").value,
    destCountry: document.getElementById("f_destCountry").value,
    isInternational: currentParsed.isInternational,
    isMultiPiece: currentParsed.isMultiPiece,
    pieceCount: currentParsed.pieceCount,
    pieceTrackingNumbers: document.getElementById("f_pieceTrackingNumbers").value || null,
    totalWeight: currentParsed.totalWeight,
    totalWeightUnit: currentParsed.totalWeightUnit,
    senderName: document.getElementById("f_senderName").value,
    senderAddress: document.getElementById("f_senderAddress").value,
    senderPhone: document.getElementById("f_senderPhone").value,
    recipientName: document.getElementById("f_recipientName").value,
    recipientAddress: document.getElementById("f_recipientAddress").value,
    recipientPhone: document.getElementById("f_recipientPhone").value,
    project,
    price: priceRaw ? parseFloat(priceRaw) : null,
    notes: document.getElementById("f_notes").value,
    parseStatus: currentParsed.parseStatus,
    sourceFile: currentParsed.sourceFile,
    submittedBy: currentUser ? currentUser.name : null,
  };
  sessionLog.push(record);
  renderLog();
  renderExpenses();

  advanceQueue();
});

// ---------- Session log table ----------
function renderLog() {
  const body = document.getElementById("logTableBody");
  const emptyRow = document.getElementById("logEmptyRow");
  body.innerHTML = "";

  if (!sessionLog.length) {
    emptyRow.hidden = false;
    return;
  }
  emptyRow.hidden = true;

  for (const r of sessionLog) {
    const tr = document.createElement("tr");

    const destPill = `<span class="pill ${r.isInternational ? "pill-intl" : "pill-domestic"}">${r.isInternational ? "Intl" : "US"}</span>`;
    const statusPill =
      r.parseStatus === "ok"
        ? '<span class="pill pill-ok">OK</span>'
        : '<span class="pill pill-review">Review</span>';
    const priceStr = r.price != null && !isNaN(r.price) ? `$${r.price.toFixed(2)}` : "—";

    const multiBadge = r.isMultiPiece ? ` <span class="pill pill-muted">×${r.pieceCount}</span>` : "";

    tr.innerHTML = `
      <td>${trackingLink(r.trackingNumber)}${multiBadge}</td>
      <td>${escapeHtml(r.shipDate || "—")}</td>
      <td>${escapeHtml(r.service || "—")}</td>
      <td>${destPill} ${escapeHtml(r.destCountry || "")}</td>
      <td>${escapeHtml(r.project)}</td>
      <td>${escapeHtml(r.submittedBy || "—")}</td>
      <td>${priceStr}</td>
      <td>${statusPill}</td>
    `;
    body.appendChild(tr);
  }
}

// ---------- Expenses by Project tab ----------
(function populateExpenseFilters() {
  const projectSel = document.getElementById("exp_project");
  for (const p of PROJECTS) {
    const opt = document.createElement("option");
    opt.value = p.name;
    opt.textContent = p.archived ? `${p.name} (archived)` : p.name;
    projectSel.appendChild(opt);
  }
  const bySel = document.getElementById("exp_submittedBy");
  for (const c of CREDENTIALS) {
    const opt = document.createElement("option");
    opt.value = c.name;
    opt.textContent = c.name;
    bySel.appendChild(opt);
  }
})();

["exp_project", "exp_submittedBy", "exp_destination", "exp_dateFrom", "exp_dateTo"].forEach((id) => {
  document.getElementById(id).addEventListener("input", renderExpenses);
  document.getElementById(id).addEventListener("change", renderExpenses);
});

document.getElementById("exp_clearFilters").addEventListener("click", () => {
  document.getElementById("exp_project").value = "";
  document.getElementById("exp_submittedBy").value = "";
  document.getElementById("exp_destination").value = "";
  document.getElementById("exp_dateFrom").value = "";
  document.getElementById("exp_dateTo").value = "";
  renderExpenses();
});

function getFilteredLog() {
  const project = document.getElementById("exp_project").value;
  const submittedBy = document.getElementById("exp_submittedBy").value;
  const destination = document.getElementById("exp_destination").value;
  const dateFrom = document.getElementById("exp_dateFrom").value;
  const dateTo = document.getElementById("exp_dateTo").value;

  return sessionLog.filter((r) => {
    if (project && r.project !== project) return false;
    if (submittedBy && r.submittedBy !== submittedBy) return false;
    if (destination === "domestic" && r.isInternational !== false) return false;
    if (destination === "international" && r.isInternational !== true) return false;
    if (dateFrom && r.shipDate && r.shipDate < dateFrom) return false;
    if (dateTo && r.shipDate && r.shipDate > dateTo) return false;
    return true;
  });
}

function renderExpenses() {
  const filtered = getFilteredLog();

  // Summary
  let total = 0;
  let priced = 0;
  for (const r of filtered) {
    if (r.price != null && !isNaN(r.price)) {
      total += r.price;
      priced++;
    }
  }
  document.getElementById("exp_summary").innerHTML =
    `$${total.toFixed(2)} <span style="font-size:13px; font-weight:500; color:var(--sub);">across ${filtered.length} shipment${filtered.length === 1 ? "" : "s"}${priced < filtered.length ? ` (${filtered.length - priced} without a price yet)` : ""}</span>`;

  // Rollup by project
  const rollup = new Map(); // project -> {count, total}
  for (const r of filtered) {
    const key = r.project || "(no project)";
    if (!rollup.has(key)) rollup.set(key, { count: 0, total: 0 });
    const entry = rollup.get(key);
    entry.count++;
    if (r.price != null && !isNaN(r.price)) entry.total += r.price;
  }
  const rollupBody = document.getElementById("exp_rollupBody");
  const rollupEmpty = document.getElementById("exp_rollupEmpty");
  rollupBody.innerHTML = "";
  const rollupRows = [...rollup.entries()].sort((a, b) => b[1].total - a[1].total);
  if (!rollupRows.length) {
    rollupEmpty.hidden = false;
  } else {
    rollupEmpty.hidden = true;
    for (const [project, entry] of rollupRows) {
      const avg = entry.count ? entry.total / entry.count : 0;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(project)}</td>
        <td>${entry.count}</td>
        <td>$${entry.total.toFixed(2)}</td>
        <td>$${avg.toFixed(2)}</td>
      `;
      rollupBody.appendChild(tr);
    }
  }

  // Detail table
  const detailBody = document.getElementById("exp_detailBody");
  const detailEmpty = document.getElementById("exp_detailEmpty");
  detailBody.innerHTML = "";
  if (!filtered.length) {
    detailEmpty.hidden = false;
  } else {
    detailEmpty.hidden = true;
    for (const r of filtered) {
      const destPill = `<span class="pill ${r.isInternational ? "pill-intl" : "pill-domestic"}">${r.isInternational ? "Intl" : "US"}</span>`;
      const priceStr = r.price != null && !isNaN(r.price) ? `$${r.price.toFixed(2)}` : "—";
      const multiBadge = r.isMultiPiece ? ` <span class="pill pill-muted">×${r.pieceCount}</span>` : "";
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${trackingLink(r.trackingNumber)}${multiBadge}</td>
        <td>${escapeHtml(r.shipDate || "—")}</td>
        <td>${escapeHtml(r.project)}</td>
        <td>${escapeHtml(r.submittedBy || "—")}</td>
        <td>${destPill}</td>
        <td>${priceStr}</td>
      `;
      detailBody.appendChild(tr);
    }
  }
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

// Links straight to FedEx's own tracking page. Verified the "trknbr"
// query param alone is enough for FedEx to attempt a real lookup (it
// also accepts an extra "trkqual" param FedEx generates internally,
// but that's not something we have or need). If a shipment's tracking
// number ages out of FedEx's own retention window, this will land on
// FedEx's "we can't find that tracking number" page rather than erroring.
function trackingLink(trackingNumber) {
  if (!trackingNumber) return "—";
  const url = `https://www.fedex.com/fedextrack/?trknbr=${encodeURIComponent(trackingNumber)}`;
  return `<a href="${url}" target="_blank" rel="noopener">${escapeHtml(trackingNumber)}</a>`;
}

renderLog();
