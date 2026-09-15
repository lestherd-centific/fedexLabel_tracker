// FedEx Shipment Cost Tracker — app logic.
// Shipment history is backed by Excel (via the two Power Automate flows
// below), not by this browser: `shipmentHistory` is loaded fresh from
// the read flow on startup, refreshed again right before every add (for
// an up-to-date duplicate check), and can be reloaded on demand with the
// Refresh button. Login and dark-mode preference are the only things
// kept in localStorage, as light per-viewer conveniences.

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const STORAGE_USER_KEY = "fedexTracker_currentUser";
const STORAGE_THEME_KEY = "fedexTracker_theme";

// Paste the Power Automate flow's "When an HTTP request is received"
// trigger URL here once that flow is built (see the design spec, §18).
// Left blank, every shipment just shows "Not connected" in the Sync
// column instead of trying to reach anything.
const POWER_AUTOMATE_URL =
  "https://default9b415834803a4da0afdcfe6b1d52d6.49.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/09/workflows/0c11aa6f378d43329be14b5836aee79a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=l566c9iR8sVBEwGVbdWO1erfuf0U0uEeWabXfm6pjzw";

// Paste the second (read-only) flow's "When an HTTP request is
// received" URL here once it's built (see the design spec, §20). Left
// blank, history just starts empty each load and duplicate-checking
// falls back to whatever's already in memory this session.
const POWER_AUTOMATE_READ_URL =
  "https://default9b415834803a4da0afdcfe6b1d52d6.49.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/10/workflows/a8681121066d4f16a53f81e25814a494/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hvEQwuxmJXIsUIxOyAUlrnZKScGDwRR4baOGEX0eZGA";

// Paste the third flow's URL here once it's built (see the design spec,
// §21) -- a POST that permanently deletes a row from Shipments by its
// `id`. Left blank, delete falls back to "remove from this view only"
// for every row, even ones already synced to Excel.
const POWER_AUTOMATE_DELETE_URL =
  "https://default9b415834803a4da0afdcfe6b1d52d6.49.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/14/workflows/ed7dcdb7f78a4ce2be55dd9624698b5c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=R_ub5Kxx8OGXEvI_eKMNsq4Ntub1iIV6fTv9M9keDpk";

const LOG_PAGE_SIZE = 10;
let logPage = 0; // 0-indexed

let currentParsed = null;
let currentUser = null;
const shipmentHistory = []; // mutated in place (push/splice/length=0), never reassigned

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

// ---------- UI helpers: toast + confirm modal (replace alert()/confirm()) ----------
// A themed, non-blocking notice in the bottom-right corner instead of the
// browser's native alert() banner. Click it, or wait, and it's gone.
function showToast(message, { type = "info", duration = 5000 } = {}) {
  const container = document.getElementById("toastContainer");
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  toast.title = "Click to dismiss";
  toast.addEventListener("click", () => toast.remove());
  container.appendChild(toast);
  if (duration) setTimeout(() => toast.remove(), duration);
}

// A themed modal instead of the browser's native confirm() dialog.
// Returns a Promise<boolean> -- true if Confirm was clicked, false for
// Cancel, clicking outside the box, or pressing Escape.
function showConfirm({ title, message, confirmLabel = "Confirm", cancelLabel = "Cancel", danger = false }) {
  return new Promise((resolve) => {
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop confirm-modal";
    backdrop.innerHTML = `
      <div class="modal">
        <h2>${escapeHtml(title)}</h2>
        <p>${escapeHtml(message)}</p>
        <div class="modal-actions">
          <button class="btn" data-action="cancel" type="button">${escapeHtml(cancelLabel)}</button>
          <button class="${danger ? "btn-danger" : "btn-primary"}" data-action="confirm" type="button">${escapeHtml(confirmLabel)}</button>
        </div>
      </div>
    `;
    function close(result) {
      backdrop.remove();
      document.removeEventListener("keydown", onKeydown);
      resolve(result);
    }
    function onKeydown(e) {
      if (e.key === "Escape") close(false);
    }
    backdrop.addEventListener("click", (e) => {
      if (e.target === backdrop) close(false); // clicked outside the modal box
    });
    backdrop.querySelector('[data-action="cancel"]').addEventListener("click", () => close(false));
    backdrop.querySelector('[data-action="confirm"]').addEventListener("click", () => close(true));
    document.addEventListener("keydown", onKeydown);
    document.body.appendChild(backdrop);
  });
}

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
    showToast("Please drop a PDF file — FedEx Ship Manager labels export as PDF.", { type: "error" });
    return;
  }
  dropzoneIdle.hidden = true;
  dropzoneBusy.hidden = false;
  try {
    const records = await parseFedexLabelFile(file);
    if (!records.length) {
      showToast("Couldn't find any FedEx label content in that PDF.", { type: "error" });
      return;
    }
    parseQueue = records;
    queueIndex = 0;
    loadQueueItem();
  } catch (err) {
    console.error(err);
    showToast("Couldn't read that PDF. It may not be a FedEx Ship Manager label, or the file is corrupted.", { type: "error" });
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

document.getElementById("addToLogBtn").addEventListener("click", async () => {
  if (!currentParsed) return;
  const project = document.getElementById("f_project").value;
  if (!project) {
    showToast("Pick a project before adding this to the log.", { type: "error" });
    return;
  }
  const trackingNumber = document.getElementById("f_trackingNumber").value.trim();

  const addBtn = document.getElementById("addToLogBtn");
  const originalLabel = addBtn.textContent;
  addBtn.disabled = true;
  addBtn.textContent = "Checking for duplicates…";
  try {
    // Refresh from Excel first (if the read flow's connected) so the
    // duplicate check below is against what's actually in Shipments
    // right now, not a stale local cache that could miss something a
    // teammate added minutes ago on a different machine. Falls back to
    // whatever's already loaded if the refresh itself fails, rather
    // than blocking the add entirely on a network hiccup.
    if (POWER_AUTOMATE_READ_URL) await refreshHistoryFromExcel({ silent: true });

    if (trackingNumber && shipmentHistory.some((r) => r.trackingNumber === trackingNumber)) {
      showToast(`Tracking # ${trackingNumber} is already in the shipment history — not adding it again.`, { type: "error" });
      return;
    }

    const priceRaw = document.getElementById("f_price").value;
    const record = {
      trackingNumber,
      shipDate: document.getElementById("f_shipDate").value,
      service: document.getElementById("f_service").value,
      destCountry: document.getElementById("f_destCountry").value,
      isInternational: currentParsed.isInternational,
      isMultiPiece: currentParsed.isMultiPiece,
      pieceCount: currentParsed.pieceCount,
      pieceTrackingNumbers: document.getElementById("f_pieceTrackingNumbers").value || null,
      totalWeight: currentParsed.totalWeight,
      totalWeightUnit: currentParsed.totalWeightUnit,
      reference: document.getElementById("f_reference").value || null,
      invoicePoDept: document.getElementById("f_invPoDept").value || null,
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
    record.syncStatus = POWER_AUTOMATE_URL ? "syncing" : "unconfigured";

    shipmentHistory.push(record);
    logPage = 0; // jump to the newest page so the just-added row is visible
    renderLog();
    renderExpenses();

    // Fires in the background -- doesn't block moving on to the next
    // queued shipment. syncRecordToExcel() updates record.syncStatus and
    // re-renders the log itself once it knows the result.
    if (POWER_AUTOMATE_URL) syncRecordToExcel(record);

    advanceQueue();
  } finally {
    addBtn.disabled = false;
    addBtn.textContent = originalLabel;
  }
});

// ---------- Sync to Excel (Power Automate) ----------
// Builds exactly the JSON body the flow's trigger expects -- the 25
// Shipments columns minus `id`/`ts`, which the flow generates itself
// (guid()/utcNow()) rather than trusting the browser's clock or a
// client-generated id.
function shipmentPayload(record) {
  return {
    trackingNumber: record.trackingNumber,
    shipDate: record.shipDate,
    service: record.service,
    destCountry: record.destCountry,
    isInternational: record.isInternational,
    isMultiPiece: record.isMultiPiece,
    pieceCount: record.pieceCount,
    pieceTrackingNumbers: record.pieceTrackingNumbers,
    totalWeight: record.totalWeight,
    totalWeightUnit: record.totalWeightUnit,
    reference: record.reference,
    invoicePoDept: record.invoicePoDept,
    senderName: record.senderName,
    senderAddress: record.senderAddress,
    senderPhone: record.senderPhone,
    recipientName: record.recipientName,
    recipientAddress: record.recipientAddress,
    recipientPhone: record.recipientPhone,
    project: record.project,
    price: record.price,
    notes: record.notes,
    parseStatus: record.parseStatus,
    sourceFile: record.sourceFile,
    submittedBy: record.submittedBy,
  };
}

async function syncRecordToExcel(record) {
  if (!POWER_AUTOMATE_URL) {
    record.syncStatus = "unconfigured";
    renderLog();
    return;
  }
  record.syncStatus = "syncing";
  renderLog();
  try {
    const res = await fetch(POWER_AUTOMATE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(shipmentPayload(record)),
    });
    if (!res.ok) throw new Error(`Flow responded ${res.status}`);
    // The write flow's Response body now echoes back the row's
    // Excel-generated `id` (see design spec §21) -- capturing it here
    // means a row added THIS session can be permanently deleted right
    // away, not just after the next history refresh re-pulls it.
    const data = await res.json().catch(() => ({}));
    if (data && data.id) record.id = data.id;
    record.syncStatus = "synced";
  } catch (err) {
    console.error("Sync to Excel failed:", err);
    record.syncStatus = "failed";
  }
  renderLog();
}

// Permanently deletes one row from Shipments by its Excel-generated
// `id`. Throws on failure -- callers decide how to handle that (leave
// the row in place rather than silently removing it locally when the
// real delete didn't actually happen).
async function deleteRecordFromExcel(record) {
  const res = await fetch(POWER_AUTOMATE_DELETE_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id: record.id }),
  });
  if (!res.ok) throw new Error(`Delete flow responded ${res.status}`);
}

// ---------- History (read from Excel) ----------
// Excel Online returns each row's fields matching the Shipments header
// row, but cell formatting can make booleans/numbers come back as
// strings (e.g. "TRUE" instead of true) -- normalize defensively so the
// rest of the app can rely on real types regardless of how a cell
// happens to be formatted in the workbook.
function toBool(v) {
  return v === true || v === "true" || v === "TRUE" || v === 1;
}
function toNumOrNull(v) {
  if (v === "" || v === undefined || v === null) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function normalizeHistoryRow(row) {
  return {
    id: row.id,
    ts: row.ts,
    trackingNumber: row.trackingNumber || "",
    shipDate: row.shipDate || "",
    service: row.service || "",
    destCountry: row.destCountry || "",
    isInternational: toBool(row.isInternational),
    isMultiPiece: toBool(row.isMultiPiece),
    pieceCount: toNumOrNull(row.pieceCount) ?? 1,
    pieceTrackingNumbers: row.pieceTrackingNumbers || null,
    totalWeight: toNumOrNull(row.totalWeight),
    totalWeightUnit: row.totalWeightUnit || null,
    reference: row.reference || null,
    invoicePoDept: row.invoicePoDept || null,
    senderName: row.senderName || "",
    senderAddress: row.senderAddress || "",
    senderPhone: row.senderPhone || "",
    recipientName: row.recipientName || "",
    recipientAddress: row.recipientAddress || "",
    recipientPhone: row.recipientPhone || "",
    project: row.project || "",
    price: toNumOrNull(row.price),
    notes: row.notes || "",
    parseStatus: row.parseStatus || "ok",
    sourceFile: row.sourceFile || "",
    submittedBy: row.submittedBy || null,
    syncStatus: "synced", // it came from Excel, so it's already there by definition
  };
}

// Calls the read flow and returns a normalized array, or throws. Callers
// decide how to handle a failure (fall back to cache, show a message).
async function fetchShipmentHistory() {
  const res = await fetch(POWER_AUTOMATE_READ_URL);
  if (!res.ok) throw new Error(`History flow responded ${res.status}`);
  const rows = await res.json();
  if (!Array.isArray(rows)) throw new Error("History flow didn't return an array");
  return rows.map(normalizeHistoryRow);
}

const historyStatusNote = document.getElementById("historyStatusNote");

// Replaces shipmentHistory's contents in place (never reassigns the
// const) with a fresh pull from Excel. `silent` skips the "Loading…"
// message for the pre-add duplicate-check call, so it doesn't flicker
// the note on every click.
async function refreshHistoryFromExcel({ silent = false } = {}) {
  if (!POWER_AUTOMATE_READ_URL) {
    if (historyStatusNote) {
      historyStatusNote.textContent =
        "History isn't connected to Excel yet — add POWER_AUTOMATE_READ_URL in app.js (see design spec §20).";
    }
    return false;
  }
  if (!silent && historyStatusNote) historyStatusNote.textContent = "Loading history from Excel…";
  try {
    const fresh = await fetchShipmentHistory();
    shipmentHistory.length = 0;
    shipmentHistory.push(...fresh);
    if (historyStatusNote) {
      historyStatusNote.textContent = `Synced with Excel as of ${new Date().toLocaleTimeString()}.`;
    }
    return true;
  } catch (err) {
    console.error("Loading history from Excel failed:", err);
    if (historyStatusNote) {
      historyStatusNote.textContent = "Couldn't load history from Excel just now — showing what's already loaded.";
    }
    return false;
  }
}

document.getElementById("historyRefreshBtn").addEventListener("click", async () => {
  await refreshHistoryFromExcel();
  logPage = 0;
  renderLog();
  renderExpenses();
});

// ---------- Session log table ----------
function renderLog() {
  const body = document.getElementById("logTableBody");
  const emptyRow = document.getElementById("logEmptyRow");
  const pagination = document.getElementById("logPagination");
  body.innerHTML = "";

  if (!shipmentHistory.length) {
    emptyRow.hidden = false;
    pagination.hidden = true;
    return;
  }
  emptyRow.hidden = true;

  // Most recent first. `ts` is an ISO timestamp (from Excel's utcNow(),
  // or unset for a just-added row this session), so a plain string
  // compare sorts correctly; unset ts sorts last within "just added".
  const sorted = [...shipmentHistory].sort((a, b) => (b.ts || "").localeCompare(a.ts || ""));

  const pageCount = Math.max(1, Math.ceil(sorted.length / LOG_PAGE_SIZE));
  logPage = Math.min(logPage, pageCount - 1);
  const pageRows = sorted.slice(logPage * LOG_PAGE_SIZE, logPage * LOG_PAGE_SIZE + LOG_PAGE_SIZE);

  for (const r of pageRows) {
    const tr = document.createElement("tr");

    const destPill = `<span class="pill ${r.isInternational ? "pill-intl" : "pill-domestic"}">${r.isInternational ? "Intl" : "US"}</span>`;
    const statusPill =
      r.parseStatus === "ok"
        ? '<span class="pill pill-ok">OK</span>'
        : '<span class="pill pill-review">Review</span>';
    const priceStr = r.price != null && !isNaN(r.price) ? `$${r.price.toFixed(2)}` : "—";

    const multiBadge = r.isMultiPiece ? ` <span class="pill pill-muted">×${r.pieceCount}</span>` : "";
    // Index within the FULL (unsorted, unpaginated) array -- delete/retry
    // splice/act on shipmentHistory itself, not the page-local slice.
    const idx = shipmentHistory.indexOf(r);
    const syncCell = syncStatusCell(r.syncStatus, idx);

    tr.innerHTML = `
      <td>${trackingLink(r.trackingNumber)}${multiBadge}</td>
      <td>${escapeHtml(r.shipDate || "—")}</td>
      <td>${escapeHtml(r.service || "—")}</td>
      <td>${destPill} ${escapeHtml(r.destCountry || "")}</td>
      <td>${escapeHtml(r.project)}</td>
      <td>${escapeHtml(r.submittedBy || "—")}</td>
      <td>${priceStr}</td>
      <td>${statusPill}</td>
      <td>${syncCell}</td>
      <td><button class="icon-btn log-delete-btn" type="button" data-idx="${idx}" title="Delete this shipment from the log">🗑</button></td>
    `;
    body.appendChild(tr);
  }

  pagination.hidden = pageCount <= 1;
  document.getElementById("logPageInfo").textContent = `Page ${logPage + 1} of ${pageCount}`;
  document.getElementById("logPrevPage").disabled = logPage === 0;
  document.getElementById("logNextPage").disabled = logPage >= pageCount - 1;
}

document.getElementById("logPrevPage").addEventListener("click", () => {
  logPage = Math.max(0, logPage - 1);
  renderLog();
});
document.getElementById("logNextPage").addEventListener("click", () => {
  logPage++;
  renderLog();
});

// The Sync column: what state this row's write to the Shipments tab is
// in. "failed" also gets a retry link right in the pill, since a live
// per-shipment sync with no way to nudge a failed write again would
// leave it stuck there with no recourse but re-entering it by hand.
function syncStatusCell(status, idx) {
  switch (status) {
    case "synced":
      return '<span class="pill pill-ok">Synced</span>';
    case "syncing":
      return '<span class="pill pill-muted">Syncing…</span>';
    case "failed":
      return `<span class="pill pill-error">Failed</span> <button class="linklike-pill" type="button" data-retry-idx="${idx}">retry</button>`;
    case "unconfigured":
      return '<span class="pill pill-muted" title="No Power Automate flow URL configured yet">Not connected</span>';
    default:
      return "—";
  }
}

// Delegated once on the (persistent) table body, not re-attached on every
// renderLog() call -- otherwise repeated renders would stack duplicate
// listeners on the same node.
document.getElementById("logTableBody").addEventListener("click", (e) => {
  const deleteBtn = e.target.closest(".log-delete-btn");
  if (deleteBtn) {
    requestDeleteLogEntry(parseInt(deleteBtn.dataset.idx, 10));
    return;
  }
  const retryBtn = e.target.closest("[data-retry-idx]");
  if (retryBtn) {
    const r = shipmentHistory[parseInt(retryBtn.dataset.retryIdx, 10)];
    if (r) syncRecordToExcel(r);
  }
});

// A row that's already synced to Excel (has a real `id`) AND has the
// delete flow configured gets permanently deleted from Shipments, for
// everyone, no undo. A row that hasn't synced yet (still syncing/failed/
// unconfigured, no `id`) was never in Excel to begin with, so deleting
// it here is just removing it from this list -- nothing to call out to.
async function requestDeleteLogEntry(idx) {
  const r = shipmentHistory[idx];
  if (!r) return;
  const label = `${r.trackingNumber || "(no tracking #)"} — ${r.project || "no project"}`;
  const canDeleteFromExcel = Boolean(r.id && POWER_AUTOMATE_DELETE_URL);

  const confirmed = await showConfirm({
    title: canDeleteFromExcel ? "Permanently delete this shipment?" : "Remove this shipment?",
    message: canDeleteFromExcel
      ? `This permanently deletes the row from the shared Shipments tab in Excel, for everyone. There's no undo.\n\n${label}`
      : `This row hasn't finished syncing to Excel yet, so it'll just be removed from this list.\n\n${label}`,
    confirmLabel: canDeleteFromExcel ? "Delete permanently" : "Remove",
    danger: true,
  });
  if (!confirmed) return;

  if (!canDeleteFromExcel) {
    shipmentHistory.splice(idx, 1);
    renderLog();
    renderExpenses();
    return;
  }

  try {
    await deleteRecordFromExcel(r);
    shipmentHistory.splice(idx, 1);
    renderLog();
    renderExpenses();
    showToast("Shipment permanently deleted from Excel.", { type: "success" });
  } catch (err) {
    console.error("Delete from Excel failed:", err);
    showToast("Couldn't delete that row from Excel — nothing was removed. Try again.", { type: "error" });
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

  return shipmentHistory.filter((r) => {
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

// Initial load: pull real history from Excel (if the read flow's
// connected) before the first render, so the table doesn't flash empty
// and then repopulate a moment later.
(async function initHistory() {
  await refreshHistoryFromExcel();
  renderLog();
  renderExpenses();
})();
