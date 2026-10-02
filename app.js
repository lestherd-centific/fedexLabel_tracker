// FedEx Shipment Cost Tracker — app logic.
// Shipment history is backed by Excel (via the two Power Automate flows
// below), not by this browser: `shipmentHistory` is loaded fresh from
// the read flow on startup, refreshed again right before every add (for
// an up-to-date duplicate check), and can be reloaded on demand with the
// Refresh button. Logins (Credentials tab) and the project list
// (Projects tab) are also read live from Excel on every load -- see
// EXCEL_LOOKUPS_URL below; data.js is no longer used. The remembered
// login and dark-mode preference are the only things kept in
// localStorage, as light per-viewer conveniences.

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

// Read-only lookups: the Hardware Tracker's own "Load" flow (a GET on
// the same shared workbook). Its response includes `credentials`
// (name/login/role) and `projects` (name/archived) alongside other
// tables this app ignores. Reused on purpose so there's no separate
// flow to maintain -- but that means changes to the Hardware Tracker's
// Load flow can affect login and projects here.
const EXCEL_LOOKUPS_URL =
  "https://default9b415834803a4da0afdcfe6b1d52d6.49.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d685d4ce87b14c37897baef1d30041eb/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tqZieZYCMgpC5imEh0w2ObioFk9tXOiTHPB-3nYmGk0";

// The Hardware Tracker's pseudo-project for unassigned stock -- not a
// real project anyone ships against, so it's left out of the dropdowns.
const HW_MASTER_PROJECT = "Master Inventory";

// Filled from Excel by loadLookups(); empty until that returns.
let CREDENTIALS = [];
let PROJECTS = [];

const LOG_PAGE_SIZE = 10;
const rowRecords = new WeakMap(); // history <tr> -> its shipment record (for the overview)
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

// Fades the startup splash out, then removes it from view.
function hideLoading() {
  const splash = document.getElementById("loadingBackdrop");
  if (splash.hidden || splash.classList.contains("splash-out")) return;
  splash.classList.add("splash-out");
  setTimeout(() => {
    splash.hidden = true;
  }, 300);
}

function showApp(cred) {
  currentUser = cred;
  hideLoading();
  document.getElementById("loginBackdrop").hidden = true;
  document.getElementById("app").hidden = false;
  document.getElementById("whoami").hidden = false;
  document.getElementById("whoamiName").textContent = `${cred.name} (${cred.role})`;
}

function showLogin() {
  currentUser = null;
  hideLoading();
  document.getElementById("loginBackdrop").hidden = false;
  document.getElementById("app").hidden = true;
  document.getElementById("whoami").hidden = true;
}

// ---------- Lookups (Credentials + Projects, live from Excel) ----------
function applyLookups(data) {
  CREDENTIALS = (data.credentials || [])
    .filter((r) => r && r.login)
    .map((r) => ({ name: toStr(r.name).trim(), login: toStr(r.login).trim(), role: toStr(r.role).trim() }));
  PROJECTS = (data.projects || [])
    .map((r) => ({ name: toStr(r && r.name).trim(), archived: toBool(r && r.archived) }))
    .filter((p) => p.name && p.name !== HW_MASTER_PROJECT);
  populateProjectDropdowns();
}

// Throws on failure. Never falls back to a hardcoded list -- Excel is
// the only source of truth for who can log in.
async function loadLookups() {
  const res = await fetch(EXCEL_LOOKUPS_URL);
  if (!res.ok) throw new Error(`Lookups flow responded ${res.status}`);
  const data = await res.json();
  if (!data || !Array.isArray(data.credentials)) throw new Error("Lookups flow returned no credentials");
  applyLookups(data);
}

let lookupsLoaded = false;
// False until the first history load from Excel finishes (either way),
// so the empty table says "Loading…" rather than "No shipments yet."
let historyLoadedOnce = false;

function setLoginBusy(busy, label) {
  const btn = document.getElementById("loginBtn");
  btn.disabled = busy;
  btn.textContent = label || (busy ? "Loading logins from Excel…" : "Enter");
}

async function attemptLogin() {
  const input = document.getElementById("loginInput");
  const errorEl = document.getElementById("loginError");
  let cred = lookupsLoaded ? findCredential(input.value) : null;

  // Not found (or the first load failed): re-pull from Excel once, so a
  // login added to the Credentials tab a minute ago works without
  // reloading the page.
  if (!cred) {
    setLoginBusy(true, "Checking Excel…");
    try {
      await loadLookups();
      lookupsLoaded = true;
      cred = findCredential(input.value);
    } catch (err) {
      console.error("Loading logins from Excel failed:", err);
      errorEl.textContent = "Couldn't reach Excel to check logins. Try again in a moment.";
      setLoginBusy(false);
      return;
    }
    setLoginBusy(false);
  }

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
  if (e.key === "Enter" && !document.getElementById("loginBtn").disabled) attemptLogin();
});
document.getElementById("logoutBtn").addEventListener("click", () => {
  localStorage.removeItem(STORAGE_USER_KEY);
  showLogin();
});

// On load: show the login modal in a "loading" state, pull Credentials
// and Projects from Excel, then either restore the remembered login (if
// it's still in the Credentials tab) or leave the modal up.
// The "Loading…" screen (in index.html) is up from the first paint; it
// stays until this check finishes, then gives way to the app (remembered,
// still-valid login) or the Sign in box (no login, removed login, or
// Excel unreachable -- with the error shown there).
(async function initSession() {
  const savedLogin = localStorage.getItem(STORAGE_USER_KEY);
  if (!savedLogin) document.getElementById("loadingMsg").textContent = "Loading… getting things ready.";
  try {
    await loadLookups();
    lookupsLoaded = true;
  } catch (err) {
    console.error("Loading logins from Excel failed:", err);
    document.getElementById("loginError").textContent =
      "Couldn't load logins from Excel. Enter your login to try again.";
  }
  setLoginBusy(false);

  const cred = lookupsLoaded && savedLogin ? findCredential(savedLogin) : null;
  if (cred) {
    showApp(cred);
    return;
  }
  if (lookupsLoaded && savedLogin) localStorage.removeItem(STORAGE_USER_KEY); // removed from Credentials
  showLogin();
  document.getElementById("loginInput").focus();
})();

// ---------- Project dropdowns (log form + Expenses filters) ----------
function populateProjectDropdowns() {
  const sel = document.getElementById("f_project");
  const keep = sel.value;
  sel.innerHTML = '<option value="" disabled selected>Choose a project…</option>';
  for (const p of PROJECTS) {
    const opt = document.createElement("option");
    opt.value = p.name;
    opt.textContent = p.archived ? `${p.name} (archived)` : p.name;
    sel.appendChild(opt);
  }
  if (keep && PROJECTS.some((p) => p.name === keep)) sel.value = keep;

  const projectSel = document.getElementById("exp_project");
  const keepExp = projectSel.value;
  projectSel.innerHTML = '<option value="">All projects</option>';
  for (const p of PROJECTS) {
    const opt = document.createElement("option");
    opt.value = p.name;
    opt.textContent = p.archived ? `${p.name} (archived)` : p.name;
    projectSel.appendChild(opt);
  }
  projectSel.value = PROJECTS.some((p) => p.name === keepExp) ? keepExp : "";

  const bySel = document.getElementById("exp_submittedBy");
  const keepBy = bySel.value;
  bySel.innerHTML = '<option value="">Everyone</option>';
  for (const c of CREDENTIALS) {
    const opt = document.createElement("option");
    opt.value = c.name;
    opt.textContent = c.name;
    bySel.appendChild(opt);
  }
  bySel.value = CREDENTIALS.some((c) => c.name === keepBy) ? keepBy : "";
}

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
    showToast("Please drop a PDF label file.", { type: "error" });
    return;
  }
  dropzoneIdle.hidden = true;
  dropzoneBusy.hidden = false;
  let records = null;
  try {
    records = await parseFedexLabelFile(file);
  } catch (err) {
    console.error(err);
    showToast("Couldn't read that PDF — the file may be corrupted.", { type: "error" });
    return;
  } finally {
    dropzoneIdle.hidden = false;
    dropzoneBusy.hidden = true;
    fileInput.value = "";
  }

  if (records.length) {
    parseQueue = records;
    queueIndex = 0;
    loadQueueItem();
    return;
  }

  // Not a FedEx label (e.g. another carrier): offer manual entry with the
  // label still shown on the left, instead of just rejecting the file.
  const manual = await showConfirm({
    title: "Not a FedEx label",
    message:
      `The app can't read "${file.name}" automatically — it may be from another carrier.\n\n` +
      "Enter the details by hand? The label stays visible next to the form.",
    confirmLabel: "Enter manually",
  });
  if (!manual) return;
  try {
    parseQueue = [await buildManualRecord(file)];
    queueIndex = 0;
    loadQueueItem();
  } catch (err) {
    console.error(err);
    showToast("Couldn't open that PDF for manual entry.", { type: "error" });
  }
}

function loadQueueItem() {
  closeLabelLightbox();
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
  document.getElementById("confirmSubIntro").textContent = parsed.isManual ? "Manual entry for" : "Auto-parsed from";
  document.getElementById("confirmSubRest").textContent = parsed.isManual
    ? " — fill in the details from the label, then add the project and price."
    : " — check anything flagged below, then add the project and price.";

  const previewImg = document.getElementById("labelPreviewImg");
  const previewBtn = document.getElementById("labelPreviewBtn");
  if (parsed.previewDataUrl) {
    previewImg.src = parsed.previewDataUrl;
    previewBtn.hidden = false;
  } else {
    previewImg.src = "";
    previewBtn.hidden = true;
  }
  document.getElementById("f_trackingNumber").value = parsed.trackingNumber || "";
  document.getElementById("f_shipDate").value = parsed.shipDate || "";
  document.getElementById("f_service").value = parsed.service || "";
  document.getElementById("f_destCountry").value = parsed.destCountry || "";
  document.getElementById("f_carrier").value = parsed.carrier || "";
  document.getElementById("f_shipScope").value = parsed.isManual ? "" : scopeFromBool(parsed.isInternational);

  // Manual entry (a label the parser can't read): banner, required
  // markers, and the label-derived fields that are normally read-only
  // become editable.
  const isManual = Boolean(parsed.isManual);
  document.getElementById("manualBanner").hidden = !isManual;
  document.querySelectorAll("#confirmCard .manual-only").forEach((el) => (el.hidden = !isManual));
  document.getElementById("f_destCountry").readOnly = !isManual;
  document.querySelectorAll("#confirmCard .field-missing").forEach((el) => el.classList.remove("field-missing"));

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
  if (parsed.isManual || parsed.serviceConfidence === "labeled") {
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

// ---------- Label preview lightbox ----------
// Click the small preview -> a large overlay. The label is re-rendered
// from the PDF page at a size that fits the window (sharp, not a
// stretched thumbnail); falls back to the thumbnail if that fails.
// "Zoom in" shows it at 2x with scrolling for reading fine print.
// Close with the button, Esc, or a click on the dark background.
const lightbox = document.getElementById("labelLightbox");
const lightboxImg = document.getElementById("labelLightboxImg");
const lightboxStage = document.getElementById("labelLightboxStage");
const lightboxZoomBtn = document.getElementById("labelLightboxZoom");
const lightboxLoading = document.getElementById("labelLightboxLoading");
const largePreviewCache = new WeakMap(); // parsed record -> data URL
let lightboxReturnFocus = null;

function setLightboxZoom(zoomed) {
  lightboxStage.classList.toggle("zoomed", zoomed);
  lightboxZoomBtn.textContent = zoomed ? "Fit to window" : "Zoom in";
}

async function openLabelLightbox() {
  const parsed = currentParsed;
  if (!parsed || !parsed.previewDataUrl) return;
  lightboxReturnFocus = document.activeElement;
  document.getElementById("labelLightboxTitle").textContent =
    parsed.trackingNumber ? `Tracking # ${parsed.trackingNumber}` : parsed.sourceFile || "Label preview";
  setLightboxZoom(false);
  lightboxImg.src = largePreviewCache.get(parsed) || parsed.previewDataUrl;
  lightbox.hidden = false;
  document.body.classList.add("lightbox-open");
  document.getElementById("labelLightboxClose").focus();

  if (!largePreviewCache.has(parsed) && parsed.previewPage) {
    lightboxLoading.hidden = false;
    try {
      // Wide enough for 2x zoom to stay crisp; capped so huge monitors
      // don't produce an enormous image.
      const width = Math.min(Math.max(window.innerWidth * 0.9, 900), 1600);
      const url = await renderLabelPreview(parsed.previewPage, width, { trim: Boolean(parsed.previewTrim) });
      largePreviewCache.set(parsed, url);
      if (!lightbox.hidden && currentParsed === parsed) lightboxImg.src = url;
    } catch (err) {
      console.error("Large label render failed, keeping the thumbnail:", err);
    } finally {
      lightboxLoading.hidden = true;
    }
  }
}

function closeLabelLightbox() {
  if (lightbox.hidden) return;
  lightbox.hidden = true;
  document.body.classList.remove("lightbox-open");
  if (lightboxReturnFocus && lightboxReturnFocus.focus) lightboxReturnFocus.focus();
}

document.getElementById("labelPreviewBtn").addEventListener("click", openLabelLightbox);
document.getElementById("labelLightboxClose").addEventListener("click", closeLabelLightbox);
lightboxZoomBtn.addEventListener("click", () => setLightboxZoom(!lightboxStage.classList.contains("zoomed")));
lightboxImg.addEventListener("click", () => setLightboxZoom(!lightboxStage.classList.contains("zoomed")));
lightbox.addEventListener("click", (e) => {
  if (e.target === lightbox || e.target === lightboxStage) closeLabelLightbox();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !lightbox.hidden) closeLabelLightbox();
});

document.getElementById("discardBtn").addEventListener("click", () => {
  advanceQueue();
});

// Required on every add: Project. On a manual entry also: carrier,
// tracking #, ship date, price and shipment type (typed by hand, so
// they're easy to miss). Missing fields get outlined and listed.
function missingRequiredFields() {
  const required = [["f_project", "Project"]];
  if (currentParsed && currentParsed.isManual) {
    required.push(
      ["f_carrier", "Carrier"],
      ["f_trackingNumber", "Tracking number"],
      ["f_shipDate", "Ship date"],
      ["f_shipScope", "Shipment type"],
      ["f_price", "Price"]
    );
  }
  const missing = [];
  for (const [id, label] of required) {
    const el = document.getElementById(id);
    const empty = String(el.value || "").trim() === "";
    el.classList.toggle("field-missing", empty);
    if (empty) missing.push({ el, label });
  }
  return missing;
}

// Clear a field's red "missing" outline as soon as it's filled in.
["input", "change"].forEach((evt) =>
  document.getElementById("confirmCard").addEventListener(evt, (e) => {
    if (e.target.classList && String(e.target.value || "").trim()) e.target.classList.remove("field-missing");
  })
);

document.getElementById("addToLogBtn").addEventListener("click", async () => {
  if (!currentParsed) return;
  const missing = missingRequiredFields();
  if (missing.length) {
    showToast(`Fill in before adding: ${missing.map((m) => m.label).join(", ")}.`, { type: "error" });
    missing[0].el.focus();
    return;
  }
  const project = document.getElementById("f_project").value;
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

    // Check the master AND every piece number against every master and
    // piece number already logged, so a piece of an existing multi-piece
    // shipment can't slip in as a "new" shipment (or vice versa).
    const pieceField = document.getElementById("f_pieceTrackingNumbers").value;
    const incoming = [trackingNumber, ...splitTrackingList(pieceField)].filter(Boolean);
    const existing = new Set();
    for (const r of shipmentHistory) {
      if (r.trackingNumber) existing.add(r.trackingNumber);
      for (const t of splitTrackingList(r.pieceTrackingNumbers)) existing.add(t);
    }
    const dup = incoming.find((t) => existing.has(t));
    if (dup) {
      const which = dup === trackingNumber ? "Tracking #" : "Piece tracking #";
      showToast(`${which} ${dup} is already in the shipment history — not adding it again.`, { type: "error" });
      return;
    }

    // Weight: save what's in the (editable) field, not the parser's
    // original value, so a corrected weight actually reaches Excel.
    const weight = parseWeightField(document.getElementById("f_weight").value, currentParsed.totalWeightUnit);
    if (weight === undefined) {
      showToast('Weight should start with a number, e.g. "12.5 LB" — or clear it.', { type: "error" });
      return;
    }

    const shipScope = document.getElementById("f_shipScope").value;
    const priceRaw = document.getElementById("f_price").value;
    const record = {
      trackingNumber,
      shipDate: document.getElementById("f_shipDate").value,
      service: document.getElementById("f_service").value,
      destCountry: document.getElementById("f_destCountry").value.trim().toUpperCase(),
      carrier: document.getElementById("f_carrier").value.trim() || null,
      shipScope: shipScope || null,
      isInternational: shipScope ? shipScope !== "domestic" : currentParsed.isInternational,
      isMultiPiece: currentParsed.isMultiPiece,
      pieceCount: currentParsed.pieceCount,
      pieceTrackingNumbers: document.getElementById("f_pieceTrackingNumbers").value || null,
      totalWeight: weight.value,
      totalWeightUnit: weight.unit,
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
      // Local-only placeholder so the new row sorts to the top right
      // away; replaced by the flow's own utcNow() on the next refresh.
      ts: new Date().toISOString(),
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

// "877031039426, 877031039427" -> ["877031039426", "877031039427"].
// Also strips inner spaces in case a number was typed as "8770 3103 9426".
function splitTrackingList(v) {
  if (v === null || v === undefined) return [];
  return String(v)
    .split(/[,;\n]+/)
    .map((s) => s.replace(/\s+/g, ""))
    .filter(Boolean);
}

// Reads the Weight field. Returns { value, unit } (both null when the
// field is empty), or undefined when there's text but no leading number.
// Accepts the multi-piece format too ("40 LB total (4 pieces: …)") --
// only the leading number + unit are used.
function parseWeightField(raw, fallbackUnit) {
  const s = String(raw || "").trim();
  if (!s) return { value: null, unit: null };
  const m = /^([\d]+(?:\.\d+)?|\.\d+)\s*(LBS?|KGS?)?\b/i.exec(s);
  if (!m) return undefined;
  let unit = m[2] ? m[2].toUpperCase().replace(/S$/, "") : fallbackUnit || null;
  return { value: parseFloat(m[1]), unit };
}

// ---------- Sync to Excel (Power Automate) ----------
// Builds exactly the JSON body the flow's trigger expects -- 26 fields,
// i.e. every Shipments column except `id`/`ts`, which the flow generates itself
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
    carrier: record.carrier || "FedEx",
    shipScope: scopeToExcel(record.shipScope),
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
// Excel's List rows returns booleans as "True"/"False" (capitalized) --
// compare case-insensitively so those aren't all read as false.
function toBool(v) {
  if (v === true || v === 1) return true;
  if (typeof v === "string") return v.trim().toLowerCase() === "true" || v.trim() === "1";
  return false;
}
function toNumOrNull(v) {
  if (v === "" || v === undefined || v === null) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}
// Excel can silently auto-detect a written string as a real Number
// (e.g. a 12-digit tracking number -- shown as scientific notation in
// Excel's own UI, which is just a display quirk, but the List rows API
// hands back the actual number, not the original string). Force text
// fields back to real strings so strict-equality checks elsewhere (the
// duplicate-tracking-number check especially) can't silently fail
// against a row Excel happened to convert.
function toStr(v) {
  return v === null || v === undefined ? "" : String(v);
}
function toStrOrNull(v) {
  return v === null || v === undefined || v === "" ? null : String(v);
}
// Same idea, but for shipDate specifically: Excel can also auto-detect
// a written ISO date string as a real Date-typed cell, in which case
// the API hands back the column's raw date-serial number (e.g. 46261)
// instead of the "2026-08-27" we originally wrote. Convert that serial
// back into the same ISO format so it displays (and sorts/filters, in
// the Expenses tab) the same regardless of which way Excel stored it.
function excelSerialToISODate(serial) {
  const ms = Date.UTC(1899, 11, 30) + Number(serial) * 86400000;
  return new Date(ms).toISOString().slice(0, 10);
}
function normalizeShipDate(v) {
  const s = toStr(v);
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  if (/^\d+(\.\d+)?$/.test(s)) return excelSerialToISODate(s);
  return s;
}

function normalizeHistoryRow(row) {
  return {
    id: row.id,
    ts: row.ts,
    trackingNumber: toStr(row.trackingNumber),
    shipDate: normalizeShipDate(row.shipDate),
    service: toStr(row.service),
    destCountry: toStr(row.destCountry),
    isInternational: toBool(row.isInternational),
    // Older rows (before these columns existed) are blank: treat them as
    // FedEx, with the type derived from isInternational.
    carrier: toStr(row.carrier).trim() || "FedEx",
    shipScope: scopeFromExcel(row.shipScope) || scopeFromBool(toBool(row.isInternational)),
    isMultiPiece: toBool(row.isMultiPiece),
    pieceCount: toNumOrNull(row.pieceCount) ?? 1,
    pieceTrackingNumbers: toStrOrNull(row.pieceTrackingNumbers),
    totalWeight: toNumOrNull(row.totalWeight),
    totalWeightUnit: toStrOrNull(row.totalWeightUnit),
    reference: toStrOrNull(row.reference),
    invoicePoDept: toStrOrNull(row.invoicePoDept),
    senderName: toStr(row.senderName),
    senderAddress: toStr(row.senderAddress),
    senderPhone: toStr(row.senderPhone),
    recipientName: toStr(row.recipientName),
    recipientAddress: toStr(row.recipientAddress),
    recipientPhone: toStr(row.recipientPhone),
    project: toStr(row.project),
    price: toNumOrNull(row.price),
    notes: toStr(row.notes),
    // Anything other than "ok" (any case) -- e.g. "needsReview" -- shows as
    // Review, so fixing a row in Excel just means typing ok here.
    parseStatus: !toStr(row.parseStatus).trim() || toStr(row.parseStatus).trim().toLowerCase() === "ok" ? "ok" : "needsReview",
    sourceFile: toStr(row.sourceFile),
    submittedBy: toStrOrNull(row.submittedBy),
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
    // Keep rows added this session that aren't in Excel yet (still
    // syncing, failed, or never configured) -- otherwise a refresh would
    // silently drop them and take a failed row's retry link with them.
    // Once one shows up in Excel (matched by id or tracking #), the
    // Excel copy wins.
    const freshIds = new Set(fresh.map((r) => r.id).filter(Boolean));
    const freshTracking = new Set(fresh.map((r) => r.trackingNumber).filter(Boolean));
    const pending = shipmentHistory.filter(
      (r) =>
        r.syncStatus !== "synced" &&
        !(r.id && freshIds.has(r.id)) &&
        !(r.trackingNumber && freshTracking.has(r.trackingNumber))
    );
    shipmentHistory.length = 0;
    shipmentHistory.push(...fresh, ...pending);
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
  emptyRow.textContent = historyLoadedOnce ? "No shipments yet." : "Loading shipments from Excel…";
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

    const destPill = scopePill(r);
    const statusPill =
      r.parseStatus === "ok" && !trackingLooksDamaged(r.trackingNumber)
        ? '<span class="pill pill-ok">OK</span>'
        : reviewPill(r);
    const priceStr = r.price != null && !isNaN(r.price) ? `$${r.price.toFixed(2)}` : "—";

    const multiBadge = r.isMultiPiece ? ` <span class="pill pill-muted">×${r.pieceCount}</span>` : "";
    const noteMark = String(r.notes || "").trim()
      ? ` <span class="note-mark" title="Has a note — click the row to read it" aria-label="Has a note">📝</span>`
      : "";
    // Index within the FULL (unsorted, unpaginated) array -- delete/retry
    // splice/act on shipmentHistory itself, not the page-local slice.
    const idx = shipmentHistory.indexOf(r);
    const syncCell = syncStatusCell(r.syncStatus, idx);

    tr.innerHTML = `
      <td>${trackingLink(r.trackingNumber, r.carrier)}${multiBadge}${noteMark}</td>
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
    // Click (or Enter on) a row -> quick overview. The row keeps a direct
    // reference to its record, so a refresh reordering the list can't
    // open the wrong shipment.
    tr.className = "log-row";
    tr.tabIndex = 0;
    tr.title = "Click for a quick overview";
    rowRecords.set(tr, r);
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

// ---------- "Review" details (computed from the row, no flow needed) ----------
// The reader's original warning list isn't saved to Excel, so this
// works out what still needs detail from the row itself: blank fields
// that the label should have filled, plus a guessed service. Each issue
// names the exact Shipments column, so it can be fixed straight in
// Excel. Some original reasons (mixed weight units, missing master
// page) leave no trace in the row and can't be shown here.
const REVIEW_FIELDS = [
  ["trackingNumber", "Tracking number"],
  ["shipDate", "Ship date"],
  ["service", "Service"],
  ["destCountry", "Destination country"],
  ["totalWeight", "Weight"],
  ["senderName", "Sender name"],
  ["senderAddress", "Sender address"],
  ["recipientName", "Recipient name"],
  ["recipientAddress", "Recipient address"],
];

function isBlank(v) {
  return v === null || v === undefined || (typeof v === "number" ? isNaN(v) : String(v).trim() === "");
}

// Excel keeps only 15 significant digits for numbers, so a long tracking
// number (e.g. Australia Post's 23 digits) written into a General-format
// cell comes back rounded -- "9.97230377603010E+22" -- and the real digits
// are gone. Spot that so the row gets flagged for a re-type in Excel.
function trackingLooksDamaged(v) {
  const s = String(v || "").trim();
  return /^\d(?:\.\d+)?e\+\d+$/i.test(s);
}

function reviewIssues(r) {
  const issues = [];
  if (trackingLooksDamaged(r.trackingNumber)) {
    issues.push("Tracking number was rounded by Excel (too long for a number cell) — retype it as text (trackingNumber)");
  }
  for (const [col, label] of REVIEW_FIELDS) {
    if (isBlank(r[col])) issues.push(`${label} is blank (${col})`);
  }
  if (!isBlank(r.totalWeight) && isBlank(r.totalWeightUnit)) issues.push("Weight unit is blank (totalWeightUnit)");
  if (r.isMultiPiece && isBlank(r.pieceTrackingNumbers)) {
    issues.push("Multi-piece, but no piece tracking numbers (pieceTrackingNumbers)");
  }
  return issues;
}

function reviewPill(r) {
  const issues = reviewIssues(r);
  const lines = issues.length
    ? ["Needs detail:", ...issues.map((i) => `• ${i}`), "", "Fix in the Shipments tab, then set parseStatus to ok."]
    : ["Flagged when added, but nothing looks blank now.", "Set parseStatus to ok in the Shipments tab to clear."];
  const count = issues.length ? ` <span class="review-count">${issues.length}</span>` : "";
  return `<span class="pill pill-review review-tip" tabindex="0" data-tip="${escapeAttr(lines.join("\n"))}" aria-label="${escapeAttr(lines.join(" "))}">Review${count}</span>`;
}

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
    return;
  }
  // Links (FedEx tracking) and the Review pill (tap shows its hover) keep
  // their own behavior; anywhere else on the row opens the overview.
  if (e.target.closest("a, button, .review-tip")) return;
  const tr = e.target.closest("tr.log-row");
  if (tr && rowRecords.has(tr)) openShipmentOverview(rowRecords.get(tr));
});
document.getElementById("logTableBody").addEventListener("keydown", (e) => {
  if (e.key !== "Enter" || !e.target.matches("tr.log-row")) return;
  // preventDefault: otherwise the same Enter "clicks" the overview's ✕
  // button (focus moves there on open) and closes it instantly.
  e.preventDefault();
  if (rowRecords.has(e.target)) openShipmentOverview(rowRecords.get(e.target));
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

  // Re-find the row now rather than trusting `idx`: a history refresh
  // can run while the confirm dialog is open and reorder/replace the
  // array, and splicing a stale index would remove the wrong row.
  const removeLocally = () => {
    let i = shipmentHistory.indexOf(r);
    if (i === -1 && r.id) i = shipmentHistory.findIndex((x) => x.id === r.id);
    if (i !== -1) shipmentHistory.splice(i, 1);
  };

  if (!canDeleteFromExcel) {
    removeLocally();
    renderLog();
    renderExpenses();
    return;
  }

  try {
    await deleteRecordFromExcel(r);
    removeLocally();
    renderLog();
    renderExpenses();
    showToast("Shipment permanently deleted from Excel.", { type: "success" });
  } catch (err) {
    console.error("Delete from Excel failed:", err);
    showToast("Couldn't delete that row from Excel — nothing was removed. Try again.", { type: "error" });
  }
}

// ---------- Shipment overview (click a history row) ----------
// A quick read-only card: who sent it where (sender -> recipient), the
// key facts, and the notes. Close with ✕, Esc, or a click outside.
const overview = document.getElementById("shipmentOverview");
let overviewReturnFocus = null;

function ovText(v, fallback = "—") {
  const s = v === null || v === undefined ? "" : String(v).trim();
  return s ? escapeHtml(s) : `<span class="ov-empty">${fallback}</span>`;
}
function ovParty(role, name, address, phone, country) {
  return `
    <div class="ov-party">
      <div class="ov-role">${role}</div>
      <div class="ov-name">${ovText(name, "No name")}</div>
      <div class="ov-addr">${ovText(address, "No address")}</div>
      ${phone ? `<div class="ov-phone">📞 ${escapeHtml(phone)}</div>` : ""}
      ${country ? `<div class="ov-country">${escapeHtml(country)}</div>` : ""}
    </div>`;
}

function openShipmentOverview(r) {
  overviewReturnFocus = document.activeElement;
  const scope = rowScope(r);
  const travel = scope === "domestic" ? "🚚" : "✈️";
  const carrier = r.carrier || "FedEx";
  const priceStr = r.price != null && !isNaN(r.price) ? `$${Number(r.price).toFixed(2)}` : null;
  const weightStr = r.totalWeight != null && !isNaN(r.totalWeight) ? `${r.totalWeight} ${r.totalWeightUnit || ""}`.trim() : null;
  const pieces = r.isMultiPiece ? `${r.pieceCount} boxes${r.pieceTrackingNumbers ? ` · ${r.pieceTrackingNumbers}` : ""}` : "1 box";
  const facts = [
    ["Carrier", carrier],
    ["Service", r.service],
    ["Ship date", r.shipDate],
    ["Shipment type", SCOPE_LABELS[scope]],
    ["Weight", weightStr],
    ["Pieces", pieces],
    ["Project", r.project],
    ["Price", priceStr],
    ["Submitted by", r.submittedBy],
    ["Reference", r.reference],
    ["Invoice / PO / Dept", r.invoicePoDept],
    ["Source file", r.sourceFile],
  ];
  const note = String(r.notes || "").trim();

  document.getElementById("ovTitle").innerHTML =
    `${r.trackingNumber ? escapeHtml(r.trackingNumber) : "No tracking #"} <span class="ov-sub">${escapeHtml(carrier)}</span>`;
  document.getElementById("ovPills").innerHTML = scopePill(r);
  document.getElementById("ovBody").innerHTML = `
    <div class="ov-route">
      ${ovParty("From", r.senderName, r.senderAddress, r.senderPhone, "")}
      <div class="ov-arrow" aria-hidden="true">
        <span class="ov-line"></span><span class="ov-icon">${travel}</span><span class="ov-line"></span>
        <div class="ov-arrow-label">${ovText(r.service, carrier)}${r.shipDate ? `<br>${escapeHtml(r.shipDate)}` : ""}</div>
      </div>
      ${ovParty("To", r.recipientName, r.recipientAddress, r.recipientPhone, r.destCountry)}
    </div>
    <dl class="ov-facts">
      ${facts.map(([k, v]) => `<div><dt>${k}</dt><dd>${ovText(v)}</dd></div>`).join("")}
    </dl>
    <div class="ov-notes">
      <div class="ov-role">📝 Notes</div>
      ${note ? `<p>${escapeHtml(note)}</p>` : `<p class="ov-empty">No notes for this shipment.</p>`}
    </div>`;
  overview.hidden = false;
  document.body.classList.add("lightbox-open");
  document.getElementById("ovClose").focus();
}

function closeShipmentOverview() {
  if (overview.hidden) return;
  overview.hidden = true;
  document.body.classList.remove("lightbox-open");
  if (overviewReturnFocus && overviewReturnFocus.focus) overviewReturnFocus.focus();
}
document.getElementById("ovClose").addEventListener("click", closeShipmentOverview);
overview.addEventListener("click", (e) => {
  if (e.target === overview) closeShipmentOverview();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !overview.hidden) closeShipmentOverview();
});

// ---------- Expenses by Project tab ----------
// (Project / Submitted-by filter options are filled by
// populateProjectDropdowns() once lookups load from Excel.)

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
    if (destination && rowScope(r) !== destination) return false;
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
      const destPill = scopePill(r);
      const priceStr = r.price != null && !isNaN(r.price) ? `$${r.price.toFixed(2)}` : "—";
      const multiBadge = r.isMultiPiece ? ` <span class="pill pill-muted">×${r.pieceCount}</span>` : "";
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${trackingLink(r.trackingNumber, r.carrier)}${multiBadge}</td>
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

// For attribute values: escapeHtml() doesn't escape quotes.
function escapeAttr(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "&#10;");
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
// ---------- Shipment type (Domestic / International / International-domestic) ----------
// Stored in the Shipments `shipScope` column as the readable label, so
// it's easy to edit straight in Excel; used in code as a short key.
const SCOPE_LABELS = {
  domestic: "Domestic",
  international: "International",
  intlDomestic: "International-domestic",
};
function scopeToExcel(code) {
  return SCOPE_LABELS[code] || "";
}
function scopeFromExcel(v) {
  const s = toStr(v).trim().toLowerCase().replace(/[\s_]+/g, "-");
  if (!s) return "";
  if (s === "domestic") return "domestic";
  if (s === "international") return "international";
  if (s === "international-domestic" || s === "intl-domestic" || s === "intldomestic") return "intlDomestic";
  return "";
}
function scopeFromBool(isInternational) {
  if (isInternational === null || isInternational === undefined) return "";
  return isInternational ? "international" : "domestic";
}
function rowScope(r) {
  return r.shipScope || scopeFromBool(r.isInternational) || "domestic";
}
function scopePill(r) {
  const scope = rowScope(r);
  if (scope === "intlDomestic") {
    return '<span class="pill pill-intldom" title="International-domestic: shipped within another country">Intl-Dom</span>';
  }
  if (scope === "international") return '<span class="pill pill-intl">Intl</span>';
  return '<span class="pill pill-domestic">US</span>';
}

function trackingLink(trackingNumber, carrier) {
  if (!trackingNumber) return "—";
  if (trackingLooksDamaged(trackingNumber)) {
    return `<span class="tracking-damaged" title="Rounded by Excel — retype it as text in the Shipments tab">${escapeHtml(trackingNumber)}</span>${carrier ? `<div class="carrier-note">${escapeHtml(carrier)}</div>` : ""}`;
  }
  // Only FedEx numbers get a link (to FedEx's tracker). Another carrier's
  // number shown as plain text, with the carrier named underneath.
  if (carrier && !/^fedex$/i.test(String(carrier).trim())) {
    return `${escapeHtml(trackingNumber)}<div class="carrier-note">${escapeHtml(carrier)}</div>`;
  }
  const url = `https://www.fedex.com/fedextrack/?trknbr=${encodeURIComponent(trackingNumber)}`;
  return `<a href="${url}" target="_blank" rel="noopener">${escapeHtml(trackingNumber)}</a>`;
}

// Initial load: pull real history from Excel (if the read flow's
// connected) before the first render, so the table doesn't flash empty
// and then repopulate a moment later.
(async function initHistory() {
  await refreshHistoryFromExcel();
  historyLoadedOnce = true;
  renderLog();
  renderExpenses();
})();
