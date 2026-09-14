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
const sessionLog = [];

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
  document.getElementById("loginBackdrop").hidden = true;
  document.getElementById("app").hidden = false;
  document.getElementById("whoami").hidden = false;
  document.getElementById("whoamiName").textContent = `${cred.name} (${cred.role})`;
}

function showLogin() {
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
    const parsed = await parseFedexLabel(file);
    currentParsed = parsed;
    populateConfirmForm(parsed);
  } catch (err) {
    console.error(err);
    alert("Couldn't read that PDF. It may not be a FedEx Ship Manager label, or the file is corrupted.");
  } finally {
    dropzoneIdle.hidden = false;
    dropzoneBusy.hidden = true;
    fileInput.value = "";
  }
}

// ---------- Confirm form ----------
function populateConfirmForm(parsed) {
  document.getElementById("sourceFileName").textContent = parsed.sourceFile;
  document.getElementById("f_trackingNumber").value = parsed.trackingNumber || "";
  document.getElementById("f_shipDate").value = parsed.shipDate || "";
  document.getElementById("f_service").value = parsed.service || "";
  document.getElementById("f_destCountry").value = parsed.destCountry || "";
  document.getElementById("f_weight").value = parsed.weight ? `${parsed.weight} ${parsed.weightUnit}` : "";
  document.getElementById("f_dimensions").value = parsed.dimensions || "";
  document.getElementById("f_reference").value = parsed.reference || "";

  const invPoDeptParts = [];
  if (parsed.invoice) invPoDeptParts.push(`INV ${parsed.invoice}`);
  if (parsed.po) invPoDeptParts.push(`PO ${parsed.po}`);
  if (parsed.dept) invPoDeptParts.push(`DEPT ${parsed.dept}`);
  document.getElementById("f_invPoDept").value = invPoDeptParts.join(" · ");

  document.getElementById("f_sender").value = (parsed.senderBlock || []).join("\n");
  document.getElementById("f_recipient").value = (parsed.recipientBlock || []).join("\n");

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
  currentParsed = null;
  document.getElementById("confirmCard").hidden = true;
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
    project,
    price: priceRaw ? parseFloat(priceRaw) : null,
    notes: document.getElementById("f_notes").value,
    parseStatus: currentParsed.parseStatus,
    sourceFile: currentParsed.sourceFile,
  };
  sessionLog.push(record);
  renderLog();

  currentParsed = null;
  document.getElementById("confirmCard").hidden = true;
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

    tr.innerHTML = `
      <td>${escapeHtml(r.trackingNumber || "—")}</td>
      <td>${escapeHtml(r.shipDate || "—")}</td>
      <td>${escapeHtml(r.service || "—")}</td>
      <td>${destPill} ${escapeHtml(r.destCountry || "")}</td>
      <td>${escapeHtml(r.project)}</td>
      <td>${priceStr}</td>
      <td>${statusPill}</td>
    `;
    body.appendChild(tr);
  }
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

renderLog();
