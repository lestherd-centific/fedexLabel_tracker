/**
 * FedEx Ship Manager label parser — browser port of fedex_label_parser.py.
 *
 * Uses pdf.js to extract text. Line reconstruction relies on each text
 * item's `hasEOL` flag (validated against pdf.js's Node build to produce
 * output matching PyMuPDF's rotation-aware extraction line-for-line —
 * no manual coordinate/rotation math needed).
 *
 * pdfjsLib must already be loaded on the page (script tag) and its
 * workerSrc configured before calling parseFedexLabel().
 */

const MONTHS = {
  JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6,
  JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12,
};

// Lines that are label metadata, never part of an address block.
const METADATA_PREFIXES = ["SIGN:", "BILL ", "NO EEI", "CAD:", "ORIGIN ID:"];

async function loadLabelPage(file) {
  const buf = await file.arrayBuffer();
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
  return doc.getPage(1);
}

async function extractLabelLines(page) {
  const textContent = await page.getTextContent();

  const lines = [];
  let cur = "";
  for (const item of textContent.items) {
    cur += item.str;
    if (item.hasEOL) {
      lines.push(cur.trim());
      cur = "";
    } else {
      cur += " ";
    }
  }
  if (cur.trim()) lines.push(cur.trim());

  return lines.filter((l) => l.length > 0);
}

// Renders the label to a PNG data URL so the confirm screen can show the
// original next to the parsed fields -- targetWidth is a CSS-pixel width,
// scaled by devicePixelRatio for a crisp image on high-DPI screens.
async function renderLabelPreview(page, targetWidth = 340) {
  const dpr = window.devicePixelRatio || 1;
  const baseViewport = page.getViewport({ scale: 1 });
  const scale = (targetWidth * dpr) / baseViewport.width;
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement("canvas");
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext("2d");
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas.toDataURL("image/png");
}

function parseShipDate(raw) {
  const m = /^(\d{2})([A-Z]{3})(\d{2})/.exec(raw);
  if (!m) return null;
  const [, day, mon, yr] = m;
  const month = MONTHS[mon];
  if (!month) return null;
  const year = 2000 + parseInt(yr, 10);
  const d = new Date(Date.UTC(year, month - 1, parseInt(day, 10)));
  if (isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

// REF:/INV:/PO:/DEPT: sometimes land on the same reconstructed line
// separated only by spaces (e.g. "PO:   DEPT: 000") rather than a
// newline. A blank field (e.g. PO blank, DEPT: right after it) needs a
// lookahead so [ \t]* skipping past the spaces doesn't then let \S*
// swallow the next label as if it were this field's value.
const RESERVED_FIELD_LABELS = ["REF:", "INV:", "PO:", "DEPT:"];

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function fieldValue(text, label) {
  const lookahead = RESERVED_FIELD_LABELS.map(escapeRe).join("|");
  const re = new RegExp(escapeRe(label) + `[ \\t]*((?:(?!${lookahead})\\S)*)`);
  const m = re.exec(text);
  return m && m[1] ? m[1] : null;
}

function extractAddressBlock(lines, startIdx, endIdx) {
  const block = [];
  for (const line of lines.slice(startIdx, endIdx)) {
    if (!line) continue;
    if (METADATA_PREFIXES.some((p) => line.startsWith(p))) continue;
    block.push(line);
  }
  return block;
}

// A line that's just digits/parens/dashes/spaces, at least 7 characters
// of them -- used to spot a phone number sitting as its own line at the
// end of an address block (true on every sample seen: domestic and
// international alike put the recipient's phone as the last address
// line, and the sender's own block never has a bare-digit line like
// this that isn't a phone).
const PHONE_LINE_RE = /^[(+]?[\d][\d\s\-().]{6,}$/;

// First line of a block is always the name; if the last remaining line
// looks like a phone, split it off; whatever's left in between is the
// address, joined into one string.
function decomposeBlock(block) {
  if (!block.length) return { name: null, address: null, phone: null };
  const name = block[0];
  let rest = block.slice(1);
  let phone = null;
  if (rest.length && PHONE_LINE_RE.test(rest[rest.length - 1])) {
    phone = rest[rest.length - 1];
    rest = rest.slice(0, -1);
  }
  return { name, address: rest.length ? rest.join(", ") : null, phone };
}

function parseLabelText(lines) {
  const text = lines.join("\n");

  const result = {
    trackingNumber: null,
    shipDateRaw: null,
    shipDate: null,
    weight: null,
    weightUnit: null,
    dimensions: null,
    reference: null,
    invoice: null,
    po: null,
    dept: null,
    originId: null,
    cad: null,
    billTo: null,
    destCountry: null,
    isInternational: null,
    service: null,
    serviceConfidence: null,
    sender: { name: null, address: null, phone: null },
    recipient: { name: null, address: null, phone: null },
    parseStatus: "ok",
    warnings: [],
  };

  // --- Tracking number: "#### #### ####" barcode caption ---
  let m = /\b(\d{4}\s\d{4}\s\d{4})\b/.exec(text);
  if (m) {
    result.trackingNumber = m[1].replace(/\s/g, "");
  } else {
    result.warnings.push("trackingNumber not found");
  }

  // --- Ship date ---
  m = /SHIP DATE:\s*(\d{2}[A-Z]{3}\d{2})/.exec(text);
  if (m) {
    result.shipDateRaw = m[1];
    result.shipDate = parseShipDate(m[1]);
  } else {
    result.warnings.push("shipDate not found");
  }

  // --- Weight ---
  m = /ACTWGT:\s*([\d.]+)\s*(LB|KG)/.exec(text);
  if (m) {
    result.weight = parseFloat(m[1]);
    result.weightUnit = m[2];
  } else {
    result.warnings.push("weight not found");
  }

  // --- Dimensions (not always present) ---
  m = /DIMS:\s*([\dx]+)\s*(IN|CM)/.exec(text);
  if (m) result.dimensions = `${m[1]} ${m[2]}`;

  // --- Simple labeled fields ---
  // Regex against the full text (not per-line) because pdf.js's hasEOL
  // line breaks land in slightly different places than PyMuPDF's did --
  // e.g. "PO:" and "DEPT:" sometimes land on the same reconstructed line
  // ("PO: 000   DEPT: 000"). [ \t]* (not \s*) keeps the match from
  // crossing a newline when a field is blank, so it doesn't swallow the
  // next label's value.
  result.reference = fieldValue(text, "REF:");
  result.invoice = fieldValue(text, "INV:");
  result.po = fieldValue(text, "PO:");
  result.dept = fieldValue(text, "DEPT:");

  m = /ORIGIN ID:\s*([A-Z]+)/.exec(text);
  if (m) result.originId = m[1];

  m = /CAD:\s*(\S+)/.exec(text);
  if (m) result.cad = m[1];

  m = /\b(BILL (?:SENDER|RECIPIENT|THIRD PARTY))\b/.exec(text);
  if (m) result.billTo = m[1];

  // --- International vs domestic: "XX-YY" destination sort code ---
  const countryMatches = [...text.matchAll(/\b[A-Z]{2}-([A-Z]{2})\b/g)];
  if (countryMatches.length) {
    result.destCountry = countryMatches[countryMatches.length - 1][1];
    result.isInternational = result.destCountry !== "US";
  } else {
    result.warnings.push("destCountry not found (sort code pattern missing)");
  }

  // --- Service type ---
  m = /\*\*\s*(.+?)\s*\*\*/.exec(text);
  if (m) {
    result.service = m[1].trim();
    result.serviceConfidence = "labeled";
  } else if (/\bIP\b/.test(text)) {
    // Seen on real international samples (US->AU, US->GB), always paired
    // with a short delivery-commitment code (EOD/EXP). Not a documented
    // field label -- an inference, flagged for confirmation, not a
    // certain read.
    result.service = "International Priority";
    result.serviceConfidence = "inferred from 'IP' code — confirm";
    result.warnings.push(
      "service type inferred from 'IP' code, not directly labeled — confirm before saving"
    );
  } else {
    result.serviceConfidence = null;
    result.warnings.push("service type not confidently parsed — enter manually");
  }

  // --- Address blocks (best effort) ---
  const originIdx = lines.findIndex((l) => l.startsWith("ORIGIN ID:"));
  const cadIdx = lines.findIndex((l) => l.startsWith("CAD:"));

  let senderBlock = [];
  if (originIdx !== -1 && cadIdx !== -1 && cadIdx > originIdx) {
    // pdf.js's hasEOL-based line breaks merge "ORIGIN ID:xxxx" with the
    // account-number/phone line that follows it into one reconstructed
    // line (PyMuPDF kept them as two separate lines) -- so the sender
    // name starts right after the ORIGIN ID line here, not two lines
    // after it.
    senderBlock = extractAddressBlock(lines, originIdx + 1, cadIdx);
  }

  let recipientBlock = [];
  if (cadIdx !== -1) {
    let endIdx = lines.length;
    for (let i = cadIdx + 1; i < lines.length; i++) {
      if (/^\*\*.*\*\*$/.test(lines[i]) || /^\S+\/\S+\/\S+$/.test(lines[i])) {
        endIdx = i;
        break;
      }
    }
    let recipientLines = extractAddressBlock(lines, cadIdx + 1, endIdx);
    if (recipientLines.length && recipientLines[0].startsWith("TO ")) {
      recipientLines[0] = recipientLines[0].slice(3).trim();
    }
    for (let i = 0; i < recipientLines.length; i++) {
      if (PHONE_LINE_RE.test(recipientLines[i])) {
        recipientLines = recipientLines.slice(0, i + 1);
        break;
      }
    }
    recipientBlock = recipientLines;
  }

  result.sender = decomposeBlock(senderBlock);
  result.recipient = decomposeBlock(recipientBlock);

  // The sender's phone (when there is one) is merged into the ORIGIN ID
  // line itself, not a separate block line -- e.g.
  // "ORIGIN ID:OTSA (206) 683-4772". Only pull it out when it's in the
  // unambiguous US "(nnn) nnn-nnnn" format: that same line can instead
  // hold a bare digit string (e.g. "14709193520" on the domestic
  // sample) that's FedEx's own account/meter number, not a phone --
  // and a bare string like that is visually indistinguishable from a
  // foreign phone number, so it's deliberately left alone rather than
  // guessed at.
  if (!result.sender.phone && originIdx !== -1) {
    const phoneMatch = /(\(\d{3}\)\s?\d{3}-\d{4})/.exec(lines[originIdx]);
    if (phoneMatch) result.sender.phone = phoneMatch[1];
  }

  if (!result.sender.name) result.warnings.push("sender name/address not confidently extracted");
  if (!result.recipient.name) result.warnings.push("recipient name/address not confidently extracted");

  if (result.warnings.length) result.parseStatus = "needsReview";

  return result;
}

async function parseFedexLabel(file) {
  const page = await loadLabelPage(file);
  const lines = await extractLabelLines(page);
  const parsed = parseLabelText(lines);
  parsed.sourceFile = file.name;
  try {
    parsed.previewDataUrl = await renderLabelPreview(page);
  } catch (err) {
    // A rendering failure shouldn't block the parse itself -- the
    // confirm screen just won't have a preview image for this one.
    console.error("Label preview render failed:", err);
    parsed.previewDataUrl = null;
  }
  return parsed;
}
