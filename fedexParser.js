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

async function loadLabelDoc(file) {
  const buf = await file.arrayBuffer();
  return pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
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
//
// Labels printed with the other team's settings (policy/terms block on
// top) carry a page /Rotate of 270 -- first seen on international
// samples (US->AU, US->GB), but it's a print setting, not an
// international trait, so any label can have it. Letting
// pdf.js apply that automatically (the default when no `rotation` is
// passed to getViewport -- it falls back to page.rotate) renders the
// label upside down: verified empirically by rendering the same page
// with PyMuPDF both ways -- honoring /Rotate=270 produced a portrait
// image with every field inverted, while forcing rotation to 0 (i.e.
// ignoring /Rotate entirely and using the page's raw, un-rotated
// content box) produced the correct, upright, readable label -- a
// landscape image, since that's genuinely the shape FedEx prints the
// sheet in before it's folded. Labels without those settings have
// /Rotate=0 already, so forcing rotation:0 is a no-op for them --
// either way the preview comes out upright, domestic or international.
async function renderLabelPreview(page, targetWidth = 340, { trim = false } = {}) {
  const dpr = window.devicePixelRatio || 1;
  const baseViewport = page.getViewport({ scale: 1, rotation: 0 });
  // trim: crop to the printed content (used for manual entries, whose
  // labels -- e.g. Australia Post on A4 -- can fill only a corner of
  // the page). Box is in fractions of the page, found once per page.
  const box = trim ? await findContentBox(page) : { x: 0, y: 0, w: 1, h: 1 };
  const scale = (targetWidth * dpr) / (baseViewport.width * box.w);
  const viewport = page.getViewport({ scale, rotation: 0 });

  const canvas = document.createElement("canvas");
  canvas.width = Math.round(viewport.width * box.w);
  canvas.height = Math.round(viewport.height * box.h);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.translate(-viewport.width * box.x, -viewport.height * box.y);
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas.toDataURL("image/png");
}

// Finds the bounding box of anything non-white on the page from a quick
// low-res render, with a little padding. Falls back to the full page if
// the page is blank or the content already fills most of it.
async function findContentBox(page) {
  if (page._contentBox) return page._contentBox;
  const full = { x: 0, y: 0, w: 1, h: 1 };
  try {
    const base = page.getViewport({ scale: 1, rotation: 0 });
    const viewport = page.getViewport({ scale: 400 / base.width, rotation: 0 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    const { data, width, height } = ctx.getImageData(0, 0, canvas.width, canvas.height);
    let minX = width, minY = height, maxX = -1, maxY = -1;
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        if (data[i] < 235 || data[i + 1] < 235 || data[i + 2] < 235) {
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }
    let box = full;
    if (maxX >= 0) {
      const pad = 0.02;
      const x = Math.max(0, minX / width - pad), y = Math.max(0, minY / height - pad);
      const w = Math.min(1, (maxX + 1) / width + pad) - x, h = Math.min(1, (maxY + 1) / height + pad) - y;
      if (w * h < 0.85) box = { x, y, w, h };
    }
    page._contentBox = box;
    return box;
  } catch (err) {
    console.error("Couldn't find the label's content area; showing the full page:", err);
    return full;
  }
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

// The team's FedEx Ship Manager service list (plus a few common ones),
// as printed on labels -> the name saved to Excel. Checked in order, so
// the more specific names (e.g. "... Freight", "... Express") come
// before the shorter ones they contain. Matching is on whole words with
// flexible spacing ("2 DAY" / "2DAY", "INTL" / "INTERNATIONAL").
const INTL = "INT(?:ERNATIONA)?'?L\\.?";
const KNOWN_SERVICES = [
  [`${INTL}\\s*PRIORITY\\s*EXPRESS`, "FedEx International Priority Express"],
  [`${INTL}\\s*PRIORITY\\s*FREIGHT`, "FedEx International Priority Freight"],
  [`${INTL}\\s*ECONOMY\\s*FREIGHT`, "FedEx International Economy Freight"],
  [`${INTL}\\s*DEFERRED\\s*FREIGHT`, "FedEx International Deferred Freight"],
  [`${INTL}\\s*CONNECT\\s*PLUS`, "FedEx International Connect Plus"],
  [`${INTL}\\s*FIRST`, "FedEx International First"],
  [`${INTL}\\s*PRIORITY`, "FedEx International Priority"],
  [`${INTL}\\s*ECONOMY`, "FedEx International Economy"],
  ["FIRST\\s*OVERNIGHT\\s*FREIGHT", "FedEx First Overnight Freight"],
  ["FIRST\\s*OVERNIGHT", "FedEx First Overnight"],
  ["PRIORITY\\s*OVERNIGHT", "FedEx Priority Overnight"],
  ["STANDARD\\s*OVERNIGHT", "FedEx Standard Overnight"],
  ["1\\s*DAY\\s*FREIGHT", "FedEx 1Day Freight"],
  ["FREIGHT\\s*PRIORITY", "FedEx Freight Priority"],
  ["FREIGHT\\s*ECONOMY", "FedEx Freight Economy"],
  ["2\\s*DAY\\s*A\\.?M\\.?", "FedEx 2Day A.M."],
  ["2\\s*DAY", "FedEx 2Day"],
  ["EXPRESS\\s*SAVER", "FedEx Express Saver"],
  ["HOME\\s*DELIVERY", "FedEx Home Delivery"],
  // "GROUND" alone could be an address ("GROUND FLOOR"), so only
  // "FEDEX GROUND" or a line that is just "GROUND" counts.
  ["FEDEX\\s*GROUND|^\\s*GROUND\\s*$", "FedEx Ground"],
].map(([re, name]) => [new RegExp(`(?:^|[^A-Z0-9])(?:${re})(?![A-Z0-9])`, "im"), name]);

// Only short, headline-style lines are searched: the policy / terms
// panel some labels carry can mention service names in its fine print,
// and those long sentences mustn't be read as this label's service.
const MAX_SERVICE_LINE_LEN = 60;

function matchKnownService(lines) {
  const upper = lines
    .filter((l) => l.length <= MAX_SERVICE_LINE_LEN)
    .join("\n")
    .toUpperCase()
    .replace(/®/g, "");
  for (const [re, name] of KNOWN_SERVICES) if (re.test(upper)) return name;
  return null;
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
  // On an MPS piece page, "Mstr# #### #### ####" (the master's number)
  // appears near the top, BEFORE this piece's own tracking number near
  // the bottom -- so take the last match in the text, not the first,
  // or a piece would incorrectly come back with its master's number.
  // Harmless on every other label, which only ever has one such match.
  const trackingMatches = [...text.matchAll(/\b(\d{4}\s\d{4}\s\d{4})\b/g)];
  if (trackingMatches.length) {
    result.trackingNumber = trackingMatches[trackingMatches.length - 1][1].replace(/\s/g, "");
  } else {
    result.warnings.push("trackingNumber not found");
  }

  // --- Ship date ---
  let m = /SHIP DATE:\s*(\d{2}[A-Z]{3}\d{2})/.exec(text);
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

  // --- International vs domestic ---
  // Two independent signals; either one makes a shipment international:
  //   1. the "IP" (International Priority) service code on the label, and
  //   2. an "XX-YY" destination sort code whose country isn't US.
  // Neither depends on the page layout: the rotated sheet with the
  // policy/terms block on top comes from the other team's print
  // settings and can appear on ANY label, domestic or international,
  // so it's never used as a signal here.
  const hasIpCode = /\bIP\b/.test(text);
  const countryMatches = [...text.matchAll(/\b[A-Z]{2}-([A-Z]{2})\b/g)];
  if (countryMatches.length) {
    result.destCountry = countryMatches[countryMatches.length - 1][1];
  } else {
    result.warnings.push("destCountry not found (sort code pattern missing)");
  }
  if (hasIpCode || (result.destCountry && result.destCountry !== "US")) {
    result.isInternational = true;
  } else if (result.destCountry === "US") {
    result.isInternational = false;
  } // else: no sort code and no IP -- unknown, left null (pill hidden)

  // --- Service type ---
  // 1. A known FedEx service name printed anywhere on the label (e.g.
  //    "PRIORITY OVERNIGHT", "2DAY"), normalized to the team's Ship
  //    Manager name ("FedEx Priority Overnight", "FedEx 2Day").
  // 2. Otherwise, whatever is printed between "** … **" as-is.
  // 3. Otherwise, the "IP" code -> FedEx International Priority. That code
  //    is the real indicator, so all three are confident reads.
  const known = matchKnownService(lines);
  m = /\*\*\s*(.+?)\s*\*\*/.exec(text);
  if (known) {
    result.service = known;
    result.serviceConfidence = "labeled";
  } else if (m) {
    result.service = m[1].trim();
    result.serviceConfidence = "labeled";
  } else if (hasIpCode) {
    result.service = "FedEx International Priority";
    result.serviceConfidence = "labeled";
  } else {
    result.serviceConfidence = null;
    result.warnings.push("service type not confidently parsed — enter manually");
  }
  // Any FedEx International service also makes the shipment international.
  if (result.service && /^FedEx International\b/.test(result.service)) result.isInternational = true;
  if (result.isInternational && result.destCountry === "US") {
    result.warnings.push("label shows an international service but a US destination — check destination");
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

// FedEx's Multiple Piece Shipment (MPS) convention: page 1 of a
// multi-piece shipment is marked "## MASTER ## " plus a standalone
// "1 of N" line; the remaining pages are marked "MPS#" plus
// "Mstr# <that same tracking number>" and their own "k of N" line. A
// plain single-piece label has none of these markers at all.
function detectPieceMeta(text) {
  const isMaster = /##\s*MASTER\s*##/.test(text);
  const posMatch = /(?:^|\n)[ \t]*(\d+)\s+of\s+(\d+)[ \t]*(?:\n|$)/.exec(text);
  const position = posMatch ? parseInt(posMatch[1], 10) : null;
  const total = posMatch ? parseInt(posMatch[2], 10) : null;
  const mstrMatch = /Mstr#\s*(\d{4}\s\d{4}\s\d{4})/.exec(text);
  const mstrRef = mstrMatch ? mstrMatch[1].replace(/\s/g, "") : null;
  const hasPieceMarkers = isMaster || position !== null || mstrRef !== null || /\bMPS#/.test(text);
  return { isMaster, position, total, mstrRef, hasPieceMarkers };
}

// Parses every page of the PDF, then groups pages into shipment records.
// Most files are the simple case: one page, no MPS markers -> one group
// with one piece, same as before. A true multi-piece shipment (several
// pages all pointing at the same master tracking number) becomes one
// group with several pieces. Pages that carry NO piece markers and
// don't share a master are each their own independent shipment --
// e.g. someone batch-downloading a handful of unrelated single-piece
// labels into one PDF -- so those come back as separate groups too,
// rather than being incorrectly merged into one record.
async function parseFedexLabelFile(file) {
  const doc = await loadLabelDoc(file);
  const pageParses = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const lines = await extractLabelLines(page);
    const text = lines.join("\n");
    const parsed = parseLabelText(lines);
    // Skip pages that aren't a label at all (e.g. a policy / terms or
    // folding-instructions page printed alongside it): no tracking
    // number, ship date or weight found anywhere on the page. Without
    // this, such a page would come back as its own bogus "shipment".
    if (!parsed.trackingNumber && !parsed.shipDate && parsed.weight == null) continue;
    const pieceMeta = detectPieceMeta(text);
    pageParses.push({ pageNumber: i, page, parsed, pieceMeta });
  }

  // Pass 1: every master / standalone page starts its own group, keyed
  // by its own tracking number.
  const groups = new Map(); // key -> { referenceIndex, pieceIndices: [] }
  pageParses.forEach((pp, idx) => {
    const { isMaster, hasPieceMarkers } = pp.pieceMeta;
    if (isMaster || !hasPieceMarkers) {
      const key = pp.parsed.trackingNumber || `__no-tracking-${idx}`;
      groups.set(key, { referenceIndex: idx, pieceIndices: [idx] });
    }
  });

  // Pass 2: every subordinate piece attaches to its referenced master's
  // group -- order-independent, so it doesn't matter whether the master
  // page physically comes first in the file.
  pageParses.forEach((pp, idx) => {
    const { isMaster, hasPieceMarkers, mstrRef } = pp.pieceMeta;
    if (isMaster || !hasPieceMarkers) return; // already handled in pass 1
    const key = mstrRef || pp.parsed.trackingNumber || `__no-tracking-${idx}`;
    if (groups.has(key)) {
      groups.get(key).pieceIndices.push(idx);
    } else {
      // This piece's master page isn't in this PDF at all -- still
      // group it (as its own shipment, best effort) rather than drop
      // it, but flag it so the confirm screen shows why.
      groups.set(key, { referenceIndex: idx, pieceIndices: [idx], missingMaster: true });
    }
  });

  // Build one shipment record per group.
  const records = [];
  for (const { referenceIndex, pieceIndices, missingMaster } of groups.values()) {
    const ref = pageParses[referenceIndex];
    const pieces = pieceIndices
      .slice()
      .sort((a, b) => a - b)
      .map((idx) => {
        const pp = pageParses[idx];
        return {
          pageNumber: pp.pageNumber,
          trackingNumber: pp.parsed.trackingNumber,
          weight: pp.parsed.weight,
          weightUnit: pp.parsed.weightUnit,
        };
      });

    const record = { ...ref.parsed }; // shared fields come from the reference (master) page
    record.carrier = "FedEx";
    record.isManual = false;
    record.warnings = [...ref.parsed.warnings]; // clone -- about to push onto it below
    record.sourceFile = file.name;
    record.isMultiPiece = pieces.length > 1;
    record.pieceCount = pieces.length;
    record.pieces = pieces;

    const weighable = pieces.filter((p) => p.weight != null);
    record.totalWeight = weighable.length ? weighable.reduce((sum, p) => sum + p.weight, 0) : null;
    record.totalWeightUnit = weighable.length ? weighable[0].weightUnit : null;
    const mixedUnits = new Set(weighable.map((p) => p.weightUnit)).size > 1;
    if (mixedUnits) record.warnings.push("pieces show mixed weight units — total may not be meaningful");

    if (missingMaster) {
      record.warnings.push(
        "this piece references a master label that isn't in this PDF — shared fields came from this piece itself, not a true master page"
      );
    }
    if (record.warnings.length) record.parseStatus = "needsReview";

    try {
      record.previewDataUrl = await renderLabelPreview(ref.page);
      // Kept so the confirm screen can re-render the label much larger
      // (the "click to enlarge" view) instead of blowing up the small
      // thumbnail. Only used in the browser; never sent to Excel.
      record.previewPage = ref.page;
    } catch (err) {
      console.error("Label preview render failed:", err);
      record.previewDataUrl = null;
    }

    records.push(record);
  }

  // Stable order: by the reference page's position in the file.
  records.sort((a, b) => a.pieces[0].pageNumber - b.pieces[0].pageNumber);
  return records;
}

// For a PDF the parser can't read (e.g. another carrier's label): a blank
// record for manual entry that still carries page 1's preview, so the
// label stays visible on the left (and can be enlarged) while the user
// types the details in.
async function buildManualRecord(file) {
  const doc = await loadLabelDoc(file);
  const page = await doc.getPage(1);
  let previewDataUrl = null;
  try {
    previewDataUrl = await renderLabelPreview(page, 340, { trim: true });
  } catch (err) {
    console.error("Label preview render failed:", err);
  }
  return {
    isManual: true,
    previewTrim: true,
    carrier: "",
    sourceFile: file.name,
    previewDataUrl,
    previewPage: previewDataUrl ? page : null,
    trackingNumber: null, shipDate: null, service: null, serviceConfidence: "labeled",
    destCountry: null, isInternational: null,
    weight: null, weightUnit: null, totalWeight: null, totalWeightUnit: null, dimensions: null,
    isMultiPiece: false, pieceCount: 1, pieces: [],
    reference: null, invoice: null, po: null, dept: null,
    sender: { name: null, address: null, phone: null },
    recipient: { name: null, address: null, phone: null },
    parseStatus: "ok",
    warnings: [],
  };
}
