/**
 * Validation.gs — data validation & manual verification workflow
 *
 * validateRecord_() scores every record extracted from email before it may
 * touch the live schedules. Records that fail are written to the
 * PENDING VERIFICATION sheet (yellow rows) for a human to approve.
 * processApprovedPending() (time-driven) commits rows marked APPROVED.
 */

/* eslint-disable no-unused-vars */

var VALIDATION = {
  pendingSheetName: "PENDING VERIFICATION",
  pendingHeaders: ["Timestamp", "Kind", "Status", "Issues", "Customer", "Invoice / PI", "BL / PRO", "Container", "Ship Date / ETA", "Qty", "Carrier / Vessel", "Note", "Source Email", "Drive File", "Raw JSON"],
  statusValues: ["NEEDS REVIEW", "APPROVED", "REJECTED", "COMMITTED"],
  colors: { needsReview: "#FFF3CD", approved: "#D9EAD3", rejected: "#F4CCCC", committed: "#E8F0FE" },
  dateWindowPastDays: 45,     // reject dates further back than this
  dateWindowFutureDays: 400   // reject dates further out than this
};

/**
 * Returns { ok: boolean, issues: string[] }.
 * A record is committable only when required identifiers exist,
 * dates parse inside a sane window, and quantities are numeric.
 */
function validateRecord_(record, kind) {
  var issues = [];

  if (kind === "inbound") {
    if (!record.pro && !record.container && !record.invoice) {
      issues.push("No B/L, container, or invoice/entry number found.");
    }
    if (!record.eta && !record.shipDate) issues.push("No ETA or ship date found.");
    if (record.eta && !isSaneDate_(record.eta)) issues.push("ETA does not parse or is outside the expected window: " + record.eta);
    if (record.container && !/^[A-Z]{4}\d{7}$/.test(String(record.container).replace(/\s/g, ""))) {
      issues.push("Container number is not ISO-format (AAAA9999999): " + record.container);
    }
  } else {
    if (!record.customer) issues.push("Customer is missing.");
    if (!record.invoice && !record.pro) issues.push("Neither invoice/PO nor PRO/BOL found.");
    if (!record.shipDate) issues.push("Ship date is missing.");
    if (record.shipDate && !isSaneDate_(record.shipDate)) issues.push("Ship date does not parse or is outside the expected window: " + record.shipDate);
  }

  if (record.qty && !/^[\d,.\s]+$/.test(String(record.qty))) issues.push("Quantity is not numeric: " + record.qty);
  if (record.parseError) issues.push(record.parseError);
  if (record._rawTextSample) issues.push("PDF parsed but no reliable identifiers were found.");

  return { ok: issues.length === 0, issues: issues };
}

function isSaneDate_(value) {
  var parsed = parseFlexibleDate_(value);
  if (!parsed) return false;
  var now = new Date();
  var past = new Date(now.getTime() - VALIDATION.dateWindowPastDays * 86400000);
  var future = new Date(now.getTime() + VALIDATION.dateWindowFutureDays * 86400000);
  return parsed >= past && parsed <= future;
}

/** Accepts M/D, M/D/YY, M/D/YYYY, YYYY-MM-DD, "Aug 3", 2026.08.03 etc. */
function parseFlexibleDate_(value) {
  var s = String(value || "").trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/.-](\d{1,2})(?:[\/.-](\d{2,4}))?$/);
  if (m) {
    var year = m[3] ? Number(m[3].length === 2 ? "20" + m[3] : m[3]) : new Date().getFullYear();
    var d = new Date(year, Number(m[1]) - 1, Number(m[2]));
    return isNaN(d.getTime()) ? null : d;
  }
  m = s.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (m) {
    var d2 = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(d2.getTime()) ? null : d2;
  }
  var d3 = new Date(s);
  return isNaN(d3.getTime()) ? null : d3;
}

/* ------------------------------------------------------------------ */
/* Pending Verification sheet                                          */
/* ------------------------------------------------------------------ */

function ensurePendingSheet_() {
  var ss = SpreadsheetApp.openById(GMAIL_PIPELINE.masterId);
  var sheet = ss.getSheetByName(VALIDATION.pendingSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(VALIDATION.pendingSheetName);
    sheet.appendRow(VALIDATION.pendingHeaders);
    sheet.getRange(1, 1, 1, VALIDATION.pendingHeaders.length).setFontWeight("bold").setBackground("#EFEFEF");
    sheet.setFrozenRows(1);
    var statusCol = VALIDATION.pendingHeaders.indexOf("Status") + 1;
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(VALIDATION.statusValues, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, 1000, 1).setDataValidation(rule);
  }
  return sheet;
}

/** Parks a questionable record for manual review (yellow row). */
function addPendingRow_(entry) {
  var sheet = ensurePendingSheet_();
  var r = entry.record || {};
  sheet.appendRow([
    new Date(),
    entry.kind || "",
    "NEEDS REVIEW",
    (entry.issues || []).join(" | "),
    r.customer || "",
    r.invoice || "",
    r.pro || "",
    r.container || "",
    r.shipDate || r.eta || "",
    r.qty || "",
    r.carrier || r.vessel || "",
    r.note || "",
    (entry.meta && entry.meta.permalink) || r._sourceEmail || "",
    entry.driveUrl || r._driveFile || "",
    JSON.stringify(r).slice(0, 5000)
  ]);
  sheet.getRange(sheet.getLastRow(), 1, 1, VALIDATION.pendingHeaders.length)
    .setBackground(VALIDATION.colors.needsReview);
}

/**
 * Time-driven: commits every pending row whose Status was manually set to
 * APPROVED, then re-colors it. REJECTED rows are greyed out and left in place
 * as an audit trail.
 */
function processApprovedPending() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var sheet = ensurePendingSheet_();
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return { committed: 0 };
    var col = {};
    VALIDATION.pendingHeaders.forEach(function (h, i) { col[h] = i; });

    var committed = 0;
    for (var r = 1; r < data.length; r++) {
      var status = String(data[r][col["Status"]] || "").trim().toUpperCase();
      var rowRange = sheet.getRange(r + 1, 1, 1, VALIDATION.pendingHeaders.length);
      if (status === "APPROVED") {
        var record;
        try { record = JSON.parse(data[r][col["Raw JSON"]] || "{}"); }
        catch (e) { record = {}; }
        // Prefer manually corrected cell values over the original extraction.
        record.customer = data[r][col["Customer"]] || record.customer;
        record.invoice = data[r][col["Invoice / PI"]] || record.invoice;
        record.pro = data[r][col["BL / PRO"]] || record.pro;
        record.container = data[r][col["Container"]] || record.container;
        record.qty = data[r][col["Qty"]] || record.qty;
        record.note = data[r][col["Note"]] || record.note;
        var when = data[r][col["Ship Date / ETA"]];
        var kind = String(data[r][col["Kind"]] || "outbound").toLowerCase();
        if (kind === "inbound") { record.eta = when || record.eta; upsertInboundRow_(record); }
        else { record.shipDate = when || record.shipDate; upsertOutboundRow_(record); }
        sheet.getRange(r + 1, col["Status"] + 1).setValue("COMMITTED");
        rowRange.setBackground(VALIDATION.colors.committed);
        committed++;
      } else if (status === "REJECTED") {
        rowRange.setBackground(VALIDATION.colors.rejected).setFontColor("#999999");
      }
    }
    return { committed: committed };
  } finally {
    lock.releaseLock();
  }
}

/** Count of rows still needing review — consumed by the KPI dashboard. */
function pendingVerificationCount_() {
  try {
    var sheet = SpreadsheetApp.openById(GMAIL_PIPELINE.masterId).getSheetByName(VALIDATION.pendingSheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    var statusCol = VALIDATION.pendingHeaders.indexOf("Status") + 1;
    return sheet.getRange(2, statusCol, sheet.getLastRow() - 1, 1).getDisplayValues()
      .filter(function (row) { return String(row[0]).trim().toUpperCase() === "NEEDS REVIEW"; }).length;
  } catch (e) { return 0; }
}
