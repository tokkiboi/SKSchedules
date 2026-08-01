/**
 * GmailPipeline.gs — StyleKorean Logistics email ingestion
 *
 * Time-driven scan of Gmail for logistics documents (출고 / 해상 / 항공,
 * arrival notices, bills of lading, entry summaries, shipping docs),
 * archives attachments to organized Drive folders, parses CSV/XLSX/PDF
 * payloads, and appends validated rows to the Inbound (IMPORTS) and
 * Outbound (WH Trucking Request) schedules in LOGISTICS MASTER 2026.
 *
 * Anything ambiguous is routed to the PENDING VERIFICATION sheet
 * (see Validation.gs) instead of being auto-committed.
 *
 * Requirements:
 *  - Advanced Drive Service enabled (Services > Drive API, identifier "Drive").
 *  - Runs as the mailbox owner (alex@stylekoreanus.com).
 *  - Trigger provisioning: see Triggers.gs setupAllTriggers().
 */

/* eslint-disable no-unused-vars */

var GMAIL_PIPELINE = {
  masterId: "1M-vZ24Yw4ZN7R7b_473cVn8kny8DznTakSsD3VQsCzc",
  inboundSheet: "IMPORTS",
  outboundSheet: "WH Trucking Request",
  driveRootName: "SK Logistics Email Archive",
  labels: {
    processed: "sk-logistics/processed",
    pending: "sk-logistics/pending-verification",
    error: "sk-logistics/error"
  },
  // Subject / body keywords. Korean: 출고 (outbound/shipped), 해상 (ocean), 항공 (air), 선적 (loading/shipment), 입고 (receiving).
  query:
    'has:attachment newer_than:7d -label:sk-logistics/processed -label:sk-logistics/error ' +
    '{subject:출고 subject:해상 subject:항공 subject:선적 subject:입고 ' +
    'subject:"arrival notice" subject:"bill of lading" subject:BOL subject:"entry summary" ' +
    'subject:"shipping documents" subject:ISF subject:"delivery order" subject:POD}',
  maxThreadsPerRun: 20,
  categories: [
    { key: "ARRIVAL_NOTICE", folder: "Arrival Notices", match: /arrival\s*notice|도착\s*통지|A\/N/i },
    { key: "BOL", folder: "Bills of Lading", match: /bill\s*of\s*lading|\bB\/?L\b|\bBOL\b|선하증권/i },
    { key: "ENTRY_SUMMARY", folder: "Entry Summaries", match: /entry\s*summary|7501|customs\s*entry|통관/i },
    { key: "WMS_EXPORT", folder: "WMS Exports", match: /wms|invoice.*issues|출고.*(list|리스트|명세)/i },
    { key: "SHIPPING_DOCS", folder: "Shipping Documents", match: /shipping\s*doc|packing\s*list|commercial\s*invoice|선적서류/i },
    { key: "OTHER", folder: "Other", match: /.*/ }
  ]
};

/** Entry point — run from a 15-minute time-driven trigger. */
function processLogisticsEmails() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    ensureLabels_();
    ensurePendingSheet_(); // Validation.gs
    var threads = GmailApp.search(GMAIL_PIPELINE.query, 0, GMAIL_PIPELINE.maxThreadsPerRun);
    var summary = { threads: threads.length, committed: 0, pending: 0, archived: 0, errors: 0 };

    threads.forEach(function (thread) {
      try {
        var result = processThread_(thread);
        summary.committed += result.committed;
        summary.pending += result.pending;
        summary.archived += result.archived;
        thread.addLabel(getLabel_(GMAIL_PIPELINE.labels.processed));
        if (result.pending > 0) thread.addLabel(getLabel_(GMAIL_PIPELINE.labels.pending));
      } catch (err) {
        summary.errors++;
        thread.addLabel(getLabel_(GMAIL_PIPELINE.labels.error));
        logPipeline_("THREAD ERROR", thread.getFirstMessageSubject(), String(err && err.message || err));
      }
    });

    logPipeline_("RUN COMPLETE", JSON.stringify(summary), "");
    return summary;
  } finally {
    lock.releaseLock();
  }
}

/** Processes every message and attachment in one thread. */
function processThread_(thread) {
  var out = { committed: 0, pending: 0, archived: 0 };
  thread.getMessages().forEach(function (message) {
    var meta = {
      subject: message.getSubject() || "",
      sender: message.getFrom() || "",
      date: message.getDate(),
      messageId: message.getId(),
      permalink: "https://mail.google.com/mail/u/0/#all/" + message.getId()
    };
    var category = classifyText_(meta.subject);

    message.getAttachments({ includeInlineImages: false }).forEach(function (attachment) {
      var name = attachment.getName() || "attachment";
      if (/\.(png|jpe?g|gif|ics|vcf)$/i.test(name)) return; // skip signatures/calendar noise
      var fileCategory = classifyText_(name) !== "OTHER" ? classifyText_(name) : category;

      // 1) Archive raw file to Drive: <root>/<YYYY>/<MM>/<Category>/
      var folder = archiveFolderFor_(meta.date, fileCategory);
      var file = folder.createFile(attachment.copyBlob().setName(stampName_(name, meta.date)));
      out.archived++;

      // 2) Extract candidate records
      var records = [];
      try {
        records = extractRecords_(attachment, name, fileCategory, meta);
      } catch (err) {
        records = [];
        addPendingRow_({
          kind: guessKind_(fileCategory, meta.subject),
          record: { parseError: String(err && err.message || err) },
          issues: ["Attachment could not be parsed automatically."],
          meta: meta, driveUrl: file.getUrl()
        });
        out.pending++;
      }

      // 3) Validate and commit or park each record
      records.forEach(function (record) {
        record._sourceEmail = meta.permalink;
        record._driveFile = file.getUrl();
        var kind = record.kind || guessKind_(fileCategory, meta.subject);
        var verdict = validateRecord_(record, kind); // Validation.gs
        if (verdict.ok) {
          var committed = kind === "inbound"
            ? upsertInboundRow_(record)
            : upsertOutboundRow_(record);
          if (committed) { out.committed++; }
          else {
            addPendingRow_({ kind: kind, record: record, issues: ["Possible duplicate — matching row already exists."], meta: meta, driveUrl: file.getUrl() });
            out.pending++;
          }
        } else {
          addPendingRow_({ kind: kind, record: record, issues: verdict.issues, meta: meta, driveUrl: file.getUrl() });
          out.pending++;
        }
      });
    });
  });
  return out;
}

/* ------------------------------------------------------------------ */
/* Extraction                                                          */
/* ------------------------------------------------------------------ */

function extractRecords_(attachment, name, category, meta) {
  var lower = name.toLowerCase();
  if (/\.csv$|\.tsv$/.test(lower)) {
    return tabularToRecords_(Utilities.parseCsv(attachment.getDataAsString("UTF-8"), /\.tsv$/.test(lower) ? "\t" : ","), category, meta);
  }
  if (/\.xlsx?$|\.xlsm$/.test(lower)) {
    return tabularToRecords_(xlsxToRows_(attachment), category, meta);
  }
  if (/\.pdf$/.test(lower)) {
    var text = pdfToText_(attachment);
    var record = pdfTextToRecord_(text, category, meta);
    return record ? [record] : [];
  }
  return [];
}

/** Converts an XLSX blob to a 2D array via a temporary Google Sheet. */
function xlsxToRows_(attachment) {
  var resource = { title: "TMP-import-" + Date.now(), mimeType: MimeType.GOOGLE_SHEETS };
  var tmp = Drive.Files.insert(resource, attachment.copyBlob(), { convert: true });
  try {
    var sheet = SpreadsheetApp.openById(tmp.id).getSheets()[0];
    return sheet.getDataRange().getDisplayValues();
  } finally {
    Drive.Files.trash(tmp.id);
  }
}

/** Extracts text from a PDF using Drive OCR (best effort). */
function pdfToText_(attachment) {
  var resource = { title: "TMP-ocr-" + Date.now(), mimeType: MimeType.GOOGLE_DOCS };
  var tmp = Drive.Files.insert(resource, attachment.copyBlob(), { convert: true, ocr: true, ocrLanguage: "en" });
  try {
    return DocumentApp.openById(tmp.id).getBody().getText() || "";
  } finally {
    Drive.Files.trash(tmp.id);
  }
}

/**
 * Maps a parsed spreadsheet (header row + data rows) into normalized records
 * using a column-alias dictionary (English + Korean headers).
 */
function tabularToRecords_(rows, category, meta) {
  if (!rows || rows.length < 2) return [];
  var aliases = {
    customer: ["CUSTOMER", "CLIENT", "ACCOUNT", "SHIP TO", "거래처", "고객사"],
    invoice: ["INVOICE", "INVOICE NO.", "INVOICE #", "PO#", "PO NUMBER", "인보이스", "PI NO.", "PI NO"],
    shipDate: ["SHIP DATE", "PU DATE", "DATE", "ETD", "출고일", "선적일"],
    eta: ["ETA", "ARRIVAL", "ARRIVAL DATE", "입고일", "도착일", "도착예정일"],
    qty: ["QTY", "QUANTITY", "수량", "CARTONS", "CTN"],
    pallets: ["PALLET", "PALLETS", "PLT", "팔렛", "파렛트"],
    carrier: ["CARRIER", "TRUCKING", "운송사", "선사", "FORWARDER"],
    pro: ["PRO#", "PRO", "TRACKING", "BOL", "BOL#", "B/L", "BL NO", "B/L NO.", "HBL", "MBL"],
    container: ["CONTAINER", "CONTAINER NO", "CNTR", "컨테이너"],
    vessel: ["VESSEL", "VESSEL/VOY", "FLIGHT", "모선", "선명"],
    sku: ["SKU", "상품코드", "ITEM CODE", "품번"],
    note: ["NOTE", "REMARK", "MEMO", "비고", "ISSUE"]
  };

  // Find header row within the first 5 rows.
  var headerRowIdx = -1, map = {};
  for (var r = 0; r < Math.min(5, rows.length); r++) {
    var candidate = headerAliasMap_(rows[r], aliases);
    if (Object.keys(candidate).length >= 2) { headerRowIdx = r; map = candidate; break; }
  }
  if (headerRowIdx === -1) throw new Error("No recognizable header row.");

  var records = [];
  for (var i = headerRowIdx + 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row.join("").trim()) continue;
    var rec = { kind: guessKind_(category, meta.subject) };
    Object.keys(map).forEach(function (field) { rec[field] = String(row[map[field]] || "").trim(); });
    if (rec.customer || rec.invoice || rec.pro || rec.container) records.push(rec);
  }
  return records;
}

function headerAliasMap_(headerRow, aliases) {
  var map = {};
  headerRow.forEach(function (cell, idx) {
    var value = String(cell || "").trim().toUpperCase();
    if (!value) return;
    Object.keys(aliases).forEach(function (field) {
      if (map[field] !== undefined) return;
      if (aliases[field].some(function (alias) { return value === alias || value.indexOf(alias) === 0; })) map[field] = idx;
    });
  });
  return map;
}

/** Pulls key fields out of OCR'd PDF text (arrival notices, BOLs, entry summaries). */
function pdfTextToRecord_(text, category, meta) {
  if (!text || text.length < 40) return null;
  var grab = function (patterns) {
    for (var i = 0; i < patterns.length; i++) {
      var m = text.match(patterns[i]);
      if (m && m[1]) return m[1].trim();
    }
    return "";
  };
  var record = {
    kind: "inbound",
    pro: grab([/(?:B\/?L|BOL|BILL OF LADING)[\s#:.NO]*([A-Z0-9-]{6,20})/i, /(?:HBL|MBL)[\s#:.NO]*([A-Z0-9-]{6,20})/i]),
    container: grab([/(?:CONTAINER|CNTR)[\s#:.NO]*([A-Z]{4}\d{7})/i, /\b([A-Z]{4}\d{7})\b/]),
    vessel: grab([/(?:VESSEL|VSL)[\s:/]*([A-Z0-9 .-]{3,30})(?:\n|VOY|$)/i, /(?:FLIGHT|FLT)[\s#:.NO]*([A-Z0-9 -]{3,12})/i]),
    eta: grab([/(?:ETA|ARRIVAL DATE|DATE OF ARRIVAL)[\s:]*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i]),
    invoice: grab([/(?:ENTRY (?:NO|NUMBER)|7501)[\s#:.]*([A-Z0-9-]{8,15})/i, /(?:INVOICE|PI)[\s#:.NO]*([A-Z0-9-]{5,20})/i]),
    qty: grab([/([0-9,]{1,9})\s*(?:CTNS?|CARTONS?|PKGS?|PACKAGES?)/i]),
    mode: /항공|AIR|FLIGHT|AWB/i.test(text) ? "AIR" : (/해상|OCEAN|VESSEL|SEA/i.test(text) ? "OCEAN" : ""),
    note: "Auto-extracted from " + category + " · " + meta.subject
  };
  return (record.pro || record.container || record.invoice) ? record : { kind: "inbound", note: record.note, _rawTextSample: text.slice(0, 400) };
}

/* ------------------------------------------------------------------ */
/* Sheet upserts (dedupe-aware)                                        */
/* ------------------------------------------------------------------ */

/** Appends an inbound record to IMPORTS unless a matching row already exists. Returns true if committed. */
function upsertInboundRow_(record) {
  var sheet = SpreadsheetApp.openById(GMAIL_PIPELINE.masterId).getSheetByName(GMAIL_PIPELINE.inboundSheet);
  if (!sheet) throw new Error("IMPORTS sheet not found.");
  var data = sheet.getDataRange().getDisplayValues();
  var headers = findHeaderRowIdx_(data);
  var map = headerMap_(data[headers]);

  var keyCols = ["B/L", "BL NO", "B/L NO.", "HBL", "MBL", "BOL", "CONTAINER", "CONTAINER NO", "CNTR"];
  var keys = [record.pro, record.container].filter(Boolean).map(function (v) { return v.toUpperCase(); });
  if (keys.length) {
    for (var r = headers + 1; r < data.length; r++) {
      var hit = keyCols.some(function (col) {
        return map[col] !== undefined && keys.indexOf(String(data[r][map[col]] || "").trim().toUpperCase()) !== -1;
      });
      if (hit) {
        // Update ETA if we have a newer one; do not duplicate the row.
        if (record.eta && map["ETA"] !== undefined) sheet.getRange(r + 1, map["ETA"] + 1).setValue(record.eta);
        return false;
      }
    }
  }

  var width = data[headers].length;
  var newRow = new Array(width).fill("");
  var put = function (names, value) {
    if (!value) return;
    for (var i = 0; i < names.length; i++) {
      if (map[names[i]] !== undefined) { newRow[map[names[i]]] = value; return; }
    }
  };
  put(["B/L", "BL NO", "B/L NO.", "BOL", "HBL"], record.pro);
  put(["CONTAINER", "CONTAINER NO", "CNTR"], record.container);
  put(["VESSEL", "VESSEL/VOY", "모선"], record.vessel);
  put(["ETA", "ARRIVAL", "도착예정일"], record.eta);
  put(["INVOICE", "INVOICE NO.", "PI NO.", "ENTRY NO"], record.invoice);
  put(["QTY", "CTNS", "CARTONS", "수량"], record.qty);
  put(["MODE", "TYPE"], record.mode);
  put(["NOTE", "REMARK", "비고"], (record.note || "") + " [auto: " + (record._sourceEmail || "email") + "]");
  put(["WEBSITE STATUS", "STATUS"], "SCHEDULED");
  sheet.appendRow(newRow);
  markAutoRow_(sheet, sheet.getLastRow());
  return true;
}

/** Appends an outbound record to WH Trucking Request unless invoice/customer+date already exist. */
function upsertOutboundRow_(record) {
  var sheet = SpreadsheetApp.openById(GMAIL_PIPELINE.masterId).getSheetByName(GMAIL_PIPELINE.outboundSheet);
  if (!sheet) throw new Error("WH Trucking Request sheet not found.");
  var data = sheet.getDataRange().getDisplayValues();
  var headers = findHeaderRowIdx_(data);
  var map = headerMap_(data[headers]);

  var invoice = String(record.invoice || "").toUpperCase();
  var custKey = (String(record.customer || "").toUpperCase().replace(/\s+/g, " ") + "___" + String(record.shipDate || "").toUpperCase());
  for (var r = headers + 1; r < data.length; r++) {
    var rowInvoices = String(map["INVOICE NO."] !== undefined ? data[r][map["INVOICE NO."]] : (map["INVOICE"] !== undefined ? data[r][map["INVOICE"]] : "")).toUpperCase();
    if (invoice && rowInvoices.split(/[\r\n,;·]+/).map(function (s) { return s.trim(); }).indexOf(invoice) !== -1) return false;
    var rowKey = String(map["CUSTOMER"] !== undefined ? data[r][map["CUSTOMER"]] : "").toUpperCase().replace(/\s+/g, " ") + "___" +
      String(map["SHIP DATE"] !== undefined ? data[r][map["SHIP DATE"]] : "").toUpperCase();
    if (record.customer && record.shipDate && rowKey === custKey) return false;
  }

  var width = data[headers].length;
  var newRow = new Array(width).fill("");
  var put = function (names, value) {
    if (!value) return;
    for (var i = 0; i < names.length; i++) {
      if (map[names[i]] !== undefined) { newRow[map[names[i]]] = value; return; }
    }
  };
  put(["CUSTOMER"], record.customer);
  put(["INVOICE NO.", "INVOICE #", "INVOICE"], record.invoice);
  put(["SHIP DATE"], record.shipDate);
  put(["PALLET TYPE", "PALLETS", "PLT"], record.pallets);
  put(["CARRIER"], record.carrier || "Trucking");
  put(["PRO#", "PRO"], record.pro);
  put(["NOTE", "REMARK"], (record.note || "Imported from email") + " [auto: " + (record._sourceEmail || "email") + "]");
  put(["WEBSITE STATUS", "STATUS"], "WORK IN PROGRESS");
  sheet.appendRow(newRow);
  markAutoRow_(sheet, sheet.getLastRow());
  return true;
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function classifyText_(text) {
  var categories = GMAIL_PIPELINE.categories;
  for (var i = 0; i < categories.length; i++) {
    if (categories[i].match.test(text || "")) return categories[i].key;
  }
  return "OTHER";
}

function guessKind_(category, subject) {
  if (category === "ARRIVAL_NOTICE" || category === "ENTRY_SUMMARY" || category === "BOL") return "inbound";
  if (/해상|항공|입고|arrival|customs|entry|vessel|container/i.test(subject || "")) return "inbound";
  return "outbound"; // 출고 / WMS exports / trucking docs
}

function categoryFolderName_(key) {
  var categories = GMAIL_PIPELINE.categories;
  for (var i = 0; i < categories.length; i++) if (categories[i].key === key) return categories[i].folder;
  return "Other";
}

function archiveFolderFor_(date, categoryKey) {
  var tz = Session.getScriptTimeZone();
  var root = getOrCreateFolder_(DriveApp.getRootFolder(), GMAIL_PIPELINE.driveRootName);
  var year = getOrCreateFolder_(root, Utilities.formatDate(date, tz, "yyyy"));
  var month = getOrCreateFolder_(year, Utilities.formatDate(date, tz, "MM"));
  return getOrCreateFolder_(month, categoryFolderName_(categoryKey));
}

function getOrCreateFolder_(parent, name) {
  var existing = parent.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parent.createFolder(name);
}

function stampName_(name, date) {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(date, tz, "yyyyMMdd-HHmm") + " " + name;
}

function ensureLabels_() {
  Object.keys(GMAIL_PIPELINE.labels).forEach(function (key) { getLabel_(GMAIL_PIPELINE.labels[key]); });
}

function getLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function findHeaderRowIdx_(data) {
  for (var r = 0; r < Math.min(4, data.length); r++) {
    var upper = data[r].map(function (c) { return String(c || "").trim().toUpperCase(); });
    if (upper.indexOf("CUSTOMER") !== -1 || upper.indexOf("STATUS") !== -1 || upper.indexOf("ETA") !== -1 ||
        upper.indexOf("B/L") !== -1 || upper.indexOf("INVOICE NO.") !== -1) return r;
  }
  return 0;
}

/** Light-blue tint marks rows committed by automation, for at-a-glance review. */
function markAutoRow_(sheet, rowIdx) {
  sheet.getRange(rowIdx, 1, 1, Math.max(sheet.getLastColumn(), 1)).setBackground("#E8F0FE");
}

function logPipeline_(event, subject, detail) {
  try {
    var ss = SpreadsheetApp.openById(GMAIL_PIPELINE.masterId);
    var log = ss.getSheetByName("PIPELINE LOG") || ss.insertSheet("PIPELINE LOG");
    if (log.getLastRow() === 0) log.appendRow(["Timestamp", "Event", "Subject", "Detail"]);
    log.appendRow([new Date(), event, subject, detail]);
    if (log.getLastRow() > 2000) log.deleteRows(2, 500); // keep the log bounded
  } catch (e) { /* logging must never break the pipeline */ }
}
