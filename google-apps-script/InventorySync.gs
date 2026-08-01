/**
 * InventorySync.gs — live Inventory Management module + KPI dashboard
 *
 * Relationally links two raw inventory workbooks to LOGISTICS MASTER 2026:
 *
 *  A. ALLOCATION workbook (17e5EYNMr3sTPhMfFVYBg3dQDC7CWxk2StwHzt_W55dY)
 *     Per-shipment SKU allocation tabs:
 *     SKU · Product Name · Brand · Barcode · Cnfm Qty · 잔여수량 (remaining)
 *     · channel columns (CAWH / iHerb / National / BK / US_Official / Moida / NY)
 *
 *  B. WMS workbook (1tNBa7c78MGL3wBNwYsDdHcHnJZFcvxrLzn_M79vN4WY)
 *     - Live stock snapshot: SKU · Location · Expiry · Total/Actual/Hold/Avail Qty
 *     - Container receiving log: Type · 차수 · PC · 입고일 · 검수 완료일
 *     - Per-container putaway tabs keyed by PI NO. / PLT NO.
 *
 * Join model:
 *   SKU (상품코드)  → product-level join across A, B, and outbound demand
 *   차수 / PC no.  → container-level join between B's receiving log and
 *                    the IMPORTS inbound schedule in LOGISTICS MASTER
 *
 * Outputs (written into LOGISTICS MASTER 2026, read live by the website):
 *   - "INVENTORY" tab      : one row per SKU (on hand + incoming + allocation)
 *   - "KPI DASHBOARD" tab  : label/value metric block, updated each run
 */

/* eslint-disable no-unused-vars */

var INVENTORY_SYNC = {
  masterId: "1M-vZ24Yw4ZN7R7b_473cVn8kny8DznTakSsD3VQsCzc",
  allocationId: "17e5EYNMr3sTPhMfFVYBg3dQDC7CWxk2StwHzt_W55dY",
  wmsId: "1tNBa7c78MGL3wBNwYsDdHcHnJZFcvxrLzn_M79vN4WY",
  inventoryTab: "INVENTORY",
  kpiTab: "KPI DASHBOARD",
  importsTab: "IMPORTS",
  runBudgetMs: 4.5 * 60 * 1000, // stay under the 6-minute Apps Script limit
  lowStockThreshold: 50
};

/** Entry point — run hourly from a time-driven trigger. */
function syncInventoryModule() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  var startedAt = Date.now();
  try {
    var incoming = readAllocationIncoming_(startedAt);   // SKU -> incoming/allocated (heaviest read goes first)
    var stock = readWmsStockSnapshot_();                 // SKU -> on-hand
    var containers = readWmsContainerLog_();             // 차수 -> receiving status
    writeInventoryTab_(stock, incoming, containers);
    updateKpiDashboard_(stock, incoming, containers);
    logPipeline_("INVENTORY SYNC", "ok",
      Object.keys(stock.bySku).length + " stocked SKUs · " +
      Object.keys(incoming.bySku).length + " inbound SKUs · " +
      containers.rows.length + " containers · " + (incoming.partial ? "PARTIAL (time budget)" : "full"));
  } catch (err) {
    logPipeline_("INVENTORY SYNC ERROR", "", String(err && err.message || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Readers                                                             */
/* ------------------------------------------------------------------ */

/** Finds the live stock tab in the WMS workbook by its header signature. */
function readWmsStockSnapshot_() {
  var ss = SpreadsheetApp.openById(INVENTORY_SYNC.wmsId);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getLastRow() < 2) continue;
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
      .map(function (c) { return String(c || "").trim().toUpperCase(); });
    if (header.indexOf("SKU") !== -1 && header.indexOf("AVAIL QTY") !== -1) {
      var col = {};
      header.forEach(function (h, idx) { col[h] = idx; });
      var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
      var bySku = {};
      rows.forEach(function (row) {
        var sku = String(row[col["SKU"]] || "").trim();
        if (!sku) return;
        var entry = bySku[sku] || (bySku[sku] = {
          name: "", brand: "", barcode: "", total: 0, actual: 0, hold: 0, avail: 0, locations: [], nearestExpiry: ""
        });
        entry.name = entry.name || String(row[col["PRODUCT NAME"]] || "");
        entry.brand = entry.brand || String(row[col["BRAND"]] || "");
        entry.barcode = entry.barcode || String(row[col["PRODUCT BARCODE"]] || "");
        entry.total += num_(row[col["TOTAL QTY"]]);
        entry.actual += num_(row[col["ACTUAL QTY"]]);
        entry.hold += num_(row[col["HOLD(PICKED)"]]) + num_(row[col["HOLD(REQ)"]]);
        entry.avail += num_(row[col["AVAIL QTY"]]);
        var loc = String(row[col["LOCATION"]] || "").trim();
        if (loc && entry.locations.indexOf(loc) === -1) entry.locations.push(loc);
        var expiry = String(row[col["EXPIRY DATE"]] || "").trim();
        if (expiry && (!entry.nearestExpiry || expiry < entry.nearestExpiry)) entry.nearestExpiry = expiry;
      });
      return { bySku: bySku, sheetName: sheet.getName() };
    }
  }
  throw new Error("WMS stock snapshot tab (SKU + Avail Qty headers) not found.");
}

/**
 * Reads the container receiving log (Type · 차수 · PC · 입고일 · 검수 완료일).
 * The header may not be on row 1 and column labels vary slightly, so this
 * scans the first 3 rows of the first 15 tabs for a row containing 차수.
 */
function readWmsContainerLog_() {
  var ss = SpreadsheetApp.openById(INVENTORY_SYNC.wmsId);
  var sheets = ss.getSheets();
  for (var i = 0; i < Math.min(sheets.length, 15); i++) {
    var sheet = sheets[i];
    if (sheet.getLastRow() < 2) continue;
    var scan = sheet.getRange(1, 1, Math.min(3, sheet.getLastRow()), sheet.getLastColumn()).getDisplayValues();
    for (var h = 0; h < scan.length; h++) {
      var header = scan[h].map(function (c) { return String(c || "").trim(); });
      if (header.indexOf("차수") === -1) continue;
      var col = {};
      header.forEach(function (name, idx) { if (name && col[name] === undefined) col[name] = idx; });
      var pick = function (row, names) {
        for (var n = 0; n < names.length; n++) {
          if (col[names[n]] !== undefined) return String(row[col[names[n]]] || "").trim();
        }
        return "";
      };
      var startRow = h + 2;
      if (sheet.getLastRow() < startRow) break;
      var rows = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, sheet.getLastColumn()).getDisplayValues()
        .filter(function (row) { return String(row[col["차수"]] || "").trim(); })
        .map(function (row) {
          return {
            type: pick(row, ["Type", "TYPE"]) || String(row[0] || "").trim(),
            shipmentCode: String(row[col["차수"]] || "").trim(),      // e.g. "TW 12", "HJ 31"
            pcNumber: pick(row, ["PC", "PC#", "PC NO", "PC NO."]),    // e.g. "PC00146273"
            receivedDate: pick(row, ["입고일", "입고 일", "입고일자"]),
            qcDoneDate: pick(row, ["검수 완료일", "검수완료일", "검수 완료"]),
            remark: pick(row, ["WHS REMARK", "REMARK", "비고"])
          };
        });
      var byCode = {};
      rows.forEach(function (r) { byCode[r.shipmentCode.toUpperCase()] = r; });
      return { rows: rows, byCode: byCode, sheetName: sheet.getName() };
    }
  }
  return { rows: [], byCode: {}, sheetName: "" };
}

/**
 * Aggregates the allocation workbook: per SKU, confirmed inbound quantity,
 * remaining-to-receive (잔여수량), and channel allocation summary.
 * Tab names are treated as shipment identifiers.
 * Respects the script run budget; sets partial=true when it had to stop early.
 */
function readAllocationIncoming_(startedAt) {
  var CHANNELS = ["CAWH", "IHERB", "HQ IHERB PO", "NATIONAL", "BK", "US_OFFICIAL", "MOIDA", "NY"];
  var ss = SpreadsheetApp.openById(INVENTORY_SYNC.allocationId);
  var sheets = ss.getSheets();
  var bySku = {};
  var partial = false;

  var allocBudgetMs = INVENTORY_SYNC.runBudgetMs - 60 * 1000; // reserve a minute for WMS reads + writes
  for (var s = 0; s < sheets.length; s++) {
    if (Date.now() - startedAt > allocBudgetMs) { partial = true; break; }
    var sheet = sheets[s];
    if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 4) continue;
    if (sheet.getLastColumn() > 40) continue; // skip the wide per-shipment tracker tab

    var data = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 300), Math.min(sheet.getLastColumn(), 16)).getDisplayValues();
    // Header can be on row 1 or 2; require SKU + Cnfm Qty to treat the tab as an allocation sheet.
    var headerIdx = -1, col = {};
    for (var r = 0; r < Math.min(3, data.length); r++) {
      var upper = data[r].map(function (c) { return String(c || "").trim().toUpperCase(); });
      if (upper.indexOf("SKU") !== -1 && (upper.indexOf("CNFM QTY") !== -1 || upper.indexOf("잔여수량") !== -1)) {
        headerIdx = r;
        upper.forEach(function (h, idx) { if (h && col[h] === undefined) col[h] = idx; });
        break;
      }
    }
    if (headerIdx === -1) continue;

    var shipment = sheet.getName();
    for (var i = headerIdx + 1; i < data.length; i++) {
      var row = data[i];
      var sku = String(row[col["SKU"]] || "").trim();
      if (!sku) continue;
      var entry = bySku[sku] || (bySku[sku] = {
        name: "", brand: "", barcode: "", confirmed: 0, remaining: 0, shipments: [], channels: {}
      });
      entry.name = entry.name || String(col["PRODUCT NAME"] !== undefined ? row[col["PRODUCT NAME"]] : "");
      entry.brand = entry.brand || String(col["BRAND"] !== undefined ? row[col["BRAND"]] : "");
      entry.barcode = entry.barcode || String(col["BARCODE"] !== undefined ? row[col["BARCODE"]] : "");
      entry.confirmed += num_(col["CNFM QTY"] !== undefined ? row[col["CNFM QTY"]] : 0);
      var remaining = num_(col["잔여수량"] !== undefined ? row[col["잔여수량"]] : 0);
      entry.remaining += remaining;
      if (remaining > 0 && entry.shipments.indexOf(shipment) === -1) entry.shipments.push(shipment);
      CHANNELS.forEach(function (channel) {
        if (col[channel] === undefined) return;
        var qty = channelQty_(row[col[channel]]);
        if (qty) entry.channels[channel] = (entry.channels[channel] || 0) + qty;
      });
    }
  }
  return { bySku: bySku, partial: partial };
}

/** Channel cells look like "30 (iHerb)", "50 (NY)", or plain numbers. */
function channelQty_(value) {
  var m = String(value || "").match(/([\d,]+)/);
  return m ? num_(m[1]) : 0;
}

/* ------------------------------------------------------------------ */
/* Writers                                                             */
/* ------------------------------------------------------------------ */

function writeInventoryTab_(stock, incoming, containers) {
  var ss = SpreadsheetApp.openById(INVENTORY_SYNC.masterId);
  var sheet = ss.getSheetByName(INVENTORY_SYNC.inventoryTab) || ss.insertSheet(INVENTORY_SYNC.inventoryTab);

  var headers = ["SKU", "Product Name", "Brand", "Barcode",
    "On Hand (Actual)", "Available", "On Hold", "Incoming (Confirmed)", "Remaining To Receive",
    "Inbound Shipments (차수)", "Locations", "Nearest Expiry", "Channel Allocation", "Flag", "Updated"];

  var skus = {};
  Object.keys(stock.bySku).forEach(function (sku) { skus[sku] = true; });
  Object.keys(incoming.bySku).forEach(function (sku) { skus[sku] = true; });

  var now = new Date();
  var rows = Object.keys(skus).sort().map(function (sku) {
    var onHand = stock.bySku[sku];
    var inbound = incoming.bySku[sku];
    var avail = onHand ? onHand.avail : 0;
    var remaining = inbound ? inbound.remaining : 0;
    var flag = "";
    if (avail <= 0 && remaining <= 0) flag = "OUT OF STOCK";
    else if (avail < INVENTORY_SYNC.lowStockThreshold && remaining <= 0) flag = "LOW STOCK";
    else if (avail < INVENTORY_SYNC.lowStockThreshold && remaining > 0) flag = "LOW — INBOUND EN ROUTE";

    var shipmentsWithStatus = inbound ? inbound.shipments.map(function (code) {
      var container = containers.byCode[String(code).toUpperCase()];
      return container && container.receivedDate ? code + " (rcvd " + container.receivedDate + ")" : code;
    }).join(", ") : "";

    var channelSummary = inbound ? Object.keys(inbound.channels).map(function (channel) {
      return channel + ":" + inbound.channels[channel];
    }).join(" · ") : "";

    return [
      sku,
      (onHand && onHand.name) || (inbound && inbound.name) || "",
      (onHand && onHand.brand) || (inbound && inbound.brand) || "",
      (onHand && onHand.barcode) || (inbound && inbound.barcode) || "",
      onHand ? onHand.actual : 0,
      avail,
      onHand ? onHand.hold : 0,
      inbound ? inbound.confirmed : 0,
      remaining,
      shipmentsWithStatus,
      onHand ? onHand.locations.slice(0, 6).join(", ") : "",
      onHand ? onHand.nearestExpiry : "",
      channelSummary,
      flag,
      now
    ];
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#EFEFEF");
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}

function updateKpiDashboard_(stock, incoming, containers) {
  var ss = SpreadsheetApp.openById(INVENTORY_SYNC.masterId);
  var sheet = ss.getSheetByName(INVENTORY_SYNC.kpiTab) || ss.insertSheet(INVENTORY_SYNC.kpiTab);

  var stockSkus = Object.keys(stock.bySku);
  var unitsOnHand = 0, unitsAvailable = 0;
  stockSkus.forEach(function (sku) { unitsOnHand += stock.bySku[sku].actual; unitsAvailable += stock.bySku[sku].avail; });

  var incomingSkus = Object.keys(incoming.bySku);
  var unitsIncoming = 0;
  incomingSkus.forEach(function (sku) { unitsIncoming += incoming.bySku[sku].remaining; });

  var lowStock = stockSkus.filter(function (sku) {
    var e = stock.bySku[sku];
    var remaining = incoming.bySku[sku] ? incoming.bySku[sku].remaining : 0;
    return e.avail < INVENTORY_SYNC.lowStockThreshold && remaining <= 0;
  }).length;

  var pendingContainers = containers.rows.filter(function (r) { return !r.receivedDate; }).length;
  var awaitingQc = containers.rows.filter(function (r) { return r.receivedDate && !r.qcDoneDate; }).length;

  var metrics = [
    ["Metric", "Value", "Updated: " + new Date()],
    ["SKUS TRACKED", stockSkus.length + incomingSkus.filter(function (s) { return !stock.bySku[s]; }).length, ""],
    ["UNITS ON HAND", unitsOnHand, ""],
    ["UNITS AVAILABLE", unitsAvailable, ""],
    ["UNITS INBOUND (REMAINING)", unitsIncoming, ""],
    ["LOW / OUT-OF-STOCK SKUS", lowStock, "avail < " + INVENTORY_SYNC.lowStockThreshold + " with nothing inbound"],
    ["CONTAINERS IN TRANSIT", pendingContainers, "receiving log rows without 입고일"],
    ["CONTAINERS AWAITING QC", awaitingQc, "received, 검수 not complete"],
    ["PENDING VERIFICATION (EMAIL)", pendingVerificationCount_(), "rows needing manual review"]
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, metrics.length, 3).setValues(metrics);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#EFEFEF");
  SpreadsheetApp.flush();
}

/* ------------------------------------------------------------------ */
/* IMPORTS enrichment                                                  */
/* ------------------------------------------------------------------ */

/**
 * Optional daily job: pushes receiving/QC dates from the WMS container log
 * onto matching IMPORTS rows (matched by 차수 code or PC number appearing in
 * any cell of the row).
 */
function enrichImportsFromContainerLog() {
  var containers = readWmsContainerLog_();
  if (!containers.rows.length) return { updated: 0 };
  var ss = SpreadsheetApp.openById(INVENTORY_SYNC.masterId);
  var sheet = ss.getSheetByName(INVENTORY_SYNC.importsTab);
  if (!sheet) return { updated: 0 };

  var data = sheet.getDataRange().getDisplayValues();
  var headerIdx = findHeaderRowIdx_(data);
  var map = headerMap_(data[headerIdx]);
  var noteCol = map["NOTE"] !== undefined ? map["NOTE"] : (map["REMARK"] !== undefined ? map["REMARK"] : null);

  var updated = 0;
  for (var r = headerIdx + 1; r < data.length; r++) {
    var rowText = data[r].join(" ").toUpperCase();
    for (var i = 0; i < containers.rows.length; i++) {
      var c = containers.rows[i];
      var hit = (c.pcNumber && rowText.indexOf(c.pcNumber.toUpperCase()) !== -1) ||
                (c.shipmentCode && rowText.indexOf(c.shipmentCode.toUpperCase()) !== -1);
      if (!hit || !c.receivedDate) continue;
      var tag = "[WMS " + c.shipmentCode + ": rcvd " + c.receivedDate + (c.qcDoneDate ? ", QC " + c.qcDoneDate : "") + "]";
      if (noteCol !== null && String(data[r][noteCol]).indexOf(tag) === -1) {
        sheet.getRange(r + 1, noteCol + 1).setValue((data[r][noteCol] ? data[r][noteCol] + " " : "") + tag);
        updated++;
      }
      break;
    }
  }
  SpreadsheetApp.flush();
  return { updated: updated };
}

function num_(value) {
  var n = Number(String(value === undefined || value === null ? "" : value).replace(/[^0-9.-]/g, ""));
  return isNaN(n) ? 0 : n;
}
