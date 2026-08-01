const SPREADSHEET_ID = "1M-vZ24Yw4ZN7R7b_473cVn8kny8DznTakSsD3VQsCzc";
const WMS_SPREADSHEET_ID = "14lH9SQzTLj8MR7UbxMfkoTDDlzhPoE8CqHV3IpK450I";

const OUTBOUND_STATUS = ["", "SCHEDULED", "WORK IN PROGRESS", "PENDING", "SHIPPING", "SHIPPED", "DELIVERED", "RECEIVED", "CANCELLED", "COMPLETED"];
const INBOUND_STATUS = ["", "SCHEDULED", "WORK IN PROGRESS", "PENDING", "SHIPPING", "SHIPPED", "DELIVERED", "RECEIVED", "CANCELLED", "COMPLETED", "N/A", "Customs Clearance", "FDA Review/Hold", "FWS Review/Hold", "Delayed"];
const ALLOWED_SHEETS = ["WH Trucking Request", "B2B/E-COM TRUCKING", "TRANSFERS", "ULTA", "IHERB", "IMPORTS", "NATIONAL ORDER PROGRESS", "Outbound Shipping Schedule", "TJX/ROSS"];

const COMPLETED_STATUSES = ["SHIPPED", "DELIVERED", "RECEIVED", "CANCELLED", "COMPLETED"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let rawContents = (e && e.postData && e.postData.contents) || "";
    if (!rawContents && e && e.parameter && e.parameter.postData) {
      rawContents = e.parameter.postData;
    }
    const request = JSON.parse(rawContents || "{}");
    validateRequest_(request);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(request.sourceSheet);
    if (!sheet) throw new Error("Source sheet not found.");

    const target = request.kind === "inbound"
      ? findInboundTarget_(sheet, request)
      : findOutboundTarget_(sheet, request);

    const allowed = (request.kind === "inbound" ? INBOUND_STATUS : OUTBOUND_STATUS).map((value) => String(value).toUpperCase());
    const status = String(request.status || "").trim();
    if (!allowed.includes(status.toUpperCase())) throw new Error("Status is not allowed.");

    const current = String(target.getDisplayValue() || "").trim();
    const requestCurrent = String(request.currentStatus || "").trim();
    const normCurrent = current.toUpperCase();
    const normRequest = requestCurrent.toUpperCase();

    // Check concurrency, tolerating default status fallbacks ("" vs "SCHEDULED")
    if (requestCurrent && normCurrent && normCurrent !== normRequest && !(normCurrent === "" && normRequest === "SCHEDULED")) {
      Logger.log("Concurrency note: Current='" + current + "', Request='" + requestCurrent + "'");
    }

    target.setValue(status);

    // Format row in Google Sheets: Grey out completed rows, reset active rows
    const rowIdx = target.getRow();
    const rowRange = sheet.getRange(rowIdx, 1, 1, Math.max(sheet.getLastColumn(), 1));
    const isCompleted = COMPLETED_STATUSES.includes(status.toUpperCase());
    if (isCompleted) {
      rowRange.setBackground("#E8EAED").setFontColor("#5F6368");
    } else {
      rowRange.setBackground(null).setFontColor(null);
    }

    SpreadsheetApp.flush();
    return json_({ ok: true, sheet: sheet.getName(), row: rowIdx, status, isCompleted });
  } catch (error) {
    return json_({ ok: false, error: String(error.message || error) });
  } finally {
    lock.releaseLock();
  }
}

function validateRequest_(request) {
  if (!["outbound", "inbound"].includes(request.kind)) throw new Error("Invalid relation kind.");
  if (!ALLOWED_SHEETS.includes(request.sourceSheet)) throw new Error("Source sheet is not allowed.");
}

function findInboundTarget_(sheet, request) {
  const row = Number(request.sourceRow);
  if (!Number.isInteger(row) || row < 3 || row > sheet.getLastRow()) throw new Error("Invalid IMPORTS source row.");
  const headers = sheet.getRange(1, 1, 3, sheet.getLastColumn()).getDisplayValues();
  const header = findHeader_(headers, ["WEBSITE STATUS", "STATUS", "INBOUND STATUS", "SHIPMENT STATUS"]);
  if (!header) throw new Error("Inbound status column not found.");
  return sheet.getRange(row, header.column);
}

function findOutboundTarget_(sheet, request) {
  const values = sheet.getDataRange().getDisplayValues();
  const header = findHeader_(values.slice(0, 4), ["WEBSITE STATUS", "STATUS", "WORK PROGRESS", "INBOUND STATUS", "SHIPMENT STATUS"]);
  if (!header) throw new Error("Status column not found.");
  const map = headerMap_(values[header.row - 1]);
  const sourceRow = Number(request.sourceRow);
  if (Number.isInteger(sourceRow) && sourceRow > header.row && sourceRow <= values.length) {
    return sheet.getRange(sourceRow, header.column);
  }
  const candidates = [];
  for (let r = header.row; r < values.length; r++) {
    const row = values[r];
    let score = 0;
    score += exact_(row, map, ["PRO#", "BOL", "BOL#"], request.pro) ? 100 : 0;
    score += exact_(row, map, ["INVOICE", "INVOICE NO.", "PO#"], request.invoice) ? 50 : 0;
    score += exact_(row, map, ["CUSTOMER", "NOTE", "DC"], request.customer) ? 20 : 0;
    score += exact_(row, map, ["SHIP DATE", "PU", "DATE"], request.shipDate) ? 10 : 0;
    if (score) candidates.push({ row: r + 1, score });
  }
  candidates.sort((a, b) => b.score - a.score);
  if (!candidates.length || (candidates[1] && candidates[0].score === candidates[1].score)) {
    throw new Error("Could not identify one unique source row.");
  }
  return sheet.getRange(candidates[0].row, header.column);
}

function findHeader_(rows, names) {
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      if (names.includes(String(rows[r][c] || "").trim().toUpperCase())) return { row: r + 1, column: c + 1 };
    }
  }
  return null;
}

function headerMap_(headers) {
  return headers.reduce((map, value, index) => {
    map[String(value || "").trim().toUpperCase()] = index;
    return map;
  }, {});
}

function exact_(row, map, names, expected) {
  const wanted = String(expected || "").trim().toUpperCase();
  if (!wanted) return false;
  const wantedParts = wanted.split(/[\r\n,;·]+/).map(p => p.trim()).filter(Boolean);

  for (const name of names) {
    if (map[name] === undefined) continue;
    const cellVal = String(row[map[name]] || "").trim().toUpperCase();
    if (cellVal === wanted) return true;
    const parts = cellVal.split(/[\r\n,;·]+/).map(p => p.trim()).filter(Boolean);
    if (parts.some(p => wantedParts.includes(p))) return true;
  }
  return false;
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Periodically scans external "WMS Invoice and Issues" sheet (14lH9SQzTLj8MR7UbxMfkoTDDlzhPoE8CqHV3IpK450I)
 * for rows where "Shipping Method" is "Trucking", combines multiple invoices
 * for the same customer & ship date into one entry, and imports/updates into "WH Trucking Request".
 */
function scanAndImportWmsTruckingOrders() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: "Lock timeout" };
  try {
    let wmsSpreadsheet;
    try {
      wmsSpreadsheet = SpreadsheetApp.openById(WMS_SPREADSHEET_ID);
    } catch (e) {
      wmsSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    const targetSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sourceSheet = wmsSpreadsheet.getSheets()[0]; // First sheet in WMS workbook
    const targetSheet = targetSpreadsheet.getSheetByName("WH Trucking Request");
    if (!sourceSheet || !targetSheet) {
      Logger.log("WMS Source sheet or WH Trucking Request sheet not found.");
      return { ok: false, error: "Source or Target sheet missing." };
    }

    const sourceData = sourceSheet.getDataRange().getDisplayValues();
    if (sourceData.length < 2) return { ok: true, imported: 0, updated: 0 };


    // Locate header row in WMS sheet
    let headerRowIdx = -1;
    let shipMethodColIdx = -1;
    let invoiceColIdx = -1;
    let customerColIdx = -1;
    let shipDateColIdx = -1;
    let palletColIdx = -1;
    let carrierColIdx = -1;
    let proColIdx = -1;
    let noteColIdx = -1;

    for (let r = 0; r < Math.min(5, sourceData.length); r++) {
      const row = sourceData[r].map(c => String(c || "").trim().toUpperCase());
      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (shipMethodColIdx === -1 && (val.includes("SHIPPING METHOD") || val.includes("SHIP METHOD"))) shipMethodColIdx = c;
        if (invoiceColIdx === -1 && (val.includes("INVOICE") || val.includes("PO#") || val.includes("PO NUMBER"))) invoiceColIdx = c;
        if (customerColIdx === -1 && (val.includes("CUSTOMER") || val.includes("CLIENT") || val.includes("ACCOUNT"))) customerColIdx = c;
        if (shipDateColIdx === -1 && (val.includes("SHIP DATE") || val.includes("DATE") || val.includes("PU DATE"))) shipDateColIdx = c;
        if (palletColIdx === -1 && (val.includes("PALLET") || val.includes("PLT") || val.includes("QTY") || val.includes("CARTONS"))) palletColIdx = c;
        if (carrierColIdx === -1 && (val.includes("CARRIER") || val.includes("TRUCKING"))) carrierColIdx = c;
        if (proColIdx === -1 && (val.includes("PRO#") || val.includes("PRO") || val.includes("TRACKING") || val.includes("BOL"))) proColIdx = c;
        if (noteColIdx === -1 && (val.includes("NOTE") || val.includes("REMARK") || val.includes("MEMO") || val.includes("ISSUE"))) noteColIdx = c;
      }
      if (shipMethodColIdx !== -1) {
        headerRowIdx = r;
        break;
      }
    }

    if (shipMethodColIdx === -1) {
      Logger.log("Shipping Method column not found in WMS sheet.");
      return { ok: false, error: "Shipping Method column missing." };
    }

    // Group Trucking entries by (Customer + Ship Date)
    const groups = new Map();
    for (let r = headerRowIdx + 1; r < sourceData.length; r++) {
      const row = sourceData[r];
      const shipMethod = String(row[shipMethodColIdx] || "").trim();
      if (shipMethod.toUpperCase() !== "TRUCKING") continue;

      const invoice = invoiceColIdx !== -1 ? String(row[invoiceColIdx] || "").trim() : "";
      const customer = customerColIdx !== -1 ? String(row[customerColIdx] || "").trim() : "";
      const shipDate = shipDateColIdx !== -1 ? String(row[shipDateColIdx] || "").trim() : "";
      const pallets = palletColIdx !== -1 ? String(row[palletColIdx] || "").trim() : "";
      const carrier = carrierColIdx !== -1 ? String(row[carrierColIdx] || "").trim() : "";
      const pro = proColIdx !== -1 ? String(row[proColIdx] || "").trim() : "";
      const note = noteColIdx !== -1 ? String(row[noteColIdx] || "").trim() : "";

      const normCust = customer.toUpperCase().replace(/\s+/g, " ").trim();
      const normDate = shipDate.toUpperCase().trim();
      const groupKey = normCust ? (normCust + "___" + normDate) : ("UNKNOWN___" + r);

      if (!groups.has(groupKey)) groups.set(groupKey, []);
      groups.get(groupKey).push({ invoice, customer, shipDate, pallets, carrier, pro, note, rowIndex: r + 1 });
    }

    // Load target sheet existing rows to avoid duplicates
    const targetData = targetSheet.getDataRange().getDisplayValues();
    const targetHeaders = targetData.length > 0 ? targetData[1] || targetData[0] : [];
    const targetMap = headerMap_(targetHeaders);

    const existingRowsMap = new Map(); // key -> row index (1-based)
    for (let r = 2; r < targetData.length; r++) {
      const row = targetData[r];
      const invs = exactVal_(row, targetMap, ["INVOICE NO.", "INVOICE #", "INVOICE"]).split(/[\r\n,;·]+/);
      const cust = exactVal_(row, targetMap, ["CUSTOMER"]).toUpperCase().replace(/\s+/g, " ").trim();
      const date = exactVal_(row, targetMap, ["SHIP DATE"]).toUpperCase().trim();
      
      if (cust && date) existingRowsMap.set(cust + "___" + date, r + 1);
      invs.forEach(inv => {
        const cleanInv = inv.trim().toUpperCase();
        if (cleanInv) existingRowsMap.set("INV___" + cleanInv, r + 1);
      });
    }

    let importedCount = 0;
    let updatedCount = 0;

    groups.forEach((items, groupKey) => {
      const customer = items[0].customer;
      const shipDate = items[0].shipDate;
      const combinedInvoices = [...new Set(items.map(i => i.invoice).filter(Boolean))].join("\n");
      const combinedCarrier = items.map(i => i.carrier).find(Boolean) || "Trucking";
      const combinedPro = [...new Set(items.map(i => i.pro).filter(Boolean))].join("\n");
      const combinedPallets = [...new Set(items.map(i => i.pallets).filter(Boolean))].join(" · ");
      const combinedNote = [...new Set(items.map(i => i.note).filter(Boolean))].join(" · ") || "Imported from WMS Invoice & Issues";

      const normCust = customer.toUpperCase().replace(/\s+/g, " ").trim();
      const normDate = shipDate.toUpperCase().trim();
      const matchKey = normCust + "___" + normDate;
      
      let matchedRowIdx = existingRowsMap.get(matchKey);
      if (!matchedRowIdx) {
        for (const item of items) {
          if (item.invoice && existingRowsMap.has("INV___" + item.invoice.toUpperCase())) {
            matchedRowIdx = existingRowsMap.get("INV___" + item.invoice.toUpperCase());
            break;
          }
        }
      }

      if (matchedRowIdx) {
        // Update existing entry if invoice list or fields changed
        const rowRange = targetSheet.getRange(matchedRowIdx, 1, 1, Math.max(targetHeaders.length, 21));
        const currentVals = rowRange.getDisplayValues()[0];

        const invCol = targetMap["INVOICE NO."] !== undefined ? targetMap["INVOICE NO."] : targetMap["INVOICE #"];
        if (invCol !== undefined && combinedInvoices) {
          const curInvs = String(currentVals[invCol] || "").trim();
          if (curInvs !== combinedInvoices) {
            targetSheet.getRange(matchedRowIdx, invCol + 1).setValue(combinedInvoices);
            updatedCount++;
          }
        }
      } else {
        // Append new combined entry for customer + ship date
        const newRow = new Array(Math.max(targetHeaders.length, 21)).fill("");
        if (targetMap["CUSTOMER"] !== undefined) newRow[targetMap["CUSTOMER"]] = customer;
        if (targetMap["INVOICE NO."] !== undefined) newRow[targetMap["INVOICE NO."]] = combinedInvoices;
        else if (targetMap["INVOICE #"] !== undefined) newRow[targetMap["INVOICE #"]] = combinedInvoices;
        if (targetMap["SHIP DATE"] !== undefined) newRow[targetMap["SHIP DATE"]] = shipDate;
        if (targetMap["PALLET TYPE"] !== undefined) newRow[targetMap["PALLET TYPE"]] = combinedPallets;
        if (targetMap["CARRIER"] !== undefined) newRow[targetMap["CARRIER"]] = combinedCarrier;
        if (targetMap["PRO#"] !== undefined) newRow[targetMap["PRO#"]] = combinedPro;
        if (targetMap["NOTE"] !== undefined) newRow[targetMap["NOTE"]] = combinedNote;
        if (targetMap["STATUS"] !== undefined) newRow[targetMap["STATUS"]] = "WORK IN PROGRESS";

        targetSheet.appendRow(newRow);
        importedCount++;
      }
    });

    SpreadsheetApp.flush();
    Logger.log("WMS Scan completed. Combined Groups: " + groups.size + ", Imported: " + importedCount + ", Updated: " + updatedCount);
    return { ok: true, groups: groups.size, imported: importedCount, updated: updatedCount };
  } catch (err) {
    Logger.log("Error in scanAndImportWmsTruckingOrders: " + err.message);
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function exactVal_(row, map, names) {
  for (const n of names) {
    if (map[n] !== undefined && row[map[n]]) return String(row[map[n]]).trim();
  }
  return "";
}

/**
 * Creates or resets the 30-minute time-driven trigger for WMS Trucking scanner.
 * Deletes all obsolete/legacy triggers in the project to ensure a clean schedule.
 */
function create30MinTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const ALLOWED_TRIGGER_HANDLERS = ["scanAndImportWmsTruckingOrders"];
  
  for (let i = 0; i < triggers.length; i++) {
    const handler = triggers[i].getHandlerFunction();
    if (!ALLOWED_TRIGGER_HANDLERS.includes(handler) || handler === "scanAndImportWmsTruckingOrders") {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("Deleted obsolete/existing trigger for handler: " + handler);
    }
  }

  ScriptApp.newTrigger("scanAndImportWmsTruckingOrders")
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log("30-minute time-driven trigger cleanly provisioned for scanAndImportWmsTruckingOrders");
}

/**
 * Adds "WEBSITE STATUS" dropdown data validation column at the end of each source sheet
 * in LOGISTICS MASTER 2026, applying the same validation rules as Column AE of IMPORTS.
 * Explicitly excludes external sheets 14lH9SQzTLj8MR7UbxMfkoTDDlzhPoE8CqHV3IpK450I and 12Aty04yiLPPqz06AFDM8Y1Log2jEOqdXDqwiUV5yVX8.
 */
function addWebsiteStatusDropdownToAllSourceSheets() {
  const EXCLUDED_SPREADSHEET_IDS = [
    "14lH9SQzTLj8MR7UbxMfkoTDDlzhPoE8CqHV3IpK450I",
    "12Aty04yiLPPqz06AFDM8Y1Log2jEOqdXDqwiUV5yVX8"
  ];
  
  const targetSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (EXCLUDED_SPREADSHEET_IDS.includes(targetSpreadsheet.getId())) {
    Logger.log("Target spreadsheet is in excluded list. Skipping.");
    return { ok: false, error: "Spreadsheet excluded." };
  }

  const TARGET_SOURCE_TABS = [
    "TRANSFERS",
    "ULTA",
    "IHERB",
    "B2B/E-COM TRUCKING",
    "WH Trucking Request",
    "NATIONAL ORDER PROGRESS",
    "Outbound Shipping Schedule",
    "TJX/ROSS"
  ];

  const STATUS_LIST = [
    "SCHEDULED",
    "WORK IN PROGRESS",
    "PENDING",
    "SHIPPING",
    "SHIPPED",
    "DELIVERED",
    "RECEIVED",
    "CANCELLED",
    "COMPLETED"
  ];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_LIST, true)
    .setAllowInvalid(false)
    .setHelpText("Select a valid Website Status from the list.")
    .build();

  let modifiedCount = 0;

  TARGET_SOURCE_TABS.forEach((tabName) => {
    const sheet = targetSpreadsheet.getSheetByName(tabName);
    if (!sheet) {
      Logger.log("Sheet tab not found: " + tabName);
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;

    // Detect header row and column
    const headers = sheet.getRange(1, 1, Math.min(3, lastRow), sheet.getLastColumn()).getDisplayValues();
    let headerRowIdx = 1;
    let colIdx = -1;

    for (let r = 0; r < headers.length; r++) {
      const row = headers[r].map(c => String(c || "").trim().toUpperCase());
      const foundIdx = row.indexOf("WEBSITE STATUS");
      if (foundIdx !== -1) {
        headerRowIdx = r + 1;
        colIdx = foundIdx + 1;
        break;
      }
    }

    // If column doesn't exist, append header to last column + 1
    if (colIdx === -1) {
      colIdx = sheet.getLastColumn() + 1;
      headerRowIdx = 2; // Default header row index for standard tabs
      sheet.getRange(headerRowIdx, colIdx).setValue("WEBSITE STATUS").setFontWeight("bold");
    }

    // Apply data validation rule down the column
    const startRow = headerRowIdx + 1;
    const numRows = Math.max(lastRow - headerRowIdx, 100);
    const range = sheet.getRange(startRow, colIdx, numRows, 1);
    range.setDataValidation(rule);

    modifiedCount++;
    Logger.log("Applied WEBSITE STATUS dropdown to sheet: " + tabName + " (Col " + colIdx + ")");
  });

  SpreadsheetApp.flush();
  return { ok: true, sheetsUpdated: modifiedCount };
}

/**
 * Deletes non-essential tabs ("Dimensions", "Reference", "Summary", "Dashboard", "Outbound Data", "Inbound_Data")
 * from LOGISTICS MASTER 2026.
 */
function deleteUnnecessaryTabs() {
  const targetSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tabsToDelete = [
    "Dimensions", "Reference", "Summary", "Dashboard",
    "Outbound Data", "Inbound_Data", "Inbound Data",
    "DIMENSIONS", "REFERENCE", "SUMMARY", "DASHBOARD",
    "OUTBOUND DATA", "INBOUND_DATA", "INBOUND DATA"
  ];
  
  let deletedCount = 0;
  tabsToDelete.forEach((tabName) => {
    const sheet = targetSpreadsheet.getSheetByName(tabName);
    if (sheet) {
      targetSpreadsheet.deleteSheet(sheet);
      deletedCount++;
      Logger.log("Deleted non-essential sheet tab: " + tabName);
    }
  });

  SpreadsheetApp.flush();
  return { ok: true, tabsDeleted: deletedCount };
}




