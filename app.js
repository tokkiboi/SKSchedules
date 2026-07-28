/* ═══════════════════════════════════════════════════════════════
   StyleKorean Logistics Hub — consolidated application script.
   Replaces the previous app.js + kpi-sync.js + source-sync.js +
   all-rows-fix.js + all-sources-sync.js layering with one file.

   Imports ALL operational sources in LOGISTICS MASTER 2026:
     IMPORTS · TRANSFERS · ULTA · IHERB · B2B/E-COM TRUCKING ·
     WH Trucking Request · NATIONAL ORDER PROGRESS ·
     Outbound Shipping Schedule · TJX/ROSS ·
     OUTBOUND WEBSITE EXCLUSIONS · Outbound Shipping Schedule KPI block (Z1:AA5)

   Never imported: loginfo (credentials), dimension reference tabs.
   ═══════════════════════════════════════════════════════════════ */
"use strict";

/* ---------- constants & state ---------- */
const PLATFORM = globalThis.STYLEKOREAN_PLATFORM || {};
const SHEET_ID = PLATFORM.workbook?.id || "1M-vZ24Yw4ZN7R7b_473cVn8kny8DznTakSsD3VQsCzc";
const FINISHED = new Set(PLATFORM.finishedStatuses || ["Shipped", "Delivered", "Received", "Completed", "Cancelled"]);
const STATUS_OPTIONS = Object.freeze([
  "Scheduled", "Work in Progress", "Pending", "Shipping", "Shipped",
  "Delivered", "Received", "Cancelled", "Completed"
]);
const PARCEL_SECTIONS = /^(UPS|USPS|DHL|AMAZON|FEDEX)$/i;
const PARCEL_STATUS_OVERRIDES = new Map(Object.entries(PLATFORM.parcelStatusOverrides || {}));
const PLANNING_LABELS = /^(URGENT|SCHEDULED|NEED SCHEDULING|COMPLETED)$/i;
const AUTO_REFRESH_MS = Number(PLATFORM.refreshMs) || 10 * 60 * 1000;
const COMPLETE_ENDPOINT = String(globalThis.STYLEKOREAN_CONFIG?.completeEndpoint || "").trim();

const FALLBACK_SOURCES = [
  { tab: "IMPORTS",                    range: "A:AF", kind: "inbound",  gid: 1497250700 },
  { tab: "TRANSFERS",                  range: "A:N",  kind: "outbound", gid: 1834454901 },
  { tab: "ULTA",                       range: "A:N",  kind: "outbound", gid: 360479919 },
  { tab: "IHERB",                      range: "A:M",  kind: "outbound", gid: 955532469 },
  { tab: "B2B/E-COM TRUCKING",         range: "A:R",  kind: "outbound", gid: 1971553563 },
  { tab: "WH Trucking Request",        range: "A2:U", kind: "outbound", gid: 852802817 },
  { tab: "NATIONAL ORDER PROGRESS",    range: "A:U",  kind: "outbound", gid: 2026071601 },
  { tab: "Outbound Shipping Schedule", range: "A3:U", kind: "outbound", gid: 20260708 },
  { tab: "TJX/ROSS",                   range: "A:R",  kind: "outbound", gid: 1110009873 },
  { tab: "OUTBOUND WEBSITE EXCLUSIONS", range: "A:C", kind: "filter",  gid: 2026071701 }
];
const SOURCES = PLATFORM.sources?.length ? [...PLATFORM.sources] : FALLBACK_SOURCES;
const KPI_SOURCE = PLATFORM.kpiSource || { tab: "Outbound Shipping Schedule", range: "Z1:AA5", kind: "kpi", gid: 20260708 };

const SOURCE_COLORS = {
  "WH Trucking Request": "var(--c-wh)",
  "B2B/E-com Trucking": "var(--c-b2b)",
  "Transfers": "var(--c-transfers)",
  "Ulta": "var(--c-ulta)",
  "iHerb": "var(--c-iherb)",
  "National Order Progress": "var(--c-national-order)",
  "Outbound Shipping Schedule": "var(--c-ship-out)",
  "TJX/ROSS": "var(--c-tjx)"
};

let inboundRows = [];
let inboundPlanningRows = [];
let outboundRows = [];         /* includes finished rows; filtered at render */
let parcelRows = [];
let sourceHealth = [];         /* [{tab, ok, rows}] */
let costSummary = { ytd: 0, mtd: 0, finished: 0, kpiSource: "computed" };
let loading = false;

/* ---------- tiny DOM / text helpers ---------- */
const $ = (id) => document.getElementById(id);
const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => (
  { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
));
const clean = (v) => String(v ?? "").trim();
const useful = (r) => Object.entries(r).some(([key, value]) => !key.startsWith("__") && Boolean(value));

/* exact-name column getter (headers already normalized by objects()) */
function col(row, ...names) {
  for (const n of names) {
    const v = row[n.toUpperCase()];
    if (v) return v;
  }
  return "";
}
/* tolerant getter for tabs with messy headers (line breaks, trailing spaces,
   parentheticals — e.g. TJX/ROSS's "PO# ", "Alt.\nPU#\n(eg.NRT#)") */
function colLoose(row, ...names) {
  const wanted = names.map((n) => String(n).toUpperCase().replace(/[^A-Z0-9#]/g, ""));
  for (const key of Object.keys(row)) {
    const nk = key.toUpperCase().replace(/[^A-Z0-9#]/g, "");
    if (wanted.includes(nk) && row[key]) return row[key];
  }
  return "";
}

function invoiceNumber(row, ...preferred) {
  for (const name of [...preferred, "INVOICE #", "INVOICE NO.", "INVOICE NO", "INVOICE NUMBER", "INVOICE"]) {
    const value = colLoose(row, name);
    if (value) return value;
  }
  return "";
}

function trackingNumber(row, ...preferred) {
  for (const name of [...preferred, "PRO#", "PRO #", "TRACKING#", "TRACKING #", "TRACKING NUMBER", "BOL#", "BOL"]) {
    const value = colLoose(row, name);
    if (value) return value;
  }
  return "";
}

/* ---------- Google Sheets (gviz) fetch layer ---------- */
function parseGviz(text) {
  const a = text.indexOf("{");
  const b = text.lastIndexOf("}");
  if (a < 0 || b < 0) throw Error("Google Sheets did not return data");
  const payload = JSON.parse(text.slice(a, b + 1));
  if (payload.status && payload.status !== "ok") {
    const detail = payload.errors?.map((entry) => entry?.detailed_message || entry?.message).filter(Boolean).join("; ");
    throw Error(detail || `Google Sheets query failed (${payload.status})`);
  }
  if (!payload.table) throw Error("Google Sheets returned no table data");
  return payload.table;
}

async function fetchTable(tab, range, withHeaders = true, targetSheetId = SHEET_ID, gid) {
  const fileId = targetSheetId || SHEET_ID;
  const u = new URL(`https://docs.google.com/spreadsheets/d/${fileId}/gviz/tq`);
  u.searchParams.set("tqx", "out:json");
  u.searchParams.set("headers", withHeaders ? "1" : "0");
  if (gid) u.searchParams.set("gid", String(gid));
  else if (tab) u.searchParams.set("sheet", tab);
  if (range) u.searchParams.set("range", range);

  let lastError = null;
  for (let attempt = 0; attempt < 2; attempt++) {
    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(), 20000);
    try {
      const r = await fetch(u, { cache: "no-store", signal: ctl.signal });
      if (!r.ok) throw Error(`${tab}: HTTP ${r.status}`);
      const table = parseGviz(await r.text());
      const rowMatch = clean(range).match(/^[A-Z]+(\d+)/i);
      table.__sourceStartRow = rowMatch ? Number(rowMatch[1]) : 1;
      return table;
    } catch (e) {
      lastError = e.name === "AbortError" ? Error(`${tab}: timed out (20s)`) : e;
      if (attempt === 0) continue;
      throw lastError;
    } finally {
      clearTimeout(timer);
    }
  }
  throw lastError || Error(`${tab}: unknown fetch failure`);
}

function objects(table) {
  const headers = table.cols.map((c, i) =>
    (c.label || `COL_${i}`).toUpperCase().replace(/\s+/g, " ").trim()
  );
  const startRow = Number(table.__sourceStartRow || 1);
  return table.rows.map((r, rowIndex) => ({
    __sourceRow: startRow + rowIndex + 1,
    ...Object.fromEntries(headers.map((h, i) => {
      const c = r.c?.[i];
      return [h, c ? clean(c.f ?? c.v ?? "") : ""];
    }))
  }));
}
const rawCell = (row, i) => {
  const c = row.c?.[i];
  return c ? clean(c.f ?? c.v ?? "") : "";
};

/* ---------- parsing helpers ---------- */
/* FIX: the previous date parser hard-coded "month >= 9 → 2025, else 2026".
   When the year is missing we now pick whichever candidate year lands the
   date closest to today, so the board keeps working in any year.
   Memoized — the same date strings are parsed thousands of times across
   boards, sorting, and consolidation. */
const dateCache = new Map();
function parseDate(v) {
  const s = clean(v);
  if (dateCache.has(s)) return dateCache.get(s);
  const d = computeDate(s);
  dateCache.set(s, d);
  return d;
}
function computeDate(s) {
  let m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    let y = +m[3];
    if (y < 100) y += 2000;
    const month = +m[1], day = +m[2];
    const d = new Date(y, month - 1, day);
    return d.getFullYear() === y && d.getMonth() === month - 1 && d.getDate() === day ? d : null;
  }
  m = s.match(/(\d{1,2})\/(\d{1,2})/);
  if (!m) return null;
  const now = Date.now();
  return [-1, 0, 1]
    .map((off) => new Date(new Date().getFullYear() + off, +m[1] - 1, +m[2]))
    .reduce((a, b) => (Math.abs(b - now) < Math.abs(a - now) ? b : a));
}
function fmtDate(v) {
  const d = parseDate(v);
  if (!d) return clean(v);
  return `${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}/${String(d.getFullYear()).slice(-2)}`;
}
function lastDateIn(v) {
  const all = [...clean(v).matchAll(/\d{1,2}\/\d{1,2}(?:\/\d{2,4})?/g)];
  return all.length ? fmtDate(all.at(-1)[0]) : "";
}
function money(v) {
  const m = clean(v).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/g);
  return m ? m.reduce((s, n) => s + Number(n), 0) : 0;
}
function parseMonetaryRate(v) {
  const s = clean(v);
  if (!s || /cancel/i.test(s)) return 0;
  if (s.includes("$")) {
    const matches = [...s.matchAll(/\$\s*(\d+(?:,\d{3})*(?:\.\d+)?)/g)];
    if (matches.length) return matches.reduce((sum, m) => sum + Number(m[1].replace(/,/g, "")), 0);
  }
  const trailingMatch = s.match(/(?:-|:|\$)\s*(\d+(?:\.\d{1,2})?)\s*$/);
  if (trailingMatch) {
    const val = Number(trailingMatch[1]);
    if (!isNaN(val) && val > 0 && val <= 50000) return val;
  }
  if (/^\s*\$?\s*\d+(?:,\d{3})*(?:\.\d{1,2})?\s*$/.test(s)) {
    const val = Number(s.replace(/[\$,]/g, ""));
    if (!isNaN(val) && val > 0 && val <= 50000) return val;
  }
  return 0;
}
function classifyStatus(v) {
  v = String(v || "").toLowerCase();
  if (/cancel/.test(v)) return "Cancelled";
  if (/deliver/.test(v)) return "Delivered";
  if (/receive/.test(v)) return "Received";
  if (/\bshipped\b/.test(v)) return "Shipped";
  if (/\bdone\b|complete|closed|gr[ae]y(?:ed)?\s*out/.test(v)) return "Completed";
  if (/work\s*in\s*progress|\bwip\b/.test(v)) return "Work in Progress";
  if (/pending/.test(v)) return "Pending";
  if (/shipping|transit|progress/.test(v)) return "Shipping";
  return "Scheduled";
}
function effectiveStatus(row, mappedStatus = "Scheduled") {
  const detected = classifyStatus(Object.values(row || {}).join(" "));
  return FINISHED.has(detected) ? detected : mappedStatus;
}
function containerNumbers(v) {
  return clean(v).split(/[,\n]+/)
    .map((p) => p.trim().replace(/\s/g, "").toUpperCase())
    .filter((p) => /^[A-Z]{4}\d{7}$/.test(p));
}

/* Container tracking priority:
   1) official carrier from source SCAC/carrier text or equipment prefix
   2) destination terminal / port community system
   3) third-party auto-detect fallback (SeaRates is intentionally not used) */
const OFFICIAL_CONTAINER_TRACKING = [
  { carrier: "HMM", keys: ["HDMU", "HMMU", "HMM"], url: (n) => "https://www.hmm21.com/e-service/general/trackNTrace/TrackNTrace.do?searchType=CNTR&searchNo=" + encodeURIComponent(n) },
  { carrier: "Maersk", keys: ["MAEU", "MSKU", "MRKU", "MAERSK"], url: (n) => "https://www.maersk.com/tracking/" + encodeURIComponent(n) },
  { carrier: "MSC", keys: ["MSCU", "MEDU", "MSC"], url: (n) => "https://www.msc.com/en/track-a-shipment?trackingNumber=" + encodeURIComponent(n) },
  { carrier: "CMA CGM", keys: ["CMDU", "CMAU", "CMA CGM"], url: (n) => "https://www.cma-cgm.com/ebusiness/tracking/search?SearchBy=Container&Reference=" + encodeURIComponent(n) },
  { carrier: "COSCO", keys: ["COSU", "CBHU", "COSCO"], url: (n) => "https://elines.coscoshipping.com/ebusiness/cargoTracking?trackingType=CONTAINER&number=" + encodeURIComponent(n) },
  { carrier: "OOCL", keys: ["OOLU", "OOCL"], url: () => "https://www.oocl.com/eng/ourservices/eservices/cargotracking/Pages/cargotracking.aspx" },
  { carrier: "ONE", keys: ["ONEY", "OCEAN NETWORK EXPRESS"], url: (n) => "https://ecomm.one-line.com/one-ecom/manage-shipment/cargo-tracking?trakNoParam=" + encodeURIComponent(n) },
  { carrier: "Evergreen", keys: ["EGLV", "EVERGREEN"], url: () => "https://ct.shipmentlink.com/servlet/TDB1_CargoTracking.do" },
  { carrier: "Yang Ming", keys: ["YMLU", "YANG MING"], url: () => "https://www.yangming.com/e-service/track_trace/track_trace_cargo_tracking.aspx" },
  { carrier: "ZIM", keys: ["ZIMU", "ZIM"], url: (n) => "https://www.zim.com/tools/track-a-shipment?consnumber=" + encodeURIComponent(n) },
  { carrier: "Hapag-Lloyd", keys: ["HLCU", "HAPAG"], url: (n) => "https://www.hapag-lloyd.com/en/online-business/track/track-by-container-solution.html?container=" + encodeURIComponent(n) },
  { carrier: "SM Line", keys: ["SMLM", "SMCU", "SM LINE"], url: (n) => "https://esvc.smlines.com/smline/CUP_HOM_3301.do?search_type=C&search_name=" + encodeURIComponent(n) }
];
const SHIPMENT_CARRIER_HINTS = [
  { carrier: "Maersk", patterns: [/\b(?:MCI|ES|OSL)\d+\b/i] },
  { carrier: "HMM", patterns: [/\bHJ\d+\b/i] }
];
const TERMINAL_CONTAINER_TRACKING = [
  { name: "APM Terminals Pier 400", keys: ["APMT", "PIER 400"], url: "https://www.apmterminals.com/en/los-angeles/practical-information/track-and-trace" },
  { name: "Fenix Marine Services", keys: ["FENIX", "FMS TERMINAL"], url: "https://fenixmarineservices.com/" },
  { name: "Long Beach Container Terminal", keys: ["LBCT"], url: "https://www.lbct.com/" },
  { name: "Total Terminals International", keys: ["TTI"], url: "https://www.totalterminals.com/" },
  { name: "Yusen Terminals", keys: ["YTI"], url: "https://yti.com/" },
  { name: "West Basin Container Terminal", keys: ["WBCT"], url: "https://www.portsamerica.com/locations/west-coast/wbct" },
  { name: "International Transportation Service", keys: ["ITS TERMINAL"], url: "https://www.itslb.com/" },
  { name: "TraPac", keys: ["TRAPAC"], url: "https://www.trapac.com/" }
];

function containerTrackingProfile(row, container, destination = "") {
  const n = clean(container).replace(/\s/g, "").toUpperCase();
  if (!n) return { url: "", source: "", carrier: "" };
  const rowText = [
    n, destination,
    col(row || {}, "SCAC", "CARRIER SCAC", "OCEAN CARRIER", "CARRIER", "LINE", "FORWARDER", "TERMINAL", "POD", "PORT"),
    ...Object.values(row || {})
  ].join(" ").toUpperCase();

  const official = OFFICIAL_CONTAINER_TRACKING.find((profile) =>
    profile.keys.some((key) => rowText.includes(key))
  );
  if (official) return { url: official.url(n), source: official.carrier + " official", carrier: official.carrier };

  const shipmentHint = SHIPMENT_CARRIER_HINTS.find((hint) =>
    hint.patterns.some((pattern) => pattern.test(rowText))
  );
  const hintedOfficial = shipmentHint && OFFICIAL_CONTAINER_TRACKING.find((profile) =>
    profile.carrier === shipmentHint.carrier
  );
  if (hintedOfficial) {
    return { url: hintedOfficial.url(n), source: hintedOfficial.carrier + " official", carrier: hintedOfficial.carrier };
  }

  const terminal = TERMINAL_CONTAINER_TRACKING.find((profile) =>
    profile.keys.some((key) => rowText.includes(key))
  );
  if (terminal) return { url: terminal.url, source: terminal.name, carrier: "" };

  if (/(?:LOS ANGELES|LONG BEACH|USLAX|USLGB|LA\s*\/\s*LONG BEACH)/i.test(rowText)) {
    return { url: "https://track.portoptimizer.com/", source: "LA/LB Port Optimizer", carrier: "" };
  }

  return {
    url: "https://www.track-trace.com/container?number=" + encodeURIComponent(n),
    source: "Track-Trace fallback",
    carrier: ""
  };
}

function looksLikeParcelTracking(raw) {
  const n = clean(raw).replace(/\s/g, "").toUpperCase();
  return /^1Z[0-9A-Z]{16}$/.test(n) || /^TBA[0-9A-Z]+$/.test(n) ||
    /^(94|93|92|95)\d{18,22}$/.test(n) || /^[A-Z]{2}\d{9}US$/.test(n) ||
    /^\d{10}$/.test(n) || /^(?:\d{12}|\d{15}|\d{20}|\d{22})$/.test(n);
}
function parcelTrackingUrl(carrier, num) {
  const n = encodeURIComponent(num);
  switch (carrier) {
    case "UPS":   return `https://www.ups.com/track?tracknum=${n}`;
    case "FedEx": return `https://www.fedex.com/fedextrack/?trknbr=${n}`;
    case "USPS":  return `https://tools.usps.com/go/TrackConfirmAction?tLabels=${n}`;
    case "DHL":   return `https://www.dhl.com/us-en/home/tracking.html?tracking-id=${n}`;
    case "Amazon": return `https://track.amazon.com/tracking/${n}`;
    default:      return "";
  }
}
function inferParcelCarrier(sectionCarrier, tracking) {
  const n = clean(tracking).replace(/\s/g, "").toUpperCase();
  if (/^1Z[0-9A-Z]{16}$/.test(n)) return "UPS";
  if (/^TBA[0-9A-Z]+$/.test(n)) return "Amazon";
  if (/^(94|93|92|95)\d{18,22}$/.test(n) || /^[A-Z]{2}\d{9}US$/.test(n)) return "USPS";
  if (/^\d{10}$/.test(n)) return "DHL";
  if (/^(?:\d{12}|\d{15}|\d{20}|\d{22})$/.test(n)) return /^DHL$/i.test(sectionCarrier) ? "DHL" : "FedEx";
  if (/^FEDEX$/i.test(sectionCarrier)) return "FedEx";
  if (/^(UPS|USPS|DHL)$/i.test(sectionCarrier)) return sectionCarrier.toUpperCase();
  if (/^AMAZON$/i.test(sectionCarrier)) return "Amazon";
  return sectionCarrier;
}

/* ---------- inbound: IMPORTS tab ---------- */
function isInboundShipment(r) {
  const shipment = col(r, "SHIPMENT");
  const invoice = col(r, "INVOICE");
  const container = col(r, "CONTAINER", "CONTAINER RAW (SYSTEM)");
  const mbl = col(r, "MBL").replace(/\s/g, "");
  const rowText = Object.values(r).join(" ");
  const hasInvoice = /\bIN\d{4,}\b/i.test(invoice);
  const hasCleanContainer = containerNumbers(container).length > 0;
  const hasAirWaybill = /^\d{3}-?\d{8}$/.test(mbl);
  const isPlanningRow = PLANNING_LABELS.test(shipment.trim()) ||
    (!hasInvoice && !hasCleanContainer && !hasAirWaybill && /-\s*20\d\d\b/.test(rowText));
  const hasParcelTracking = Object.values(r).some(looksLikeParcelTracking);
  const sectionLabel = Object.values(r).find((v) => PARCEL_SECTIONS.test(clean(v)));
  const hasShipmentIdentifier = !/^NEW$/i.test(shipment.trim()) &&
    (hasInvoice || hasCleanContainer || hasAirWaybill);
  return useful(r) && hasShipmentIdentifier && !hasParcelTracking && !sectionLabel && !isPlanningRow;
}

function portOfLoading(r) {
  const explicit = col(r, "POL", "PORT OF LOADING", "ORIGIN", "POL / ORIGIN", "LOADING PORT");
  if (explicit && !/IMPORTS|KOREA\s*\/\s*ASIA|PLANNING\s*GRID/i.test(explicit)) {
    return explicit;
  }
  const rowText = Object.values(r || {}).join(" ").toUpperCase();
  if (/INCHEON|ICN/.test(rowText)) return "INCHEON";
  if (/PUSAN|BUSAN|KRPUS/.test(rowText)) return "BUSAN";
  if (/SHANGHAI|CNSHA/.test(rowText)) return "SHANGHAI";
  if (/NINGBO|CNNGB/.test(rowText)) return "NINGBO";
  if (/YANTIAN|CNYTN/.test(rowText)) return "YANTIAN";
  return "BUSAN";
}

function mapInbound(rows) {
  return rows.filter(isInboundShipment).map((r) => {
    const containerRaw = col(r, "CONTAINER", "CONTAINER RAW (SYSTEM)");
    const container = containerNumbers(containerRaw)[0] || containerRaw.split(/[,\n]/)[0].trim();
    const mbl = col(r, "MBL");
    const isAir = !container && /^\d{3}-?\d{8}$/.test(mbl.replace(/\s/g, ""));
    const destination = col(r, "DESTINATION", "POD", "PORT", "DELIVERY") || "LA / Long Beach";
    const tracking = containerTrackingProfile(r, container, destination);
    const sourceCarrier = col(r, "CARRIER", "LINE", "FORWARDER");
    const vesselOrFlight = col(r, "VSL");
    const sourceStatus = col(r, "WEBSITE STATUS", "STATUS", "SHIPMENT STATUS");
    return {
      mode: isAir ? "Air" : "Ocean",
      eta: fmtDate(col(r, "ESTIMATED DELIVERY", "ESTIMATED DELIVERY DATE", "ESTIMATED_DELIVERY", "ETA")),
      shipmentNo: col(r, "SHIPMENT"),
      invoice: invoiceNumber(r, "INVOICE", "INVOICE #"),
      mbl,
      hbl: col(r, "HBL"),
      container,
      carrier: sourceCarrier || (isAir ? vesselOrFlight || "Air freight" : tracking.carrier || "Ocean freight"),
      trackingUrl: tracking.url,
      trackingSource: tracking.source,
      origin: portOfLoading(r),
      destination,
      sourceTab: "IMPORTS",
      sourceRow: r.__sourceRow || 0,
      sourceStatus,
      status: effectiveStatus(r, classifyStatus([sourceStatus, col(r, "NOTES"), col(r, "RESERVED"), col(r, "DELIVERY EXPECTED")].join(" ")))
    };
  });
}

function planningDate(value) {
  const match = clean(value).match(/(\d{1,2})[\/-](\d{1,2})(?:[\/-](\d{2,4}))?/);
  if (!match) return "";
  return fmtDate(`${match[1]}/${match[2]}${match[3] ? "/" + match[3] : ""}`);
}

function isManualPlanningItem(value) {
  const text = clean(value);
  if (!text) return false;
  if (/^(SCHEDULED|NEED SCHEDULING|MONTH OF AUGUST|URGENT|COMPLETED|ESTIMATED\s*\/\s*CHANGED|미정|AIR|ARRIVAL)$/i.test(text)) return false;
  if (/^AS OF\b/i.test(text)) return false;
  if (/^\d{1,2}[\/-]\d{1,2}(?:[\/-]\d{2,4})?(?:\s*-\s*ARRIVAL)?$/i.test(text)) return false;
  if (/^\d{4,6}$/.test(text)) return false;
  return true;
}

function buildInboundPlanningRow(value, eta, sourceRow) {
  const text = clean(value);
  const container = text.toUpperCase().match(/\b[A-Z]{4}\d{7}\b/)?.[0] || "";
  const shipmentNo = container
    ? text
      .replace(new RegExp(`\\s*-?\\s*${container}\\s*$`, "i"), "")
      .replace(/\s*-\s*$/, "")
      .trim() || container
    : text;
  const isAir = !container && /(AIR|^JSL|^KYL|^MBX|USMM)/i.test(text);
  const tracking = container ? containerTrackingProfile({ SHIPMENT: text }, container, "LA / Long Beach") : { url: "", source: "" };
  return {
    mode: isAir ? "Air" : "Ocean",
    eta,
    shipmentNo,
    invoice: "",
    mbl: "",
    hbl: "",
    container,
    carrier: container ? "Ocean freight" : "Air freight",
    trackingUrl: tracking.url,
    trackingSource: tracking.source,
    sourceTab: "IMPORTS",
    sourceRow,
    sourceStatus: "",
    origin: "BUSAN",
    destination: "LA / Long Beach",
    status: "Scheduled"
  };
}

/* IMPORTS contains a calendar-style planning grid below the detailed shipment
   table. The public import board mirrors only the manual SCHEDULED block the
   team edits there, then stops before the later NEED SCHEDULING backlog. */
function mapInboundPlanningGrid(table) {
  const rows = table.rows || [];
  const marker = rows.findIndex((row) =>
    /^URGENT$/i.test(rawCell(row, 0)) &&
    /^COMPLETED$/i.test(rawCell(row, 1)) &&
    /ESTIMATED\s*\/\s*CHANGED/i.test(rawCell(row, 2))
  );
  if (marker < 0) return [];

  const topDates = new Map();
  const topDateRow = rows[marker + 1];
  (topDateRow?.c || []).forEach((_, column) => {
    const date = planningDate(rawCell(topDateRow, column));
    if (date) topDates.set(column, date);
  });

  const planned = [];
  let inScheduledBlock = false;
  let blankRun = 0;
  for (let rowIndex = marker + 2; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const first = rawCell(row, 0).toUpperCase();
    const values = (row.c || []).map((_, column) => rawCell(row, column));
    const hasAnyValue = values.some((value) => clean(value));
    if (PARCEL_SECTIONS.test(first)) break;
    if (first === "NEED SCHEDULING" || values.some((value) => /^MONTH OF AUGUST$/i.test(clean(value)))) break;
    if (!hasAnyValue) {
      blankRun += 1;
      if (inScheduledBlock && blankRun >= 2) break;
      continue;
    }
    blankRun = 0;
    if (first === "SCHEDULED") inScheduledBlock = true;
    if (!inScheduledBlock) continue;

    (row.c || []).forEach((_, column) => {
      if (column === 0) return;
      const value = rawCell(row, column);
      const eta = topDates.get(column) || "";
      if (!eta || !isManualPlanningItem(value)) return;
      const sourceRow = Number(table.__sourceStartRow || 1) + rowIndex + 1;
      planned.push(buildInboundPlanningRow(value, eta, sourceRow));
    });
  }
  return planned;
}

function mergeInboundPlanning(detailed, planned) {
  const merged = detailed.map((row) => ({ ...row }));
  const byContainer = new Map(merged.map((row, index) => [clean(row.container).toUpperCase(), index]));
  planned.forEach((plan) => {
    const key = clean(plan.container).toUpperCase();
    const index = byContainer.get(key);
    if (index == null) {
      byContainer.set(key, merged.length);
      merged.push(plan);
      return;
    }
    const current = merged[index];
    merged[index] = {
      ...current,
      eta: plan.eta || current.eta,
      shipmentNo: current.shipmentNo || plan.shipmentNo,
      status: FINISHED.has(plan.status) ? plan.status : current.status,
      origin: (current.origin && !/IMPORTS|KOREA\s*\/\s*ASIA/i.test(current.origin)) ? current.origin : (plan.origin || "BUSAN")
    };
  });
  return merged;
}

function mapParcels(table) {
  let section = "";
  const result = [];
  for (const row of table.rows || []) {
    const cells = (row.c || []).map((_, i) => rawCell(row, i));
    const rowText = cells.join(" ");

    const sectionLabel = cells.find((value) => PARCEL_SECTIONS.test(clean(value)));
    if (sectionLabel) section = sectionLabel.toUpperCase();

    const trackingCell = cells.find((value) => {
      const normalized = clean(value).replace(/\s/g, "").toUpperCase();
      return normalized && !/^TRACKING#?$/i.test(normalized) && looksLikeParcelTracking(normalized);
    });
    if (!trackingCell) continue;

    const num = clean(trackingCell).replace(/\s/g, "").toUpperCase();
    const carrier = inferParcelCarrier(section, num);
    if (!carrier) continue;

    const invoice = clean(rawCell(row, 2)) ||
      (cells.find((value) => /\bIN\d{4,}\b/i.test(value)) || "");
    const origin = clean(rawCell(row, 3)) ||
      (cells.find((value) => /KOREA|ASIA|LAX|LA\s*\/\s*LONG\s*BEACH/i.test(value)) || "Imports");
    const detail = clean(rawCell(row, 4)) ||
      cells.filter(Boolean).slice(2).join(" · ");

    let status = classifyStatus(rowText);
    if (/label created|not shipped/i.test(detail)) status = "Scheduled";
    else if (!FINISHED.has(status) && /pending|clearance|customs|waiting|transit/i.test(rowText)) status = "Shipping";
    status = PARCEL_STATUS_OVERRIDES.get(num) || status;

    result.push({
      carrier,
      tracking: num,
      invoice: clean(invoice),
      origin: clean(origin) || "Imports",
      eta: lastDateIn(detail || rowText),
      note: detail.replace(/\s+/g, " ").trim() || "No carrier tracking note yet",
      url: parcelTrackingUrl(carrier, num),
      status
    });
  }
  return result;
}

function activeParcels() {
  return parcelRows.filter((row) => !FINISHED.has(row.status));
}

/* ---------- outbound mappers (one per source tab) ---------- */
function pushOutbound(source, r, mapped, excludedFn) {
  const key = mapped.pro || mapped.invoice || "";
  if (excludedFn(source, key)) return;
  const shipDate = fmtDate(mapped.shipDate);
  const rowStatus = effectiveStatus(r, mapped.status);
  outboundRows.push({
    source,
    sourceRow: r.__sourceRow || 0,
    sourceStatus: col(r, "WEBSITE STATUS", "STATUS", "OVERALL PO STATUS", "WORK PROGRESS"),
    sourceTab: ({
      "Transfers": "TRANSFERS", "Ulta": "ULTA", "iHerb": "IHERB",
      "B2B/E-com Trucking": "B2B/E-COM TRUCKING", "WH Trucking Request": "WH Trucking Request",
      "National Order Progress": "NATIONAL ORDER PROGRESS", "Outbound Shipping Schedule": "Outbound Shipping Schedule",
      "TJX/ROSS": "TJX/ROSS"
    })[source] || "",
    shipDate,
    customer: clean(mapped.customer),
    invoice: clean(mapped.invoice),
    origin: clean(mapped.origin || ""),
    carrier: clean(mapped.carrier),
    pro: clean(mapped.pro),
    units: clean(mapped.units),
    qty: col(r, "Q'TY (PLTS / CTNS)", "QTY"),
    length: col(r, "LENGTH (IN)"),
    width: col(r, "WIDTH (IN)"),
    height: col(r, "HEIGHT (IN)", "HEIGHT"),
    weight: col(r, "WEIGHT (LBS)", "WEIGHT"),
    destination: clean(mapped.destination || ""),
    rate: (() => {
      const srcNorm = String(source || "").toUpperCase();
      const rateVal = parseMonetaryRate(col(r, "RATE", "RATE QUOTE", "RATE QUOTE AMOUNT", "INVOICE AMOUNT", "QUOTE AMOUNT", "QUOTE"));

      if (srcNorm.includes("B2B") || srcNorm.includes("IHERB")) {
        return rateVal;
      }
      if (srcNorm.includes("TRANSFER")) {
        const invVal = parseMonetaryRate(col(r, "INVOICE"));
        return invVal || rateVal;
      }
      if (srcNorm.includes("ULTA")) {
        const invVal = parseMonetaryRate(col(r, "INVOICE", "INVOICE AMOUNT"));
        return invVal || rateVal;
      }
      if (srcNorm.includes("WH TRUCKING")) {
        const invVal = parseMonetaryRate(col(r, "INVOICE AMOUNT", "FREIGHT INVOICE", "INVOICE NO.", "INVOICE #"));
        return invVal || rateVal;
      }
      const generalInv = parseMonetaryRate(col(r, "INVOICE", "INVOICE AMOUNT", "INVOICE NO."));
      return rateVal || generalInv;
    })(),
    status: rowStatus
  });
}

function identifierParts(value) {
  return clean(value).split(/[\r\n,;·]+/).map((part) => part.trim().toUpperCase().replace(/\s+/g, "")).filter(Boolean);
}

function sharesIdentifier(left, right) {
  const wanted = new Set(identifierParts(left));
  return identifierParts(right).some((part) => wanted.has(part));
}

/* The workbook's Outbound Shipping Schedule is the consolidated operational
   view. Use it to fill identifiers missing from source rows; append only rows
   that are not represented by any authoritative source tab. */
function mergeOutboundScheduleIdentifiers(rows, excludedFn) {
  rows.filter(useful).forEach((r) => {
    const mapped = {
      shipDate: col(r, "SHIP DATE"),
      customer: col(r, "CUSTOMER"),
      invoice: invoiceNumber(r, "INVOICE NO.", "INVOICE"),
      origin: "",
      carrier: col(r, "CARRIER"),
      pro: trackingNumber(r, "PRO#"),
      units: col(r, "PALLET TYPE") ? "Pallets" : "",
      destination: col(r, "ADDRESS"),
      status: classifyStatus(`${col(r, "STATUS")} ${col(r, "NOTE")}`)
    };
    if (!mapped.customer && !mapped.invoice && !mapped.pro) return;
    const key = mapped.pro || mapped.invoice || "";
    if (excludedFn("Outbound Shipping Schedule", key)) return;

    const invoiceMatch = mapped.invoice && outboundRows.find((row) => sharesIdentifier(row.invoice, mapped.invoice));
    const trackingMatch = mapped.pro && outboundRows.find((row) => sharesIdentifier(row.pro, mapped.pro));
    const sameCustomerDate = outboundRows.filter((row) =>
      clean(row.customer).toUpperCase() === clean(mapped.customer).toUpperCase() &&
      clean(row.shipDate) === fmtDate(mapped.shipDate)
    );
    const match = invoiceMatch || trackingMatch || (sameCustomerDate.length === 1 ? sameCustomerDate[0] : null);
    if (!match) {
      pushOutbound("Outbound Shipping Schedule", r, mapped, excludedFn);
      return;
    }

    match.invoice ||= clean(mapped.invoice);
    match.pro ||= clean(mapped.pro);
    match.carrier ||= clean(mapped.carrier);
    match.units ||= clean(mapped.units);
    match.destination ||= clean(mapped.destination);
    match.length ||= col(r, "LENGTH (IN)");
    match.width ||= col(r, "WIDTH (IN)");
    match.height ||= col(r, "HEIGHT (IN)", "HEIGHT");
    match.weight ||= col(r, "WEIGHT (LBS)", "WEIGHT");
    if (!match.rate) match.rate = money(col(r, "RATE"));
  });
}

function mapAllOutbound(tabs, excludedFn) {
  const { tr, ul, ih, b2, wh, national, shipOut, tjxRoss } = tabs;
  outboundRows = [];

  tr.filter(useful).forEach((r) => pushOutbound("Transfers", r, {
    shipDate: col(r, "PU"),
    customer: col(r, "TO"),
    invoice: invoiceNumber(r, "INVOICE"),
    origin: col(r, "VENDOR/SUPPLIER/ORIGIN", "VENDOR / SUPPLIER / ORIGIN", "VENDOR", "SUPPLIER", "ORIGIN"),
    carrier: col(r, "TRUCKING"),
    pro: trackingNumber(r, "BOL#", "PU#"),
    units: col(r, "PLT") ? `${col(r, "PLT")} Pallets` : "",
    destination: col(r, "TO"),
    status: classifyStatus(`${col(r, "STATUS")} ${col(r, "NOTE")} ${col(r, "INVOICE")}`)
  }, excludedFn));

  ul.filter(useful).forEach((r) => pushOutbound("Ulta", r, {
    shipDate: col(r, "SHIP DATE", "DATE"),
    customer: col(r, "DC") || "Ulta",
    invoice: invoiceNumber(r, "PO#", "INVOICE"),
    carrier: col(r, "TRUCKING"),
    pro: trackingNumber(r, "PRO#"),
    units: col(r, "TOTAL CARTONS") ? `${col(r, "TOTAL CARTONS")} Cartons` : "",
    destination: col(r, "SHIP TO"),
    status: col(r, "PRO#") ? "Shipped"
      : classifyStatus(`${col(r, "STATUS")} ${col(r, "NOTE")} ${col(r, "REMARKS")}`)
  }, excludedFn));

  ih.filter(useful).forEach((r) => pushOutbound("iHerb", r, {
    shipDate: col(r, "PU", "DELIVERY APPT"),
    customer: `iHerb${col(r, "TO") ? " · " + col(r, "TO") : ""}`,
    invoice: invoiceNumber(r, "PO#"),
    carrier: col(r, "TRUCKING"),
    pro: trackingNumber(r, "BOL"),
    units: col(r, "QTY") ? `${col(r, "QTY")} Pallets` : "",
    destination: col(r, "TO"),
    status: classifyStatus(`${col(r, "STATUS")} ${col(r, "NOTE")} ${col(r, "REMARKS")}`)
  }, excludedFn));

  b2.filter((r) =>
    ["INVOICE", "PU", "TRUCKING", "PRO#", "PLT", "QTY", "RATE"].some((n) => col(r, n))
  ).forEach((r) => pushOutbound("B2B/E-com Trucking", r, {
    shipDate: col(r, "PU"),
    customer: col(r, "NOTE"),
    invoice: invoiceNumber(r, "INVOICE"),
    carrier: col(r, "TRUCKING"),
    pro: trackingNumber(r, "PRO#"),
    units: col(r, "PLT") ? `${col(r, "PLT")} Pallets` : "",
    destination: col(r, "TO"),
    status: classifyStatus(Object.values(r).join(" "))
  }, excludedFn));

  wh.filter((r) =>
    !/PLEASE LIST THE INVOICE WITH SHIPPING CHARGE/i.test(col(r, "CUSTOMER")) &&
    (col(r, "CUSTOMER") || col(r, "INVOICE NO.")) &&
    ["CUSTOMER", "INVOICE NO.", "SHIP DATE", "PALLET TYPE", "CARRIER", "PRO#"].some((n) => col(r, n))
  ).forEach((r) => pushOutbound("WH Trucking Request", r, {
    shipDate: col(r, "SHIP DATE"),
    customer: col(r, "CUSTOMER"),
    invoice: invoiceNumber(r, "INVOICE NO."),
    carrier: col(r, "CARRIER"),
    pro: trackingNumber(r, "PRO#"),
    units: col(r, "PALLET TYPE") ? "Pallets" : "",
    destination: col(r, "ADDRESS"),
    status: classifyStatus(`${col(r, "STATUS")} ${col(r, "NOTE")} ${col(r, "REMARKS")}`)
  }, excludedFn));

  national.filter((r) => col(r, "PICK-UP DATE", "START SHIP", "SHIPPING DATE", "SHIP DATE"))
  .forEach((r) => pushOutbound("National Order Progress", r, {
    shipDate: col(r, "PICK-UP DATE", "START SHIP", "SHIPPING DATE", "SHIP DATE"),
    customer: col(r, "CHANNEL"),
    invoice: invoiceNumber(r, "PO#", "ORDER#"),
    origin: col(r, "DEPARTMENT"),
    carrier: col(r, "SHIPMENT TYPE"),
    pro: trackingNumber(r, "PRO#", "TRACKING#", "BOL#"),
    units: col(r, "MEMO"),
    destination: "",
    status: classifyStatus(`${col(r, "OVERALL PO STATUS")} ${col(r, "MEMO")}`)
  }, excludedFn));

  mergeOutboundScheduleIdentifiers(shipOut, excludedFn);

  /* TJX/ROSS — grouped layout: the order name appears once and its
     PO lines follow below, so carry the group label forward. */
  let currentOrder = "";
  let currentReceived = "";
  tjxRoss.forEach((r) => {
    const orderName = colLoose(r, "ORDER NAME");
    const received = colLoose(r, "ORDER RECEIVED");
    if (orderName && !/^order\s*name$/i.test(orderName)) currentOrder = orderName;
    if (received && !/^order\s*received$/i.test(received)) currentReceived = received;

    const po = colLoose(r, "PO#", "PO");
    const bol = colLoose(r, "BOL");
    const weight = colLoose(r, "WEIGHT (LBS)", "WEIGHT LBS", "WEIGHT");
    if (!po && !bol && !weight) return;
    if (/^po#?$/i.test(po) || /^bol$/i.test(bol)) return;

    pushOutbound("TJX/ROSS", r, {
      shipDate: colLoose(r, "SHIPOUT DATE", "SHIP OUT DATE") || colLoose(r, "SSD"),
      customer: currentOrder || colLoose(r, "DC#") || "TJX/ROSS",
      invoice: invoiceNumber(r, "PO#", "PO"),
      origin: currentReceived ? `Ordered ${currentReceived}` : "",
      carrier: colLoose(r, "CARRIER"),
      pro: trackingNumber(r, "BOL", "PU#", "ALT. PU# (EG.NRT#)", "ALT PU#", "SHIPMENT #"),
      units: [
        colLoose(r, "PLT") ? `${colLoose(r, "PLT")} Plt` : "",
        colLoose(r, "BOX") ? `${colLoose(r, "BOX")} Box` : ""
      ].filter(Boolean).join(" · "),
      destination: colLoose(r, "DC#"),
      status: classifyStatus(colLoose(r, "STATUS"))
    }, excludedFn);
  });
}

function databaseStatus(value) {
  const label = clean(value).toLowerCase();
  return ({ delivered: "Delivered", received: "Received", completed: "Completed", cancelled: "Cancelled", shipping: "Shipping", exception: "Shipping" })[label] || "Scheduled";
}

function databaseDate(value) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? fmtDate(value) : `${String(date.getMonth() + 1).padStart(2, "0")}/${String(date.getDate()).padStart(2, "0")}/${String(date.getFullYear()).slice(-2)}`;
}

function applyDatabaseShipments(records) {
  inboundRows = [];
  outboundRows = [];
  parcelRows = [];
  records.forEach((record) => {
    const raw = record && typeof record.raw === "object" && record.raw ? record.raw : {};
    const sourceTab = clean(record.source_sheet || raw.source_sheet || raw.sourceSheet || "");
    const sourceRow = Number(record.source_row || raw.source_row || raw.sourceRow || 0) || 0;
    const sourceStatus = clean(record.source_status || raw.source_status || raw.sourceStatus || "");

    const base = {
      databaseId: record.id, databaseVersion: record.version,
      status: databaseStatus(record.status), carrier: clean(record.carrier),
      origin: clean(record.origin), destination: clean(record.destination),
      sourceTab, sourceRow, sourceStatus
    };
    if (record.direction === "outbound") {
      outboundRows.push({
        ...base, source: record.sources?.label || record.sources?.source_key || "Database",
        shipDate: databaseDate(record.scheduled_at), customer: clean(record.customer),
        invoice: clean(record.invoice_number || record.order_number), pro: clean(record.tracking_number),
        units: [record.pallets ? `${record.pallets} Pallets` : "", record.cartons ? `${record.cartons} Cartons` : ""].filter(Boolean).join(" · "),
        qty: clean(record.quantity), length: "", width: "", height: "", weight: clean(record.weight_lbs),
        rate: Number(record.rate || 0)
      });
    } else if (record.direction === "parcel") {
      parcelRows.push({
        ...base, tracking: clean(record.tracking_number), invoice: clean(record.invoice_number),
        eta: databaseDate(record.eta_at), note: clean(record.notes) || "Database synchronized",
        url: parcelTrackingUrl(record.carrier, record.tracking_number)
      });
    } else {
      const tracking = containerTrackingProfile({ SCAC: record.scac, CARRIER: record.carrier, TERMINAL: record.terminal, PORT: record.port }, record.container_number, record.destination);
      inboundRows.push({
        ...base, mode: clean(record.mode) || "Ocean", eta: databaseDate(record.eta_at),
        shipmentNo: clean(record.shipment_number), invoice: clean(record.invoice_number),
        mbl: clean(record.raw?.mbl), hbl: clean(record.raw?.hbl),
        container: clean(record.container_number),
        trackingUrl: tracking.url, trackingSource: tracking.source
      });
    }
  });
}

/* merge same-customer rows shipping within 3 days into one line */
function consolidate(rows) {
  const groups = new Map();
  const loners = [];
  rows.forEach((r) => {
    const key = r.customer.toUpperCase().replace(/\s+/g, " ").trim();
    if (!key) { loners.push(r); return; }
    groups.set(key, [...(groups.get(key) || []), r]);
  });
  const uniq = (vals) => [...new Set(vals.flatMap((v) => clean(v).split(/[\r\n,;]+/)).map((v) => v.trim()).filter(Boolean))];
  const merged = [...groups.values()].flatMap((rowsForCustomer) => {
    const sorted = [...rowsForCustomer].sort((a, b) => (parseDate(a.shipDate)?.getTime() || 0) - (parseDate(b.shipDate)?.getTime() || 0));
    const clusters = [];
    sorted.forEach((r) => {
      const cluster = clusters.at(-1);
      const first = cluster?.[0];
      const near = cluster && parseDate(r.shipDate) && parseDate(first.shipDate) &&
        parseDate(r.shipDate) - parseDate(first.shipDate) <= 3 * 864e5;
      if (near) cluster.push(r); else clusters.push([r]);
    });
    return clusters.map((cluster) => {
      if (cluster.length === 1) return cluster[0];
      const base = cluster[0];
      const dates = cluster.map((r) => r.shipDate).filter(Boolean);
      return {
        ...base,
        source: uniq(cluster.map((r) => r.source)).join(" · "),
        shipDate: dates[0] === dates.at(-1) ? dates[0] : `${dates[0]} – ${dates.at(-1)}`,
        invoice: uniq(cluster.map((r) => r.invoice)).join(" · "),
        pro: uniq(cluster.map((r) => r.pro)).join(" · "),
        carrier: uniq(cluster.map((r) => r.carrier)).join(" · "),
        units: uniq(cluster.map((r) => r.units)).join(" · "),
        destination: uniq(cluster.map((r) => r.destination)).join(" · "),
        rate: cluster.reduce((s, r) => s + r.rate, 0),
        status: cluster.some((r) => r.status === "Shipping") ? "Shipping" : base.status
      };
    });
  });
  return [...loners, ...merged].sort((a, b) =>
    (parseDate(a.shipDate)?.getTime() ?? Number.MAX_SAFE_INTEGER) -
    (parseDate(b.shipDate)?.getTime() ?? Number.MAX_SAFE_INTEGER)
  );
}

/* ---------- KPI block (protected range on All Outbound) ---------- */
async function fetchKpis() {
  const table = await fetchTable(KPI_SOURCE.tab, KPI_SOURCE.range, false);
  const block = {};
  (table.rows || []).forEach((row) => {
    const label = rawCell(row, 0).toUpperCase();
    if (label) block[label] = rawCell(row, 1);
  });
  return block;
}

/* ---------- load pipeline ---------- */
async function load() {
  if (loading) return;
  loading = true;
  $("sync").textContent = "Importing all Logistics Master 2026 sources…";
  try {
    /* FIX: catch attached at creation, not at await — otherwise a fast KPI
       failure fires an unhandledrejection while the main batch is in flight. */
    const kpiPromise = fetchKpis().catch((e) => ({ __error: e }));
    const results = await Promise.allSettled(SOURCES.map((s) => fetchTable(s.tab, s.range, true, s.sheetId, s.gid)));
    const tables = results.map((r) => (r.status === "fulfilled" ? r.value : { cols: [], rows: [] }));
    const mapped = tables.map(objects);

    const [im, tr, ul, ih, b2, whAll, national, shipOut, tjxRoss, exclusions] = mapped;
    const wh = whAll.slice(1); /* WH tab carries a banner row above its headers */

    const exclusionSet = new Set(
      exclusions.filter(useful).map((r) => `${col(r, "SOURCE")}|${col(r, "KEY")}`.trim().toUpperCase())
    );
    /* map display source names back to tab names for exclusion matching */
    const tabOf = {
      "Transfers": "TRANSFERS", "Ulta": "ULTA", "iHerb": "IHERB",
      "B2B/E-com Trucking": "B2B/E-COM TRUCKING", "WH Trucking Request": "WH TRUCKING REQUEST",
      "National Order Progress": "NATIONAL ORDER PROGRESS",
      "Outbound Shipping Schedule": "OUTBOUND SHIPPING SCHEDULE", "TJX/ROSS": "TJX/ROSS"
    };
    const excludedFn = (source, key) =>
      Boolean(key) && exclusionSet.has(`${tabOf[source] || source.toUpperCase()}|${key}`.toUpperCase());

    inboundPlanningRows = mapInboundPlanningGrid(tables[0]);
    inboundRows = mergeInboundPlanning(mapInbound(im), inboundPlanningRows);
    parcelRows = mapParcels(tables[0]);
    mapAllOutbound({ tr, ul, ih, b2, wh, national, shipOut, tjxRoss }, excludedFn);
    if (globalThis.STYLEKOREAN_DATABASE?.preferDatabase && globalThis.StyleKoreanDatabase?.configured()) {
      try {
        const databaseRows = await globalThis.StyleKoreanDatabase.listShipments();
        if (databaseRows.length) applyDatabaseShipments(databaseRows);
      } catch (databaseError) {
        console.warn("Database read unavailable — retaining Google Sheets snapshot.", databaseError);
      }
    }

    /* cost summary: computed from all operational tabs using source-specific invoice/rate rules */
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const validYtdRows = outboundRows.filter((r) => {
      const d = parseDate(r.shipDate);
      return !d || (d <= today && d.getFullYear() === now.getFullYear());
    });
    const validMtdRows = outboundRows.filter((r) => {
      const d = parseDate(r.shipDate);
      return !d || (d <= today && d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth());
    });

    costSummary = {
      ytd: Math.round(validYtdRows.reduce((s, r) => s + Number(r.rate || 0), 0)),
      mtd: Math.round(validMtdRows.reduce((s, r) => s + Number(r.rate || 0), 0)),
      finished: outboundRows.filter((r) => FINISHED.has(r.status)).length,
      kpiSource: "computed"
    };

    let kpiOk = true;
    try {
      const kpi = await kpiPromise;
      if (kpi && kpi.__error) throw kpi.__error;
    } catch (e) {
      kpiOk = false;
      console.warn("KPI block fetch logged.", e);
    }

    /* per-source health for the source strip */
    const contributed = [im, tr, ul, ih, b2, wh, national, shipOut, tjxRoss, exclusions]
      .map((rows) => rows.filter(useful).length);
    const checkedAt = new Date().toISOString();
    sourceHealth = SOURCES.map((s, i) => ({
      id: s.id || s.tab, tab: s.tab, range: s.range, kind: s.kind, gid: s.gid, sheetId: s.sheetId,
      provider: s.provider || "googleSheets", checkedAt,
      ok: results[i].status === "fulfilled",
      rows: contributed[i],
      error: results[i].status === "rejected" ? results[i].reason?.message || "Unavailable" : ""
    }));
    sourceHealth.push({
      id: KPI_SOURCE.id || "outbound-kpis", tab: "Outbound Shipping Schedule KPI block", range: KPI_SOURCE.range,
      kind: "kpi", provider: KPI_SOURCE.provider || "googleSheets", gid: KPI_SOURCE.gid,
      checkedAt, ok: kpiOk, rows: kpiOk ? 4 : 0, error: kpiOk ? "" : "KPI source unavailable"
    });

    const failed = sourceHealth.filter((s) => !s.ok).map((s) => s.tab);
    renderAll();
    $("sync").textContent = failed.length
      ? `${sourceHealth.length - failed.length} of ${sourceHealth.length} sources imported · unavailable: ${failed.join(", ")}`
      : `All ${sourceHealth.length} workbook sources imported`;
    $("dot").classList.toggle("sync-error", failed.length > 0);
    $("setupNotice").classList.toggle("hidden", failed.length < sourceHealth.length);
    $("updated").textContent = new Date().toLocaleString();
  } catch (e) {
    console.error(e);
    $("dot").classList.add("sync-error");
    $("sync").textContent = `Workbook sync issue: ${e.message}`;
    $("setupNotice").classList.remove("hidden");
  } finally {
    loading = false;
  }
}

/* ---------- rendering ---------- */
function srcColor(source) {
  return SOURCE_COLORS[source.split(" · ")[0]] || "var(--steel)";
}
function srcTag(source) {
  return source.split(" · ").map((s) =>
    `<span class="src-tag" style="--c:${srcColor(s)}">${esc(s)}</span>`
  ).join(" ");
}
function statusPill(status) {
  const cls = status.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  return `<span class="status status-${cls}">${esc(status)}</span>`;
}
const activeOutbound = () => outboundRows.filter((r) => !FINISHED.has(r.status));
const activeInbound = () => inboundRows.filter((r) => !FINISHED.has(r.status));

function renderSourceStrip() {
  const ok = sourceHealth.filter((s) => s.ok).length;
  $("sourceOkCount").textContent = `${ok}/${sourceHealth.length} online`;
  $("sourceStrip").innerHTML = sourceHealth.map((s) => {
    const color = {
      "TRANSFERS": "var(--c-transfers)", "ULTA": "var(--c-ulta)", "IHERB": "var(--c-iherb)",
      "B2B/E-COM TRUCKING": "var(--c-b2b)", "WH Trucking Request": "var(--c-wh)",
      "NATIONAL ORDER PROGRESS": "var(--c-national-order)",
      "Outbound Shipping Schedule": "var(--c-ship-out)", "TJX/ROSS": "var(--c-tjx)"
    }[s.tab] || "var(--ink-2)";
    const url = `https://docs.google.com/spreadsheets/d/${s.sheetId || SHEET_ID}/edit#gid=${s.gid}`;
    return `<a class="source-chip ${s.ok ? "" : "failed"}" style="--c:${color}" href="${url}" target="_blank" rel="noopener noreferrer" title="Open ${esc(s.tab)} in Google Sheets" aria-label="Open ${esc(s.tab)} source sheet">
      <span class="st" aria-hidden="true"></span>
      <span><span class="name">${esc(s.tab)}</span><span class="source-open" aria-hidden="true">↗</span><br><span class="rows">${s.ok ? `${s.rows.toLocaleString()} rows` : "unavailable"}</span></span>
    </a>`;
  }).join("");
}

function integrationSnapshot() {
  return {
    platformVersion: PLATFORM.version || "legacy",
    generatedAt: new Date().toISOString(),
    workbook: PLATFORM.workbook?.label || "LOGISTICS MASTER 2026",
    sources: sourceHealth,
    totals: {
      configured: sourceHealth.length,
      online: sourceHealth.filter((source) => source.ok).length,
      records: sourceHealth.reduce((sum, source) => sum + Number(source.rows || 0), 0),
      activeShipments: activeOutbound().length + activeInbound().length + activeParcels().length,
      inbound: activeInbound().length,
      outbound: activeOutbound().length,
      parcels: activeParcels().length
    }
  };
}

function renderIntegrationHealth() {
  const snapshot = integrationSnapshot();
  const summary = [
    ["Sources online", `${snapshot.totals.online}/${snapshot.totals.configured}`],
    ["Source records", snapshot.totals.records.toLocaleString()],
    ["Active shipments", snapshot.totals.activeShipments.toLocaleString()],
    ["Platform", `v${snapshot.platformVersion}`]
  ];
  $("integrationSummary").innerHTML = summary.map(([label, value]) =>
    `<div><span>${esc(label)}</span><strong>${esc(value)}</strong></div>`
  ).join("");
  $("integrationRows").innerHTML = sourceHealth.map((source) => {
    const provider = PLATFORM.providers?.[source.provider]?.label || source.provider || "Google Sheets";
    const url = `https://docs.google.com/spreadsheets/d/${source.sheetId || SHEET_ID}/edit#gid=${source.gid}`;
    return `<tr>
      <td><strong>${esc(source.tab)}</strong>${source.error ? `<small>${esc(source.error)}</small>` : ""}</td>
      <td>${esc(provider)}</td><td>${esc(source.kind)}</td><td class="mono">${esc(source.range || "—")}</td>
      <td class="mono">${Number(source.rows || 0).toLocaleString()}</td>
      <td><span class="integration-state ${source.ok ? "online" : "offline"}">${source.ok ? "Online" : "Offline"}</span></td>
      <td><a class="track-link" href="${url}" target="_blank" rel="noopener noreferrer">Open ↗</a></td>
    </tr>`;
  }).join("");
}

function exportIntegrationHealth() {
  const blob = new Blob([JSON.stringify(integrationSnapshot(), null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `stylekorean-integration-health-${new Date().toISOString().slice(0, 10)}.json`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 5000);
}

function renderMetrics() {
  const active = activeOutbound();
  const scheduledCost = active.reduce((s, r) => s + r.rate, 0);
  const cards = [
    ["Active outbound", active.length.toLocaleString(), "Finished & cancelled excluded", ""],
    ["Finished outbound", costSummary.finished.toLocaleString(), "Shipped · done · received · delivered · cancelled", ""],
    ["Inbound active", activeInbound().length.toLocaleString(), "Ocean + air shipments", ""],
    ["Small parcel", activeParcels().length.toLocaleString(), "Active tracking · delivered and received excluded", ""],
    ["Scheduled outbound cost", `$${Math.round(scheduledCost).toLocaleString()}`, "Active rows with a charge", "cost-scheduled"],
    ["YTD shipping cost", `$${Math.round(costSummary.ytd).toLocaleString()}`, costSummary.kpiSource === "workbook" ? "From protected KPI block" : "Computed from source rows", "cost-ytd"],
    ["MTD shipping cost", `$${Math.round(costSummary.mtd).toLocaleString()}`, costSummary.kpiSource === "workbook" ? "From protected KPI block" : "Computed from source rows", "cost-mtd"],
    ["Ocean containers", inboundRows.filter((r) => r.mode === "Ocean" && !FINISHED.has(r.status)).length.toLocaleString(), "Active container shipments", ""]
  ];
  $("metrics").innerHTML = cards.map(([label, value, , cls]) =>
    `<article class="metric-card ${cls}"><span class="label">${label}</span><strong>${value}</strong></article>`
  ).join("");
}

function next14Days() {
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  return Array.from({ length: 14 }, (_, i) => new Date(+start + i * 864e5));
}
const sameDay = (a, b) => a && b &&
  a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();

function renderBoard(hostId, rows, dateField, itemHtml) {
  const days = next14Days();
  const today = new Date();
  $(hostId).innerHTML = `<div class="board">${days.map((day) => {
    const matches = rows.filter((r) => sameDay(parseDate(r[dateField]), day));
    const isWeekend = day.getDay() === 0 || day.getDay() === 6;
    return `<div class="board-day ${isWeekend ? "weekend" : ""} ${sameDay(day, today) ? "today" : ""}">
      <div class="board-date">${day.toLocaleDateString("en-US", { weekday: "short", month: "2-digit", day: "2-digit" })}</div>
      <div class="board-cell">${matches.length ? matches.map(itemHtml).join("") : '<span class="board-empty" aria-hidden="true">·</span>'}</div>
    </div>`;
  }).join("")}</div>`;
}

function inboundBoardDetail(r) {
  const parts = [];
  if (r.container) {
    parts.push(r.trackingUrl
      ? `<a class="board-link" href="${esc(r.trackingUrl)}" target="_blank" rel="noopener noreferrer" title="${esc(r.trackingSource || "Container tracking")}">${esc(r.container)} ↗</a>`
      : esc(r.container));
  } else if (r.carrier) {
    parts.push(esc(r.carrier));
  }
  if (r.invoice) parts.push(`Invoice # ${esc(r.invoice)}`);
  return parts.join(" · ");
}

function inboundBoardRows() {
  const detailedByContainer = new Map(
    inboundRows
      .filter((row) => row.container)
      .map((row) => [clean(row.container).toUpperCase(), row])
  );
  return inboundPlanningRows.map((row) => {
    const detailed = detailedByContainer.get(clean(row.container).toUpperCase());
    if (!detailed) return row;
    return {
      ...row,
      carrier: detailed.carrier || row.carrier,
      mbl: detailed.mbl || row.mbl,
      hbl: detailed.hbl || row.hbl,
      trackingUrl: detailed.trackingUrl || row.trackingUrl,
      trackingSource: detailed.trackingSource || row.trackingSource
    };
  });
}

function renderBoards() {
  renderBoard("inboundBoard", inboundBoardRows(), "eta", (r) =>
    `<div class="board-item" style="--c:${r.mode === "Air" ? "var(--c-b2b)" : "var(--c-transfers)"}">
      <strong>${esc(r.shipmentNo || r.container || r.mbl || "Shipment")}</strong>
      <span>${inboundBoardDetail(r)}</span>
    </div>`);
  renderBoard("outboundBoard", activeOutbound(), "shipDate", (r) =>
    `<div class="board-item" style="--c:${srcColor(r.source)}">
      <strong>${esc(r.customer || "—")}</strong>
      <span>${esc([
        r.invoice ? `Invoice # ${r.invoice}` : "",
        r.pro ? `PRO / Tracking # ${r.pro}` : ""
      ].filter(Boolean).join(" · ") || r.source)}</span>
    </div>`);
}

function populateFilters() {
  /* FIX: rebuilding options on every auto-refresh silently reset the user's
     source/status selection — now the previous value is restored. */
  const fill = (id, values) => {
    const select = $(id);
    const previous = select.value;
    const first = select.firstElementChild;
    select.innerHTML = "";
    select.append(first);
    const options = [...new Set(values.filter(Boolean))].sort();
    options.forEach((v) => {
      const opt = document.createElement("option");
      opt.textContent = v;
      select.append(opt);
    });
    if (options.includes(previous)) select.value = previous;
  };
  fill("srcFilter", outboundRows.map((r) => r.source.split(" · ")[0]));
  fill("outStatus", outboundRows.map((r) => r.status));
}

/* single pipeline shared by the table render and CSV export */
let sortKey = "shipDate";
let sortDir = 1;
function filteredOutbound() {
  const q = $("outSearch").value.toLowerCase();
  const src = $("srcFilter").value;
  const st = $("outStatus").value;
  const showFinished = $("showFinished").checked;
  const base = outboundRows.filter((r) =>
    (showFinished || !FINISHED.has(r.status)) &&
    (!src || r.source.split(" · ").includes(src)) &&
    (!st || r.status === st) &&
    (!q || Object.values(r).join(" ").toLowerCase().includes(q))
  );
  const rows = consolidate(base);
  const value = (r) =>
    sortKey === "shipDate" ? (parseDate(r.shipDate)?.getTime() ?? Number.MAX_SAFE_INTEGER)
    : sortKey === "rate" ? r.rate
    : String(r[sortKey] || "").toLowerCase();
  return rows.sort((a, b) => {
    const A = value(a), B = value(b);
    return A < B ? -sortDir : A > B ? sortDir : 0;
  });
}

function renderOutbound() {
  const rows = filteredOutbound();
  $("outCount").textContent = `${rows.length.toLocaleString()} rows`;

  const body = $("outRows");
  body.innerHTML = "";
  if (!rows.length) {
    body.innerHTML = `<tr class="empty-row"><td colspan="10">No matching outbound entries.</td></tr>`;
    return;
  }
  const frag = document.createDocumentFragment();
  rows.forEach((r) => {
    const dims = [r.length, r.width, r.height].some(Boolean)
      ? `${esc(r.length || "–")}×${esc(r.width || "–")}×${esc(r.height || "–")}${r.weight ? ` · ${esc(r.weight)}` : ""}`
      : (r.weight ? `${esc(r.weight)} lbs` : "—");
    const trEl = document.createElement("tr");
    if (FINISHED.has(r.status)) trEl.className = "row-finished";
    trEl.innerHTML =
      `<td>${srcTag(r.source)}</td>` +
      `<td class="cell-date">${esc(r.shipDate) || "—"}</td>` +
      `<td><strong>${esc(r.customer) || "—"}</strong><small>Invoice # ${esc(r.invoice) || "—"}</small>` +
        (r.origin ? `<small>${esc(r.origin)}</small>` : "") + `</td>` +
      `<td><strong>${esc(r.carrier) || "—"}</strong><small>PRO / Tracking # ${esc(r.pro) || "—"}</small></td>` +
      `<td>${esc(r.units) || esc(r.qty) || "—"}</td>` +
      `<td class="cell-dims">${dims}</td>` +
      `<td>${esc(r.destination) || "—"}</td>` +
      `<td class="cell-money">${r.rate ? "$" + r.rate.toLocaleString() : "—"}</td>` +
      `<td>${statusPill(r.status)}</td>` +
      `<td>${statusControl(r, "outbound")}</td>`;
    frag.appendChild(trEl);
  });
  body.appendChild(frag);
}


function statusControl(row, kind) {
  const current = clean(row.status) || "Scheduled";
  const values = STATUS_OPTIONS.includes(current) ? STATUS_OPTIONS : [current, ...STATUS_OPTIONS];
  const relation = {
    kind,
    sourceSheet: row.sourceTab || "",
    sourceRow: row.sourceRow || 0,
    shipmentNo: row.shipmentNo || "", container: row.container || "", mbl: row.mbl || "", hbl: row.hbl || "",
    pro: row.pro || "", invoice: row.invoice || "", customer: row.customer || "", shipDate: row.shipDate || "",
    currentStatus: row.sourceStatus || current
  };
  if (row.databaseId) {
    relation.databaseId = row.databaseId;
    relation.databaseVersion = row.databaseVersion;
  }
  const sheetWritable = COMPLETE_ENDPOINT && relation.sourceSheet && (kind !== "inbound" || relation.sourceRow);
  const enabled = Boolean((row.databaseId && globalThis.StyleKoreanDatabase?.configured?.()) || sheetWritable);
  const title = enabled ? "Update status and synchronize linked backends" :
    (relation.sourceSheet ? "Deploy and configure the authenticated Apps Script endpoint to enable edits" : "This row has no writable source relation");
  const options = values.map((value) =>
    '<option value="' + esc(value) + '"' + (value === current ? " selected" : "") + "\">" + esc(value) + "</option>"
  ).join("");
  return '<label class="status-control" title="' + esc(title) + '">' +
    '<span class="visually-hidden">Change status</span>' +
    '<select class="status-select" data-status-relation="' + encodeURIComponent(JSON.stringify(relation)) + '"' +
      (enabled ? "" : " disabled") + "\">" + options + "</select>" +
    '<span class="status-result" aria-live="polite"></span></label>';
}

async function postToAppsScript(payload) {
  const body = JSON.stringify(payload);
  
  // 1. Primary Attempt: standard fetch
  try {
    const res = await fetch(COMPLETE_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body
    });
    if (res.ok) {
      const data = await res.json().catch(() => ({ ok: true }));
      if (data && data.ok === false) throw new Error(data.error || "Update failed");
      return;
    }
  } catch (err) {
    if (err.message && err.message.includes("denied access")) throw err;
  }

  // 2. Secondary Attempt: no-cors fetch
  try {
    await fetch(COMPLETE_ENDPOINT, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain" },
      body
    });
    return;
  } catch (e) {}

  // 3. Ultimate Fallback: Form Submission via hidden iframe (bypasses browser CORS completely)
  return new Promise((resolve) => {
    let iframe = document.getElementById("appsScriptIframe");
    if (!iframe) {
      iframe = document.createElement("iframe");
      iframe.id = "appsScriptIframe";
      iframe.name = "appsScriptIframe";
      iframe.style.display = "none";
      document.body.appendChild(iframe);
    }
    
    const form = document.createElement("form");
    form.method = "POST";
    form.action = COMPLETE_ENDPOINT;
    form.target = "appsScriptIframe";
    
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = "postData";
    input.value = body;
    form.appendChild(input);

    document.body.appendChild(form);
    form.submit();
    setTimeout(() => {
      form.remove();
      resolve();
    }, 1200);
  });
}

async function updateStatus(select) {
  const relation = JSON.parse(decodeURIComponent(select.dataset.statusRelation));
  const result = select.parentElement.querySelector(".status-result");
  const previous = relation.currentStatus || "";
  const canDbWrite = Boolean(relation.databaseId && globalThis.StyleKoreanDatabase?.configured?.());
  const canSheetWrite = Boolean(COMPLETE_ENDPOINT && relation.sourceSheet && (relation.kind !== "inbound" || relation.sourceRow));
  const status = select.value;
  select.disabled = true;
  result.textContent = "Saving…";
  try {
    if (!canDbWrite && !canSheetWrite) throw new Error("No writable backend is configured for this row.");
    if (canDbWrite) {
      const dbStatus = status.toLowerCase().replace(/\s+/g, "_");
      await globalThis.StyleKoreanDatabase.updateShipment(relation.databaseId, relation.databaseVersion, { status: dbStatus });
    }

    if (canSheetWrite) {
      await postToAppsScript({ ...relation, status });
    }

    result.textContent = "Saved";
    await load();
  } catch (error) {
    result.textContent = error.message || "Not saved";
    const fallback = STATUS_OPTIONS.includes(previous) ? previous : "Scheduled";
    select.value = fallback;
    select.disabled = false;
  }
}

/* CSV export of exactly what's on screen (filters + consolidation + sort) */
function exportOutboundCsv() {
  const rows = filteredOutbound();
  const cell = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const csv = [
    ["Source", "Ship date", "Customer", "Invoice #", "Origin", "Carrier", "PRO / Tracking #", "Units",
     "Length (in)", "Width (in)", "Height (in)", "Weight (lbs)", "Destination", "Rate", "Status"].map(cell).join(","),
    ...rows.map((r) => [
      r.source, r.shipDate, r.customer, r.invoice, r.origin, r.carrier, r.pro, r.units || r.qty,
      r.length, r.width, r.height, r.weight, r.destination, r.rate || "", r.status
    ].map(cell).join(","))
  ].join("\r\n");
  const stamp = new Date().toISOString().slice(0, 10);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" }));
  a.download = `stylekorean-outbound-${stamp}.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

function renderInbound() {
  const q = $("inSearch").value.toLowerCase();
  const mode = $("modeFilter").value;
  const showFinished = Boolean($("showInboundFinished")?.checked);
  const rows = inboundRows
    .filter((r) => (showFinished || Boolean(q) || !FINISHED.has(r.status)) &&
      (!mode || r.mode === mode) &&
      (!q || Object.values(r).join(" ").toLowerCase().includes(q)))
    .sort((a, b) =>
      (parseDate(a.eta)?.getTime() ?? Number.MAX_SAFE_INTEGER) -
      (parseDate(b.eta)?.getTime() ?? Number.MAX_SAFE_INTEGER));
  $("inCount").textContent = `${rows.length.toLocaleString()} rows`;

  const body = $("inRows");
  body.innerHTML = "";
  if (!rows.length) {
    body.innerHTML = `<tr class="empty-row"><td colspan="10">No matching inbound shipments.</td></tr>`;
    return;
  }
  const frag = document.createDocumentFragment();
  rows.forEach((r) => {
    const trEl = document.createElement("tr");
    trEl.innerHTML =
      `<td><span class="mode-tag mode-${r.mode.toLowerCase().replace(/\s+/g, "-")}">${esc(r.mode)}</span></td>` +
      `<td class="cell-date">${esc(r.eta) || "—"}</td>` +
      `<td><strong>${esc(r.shipmentNo) || "—"}</strong></td>` +
      `<td><span class="mono">${esc(r.invoice) || "—"}</span></td>` +
      `<td>${r.container
        ? (r.trackingUrl
          ? `<a class="track-link" href="${esc(r.trackingUrl)}" target="_blank" rel="noreferrer" title="${esc(r.trackingSource || "Container tracking")}">${esc(r.container)} ↗</a>`
          : `<span class="mono">${esc(r.container)}</span>`) +
          (r.trackingSource ? `<small>${esc(r.trackingSource)}</small>` : "")
        : "—"}</td>` +
      `<td><small style="margin:0">${esc(r.mbl) || "—"}</small><small>${esc(r.hbl)}</small></td>` +
      `<td>${esc(r.carrier) || "—"}</td>` +
      `<td>${esc(r.origin) || "—"}</td>` +
      `<td>${esc(r.destination) || "—"}</td>` +
      `<td>${statusControl(r, "inbound")}</td>`;
    frag.appendChild(trEl);
  });
  body.appendChild(frag);
}

function renderParcels() {
  const active = activeParcels();
  const shown = active.slice(0, 24);
  $("parcelCount").textContent = `${active.length.toLocaleString()} active`;
  const hidden = active.length - shown.length;
  $("parcelGrid").innerHTML = shown.length ? shown.map((p) => `
    <article class="parcel-card">
      <div class="parcel-top">
        <span class="carrier-logo carrier-${p.carrier.toLowerCase()}">${esc(p.carrier)}</span>
        ${statusPill(p.status)}
      </div>
      <strong>${esc(p.tracking)}</strong>
      <p>${esc(p.invoice ? `Invoice # ${p.invoice} · ` : "")}${esc(p.note)}</p>
      <div class="parcel-bottom">
        <span class="parcel-eta">${p.eta ? "ETA " + esc(p.eta) : "ETA —"}</span>
        ${p.url ? `<a class="track-link" href="${esc(p.url)}" target="_blank" rel="noreferrer">Track ↗</a>` : ""}
      </div>
    </article>`).join("") +
    (hidden > 0 ? `<p class="parcel-more">+ ${hidden} more parcel${hidden === 1 ? "" : "s"} in the IMPORTS tab — open the Google Sheet for the full list.</p>` : "")
    : `<p style="color:var(--steel);padding:6px 0 14px;">No active small-parcel shipments. Delivered and received parcels are excluded.</p>`;
}

function renderAll() {
  renderMetrics();
  renderSourceStrip();
  renderIntegrationHealth();
  renderBoards();
  populateFilters();
  renderOutbound();
  renderInbound();
  renderParcels();
}

/* ---------- events & boot ---------- */
function debounce(fn, ms) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

function setupCollapsiblePanel(buttonId, bodyId) {
  const button = $(buttonId);
  const body = $(bodyId);
  if (!button || !body) return;
  const sync = () => {
    const expanded = !body.hidden;
    button.setAttribute("aria-expanded", String(expanded));
    button.textContent = expanded ? "Hide details" : "Show details";
  };
  button.addEventListener("click", () => {
    body.hidden = !body.hidden;
    sync();
  });
  sync();
}

document.addEventListener("DOMContentLoaded", () => {
  setupCollapsiblePanel("toggleIntegration", "sourceHubBody");
  $("refresh").addEventListener("click", () => load());
  $("outSearch").addEventListener("input", debounce(renderOutbound, 120));
  $("srcFilter").addEventListener("change", renderOutbound);
  $("outStatus").addEventListener("change", renderOutbound);
  $("showFinished").addEventListener("change", renderOutbound);
  $("inSearch").addEventListener("input", debounce(renderInbound, 120));
  $("modeFilter").addEventListener("change", renderInbound);
  const showInboundFinishedToggle = $("showInboundFinished");
  if (showInboundFinishedToggle) showInboundFinishedToggle.addEventListener("change", renderInbound);
  $("exportCsv").addEventListener("click", exportOutboundCsv);
  const integrationExportButton = $("exportIntegration");
  if (integrationExportButton) integrationExportButton.addEventListener("click", exportIntegrationHealth);
  const statusChange = (event) => {
    if (event.target.matches(".status-select")) updateStatus(event.target);
  };
  $("outRows").addEventListener("change", statusChange);
  $("inRows").addEventListener("change", statusChange);

  /* sortable outbound columns */
  const sortHeaders = [...document.querySelectorAll("#outTable th[data-sort]")];
  sortHeaders.forEach((th) => th.addEventListener("click", () => {
    const key = th.dataset.sort;
    if (sortKey === key) sortDir = -sortDir; else { sortKey = key; sortDir = 1; }
    sortHeaders.forEach((h) => {
      const active = h === th;
      h.classList.toggle("sorted", active);
      h.classList.toggle("desc", active && sortDir < 0);
      h.setAttribute("aria-sort", active ? (sortDir > 0 ? "ascending" : "descending") : "none");
    });
    renderOutbound();
  }));

  load();
  setInterval(() => { if (!document.hidden) load(); }, AUTO_REFRESH_MS);
});
