/**
 * Triggers.gs — one-click provisioning of every time-driven job,
 * plus the GitHub Pages redeploy hook.
 *
 * Run setupAllTriggers() once (and again after changing schedules).
 *
 * Script Properties (File > Project properties > Script properties):
 *   GITHUB_TOKEN  — fine-grained PAT, repo tokkiboi/SKSchedules,
 *                   permission: Contents: Read & write (for repository_dispatch)
 *   GITHUB_REPO   — optional override, default "tokkiboi/SKSchedules"
 */

/* eslint-disable no-unused-vars */

var TRIGGER_PLAN = [
  { handler: "processLogisticsEmails", minutes: 15 },          // Gmail ingestion
  { handler: "processApprovedPending", minutes: 30 },          // commit human-approved rows
  { handler: "scanAndImportWmsTruckingOrders", minutes: 30 },  // existing WMS trucking scanner (Code.gs)
  { handler: "syncInventoryModule", minutes: 60 },             // inventory + KPI rebuild
  { handler: "enrichImportsFromContainerLog", daily: 6 },      // 6 AM daily
  { handler: "requestSiteRedeploy", daily: 7 }                 // 7 AM daily safety redeploy
];

function setupAllTriggers() {
  var handlers = TRIGGER_PLAN.map(function (t) { return t.handler; });
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (handlers.indexOf(trigger.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(trigger);
  });

  TRIGGER_PLAN.forEach(function (t) {
    var builder = ScriptApp.newTrigger(t.handler).timeBased();
    if (t.minutes) builder.everyMinutes(t.minutes).create();
    else builder.everyDays(1).atHour(t.daily).create();
  });

  Logger.log("Provisioned " + TRIGGER_PLAN.length + " triggers.");
  return TRIGGER_PLAN;
}

/**
 * Fires a repository_dispatch event so GitHub Actions redeploys the site.
 * The frontend reads sheet data live at runtime, so this is only needed to
 * refresh statically-baked content and to keep Pages caches warm.
 */
function requestSiteRedeploy() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("GITHUB_TOKEN");
  if (!token) { Logger.log("GITHUB_TOKEN script property not set — skipping redeploy."); return; }
  var repo = props.getProperty("GITHUB_REPO") || "tokkiboi/SKSchedules";

  var response = UrlFetchApp.fetch("https://api.github.com/repos/" + repo + "/dispatches", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    payload: JSON.stringify({ event_type: "sheet-data-changed", client_payload: { at: new Date().toISOString() } }),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code >= 300) throw new Error("repository_dispatch failed: HTTP " + code + " " + response.getContentText().slice(0, 200));
  Logger.log("Redeploy requested (HTTP " + code + ").");
}
