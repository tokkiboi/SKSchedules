# StyleKorean Logistics Hub — System Blueprint

Production architecture for the SKSchedules web application, its Google Workspace backend, and the automated email → Drive → Sheets → Website pipeline.

Repo: `tokkiboi/SKSchedules` · Site: GitHub Pages (custom domain via `CNAME`) · Owner mailbox: `alex@stylekoreanus.com`

---

## 1. System architecture & data flow

```
                         ┌──────────────────────────────────────────────┐
                         │                 GMAIL (alex@)                │
                         │  출고 / 해상 / 항공 / arrival notice / BOL /   │
                         │  entry summary / shipping docs (attachments) │
                         └──────────────┬───────────────────────────────┘
                                        │ every 15 min (GmailPipeline.gs)
                        ┌───────────────┼──────────────────┐
                        ▼               ▼                  ▼
              ┌──────────────┐  ┌───────────────┐  ┌───────────────────┐
              │ GOOGLE DRIVE │  │  VALIDATION    │  │  PARSERS          │
              │ SK Logistics │  │  (Validation.gs│  │  CSV / XLSX (Drive │
              │ Email Archive│  │  scoring)      │  │  convert) / PDF OCR│
              │ YYYY/MM/Cat. │  └──────┬────────┘  └───────────────────┘
              └──────────────┘         │ pass            │ fail
                                       ▼                 ▼
        ┌──────────────────────────────────────┐  ┌─────────────────────┐
        │      LOGISTICS MASTER 2026           │  │ PENDING VERIFICATION │
        │  (1M-vZ24Yw4ZN7R7b_473cVn8kny8Dzn…)  │  │ tab (yellow rows,    │
        │  IMPORTS (inbound) · WH Trucking     │◀─│ APPROVED → commit    │
        │  Request + 8 outbound tabs           │  │ every 30 min)        │
        │  INVENTORY · KPI DASHBOARD ·         │  └─────────────────────┘
        │  PIPELINE LOG                        │
        └────────▲──────────────▲──────────────┘
                 │ hourly       │ 30 min (existing scanner in Code.gs)
   ┌─────────────┴───┐   ┌──────┴──────────────┐
   │ ALLOCATION WB   │   │ WMS workbooks       │
   │ 17e5EYNMr…      │   │ 1tNBa7c78… (stock,  │
   │ per-shipment SKU│   │ container log,      │
   │ channel alloc.  │   │ putaway) · 14lH9SQ… │
   └─────────────────┘   │ (invoice & issues)  │
                         └─────────────────────┘

        LOGISTICS MASTER 2026 ──(public gviz JSON, no key needed)──▶ WEB FRONTEND
                                                                    (GitHub Pages)
        GitHub repo main ──push──▶ GitHub Actions deploy-pages.yml ──▶ Pages CDN
        Apps Script ──repository_dispatch "sheet-data-changed"──▶ same workflow
```

Key properties:

- **Reads are serverless.** The browser pulls sheet data directly from Google's gviz endpoint at page load and every 30 minutes — no rebuild is needed for data changes. Redeploys refresh statically-baked content and CDN caches only.
- **Writes are guarded.** Status edits go through the domain-restricted Apps Script web app (`doPost` in `Code.gs`); email-derived rows go through validation, and anything ambiguous stops in PENDING VERIFICATION.
- **Single writer.** All mutation happens inside one Apps Script project under `LockService`, so concurrent triggers can't corrupt the workbook.

## 2. Apps Script modules (in `google-apps-script/`)

| File | Purpose | Trigger |
|---|---|---|
| `Code.gs` | Status write-back web app (`doPost`), WMS trucking scanner, dropdown provisioning. **Fixed:** the file previously contained two full concatenated copies (duplicate `const SPREADSHEET_ID` — a load-time syntax error); only the newer copy remains. | Web app + 30 min |
| `GmailPipeline.gs` | Gmail scan → Drive archive → parse (CSV / XLSX via Drive convert / PDF via Drive OCR) → validate → upsert into IMPORTS / WH Trucking Request. Labels threads `sk-logistics/processed·pending-verification·error`. Auto-committed rows are tinted light blue. | 15 min |
| `Validation.gs` | Field/date/format validation, PENDING VERIFICATION sheet (yellow = needs review), `processApprovedPending()` commits human-approved rows using manually corrected cell values. | 30 min |
| `InventorySync.gs` | Reads WMS stock snapshot + container log and the allocation workbook; writes `INVENTORY` and `KPI DASHBOARD` tabs; `enrichImportsFromContainerLog()` stamps receiving/QC dates onto IMPORTS rows. | 60 min / daily |
| `Triggers.gs` | `setupAllTriggers()` provisions every schedule; `requestSiteRedeploy()` fires the GitHub `repository_dispatch`. | daily |

### Gmail filter

```
has:attachment newer_than:7d -label:sk-logistics/processed -label:sk-logistics/error
{subject:출고 subject:해상 subject:항공 subject:선적 subject:입고
 subject:"arrival notice" subject:"bill of lading" subject:BOL
 subject:"entry summary" subject:"shipping documents" subject:ISF
 subject:"delivery order" subject:POD}
```

Adjust in `GMAIL_PIPELINE.query`. Add `from:` clauses per carrier/forwarder as they become known — sender filters are far more precise than subject filters.

### Drive layout

`SK Logistics Email Archive / <YYYY> / <MM> / {Arrival Notices, Bills of Lading, Entry Summaries, WMS Exports, Shipping Documents, Other}` — every raw attachment is stamped `yyyyMMdd-HHmm <original name>`.

### Validation rules (auto-commit requires all)

- Inbound: at least one of B/L·container·invoice/entry no.; an ETA or ship date inside −45/+400 days; container numbers must be ISO format `AAAA9999999`.
- Outbound: customer + ship date + (invoice or PRO/BOL); same date window.
- Quantities numeric; unparseable attachments and PDFs without reliable identifiers always go to PENDING VERIFICATION.
- Duplicate protection: inbound matched on B/L or container; outbound on invoice or customer+ship-date (multi-invoice cells are split on newlines/commas, same as the existing scanner).

### Verification workflow

1. Row lands in **PENDING VERIFICATION** (yellow), with issues, source-email link, Drive file link, and raw JSON.
2. A human corrects any cell (Customer, Invoice, BL/PRO, Container, dates, Qty) and sets **Status = APPROVED** (dropdown).
3. Within 30 minutes `processApprovedPending()` commits it through the same upsert path (corrected cells win over the original extraction) and recolors the row blue (COMMITTED). REJECTED rows grey out and stay as an audit trail.

## 3. Inventory relational model

Join keys discovered from the two raw workbooks:

| Key | Source | Role |
|---|---|---|
| `SKU` / 상품코드 | every tab in both workbooks | primary product key (~3,700 SKUs) |
| Barcode / UPC | allocation `Barcode`, WMS `Product BarCode`/`UPC` | secondary product match |
| 차수 (shipment code: `TW n`, `HJ n`, `ES n`, `OSL n`, `ER n`, `MCI n`) | WMS container log ↔ allocation workbook tabs | container-level join |
| `PC########` | WMS container log | container id ↔ IMPORTS rows |
| `PI########` | WMS putaway tabs | purchase-invoice join to inbound schedule |
| 입고일 / 검수 완료일 | WMS container log | receiving + QC milestone dates |

Derived `INVENTORY` tab (one row per SKU):

```
On Hand (Actual) / Available / On Hold   ← WMS live stock snapshot (summed across locations)
Incoming (Confirmed) / Remaining To Receive ← Σ allocation tabs (Cnfm Qty, 잔여수량)
Inbound Shipments (차수)                  ← allocation tabs with remaining > 0,
                                            annotated "(rcvd <입고일>)" from the container log
Channel Allocation                        ← CAWH / iHerb / National / BK / US_Official / Moida / NY
Flag                                      ← OUT OF STOCK · LOW STOCK · LOW — INBOUND EN ROUTE
```

`KPI DASHBOARD` tab metrics: SKUs tracked, units on hand/available, units inbound remaining, low/out-of-stock SKUs, containers in transit (no 입고일), containers awaiting QC, pending-verification count. The website reads both tabs live; `inventory.html` renders them with search, flag filter, and column sorting.

Scale note: the allocation workbook has ~190 tabs. `readAllocationIncoming_()` respects a 4.5-minute budget and marks the run `PARTIAL` in PIPELINE LOG if it stops early; hourly runs converge quickly. If the workbook keeps growing, archive fully-received shipment tabs to a second workbook.

## 4. Deployment strategy

**Recommendation: stay on GitHub Pages.** The site is a static export with no build step, Pages is already live behind the custom domain, and all dynamic data arrives client-side from Google Sheets. Vercel/Netlify would only pay off if you later need server-side code (e.g., replacing the lost `/api/status` endpoint with something beyond Apps Script).

`.github/workflows/deploy-pages.yml` (replaces the Jekyll workflow):

- `push` to `main` → sanity-check JS (`node --check`) → upload repo root → deploy.
- `repository_dispatch: sheet-data-changed` → same deploy, fired daily (or on demand) by `requestSiteRedeploy()` in Apps Script.
- `workflow_dispatch` for manual runs.

Repo settings required once: **Settings → Pages → Source: GitHub Actions.**

## 5. Setup runbook

1. **Apps Script** — open the existing `StyleKorean Logistics Sync` project; replace `Code.gs`; add `GmailPipeline.gs`, `Validation.gs`, `InventorySync.gs`, `Triggers.gs`.
2. **Enable the Advanced Drive Service** (Editor → Services → Drive API, identifier `Drive`) — required for XLSX conversion and PDF OCR.
3. **Script properties** — set `GITHUB_TOKEN` (fine-grained PAT for `tokkiboi/SKSchedules`, Contents: read-write). Optional `GITHUB_REPO`.
4. Run `setupAllTriggers()` once and grant the Gmail/Drive/Sheets scopes.
5. Run `syncInventoryModule()` once manually; confirm `INVENTORY` and `KPI DASHBOARD` tabs appear in LOGISTICS MASTER 2026.
6. **Repo** — commit this branch (`git pull` first; the local checkout may be behind `origin/main`), confirm Pages source is *GitHub Actions*, delete `jekyll-gh-pages.yml` if it still exists on remote.
7. Verify `https://<domain>/inventory.html` loads and the INVENTORY nav button appears on the board.

## 6. Security notes

- The master workbook is readable via gviz without auth — that is what makes the site serverless. Keep genuinely sensitive data (costs beyond what the KPI block already exposes, supplier pricing) out of the public tabs, or move reads behind an Apps Script proxy later.
- Keep the `doPost` web app restricted to `@stylekoreanus.com` (as `site-config.js` assumes). Never deploy it as "Anyone".
- The GitHub PAT lives only in Script Properties — never in the repo.
- `database-config.js` (Supabase) stays disabled; if enabled later, publishable key + RLS only, per the comment already in that file.

## 7. Known limitations / next steps

- PDF extraction is OCR + regex — good for arrival notices and BOLs with standard layouts; oddly formatted docs will land in PENDING VERIFICATION by design.
- Gmail quotas: MailApp/Gmail read quotas are generous, but `maxThreadsPerRun` is capped at 20/run to stay within the 6-minute execution limit.
- The allocation workbook's tab names are the shipment identifiers; renaming tabs changes 차수 join behavior.
- Consider adding a weekly Drive-folder digest email (files archived, rows committed, rows pending) — `PIPELINE LOG` already collects the data.
