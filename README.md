# Quantum Matrix: AGLxAR Solution

**Live app:** https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app  
**Password:** `robotics2026`  
**Carrier portal:** https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app/vendor_upload  
**Owner:** Dominique Kennedy (kennewdo) — Amazon Global Logistics  
**Repo:** domkennedy5/robotics-container-tracker (Streamlit Cloud, auto-deploy on push)

---

## What This Does

Single-app command center for the Amazon Robotics Dray Program. Three primary functions:

1. **Weekly Delivery Planning** — build and distribute the carrier delivery schedule every Friday
2. **WBR Generation** — auto-build the Monday WBR slide and Perjen-style bridge submitted to NA Destination Ops leadership
3. **Container Lifecycle Tracking** — port → yard → FC → empty return visibility across all active carriers

All data persists to SQLite (`tracker.db`) synced to S3 after every write. No local-only state.

---

## Tab Reference

### Tab 1 — WBR

Four sub-tabs: **Summary · Trends · Build · History**

**Build** is the primary sub-tab. Upload 3 required source files to generate the weekly WBR slide (PDF + PPTX) and Perjen-style bridge in one click.

| File | Where to pull |
|---|---|
| GVT Data (`GVT Data WK##.xlsx`) | GM DCM Reports → Inbound Container Milestone (Customer=AMZ, Equip=AMAZON Robotics) |
| OBLT Data | Ocean Bridge Logistics Tracking |
| Inbound Loads | Amazon Robotics Inbound Loads Report (run Monday morning) |

**Submission SOP**

| Field | Value |
|---|---|
| To | `doc+destops-36@fusion.amazon.dev` |
| Subject | `NA Destination Ops WBR_Robotics` |
| Deadline | Monday by 2:00 PM CT |
| Attachment | `GLS_Robotics_YYYY-M-D.pdf` |
| Fusion doc set | https://fusion.amazon.dev/documentset/DOCUMENTSET%23f42d63bd-a96a-4f29-a356-a66ac18602a9 |

**SLA Targets**

| Metric | Target |
|---|---|
| AV→OA transit | ≤ 3 days |
| OA→Del transit (BOS) | ≤ 3 days |
| Empty→Term return | ≤ 3 days |
| On-Time to Promise (OTP) | ≥ 95% |

**Summary** shows the 6-week trend table and current-week metrics.  
**Trends** shows the 6-chart slide preview (AV→OA, OA→Del, OTP, Volume, E2E, carrier scorecard).  
**History** shows all past generated WBRs; re-download any prior week.

If prior weeks are missing, the ⚠️ warning expander lets you enter historical W-5 through W-1 values manually.

---

### Tab 2 — Business Review

Aggregates WBR weekly data into monthly, quarterly, or annual rollups. Requires WBR data generated in the WBR tab first.

- **Cadence selector**: Week / Month / Quarter / Year
- Shows volume totals, SLA performance, and WoW delta vs. prior period
- All charts built from `wbr_results` DB table — no re-upload needed

---

### Tab 3 — Planning (Delivery Plan Scheduler)

Nine sub-tabs: **Plan Builder · All Sites · By Site · Carrier View · WoW / History · SRF · Config · Import History · SOP Guide**

**Owner:** Dominique Kennedy | **Frequency:** Weekly, every Friday | **Deadline:** 3:00 PM ET

| Sub-tab | What it does |
|---|---|
| Plan Builder | Add containers (single or bulk paste), parse stakeholder requests, level-load across the week |
| All Sites | Full-week compiled view across all sites — Excel export |
| By Site | Per-site view with carrier badges, daily Slack notification generator, mid-week adjustment tool |
| Carrier View | Shareable carrier-specific schedule + pre-send validation checklist + Excel export |
| WoW / History | Week-over-week delta (new / rolled / dropped) + full history search |
| SRF | Site Receiving Form generator |
| Config | Sites table, carriers table, site–carrier map — all editable in-app |
| Import History | Upload ToteASERs Robotics DBR Tracker.xlsx to backfill container history |
| SOP Guide | Full step-by-step SOP (Steps 1–9), GVT glossary, file naming reference |

**RIC6 Receiving Constraints**

| Parameter | Value |
|---|---|
| Receiving window | 7:30 AM – 4:30 PM |
| Lunch (no deliveries) | 12:00 PM – 1:00 PM |
| Last arrival | 3:30 PM |
| Saturday | 5 slots only — 7:30–11:30 AM, hard cap |
| Weekly cap | 115 total loads across all programs |
| HDDR preference | Before lunch (7:30–11:30 AM), one slot per day |

**Bulk Paste format** (Plan Builder → Step 2):
```
TCNU3773041  HDDR  RIC6
MRKU4103422  ATMI  ILM1
CSNU8812340  ARVY  DBM6
```
Carrier and Site are optional if defaults are set. Level-load checkbox distributes evenly across active days.

---

### Tab 4 — Insights

Three sub-tabs: **Overview · Empty Returns · Lane Costs**

**Overview** — weekly operations briefing auto-generated from all loaded data sources.
- DBR Snapshot: delivery count, empty return overdue/due-soon, on-vessel count, demurrage holds
- Detention & Demurrage Risk: flags containers approaching LFD/demurrage dates (from Inbound Loads)
- Carrier Submission Activity: recent submission log by carrier

**Empty Returns** — overdue / due-soon / on-track flag view from the DBR Empty Returns sheet.  
**Lane Costs** — drayage rate lane management by SCAC + port + destination. Cost comparison simulator.

---

### Tab 5 — Container Lookup

Upload the weekly carrier DBR Excel → paste any container IDs → instant cross-sheet results.

- Searches: Delivery Appointments, Empty Returns, On Vessel, Canceled, Demurrage, Accessorials
- Check-digit-tolerant matching (dashes optional; TCNU389902-4 and TCNU389902 both match)
- Download results as Excel

**How to use:**
1. Upload the weekly carrier DBR in the **sidebar**
2. Paste container IDs (one per line) in the search box
3. Click Search → results show sheet, status, key detail for each container

---

### Tab 6 — Carriers

Two sub-tabs: **Submit · Data**

**Submit** — the primary intake point for carrier DBR files.
- Share the vendor portal link with carriers for direct self-submission
- Or upload files manually here (supports AGL standard template + ARVY/HUDD/ATMI legacy formats)
- **DBR Receipt Tracker**: compliance grid showing which carriers submitted each day of the selected week; missing submissions flagged automatically
- **Send Reminder**: SES email to carrier contact for any missing submission
- Download the AGL Carrier Template

**Data** — structured view of all raw carrier submissions stored in DB.
- Filter by carrier and sheet type (Delivery, Empty Returns, Demurrage, Accessorials, ODY)
- Shows latest status per container (duplicates collapsed; full history retained)
- **Backfill** expander: upload the ToteASERs Robotics DBR Tracker.xlsx to populate `inbound_containers` from the AGL-side Delivery Plan sheet

**Carrier Portal URL:**  
`https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app/vendor_upload`  
Carriers authenticate with a per-carrier password set in DBR Dashboard → By Carrier → Carrier Admin.

---

### Tab 7 — DBR Dashboard

Live container lifecycle view sourced from `inbound_containers` table. Updated when carriers submit their daily DBR or when a backfill is run.

**Global filters** (apply across all sub-tabs): Carrier, FC Destination, Status, Date Range (FC Sched Del)

**KPI row:** Total Containers · At Yard · Delivered to FC · Pending FC Delivery · Awaiting Empty Return · Empty Returned

Four sub-tabs:

| Sub-tab | What it shows |
|---|---|
| Today / Upcoming | Containers with FC Sched Del = today or tomorrow (excludes delivered/returned) |
| Late / At Risk | Containers past their FC Sched Del that are not yet delivered — sorted by days late |
| Empty Returns | Containers with FC Act Del recorded but no Empty Returned to Port date — sorted by days at FC |
| By Carrier | Full Delivery Plan column set for the selected carrier; on-time %; last DBR receipt date; Carrier Admin (contact info + portal password) |

Column labels match the Delivery Plan sheet exactly: Container #, Port/Region, Yard Sched Del, Yard Act Del, FC Destination, FC Sched Del, FC Act Del, Empty Returned to Port, Date Received, Live/Drop.

---

### Tab 8 — Inbound Forecast

Three panels: **Pipeline Forecast · Allocation Manager · Simulator**

Powered by the `ar_inbound_unified` table — upload the ARVY Inbound report or equivalent to populate.

**Pipeline Forecast** — weekly inbound container volume forecast by FC and carrier, with ocean ETA visibility.  
**Allocation Manager** — set per-carrier allocation targets by site; track actual vs. target.  
**Simulator** — model load distribution changes (e.g. shifting allocation % between carriers) and see capacity impact.

---

## Carriers

| SCAC | Carrier | Ops Contact | Email |
|---|---|---|---|
| ATMI | Cargomatic | Tyler Domingues | tdomingues@cargomatic.com |
| ARVY | Arrive Logistics | Tyler Spangler | tspangler@arrivelogistics.com |
| HDDR – RIC6 | Maersk | Sandji Ruffin | sandji.ruffin@maersk.com |
| HDDR – ILM1 | Maersk | Jerry Nesbit | jerry.nesbit@maersk.com |
| HDDR – LAX | Maersk | Desirae Swain / Ailua Osoimalo | desirae.swain@maersk.com |
| RKNE | RoadOne | Mark Brennan | — |
| TGHE | Tighe | TBD | — |

**DBR email subjects to expect Thursday EOD:**
```
ATMI  : [EXTERNAL] Amazon Robotics - DBR [DATE]
ARVY  : [EXTERNAL] Robotics DBR [DATE]
HUDD RIC6 : [EXTERNAL] RIC6 Delivery Plan Update (HUDD)
HUDD ILM1 : [EXTERNAL] ILM1 Delivery Plan Update / HUDD
HUDD LAX  : RE: DBR Bridges Report - HUDD - [DATE]
```

---

## File Naming Conventions

| File | Convention | Example |
|---|---|---|
| Carrier DBR | `[CARRIER] DBR M.DD.YY.xlsx` | `ATMI DBR 7.17.26.xlsx` |
| HUDD site-specific | `HUDD [SITE] DBR M.DD.YY.xlsx` | `HUDD RIC6 DBR 7.17.26.xlsx` |
| Delivery Plan export | `[SITE] Delivery Plan M.DD-M.DD.xlsx` | `RIC6 Delivery Plan 7.21-7.25.xlsx` |
| GVT export | `GVT Data WK##.xlsx` | `GVT Data WK30.xlsx` |
| WBR PDF | `GLS_Robotics_YYYY-M-D.pdf` | `GLS_Robotics_2026-7-28.pdf` |

---

## Data Architecture

```
app.py                Main Streamlit app (8 tabs)
wbr_engine.py         WBR data parsing + metrics computation
wbr_pdf.py            PDF slide generator (pixel-faithful to gold standard)
wbr_pptx.py           PPTX converter
inbound_forecast.py   Inbound Forecast tab (Pipeline / Allocation / Simulator)
data_sync.py          S3 sync — pull on startup, push after every write
utils.py              Container ID normalization + carrier file parsing
lambda/               Email ingest + alert Lambdas (deploy separately)
deploy/SETUP.md       Lambda + SES deployment guide
tracker.db            SQLite — local + S3-synced
```

**Key DB tables:**

| Table | Purpose |
|---|---|
| `delivery_plan` | Container delivery plan entries (Planning tab) |
| `plan_sites` | Site config (capacity, port, constraints, site type) |
| `plan_carriers` | Carrier config |
| `plan_site_carrier` | Site–carrier mapping (priority time, allocation %) |
| `plan_week_config` | Per-week active receiving days |
| `wbr_results` | Historical WBR metrics by week (feeds WBR + Business Review) |
| `wbr_context_notes` | Operational notes per week (auto-injected into bridge) |
| `carrier_submissions` | All raw carrier DBR file rows |
| `dbr_receipts` | DBR receipt tracking per carrier per week |
| `inbound_containers` | Container lifecycle: port → yard → FC → empty return (feeds DBR Dashboard) |
| `carrier_contacts` | Carrier contact info + portal passwords + reminder toggle |
| `ar_inbound_unified` | Inbound Forecast source data |
| `lookup_log` | Container Lookup search history |

**Infrastructure:**

| Resource | Value |
|---|---|
| S3 bucket | `robotics-container-tracker` (us-east-1, account 844000647671) |
| IAM user | `robotics-tracker-reader` — read-only S3 for Streamlit Cloud |
| Streamlit Cloud | Auto-redeploys on `git push` to `main` (~60s) |
| Secrets | AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, APP_PASSWORD in Streamlit Cloud secrets |

---

## Local Development

```bash
pip install -r requirements.txt
# double-click launch.bat  OR:
venv\Scripts\streamlit run app.py
# password: robotics2026
```

DB syncs from S3 on startup if AWS credentials are set in `.streamlit/secrets.toml`.
