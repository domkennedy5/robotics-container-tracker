# AGL Robotics Container Tracker

**Live app:** https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app  
**Password:** `robotics2026`  
**Owner:** Dominique Kennedy (kennewdo) — Amazon Global Logistics

---

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
# password: robotics2026
```

---

## What This Does

Single-app command center for the Amazon Robotics Dray Program. Two primary workflows:

1. **Weekly Delivery Planning** — build, validate, and distribute the carrier delivery schedule every Friday
2. **WBR Generation** — auto-build the Monday WBR slide and bridge submitted to NA Destination Ops leadership

All data persists to SQLite (`tracker.db`) synced to S3 after every write. No local-only state.

---

## Tab Reference

### Tab 1 — Container Lookup
Upload the weekly DBR Excel → paste any container IDs (one per line) → instant cross-sheet results.  
Searches: Delivery Appointments, Empty Returns, On Vessel, Canceled, Demurrage, Accessorials.  
Check-digit-tolerant matching (dashes optional). Download results as Excel.

**How to use:**
1. Upload the weekly DBR in the **sidebar**
2. Paste container IDs (one per line) in the search box
3. Click Search → results show which sheet each container appears on + status

---

### Tab 2 — Carrier Submission
Carriers submit container status updates directly, or upload an ARVY-style Excel template.  
All submissions log to DB with timestamp and are searchable/filterable/exportable.

**DBR Receipt Tracker** — tracks whether each carrier (ATMI, ARVY, HDDR, RKNE, TGHE) has submitted their weekly DBR.  
Missing submissions are auto-flagged. Reminder messages are auto-drafted.

---

### Tab 3 — Empty Returns Dashboard
Upload DBR → see overdue / due-soon / on-track empty returns, color-coded 🔴🟡🟢.  
Filter by carrier or status. Export to Excel.

**SLA:** Empty returns due within 3 business days of delivery.

---

### Tab 4 — Carrier Data
Structured view of all DBR submissions. Filter by carrier or sheet type.  
Sub-tabs: Delivery, Empty Returns, Demurrage, Accessorials.  
Shows latest status per container (duplicates collapsed; full history retained in DB).

---

### Tab 5 — Lane Costs
Drayage rate lane management by SCAC + port + destination.  
Simulator shows cost comparison across carriers for a given lane.

---

### Tab 6 — Insights
Aggregated analytics: detention/demurrage risk, SLA performance, upcoming ETAs.  
Sections activate automatically as data is uploaded (DBR, GVT, Inbound Loads).

---

### Tab 7 — Planning (Weekly Delivery Scheduler)

Full SOP is documented inside the app at **Planning → SOP Guide** tab.  
Below is a condensed reference.

**SOP Metadata**

| Field | Value |
|---|---|
| Owner | Dominique Kennedy (kennewdo) |
| Frequency | Weekly — Every Friday |
| Plan Deadline | 3:00 PM ET / 2:00 PM CT / 12:00 PM PT |
| Carrier DBR Deadline | 3:00 PM CT Thursday (primary input) |
| Sites | RIC6, ILM1, DBM6/SAV, SJC8, XPH1-IAG1, LAX/West Coast |
| SharePoint | AGLRobotics → Carrier DBRs → subfolders: ATMI / ARVY / HUDD |

**Carriers**

| SCAC | Carrier | Ops Contact | Email |
|---|---|---|---|
| ATMI | Cargomatic | Tyler Domingues | tdomingues@cargomatic.com |
| ARVY | Arrive Logistics | Tyler Spangler | tspangler@arrivelogistics.com |
| HDDR – RIC6 | Maersk | Sandji Ruffin | sandji.ruffin@maersk.com |
| HDDR – ILM1 | Maersk | Jerry Nesbit | jerry.nesbit@maersk.com |
| HDDR – LAX | Maersk | Desirae Swain / Ailua Osoimalo | desirae.swain@maersk.com |

**Planning sub-tabs**

| Sub-tab | What it does |
|---|---|
| Plan Builder | Add containers (single or bulk paste), parse stakeholder requests, level-load across week |
| All Sites | Compiled view of full week across all sites — Excel export |
| By Site | Per-site view with carrier badges, daily Slack notification generator, mid-week adjustment tool |
| Carrier View | Shareable carrier-specific schedule + pre-send validation checklist + Excel export |
| WoW / History | Week-over-week delta (new / rolled / dropped) + full history search |
| Config | Sites table, carriers table, site–carrier map — all editable in-app |
| Import History | Upload ToteASERs Robotics DBR Tracker.xlsx to backfill history |
| **SOP Guide** | Full step-by-step SOP (Steps 1–9), GVT glossary, file naming reference |

**RIC6 Receiving Constraints**

| Parameter | Value |
|---|---|
| Receiving window | 7:30 AM – 4:30 PM |
| Lunch (no deliveries) | 12:00 PM – 1:00 PM |
| Last arrival | 3:30 PM |
| Max loads/day | 11–12 |
| No-receiving day | Monday |
| HDDR preference | Before lunch (7:30–11:30 AM) |

**Bulk Paste format** (Plan Builder → Step 2):
```
TCNU3773041  HDDR  RIC6
MRKU4103422  ATMI  ILM1
CSNU8812340  ARVY  DBM6
```
Carrier and Site are optional if defaults are set. Level-load checkbox distributes evenly across active days.

**DBR email subjects to expect Thursday EOD:**
```
ATMI  : [EXTERNAL] Amazon Robotics - DBR [DATE]
ARVY  : [EXTERNAL] Robotics DBR [DATE]
HUDD RIC6: [EXTERNAL] RIC6 Delivery Plan Update (HUDD)
HUDD ILM1: [EXTERNAL] ILM1 Delivery Plan Update / HUDD
HUDD LAX : RE: DBR Bridges Report - HUDD - [DATE]
```

---

### Tab 8 — WBR Generator

Generates the Monday WBR PDF + Perjen-style bridge, submitted to NA Destination Ops leadership.

**Submission SOP**

| Field | Value |
|---|---|
| To | `doc+destops-36@fusion.amazon.dev` |
| Subject | `NA Destination Ops WBR_Robotics` |
| Deadline | Monday by 2:00 PM CT |
| Attachment | `GLS_Robotics_YYYY-M-D.pdf` |
| Fusion doc set | https://fusion.amazon.dev/documentset/DOCUMENTSET%23f42d63bd-a96a-4f29-a356-a66ac18602a9 |

**Step-by-step:**

1. **(Optional) Enter Context Notes** — type operational callouts before uploading. Notes save to DB by week and auto-inject into the bridge at generate time.
2. **Upload 3 required files:**

   | File | Where to pull | Filter / Notes |
   |---|---|---|
   | **GVT Data** (`GVT Data WK##.xlsx`) | GM DCM Reports → Inbound Container Milestone | Customer=AMZ, Equip Category=AMAZON Robotics, ETA 6–12 wks back + forward |
   | **OBLT Data** | Ocean Bridge Logistics Tracking | All AV/OA/RD statuses for reporting week |
   | **Inbound Loads** | Amazon Robotics Inbound Loads Report | Run Monday morning |
   | *(Optional)* **Import Shipment Status** | CDS / Import Shipment Status report | Enriches Enhanced WBR forward look |

3. **Set report date** — defaults to most recent Monday. Adjust only if re-running a prior week.
4. **Click "Generate Both WBR Outputs"**
5. **Download PDF** → use "📧 Open in Email Client" button to pre-fill To/Subject/body
6. **After sending** → click "✅ Mark as Sent" and verify slide appears in Fusion doc set

**If prior weeks are missing:**  
The ⚠️ warning expander (auto-expanded) lets you enter historical values manually for W-5 through W-1. Values fill the chart columns on the slide.

**What the slide contains:**
- 6-week trend charts: AV→OA SLA%, OA→Del SLA%, OTP (top row) + Volume / E2E (bottom row)
- Summary table: 6 prior weeks + current week + totals column
- SLA Goals box (AV→OA ≤3d, OA→Del ≤3d BOS, Empty→Term ≤3d, OTP ≥95%)

**Bridge (Perjen format):**  
Auto-generated below the slide — `[Volume] / [AV→OA] / [OA→Del] / [Empty→Term] / [E2E/OTP]`  
Includes WoW deltas, dominant carrier callout, context notes auto-injected.  
Fully editable before sending. Carrier scorecard and root cause are in the right column (Enhanced WBR).

**WBR SLA thresholds:**

| Metric | SLA Target | Green |
|---|---|---|
| AV→OA transit | ≤ 3 days | ≥ 95% |
| OA→Del transit (BOS) | ≤ 3 days | ≥ 95% |
| Empty→Term return | ≤ 3 days | ≥ 95% |
| On-Time to Promise (OTP) | ≥ 95% | — |

**GVT Container Status Glossary** (for WBR data pulls):

| Status | Meaning |
|---|---|
| In Yard Full | At dray yard — ready to deliver |
| Dispatched to Destination | En route to FC or refused/returning |
| Not Ready | At port — pending customs or carrier release |
| On Water | Still at sea |
| Closed | Delivered and empty returned |

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
app.py              Main Streamlit app (8 tabs)
wbr_engine.py       WBR data parsing + metrics computation
wbr_pdf.py          PDF slide generator (pixel-faithful to gold standard)
data_sync.py        S3 sync — pull on startup, push after every write
utils.py            Container ID normalization helpers
lambda/             Email ingest + alert Lambdas (deploy separately)
deploy/SETUP.md     Lambda + SES deployment guide
tracker.db          SQLite — local + S3-synced
```

**Key DB tables:**

| Table | Purpose |
|---|---|
| `delivery_plan` | All container delivery plan entries |
| `plan_sites` | Site config (capacity, port, constraints) |
| `plan_carriers` | Carrier config (contact, email) |
| `plan_site_carrier` | Site–carrier mapping (priority time, allocation %) |
| `plan_week_config` | Per-week active receiving days |
| `wbr_results` | Historical WBR metrics by week |
| `wbr_context_notes` | Operational notes per week (auto-injected into bridge) |
| `carrier_submissions` | All carrier DBR submissions |
| `dbr_receipts` | DBR receipt tracking per carrier per week |

**Infrastructure:**

| Resource | Value |
|---|---|
| S3 bucket | `robotics-container-tracker` (us-east-1, account 844000647671) |
| IAM user | `robotics-tracker-reader` — read-only S3 for Streamlit Cloud |
| Streamlit Cloud | Auto-redeploys on `git push` to `main` (~60s) |
| Secrets | AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, APP_PASSWORD in Streamlit Cloud secrets |
