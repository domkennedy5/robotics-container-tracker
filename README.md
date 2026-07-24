# Robotics Container Tracker

Live app: https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app  
Password: `robotics2026`

## Local run
```
pip install -r requirements.txt
streamlit run app.py
```

---

## Tabs

### Container Lookup
Upload the weekly DBR Excel → paste any list of container IDs → get instant cross-sheet results.  
Searches across: Delivery Appointments, Empty Returns, On Vessel, Canceled, Demurrage, Accessorials.  
Check-digit-tolerant matching (dashes optional). Download results as Excel.

### Carrier Submission
Carriers fill in container updates directly, or upload an ARVY-style Excel template.  
All submissions log to `tracker.db` (SQLite, S3-synced) with timestamp.  
Full log is searchable/filterable/exportable.

### Empty Returns Dashboard
Upload DBR → see overdue/due-soon/on-track empty returns, color-coded 🔴🟡🟢. Exportable.

### Detention / Demurrage Risk
Flags containers at LFD risk based on current DBR status and port logic.

### Carrier Scorecard
Aggregate SLA performance by carrier from submission history.

### Rate Lanes
Manage drayage rate lanes by SCAC + port + destination.

### WoW / History
Week-over-week trend charts for AV→OA SLA, OA→Del SLA, OTP, volume.

### Planning
Delivery scheduling board — assign containers to sites, carriers, and time slots by week. Includes DBR receipt tracker, site/carrier config, and capacity management.

### WBR Generator (Tab 8)
**Automated WBR slide builder** — generates the Monday WBR PDF submitted to Mitch/DestOps.

**How to use:**

1. **Context Notes** (optional) — type any operational callouts before uploading files. Notes save to DB by week and auto-inject into the bridge at generate time.

2. **Upload 3 files:**
   - **GVT Data** — Global Visibility Tool export, filtered to the reporting week (Sun–Sat), Robotics BU, all markets
   - **OBLT Data** — Ocean Bridge Logistics Tracking export with AV/OA/RD statuses
   - **Inbound Loads** — Amazon Robotics Inbound Loads Report (run Monday morning)
   - *(Optional)* **Import Shipment Status** — enriches the Enhanced WBR forward look with vessel/ETA data

3. **Set report date** — defaults to most recent Monday

4. **Click Generate Both WBR Outputs**

5. **Download the PDF** → submit per the SOP below

**WBR submission SOP:**  
To: `doc+destops-36@fusion.amazon.dev`  
Subject: `NA Destination Ops WBR_Robotics`  
Deadline: Monday by 2:00 PM CT  
Attachment: `GLS_Robotics_YYYY-M-D.pdf`

The app provides an "Open in Email Client" button that pre-fills To/Subject/body.  
If that doesn't work: use the **Manual submission instructions** expander — numbered steps + copy-paste email body.

**What the slide contains:**
- 6-week trend charts: AV→OA SLA%, OA→Del SLA%, OTP (top row) + volume/E2E (bottom row)
- Summary table with 6 prior weeks + current week total column
- SLA Goals box (AV→OA ≤3d, OA→Del ≤3d BOS, Empty→Term ≤3d, OTP ≥95%)

**Bridge format (Perjen-style):**  
Auto-generated below the slide — `[Volume] / [AV→OA] / [OA→Del] / [Empty→Term] / [E2E/OTP] / [Op Callouts]`  
Includes WoW deltas, carrier breakdown (top 3), and context notes injected automatically.  
Fully editable in-app before sending.

---

## Data persistence
- SQLite `tracker.db` — committed to repo + synced to/from S3 after every write
- S3 bucket: `robotics-container-tracker` (us-east-1, account 844000647671)
- IAM user `robotics-tracker-reader` provides read-only S3 access for Streamlit Cloud

## Architecture
```
app.py              Main Streamlit frontend (8 tabs)
wbr_engine.py       WBR data parsing + metrics computation
wbr_pdf.py          PDF slide generator (pixel-faithful to 7/20 gold standard)
data_sync.py        S3 sync (pull on startup, push after writes)
utils.py            Container normalization helpers
lambda/             Email ingest + alert Lambdas (deploy separately)
deploy/SETUP.md     Lambda + SES deployment instructions
```
