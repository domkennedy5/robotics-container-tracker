# Robotics Container Tracker — Build Spec for Kiro

## Context
This tool replaces a manual process where an Amazon Global Logistics (AGL) PM
receives lists of container IDs from internal partners and has to hand-look them
up in a weekly Excel file (the DBR). It also collects status updates that carriers
currently email in as Excel attachments.

The working prototype (`app.py`, 557 lines) already handles:
- Container lookup across 6 DBR sheets with check-digit-tolerant matching
- Carrier submission web form logged to SQLite
- Empty returns dashboard with overdue/due-soon flags

**This spec defines what needs to be built on top of that.**

---

## Project structure (target)

```
robotics-tracker/
├── app.py                  # Streamlit frontend (existing, modify as needed)
├── data_sync.py            # S3 sync module (NEW — modeled on Signal Tower pattern)
├── lambda/
│   └── email_ingest.py     # Lambda function for email-to-submission pipeline (NEW)
├── deploy/
│   └── SETUP.md            # Deployment instructions (NEW)
├── .streamlit/
│   └── secrets.toml        # Local secrets template (NEW — DO NOT commit real values)
├── requirements.txt        # (existing, update as needed)
└── README.md               # (existing)
```

---

## AWS account context
- Account: 844000647671
- Region: us-east-1
- S3 bucket to create: `robotics-container-tracker` (new, separate from signal-tower-data)
- SES: already in production mode on this account
- IAM user to create: `robotics-tracker-reader` (read-only S3, for Streamlit Cloud)
- Lambda role to create: `robotics-tracker-lambda-role` (S3 full + SES send)

---

## Task 1 — S3 sync (`data_sync.py`)

### What it does
Mirrors the Signal Tower `data_sync.py` pattern:
- On app load: pull latest DBR file from S3 if newer than local copy
- After a new DBR is uploaded via the app: push to S3
- SQLite tracker.db: pull from S3 on startup, push after every write

### S3 key structure
```
s3://robotics-container-tracker/
  dbr/latest.xlsx          # current DBR file
  dbr/archive/YYYY-MM-DD.xlsx  # archived on each upload
  db/tracker.db            # SQLite database
  last_updated.json        # { "dbr": "ISO timestamp", "db": "ISO timestamp" }
```

### Implementation notes
- Use boto3 with credentials from Streamlit secrets (AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY)
- Cache TTL: 300s (same as Signal Tower)
- Fallback to local file if S3 unreachable
- `pull_dbr_from_s3()` → returns local path
- `push_dbr_to_s3(local_path)` → uploads + archives
- `pull_db_from_s3()` → downloads tracker.db
- `push_db_to_s3()` → uploads tracker.db
- Call `pull_db_from_s3()` at app startup
- Call `push_db_to_s3()` after every carrier submission insert

---

## Task 2 — DBR upload flow (update `app.py`)

Currently the DBR is uploaded each session and lives only in memory.

**Change:** When user uploads a DBR file in the sidebar:
1. Save to local `dbr/latest.xlsx`
2. Call `push_dbr_to_s3()` to persist it
3. On next app load, `pull_dbr_from_s3()` auto-loads it — no re-upload needed
4. Show last-updated timestamp in sidebar ("DBR last updated: July 22, 2026")
5. Keep the manual upload option so user can push a new file anytime

---

## Task 3 — Email ingestion Lambda (`lambda/email_ingest.py`)

### Flow
```
Carrier sends email with Excel/CSV attachment
    → SES receives it (requires verified domain or use existing SES setup)
    → SES rule: save raw email to S3 at s3://robotics-container-tracker/inbound/{message-id}
    → S3 event triggers Lambda
    → Lambda parses email (extract sender, subject, attachments)
    → For each Excel/CSV attachment: parse container rows (same logic as app.py carrier file parser)
    → Insert rows into tracker.db (pull from S3, write, push back)
    → Send confirmation reply to sender via SES
    → Log to s3://robotics-container-tracker/ingest-log/YYYY-MM-DD.jsonl
```

### Lambda implementation details
```python
# Dependencies (Lambda layer or zip): boto3, openpyxl, pandas, email (stdlib)
# Environment variables:
#   S3_BUCKET = robotics-container-tracker
#   SES_FROM_EMAIL = <verified sender address>
#   DB_S3_KEY = db/tracker.db

# Handler signature: handler(event, context)
# event comes from S3 trigger (s3:ObjectCreated on prefix inbound/)
```

### Email parsing
- Use Python `email` stdlib to parse raw .eml from S3
- Extract: sender address, subject, all attachments
- For each attachment with `.xlsx` or `.csv` extension: run the same `parse_carrier_file()` logic already in app.py (refactor that function into a shared `utils.py` so both app.py and the Lambda can import it)

### Container row schema (matches existing DB table)
```sql
carrier_submissions (
    id, submitted_at, carrier_name, container_id,
    terminal, status, notes, source_file
)
-- carrier_name = parsed from email sender / subject line
-- source_file  = original attachment filename
```

### Confirmation email template
```
Subject: Re: [original subject] — Received ✓
Body:
  Hi,
  We received your submission and logged [N] containers.
  Containers processed: [list]
  Submitted at: [timestamp UTC]
  
  If anything looks wrong reply to this email.
  — Amazon Global Logistics Container Tracker
```

---

## Task 4 — SES notification for overdue empty returns

### Trigger
A scheduled Lambda (EventBridge, daily at 8 AM EST) checks the DBR's Empty Returns
sheet (loaded from S3) and sends reminders.

### Logic
```python
# Pull latest DBR from S3
# Load Empty Returns sheet
# For each container where:
#   - Status != 'TERMINATED'
#   - Empty Return Due Date <= today + 2 days (due soon) OR < today (overdue)
#   - Not already notified today (check alerts_log table in tracker.db)
# Send SES email to configured recipient list
# Write to alerts_log: (alert_type, container_id, sent_at)
```

### alerts_log table (add to DB schema in app.py `init_db()`)
```sql
CREATE TABLE IF NOT EXISTS alerts_log (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    alert_type   TEXT NOT NULL,   -- 'empty_return_overdue' | 'empty_return_due_soon'
    container_id TEXT NOT NULL,
    sent_at      TEXT NOT NULL,
    recipient    TEXT
);
```

### Alert email format
```
Subject: [ACTION REQUIRED] Empty Return Due — {N} containers
Body: table of container | terminal | due date | days overdue
```

---

## Task 5 — Streamlit secrets / config

### `.streamlit/secrets.toml` (template — user fills in real values)
```toml
[aws]
AWS_ACCESS_KEY_ID     = "FILL_IN"
AWS_SECRET_ACCESS_KEY = "FILL_IN"
AWS_REGION            = "us-east-1"
S3_BUCKET             = "robotics-container-tracker"

[notifications]
ALERT_RECIPIENTS      = ["email@amazon.com"]
SES_FROM_EMAIL        = "FILL_IN_VERIFIED_SES_EMAIL"

[app]
PASSWORD              = "FILL_IN"   # simple password gate, same as Signal Tower
```

### Secrets access pattern in app.py
```python
import streamlit as st
AWS_KEY    = st.secrets["aws"]["AWS_ACCESS_KEY_ID"]
AWS_SECRET = st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"]
REGION     = st.secrets["aws"]["AWS_REGION"]
BUCKET     = st.secrets["aws"]["S3_BUCKET"]
```

---

## Task 6 — Simple password gate (update `app.py`)

Add the same password gate pattern used in Signal Tower:
- Check `st.session_state.authenticated`
- If not authenticated: show password input, compare to `st.secrets["app"]["PASSWORD"]`
- If authenticated: show full app
- Gate applies to all tabs

---

## Task 7 — `deploy/SETUP.md`

Document these steps:
1. Create S3 bucket `robotics-container-tracker`
2. Create IAM user `robotics-tracker-reader` with S3 read policy, generate keys
3. Set up SES verified email for sending
4. Set up SES receiving rule (if using email ingestion): inbound → S3 → Lambda trigger
5. Deploy Lambda (`lambda/email_ingest.py`) with role `robotics-tracker-lambda-role`
6. Set EventBridge rule for daily empty-returns alert Lambda
7. Deploy to Streamlit Cloud: connect GitHub repo, set secrets in Streamlit Cloud dashboard

---

## Constraints / conventions
- **No credentials in code** — all via st.secrets or Lambda env vars
- **Fallback gracefully** — if S3 unreachable, fall back to local SQLite / local DBR file
- **Shared logic** — refactor `parse_carrier_file()` and `normalize_container()` into `utils.py` so Lambda and app.py share them without duplication
- **Keep app.py clean** — import from data_sync.py and utils.py, don't inline everything
- **Match Signal Tower patterns** — same S3 sync TTL, same secrets structure, same password gate

---

## What NOT to change
- The three-tab UI structure (Lookup / Carrier Submission / Empty Returns)
- The `containers_match()` prefix-matching logic
- The `SHEET_CFG` dict defining which DBR sheets to search
- The existing SQLite schema for `carrier_submissions` and `lookup_log`

---

## Kiro: start here
1. Read `app.py` in full first
2. Create `utils.py` — extract `normalize_container`, `containers_match`, `parse_container_list`, `parse_carrier_file` from app.py
3. Create `data_sync.py` — S3 sync module per Task 1
4. Update `app.py` — import from utils.py and data_sync.py, add password gate, add S3-backed DBR upload
5. Create `lambda/email_ingest.py` — per Task 3
6. Add `alerts_log` table to `init_db()` in app.py
7. Create `lambda/empty_return_alerts.py` — scheduled alert Lambda per Task 4
8. Create `.streamlit/secrets.toml` template
9. Create `deploy/SETUP.md`
10. Update `requirements.txt`
