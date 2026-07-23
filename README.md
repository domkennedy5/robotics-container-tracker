# Robotics Container Tracker

## Local run
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Cloud)
1. Push this folder to a GitHub repo
2. Go to share.streamlit.io → New app → point at app.py
3. No secrets needed (SQLite is local; add S3 sync later if needed)

## How it works

### Container Lookup (Tab 1)
- Upload the weekly DBR Excel file in the sidebar
- Paste any list of container IDs (dashes/no dashes both work)
- Searches across: Delivery Appointments, Empty Returns, On Vessel, Canceled, Demurrage, Accessorials
- Download results as Excel

### Carrier Submission (Tab 2)
- Carriers fill in their carrier name, containers, and status
- Or upload their ARVY-style Excel file directly
- Everything gets stored in tracker.db (SQLite) with a timestamp
- Full log is searchable/filterable and exportable

### Empty Returns Dashboard (Tab 3)
- Upload DBR → instantly see overdue / due-soon / on-track empty returns
- Color-coded 🔴🟡🟢
- Exportable

## Roadmap (future)
- Email reminders for overdue/due-soon empty returns via SES
- Carrier-facing public submission link
- S3 sync for persistence across Streamlit Cloud restarts
- Auto-ingest DBR from S3 (if uploaded by someone daily)
