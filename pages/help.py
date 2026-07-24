from __future__ import annotations
import streamlit as st

st.set_page_config(
    page_title="How to Use — Container Tracker",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_URL    = "https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app"
VENDOR_URL = f"{APP_URL}/vendor_upload"

st.title("AGL Robotics Container Tracker — User Guide")
st.caption("Quick reference for the AGL ops team. Last updated July 2026.")
st.divider()

# ── What this tool does ───────────────────────────────────────────────────────
st.markdown("## What This Tool Does")
st.markdown("""
The **AGL Robotics Container Tracker** is a single place to manage inbound robotics container
deliveries — from weekly WBR reporting through carrier tracking, delivery planning, and risk monitoring.

| Workflow | What you need | Where to go |
|---|---|---|
| Generate the weekly WBR slide and bridge | GVT, OBLT, Inbound Loads exports | **WBR Generator** tab |
| Build and manage the weekly delivery plan | Nothing | **Planning** tab |
| Review risk and weekly summaries | Weekly DBR (sidebar) | **Insights** tab |
| Look up container status | Weekly DBR (sidebar) | **Container Lookup** tab |
| Receive carrier status data | Nothing | **Carrier Submission** tab |
| Track empty return obligations | Weekly DBR (sidebar) | **Empty Returns** tab |
| View all carrier submissions | Nothing | **Carrier Data** tab |
| Model costs across lanes | Rate card loaded | **Lane Costs** tab |
""")
st.divider()

# ── Weekly workflow ───────────────────────────────────────────────────────────
st.markdown("## Monday Morning Workflow")
st.markdown("Do these steps every Monday. Total time: about 15–20 minutes.")

steps = [
    (
        "1. Pull GVT, OBLT, and Inbound Loads",
        "External systems → WBR Generator",
        "Before opening the app, pull three reports:\n\n"
        "- **GVT** — Export Robotics containers for the week that just closed (Sun–Sat). "
        "Filter Ready Date to the prior week's date range. Save as `GVT Data WK##.xlsx`.\n"
        "- **OBLT** — Export milestone events (AV/OA/VD/RD) for the same containers. "
        "Save as `OBLT WK##.xlsx`.\n"
        "- **Inbound Loads Report** — Run from the Reporting Portal as of Monday morning. "
        "This contains PO Promised Date and Actual Arrival for OTP scoring.\n"
        "- **Import Shipment Status** *(optional)* — Run from Reporting Portal → Import Shipment Status. "
        "Provides EDI tracking dates and vessel/ETA data that auto-populates the Enhanced WBR forward look table.\n\n"
        "The WBR Generator tab shows exact date ranges and pull steps for the current week automatically — "
        "expand each **How to pull** instruction before going to the source system."
    ),
    (
        "2. Generate the WBR slide",
        "WBR Generator tab",
        "Open the **WBR Generator** tab. The tab shows the current WBR week and due date at the top.\n\n"
        "1. Upload GVT, OBLT, and Inbound Loads using the three uploaders\n"
        "2. Confirm the report date (defaults to this Monday)\n"
        "3. Review the **Data Preview** and **Data Quality** panel — check container count and OBLT coverage\n"
        "4. Click **Generate Both WBR Outputs**\n"
        "5. **Left column — Standard WBR:** Download `GLS_Robotics_YYYY-M-D.pdf`, edit the bridge, "
        "click **📧 Open in Email Client** to send to Mitch. Deadline: **2:00 PM CT**.\n"
        "6. **Right column — Enhanced WBR:** Carrier scorecard by AV→OA and OA→Del, root cause notes, "
        "forward look — for the internal Robotics team meeting.\n\n"
        "If Import Shipment Status is uploaded, the forward look table auto-populates with live vessel names, "
        "voyage numbers, and FinalDestETAs for all containers arriving in the next 14 days."
    ),
    (
        "3. Load the weekly DBR",
        "Sidebar",
        "Open the sidebar (top-left arrow). Under **DBR File**, upload the current week's "
        "DBR Excel. The app parses all sheets automatically and shows row counts. "
        "This activates Container Lookup, Empty Returns, and the Insights DBR Snapshot."
    ),
    (
        "4. Upload supplemental reports",
        "Sidebar → Additional Reports",
        "Upload the **Import Shipment Status** and **Inbound Loads Report** under "
        "**Additional Reports** in the sidebar. The Inbound Loads Report is the most "
        "time-sensitive — it contains detention and demurrage free-time deadlines. "
        "(If you already pulled it for the WBR, upload the same file here.)"
    ),
    (
        "5. Refresh historical data",
        "Planning → Import History",
        "Download `ToteASERs Robotics DBR Tracker.xlsx` from SharePoint. "
        "Go to **Planning → Import History**, upload the file, click **Import**. "
        "New containers since last week are added automatically — duplicates are skipped. "
        "This keeps the WoW / History tab current."
    ),
    (
        "6. Build the week's delivery plan",
        "Planning → Plan Builder",
        "Select the week at the top of the Planning tab. Paste Miguel's or site ops' "
        "message into **Step 1 — Parse Stakeholder Request** and click Parse. "
        "Add containers using Single or Bulk Paste. Review the week grid at the bottom."
    ),
    (
        "7. Send daily notifications",
        "Planning → By Site",
        "Select a site, open **Daily Notification — Copy for Slack**, pick the delivery "
        "date, copy the generated text block, paste into Slack."
    ),
    (
        "8. Check for risk",
        "Insights tab",
        "Review Detention & Demurrage Risk and the LFD Risk table. These flag containers "
        "with deadlines within 7 days. Check this before finalizing the weekly plan — "
        "it affects which containers need priority scheduling."
    ),
]

for title, location, desc in steps:
    with st.expander(f"**{title}** — *{location}*"):
        st.markdown(desc)

st.divider()

# ── Quick reference ───────────────────────────────────────────────────────────
st.markdown("## Quick Reference")
st.markdown("Find what you need without reading the full guide.")
st.markdown("""
| I want to... | Go here |
|---|---|
| Generate the weekly WBR slide (Mitch) | WBR Generator — upload 3 files → Generate Both |
| Email the standard WBR to Mitch | WBR Generator → Standard column → 📧 Open in Email Client |
| See the carrier-level WBR breakdown | WBR Generator → Enhanced column |
| See which week's data to pull for the WBR | WBR Generator — due date banner + How to pull expanders |
| Find a container's current status | Container Lookup — paste ID, click Search |
| See which empty returns are overdue | Empty Returns tab |
| Build next week's delivery schedule | Planning → Plan Builder |
| Send tomorrow's delivery list to a site | Planning → By Site → Daily Notification |
| Share the plan with a carrier | Planning → Carrier View → Export |
| Compare this week to last week | Planning → WoW / History |
| Update a container's time, site, or status | Planning → By Site or All Sites → Edit Entry |
| See containers at risk of detention fees | Insights → Detention & Demurrage Risk |
| Load the weekly DBR | Sidebar → DBR File |
| Load the Inbound Loads or Shipment Status report | Sidebar → Additional Reports |
| Add a new site or carrier | Planning → Config |
| Load historical container data | Planning → Import History |
| Submit carrier status data | Carrier Submission tab or Vendor Portal |
| Model lane costs and scenarios | Lane Costs tab |
""")
st.divider()

# ── Tab reference ─────────────────────────────────────────────────────────────
st.markdown("## Tab Reference")
st.caption("Tabs appear in this order in the app.")

tab_ref = st.tabs([
    "WBR Generator",
    "Planning",
    "Insights",
    "Container Lookup",
    "Carrier Submission",
    "Carrier Data",
    "Empty Returns",
    "Lane Costs",
])

with tab_ref[0]:
    st.markdown("""
**Generates both WBR deliverables — Standard slide (for Mitch) and Enhanced analysis (for Robotics team) — in one click.**

The tab auto-detects which week you're working on and shows a live due date countdown at the top.
A translucent slide preview updates as you upload each file, giving you a visual progress indicator before you generate.

**What you need:**
| File | Required | Source | Notes |
|---|---|---|---|
| GVT Data (.xlsx) | ✅ | Global Visibility Tool | Robotics containers, filtered to reporting week |
| OBLT Data (.xlsx) | ✅ | Ocean Bridge Logistics Tracking | Milestone events for the same containers |
| Inbound Loads (.xlsx) | ✅ | Reporting Portal → Amazon Robotics | Run Monday morning before generating |
| Import Shipment Status (.xlsx) | Optional | Reporting Portal → Import Shipment Status | Enriches Enhanced WBR forward look with live vessel/ETA data |

**To use:**
1. Open the **How to pull** expanders — they show exact date ranges for the current week
2. Pull the three files from their source systems
3. Upload all three using the uploaders
4. Confirm the report date (defaults to this Monday)
5. Review the **Data Preview** expander to inspect raw files; check **Data Quality** warnings
6. Click **Generate Both WBR Outputs**

**Standard WBR (left column):**
- Downloads `GLS_Robotics_YYYY-M-D.pdf` — the slide Mitch receives
- Editable bridge pre-filled with all KPIs for the week
- **📧 Open in Email Client** — opens your email with recipient, subject, and body pre-filled
- Submit to: `doc+destops-36@fusion.amazon.dev` · Subject: `NA Destination Ops WBR_Robotics` · **Deadline: 2:00 PM CT**

**Enhanced WBR (right column):**
- Carrier scorecard — AV→OA and OA→Del performance by carrier (ATMI, RKNE, HDDR, ARVY)
- Root cause notes per carrier and classification (a–e)
- 2-week forward look table
- Editable enhanced bridge for the internal Robotics team meeting

**Slide preview:**
- Gray/translucent = no data loaded for that section
- Amber = file uploaded, prior week shown as reference
- Green = current week data confirmed

**Prior week carry-forward:**
Previously generated weeks are saved to the database automatically. Use the **Seed W24–W28** button
(available when no files are uploaded) to pre-load historical data from the first slide.
""")

with tab_ref[1]:
    st.markdown("""
**Full delivery plan builder — standalone, no DBR upload required.**

| Sub-tab | Use it to |
|---|---|
| Plan Builder | Parse stakeholder requests, add containers (single or bulk), review the week grid |
| All Sites | Full compiled view of the week; Excel export |
| By Site | Site-specific view, daily Slack notification, inline edit |
| Carrier View | Carrier-specific view, shareable Excel (no internal status) |
| WoW / History | Compare this week vs. prior weeks; search historical containers |
| Config | Add/edit sites, carriers, and site-carrier mappings |
| Import History | Upload the DBR Tracker to load historical container data |

**Week selector** at the top controls which week you're viewing or editing.

**Receiving days** — the day toggles below the week selector let you activate/deactivate
days for that week. Mon–Fri is the default. Changing active days offers to redistribute
any unassigned containers automatically.

**Edit Entry** — available in By Site and All Sites. Select a container from the dropdown
and update date, time, site, carrier, product, status, qty, or notes. Check **Delete this entry**
to remove it.

**Daily Notification** — in By Site, open the expander, pick a date, copy the generated
text block, paste directly into Slack.
""")

with tab_ref[2]:
    st.markdown("""
**Auto-generated weekly summary across all loaded data sources.**

Each section activates as its source data is uploaded:

| Section | Requires | Shows |
|---|---|---|
| DBR Snapshot | DBR in sidebar | Delivery counts, empty return risk, LFD risk table, status/terminal breakdown |
| Detention & Demurrage Risk | Inbound Loads Report | Containers within 7 days of free-time expiry — most actionable section |
| In-Transit Pipeline | Import Shipment Status | Active containers by status, vessel, FC destination |

Upload supplemental reports via **Sidebar → Additional Reports**.

**Detention & Demurrage Risk is the most time-sensitive section.** Review it before finalizing
the weekly delivery plan — containers approaching their free-time deadline need priority scheduling.
""")

with tab_ref[3]:
    st.markdown("""
**Searches the DBR for any container ID you paste in.** Returns sheet, status, terminal,
FC/building, appointment date, and LFD for every match.

**To use:**
1. Upload the DBR in the sidebar first (required)
2. Paste container IDs — one per line or comma-separated; dashes and check digits are optional
3. Enter your name (optional — logged for audit trail)
4. Click **Search**
5. Download results as Excel if needed

**Searches across:** Delivery Appointments, Empty Returns, On Vessel, Canceled, Demurrage, Accessorials

**Not found?** The container is not in the current DBR. It may be delivered, not yet entered,
or on a different vessel not yet in the system. Try searching in Planning → WoW / History for
historical containers.
""")

with tab_ref[4]:
    st.markdown(f"""
**Two ways for carriers to submit their weekly container status:**

**Option 1 — Template upload (preferred):**
1. Download the AGL Carrier Template from the button at the top of the tab
2. Give it to your carrier to fill out
3. Enter carrier name, upload the completed file, click **Confirm & Submit All**

**Option 2 — Vendor Portal (carrier self-service):**
Share this URL with your carrier — no login required:
`{VENDOR_URL}`

**Template sheets:** Delivery, ILM1, RIC6, Empty Return, ODY/Storage, Demurrage, Accessorials.
Carriers only fill in sheets relevant to that week.

**Submission Log** at the bottom of the tab shows all historical entries, filterable by carrier,
status, and source. Export to Excel any time.

**Onboarding a new carrier:** Send them the vendor portal URL. On their first submission
their name auto-populates in all filters throughout the app.
""")

with tab_ref[5]:
    st.markdown("""
**Full view of all carrier submissions, organized by sheet type.**

This is the historical record of everything carriers have submitted — both via the Carrier
Submission tab and the Vendor Portal.

**Filters:** Carrier name, Sheet type, Within SLA

Sub-tabs appear automatically for each sheet type present (Delivery, Empty Returns, Demurrage, etc.).
Only columns with data are shown — no empty columns cluttering the view.

**Export** all carrier data to Excel from the button at the bottom of the tab.

Use this tab when you need to audit what a specific carrier submitted, compare across carriers,
or pull data for a dispute or review.
""")

with tab_ref[6]:
    st.markdown("""
**Tracks containers that still need to be returned to the port terminal after FC delivery.**
When a container is delivered to an FC, the dray carrier has a deadline to return the empty
to the terminal. Missing that deadline results in per diem fees.

Requires DBR to be loaded in the sidebar.

**Filters:** Overdue / Due Soon (≤3 days) / All Open

**Columns to watch:** Container #, Terminal, Empty Return Due Date, Days Until Due, Alert

The table shows only open obligations, sorted by most urgent at the top.

**Terminated = resolved.** If a container's status is Terminated, the return obligation is
closed — it was returned or the requirement was waived. Terminated containers are never
counted as overdue. They appear in a collapsed "Terminated / Resolved" section at the
bottom for reference only.

Export the filtered view to Excel as needed.
""")

with tab_ref[7]:
    st.markdown("""
**Rate matrix and scenario cost builder by port → destination → carrier.**

Requires rate data to be loaded by the app admin. Shows an info message if empty.

**Rate Matrix:** All active lanes with carrier options side by side.
Green = cheapest on that lane. Red = most expensive. Use this to quickly see the cost spread
for a given port-to-destination pair.

**Scenario Builder:**
1. Select Port → Destination Node → Carrier
2. Enter container count → click **Add Lane to Scenario**
3. Build a multi-lane scenario; total cost and savings vs. cheapest are calculated automatically
4. Edit container counts inline after adding

Export scenario to Excel for sharing with stakeholders.

**Rate data not loaded?** Contact Dominique Kennedy (AGL) to get the rate card entered.
""")

st.divider()

# ── Data sources ──────────────────────────────────────────────────────────────
st.markdown("## Data Sources")

col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.markdown("#### GVT — Global Visibility Tool (Weekly, WBR)")
        st.markdown("""
Export of all Robotics containers for the reporting week.

**Provides:** Container IDs, Ready Date/Time, market, port, carrier (SCAC), facility,
status, empty return date, enter facility date.

**Load via:** WBR Generator → GVT Data uploader

**Cadence:** Weekly — pull Monday morning for the prior Sun–Sat week.
""")

    with st.container(border=True):
        st.markdown("#### OBLT — Ocean Bridge Logistics Tracking (Weekly, WBR)")
        st.markdown("""
Milestone event export for all Robotics containers.

**Provides:** AV (Available), OA (Outgate/Pickup), VD (Vessel Departed), RD (Received/Delivered)
event timestamps per container. Used to compute AV→OA and OA→Delivery transit times.

**Load via:** WBR Generator → OBLT Data uploader

**Cadence:** Weekly — pull same day as GVT.
""")

    with st.container(border=True):
        st.markdown("#### DBR — Delivery Booking Report (Weekly)")
        st.markdown("""
The weekly Delivery Booking Report Excel maintained by the drayage team.

**Provides:** Container status across Delivery Appointments, Empty Returns,
On Vessel, Canceled, Demurrage, Accessorials.

**Load via:** Sidebar → DBR File → Upload

**Cadence:** Weekly — upload each Monday or when received.
""")

with col2:
    with st.container(border=True):
        st.markdown("#### Inbound Loads Report (Weekly, WBR + Insights)")
        st.markdown("""
Amazon's internal AR Inbound Loads report with PO-level container detail.

**Provides (WBR):** PO Promised Date and Actual Arrival — used to compute OTP (On-Time to Promise).

**Provides (Insights):** Last Free Detention time, Last Free Demurrage time,
gate-out terminal time, estimated appointment date, empty return time.

**Load via:** WBR Generator → Inbound Loads uploader AND Sidebar → Additional Reports

**Most time-sensitive source** — upload Monday morning before generating the WBR slide.
""")

    with st.container(border=True):
        st.markdown("#### Import Shipment Status (Weekly)")
        st.markdown("""
Shipment lifecycle report from Amazon's TMS.

**Provides:** EDI dates (ingate, departure, arrival, available, out-gate,
empty return, delivered), vessel/voyage, carrier code, discharge port ETA,
destination FC, vendor name.

**Load via:** Sidebar → Additional Reports → Import Shipment Status

**Source:** Emailed report — filename typically "Import Shipment Status.xlsx"
""")

    with st.container(border=True):
        st.markdown("#### DBR Tracker — Historical (Weekly)")
        st.markdown("""
The master `ToteASERs Robotics DBR Tracker.xlsx` on SharePoint.

**Provides:** Full delivery history by container — carrier, site, scheduled
and actual FC delivery dates, product type, quantity, status.

**Load via:** Planning → Import History → Upload → Import

**Cadence:** Weekly — upload Monday to capture prior week's new containers.
""")

st.divider()

# ── Vendor portal ─────────────────────────────────────────────────────────────
st.markdown("## Vendor Portal")
st.markdown("A public-facing, no-login page for carriers to submit weekly status templates.")
st.code(VENDOR_URL, language=None)
st.markdown(f"""
**What carriers do:**
1. Go to the URL above
2. Download the AGL Carrier Template
3. Fill in their container data across the relevant sheets
4. Enter company name and contact name
5. Upload and submit — data goes directly into the tracker

**What you see:** All submissions appear in Carrier Data and Carrier Submission tabs,
tagged `source = vendor_portal` so you can filter them.

**Onboarding a new carrier:** Send them the vendor portal URL. On their first submission
their name auto-populates in all filters.
""")
st.divider()

# ── FAQ ───────────────────────────────────────────────────────────────────────
st.markdown("## FAQ")

faqs = [
    (
        "The WBR tab is showing the wrong week number.",
        "The WBR tab anchors to the most recent Monday. If it's showing W30 but you're working on W29, "
        "the report date picker defaults to last Monday — confirm it's set to the correct submission Monday "
        "and check that the GVT file you're uploading covers the right Sun–Sat range. "
        "The 'How to pull' expanders show the exact dates to use for each source system."
    ),
    (
        "The slide preview boxes are amber instead of green after uploading files.",
        "Amber means a file is uploaded but the metrics haven't been generated yet — the preview is "
        "showing last week's values as a reference. Click **Generate Both WBR Outputs** and the boxes "
        "will fill with current week values at full opacity."
    ),
    (
        "Container count looks too low on the WBR.",
        "Two things to check: (1) The GVT Ready Date range — make sure you pulled Sun–Sat for the right week. "
        "The 'How to pull GVT' expander shows the exact dates. (2) Check the Data Quality panel after "
        "upload — it will flag if OBLT coverage is low or if the container count is below typical range."
    ),
    (
        "The 📧 Open in Email Client button doesn't open my email.",
        "The button generates a mailto: link. If your browser blocks it, expand **Email body (copy-paste)** "
        "directly below the button and copy the pre-filled content manually into your email client. "
        "The recipient and subject line are shown directly above the button."
    ),
    (
        "The app is showing old DBR data. How do I refresh it?",
        "Upload the new DBR in the sidebar. The app syncs from S3 on load — if a teammate "
        "already uploaded a newer file, refresh the page and you'll see it. "
        "If not, upload the file yourself and it overwrites S3 automatically."
    ),
    (
        "A container isn't showing up in the search.",
        "A few possible causes: (1) It's not in the current DBR — it may be delivered or "
        "on a vessel not yet in the system. (2) Try just the first 10 characters without "
        "any suffix. (3) Search historical containers in Planning → WoW / History."
    ),
    (
        "A carrier says they submitted but I don't see it.",
        "Check the Carrier Submission tab's log and filter by carrier name. If it's not "
        "there, ask them to resubmit. Common issues: old template format, or they didn't "
        "enter a company name before submitting."
    ),
    (
        "Detention / Demurrage Risk section is empty.",
        "That section requires the Inbound Loads Report. Upload it in Sidebar → Additional "
        "Reports. If it still shows nothing after upload, either all containers have already "
        "gated out (good), or you uploaded the wrong file — check you're using the AR "
        "Inbound Loads, not the Import Shipment Status."
    ),
    (
        "The Planning tab shows no containers for this week.",
        "Either no containers have been added yet for the selected week, or the week "
        "selector is on a week with no entries. Use the arrow buttons next to the date "
        "to move to the right week, or go to Plan Builder and add containers."
    ),
    (
        "How do I share access with someone on the team?",
        f"Send them the app URL ({APP_URL}) and the password: **robotics2026**. "
        "The vendor portal does not require a password — it's public-facing by design."
    ),
    (
        "How often should I upload the DBR Tracker for Import History?",
        "Weekly — Monday morning before you build the plan. Each upload only adds "
        "containers that aren't already in the database, so it's safe to run it "
        "more often if needed."
    ),
    (
        "Lane Costs shows 'No rate data loaded'.",
        "Rate card data is managed by the app admin. Contact Dominique Kennedy (AGL) "
        "to get rates loaded. This tab is a framework for when rate data is available."
    ),
]

for question, answer in faqs:
    with st.expander(f"**{question}**"):
        st.markdown(answer)

st.divider()
st.caption(f"Password: robotics2026  ·  App: {APP_URL}  ·  Questions? Contact Dominique Kennedy, AGL")
