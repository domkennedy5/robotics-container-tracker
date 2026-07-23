from __future__ import annotations
import streamlit as st

st.set_page_config(
    page_title="How to Use — Container Tracker",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_URL = "https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app"
VENDOR_URL = f"{APP_URL}/vendor_upload"

st.title("📖 AGL Robotics Container Tracker — User Guide")
st.caption("Reference guide for the AGL ops team. Last updated July 2026.")

# ── Quick navigation ──────────────────────────────────────────────────────────
st.markdown("""
**Jump to a section:**
[Quick Start](#quick-start) · [Tab Reference](#tab-reference) · [Data Sources](#data-sources) · [Vendor Portal](#vendor-portal) · [FAQ](#faq)
""")
st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — WHAT THIS TOOL DOES
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## What This Tool Does")
st.markdown("""
The **AGL Robotics Container Tracker** replaces manual container lookups across the weekly DBR 
and related Excel reports. It gives you a single place to:

- **Find any container instantly** across all DBR sheets (Delivery Appts, Empty Returns, On Vessel, Demurrage, etc.)
- **Track detention and demurrage deadlines** before costs start accruing
- **Receive carrier submissions** through a standardized portal instead of ad-hoc emails
- **Project dray costs** using the 2025–26 Robotics rate card by lane
- **See a weekly intelligence summary** — action items, LFD risk, carrier SLA, and in-transit pipeline — without building it manually

All data is synced to S3 so updates are shared across users automatically.
""")
st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — QUICK START
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## Quick Start")
st.markdown("**Weekly setup takes about 2 minutes. Do this every time you get a new report.**")

steps = [
    ("1. Upload the DBR", "sidebar",
     "In the **sidebar**, under **DBR File**, upload the latest weekly DBR Excel file. "
     "The app parses all 6 sheets automatically and shows row counts. This unlocks Container Lookup, "
     "Empty Returns, and most of the Insights tab."),
    ("2. Upload Import Shipment Status", "sidebar",
     "Under **Additional Reports** in the sidebar, upload the **Import Shipment Status** Excel. "
     "This adds vessel names, discharge ETAs, container availability dates, and delivery status "
     "for every active container."),
    ("3. Upload Inbound Loads Report", "sidebar",
     "Upload the **Inbound Loads Report** Excel in the same sidebar section. "
     "This is the most time-sensitive one — it contains **Last Free Detention** and "
     "**Last Free Demurrage** deadlines. The Insights tab will flag any containers "
     "with deadlines within 7 days."),
    ("4. Check the Insights tab", "Tab 6",
     "Go to **📈 Insights**. Review the **Action Items** section at the bottom first — "
     "it surfaces the highest-priority flags across all data sources. Then review "
     "Detention & Demurrage Risk and the LFD Risk table."),
    ("5. Look up containers as needed", "Tab 1",
     "Paste any list of container IDs into **🔍 Container Lookup**. Dashes, no dashes, "
     "check digit or not — the tool normalizes all formats. Results show which DBR sheet "
     "each container appears on and its current status."),
]

for title, location, desc in steps:
    with st.expander(f"**{title}** _(via {location})_"):
        st.markdown(desc)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — TAB REFERENCE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## Tab Reference", anchor="tab-reference")

tab_ref = st.tabs([
    "🔍 Container Lookup",
    "📤 Carrier Submission",
    "📋 Empty Returns",
    "📊 Carrier Data",
    "🎛️ Allocation",
    "📈 Insights",
])

with tab_ref[0]:
    st.markdown("### Container Lookup")
    st.markdown("""
**What it does:** Searches the loaded DBR for any container ID you paste in. Returns the sheet, 
status, terminal, FC/Building, appointment date, and LFD for every match.

**How to use:**
1. Paste container IDs in the text box — one per line, or comma-separated
2. Format doesn't matter: `TCNU3899024`, `TCNU389902-4`, and `TCNU389902` all match the same container
3. Enter your name (optional — logged for audit trail)
4. Click **Search**
5. Download results as Excel if needed

**Searches across:** Delivery Appointments · Empty Returns · On Vessel · Canceled · Demurrage · Accessorials

**Not found?** The container isn't in the current DBR. It may be delivered, not yet created, or on a different vessel not yet in the system.
    """)

with tab_ref[1]:
    st.markdown("### Carrier Submission Portal")
    st.markdown(f"""
**What it does:** Two ways for carriers to submit their weekly container status:
1. **Template upload** (recommended) — upload a filled AGL Carrier Template directly in this tab
2. **Vendor Portal** — share `{VENDOR_URL}` with the carrier; they go there directly, no login required

**Template sheets:**
| Sheet | Used for |
|-------|----------|
| Delivery | Standard container deliveries |
| ILM1 | ILM1-specific deliveries |
| RIC6 | RIC6-specific deliveries |
| Empty Return | Empty container return appointments |
| ODY/Storage | Containers in off-dock storage |
| Demurrage | Containers accruing demurrage charges |
| Accessorials | Accessorial charge records |

**Submission log** at the bottom of the tab shows all submitted containers with filter options. Export to Excel any time.

**Note:** The legacy manual form (below the template uploader) still works but the template is preferred — it captures more structured data.
    """)

with tab_ref[2]:
    st.markdown("### Empty Returns Dashboard")
    st.markdown("""
**What it does:** Dedicated view of the DBR's Empty Returns sheet with due-date tracking and risk flagging.

**Key features:**
- **4 KPI cards:** Active, Overdue, Due ≤3 days, On track
- **Radio filter:** switch between Overdue / Due Soon / All Active / All incl. Terminated
- Sorted by days until due — most urgent at top
- Export the filtered view to Excel

**Color coding:**
- 🔴 OVERDUE — empty return is past due
- 🟡 Due soon — 3 days or less remaining
- 🟢 OK — more than 3 days remaining

**Terminated section** at the bottom shows containers with TERMINATED status for record-keeping.
    """)

with tab_ref[3]:
    st.markdown("### Carrier Data")
    st.markdown("""
**What it does:** Structured view of everything submitted by carriers through the template portal 
or vendor portal. Organized by sheet type.

**Filters:** Carrier name · Sheet type · Within SLA  

**Tabs inside the view:** Each sheet type (Delivery, Empty Return, Demurrage, etc.) gets its own 
sub-tab, and only non-empty columns are shown.

**Export:** Full carrier data download available at the bottom.

**This tab grows over time** as carriers submit weekly. It's the historical record of all 
carrier-reported container status.
    """)

with tab_ref[4]:
    st.markdown("### Lane Cost Simulator")
    st.markdown("""
**What it does:** Uses the 2025–26 Robotics rate card to project dray costs by lane (port → destination).

**Rate Matrix (top):**
- Shows every active lane with all carrier options and rates side by side
- **Green** = cheapest carrier on that lane · **Red** = most expensive
- Filter by port of arrival, rate source, and lane type (Static / GF/BF / Backup)

**Scenario Builder:**
1. Select a **Port of Arrival** (USBOS, USEWR, USLAX, etc.)
2. Select a **Destination Node** (A100 = Tighe Woburn, A310 = DCB2 Mansfield, etc.)
3. Pick a **Carrier** — dropdown shows all carriers serving that lane with their rate
4. Enter **container count**
5. Click **Add Lane** — it builds up a running portfolio

**Portfolio view:**
- Edit container counts inline after adding
- Cost by carrier breakdown + bar chart
- Cost by port breakdown + bar chart
- **Savings vs Cheapest** — shows how much you'd save by switching to the cheapest available carrier on each lane

**Tip:** Use "Static" lanes for your primary allocation, "GF/BF" and "Backup" for overflow or contingency scenarios.
    """)

with tab_ref[5]:
    st.markdown("### Insights")
    st.markdown("""
**What it does:** Auto-generated weekly intelligence across all loaded data sources. Designed to 
answer "what needs my attention right now?" without manual analysis.

**Sections:**

**DBR Snapshot**
KPI cards for all active containers. LFD Risk table shows any delivery appointment 
with Last Free Day within 5 days. Status and terminal breakdowns.

**Detention & Demurrage Risk** _(requires Inbound Loads Report)_
The most operationally critical section. Flags containers where:
- **Detention** free time expires within 7 days (container available at terminal but not yet picked up)
- **Demurrage** free time expires within 7 days (container still sitting at port/on vessel)

Both accrue daily fees once the free period ends. The earlier you see this, the cheaper.

**In-Transit Pipeline** _(requires Import Shipment Status)_
Active containers by status, FC destination, and carrier. Includes an "Arriving at Port in 30 days" 
table so you can anticipate upcoming dray volume.

**Carrier Intelligence**
SLA rate per carrier (color-coded), submission volume, and a freshness check — flags carriers 
who haven't submitted in 7+ days.

**Action Items**
Auto-prioritized flags across all sources. Check this section first every week. 
Items are ranked: 🔴 Critical (costs accruing or imminent deadline) → 🟡 Urgent → 🟡 Attention.
    """)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — DATA SOURCES
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## Data Sources", anchor="data-sources")
st.markdown("""
The tool pulls from four data sources. Here's what each contributes and how to get it into the app.
""")

col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.markdown("#### 📁 DBR (Weekly)")
        st.markdown("""
**What it is:** The weekly Delivery Booking Report Excel file maintained by the drayage team.

**What it provides:** Container status across 6 sheets — Delivery Appointments, Empty Returns, 
On Vessel, CANCELED, Demurrage, Accessorials.

**How to load:** Sidebar → DBR File → Upload. Saved to S3 automatically.

**Refresh cadence:** Weekly (upload the new file each Monday or when received).
        """)

    with st.container(border=True):
        st.markdown("#### 📊 Import Shipment Status (Weekly)")
        st.markdown("""
**What it is:** A 73-column shipment lifecycle report from Amazon's procurement/TMS system 
covering all Robotics ocean containers.

**What it provides:** Container-level EDI dates (ingate, departure, arrival, available, 
out-gate, empty return, delivered, customs cleared), vessel/voyage, carrier code, 
discharge port ETA, destination FC, vendor name.

**How to load:** Sidebar → Additional Reports → Import Shipment Status → Upload.

**Source:** Emailed report — filename is typically "Import Shipment Status.xlsx".
        """)

with col2:
    with st.container(border=True):
        st.markdown("#### 📊 Inbound Loads Report (Weekly/Daily)")
        st.markdown("""
**What it is:** Amazon's internal supply chain report (AR Inbound Loads) with PO-level 
detail for every Robotics container in the pipeline.

**What it provides:** **Last Free Detention Time**, **Last Free Demurrage Time**, 
Gate-out Terminal Time, Estimated Appointment Date, actual arrival at final destination, 
empty container return time.

**How to load:** Sidebar → Additional Reports → Inbound Loads Report → Upload.

**Source:** Emailed report — filename includes a timestamp 
(e.g. "Amazon Robotics Inbound Loads Report23-Jul-2026 083003.xlsx").

**⚠️ Most time-sensitive source** — upload this first when it arrives.
        """)

    with st.container(border=True):
        st.markdown("#### 📤 Carrier Submissions (Ongoing)")
        st.markdown("""
**What it is:** Container status data submitted by dray carriers using the AGL Carrier Template.

**What it provides:** Per-container delivery status, SLA compliance, empty return dates, 
appointment dates, accessorial charges — as reported by the carrier.

**How to load:** Carriers submit via the **Vendor Portal** or you can upload directly 
in the **Carrier Submission tab**.

**Source:** Ongoing — carriers submit weekly. Data persists in the database.
        """)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — VENDOR PORTAL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## Vendor Portal", anchor="vendor-portal")
st.markdown(f"""
The vendor portal is a **public-facing, no-login page** for dray carriers to submit their 
weekly container status templates.

**URL to share with carriers:**
""")
st.code(VENDOR_URL, language=None)
st.markdown(f"""
**What carriers do:**
1. Go to the URL above
2. Download the AGL Carrier Template (button on the page)
3. Fill in their container data across the relevant sheets
4. Enter their company name and contact name
5. Upload the completed file
6. Preview and submit — data goes directly into the tracker

**What you get:**
- All submissions appear in the **Carrier Data tab** and the **Carrier Submission tab** log
- Submissions are tagged `source = vendor_portal` so you can filter them
- The Insights tab tracks which carriers have submitted recently

**Onboarding a new carrier:**
1. Send them the vendor portal URL and the AGL Carrier Template download link
2. On their first submission, their carrier name will auto-populate in filters
3. The Insights tab will start tracking their SLA performance as data accumulates

**Template sheets:** The template has 7 sheets — carriers only need to fill in the 
sheets relevant to their activity that week. Empty sheets are ignored on upload.
""")
st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — FAQ
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## FAQ", anchor="faq")

faqs = [
    ("The app is showing old DBR data. How do I refresh it?",
     "Upload the new DBR file in the sidebar. The app pulls from S3 on load, "
     "so if someone else on the team already uploaded a newer file, you'll see it "
     "on your next page refresh. If not, upload the new file — it overwrites the "
     "S3 copy and the cache refreshes automatically."),

    ("A container I'm looking for isn't showing up in the search.",
     "A few possible reasons: (1) It's not in the current week's DBR — check if "
     "it might be in a previous week. (2) The container ID format is unusual — try "
     "just the first 10 characters without any suffix. (3) It's already delivered "
     "and not included in the active DBR. If you need to search historical containers, "
     "that's a feature to add — note it and flag it."),

    ("A carrier says they submitted but I don't see it in the Carrier Data tab.",
     "Check the Carrier Submission tab's submission log and filter by carrier name. "
     "If it's not there, the submission may have failed — ask them to resubmit via the vendor portal. "
     "Common issues: they used an old template format, or they submitted without entering their company name."),

    ("The Detention/Demurrage Risk section is empty.",
     "That section requires the Inbound Loads Report to be uploaded. Upload it in the sidebar "
     "under Additional Reports. If it still shows nothing after upload, either all containers "
     "have already gated out (which is good) or the file didn't parse correctly — check that "
     "you're uploading the right file (AR Inbound Loads, not the Import Shipment Status)."),

    ("How do I share access to the main app with someone on the team?",
     f"Send them the app URL ({APP_URL}) and the password: **robotics2026**. "
     "They'll need to enter the password on first visit. The vendor portal does not "
     "require a password — it's public-facing by design."),

    ("The lane cost simulator says 'No rates on this lane' for a destination.",
     "The rate card only covers lanes where a dray bid has been completed. Some newer "
     "destination nodes (A322 and above) are marked 'NO - Need SIM' in the Robotics "
     "Lanes file and don't have rates yet. For those lanes, work with the sourcing team "
     "to get rates added to the rate card, then re-import."),

    ("Can I add a carrier to the rate card?",
     "Rate card data is managed by re-importing the rate card Excel files. If rates change "
     "or a new carrier is added, contact the app admin to run the import script. This is "
     "intentionally not a self-service operation to avoid accidental overrides."),

    ("How often does the data auto-refresh?",
     "The app does not auto-pull new reports. All data sources (DBR, Shipment Status, "
     "Inbound Loads) require a manual upload. The team is working on automatic email "
     "ingestion — once configured, reports will be processed within minutes of arriving "
     "in your inbox."),
]

for question, answer in faqs:
    with st.expander(f"**{question}**"):
        st.markdown(answer)

st.divider()
st.caption("For issues or feature requests, contact the app owner. Built by AGL Ops · July 2026.")
