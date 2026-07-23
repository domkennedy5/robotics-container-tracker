from __future__ import annotations
import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime

st.set_page_config(
    page_title="AGL Carrier Submission Portal",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Hide sidebar navigation entirely — carriers should never see the internal app
st.markdown("""
<style>
[data-testid="stSidebarNav"] { display: none; }
section[data-testid="stSidebar"] { display: none; }
[data-testid="collapsedControl"] { display: none; }
</style>
""", unsafe_allow_html=True)

# On Streamlit Cloud the CWD is always the repo root — reliable across all pages
BASE_DIR = os.getcwd()
DB_PATH  = os.path.join(BASE_DIR, "tracker.db")

# Shared modules are importable directly — Streamlit Cloud adds repo root to sys.path
from utils import parse_carrier_template
import data_sync

# ── AWS config ────────────────────────────────────────────────────────────────
def _aws_cfg():
    try:
        return (
            st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
            st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
            st.secrets["aws"]["AWS_REGION"],
            st.secrets["aws"]["S3_BUCKET"],
        )
    except Exception:
        return None, None, None, None

AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET = _aws_cfg()
S3_ENABLED = all([AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET])

# Pull latest DB from S3 once per session so submissions write to current data
if S3_ENABLED and not st.session_state.get("vendor_s3_done"):
    data_sync.pull_db_from_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
    st.session_state.vendor_s3_done = True

def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

# ── Page header ───────────────────────────────────────────────────────────────
st.markdown("## 📦 AGL Robotics — Carrier Submission Portal")
st.markdown("Use this page to submit your weekly container status update to the AGL team.")
st.divider()

# ── Step 1: Download template ─────────────────────────────────────────────────
st.markdown("### Step 1 — Download the template")
st.markdown(
    "If you don't already have the AGL Carrier Template, download it here. "
    "Fill in the sheets that apply to your containers this week."
)

tmpl_path = os.path.join(BASE_DIR, "agl_carrier_template.xlsx")
if os.path.exists(tmpl_path):
    with open(tmpl_path, "rb") as f:
        tmpl_bytes = f.read()
    st.download_button(
        "⬇️ Download AGL Carrier Template (.xlsx)",
        data=tmpl_bytes,
        file_name="AGL_Carrier_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
else:
    st.info("Template file not found — contact your AGL point of contact for the template.")

st.divider()

# ── Step 2: Fill in info and upload ──────────────────────────────────────────
st.markdown("### Step 2 — Upload your completed template")

carrier_name = st.text_input(
    "Company / Carrier Name *",
    placeholder="e.g. Arrive Logistics, Atlas Marine, Roadone…",
)
submitter = st.text_input(
    "Your Name *",
    placeholder="e.g. Jane Smith",
)
carrier_file = st.file_uploader(
    "Upload completed AGL Carrier Template *",
    type=["xlsx", "csv"],
    help="Must use the AGL Carrier Template format.",
)

# ── Validation & preview ──────────────────────────────────────────────────────
if carrier_file and not carrier_name.strip():
    st.warning("Please enter your company name before submitting.")

elif carrier_file and not submitter.strip():
    st.warning("Please enter your name before submitting.")

elif carrier_file and carrier_name.strip() and submitter.strip():
    # cache bytes so re-renders don't exhaust the file object
    cache_key = f"vp_{carrier_file.name}_{carrier_file.size}"
    if cache_key not in st.session_state:
        st.session_state[cache_key] = carrier_file.read()
    raw = st.session_state[cache_key]

    with st.spinner("Reading file…"):
        parsed = parse_carrier_template(raw, carrier_file.name)

    if not parsed:
        st.error(
            "No container data found in this file. "
            "Make sure you are using the AGL Carrier Template and that at least one sheet is filled in."
        )
    else:
        total = sum(len(v) for v in parsed.values())
        st.success(
            f"✅ Found **{total} containers** across **{len(parsed)} sheet(s)**. "
            "Review below, then click Submit."
        )

        sheet_labels = {
            "delivery":      "📦 Delivery",
            "delivery_ilm1": "🏭 ILM1",
            "delivery_ric6": "🏭 RIC6",
            "empty_return":  "🔁 Empty Returns",
            "ody":           "🏗️ Storage/ODY",
            "demurrage":     "⚠️ Demurrage",
            "accessorial":   "💰 Accessorials",
        }
        preview_tabs = st.tabs([sheet_labels.get(k, k) for k in parsed.keys()])
        for ptab, (stype, rows) in zip(preview_tabs, parsed.items()):
            with ptab:
                df_prev = pd.DataFrame(rows).drop(columns=["_raw_container"], errors="ignore")
                st.dataframe(df_prev, use_container_width=True, hide_index=True)

        st.divider()
        if st.button("✅ Confirm & Submit", type="primary", use_container_width=True):
            now  = datetime.now().isoformat()
            conn = get_db()
            count = 0
            try:
                for stype, rows in parsed.items():
                    for row in rows:
                        conn.execute(
                            """INSERT INTO carrier_submissions
                               (submitted_at, carrier_name, container_id, sheet_type,
                                port, terminal, fc_building, flexi_id, outgate_date,
                                delivery_date, status, within_sla, sla_notes,
                                empty_return_due, appointment_date, accessorial_type,
                                notes, source_file, source)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (now, carrier_name.strip(), row["container_id"], stype,
                             row.get("port"), row.get("terminal"), row.get("fc_building"),
                             row.get("flexi_id"), row.get("outgate_date"),
                             row.get("delivery_date"), row.get("status"),
                             row.get("within_sla"), row.get("sla_notes"),
                             row.get("empty_return_due"), row.get("appointment_date"),
                             row.get("accessorial_type"), row.get("notes"),
                             carrier_file.name, "vendor_portal"),
                        )
                        count += 1
                conn.commit()
            except Exception as e:
                st.error(f"Submission failed — please try again or contact your AGL point of contact. ({e})")
            finally:
                conn.close()

            if count > 0:
                if S3_ENABLED:
                    data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                st.balloons()
                st.success(
                    f"**Submission received.** {count} containers logged from **{carrier_name}**,  \n"
                    f"submitted by {submitter} on {datetime.now().strftime('%B %d, %Y at %H:%M UTC')}."
                )
                st.info("You can close this window. The AGL team will review your submission.")
                if cache_key in st.session_state:
                    del st.session_state[cache_key]

st.divider()
st.caption("AGL Robotics · Questions? Contact your AGL point of contact.")
