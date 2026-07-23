from __future__ import annotations
import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime

# must be first st call
st.set_page_config(
    page_title="AGL Carrier Submission Portal",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── import shared helpers from parent app ────────────────────────────────────
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils import parse_carrier_template
import data_sync

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH  = os.path.join(BASE_DIR, "tracker.db")

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

# ── pull latest DB from S3 on first load ─────────────────────────────────────
if S3_ENABLED and not st.session_state.get("vendor_s3_done"):
    data_sync.pull_db_from_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
    st.session_state.vendor_s3_done = True

def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

# ── UI ────────────────────────────────────────────────────────────────────────
st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Amazon_logo.svg/320px-Amazon_logo.svg.png", width=120)
st.title("AGL Robotics — Carrier Submission Portal")
st.caption("Submit your weekly container status update using the AGL Carrier Template.")

st.divider()

# template download
tmpl_path = os.path.join(BASE_DIR, "agl_carrier_template.xlsx")
if os.path.exists(tmpl_path):
    with open(tmpl_path, "rb") as f:
        st.download_button(
            "⬇️ Download the AGL Carrier Template",
            data=f.read(),
            file_name="AGL_Carrier_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    st.caption("Fill in the template, then upload your completed file below.")

st.divider()

# submission form
carrier_name = st.text_input(
    "Company / Carrier Name *",
    placeholder="e.g. Arrive Logistics, Atlas Marine, Roadone…",
)
submitter    = st.text_input("Your Name *", placeholder="e.g. Jane Smith")
carrier_file = st.file_uploader(
    "Upload completed AGL Carrier Template *",
    type=["xlsx", "csv"],
    help="Only .xlsx or .csv accepted. Must use the AGL template format.",
)

if carrier_file and carrier_name.strip() and submitter.strip():
    cache_key = f"vp_bytes_{carrier_file.name}_{carrier_file.size}"
    if cache_key not in st.session_state:
        st.session_state[cache_key] = carrier_file.read()
    raw = st.session_state[cache_key]
    parsed = parse_carrier_template(raw, carrier_file.name)

    if not parsed:
        st.error("No container data found. Make sure you are using the AGL Carrier Template and the sheets are filled in.")
    else:
        total = sum(len(v) for v in parsed.values())
        st.success(f"Found **{total} containers** across **{len(parsed)} sheet(s)**. Review the preview below, then submit.")

        sheet_labels = {
            "delivery":      "📦 Delivery",
            "delivery_ilm1": "🏭 ILM1",
            "delivery_ric6": "🏭 RIC6",
            "empty_return":  "🔁 Empty Returns",
            "ody":           "🏗️ Storage/ODY",
            "demurrage":     "⚠️ Demurrage",
            "accessorial":   "💰 Accessorials",
        }
        ptabs = st.tabs([sheet_labels.get(k, k) for k in parsed.keys()])
        for ptab, (stype, rows) in zip(ptabs, parsed.items()):
            with ptab:
                df_prev = pd.DataFrame(rows).drop(columns=["_raw_container"], errors="ignore")
                st.dataframe(df_prev, use_container_width=True, hide_index=True)

        if st.button("✅ Submit", type="primary", use_container_width=True):
            now  = datetime.now().isoformat()
            conn = get_db()
            count = 0
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
            conn.close()
            if S3_ENABLED:
                data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
            st.balloons()
            st.success(
                f"✅ **Submission received** — {count} containers logged from **{carrier_name}**. "
                f"Submitted by {submitter} at {datetime.now().strftime('%b %d, %Y %H:%M UTC')}."
            )
            st.info("You can close this window. The AGL team has been notified.")
            # clear cache so user can submit again fresh
            if cache_key in st.session_state:
                del st.session_state[cache_key]

elif carrier_file and not (carrier_name.strip() and submitter.strip()):
    st.warning("Enter your company name and your name before uploading.")

st.divider()
st.caption("Questions? Contact your AGL point of contact. This portal is for AGL Robotics carriers only.")
