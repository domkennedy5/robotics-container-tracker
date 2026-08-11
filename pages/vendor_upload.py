from __future__ import annotations
import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

st.set_page_config(
    page_title="AGL Carrier Submission Portal",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Hide sidebar — carriers never see the internal app
st.markdown("""
<style>
[data-testid="stSidebarNav"]  { display: none; }
section[data-testid="stSidebar"] { display: none; }
[data-testid="collapsedControl"] { display: none; }
</style>
""", unsafe_allow_html=True)

_PORTAL_PASSWORD = "ARcarrier"
_PHX = ZoneInfo("America/Phoenix")

BASE_DIR = os.getcwd()
DB_PATH  = os.path.join(BASE_DIR, "tracker.db")

from utils import parse_carrier_template
import data_sync

# ── AWS ───────────────────────────────────────────────────────────────────────
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

if S3_ENABLED and not st.session_state.get("vendor_s3_done"):
    data_sync.pull_db_from_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
    st.session_state.vendor_s3_done = True

def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

# ── Password gate ─────────────────────────────────────────────────────────────
if not st.session_state.get("portal_auth"):
    st.markdown("## 📦 AGL Robotics — Carrier Portal")
    st.markdown("Enter the portal password provided by your AGL point of contact.")
    _pw_input = st.text_input("Password", type="password", key="portal_pw_field")
    if st.button("Enter", type="primary"):
        if _pw_input == _PORTAL_PASSWORD:
            st.session_state.portal_auth = True
            st.rerun()
        else:
            st.error("Incorrect password. Contact your AGL point of contact if you need access.")
    st.stop()

# ── Authenticated: portal ─────────────────────────────────────────────────────
st.markdown("## 📦 AGL Robotics — Carrier Submission Portal")
st.markdown("Submit your daily container status update to the AGL team.")
st.divider()

# Step 1 — Template download
st.markdown("### Step 1 — Download the template (if needed)")
st.markdown(
    "Use the AGL Carrier Template to ensure your data is parsed correctly. "
    "Fill in the sheets that apply to your containers today."
)
tmpl_path = os.path.join(BASE_DIR, "agl_carrier_template.xlsx")
if os.path.exists(tmpl_path):
    with open(tmpl_path, "rb") as _f:
        st.download_button(
            "⬇️ Download AGL Carrier Template (.xlsx)",
            data=_f.read(),
            file_name="AGL_Carrier_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
else:
    st.info("Template not on file — contact your AGL point of contact.")

st.divider()

# Step 2 — Carrier info + upload
st.markdown("### Step 2 — Enter your carrier details and upload your DBR")

_col1, _col2 = st.columns(2)
with _col1:
    scac_raw = st.text_input(
        "Your SCAC Code *",
        placeholder="e.g. ATMI, ARVY, AOYV, HDDR…",
        help="Enter your 4-letter Standard Carrier Alpha Code exactly as assigned.",
        key="vp_scac",
    )
    scac_code = scac_raw.strip().upper() if scac_raw.strip() else ""
with _col2:
    submitter = st.text_input(
        "Your Name *",
        placeholder="e.g. Jane Smith",
        key="vp_submitter",
    )

file_date = st.date_input(
    "Date this data reflects *",
    value=datetime.now(_PHX).date(),
    key="vp_file_date",
    help="Which day's status does this file represent? Defaults to today (MST).",
)

carrier_file = st.file_uploader(
    "Upload your daily DBR file *",
    type=["xlsx", "csv"],
    help="Excel or CSV. The AGL template is preferred; your native format is also accepted.",
    key="vp_file",
)

# ── Validation ─────────────────────────────────────────────────────────────────
_ready = carrier_file and scac_code and submitter.strip()

if carrier_file and not scac_code:
    st.warning("Enter your SCAC code before submitting.")
if carrier_file and not submitter.strip():
    st.warning("Enter your name before submitting.")

if _ready:
    _cache_key = f"vp_{carrier_file.name}_{carrier_file.size}"
    if _cache_key not in st.session_state:
        st.session_state[_cache_key] = carrier_file.read()
    raw = st.session_state[_cache_key]

    with st.spinner("Reading file…"):
        parsed = parse_carrier_template(raw, carrier_file.name)

    if not parsed:
        st.error(
            "No container data found. "
            "Make sure the file uses the AGL Carrier Template format, or that at least one sheet contains container IDs."
        )
    else:
        total_rows = sum(len(v) for v in parsed.values())
        st.success(f"✅ Found **{total_rows} containers** across **{len(parsed)} sheet(s)**. Review below, then confirm.")

        _sheet_labels = {
            "delivery":       "📦 Delivery",
            "delivery_ilm1":  "🏭 ILM1",
            "delivery_ric6":  "🏭 RIC6",
            "delivery_dbm6":  "🏭 DBM6",
            "empty_return":   "🔁 Empty Returns",
            "ody":            "🏗️ Storage/ODY",
            "demurrage":      "⚠️ Demurrage",
            "accessorial":    "💰 Accessorials",
        }
        _preview_tabs = st.tabs([_sheet_labels.get(k, k) for k in parsed.keys()])
        for _ptab, (_stype, _rows) in zip(_preview_tabs, parsed.items()):
            with _ptab:
                _df_prev = pd.DataFrame(_rows).drop(columns=["_raw_container"], errors="ignore")
                st.dataframe(_df_prev, use_container_width=True, hide_index=True)

        st.divider()

        if st.button("✅ Confirm & Submit", type="primary", use_container_width=True, key="vp_submit_btn"):
            _now     = datetime.now(_PHX).isoformat()
            _today   = file_date.isoformat()
            _wk_start = (file_date - timedelta(days=file_date.weekday())).isoformat()
            _conn    = get_db()
            _count   = 0

            # Delivery sheet types that feed inbound_containers
            _DELIVERY_TYPES = {"delivery", "delivery_ilm1", "delivery_ric6", "delivery_dbm6"}
            # Map sheet type → fc_dest override when fc_building field is absent
            _SHEET_FC = {
                "delivery_ilm1": "ILM1",
                "delivery_ric6": "RIC6",
                "delivery_dbm6": "DBM6",
            }

            try:
                for _stype, _rows in parsed.items():
                    for _row in _rows:
                        # ── 1. carrier_submissions (existing log) ───────────────
                        _conn.execute("""
                            INSERT INTO carrier_submissions
                                (submitted_at, carrier_name, container_id, sheet_type,
                                 port, terminal, fc_building, flexi_id, outgate_date,
                                 delivery_date, status, within_sla, sla_notes,
                                 empty_return_due, appointment_date, accessorial_type,
                                 notes, source_file, source)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """, (
                            _today, scac_code, _row["container_id"], _stype,
                            _row.get("port"), _row.get("terminal"), _row.get("fc_building"),
                            _row.get("flexi_id"), _row.get("outgate_date"),
                            _row.get("delivery_date"), _row.get("status"),
                            _row.get("within_sla"), _row.get("sla_notes"),
                            _row.get("empty_return_due"), _row.get("appointment_date"),
                            _row.get("accessorial_type"), _row.get("notes"),
                            carrier_file.name, "vendor_portal",
                        ))
                        _count += 1

                        # ── 2. inbound_containers (lifecycle tracker) ───────────
                        if _stype in _DELIVERY_TYPES:
                            _has_del = bool(_row.get("delivery_date"))
                            _has_og  = bool(_row.get("outgate_date"))
                            if _has_del:
                                _ic_status = "delivered"
                                _fc_actual = str(_row["delivery_date"])
                            elif _has_og:
                                _ic_status = "at_yard"
                                _fc_actual = None
                            else:
                                _ic_status = "pending"
                                _fc_actual = None

                            _fc_dest = _SHEET_FC.get(_stype) or _row.get("fc_building")

                            _conn.execute("""
                                INSERT INTO inbound_containers
                                    (container_id, carrier, port_region, fc_dest,
                                     yard_actual, fc_actual, status, notes,
                                     last_updated, source_file)
                                VALUES (?,?,?,?,?,?,?,?,?,?)
                                ON CONFLICT(container_id) DO UPDATE SET
                                    carrier     = COALESCE(excluded.carrier,     carrier),
                                    port_region = COALESCE(excluded.port_region, port_region),
                                    fc_dest     = COALESCE(excluded.fc_dest,     fc_dest),
                                    yard_actual = COALESCE(excluded.yard_actual, yard_actual),
                                    fc_actual   = COALESCE(excluded.fc_actual,   fc_actual),
                                    status      = excluded.status,
                                    last_updated = excluded.last_updated,
                                    source_file  = excluded.source_file
                            """, (
                                _row["container_id"], scac_code, _row.get("port"),
                                _fc_dest, _row.get("outgate_date"), _fc_actual,
                                _ic_status, _row.get("notes"), _now, carrier_file.name,
                            ))

                        elif _stype == "empty_return":
                            # Mark as empty_returned if status indicates return complete
                            _ret_status = str(_row.get("status") or "").lower()
                            _is_ret = any(kw in _ret_status for kw in
                                          ("return", "complete", "done", "returned", "at port"))
                            if _is_ret:
                                _ret_date = _row.get("appointment_date") or _today
                                _conn.execute("""
                                    UPDATE inbound_containers SET
                                        empty_returned = COALESCE(empty_returned, ?),
                                        status         = 'empty_returned',
                                        last_updated   = ?
                                    WHERE container_id = ?
                                """, (_ret_date, _now, _row["container_id"]))

                # ── 3. dbr_receipts (daily compliance log) ─────────────────────
                _conn.execute("""
                    INSERT OR IGNORE INTO dbr_receipts
                        (carrier, week_start, received_date, received_via, file_name, logged_at)
                    VALUES (?,?,?,?,?,?)
                """, (scac_code, _wk_start, _today, "vendor_portal", carrier_file.name, _now))

                _conn.commit()

            except Exception as _e:
                _conn.rollback()
                st.error(f"Submission failed — please try again or contact AGL. ({_e})")
                st.stop()
            finally:
                _conn.close()

            if _count > 0:
                if S3_ENABLED:
                    data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                st.balloons()
                st.success(
                    f"**Submission received.** {_count} container records logged from **{scac_code}**,  \n"
                    f"submitted by {submitter.strip()} — {datetime.now(_PHX).strftime('%B %d, %Y %H:%M MST')}."
                )
                st.info("You can close this window. The AGL team will review your submission.")
                if _cache_key in st.session_state:
                    del st.session_state[_cache_key]

st.divider()
st.caption("AGL Robotics Logistics · Questions? Contact your AGL point of contact.")
