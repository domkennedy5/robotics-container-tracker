from __future__ import annotations
import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime, date, timedelta

st.set_page_config(
    page_title="AGL — Delivery Schedule",
    page_icon="\U0001f4e6",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Hide sidebar nav — carriers should never see internal app pages
st.markdown("""
<style>
[data-testid="stSidebarNav"]  { display: none; }
section[data-testid="stSidebar"] { display: none; }
[data-testid="collapsedControl"] { display: none; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = os.getcwd()
DB_PATH  = os.path.join(BASE_DIR, "tracker.db")

import data_sync

# ── AWS / S3 ──────────────────────────────────────────────────────────────────
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

if S3_ENABLED and not st.session_state.get("cp_s3_pulled"):
    data_sync.pull_db_from_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
    st.session_state.cp_s3_pulled = True

def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

# ── Token auth ────────────────────────────────────────────────────────────────
_token = st.query_params.get("token", "")

if not _token:
    st.error("\U0001f512 No access token provided. Please use the link sent by your AGL coordinator.")
    st.stop()

conn = get_db()
_tok_row = conn.execute(
    "SELECT scac FROM plan_carrier_tokens WHERE token=?", (_token,)
).fetchone()
conn.close()

if not _tok_row:
    st.error("\U0001f6ab Invalid or expired access token. Please contact your AGL coordinator for a new link.")
    st.stop()

_scac = _tok_row["scac"]

# ── Carrier metadata ──────────────────────────────────────────────────────────
conn = get_db()
_carr_row = conn.execute(
    "SELECT carrier_name, ops_contact FROM plan_carriers WHERE scac=?", (_scac,)
).fetchone()
conn.close()

_cname    = _carr_row["carrier_name"]    if _carr_row else _scac
_ccontact = _carr_row["ops_contact"]     if _carr_row else ""

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"## \U0001f4e6 AGL Robotics — Delivery Schedule")
st.markdown(
    f"**{_cname} ({_scac})**"
    + (f"  \u00b7  Contact: {_ccontact}" if _ccontact else "")
)
st.caption(
    "Review your scheduled container deliveries below. "
    "Check the box next to each container to confirm your appointment, "
    "add a note if needed, then click **Submit Confirmations**."
)
st.divider()

# ── Week selector ─────────────────────────────────────────────────────────────
_today = date.today()
# Default to next Monday (or current Mon if today is Mon)
_days_to_mon = (7 - _today.weekday()) % 7
_next_mon    = _today + timedelta(days=_days_to_mon if _days_to_mon > 0 else 7)

_wk_pick = st.date_input(
    "Select week (any day in the week)", value=_next_mon, key="cp_week_pick"
)
_ws = _wk_pick - timedelta(days=_wk_pick.weekday())  # normalize to Monday
_we = _ws + timedelta(days=6)
st.caption(f"Showing plan for **{_ws.strftime('%B %d').lstrip('0')} \u2013 {_we.strftime('%B %d, %Y').lstrip('0')}**")

# ── Load plan ─────────────────────────────────────────────────────────────────
conn = get_db()
_plan_df = pd.read_sql(
    """SELECT id, appt_date, appt_time, container_id, site_code,
              product_type, qty, notes,
              COALESCE(carrier_confirmed,0) AS carrier_confirmed,
              confirmed_at, carrier_note
       FROM delivery_plan
       WHERE carrier=? AND week_start=? AND status != 'PENDING'
       ORDER BY appt_date, slot_num""",
    conn, params=[_scac, _ws.isoformat()]
)
conn.close()

if _plan_df.empty:
    st.info(
        f"No scheduled deliveries found for {_cname} for the week of "
        f"{_ws.strftime('%B %d, %Y').lstrip('0')}.  \n"
        f"If you believe this is an error, contact your AGL coordinator."
    )
    st.stop()

# ── Build editable confirmation table ────────────────────────────────────────
def _fmt_day(d):
    try:
        return date.fromisoformat(str(d)).strftime("%a %m/%d")
    except Exception:
        return "TBD"

_disp_df = pd.DataFrame({
    "id":        _plan_df["id"],
    "Confirm":   _plan_df["carrier_confirmed"].astype(bool),
    "Day":       _plan_df["appt_date"].apply(_fmt_day),
    "Site":      _plan_df["site_code"].fillna("").astype(str),
    "Appt Time": _plan_df["appt_time"].fillna("").astype(str),
    "Container": _plan_df["container_id"].astype(str),
    "Product":   _plan_df["product_type"].fillna("").astype(str),
    "Your Notes":_plan_df["carrier_note"].fillna("").astype(str),
})

# Metrics row
_n_conf  = int(_disp_df["Confirm"].sum())
_n_total = len(_disp_df)
mc1, mc2, mc3 = st.columns(3)
with mc1: st.metric("Total deliveries", _n_total)
with mc2: st.metric("Confirmed", _n_conf)
with mc3: st.metric("Awaiting confirmation", _n_total - _n_conf)

st.divider()

_edited = st.data_editor(
    _disp_df.drop(columns=["id"]),
    column_config={
        "Confirm":    st.column_config.CheckboxColumn("\u2705 Confirm", help="Check to confirm this appointment"),
        "Day":        st.column_config.TextColumn("Day",         disabled=True),
        "Site":       st.column_config.TextColumn("Site",        disabled=True, width="small"),
        "Appt Time":  st.column_config.TextColumn("Appt Time",   disabled=True, width="small"),
        "Container":  st.column_config.TextColumn("Container #", disabled=True),
        "Product":    st.column_config.TextColumn("Product",     disabled=True),
        "Your Notes": st.column_config.TextColumn("Your Notes",  help="Optional — flag issues, substitutions, etc."),
    },
    use_container_width=True,
    hide_index=True,
    key="cp_editor",
    height=min(500, 38 + _n_total * 38),
)

st.divider()

# ── Submit ─────────────────────────────────────────────────────────────────────
_sub_c1, _sub_c2 = st.columns([2, 3])
with _sub_c1:
    if st.button("\U0001f4be Submit Confirmations", type="primary", use_container_width=True):
        conn = get_db()
        _now = datetime.now().isoformat()
        _updated = 0
        for i, row in _edited.iterrows():
            _rid  = int(_disp_df.iloc[i]["id"])
            _conf = 1 if row["Confirm"] else 0
            _note = str(row["Your Notes"]).strip()
            _cat  = _now if _conf else None
            conn.execute(
                "UPDATE delivery_plan SET carrier_confirmed=?, confirmed_at=?, carrier_note=?, updated_at=? WHERE id=?",
                (_conf, _cat, _note, _now, _rid)
            )
            _updated += 1
        conn.commit()
        conn.close()
        # Push to S3 so admin sees confirmations immediately
        if S3_ENABLED:
            try:
                data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
            except Exception as _e:
                pass  # Non-fatal — DB still updated locally
        _newly_conf = int(_edited["Confirm"].sum())
        st.success(
            f"\u2705 Saved. {_newly_conf} container(s) confirmed, "
            f"{_n_total - _newly_conf} unconfirmed."
        )
        st.session_state.cp_s3_pulled = False  # force re-pull on next load
        st.rerun()

with _sub_c2:
    st.caption(
        "Confirmations are saved immediately and visible to your AGL coordinator. "
        "You can return to this link to update at any time."
    )

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "AGL Robotics Logistics  \u00b7  "
    "Questions? Reply to the email from your coordinator or contact AGL Inbound directly."
)
