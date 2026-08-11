from __future__ import annotations
import streamlit as st
import pandas as pd
import sqlite3
import os
import hashlib
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
[data-testid="stSidebarNav"]     { display: none; }
section[data-testid="stSidebar"] { display: none; }
[data-testid="collapsedControl"] { display: none; }
</style>
""", unsafe_allow_html=True)

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

# ── Per-carrier password gate ─────────────────────────────────────────────────
# Each carrier has a unique password stored as sha256 in carrier_contacts.portal_password.
# On login the hash is matched and the SCAC auto-populates — no manual SCAC entry.
if not st.session_state.get("portal_auth"):
    st.markdown("## 📦 AGL Robotics — Carrier Portal")
    st.markdown("Enter the portal password provided by your AGL point of contact.")
    _pw = st.text_input("Password", type="password", key="portal_pw_field")
    if st.button("Enter", type="primary"):
        if _pw.strip():
            _hashed = hashlib.sha256(_pw.strip().encode()).hexdigest()
            _aconn = get_db()
            _row = _aconn.execute(
                "SELECT scac, display_name, contact_name FROM carrier_contacts "
                "WHERE portal_password = ? LIMIT 1",
                (_hashed,),
            ).fetchone()
            _aconn.close()
            if _row:
                st.session_state.portal_auth         = True
                st.session_state.portal_scac         = _row["scac"]
                st.session_state.portal_display_name = _row["display_name"] or _row["scac"]
                st.session_state.portal_contact_name = _row["contact_name"] or ""
                st.rerun()
            else:
                st.error("Incorrect password. Contact your AGL point of contact if you need access.")
        else:
            st.error("Please enter a password.")
    st.stop()

# ── Authenticated ─────────────────────────────────────────────────────────────
_portal_scac    = st.session_state.portal_scac
_portal_display = st.session_state.portal_display_name

st.markdown("## 📦 AGL Robotics — Carrier Submission Portal")
st.markdown("Submit your daily container status update to the AGL team.")
st.info("🏷️  Logged in as **" + _portal_display + "** (" + _portal_scac + ")")
st.divider()

# Step 1 — Template
st.markdown("### Step 1 — Download the template (if needed)")
st.markdown("Use the AGL Carrier Template to ensure your data is parsed correctly.")
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

# ── Uploader key counter ──────────────────────────────────────────────────────
if "vp_uploader_key" not in st.session_state:
    st.session_state.vp_uploader_key = 0

# ── Acknowledge-and-clear banner ──────────────────────────────────────────────
if st.session_state.get("vp_submit_result"):
    _res = st.session_state.vp_submit_result
    _fnames = ", ".join(_res.get("files", []))
    st.success(
        "✅  **Submission received.**\n\n"
        "**" + str(_res["count"]) + "** container records logged from **" + _res["scac"] + "**  \n"
        "Submitted by " + _res["submitter"] + " — " + _res["timestamp"] + "  \n"
        "Files: " + _fnames
    )
    st.info("Your data has been captured and stored. The AGL team will review your submission.")
    if st.button("✔️  Acknowledge & Clear", type="primary", use_container_width=True, key="vp_ack"):
        del st.session_state["vp_submit_result"]
        st.session_state.vp_uploader_key += 1
        # Clear cached file bytes
        for k in [k for k in st.session_state if k.startswith("vp_bytes_")]:
            del st.session_state[k]
        st.rerun()
else:
    # Step 2 — Upload
    st.markdown("### Step 2 — Upload your DBR file(s)")

    file_date = st.date_input(
        "Date this data reflects *",
        value=datetime.now(_PHX).date(),
        key="vp_file_date",
        help="Which day does this file represent? Defaults to today (MST).",
    )

    carrier_files = st.file_uploader(
        "Upload your daily DBR file(s) *",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        help="Excel or CSV. Drop multiple files at once if submitting for more than one site or date.",
        key="vp_files_" + str(st.session_state.vp_uploader_key),
    )

    # ── Validation ────────────────────────────────────────────────────────────
    _ready = bool(carrier_files)

    if _ready:
        # Cache bytes per file so re-renders don't exhaust the file objects
        _all_parsed = {}
        _parse_errs = []
        for _cf in carrier_files:
            _ck = "vp_bytes_" + _cf.name + "_" + str(_cf.size)
            if _ck not in st.session_state:
                st.session_state[_ck] = _cf.read()
            _raw = st.session_state[_ck]
            with st.spinner("Reading " + _cf.name + "…"):
                try:
                    _p = parse_carrier_template(_raw, _cf.name)
                    if _p:
                        for _stype, _rows in _p.items():
                            _all_parsed.setdefault(_stype, [])
                            for _r in _rows:
                                _r["_src"] = _cf.name
                            _all_parsed[_stype].extend(_rows)
                    else:
                        _parse_errs.append(_cf.name + ": no container data found")
                except Exception as _e:
                    _parse_errs.append(_cf.name + ": " + str(_e))

        for _err in _parse_errs:
            st.warning(_err)

        if not _all_parsed:
            st.error(
                "No container data found in the uploaded file(s). "
                "Ensure you are using the AGL Carrier Template format and at least one sheet is filled in."
            )
        else:
            _total = sum(len(v) for v in _all_parsed.values())
            _n_files = len(carrier_files)
            st.success(
                "✅ Found **" + str(_total) + " containers** across **" + str(len(_all_parsed)) + " sheet(s)** "
                "from **" + str(_n_files) + " file(s)**. Review below, then confirm."
            )

            _sheet_labels = {
                "delivery":      "📦 Delivery",
                "delivery_ilm1": "🏭 ILM1",
                "delivery_ric6": "🏭 RIC6",
                "delivery_dbm6": "🏭 DBM6",
                "empty_return":  "🔁 Empty Returns",
                "ody":           "🏗️ Storage/ODY",
                "demurrage":     "⚠️ Demurrage",
                "accessorial":   "💰 Accessorials",
            }
            _ptabs = st.tabs([_sheet_labels.get(k, k) for k in _all_parsed.keys()])
            for _ptab, (_stype, _rows) in zip(_ptabs, _all_parsed.items()):
                with _ptab:
                    _df = pd.DataFrame(_rows).drop(columns=["_raw_container", "_src"], errors="ignore")
                    st.dataframe(_df, use_container_width=True, hide_index=True)

            st.divider()

            if st.button("✅ Confirm & Submit", type="primary", use_container_width=True, key="vp_submit_btn"):
                _now     = datetime.now(_PHX).isoformat()
                _today   = file_date.isoformat()
                _wk      = (file_date - timedelta(days=file_date.weekday())).isoformat()
                _conn    = get_db()
                _count   = 0
                _DEL_T   = {"delivery","delivery_ilm1","delivery_ric6","delivery_dbm6"}
                _FC_MAP  = {"delivery_ilm1":"ILM1","delivery_ric6":"RIC6","delivery_dbm6":"DBM6"}
                scac_code   = _portal_scac
                submitter   = _portal_display

                try:
                    for _stype, _rows in _all_parsed.items():
                        for _row in _rows:
                            # carrier_submissions log
                            _conn.execute(
                                "INSERT INTO carrier_submissions"
                                " (submitted_at, carrier_name, container_id, sheet_type,"
                                "  port, terminal, fc_building, flexi_id, outgate_date,"
                                "  delivery_date, status, within_sla, sla_notes,"
                                "  empty_return_due, appointment_date, accessorial_type,"
                                "  notes, source_file, source)"
                                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                (_today, scac_code, _row["container_id"], _stype,
                                 _row.get("port"), _row.get("terminal"), _row.get("fc_building"),
                                 _row.get("flexi_id"), _row.get("outgate_date"),
                                 _row.get("delivery_date"), _row.get("status"),
                                 _row.get("within_sla"), _row.get("sla_notes"),
                                 _row.get("empty_return_due"), _row.get("appointment_date"),
                                 _row.get("accessorial_type"), _row.get("notes"),
                                 _row.get("_src"), "vendor_portal"),
                            )
                            _count += 1

                            # inbound_containers upsert (feeds DBR Dashboard)
                            if _stype in _DEL_T:
                                _hd = bool(_row.get("delivery_date"))
                                _ho = bool(_row.get("outgate_date"))
                                if _hd:   _ic_st, _fa = "delivered", str(_row["delivery_date"])
                                elif _ho: _ic_st, _fa = "at_yard", None
                                else:     _ic_st, _fa = "pending", None
                                _fd = _FC_MAP.get(_stype) or _row.get("fc_building")
                                _conn.execute(
                                    "INSERT INTO inbound_containers"
                                    " (container_id,carrier,port_region,fc_dest,"
                                    "  yard_actual,fc_actual,status,notes,last_updated,source_file)"
                                    " VALUES (?,?,?,?,?,?,?,?,?,?)"
                                    " ON CONFLICT(container_id) DO UPDATE SET"
                                    "  carrier=COALESCE(excluded.carrier,carrier),"
                                    "  port_region=COALESCE(excluded.port_region,port_region),"
                                    "  fc_dest=COALESCE(excluded.fc_dest,fc_dest),"
                                    "  yard_actual=COALESCE(excluded.yard_actual,yard_actual),"
                                    "  fc_actual=COALESCE(excluded.fc_actual,fc_actual),"
                                    "  status=excluded.status,"
                                    "  last_updated=excluded.last_updated,"
                                    "  source_file=excluded.source_file",
                                    (_row["container_id"], scac_code, _row.get("port"), _fd,
                                     _row.get("outgate_date"), _fa,
                                     _ic_st, _row.get("notes"), _now, _row.get("_src")),
                                )
                            elif _stype == "empty_return":
                                _ret = str(_row.get("status") or "").lower()
                                if any(kw in _ret for kw in ("return","complete","done","at port")):
                                    _rd = _row.get("appointment_date") or _today
                                    _conn.execute(
                                        "UPDATE inbound_containers SET"
                                        " empty_returned=COALESCE(empty_returned,?),"
                                        " status='empty_returned', last_updated=?"
                                        " WHERE container_id=?",
                                        (_rd, _now, _row["container_id"]),
                                    )

                    # dbr_receipts daily compliance log
                    _conn.execute(
                        "INSERT OR IGNORE INTO dbr_receipts"
                        " (carrier,week_start,received_date,received_via,file_name,logged_at)"
                        " VALUES (?,?,?,?,?,?)",
                        (scac_code, _wk, _today, "vendor_portal",
                         "; ".join(cf.name for cf in carrier_files), _now),
                    )
                    _conn.commit()

                except Exception as _e:
                    _conn.rollback()
                    st.error("Submission failed — please try again or contact AGL. (" + str(_e) + ")")
                    st.stop()
                finally:
                    _conn.close()

                if _count > 0:
                    if S3_ENABLED:
                        data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                    st.session_state["vp_submit_result"] = {
                        "count":     _count,
                        "scac":      scac_code,
                        "submitter": submitter,
                        "timestamp": datetime.now(_PHX).strftime("%B %d, %Y %H:%M MST"),
                        "files":     [cf.name for cf in carrier_files],
                    }
                    st.rerun()

st.divider()
st.caption("AGL Robotics Logistics · Questions? Contact your AGL point of contact.")
