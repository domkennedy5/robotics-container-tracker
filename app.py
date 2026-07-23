from __future__ import annotations
import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import io
import os
from datetime import datetime, date

from utils import (
    normalize_container, containers_match,
    parse_container_list, parse_carrier_file,
    parse_carrier_template, TEMPLATE_SHEET_MAP,
)
import data_sync

# ── page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Robotics Container Tracker",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "tracker.db")

# ── AWS secrets (graceful fallback if not configured) ─────────────────────────
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

# ── password gate ──────────────────────────────────────────────────────────────
def _check_password():
    try:
        correct = st.secrets["app"]["PASSWORD"]
    except Exception:
        return True  # no password configured — open access (local dev)

    if st.session_state.get("authenticated"):
        return True

    st.title("🚢 Robotics Container Tracker")
    pwd = st.text_input("Password", type="password")
    if st.button("Sign in"):
        if pwd == correct:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False

if not _check_password():
    st.stop()

# ── S3 startup sync ────────────────────────────────────────────────────────────
if S3_ENABLED and not st.session_state.get("s3_startup_done"):
    data_sync.pull_db_from_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
    st.session_state.s3_startup_done = True

# ── database ───────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS carrier_submissions (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            submitted_at     TEXT    NOT NULL,
            carrier_name     TEXT    NOT NULL,
            container_id     TEXT    NOT NULL,
            sheet_type       TEXT,
            port             TEXT,
            terminal         TEXT,
            fc_building      TEXT,
            flexi_id         TEXT,
            outgate_date     TEXT,
            delivery_date    TEXT,
            status           TEXT,
            within_sla       TEXT,
            sla_notes        TEXT,
            empty_return_due TEXT,
            appointment_date TEXT,
            accessorial_type TEXT,
            notes            TEXT,
            source_file      TEXT,
            source           TEXT DEFAULT 'web'
        );
        CREATE TABLE IF NOT EXISTS lookup_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            searched_at     TEXT    NOT NULL,
            requester       TEXT,
            container_ids   TEXT    NOT NULL,
            found_count     INTEGER,
            not_found_count INTEGER
        );
        CREATE TABLE IF NOT EXISTS alerts_log (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            alert_type   TEXT NOT NULL,
            container_id TEXT NOT NULL,
            sent_at      TEXT NOT NULL,
            recipient    TEXT
        );
    """)
    conn.commit()
    conn.close()

init_db()

def migrate_db():
    """Add columns missing in old-schema DBs pulled from S3."""
    new_cols = [
        ("sheet_type",       "TEXT"),
        ("port",             "TEXT"),
        ("fc_building",      "TEXT"),
        ("flexi_id",         "TEXT"),
        ("outgate_date",     "TEXT"),
        ("delivery_date",    "TEXT"),
        ("within_sla",       "TEXT"),
        ("sla_notes",        "TEXT"),
        ("empty_return_due", "TEXT"),
        ("appointment_date", "TEXT"),
        ("accessorial_type", "TEXT"),
    ]
    conn = get_db()
    for col, col_type in new_cols:
        try:
            conn.execute(f"ALTER TABLE carrier_submissions ADD COLUMN {col} {col_type}")
        except sqlite3.OperationalError:
            pass  # column already exists
    conn.commit()
    conn.close()

migrate_db()

def db_write(sql: str, params: tuple):
    """Write to DB then push to S3."""
    conn = get_db()
    conn.execute(sql, params)
    conn.commit()
    conn.close()
    if S3_ENABLED:
        data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)

# ── DBR config & loader ────────────────────────────────────────────────────────
SHEET_CFG = {
    "Delivery Appointments": {
        "container_col": "Container #",
        "display_cols": ["Location", "Container #", "Status", "Terminal ",
                         "FC/Building", "Del Appointment Date/Time", "LFD", "Outgate Date"],
        "status_label": "Delivery Appointments",
    },
    "Empty Returns": {
        "container_col": "Container #",
        "display_cols": ["Container #", "Terminal", "Empty Return Due Date",
                         "Appointment Date", "Status", "If Outside Window - Reason"],
        "status_label": "Empty Returns",
    },
    "On Vessel": {
        "container_col": "Container #",
        "display_cols": ["Location", "Container #", "Status", "Terminal ",
                         "LFD", "Appointment Date/Time", "Flexi ID"],
        "status_label": "On Vessel",
    },
    "CANCELED": {
        "container_col": "Container #",
        "display_cols": ["Location", "Container #", "Status", "Terminal ",
                         "FC/Building", "Del Appointment Date/Time", "LFD"],
        "status_label": "Canceled",
    },
    "Demurrage": {
        "container_col": "Container #",
        "display_cols": ["Container #", "Terminal ", "Type of Hold", "Days in Hold", "Bridge"],
        "status_label": "Demurrage",
    },
    "Accessorials": {
        "container_col": "Container #",
        "display_cols": ["Container #", "Accessorial", "Date(s) Incurred",
                         "Reason for Charge", "Notes"],
        "status_label": "Accessorials",
    },
}

@st.cache_data(show_spinner="Parsing DBR…", ttl=3600)
def load_dbr(file_bytes: bytes) -> dict:
    # pd.read_excel handles merged cells, varying row widths, and None headers
    sheets = {}
    for sheet_name, cfg in SHEET_CFG.items():
        try:
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                sheet_name=sheet_name,
                header=0,
                dtype=str,
                engine="openpyxl",
            )
        except Exception:
            continue
        df.columns = [str(c).strip() if pd.notna(c) else f"_col_{i}"
                      for i, c in enumerate(df.columns)]
        cont_col = cfg["container_col"]
        if cont_col in df.columns:
            df = df[df[cont_col].notna() & (df[cont_col].str.strip() != "")].copy()
            df["_norm"] = df[cont_col].apply(lambda x: normalize_container(str(x)))
        sheets[sheet_name] = df
    return sheets

def lookup_containers(dbr_sheets: dict, query_ids: list) -> tuple:
    norm_map = {normalize_container(cid): cid for cid in query_ids}
    results, matched = [], set()

    for sheet_name, cfg in SHEET_CFG.items():
        if sheet_name not in dbr_sheets:
            continue
        df = dbr_sheets[sheet_name].copy()
        if "_norm" not in df.columns:
            continue

        def find_match(dbr_norm):
            for qn in norm_map:
                if containers_match(qn, dbr_norm):
                    return qn
            return None

        df["_mq"] = df["_norm"].apply(find_match)
        hits = df[df["_mq"].notna()].copy()
        if hits.empty:
            continue

        keep = [c for c in cfg["display_cols"] if c in hits.columns]
        hits = hits[keep + ["_mq"]].copy()
        hits.rename(columns={cfg["container_col"]: "Container #"}, inplace=True, errors="ignore")
        hits.insert(0, "Sheet", cfg["status_label"])
        hits.insert(1, "Searched As", hits["_mq"].map(lambda n: norm_map.get(n, n)))
        hits.drop(columns=["_mq"], inplace=True)
        hits.columns = [str(c).strip() for c in hits.columns]
        results.append(hits)
        for qn in df[df["_mq"].notna()]["_mq"].tolist():
            matched.add(qn)

    not_found = [orig for norm, orig in norm_map.items() if norm not in matched]
    combined = pd.concat(results, ignore_index=True) if results else pd.DataFrame()
    return combined, not_found

# ── sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 DBR File")

    # try to auto-load from S3 on first run
    if S3_ENABLED and not st.session_state.get("dbr_auto_loaded"):
        with st.spinner("Checking S3 for latest DBR…"):
            dbr_bytes_s3 = data_sync.pull_dbr_from_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
            if dbr_bytes_s3:
                st.session_state.dbr_bytes = dbr_bytes_s3
                last_updated = data_sync.get_dbr_last_updated(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                st.session_state.dbr_last_updated = last_updated
        st.session_state.dbr_auto_loaded = True

    dbr_file = st.file_uploader(
        "Upload new DBR Excel", type=["xlsx"],
        help="Upload a new weekly DBR. It will be saved to S3 automatically."
    )

    if dbr_file:
        raw = dbr_file.read()
        st.session_state.dbr_bytes = raw
        st.session_state.dbr_last_updated = datetime.utcnow().isoformat()
        if S3_ENABLED:
            with st.spinner("Saving to S3…"):
                ok = data_sync.push_dbr_to_s3(raw, AWS_KEY, AWS_SECRET,
                                               AWS_REGION, S3_BUCKET, dbr_file.name)
            st.success("Saved to S3 ✓" if ok else "Saved locally (S3 unavailable)")
        else:
            st.success(f"Loaded: {dbr_file.name}")

    if st.session_state.get("dbr_bytes"):
        dbr_sheets = load_dbr(st.session_state.dbr_bytes)
        ts = st.session_state.get("dbr_last_updated", "")
        if ts:
            try:
                friendly = datetime.fromisoformat(ts).strftime("%b %d, %Y %H:%M UTC")
            except Exception:
                friendly = ts
            st.caption(f"DBR last updated: {friendly}")
        st.caption(f"Delivery Appts: {len(dbr_sheets.get('Delivery Appointments', pd.DataFrame()))} rows")
        st.caption(f"Empty Returns: {len(dbr_sheets.get('Empty Returns', pd.DataFrame()))} rows")
    else:
        dbr_sheets = None
        st.info("No DBR loaded yet." if S3_ENABLED else "Upload the DBR to enable lookups.")

    st.divider()
    st.caption("🚢 **Container Tracker v1.1**")
    if S3_ENABLED:
        st.caption("☁️ S3 sync active")

# ── tabs ───────────────────────────────────────────────────────────────────────
st.title("🚢 Robotics Container Tracker")
tab1, tab2, tab3, tab4 = st.tabs(["🔍 Container Lookup", "📤 Carrier Submission", "📋 Empty Returns", "📊 Carrier Data"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Container Lookup
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Container Lookup")
    st.caption("Paste any list of IDs — dashes, no dashes, with or without check digit.")

    col_a, col_b = st.columns([2, 1])
    with col_a:
        lookup_input = st.text_area(
            "Container IDs", height=260, label_visibility="collapsed",
            placeholder="TCNU389902-4\nONEU624470-0\nGCXU542076\n...",
        )
    with col_b:
        requester = st.text_input("Your name (optional)", placeholder="e.g. Dominique")
        st.caption("Searches: Delivery Appts · Empty Returns · On Vessel · Canceled · Demurrage · Accessorials")
        search_btn = st.button("🔍 Search", type="primary", use_container_width=True)

    if search_btn:
        if not dbr_sheets:
            st.error("No DBR loaded — upload one in the sidebar.")
        elif not lookup_input.strip():
            st.warning("Paste at least one container ID.")
        else:
            query_ids = parse_container_list(lookup_input)
            results_df, not_found = lookup_containers(dbr_sheets, query_ids)

            db_write(
                "INSERT INTO lookup_log (searched_at, requester, container_ids, found_count, not_found_count) VALUES (?,?,?,?,?)",
                (datetime.now().isoformat(), requester or "anonymous",
                 "\n".join(query_ids), len(query_ids) - len(not_found), len(not_found))
            )

            m1, m2, m3 = st.columns(3)
            m1.metric("Queried", len(query_ids))
            m2.metric("Found", len(query_ids) - len(not_found))
            m3.metric("Not found", len(not_found))
            st.divider()

            if not results_df.empty:
                st.dataframe(results_df, use_container_width=True, hide_index=True)
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    results_df.to_excel(w, index=False, sheet_name="Found")
                    if not_found:
                        pd.DataFrame({"Not Found": not_found}).to_excel(
                            w, index=False, sheet_name="Not Found")
                buf.seek(0)
                st.download_button("⬇️ Download results (.xlsx)", data=buf,
                    file_name=f"container_lookup_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("None of the queried containers found in the DBR.")

            if not_found:
                st.markdown("### ❌ Not Found in DBR")
                st.dataframe(pd.DataFrame({"Container ID": not_found}),
                             use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Carrier Submission
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Carrier Submission Portal")

    col_info, col_tmpl = st.columns([3, 1])
    with col_info:
        st.info(
            "**Two ways to submit:**  \n"
            "1. Upload your filled-out AGL template below  \n"
            "2. Email your template to the inbound address — it will be processed automatically"
        )
    with col_tmpl:
        tmpl_path = os.path.join(BASE_DIR, "agl_carrier_template.xlsx")
        if os.path.exists(tmpl_path):
            with open(tmpl_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Carrier Template",
                    data=f.read(),
                    file_name="AGL_Carrier_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Fill this out and upload below or email it in"
                )

    st.divider()

    # ── Template upload + preview ──────────────────────────────────────────────
    c1, c2 = st.columns([2, 1])
    with c1:
        carrier_name_input = st.text_input("Carrier / Company Name *", placeholder="e.g. ARVY, XPO…", key="cn_top")
    with c2:
        carrier_file_top = st.file_uploader(
            "Upload completed template", type=["xlsx", "csv"],
            help="Upload your filled AGL carrier template",
            key="cf_top"
        )

    if carrier_file_top and carrier_name_input.strip():
        # cache bytes in session_state — file object resets across reruns in Streamlit
        cache_key = f"cf_bytes_{carrier_file_top.name}_{carrier_file_top.size}"
        if cache_key not in st.session_state:
            st.session_state[cache_key] = carrier_file_top.read()
        raw_bytes = st.session_state[cache_key]
        parsed = parse_carrier_template(raw_bytes, carrier_file_top.name)

        if not parsed:
            st.warning("No container data found in this file. Make sure you're using the AGL Carrier Template.")
        else:
            # count totals
            total = sum(len(v) for v in parsed.values())
            st.success(f"Found **{total} containers** across **{len(parsed)} sheet(s)** — review below then click Submit.")

            # show preview tabs per sheet type
            sheet_labels = {
                "delivery":       "📦 Delivery",
                "delivery_ilm1":  "🏭 ILM1",
                "delivery_ric6":  "🏭 RIC6",
                "empty_return":   "🔁 Empty Returns",
                "ody":            "🏗️ Storage/ODY",
                "demurrage":      "⚠️ Demurrage",
                "accessorial":    "💰 Accessorials",
            }
            preview_tabs = st.tabs([sheet_labels.get(k, k) for k in parsed.keys()])
            for ptab, (sheet_type, rows) in zip(preview_tabs, parsed.items()):
                with ptab:
                    df_prev = pd.DataFrame(rows).drop(columns=["_raw_container"], errors="ignore")
                    st.dataframe(df_prev, use_container_width=True, hide_index=True)

            if st.button("✅ Confirm & Submit All", type="primary"):
                now = datetime.now().isoformat()
                conn = get_db()
                count = 0
                for sheet_type, rows in parsed.items():
                    for row in rows:
                        conn.execute(
                            """INSERT INTO carrier_submissions
                               (submitted_at, carrier_name, container_id, sheet_type,
                                port, terminal, fc_building, flexi_id, outgate_date,
                                delivery_date, status, within_sla, sla_notes,
                                empty_return_due, appointment_date, accessorial_type,
                                notes, source_file, source)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (now, carrier_name_input.strip(), row["container_id"],
                             sheet_type,
                             row.get("port"), row.get("terminal"), row.get("fc_building"),
                             row.get("flexi_id"), row.get("outgate_date"),
                             row.get("delivery_date"), row.get("status"),
                             row.get("within_sla"), row.get("sla_notes"),
                             row.get("empty_return_due"), row.get("appointment_date"),
                             row.get("accessorial_type"), row.get("notes"),
                             carrier_file_top.name, "web")
                        )
                        count += 1
                conn.commit()
                conn.close()
                if S3_ENABLED:
                    data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                st.success(f"✅ Logged {count} containers from **{carrier_name_input}**")
                st.rerun()

    elif carrier_file_top and not carrier_name_input.strip():
        st.warning("Enter your carrier name above before uploading.")

    if S3_ENABLED:
        try:
            inbound_email = st.secrets["notifications"].get("INBOUND_EMAIL", "")
            if inbound_email:
                st.code(inbound_email, language=None)
        except Exception:
            pass

    with st.form("carrier_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            carrier_name = st.text_input("Carrier / Company Name *", placeholder="e.g. ARVY, XPO…")
            terminal     = st.text_input("Terminal", placeholder="e.g. EWR, BOS, CHI")
        with c2:
            submitted_status = st.selectbox("Status", [
                "Delivery Scheduled",
                "Delivered pending empty return",
                "Empty Return Scheduled",
                "TERMINATED",
                "On Vessel",
                "Customs Hold",
                "Other",
            ])
            notes = st.text_input("Notes", placeholder="Delays, exceptions…")

        container_text = st.text_area(
            "Container IDs (one per line) *", height=180,
            placeholder="TCNU389902-4\nONEU624470-0\n…",
        )
        carrier_file = st.file_uploader(
            "Or upload your status file", type=["xlsx", "csv"],
            help="ARVY-format Excel or CSV — containers extracted automatically"
        )
        submitted = st.form_submit_button("📤 Submit", type="primary")

    if submitted:
        if not carrier_name.strip():
            st.error("Carrier name is required.")
        else:
            ids = []
            if container_text.strip():
                ids.extend(parse_container_list(container_text))
            source_file = None
            if carrier_file:
                source_file = carrier_file.name
                fdf = parse_carrier_file(carrier_file.read(), carrier_file.name)
                if not fdf.empty:
                    ids.extend(fdf["container_id"].tolist())

            ids = list(dict.fromkeys(ids))
            if not ids:
                st.error("No container IDs found.")
            else:
                now = datetime.now().isoformat()
                conn = get_db()
                for cid in ids:
                    conn.execute(
                        """INSERT INTO carrier_submissions
                           (submitted_at, carrier_name, container_id, terminal,
                            status, notes, source_file, source)
                           VALUES (?,?,?,?,?,?,?,?)""",
                        (now, carrier_name.strip(), normalize_container(cid),
                         terminal.strip() or None, submitted_status,
                         notes.strip() or None, source_file, "web")
                    )
                conn.commit()
                conn.close()
                if S3_ENABLED:
                    data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                st.success(f"✅ Logged {len(ids)} containers from **{carrier_name}**")

    st.divider()
    st.markdown("### Submission Log")

    conn = get_db()
    log_df = pd.read_sql(
        """SELECT submitted_at, carrier_name, container_id, terminal,
                  status, notes, source
           FROM carrier_submissions ORDER BY submitted_at DESC LIMIT 500""",
        conn
    )
    conn.close()

    if log_df.empty:
        st.info("No submissions yet.")
    else:
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            f_carrier = st.multiselect("Carrier", sorted(log_df["carrier_name"].unique()))
        with fc2:
            f_status = st.multiselect("Status", sorted(log_df["status"].dropna().unique()))
        with fc3:
            f_source = st.multiselect("Source", sorted(log_df["source"].dropna().unique()))

        filt = log_df.copy()
        if f_carrier: filt = filt[filt["carrier_name"].isin(f_carrier)]
        if f_status:  filt = filt[filt["status"].isin(f_status)]
        if f_source:  filt = filt[filt["source"].isin(f_source)]

        st.dataframe(filt, use_container_width=True, hide_index=True)

        buf = io.BytesIO()
        filt.to_excel(buf, index=False)
        buf.seek(0)
        st.download_button("⬇️ Export log (.xlsx)", data=buf,
            file_name=f"carrier_log_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Empty Returns Dashboard
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Empty Returns Dashboard")

    if not dbr_sheets:
        st.info("Upload the DBR file to see empty returns data.")
    else:
        er_raw = dbr_sheets.get("Empty Returns", pd.DataFrame()).copy()
        if er_raw.empty:
            st.warning("No Empty Returns sheet found in this DBR.")
        else:
            er_raw.columns = [str(c).strip() for c in er_raw.columns]
            er_raw["Empty Return Due Date"] = pd.to_datetime(
                er_raw["Empty Return Due Date"], errors="coerce")
            today = pd.Timestamp(date.today())

            active = er_raw[
                er_raw["Status"].fillna("").str.upper().ne("TERMINATED") &
                er_raw["Empty Return Due Date"].notna()
            ].copy()
            active["Days Until Due"] = (active["Empty Return Due Date"] - today).dt.days
            active["⚠️"] = active["Days Until Due"].apply(
                lambda d: "🔴 OVERDUE" if d < 0 else ("🟡 Due soon" if d <= 3 else "🟢 OK")
            )

            overdue   = active[active["Days Until Due"] < 0]
            due_soon  = active[(active["Days Until Due"] >= 0) & (active["Days Until Due"] <= 3)]
            on_track  = active[active["Days Until Due"] > 3]

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Active", len(active))
            m2.metric("🔴 Overdue", len(overdue))
            m3.metric("🟡 Due ≤ 3 days", len(due_soon))
            m4.metric("🟢 On track", len(on_track))
            st.divider()

            view = st.radio("Show", ["🔴 Overdue", "🟡 Due Soon (≤3 days)",
                                      "All Active", "All incl. Terminated"],
                            horizontal=True)
            display = (overdue if view == "🔴 Overdue"
                       else due_soon if view == "🟡 Due Soon (≤3 days)"
                       else active   if view == "All Active"
                       else er_raw)

            show_cols = [c for c in ["⚠️", "Container #", "Terminal",
                                      "Empty Return Due Date", "Appointment Date",
                                      "Status", "Days Until Due",
                                      "If Outside Window - Reason"]
                         if c in display.columns]

            st.dataframe(
                display[show_cols].sort_values("Days Until Due", ascending=True)
                if "Days Until Due" in display.columns else display[show_cols],
                use_container_width=True, hide_index=True
            )

            if not display.empty:
                buf = io.BytesIO()
                display.to_excel(buf, index=False)
                buf.seek(0)
                st.download_button("⬇️ Export (.xlsx)", data=buf,
                    file_name=f"empty_returns_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.divider()
            st.markdown("### Terminated")
            term = er_raw[er_raw["Status"].fillna("").str.upper() == "TERMINATED"] if "Status" in er_raw.columns else pd.DataFrame()
            if not term.empty:
                st.dataframe(term, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — Carrier Data (structured view of all submissions)
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Carrier Submission Data")
    st.caption("All data submitted by carriers via template upload or email.")

    conn = get_db()
    all_df = pd.read_sql(
        """SELECT submitted_at, carrier_name, container_id, sheet_type,
                  port, terminal, fc_building, delivery_date, status,
                  within_sla, empty_return_due, appointment_date,
                  accessorial_type, notes, source
           FROM carrier_submissions ORDER BY submitted_at DESC LIMIT 1000""",
        conn
    )
    conn.close()

    if all_df.empty:
        st.info("No carrier submissions yet.")
    else:
        # top filters
        f1, f2, f3 = st.columns(3)
        with f1:
            f_carrier = st.multiselect("Carrier", sorted(all_df["carrier_name"].unique()), key="cd_carrier")
        with f2:
            f_sheet = st.multiselect("Sheet type", sorted(all_df["sheet_type"].dropna().unique()), key="cd_sheet")
        with f3:
            f_sla = st.multiselect("Within SLA?", sorted(all_df["within_sla"].dropna().unique()), key="cd_sla")

        filt = all_df.copy()
        if f_carrier: filt = filt[filt["carrier_name"].isin(f_carrier)]
        if f_sheet:   filt = filt[filt["sheet_type"].isin(f_sheet)]
        if f_sla:     filt = filt[filt["within_sla"].isin(f_sla)]

        # metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total containers", len(filt))
        m2.metric("Carriers", filt["carrier_name"].nunique())
        sla_counts = filt["within_sla"].value_counts()
        m3.metric("Within SLA", int(sla_counts.get("YES", 0)))
        m4.metric("Outside SLA", int(sla_counts.get("NO", 0)))

        st.divider()

        # tabbed by sheet type
        sheet_types = sorted(filt["sheet_type"].dropna().unique())
        if sheet_types:
            sheet_labels = {
                "delivery":       "📦 Delivery",
                "delivery_ilm1":  "🏭 ILM1",
                "delivery_ric6":  "🏭 RIC6",
                "empty_return":   "🔁 Empty Returns",
                "ody":            "🏗️ ODY/Storage",
                "demurrage":      "⚠️ Demurrage",
                "accessorial":    "💰 Accessorials",
            }
            stabs = st.tabs([sheet_labels.get(s, s) for s in sheet_types])
            for stab, stype in zip(stabs, sheet_types):
                with stab:
                    sdf = filt[filt["sheet_type"] == stype].copy()
                    # show only non-empty columns
                    nonempty = [c for c in sdf.columns if sdf[c].notna().any() and sdf[c].astype(str).str.strip().ne("").any()]
                    st.dataframe(sdf[nonempty], use_container_width=True, hide_index=True)

        st.divider()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            for stype in filt["sheet_type"].dropna().unique():
                sdf = filt[filt["sheet_type"] == stype]
                sdf.to_excel(w, index=False, sheet_name=stype[:31])
        buf.seek(0)
        st.download_button("⬇️ Export all carrier data (.xlsx)", data=buf,
            file_name=f"carrier_data_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
