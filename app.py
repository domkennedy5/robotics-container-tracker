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
VENDOR_PORTAL_URL = "https://robotics-container-tracker-7uf88f7ez9tga3k44phfjm.streamlit.app/vendor_upload"

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
        CREATE TABLE IF NOT EXISTS carrier_rates (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            carrier_name        TEXT NOT NULL UNIQUE,
            rate_per_container  REAL DEFAULT 0.0,
            notes               TEXT
        );
        CREATE TABLE IF NOT EXISTS rate_lanes (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            scac         TEXT NOT NULL,
            carrier_name TEXT,
            port         TEXT NOT NULL,
            destination  TEXT NOT NULL,
            dest_address TEXT,
            base_rate    REAL NOT NULL,
            lane_type    TEXT,
            lane_id      TEXT,
            rate_source  TEXT,
            imported_at  TEXT
        );
        CREATE TABLE IF NOT EXISTS route_lanes (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            origin_port       TEXT,
            dest_port         TEXT,
            transit_notes     TEXT,
            node_code         TEXT NOT NULL,
            node_address      TEXT,
            fc_code           TEXT,
            node_type         TEXT,
            program_id        TEXT,
            warehouse_partner TEXT,
            dray_bid_done     TEXT,
            cost_cards_done   TEXT,
            imported_at       TEXT
        );
        CREATE TABLE IF NOT EXISTS shipment_status (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            container_no        TEXT,
            current_status      TEXT,
            fc                  TEXT,
            vendor_name         TEXT,
            po_no               TEXT,
            edd                 TEXT,
            vessel_name         TEXT,
            voyage_no           TEXT,
            carrier_code        TEXT,
            discharge_city      TEXT,
            discharge_eta       TEXT,
            final_dest_eta      TEXT,
            container_available TEXT,
            container_out       TEXT,
            empty_return        TEXT,
            delivered           TEXT,
            customs_cleared     TEXT,
            container_size      TEXT,
            load_port_country   TEXT,
            load_port_city      TEXT,
            source_file         TEXT,
            imported_at         TEXT
        );
        CREATE TABLE IF NOT EXISTS inbound_loads (
            id                       INTEGER PRIMARY KEY AUTOINCREMENT,
            container_id             TEXT,
            supplier                 TEXT,
            po_number                TEXT,
            location_code            TEXT,
            destination_fc           TEXT,
            origin_port              TEXT,
            dest_port                TEXT,
            vessel_name              TEXT,
            eta_discharge            TEXT,
            actual_arrival_discharge TEXT,
            gateout_terminal         TEXT,
            last_free_detention      TEXT,
            last_free_demurrage      TEXT,
            est_appointment          TEXT,
            eta_final                TEXT,
            actual_arrival_final     TEXT,
            empty_return_time        TEXT,
            source_file              TEXT,
            imported_at              TEXT
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


# ── Report parsers (Import Shipment Status + Inbound Loads) ───────────────────
_SS_COLS = {
    "ContainerNo": "container_no", "CurrentStatus": "current_status",
    "FC": "fc", "VendorName": "vendor_name", "PONo": "po_no",
    "EDD": "edd", "VesselName": "vessel_name", "VoyageNo": "voyage_no",
    "CarrierCode": "carrier_code", "DischargeCity": "discharge_city",
    "DischargePortETA": "discharge_eta", "FinalDestETA": "final_dest_eta",
    "EDIContainerAvailable": "container_available", "EDIContainerOut": "container_out",
    "EDIEmptyReturn": "empty_return", "EDIDelivered": "delivered",
    "EDIClearedCustoms": "customs_cleared", "ContainerSize": "container_size",
    "LoadPortCountry": "load_port_country", "LoadPortCity": "load_port_city",
}
_IL_COLS = {
    "Container Id": "container_id", "Supplier": "supplier",
    "PO Number": "po_number", "Location Code": "location_code",
    "Destination FC": "destination_fc", "International Origin": "origin_port",
    "International Destination": "dest_port", "Vessel Name": "vessel_name",
    "ETA Arrival At Discharge Port": "eta_discharge",
    "Actual Arrival At Discharge Port": "actual_arrival_discharge",
    "Gateout Terminal Time": "gateout_terminal",
    "Last Free Detention Time": "last_free_detention",
    "Last Free Demurrage Time": "last_free_demurrage",
    "Estimated Appointment Date": "est_appointment",
    "ETA To Final Destination": "eta_final",
    "Actual Arrival At Final Destination": "actual_arrival_final",
    "Empty Container Return Time": "empty_return_time",
}

def parse_shipment_status_file(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, engine="openpyxl")
    available = {k: v for k, v in _SS_COLS.items() if k in df.columns}
    out = df[list(available.keys())].rename(columns=available).copy()
    out = out[out["container_no"].notna() & (out["container_no"].str.strip() != "")]
    out = out.drop_duplicates(subset=["container_no"], keep="last")
    return out.reset_index(drop=True)

def parse_inbound_loads_file(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, engine="openpyxl")
    available = {k: v for k, v in _IL_COLS.items() if k in df.columns}
    out = df[list(available.keys())].rename(columns=available).copy()
    out = out[out["container_id"].notna() & (out["container_id"].str.strip() != "")]
    out = out.drop_duplicates(subset=["container_id"], keep="last")
    return out.reset_index(drop=True)

def upsert_report(df: pd.DataFrame, table: str, key_col: str, source_file: str):
    """Delete existing rows for containers in this upload, then insert fresh."""
    now = datetime.now().isoformat()
    df = df.copy()
    df["source_file"] = source_file
    df["imported_at"]  = now
    conn = get_db()
    # remove stale rows for these container IDs
    ids = df[key_col].dropna().unique().tolist()
    if ids:
        placeholders = ",".join("?" * len(ids))
        conn.execute(f"DELETE FROM {table} WHERE {key_col} IN ({placeholders})", ids)
    df.to_sql(table, conn, if_exists="append", index=False)
    conn.commit()
    conn.close()
    if S3_ENABLED:
        data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
    return len(df)


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
    st.markdown("**📋 Additional Reports**")
    st.caption("Upload emailed reports to enrich tracking data.")

    # Import Shipment Status
    ss_file = st.file_uploader(
        "Import Shipment Status", type=["xlsx"],
        help="Weekly shipment status report — adds EDI dates, vessel, customs status per container",
        key="ss_upload"
    )
    if ss_file:
        with st.spinner("Parsing…"):
            ss_bytes = ss_file.read()
            ss_df = parse_shipment_status_file(ss_bytes)
            n_ss = upsert_report(ss_df, "shipment_status", "container_no", ss_file.name)
            st.session_state.ss_loaded = True
            st.session_state.ss_source = ss_file.name
        st.success(f"{n_ss:,} containers loaded ✓")

    if st.session_state.get("ss_loaded"):
        st.caption(f"✓ Shipment Status: {st.session_state.get('ss_source','')}")

    # Inbound Loads Report
    il_file = st.file_uploader(
        "Inbound Loads Report", type=["xlsx"],
        help="AR Inbound Loads — adds detention/demurrage free-time deadlines per container",
        key="il_upload"
    )
    if il_file:
        with st.spinner("Parsing…"):
            il_bytes = il_file.read()
            il_df = parse_inbound_loads_file(il_bytes)
            n_il = upsert_report(il_df, "inbound_loads", "container_id", il_file.name)
            st.session_state.il_loaded = True
            st.session_state.il_source = il_file.name
        st.success(f"{n_il:,} containers loaded ✓")

    if st.session_state.get("il_loaded"):
        st.caption(f"✓ Inbound Loads: {st.session_state.get('il_source','')}")

    st.divider()
    st.caption("🚢 **Container Tracker v1.1**")
    if S3_ENABLED:
        st.caption("☁️ S3 sync active")
    st.page_link("pages/help.py", label="📖 User Guide", icon="📖")

# ── tabs ───────────────────────────────────────────────────────────────────────
st.title("🚢 Robotics Container Tracker")
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["🔍 Container Lookup", "📤 Carrier Submission", "📋 Empty Returns", "📊 Carrier Data", "🎛️ Allocation", "📈 Insights"])


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
        st.caption(f"📎 Vendor portal (share with carriers): [{VENDOR_PORTAL_URL}]({VENDOR_PORTAL_URL})")

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


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — Lane Cost Simulator
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("Lane Cost Simulator")
    st.caption("Rate matrix by port + destination, carrier options, and scenario cost projection.")

    # ── Load rate + route data ────────────────────────────────────────────────
    conn = get_db()
    rates_df  = pd.read_sql("SELECT * FROM rate_lanes  ORDER BY port, destination, base_rate", conn)
    routes_df = pd.read_sql(
        "SELECT DISTINCT node_code, fc_code, node_address, dest_port, node_type, warehouse_partner, dray_bid_done "
        "FROM route_lanes ORDER BY node_code", conn
    )
    conn.close()

    if rates_df.empty:
        st.info(
            "No rate data loaded yet. Rate cards are imported by the AGL ops team — "
            "contact the app admin to refresh."
        )
        st.stop()

    # ── Build node lookups ────────────────────────────────────────────────────
    # rate_lanes.destination matches either node_code (A100) or fc_code (ARW0)
    fc_to_node   = routes_df.dropna(subset=["fc_code"]).set_index("fc_code")["node_code"].to_dict()
    node_to_fc   = routes_df.dropna(subset=["fc_code"]).set_index("node_code")["fc_code"].to_dict()
    node_to_addr = routes_df.set_index("node_code")["node_address"].to_dict()
    node_to_type = routes_df.set_index("node_code")["node_type"].to_dict()

    # normalize destination to node_code where possible
    rates_df["node_code"] = rates_df["destination"].apply(
        lambda d: d if str(d) in node_to_fc else fc_to_node.get(str(d), str(d))
    )

    # ── Global filters ────────────────────────────────────────────────────────
    gf1, gf2, gf3 = st.columns(3)
    all_ports   = sorted(rates_df["port"].dropna().unique())
    all_sources = sorted(rates_df["rate_source"].dropna().unique())
    all_types   = sorted(rates_df["lane_type"].dropna().unique())

    with gf1:
        g_port = st.selectbox("Port of Arrival", ["All"] + all_ports, key="g_port")
    with gf2:
        g_sources = st.multiselect("Rate Source", all_sources,
            default=["robotics_2026"] if "robotics_2026" in all_sources else all_sources,
            key="g_sources")
    with gf3:
        g_types = st.multiselect("Lane Type", all_types, default=all_types, key="g_types")

    rfilt = rates_df.copy()
    if g_port != "All":
        rfilt = rfilt[rfilt["port"] == g_port]
    if g_sources:
        rfilt = rfilt[rfilt["rate_source"].isin(g_sources)]
    if g_types:
        rfilt = rfilt[rfilt["lane_type"].isin(g_types)]

    if rfilt.empty:
        st.warning("No rates for the selected filters.")
        st.stop()

    # ── Rate Matrix ────────────────────────────────────────────────────────────
    with st.expander("Rate Matrix — all carrier options by lane", expanded=True):
        st.caption("Green = cheapest carrier per lane  ·  Red = most expensive  ·  — = no rate on this lane")

        pivot = rfilt.pivot_table(
            index=["port", "node_code"],
            columns="carrier_name",
            values="base_rate",
            aggfunc="min",
        ).reset_index()
        pivot.columns.name = None
        carrier_cols = [c for c in pivot.columns if c not in ["port", "node_code"]]

        pivot["FC"] = pivot["node_code"].map(node_to_fc)
        pivot["Type"] = pivot["node_code"].map(node_to_type)
        pivot["Facility"] = pivot["node_code"].map(
            lambda n: (node_to_addr.get(n) or "")[:55]
        )

        if carrier_cols:
            pivot["★ Cheapest"] = pivot[carrier_cols].idxmin(axis=1)
            pivot["Low $"]  = pivot[carrier_cols].min(axis=1)
            pivot["Spread"] = pivot[carrier_cols].max(axis=1) - pivot["Low $"]

        col_order = (["port", "node_code", "FC", "Type", "Facility"]
                     + carrier_cols
                     + ["★ Cheapest", "Low $", "Spread"])
        col_order = [c for c in col_order if c in pivot.columns]
        disp = pivot[col_order].rename(columns={"port": "Port", "node_code": "Node"}).sort_values(["Port", "Node"])

        fmt = {c: "${:,.0f}" for c in carrier_cols}
        fmt.update({"Low $": "${:,.0f}", "Spread": "${:,.0f}"})

        try:
            styled = disp.style.format(fmt, na_rep="—")
            if len(carrier_cols) > 1:
                styled = (styled
                    .highlight_min(subset=carrier_cols, axis=1, color="#c8f7c5")
                    .highlight_max(subset=carrier_cols, axis=1, color="#fad7d7"))
            st.dataframe(styled, use_container_width=True, hide_index=True)
        except Exception:
            st.dataframe(disp, use_container_width=True, hide_index=True)

    st.divider()

    # ── Scenario Builder ───────────────────────────────────────────────────────
    st.markdown("### Scenario Builder")
    st.caption("Select port → node → carrier, enter container count, build your cost projection.")

    sb_ports = sorted(rfilt["port"].dropna().unique())
    sb1, sb2, sb3, sb4 = st.columns([1.2, 1.5, 2, 1])

    with sb1:
        sb_port = st.selectbox("Port", sb_ports, key="sb_port")

    sb_nodes = sorted(rfilt[rfilt["port"] == sb_port]["node_code"].dropna().unique())
    with sb2:
        sb_node = st.selectbox(
            "Node", sb_nodes, key="sb_node",
            format_func=lambda n: f"{n}  ({node_to_fc.get(n, '?')})",
        )

    carriers_here = (
        rfilt[(rfilt["port"] == sb_port) & (rfilt["node_code"] == sb_node)]
        [["carrier_name", "base_rate", "lane_type"]]
        .drop_duplicates()
        .sort_values("base_rate")
    )
    c_opts = [
        f"{r.carrier_name}  |  ${r.base_rate:,.0f}  [{r.lane_type}]"
        for _, r in carriers_here.iterrows()
    ]

    with sb3:
        sb_carrier_sel = st.selectbox("Carrier & Rate", c_opts if c_opts else ["No rates on this lane"], key="sb_carrier")
    with sb4:
        sb_count = st.number_input("Containers", min_value=1, value=10, step=1, key="sb_count")

    if st.button("➕ Add Lane to Scenario", type="primary"):
        if c_opts and sb_carrier_sel != "No rates on this lane":
            carrier_nm = sb_carrier_sel.split("  |  $")[0].strip()
            rate_val   = float(sb_carrier_sel.split("$")[1].split("  ")[0].replace(",", ""))
            fc   = node_to_fc.get(sb_node, "")
            addr = (node_to_addr.get(sb_node) or "")[:60]
            if "scenario_v2" not in st.session_state:
                st.session_state.scenario_v2 = []
            st.session_state.scenario_v2.append({
                "Port":       sb_port,
                "Node":       sb_node,
                "FC":         fc,
                "Facility":   addr,
                "Carrier":    carrier_nm,
                "Containers": sb_count,
                "Rate ($)":   rate_val,
                "Cost ($)":   round(sb_count * rate_val, 2),
            })
            st.rerun()

    # ── Portfolio ─────────────────────────────────────────────────────────────
    items = st.session_state.get("scenario_v2", [])

    if not items:
        st.info("Add lanes above to build your cost scenario.")
    else:
        sc_df = pd.DataFrame(items)
        total_ctr  = int(sc_df["Containers"].sum())
        total_cost = float(sc_df["Cost ($)"].sum())
        avg_rate   = total_cost / total_ctr if total_ctr else 0

        pm1, pm2, pm3 = st.columns(3)
        pm1.metric("Total Containers",    f"{total_ctr:,}")
        pm2.metric("Total Projected Cost", f"${total_cost:,.0f}")
        pm3.metric("Avg Cost / Container", f"${avg_rate:,.0f}")

        edited_sc = st.data_editor(
            sc_df,
            column_config={
                "Port":       st.column_config.TextColumn("Port",     disabled=True),
                "Node":       st.column_config.TextColumn("Node",     disabled=True),
                "FC":         st.column_config.TextColumn("FC",       disabled=True),
                "Facility":   st.column_config.TextColumn("Facility", disabled=True),
                "Carrier":    st.column_config.TextColumn("Carrier",  disabled=True),
                "Containers": st.column_config.NumberColumn("Containers", min_value=0),
                "Rate ($)":   st.column_config.NumberColumn("Rate ($)",   format="$%d", disabled=True),
                "Cost ($)":   st.column_config.NumberColumn("Cost ($)",   format="$%d", disabled=True),
            },
            use_container_width=True, hide_index=True, num_rows="dynamic", key="sc_editor",
        )

        # recalc costs if containers edited
        if not edited_sc.equals(sc_df):
            edited_sc["Cost ($)"] = edited_sc["Containers"] * edited_sc["Rate ($)"]
            st.session_state.scenario_v2 = edited_sc.to_dict("records")
            sc_df = edited_sc.copy()
            total_cost = float(sc_df["Cost ($)"].sum())
            total_ctr  = int(sc_df["Containers"].sum())

        dl_c, cl_c = st.columns([3, 1])
        with dl_c:
            buf = io.BytesIO()
            sc_df.to_excel(buf, index=False)
            buf.seek(0)
            st.download_button("⬇️ Export scenario (.xlsx)", buf,
                file_name=f"dray_scenario_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with cl_c:
            if st.button("🗑️ Clear All", type="secondary", use_container_width=True):
                st.session_state.scenario_v2 = []
                st.rerun()

        st.divider()

        # ── By carrier + by port breakdown ────────────────────────────────────
        lc, rc = st.columns(2)

        with lc:
            st.markdown("#### Cost by Carrier")
            by_c = sc_df.groupby("Carrier").agg(
                Containers=("Containers", "sum"),
                Cost=("Cost ($)", "sum"),
            ).reset_index()
            by_c["% Total"] = (by_c["Cost"] / total_cost * 100).round(1) if total_cost else 0
            by_c.columns = ["Carrier", "Containers", "Cost ($)", "% Total"]
            st.dataframe(
                by_c.style.format({"Cost ($)": "${:,.0f}", "% Total": "{:.1f}%"}),
                use_container_width=True, hide_index=True,
            )
            st.bar_chart(by_c.set_index("Carrier")["Cost ($)"])

        with rc:
            st.markdown("#### Cost by Port")
            by_p = sc_df.groupby("Port").agg(
                Containers=("Containers", "sum"),
                Cost=("Cost ($)", "sum"),
            ).reset_index()
            by_p["% Total"] = (by_p["Cost"] / total_cost * 100).round(1) if total_cost else 0
            by_p.columns = ["Port", "Containers", "Cost ($)", "% Total"]
            st.dataframe(
                by_p.style.format({"Cost ($)": "${:,.0f}", "% Total": "{:.1f}%"}),
                use_container_width=True, hide_index=True,
            )
            st.bar_chart(by_p.set_index("Port")["Cost ($)"])

        # ── Savings vs cheapest ───────────────────────────────────────────────
        st.divider()
        st.markdown("#### Savings vs Cheapest Available")

        sav_rows = []
        for _, row in sc_df.iterrows():
            lane_r = rfilt[(rfilt["port"] == row["Port"]) & (rfilt["node_code"] == row["Node"])]
            if lane_r.empty:
                continue
            min_idx      = lane_r["base_rate"].idxmin()
            min_rate     = float(lane_r.loc[min_idx, "base_rate"])
            min_carrier  = str(lane_r.loc[min_idx, "carrier_name"])
            current_cost = float(row["Cost ($)"])
            cheapest     = row["Containers"] * min_rate
            sav_rows.append({
                "Lane":          f"{row['Port']} → {row['Node']} ({row['FC']})",
                "Selected":      row["Carrier"],
                "Rate ($)":      row["Rate ($)"],
                "Cheapest":      min_carrier,
                "Cheapest Rate": min_rate,
                "Containers":    row["Containers"],
                "Current Cost":  current_cost,
                "Min Cost":      cheapest,
                "Savings":       current_cost - cheapest,
            })

        if sav_rows:
            sav_df = pd.DataFrame(sav_rows)
            total_sav = float(sav_df["Savings"].sum())

            if total_sav > 1:
                st.info(f"Switching to cheapest available on all lanes would save **${total_sav:,.0f}**")
            else:
                st.success("You're already using the cheapest available carrier on all lanes.")

            st.dataframe(
                sav_df.style.format({
                    "Rate ($)":      "${:,.0f}",
                    "Cheapest Rate": "${:,.0f}",
                    "Current Cost":  "${:,.0f}",
                    "Min Cost":      "${:,.0f}",
                    "Savings":       "${:,.0f}",
                }),
                use_container_width=True, hide_index=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — Insights
# ══════════════════════════════════════════════════════════════════════════════
with tab6:
    st.subheader("Weekly Insights")
    st.caption("Auto-generated from DBR, carrier submissions, shipment status, and inbound loads data.")

    today_ts = pd.Timestamp(date.today())

    def _risk_flag(d):
        if pd.isna(d): return "—"
        if d < 0:   return f"🔴 {abs(int(d))}d overdue"
        if d == 0:  return "🟡 Due today"
        if d <= 3:  return f"🟡 {int(d)}d"
        return f"🟢 {int(d)}d"

    # ── load supplemental data ────────────────────────────────────────────────
    conn = get_db()
    ss_df = pd.read_sql("SELECT * FROM shipment_status  ORDER BY imported_at DESC", conn)
    il_df = pd.read_sql("SELECT * FROM inbound_loads    ORDER BY imported_at DESC", conn)
    sub_df = pd.read_sql(
        "SELECT submitted_at, carrier_name, container_id, sheet_type, "
        "terminal, fc_building, status, within_sla, source FROM carrier_submissions "
        "WHERE carrier_name IS NOT NULL AND carrier_name != '' ORDER BY submitted_at DESC",
        conn,
    )
    conn.close()

    # normalise date columns
    for col in ["last_free_detention", "last_free_demurrage", "eta_discharge",
                "gateout_terminal", "eta_final", "actual_arrival_final", "empty_return_time"]:
        if col in il_df.columns:
            il_df[col] = pd.to_datetime(il_df[col], errors="coerce")

    for col in ["discharge_eta", "final_dest_eta", "container_available",
                "container_out", "empty_return", "delivered"]:
        if col in ss_df.columns:
            ss_df[col] = pd.to_datetime(ss_df[col], errors="coerce")

    if not sub_df.empty:
        sub_df["submitted_at"] = pd.to_datetime(sub_df["submitted_at"], errors="coerce")

    has_il = not il_df.empty
    has_ss = not ss_df.empty

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 1 — DBR Snapshot
    # ══════════════════════════════════════════════════════════════════════
    st.markdown("### DBR Snapshot")
    if not dbr_sheets:
        st.info("Load a DBR in the sidebar to activate this section.")
    else:
        _del = dbr_sheets.get("Delivery Appointments", pd.DataFrame())
        _er  = dbr_sheets.get("Empty Returns",         pd.DataFrame())
        _vsl = dbr_sheets.get("On Vessel",             pd.DataFrame())
        _dem = dbr_sheets.get("Demurrage",             pd.DataFrame())

        overdue_ct = due_soon_ct = 0
        if not _er.empty and "Empty Return Due Date" in _er.columns:
            er_c = _er.copy()
            er_c["_erd"]  = pd.to_datetime(er_c["Empty Return Due Date"], errors="coerce")
            er_c["_days"] = (er_c["_erd"] - today_ts).dt.days
            er_act = er_c if "Status" not in er_c.columns else \
                     er_c[er_c["Status"].fillna("").str.upper().ne("TERMINATED")]
            overdue_ct  = int((er_act["_days"] < 0).sum())
            due_soon_ct = int(((er_act["_days"] >= 0) & (er_act["_days"] <= 3)).sum())

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Delivery Appts",    len(_del))
        c2.metric("Empty Returns",     len(_er))
        c3.metric("On Vessel",         len(_vsl))
        c4.metric("Demurrage",         len(_dem))
        c5.metric("🔴 ER Overdue",     overdue_ct)
        c6.metric("🟡 ER Due ≤3 days", due_soon_ct)

        # LFD risk
        if not _del.empty and "LFD" in _del.columns:
            lfd = _del.copy()
            lfd["_lfd"]  = pd.to_datetime(lfd["LFD"], errors="coerce")
            lfd["Days to LFD"] = (lfd["_lfd"] - today_ts).dt.days
            risk = lfd[lfd["Days to LFD"].notna() & (lfd["Days to LFD"] <= 5)].copy()
            if not risk.empty:
                risk["⚠️"] = risk["Days to LFD"].apply(_risk_flag)
                show = [c for c in ["⚠️","Container #","Terminal ","FC/Building","Status","LFD","Days to LFD"] if c in risk.columns]
                st.divider()
                st.markdown(f"#### ⏰ LFD Risk — {len(risk)} containers within 5 days")
                st.dataframe(risk[show].sort_values("Days to LFD"), use_container_width=True, hide_index=True)

        # Status + terminal breakdown
        if not _del.empty:
            s_col, t_col_name = st.columns(2)
            if "Status" in _del.columns:
                with s_col:
                    st.markdown("#### Delivery Status")
                    vc = _del["Status"].fillna("Unknown").value_counts().reset_index()
                    vc.columns = ["Status","Count"]
                    st.dataframe(vc, use_container_width=True, hide_index=True)
            term_col = "Terminal " if "Terminal " in _del.columns else ("Terminal" if "Terminal" in _del.columns else None)
            if term_col:
                with t_col_name:
                    st.markdown("#### By Terminal")
                    tc = _del[term_col].fillna("Unknown").value_counts().head(10).reset_index()
                    tc.columns = ["Terminal","Count"]
                    st.bar_chart(tc.set_index("Terminal")["Count"])

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 2 — Detention & Demurrage Risk (from Inbound Loads)
    # ══════════════════════════════════════════════════════════════════════
    st.divider()
    st.markdown("### Detention & Demurrage Risk")

    if not has_il:
        st.info("Upload the **Inbound Loads Report** in the sidebar to activate detention/demurrage tracking.")
    else:
        # detention = clock starts when container available at terminal;
        #             container not yet gated out; last_free_detention approaching
        det = il_df.copy()
        det["days_to_det"] = (det["last_free_detention"] - today_ts).dt.days
        det["days_to_dem"] = (det["last_free_demurrage"] - today_ts).dt.days

        # active detention risk: has free-time deadline, not yet gated out
        det_risk = det[
            det["last_free_detention"].notna() &
            (det["gateout_terminal"].isna() | (det["gateout_terminal"] > today_ts)) &
            (det["days_to_det"] <= 7)
        ].copy()

        dem_risk = det[
            det["last_free_demurrage"].notna() &
            (det["actual_arrival_discharge"].isna() | det["gateout_terminal"].isna()) &
            (det["days_to_dem"] <= 7)
        ].copy()

        d1, d2, d3, d4 = st.columns(4)
        det_overdue = int((det["days_to_det"].dropna() < 0).sum())
        dem_overdue = int((det["days_to_dem"].dropna() < 0).sum())
        d1.metric("🔴 Detention Overdue",  det_overdue)
        d2.metric("🟡 Detention ≤7 days",  max(0, len(det_risk) - det_overdue))
        d3.metric("🔴 Demurrage Overdue",  dem_overdue)
        d4.metric("🟡 Demurrage ≤7 days",  max(0, len(dem_risk) - dem_overdue))

        if not det_risk.empty or not dem_risk.empty:
            drisk_tab, ddem_tab = st.tabs(["⚠️ Detention Risk", "⚠️ Demurrage Risk"])

            with drisk_tab:
                if det_risk.empty:
                    st.success("No detention risk in next 7 days.")
                else:
                    det_risk["⚠️"] = det_risk["days_to_det"].apply(_risk_flag)
                    show_d = [c for c in ["⚠️","container_id","supplier","destination_fc",
                                          "dest_port","last_free_detention","days_to_det",
                                          "eta_discharge","gateout_terminal"]
                              if c in det_risk.columns]
                    st.dataframe(
                        det_risk[show_d].sort_values("days_to_det"),
                        use_container_width=True, hide_index=True
                    )

            with ddem_tab:
                if dem_risk.empty:
                    st.success("No demurrage risk in next 7 days.")
                else:
                    dem_risk["⚠️"] = dem_risk["days_to_dem"].apply(_risk_flag)
                    show_m = [c for c in ["⚠️","container_id","supplier","destination_fc",
                                          "dest_port","last_free_demurrage","days_to_dem",
                                          "eta_discharge","actual_arrival_discharge"]
                              if c in dem_risk.columns]
                    st.dataframe(
                        dem_risk[show_m].sort_values("days_to_dem"),
                        use_container_width=True, hide_index=True
                    )

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 3 — In-Transit Pipeline (from Shipment Status)
    # ══════════════════════════════════════════════════════════════════════
    st.divider()
    st.markdown("### In-Transit Pipeline")

    if not has_ss:
        st.info("Upload the **Import Shipment Status** report in the sidebar to activate pipeline tracking.")
    else:
        # active = not yet delivered
        active_ss = ss_df[ss_df["delivered"].isna()].copy() if "delivered" in ss_df.columns else ss_df.copy()

        p1,p2,p3,p4 = st.columns(4)
        p1.metric("Total Active Containers",  len(active_ss))
        p2.metric("Unique FCs",     active_ss["fc"].nunique() if "fc" in active_ss.columns else 0)
        p3.metric("Unique Vessels", active_ss["vessel_name"].nunique() if "vessel_name" in active_ss.columns else 0)
        p4.metric("Unique Carriers",active_ss["carrier_code"].nunique() if "carrier_code" in active_ss.columns else 0)

        col_pip1, col_pip2 = st.columns(2)

        with col_pip1:
            st.markdown("#### By Status")
            if "current_status" in active_ss.columns:
                sc = active_ss["current_status"].fillna("Unknown").value_counts().reset_index()
                sc.columns = ["Status","Containers"]
                st.dataframe(sc, use_container_width=True, hide_index=True)

        with col_pip2:
            st.markdown("#### By Destination FC")
            if "fc" in active_ss.columns:
                fc_c = active_ss["fc"].fillna("Unknown").value_counts().head(12).reset_index()
                fc_c.columns = ["FC","Containers"]
                st.bar_chart(fc_c.set_index("FC")["Containers"])

        # ETA table: upcoming arrivals in next 30 days
        if "discharge_eta" in active_ss.columns:
            soon = active_ss[
                active_ss["discharge_eta"].notna() &
                (active_ss["discharge_eta"] >= today_ts) &
                (active_ss["discharge_eta"] <= today_ts + pd.Timedelta(days=30))
            ].copy()
            if not soon.empty:
                soon["Days to ETA"] = (soon["discharge_eta"] - today_ts).dt.days
                st.divider()
                st.markdown(f"#### 🚢 Arriving at Port — Next 30 Days ({len(soon)} containers)")
                show_s = [c for c in ["container_no","vendor_name","fc","vessel_name",
                                      "discharge_city","discharge_eta","Days to ETA","current_status"]
                          if c in soon.columns]
                st.dataframe(
                    soon[show_s].sort_values("Days to ETA"),
                    use_container_width=True, hide_index=True
                )

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 4 — Carrier Intelligence
    # ══════════════════════════════════════════════════════════════════════
    st.divider()
    st.markdown("### Carrier Intelligence")

    if sub_df.empty:
        st.info("No carrier submissions yet — insights will appear as data comes in.")
    else:
        s1,s2,s3,s4 = st.columns(4)
        s1.metric("Total Submissions", len(sub_df))
        s2.metric("Unique Carriers",   sub_df["carrier_name"].nunique())
        s3.metric("Unique Containers", sub_df["container_id"].nunique())
        last_s = sub_df["submitted_at"].max()
        s4.metric("Last Submission", last_s.strftime("%b %d %H:%M") if pd.notna(last_s) else "—")

        ci1, ci2 = st.columns(2)
        with ci1:
            st.markdown("#### Submission Volume")
            vol = sub_df.groupby("carrier_name").size().sort_values(ascending=False).reset_index()
            vol.columns = ["Carrier","Submissions"]
            st.bar_chart(vol.set_index("Carrier")["Submissions"])

        with ci2:
            if "within_sla" in sub_df.columns and sub_df["within_sla"].notna().any():
                st.markdown("#### SLA Rate by Carrier")
                sla = sub_df.dropna(subset=["within_sla"]).copy()
                sla["within_sla"] = sla["within_sla"].str.upper().str.strip()
                sp = sla.groupby(["carrier_name","within_sla"]).size().unstack(fill_value=0).reset_index()
                for c in ["YES","NO"]:
                    if c not in sp.columns: sp[c] = 0
                sp["Total"] = sp["YES"] + sp["NO"]
                sp["SLA %"] = (sp["YES"] / sp["Total"].replace(0,1) * 100).round(1)
                sp = sp.rename(columns={"carrier_name":"Carrier","YES":"Within SLA","NO":"Outside SLA"})
                sp = sp[["Carrier","Within SLA","Outside SLA","Total","SLA %"]].sort_values("SLA %",ascending=False)
                try:
                    st.dataframe(
                        sp.style.format({"SLA %":"{:.1f}%"})
                          .background_gradient(subset=["SLA %"],cmap="RdYlGn",vmin=0,vmax=100),
                        use_container_width=True, hide_index=True
                    )
                except Exception:
                    st.dataframe(sp, use_container_width=True, hide_index=True)

        # freshness
        fresh = (sub_df.groupby("carrier_name")["submitted_at"].max()
                 .reset_index().rename(columns={"carrier_name":"Carrier","submitted_at":"Last Submitted"})
                 .sort_values("Last Submitted",ascending=False))
        fresh["Days Since"] = (today_ts - fresh["Last Submitted"].dt.normalize()).dt.days
        fresh["⚠️"] = fresh["Days Since"].apply(
            lambda d: "🔴 Stale (7+ days)" if d>=7 else ("🟡 4-6 days" if d>=4 else "🟢 Recent")
        )
        fresh["Last Submitted"] = fresh["Last Submitted"].dt.strftime("%b %d %H:%M")
        st.markdown("#### Submission Freshness")
        st.dataframe(fresh, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 5 — Action Items
    # ══════════════════════════════════════════════════════════════════════
    st.divider()
    st.markdown("### 🚨 Action Items")
    st.caption("Priority flags auto-generated from all loaded data sources.")

    actions = []

    # From DBR
    if dbr_sheets:
        if overdue_ct > 0:
            actions.append(("🔴 Critical", f"{overdue_ct} container(s) with overdue empty returns — DBR",
                            "Go to Empty Returns tab for details."))
        if due_soon_ct > 0:
            actions.append(("🟡 Urgent", f"{due_soon_ct} container(s) with empty return due ≤3 days — DBR",
                            "Schedule appointments immediately."))
        if not _dem.empty:
            actions.append(("🔴 Critical", f"{len(_dem)} container(s) currently in demurrage — DBR",
                            "Costs accruing daily. Expedite gateout."))
        if dbr_sheets and not _del.empty and "LFD" in _del.columns:
            _lfd_chk = _del.copy()
            _lfd_chk["_lfd"]  = pd.to_datetime(_lfd_chk["LFD"], errors="coerce")
            _lfd_chk["_days"] = (_lfd_chk["_lfd"] - today_ts).dt.days
            lfd_r = _lfd_chk[_lfd_chk["_days"].notna() & (_lfd_chk["_days"] <= 2)]
            if not lfd_r.empty:
                actions.append(("🔴 Critical", f"{len(lfd_r)} container(s) with LFD within 2 days — DBR",
                                "Expedite or notify carrier."))

    # From Inbound Loads
    if has_il:
        if det_overdue > 0:
            actions.append(("🔴 Critical", f"{det_overdue} container(s) past detention free time — Inbound Loads",
                            "Detention fees accruing. Contact dray carrier immediately."))
        if dem_overdue > 0:
            actions.append(("🔴 Critical", f"{dem_overdue} container(s) past demurrage free time — Inbound Loads",
                            "Demurrage fees accruing. Contact ocean carrier / terminal."))
        det_warn = len(det_risk) - det_overdue
        if det_warn > 0:
            actions.append(("🟡 Urgent", f"{det_warn} container(s) with detention deadline within 7 days",
                            "Schedule dray pickup before free time expires."))

    # From Carrier Submissions
    if not sub_df.empty and "fresh" in dir():
        stale = fresh[fresh["Days Since"] >= 7] if "Days Since" in fresh.columns else pd.DataFrame()
        if not stale.empty:
            carriers_stale = ", ".join(stale["Carrier"].tolist())
            actions.append(("🟡 Attention", f"{len(stale)} carrier(s) with no submission in 7+ days",
                            f"Follow up: {carriers_stale}"))

    if not actions:
        st.success("No open action items across all loaded data sources.")
    else:
        priority_order = {"🔴 Critical": 0, "🟡 Urgent": 1, "🟡 Attention": 2}
        actions.sort(key=lambda x: priority_order.get(x[0], 9))
        for priority, title, detail in actions:
            with st.container(border=True):
                pr_col, txt_col = st.columns([1, 5])
                pr_col.markdown(f"**{priority}**")
                txt_col.markdown(f"**{title}**  \n{detail}")
