from __future__ import annotations
import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import io
import os
from datetime import datetime, date, timedelta

from utils import (
    normalize_container, containers_match,
    parse_container_list, parse_carrier_file,
    parse_carrier_template, TEMPLATE_SHEET_MAP,
)
import data_sync

# ── page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Robotics Container Tracker",
    page_icon=None,
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

    st.title("Robotics Container Tracker")
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
        CREATE TABLE IF NOT EXISTS delivery_plan (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            week_start   TEXT NOT NULL,
            appt_date    TEXT,
            appt_time    TEXT,
            slot_num     INTEGER,
            container_id TEXT NOT NULL,
            carrier      TEXT NOT NULL,
            product_type TEXT DEFAULT 'SHELF, HDTP',
            qty          INTEGER,
            notes        TEXT,
            status       TEXT DEFAULT 'SCHEDULED',
            created_at   TEXT,
            updated_at   TEXT
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
        CREATE TABLE IF NOT EXISTS dbr_receipts (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            carrier       TEXT NOT NULL,
            week_start    TEXT NOT NULL,
            received_date TEXT NOT NULL,
            received_via  TEXT DEFAULT 'manual',
            file_name     TEXT,
            notes         TEXT,
            logged_at     TEXT NOT NULL,
            UNIQUE(carrier, week_start, received_date)
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
    st.header("DBR File")

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
            st.success("Saved to S3" if ok else "Saved locally (S3 unavailable)")
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
    st.markdown("**Additional Reports**")
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
        st.success(f"{n_ss:,} containers loaded")

    if st.session_state.get("ss_loaded"):
        st.caption(f"Shipment Status: {st.session_state.get('ss_source','')}")

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
        st.success(f"{n_il:,} containers loaded")

    if st.session_state.get("il_loaded"):
        st.caption(f"Inbound Loads: {st.session_state.get('il_source','')}")

    st.divider()
    st.caption("Container Tracker v1.2")
    if S3_ENABLED:
        st.caption("S3 sync active")

# ── tabs ───────────────────────────────────────────────────────────────────────
st.title("Robotics Container Tracker")
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["Container Lookup", "Carrier Submission", "Empty Returns", "Carrier Data", "Lane Costs", "Insights", "Planning"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Container Lookup
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Container Lookup")
    st.caption("🎯 **Purpose:** Find the current status of any container across all DBR sheets. Paste a list of IDs and get instant results — no manual spreadsheet searching required.")
    st.caption("📋 **How to use:** Upload the DBR in the sidebar, paste container IDs (one per line) into the box, then click Search.")

    col_a, col_b = st.columns([2, 1])
    with col_a:
        lookup_input = st.text_area(
            "Container IDs", height=260, label_visibility="collapsed",
            placeholder="TCNU389902-4\nONEU624470-0\nGCXU542076\n...",
        )
    with col_b:
        requester = st.text_input("Your name (optional)", placeholder="e.g. Dominique")
        st.caption("Searches across: Delivery Appts · Empty Returns · On Vessel · Canceled · Demurrage · Accessorials · IDs with or without check digit both work")
        search_btn = st.button("Search", type="primary", use_container_width=True)

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

            # ── Status + Detail helpers ────────────────────────────────────────────────
            def _smart_status(sheet, row):
                sv = str(row.get("Status", "") or "").strip().lower()
                if sheet == "Delivery Appointments":
                    if "delivered" in sv: return "Delivered"
                    if "in transit" in sv: return "In Transit"
                    return "Scheduled for Delivery"
                if sheet == "Empty Returns":
                    if "terminat" in sv or "returned" in sv:
                        return "Delivered / Empty Returned"
                    return "Delivered / Pending Return"
                if sheet == "On Vessel":    return "On Vessel"
                if sheet == "Canceled":     return "Canceled"
                if sheet == "Demurrage":    return "Demurrage Hold"
                if sheet == "Accessorials": return "Accessorial Charge"
                return sheet

            def _key_detail(sheet, row):
                def g(col): return str(row.get(col) or "").strip()
                if sheet == "Delivery Appointments":
                    appt = g("Del Appointment Date/Time") or g("Outgate Date")
                    return " · ".join(p for p in [g("FC/Building"), appt] if p)
                if sheet == "Empty Returns":
                    due = g("Empty Return Due Date")
                    return " · ".join(p for p in [g("Terminal"), ("Due " + due) if due else ""] if p)
                if sheet == "On Vessel":
                    lfd = g("LFD")
                    return " · ".join(p for p in [g("Location"), ("LFD " + lfd) if lfd else ""] if p)
                if sheet == "Canceled":
                    return " · ".join(p for p in [g("FC/Building"), g("Del Appointment Date/Time")] if p)
                if sheet == "Demurrage":
                    days = g("Days in Hold")
                    return " · ".join(p for p in [g("Type of Hold"), (days + " days") if days else ""] if p)
                if sheet == "Accessorials":
                    return " · ".join(p for p in [g("Accessorial"), g("Date(s) Incurred")] if p)
                return ""

            # ── Build summary table ───────────────────────────────────────────────
            summary_rows = []
            if not results_df.empty:
                for _, row in results_df.iterrows():
                    sheet = row.get("Sheet", "")
                    container = row.get("Container #") or row.get("Searched As", "")
                    summary_rows.append({
                        "Container": str(container).strip(),
                        "Status":    _smart_status(sheet, row),
                        "Detail":    _key_detail(sheet, row),
                    })
            summary_df = (pd.DataFrame(summary_rows) if summary_rows
                          else pd.DataFrame(columns=["Container", "Status", "Detail"]))

            # ── Status bucket metrics ──────────────────────────────────────────────────
            BUCKET_ORDER = [
                "Scheduled for Delivery", "In Transit", "On Vessel",
                "Delivered", "Delivered / Pending Return", "Delivered / Empty Returned",
                "Demurrage Hold", "Accessorial Charge", "Canceled",
            ]
            counts = {b: 0 for b in BUCKET_ORDER}
            for s in (summary_df["Status"] if not summary_df.empty else []):
                if s in counts: counts[s] += 1
            active = [(k, v) for k, v in counts.items() if v > 0]
            active.append(("Not Found", len(not_found)))
            metric_cols = st.columns(len(active))
            for col, (label, val) in zip(metric_cols, active):
                col.metric(label, val)

            st.divider()

            # ── Consolidated summary table ────────────────────────────────────────────
            if not summary_df.empty:
                st.dataframe(
                    summary_df, use_container_width=True, hide_index=True,
                    column_config={
                        "Container": st.column_config.TextColumn("Container", width="small"),
                        "Status":    st.column_config.TextColumn("Status",    width="medium"),
                        "Detail":    st.column_config.TextColumn("Detail",    width="large"),
                    },
                )
            else:
                st.warning("None of the queried containers found in the DBR.")

            if not_found:
                st.markdown(f"**Not found in this DBR ({len(not_found)}):**")
                st.dataframe(pd.DataFrame({"Container ID": not_found}),
                             use_container_width=True, hide_index=True)

            if not summary_df.empty or not_found:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    if not summary_df.empty:
                        summary_df.to_excel(w, index=False, sheet_name="Summary")
                        results_df.to_excel(w, index=False, sheet_name="Full Details")
                    if not_found:
                        pd.DataFrame({"Not Found": not_found}).to_excel(
                            w, index=False, sheet_name="Not Found")
                buf.seek(0)
                st.download_button(
                    "⬇️ Download results (.xlsx)", data=buf,
                    file_name=f"container_lookup_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            if not results_df.empty:
                with st.expander("🔍 Full details (all fields)"):
                    st.dataframe(results_df, use_container_width=True, hide_index=True)




# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Carrier Submission
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Carrier Submission Portal")
    st.caption("🎯 **Purpose:** The single intake point for all carrier status updates. Carriers upload their weekly template here so you have one place to review submissions instead of managing files across email.")
    st.caption("📋 **How to use:** Upload a completed AGL Carrier Template and click Confirm & Submit, or share the vendor portal link with carriers to submit directly.")

    _vp_col, _tmpl_col = st.columns([3, 1])
    with _vp_col:
        st.info(
            f"**🔗 Share this link with your carriers:**  \n"
            f"[{VENDOR_PORTAL_URL}]({VENDOR_PORTAL_URL})  \n"
            "Carriers submit directly through this portal — no manual upload needed on your end."
        )
    with _tmpl_col:
        tmpl_path = os.path.join(BASE_DIR, "agl_carrier_template.xlsx")
        if os.path.exists(tmpl_path):
            with open(tmpl_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Carrier Template",
                    data=f.read(),
                    file_name="AGL_Carrier_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    st.divider()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # DBR Receipt Tracker
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.markdown("### 📋 DBR Receipt Tracker")
    st.caption("Track whether each carrier has submitted their weekly DBR. Use the week selector to navigate; missing submissions are flagged automatically.")

    _today_rd  = date.today()
    _this_mon  = _today_rd - timedelta(days=_today_rd.weekday())
    _wk_input  = st.date_input(
        "Week (select any day)", value=_this_mon, key="receipt_wk_sel",
        label_visibility="collapsed",
    )
    _wk_start = _wk_input - timedelta(days=_wk_input.weekday())
    _wk_end   = _wk_start + timedelta(days=4)
    st.caption(f"**{_wk_start.strftime('%b %-d')} – {_wk_end.strftime('%b %-d, %Y')}**")

    TRACKED_CARRIERS = ["ATMI", "ARVY", "HDDR", "RKNE", "TGHE"]

    _conn_rd = get_db()
    _rec_df  = pd.read_sql(
        "SELECT carrier, received_date FROM dbr_receipts WHERE week_start = ?",
        _conn_rd, params=[_wk_start.isoformat()]
    )
    _conn_rd.close()

    _rec_map = {}
    for _, _r in _rec_df.iterrows():
        _rec_map.setdefault(_r["carrier"], set()).add(_r["received_date"])

    _days     = [_wk_start + timedelta(days=i) for i in range(5)]
    _day_cols = [f"{d.strftime('%a')} {d.month}/{d.day}" for d in _days]

    _grid_rows, _missing_carriers = [], []
    for _carrier in TRACKED_CARRIERS:
        _row = {"Carrier": _carrier}
        _got = False
        for _d, _dc in zip(_days, _day_cols):
            if _d.isoformat() in _rec_map.get(_carrier, set()):
                _row[_dc] = "✅"
                _got = True
            else:
                _row[_dc] = ""
        if _got:
            _row["Status"] = "✅ Received"
        elif _wk_start <= _today_rd:
            _row["Status"] = "⚠️ Missing"
            _missing_carriers.append(_carrier)
        else:
            _row["Status"] = "—"
        _grid_rows.append(_row)

    _grid_df = pd.DataFrame(_grid_rows)
    st.dataframe(
        _grid_df, use_container_width=True, hide_index=True,
        column_config={
            "Carrier": st.column_config.TextColumn("Carrier", width="small"),
            "Status":  st.column_config.TextColumn("Status",  width="medium"),
            **{dc: st.column_config.TextColumn(dc, width="small") for dc in _day_cols},
        },
    )

    if _missing_carriers and _wk_start <= _today_rd:
        st.warning(f"**{len(_missing_carriers)} carrier(s) missing this week:** {', '.join(_missing_carriers)}")
        for _mc in _missing_carriers:
            with st.expander(f"📧 Follow-up message — {_mc}"):
                _tmpl = (
                    f"Hi {_mc} Team,\n\n"
                    f"We haven\u2019t received your container status DBR for the week of "
                    f"{_wk_start.strftime('%B %-d, %Y')}. Could you please submit your weekly "
                    f"update at your earliest convenience?\n\n"
                    f"You can submit via the AGL carrier portal: {VENDOR_PORTAL_URL}\n\n"
                    f"Thank you,\nAGL Robotics Logistics"
                )
                st.code(_tmpl.replace("\n", "\n"), language=None)

    st.divider()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Bulk DBR File Upload
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.markdown("### 📤 Submit DBR Files")
    st.caption("Upload one or more files in one shot — mix carriers, mix weeks. Carrier is auto-detected from the filename; override per file if needed.")

    def _detect_carrier_from_name(fname: str) -> str:
        fn = fname.upper()
        for sc in ["ATMI", "ARVY", "RKNE", "TGHE"]:
            if sc in fn: return sc
        if "HUDD" in fn or "HDDR" in fn: return "HDDR"
        return ""

    _CARRIER_OPTS = ["", "ATMI", "ARVY", "HDDR", "RKNE", "TGHE"]

    bulk_files = st.file_uploader(
        "Drop DBR files here",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="bulk_dbr_upload",
        label_visibility="collapsed",
    )

    if bulk_files:
        st.caption(f"{len(bulk_files)} file(s) selected — confirm carrier for each:")

        _file_carrier_map = {}
        for _fi, _uf in enumerate(bulk_files):
            _det = _detect_carrier_from_name(_uf.name)
            _hd1, _hd2 = st.columns([4, 1])
            with _hd1:
                st.markdown(f"&nbsp;&nbsp;📄 `{_uf.name}`")
            with _hd2:
                _chosen = st.selectbox(
                    "Carrier",
                    _CARRIER_OPTS,
                    index=_CARRIER_OPTS.index(_det) if _det in _CARRIER_OPTS else 0,
                    key=f"bulk_c_{_fi}",
                    label_visibility="collapsed",
                )
            _file_carrier_map[_fi] = (_uf, _chosen)

        _unassigned = [_uf.name for _, (_uf, _c) in _file_carrier_map.items() if not _c]
        if _unassigned:
            st.warning(f"Assign carrier for: {', '.join(_unassigned)}")
        else:
            # Parse all files, group by carrier
            _all_parsed = {}
            _parse_errs = []
            for _fi, (_uf, _carrier) in _file_carrier_map.items():
                _ck = f"bulk_{_uf.name}_{_uf.size}"
                if _ck not in st.session_state:
                    st.session_state[_ck] = _uf.read()
                _raw = st.session_state[_ck]
                try:
                    _p = parse_carrier_template(_raw, _uf.name)
                    if _p:
                        if _carrier not in _all_parsed:
                            _all_parsed[_carrier] = {}
                        for _stype, _rows in _p.items():
                            _all_parsed[_carrier].setdefault(_stype, [])
                            for _r in _rows:
                                _r["_src"] = _uf.name
                            _all_parsed[_carrier][_stype].extend(_rows)
                    else:
                        _parse_errs.append(f"{_uf.name}: no container data found")
                except Exception as _e:
                    _parse_errs.append(f"{_uf.name}: {_e}")

            for _err in _parse_errs:
                st.warning(_err)

            if _all_parsed:
                _tot = sum(len(_r) for _cd in _all_parsed.values() for _r in _cd.values())
                st.success(f"Ready: **{_tot} containers** across **{len(_all_parsed)} carrier(s)**")

                for _carrier, _sheets in _all_parsed.items():
                    _ctot = sum(len(_r) for _r in _sheets.values())
                    with st.expander(f"🔎 Preview — {_carrier} ({_ctot} containers)"):
                        _slabels = {"delivery": "Delivery", "delivery_ilm1": "ILM1",
                                    "delivery_ric6": "RIC6", "empty_return": "Empty Returns",
                                    "demurrage": "Demurrage", "accessorial": "Accessorials", "ody": "ODY"}
                        _ptabs = st.tabs([_slabels.get(_k, _k) for _k in _sheets])
                        for _ptab, (_st, _rows) in zip(_ptabs, _sheets.items()):
                            with _ptab:
                                _df_p = pd.DataFrame(_rows).drop(columns=["_raw_container","_src"], errors="ignore")
                                st.dataframe(_df_p, use_container_width=True, hide_index=True)

                if st.button("✅ Confirm & Submit All", type="primary", key="bulk_submit_btn"):
                    _now = datetime.now().isoformat()
                    _conn = get_db()
                    _logged = 0
                    for _carrier, _sheets in _all_parsed.items():
                        for _stype, _rows in _sheets.items():
                            for _row in _rows:
                                _conn.execute(
                                    """INSERT INTO carrier_submissions
                                       (submitted_at, carrier_name, container_id, sheet_type,
                                        port, terminal, fc_building, flexi_id, outgate_date,
                                        delivery_date, status, within_sla, sla_notes,
                                        empty_return_due, appointment_date, accessorial_type,
                                        notes, source_file, source)
                                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                    (_now, _carrier, _row["container_id"], _stype,
                                     _row.get("port"), _row.get("terminal"), _row.get("fc_building"),
                                     _row.get("flexi_id"), _row.get("outgate_date"),
                                     _row.get("delivery_date"), _row.get("status"),
                                     _row.get("within_sla"), _row.get("sla_notes"),
                                     _row.get("empty_return_due"), _row.get("appointment_date"),
                                     _row.get("accessorial_type"), _row.get("notes"),
                                     _row.get("_src"), "web")
                                )
                                _logged += 1
                        _wk = (date.today() - timedelta(days=date.today().weekday())).isoformat()
                        _conn.execute(
                            "INSERT OR IGNORE INTO dbr_receipts (carrier, week_start, received_date, received_via, logged_at) VALUES (?,?,?,?,?)",
                            (_carrier, _wk, date.today().isoformat(), "portal", _now)
                        )
                    _conn.commit()
                    _conn.close()
                    if S3_ENABLED:
                        data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                    st.success(f"✅ Submitted {_logged} containers across {len(_all_parsed)} carrier(s)")
                    st.rerun()

    st.divider()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Manual entry (fallback)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    with st.expander("⌨️ Enter containers manually (fallback — use file upload when possible)"):
        st.caption("Use this only if a carrier cannot provide a file. Paste container IDs directly.")
        with st.form("carrier_form_manual", clear_on_submit=True):
            _fm1, _fm2 = st.columns(2)
            with _fm1:
                _m_carrier  = st.selectbox("Carrier *", ["", "ATMI", "ARVY", "HDDR", "RKNE", "TGHE"], key="man_carrier")
                _m_terminal = st.text_input("Terminal", placeholder="EWR, BOS, CHI…", key="man_term")
            with _fm2:
                _m_status = st.selectbox("Status", [
                    "Delivery Scheduled", "Delivered pending empty return",
                    "Empty Return Scheduled", "TERMINATED", "On Vessel",
                    "Customs Hold", "Other",
                ], key="man_status")
                _m_notes = st.text_input("Notes", placeholder="Delays, exceptions…", key="man_notes")
            _m_ids_text = st.text_area("Container IDs (one per line) *", height=140, key="man_ids")
            _m_submit   = st.form_submit_button("Submit", type="primary")
        if _m_submit:
            if not _m_carrier:
                st.error("Select a carrier.")
            elif not _m_ids_text.strip():
                st.error("Paste at least one container ID.")
            else:
                _mids = parse_container_list(_m_ids_text)
                _mn   = datetime.now().isoformat()
                _mc   = get_db()
                for _cid in _mids:
                    _mc.execute(
                        "INSERT INTO carrier_submissions (submitted_at, carrier_name, container_id, terminal, status, notes, source) VALUES (?,?,?,?,?,?,?)",
                        (_mn, _m_carrier, normalize_container(_cid), _m_terminal.strip() or None, _m_status, _m_notes.strip() or None, "manual")
                    )
                _mwk = (date.today() - timedelta(days=date.today().weekday())).isoformat()
                _mc.execute(
                    "INSERT OR IGNORE INTO dbr_receipts (carrier, week_start, received_date, received_via, logged_at) VALUES (?,?,?,?,?)",
                    (_m_carrier, _mwk, date.today().isoformat(), "manual", _mn)
                )
                _mc.commit(); _mc.close()
                if S3_ENABLED:
                    data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                st.success(f"Logged {len(_mids)} containers from **{_m_carrier}**")

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Submission Log
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.markdown("### Submission Log")

    _sl_conn = get_db()
    _log_df  = pd.read_sql(
        """SELECT submitted_at, carrier_name, container_id, sheet_type,
                  terminal, status, notes, source_file, source
           FROM carrier_submissions ORDER BY submitted_at DESC LIMIT 500""",
        _sl_conn
    )
    _sl_conn.close()

    if _log_df.empty:
        st.info("No submissions yet.")
    else:
        _sc1, _sc2, _sc3 = st.columns(3)
        with _sc1:
            _sf_carrier = st.multiselect("Carrier", sorted(_log_df["carrier_name"].unique()), key="sl_carrier")
        with _sc2:
            _sf_status  = st.multiselect("Status",  sorted(_log_df["status"].dropna().unique()), key="sl_status")
        with _sc3:
            _sf_source  = st.multiselect("Source",  sorted(_log_df["source"].dropna().unique()), key="sl_source")

        _filt = _log_df.copy()
        if _sf_carrier: _filt = _filt[_filt["carrier_name"].isin(_sf_carrier)]
        if _sf_status:  _filt = _filt[_filt["status"].isin(_sf_status)]
        if _sf_source:  _filt = _filt[_filt["source"].isin(_sf_source)]

        st.dataframe(_filt, use_container_width=True, hide_index=True)

        _sl_buf = io.BytesIO()
        _filt.to_excel(_sl_buf, index=False)
        _sl_buf.seek(0)
        st.download_button(
            "⬇️ Export log (.xlsx)", data=_sl_buf,
            file_name=f"carrier_log_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )



# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Empty Returns Dashboard
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Empty Returns")
    st.caption("🎯 **Purpose:** Shows which containers still owe a terminal return after FC delivery, ranked by urgency. Overdue returns can trigger fees — this view tells you what needs carrier follow-up today.")
    st.caption("📋 **How to use:** Upload the DBR in the sidebar, then check Overdue and Due Soon first — those need immediate carrier follow-up.")

    if not dbr_sheets:
        st.info("Upload the DBR file in the sidebar to see empty returns data.")
    else:
        er_raw = dbr_sheets.get("Empty Returns", pd.DataFrame()).copy()
        if er_raw.empty:
            st.warning("No Empty Returns sheet found in this DBR.")
        else:
            er_raw.columns = [str(c).strip() for c in er_raw.columns]
            er_raw["Empty Return Due Date"] = pd.to_datetime(
                er_raw["Empty Return Due Date"], errors="coerce")
            today = pd.Timestamp(date.today())

            # Terminated = resolved. Only open containers are tracked for risk.
            terminated = er_raw[
                er_raw["Status"].fillna("").str.upper() == "TERMINATED"
            ] if "Status" in er_raw.columns else pd.DataFrame()

            active = er_raw[
                er_raw["Status"].fillna("").str.upper().ne("TERMINATED") &
                er_raw["Empty Return Due Date"].notna()
            ].copy()
            active["Days Until Due"] = (active["Empty Return Due Date"] - today).dt.days
            active["Alert"] = active["Days Until Due"].apply(
                lambda d: "OVERDUE" if d < 0 else ("Due soon" if d <= 3 else "OK")
            )

            overdue  = active[active["Days Until Due"] < 0]
            due_soon = active[(active["Days Until Due"] >= 0) & (active["Days Until Due"] <= 3)]
            on_track = active[active["Days Until Due"] > 3]

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Open",         len(active))
            m2.metric("Overdue",      len(overdue))
            m3.metric("Due ≤ 3 days", len(due_soon))
            m4.metric("On track",     len(on_track))
            m5.metric("Terminated",   len(terminated), help="Return confirmed or obligation closed — no action needed")
            st.divider()

            view = st.radio(
                "Show",
                ["Overdue", "Due Soon (≤3 days)", "All Open"],
                horizontal=True,
            )
            display = (
                overdue  if view == "Overdue" else
                due_soon if view == "Due Soon (≤3 days)" else
                active
            )

            show_cols = [c for c in [
                "Alert", "Container #", "Terminal",
                "Empty Return Due Date", "Appointment Date",
                "Status", "Days Until Due", "If Outside Window - Reason",
            ] if c in display.columns]

            if display.empty:
                st.success("No containers in this category.")
            else:
                st.dataframe(
                    display[show_cols].sort_values("Days Until Due", ascending=True)
                    if "Days Until Due" in display.columns else display[show_cols],
                    use_container_width=True, hide_index=True,
                )
                buf = io.BytesIO()
                display.to_excel(buf, index=False)
                buf.seek(0)
                st.download_button(
                    "Export (.xlsx)", data=buf,
                    file_name=f"empty_returns_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            # Terminated — collapsed by default, clearly labelled as resolved
            if not terminated.empty:
                with st.expander(f"Terminated / Resolved ({len(terminated)} containers)", expanded=False):
                    st.caption("These containers have had their return obligation closed. No action needed.")
                    term_cols = [c for c in ["Container #", "Terminal", "Empty Return Due Date",
                                              "Status", "If Outside Window - Reason"]
                                 if c in terminated.columns]
                    st.dataframe(terminated[term_cols] if term_cols else terminated,
                                 use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — Carrier Data (structured view of all submissions)
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Carrier Data")
    st.caption("🎯 **Purpose:** A full record of every container status submission from carriers. Use this to verify what a carrier reported, check SLA performance, or audit delivery and return activity.")
    st.caption("📋 **How to use:** Filter by carrier or sheet type at the top, then drill into sub-tabs for Delivery, Empty Returns, Demurrage, and Accessorials.")

    conn = get_db()
    all_df = pd.read_sql(
        """SELECT submitted_at, carrier_name, container_id, sheet_type,
                  port, terminal, fc_building, delivery_date, status,
                  within_sla, empty_return_due, appointment_date,
                  accessorial_type, notes, source
           FROM carrier_submissions ORDER BY submitted_at DESC""",
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
                "delivery":       "Delivery",
                "delivery_ilm1":  "ILM1",
                "delivery_ric6":  "RIC6",
                "empty_return":   "Empty Returns",
                "ody":            "ODY/Storage",
                "demurrage":      "Demurrage",
                "accessorial":    "Accessorials",
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
    st.caption("🎯 **Purpose:** Compare drayage rates across carriers by lane and model the total cost of a delivery scenario. Use this when allocating containers across carriers or evaluating cost trade-offs.")
    st.caption("📋 **How to use:** Check the Rate Matrix for cheapest carrier per lane, then use the Scenario Builder to model cost by entering lane and container count.")

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

    if not rates_df.empty:
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
                pivot["Cheapest"] = pivot[carrier_cols].idxmin(axis=1)
                pivot["Low $"]  = pivot[carrier_cols].min(axis=1)
                pivot["Spread"] = pivot[carrier_cols].max(axis=1) - pivot["Low $"]

            col_order = (["port", "node_code", "FC", "Type", "Facility"]
                         + carrier_cols
                         + ["Cheapest", "Low $", "Spread"])
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

        if st.button("Add Lane to Scenario", type="primary"):
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
                if st.button("Clear All", type="secondary", use_container_width=True):
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
    st.caption("🎯 **Purpose:** Your weekly operations briefing — auto-generated from all loaded data sources. Start here to see what needs attention before you open anything else.")
    st.caption("📋 **How to use:** Upload the DBR and supplemental reports in the sidebar — sections activate automatically. Start with Detention & Demurrage Risk.")

    today_ts = pd.Timestamp(date.today())

    def _risk_flag(d):
        if pd.isna(d): return "—"
        if d < 0:   return f"{abs(int(d))}d overdue"
        if d == 0:  return "Due today"
        if d <= 3:  return f"{int(d)}d"
        return f"{int(d)}d ok"

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
        c5.metric("ER Overdue",     overdue_ct)
        c6.metric("ER Due ≤3 days", due_soon_ct)

        # LFD risk
        if not _del.empty and "LFD" in _del.columns:
            lfd = _del.copy()
            lfd["_lfd"]  = pd.to_datetime(lfd["LFD"], errors="coerce")
            lfd["Days to LFD"] = (lfd["_lfd"] - today_ts).dt.days
            risk = lfd[lfd["Days to LFD"].notna() & (lfd["Days to LFD"] <= 5)].copy()
            if not risk.empty:
                risk["Alert"] = risk["Days to LFD"].apply(_risk_flag)
                show = [c for c in ["Alert","Container #","Terminal ","FC/Building","Status","LFD","Days to LFD"] if c in risk.columns]
                st.divider()
                st.markdown(f"#### LFD Risk — {len(risk)} containers within 5 days")
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
        d1.metric("Detention Overdue",  det_overdue)
        d2.metric("Detention ≤7 days",  max(0, len(det_risk) - det_overdue))
        d3.metric("Demurrage Overdue",  dem_overdue)
        d4.metric("Demurrage ≤7 days",  max(0, len(dem_risk) - dem_overdue))

        if not det_risk.empty or not dem_risk.empty:
            drisk_tab, ddem_tab = st.tabs(["Detention Risk", "Demurrage Risk"])

            with drisk_tab:
                if det_risk.empty:
                    st.success("No detention risk in next 7 days.")
                else:
                    det_risk["Alert"] = det_risk["days_to_det"].apply(_risk_flag)
                    show_d = [c for c in ["Alert","container_id","supplier","destination_fc",
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
                    dem_risk["Alert"] = dem_risk["days_to_dem"].apply(_risk_flag)
                    show_m = [c for c in ["Alert","container_id","supplier","destination_fc",
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
                st.markdown(f"#### Arriving at Port — Next 30 Days ({len(soon)} containers)")
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
        fresh["Alert"] = fresh["Days Since"].apply(
            lambda d: "Stale (7+ days)" if d>=7 else ("4-6 days" if d>=4 else "Recent")
        )
        fresh["Last Submitted"] = fresh["Last Submitted"].dt.strftime("%b %d %H:%M")
        st.markdown("#### Submission Freshness")
        st.dataframe(fresh, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 5 — Action Items
    # ══════════════════════════════════════════════════════════════════════
    st.divider()
    st.markdown("### Action Items")
    st.caption("Priority flags auto-generated from all loaded data sources.")

    actions = []

    # From DBR
    if dbr_sheets:
        if overdue_ct > 0:
            actions.append(("Critical", f"{overdue_ct} container(s) with overdue empty returns — DBR",
                            "Go to Empty Returns tab for details."))
        if due_soon_ct > 0:
            actions.append(("Urgent", f"{due_soon_ct} container(s) with empty return due ≤3 days — DBR",
                            "Schedule appointments immediately."))
        if not _dem.empty:
            actions.append(("Critical", f"{len(_dem)} container(s) currently in demurrage — DBR",
                            "Costs accruing daily. Expedite gateout."))
        if dbr_sheets and not _del.empty and "LFD" in _del.columns:
            _lfd_chk = _del.copy()
            _lfd_chk["_lfd"]  = pd.to_datetime(_lfd_chk["LFD"], errors="coerce")
            _lfd_chk["_days"] = (_lfd_chk["_lfd"] - today_ts).dt.days
            lfd_r = _lfd_chk[_lfd_chk["_days"].notna() & (_lfd_chk["_days"] <= 2)]
            if not lfd_r.empty:
                actions.append(("Critical", f"{len(lfd_r)} container(s) with LFD within 2 days — DBR",
                                "Expedite or notify carrier."))

    # From Inbound Loads
    if has_il:
        if det_overdue > 0:
            actions.append(("Critical", f"{det_overdue} container(s) past detention free time — Inbound Loads",
                            "Detention fees accruing. Contact dray carrier immediately."))
        if dem_overdue > 0:
            actions.append(("Critical", f"{dem_overdue} container(s) past demurrage free time — Inbound Loads",
                            "Demurrage fees accruing. Contact ocean carrier / terminal."))
        det_warn = len(det_risk) - det_overdue
        if det_warn > 0:
            actions.append(("Urgent", f"{det_warn} container(s) with detention deadline within 7 days",
                            "Schedule dray pickup before free time expires."))

    # From Carrier Submissions
    if not sub_df.empty and "fresh" in dir():
        stale = fresh[fresh["Days Since"] >= 7] if "Days Since" in fresh.columns else pd.DataFrame()
        if not stale.empty:
            carriers_stale = ", ".join(stale["Carrier"].tolist())
            actions.append(("Attention", f"{len(stale)} carrier(s) with no submission in 7+ days",
                            f"Follow up: {carriers_stale}"))

    if not actions:
        st.success("No open action items across all loaded data sources.")
    else:
        priority_order = {"Critical": 0, "Urgent": 1, "Attention": 2}
        actions.sort(key=lambda x: priority_order.get(x[0], 9))
        for priority, title, detail in actions:
            with st.container(border=True):
                pr_col, txt_col = st.columns([1, 5])
                pr_col.markdown(f"**{priority}**")
                txt_col.markdown(f"**{title}**  \n{detail}")



# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — Planning / Scheduler (v2)
# ══════════════════════════════════════════════════════════════════════════════

_PLAN_SLOTS = [
    ("7:30 AM", 1), ("8:30 AM", 2), ("9:30 AM", 3),
    ("10:30 AM", 4), ("11:30 AM", 5),
    ("1:00 PM", 6), ("2:00 PM", 7), ("3:00 PM", 8), ("3:30 PM", 9),
]
_PLAN_PRODUCTS = ["SHELF, HDTP", "STRAP, HDTP", "SHELF, CLDTP", "HDPT 2.0", "BASE ASSEMBLY, HDTP", "OTHER"]
_SCAC_COLOR = {
    "ATMI": "#E8A020", "ARVY": "#4285F4",
    "HDDR": "#34A853", "RKNE": "#9C27B0", "TGHE": "#E91E63",
}
_PLAN_STATUSES = ["SCHEDULED", "DELIVERED", "REJECTED", "RESCHEDULED", "HOLD", "PENDING"]
_STATUS_BADGE = {
    "SCHEDULED": "Scheduled", "DELIVERED": "Delivered", "REJECTED": "Rejected",
    "RESCHEDULED": "Rescheduled", "HOLD": "Hold", "PENDING": "Pending",
}

def _init_plan_config():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS plan_sites (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            site_code    TEXT NOT NULL UNIQUE,
            fc_name      TEXT,
            port         TEXT,
            region       TEXT,
            capacity_day INTEGER DEFAULT 9,
            active       INTEGER DEFAULT 1,
            notes        TEXT
        );
        CREATE TABLE IF NOT EXISTS plan_carriers (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            scac         TEXT NOT NULL UNIQUE,
            carrier_name TEXT NOT NULL,
            ops_contact  TEXT,
            ops_email    TEXT,
            active       INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS plan_site_carrier (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            site_code          TEXT NOT NULL,
            scac               TEXT NOT NULL,
            site_contact       TEXT,
            site_contact_email TEXT,
            priority_time      TEXT,
            allocation_pct     REAL DEFAULT 0,
            active             INTEGER DEFAULT 1,
            notes              TEXT,
            UNIQUE(site_code, scac)
        );
    """)
    conn.commit()
    try:
        conn.execute("ALTER TABLE delivery_plan ADD COLUMN site_code TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass
    if conn.execute("SELECT COUNT(*) FROM plan_sites").fetchone()[0] == 0:
        conn.executemany(
            "INSERT OR IGNORE INTO plan_sites (site_code,fc_name,port,region,capacity_day,active,notes) VALUES (?,?,?,?,?,?,?)",
            [
                ("RIC6","Richmond FC",       "ORF","East",      9,1,""),
                ("ILM1","Wilmington NC",     "ORF","East",      9,1,"Knightrider (NCKR) for overflow"),
                ("IAG1","Niagara NY",        "ORF","East",      4,1,"Via XPH1 transload — 53ft trailer"),
                ("DBM6","DBM6",              "ORF","East",      6,1,""),
                ("ATL2","ATL2",              "SAV","Southeast", 6,1,""),
                ("DRO1","DRO1",              "ORF","East",      4,1,""),
                ("LAX", "LAX / West Coast",  "LAX","West",      6,1,""),
                ("OAK", "OAK / Northern CA", "OAK","West",      6,1,""),
                ("RSW9","Fort Myers FL",     "MIA","Southeast", 4,0,"NEW — starts 11/2/2026"),
            ]
        )
    if conn.execute("SELECT COUNT(*) FROM plan_carriers").fetchone()[0] == 0:
        conn.executemany(
            "INSERT OR IGNORE INTO plan_carriers (scac,carrier_name,ops_contact,ops_email) VALUES (?,?,?,?)",
            [
                ("ATMI","Cargomatic",       "Tyler Domingues",    "tdomingues@cargomatic.com"),
                ("ARVY","Arrive Logistics", "Tyler Spangler",     ""),
                ("HDDR","Maersk/HUDD",      "Emily Christenbury", ""),
                ("RKNE","Road One",         "Mark Brennan",       ""),
                ("TGHE","Tighe Logistics",  "Jose Alvarado",      "jalvarado@tighe-co.com"),
            ]
        )
    if conn.execute("SELECT COUNT(*) FROM plan_site_carrier").fetchone()[0] == 0:
        conn.executemany(
            "INSERT OR IGNORE INTO plan_site_carrier (site_code,scac,site_contact,site_contact_email,priority_time,allocation_pct,active,notes) VALUES (?,?,?,?,?,?,?,?)",
            [
                ("RIC6","ATMI","Tyler Domingues",    "tdomingues@cargomatic.com","8:30 AM",60.0,1,""),
                ("RIC6","HDDR","Sandji Ruffin",      "",                         "7:30 AM",20.0,1,"Before lunch"),
                ("RIC6","ARVY","Tyler Spangler",     "",                         "8:30 AM",20.0,1,""),
                ("ILM1","HDDR","Jerry Nesbit",       "",                         "7:30 AM",100.0,1,""),
                ("ILM1","ATMI","Tyler Domingues",    "tdomingues@cargomatic.com","8:30 AM",0.0, 1,"As needed"),
                ("IAG1","ATMI","Tyler Domingues",    "tdomingues@cargomatic.com","8:30 AM",100.0,1,"Via XPH1 transload"),
                ("DBM6","ATMI","Tyler Domingues",    "tdomingues@cargomatic.com","8:30 AM",100.0,1,""),
                ("DRO1","ATMI","Tyler Domingues",    "tdomingues@cargomatic.com","8:30 AM",100.0,1,""),
                ("ATL2","ATMI","Tyler Domingues",    "tdomingues@cargomatic.com","8:30 AM",50.0,1,""),
                ("ATL2","RKNE","Michelle Fileccia",  "",                         "8:30 AM",50.0,1,""),
                ("LAX", "ATMI","Anel Nunez",         "",                         "8:30 AM",50.0,1,""),
                ("LAX", "HDDR","Ailua Osoimalo",     "",                         "7:30 AM",50.0,1,"Desirae Swain backup"),
                ("OAK", "ATMI","Tyler Domingues",    "tdomingues@cargomatic.com","8:30 AM",20.0,1,""),
                ("OAK", "HDDR","Emily Christenbury", "",                         "7:30 AM",40.0,1,""),
                ("OAK", "ARVY","Tyler Spangler",     "",                         "8:30 AM",40.0,1,""),
                ("RSW9","RKNE","Chuck Henderson",    "",                         "8:30 AM",100.0,0,"NEW 11/2/26"),
            ]
        )
    # week config table — stores active receiving days per week
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS plan_week_config (
            week_start  TEXT NOT NULL UNIQUE,
            active_days TEXT NOT NULL DEFAULT 'Mon,Tue,Wed,Thu,Fri',
            updated_at  TEXT
        );
    """)
    conn.commit()
    conn.close()

_init_plan_config()

_ALL_DAYS   = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]
_DEFAULT_DAYS = ["Mon","Tue","Wed","Thu","Fri"]

def _get_week_config(week_iso: str) -> list[str]:
    conn = get_db()
    row = conn.execute("SELECT active_days FROM plan_week_config WHERE week_start=?", (week_iso,)).fetchone()
    conn.close()
    if row:
        return [d.strip() for d in row[0].split(",") if d.strip()]
    return list(_DEFAULT_DAYS)

def _save_week_config(week_iso: str, active_days: list[str]):
    conn = get_db()
    conn.execute(
        """INSERT INTO plan_week_config (week_start, active_days, updated_at)
           VALUES (?,?,?)
           ON CONFLICT(week_start) DO UPDATE SET active_days=excluded.active_days, updated_at=excluded.updated_at""",
        (week_iso, ",".join(active_days), datetime.utcnow().isoformat())
    )
    conn.commit()
    conn.close()



# ── plan DB helpers ───────────────────────────────────────────────────────────

def _plan_week_start(d: date) -> date:
    return d - timedelta(days=d.weekday())

def _get_plan(week_iso: str) -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM delivery_plan WHERE week_start=? ORDER BY appt_date, slot_num, id",
        conn, params=(week_iso,)
    )
    conn.close()
    return df

def _get_sites_df() -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM plan_sites ORDER BY site_code", conn)
    conn.close()
    return df

def _get_carriers_df() -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM plan_carriers ORDER BY scac", conn)
    conn.close()
    return df

def _get_scm_df() -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query(
        """SELECT m.*, s.fc_name, c.carrier_name
           FROM plan_site_carrier m
           LEFT JOIN plan_sites s ON m.site_code=s.site_code
           LEFT JOIN plan_carriers c ON m.scac=c.scac
           ORDER BY m.site_code, m.scac""", conn)
    conn.close()
    return df

def _add_plan_entry(week_start, appt_date, appt_time, slot_num,
                    container_id, scac, site_code, product_type, qty, notes, status):
    db_write(
        """INSERT INTO delivery_plan
           (week_start,appt_date,appt_time,slot_num,container_id,
            carrier,site_code,product_type,qty,notes,status,created_at,updated_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (week_start.isoformat(),
         appt_date.isoformat() if appt_date else None,
         appt_time, slot_num, container_id.upper().strip(),
         scac, site_code, product_type,
         int(qty) if qty else None,
         notes.strip() if notes else None, status,
         datetime.utcnow().isoformat(), datetime.utcnow().isoformat())
    )

def _update_plan_entry(row_id: int, **kwargs):
    conn = get_db()
    kwargs["updated_at"] = datetime.utcnow().isoformat()
    set_clause = ", ".join(f"{k}=?" for k in kwargs)
    conn.execute(f"UPDATE delivery_plan SET {set_clause} WHERE id=?",
                 list(kwargs.values()) + [row_id])
    conn.commit(); conn.close()
    if S3_ENABLED:
        data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)

def _delete_plan_entry(row_id: int):
    conn = get_db()
    conn.execute("DELETE FROM delivery_plan WHERE id=?", (row_id,))
    conn.commit(); conn.close()
    if S3_ENABLED:
        data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)

def _config_write(sql: str, params: tuple):
    conn = get_db()
    conn.execute(sql, params)
    conn.commit(); conn.close()


# ── DBR Tracker historical importer ──────────────────────────────────────────
def _parse_dbr_tracker(file_bytes: bytes) -> tuple:
    """Parse ToteASERs Robotics DBR Tracker.xlsx -> list of delivery_plan dicts.
    Returns (rows: list[dict], errors: list[str]).
    """
    import io as _bio
    wb = openpyxl.load_workbook(_bio.BytesIO(file_bytes), data_only=True)

    if "Delivery Plan" not in wb.sheetnames:
        return [], ["File missing 'Delivery Plan' sheet — is this the DBR Tracker?"]

    # Build container -> (product_type, qty) from Consolidated sheet
    prod_lookup: dict = {}
    if "Consolidated" in wb.sheetnames:
        for i, row in enumerate(wb["Consolidated"].iter_rows(values_only=True)):
            if i == 0:
                continue
            cid = row[2]
            if cid:
                prod_lookup[str(cid).strip()] = (
                    str(row[8]).strip() if row[8] else None,
                    int(row[9]) if row[9] else None,
                )

    def _scac(carrier_str):
        # "ARVY (Arrive)" -> "ARVY"; normalize HUDD → HDDR (Maersk)
        if not carrier_str:
            return None
        code = str(carrier_str).strip().split()[0].upper()
        return "HDDR" if code == "HUDD" else code

    STATUS_MAP = {
        "Delivered": "delivered",
        "At Yard":   "at_yard",
        "Pending":   "planned",
    }

    rows, errors = [], []
    for i, row in enumerate(wb["Delivery Plan"].iter_rows(values_only=True)):
        if i <= 1:
            continue  # row 0 = grouping label, row 1 = headers
        container = row[3]
        if not container:
            continue
        container = str(container).strip()
        fc_dest  = row[7]
        fc_sched = row[8]
        if not fc_dest or not fc_sched:
            errors.append(f"Row {i+1}: missing FC Destination or FC Sched Del for {container} — skipped")
            continue
        # Resolve appt_date
        if hasattr(fc_sched, "date"):
            appt_date = fc_sched.date()
        else:
            try:
                appt_date = datetime.strptime(str(fc_sched), "%Y-%m-%d").date()
            except Exception:
                errors.append(f"Row {i+1}: unparseable date '{fc_sched}' for {container} — skipped")
                continue
        # Notes: merge Notes field + Live/Drop
        note_parts = []
        if row[12]:
            note_parts.append(str(row[12]).strip())
        if row[14]:
            note_parts.append(f"[{str(row[14]).strip()}]")
        prod_type, qty = prod_lookup.get(container, (None, None))
        rows.append({
            "week_start":   _plan_week_start(appt_date).isoformat(),
            "appt_date":    appt_date.isoformat(),
            "appt_time":    None,
            "slot_num":     None,
            "container_id": container,
            "carrier":      _scac(row[1]),
            "site_code":    str(fc_dest).strip().upper(),
            "product_type": prod_type,
            "qty":          qty,
            "notes":        " | ".join(note_parts) if note_parts else None,
            "status":       STATUS_MAP.get(str(row[11]).strip() if row[11] else "", "planned"),
        })
    return rows, errors

# ── SharePoint DBR fetch ─────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner="Fetching DBR Tracker from SharePoint...")
def _fetch_dbr_from_sharepoint(tenant_id: str, client_id: str, client_secret: str) -> bytes | None:
    """Download the master DBR Tracker from SharePoint via MSAL client credentials.
    Cached for 1 hour to avoid hammering SharePoint on every rerun.
    Returns raw xlsx bytes, or None on failure.
    """
    try:
        import msal
        import requests as _rq
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
        result = app.acquire_token_for_client(
            scopes=["https://amazon.sharepoint.com/.default"]
        )
        token = result.get("access_token")
        if not token:
            return None
        # Download via SharePoint REST API using the server-relative path
        file_path = "/sites/AGLRobotics/Shared%20Documents/ToteASERs%20Robotics%20DBR%20Tracker.xlsx"
        url = (
            "https://amazon.sharepoint.com/sites/AGLRobotics"
            f"/_api/web/GetFileByServerRelativePath(decodedurl='{file_path}')/$value"
        )
        resp = _rq.get(
            url,
            headers={"Authorization": f"Bearer {token}", "Accept": "application/octet-stream"},
            timeout=30,
        )
        resp.raise_for_status()
        return resp.content
    except Exception as _sp_exc:
        return str(_sp_exc)  # return error string so caller can distinguish


def _sp_secrets() -> tuple[str | None, str | None, str | None]:
    """Return (tenant_id, client_id, client_secret) from Streamlit secrets, or Nones."""
    try:
        cfg = st.secrets["sharepoint"]
        return (
            cfg.get("TENANT_ID") or cfg.get("tenant_id"),
            cfg.get("CLIENT_ID")  or cfg.get("client_id"),
            cfg.get("CLIENT_SECRET") or cfg.get("client_secret"),
        )
    except Exception:
        return None, None, None


# ── request parser ────────────────────────────────────────────────────────────
import re as _re

def _parse_site_request(text: str) -> list[dict]:
    """
    Parse a stakeholder request message into a list of material line items.
    Returns: [{"material": str, "qty": int, "unit": str, "level_load": bool, "include_sat": bool, "note": str}]
    """
    results = []
    lines = text.strip().splitlines()
    mat_pattern = _re.compile(
        r"([A-Z][A-Z ,./&-]+?)[\s:]+(\d+)\s*(TLs?|containers?)?\s*(.*)$",
        _re.IGNORECASE
    )
    for line in lines:
        line = line.strip().lstrip("•-*").strip()
        m = mat_pattern.match(line)
        if not m:
            continue
        material = m.group(1).strip().upper().rstrip(",").rstrip()
        qty_str  = m.group(2)
        unit     = (m.group(3) or "containers").lower()
        rest     = (m.group(4) or "").lower()
        # normalise material names
        material = material.replace("  ", " ")
        if "shelf" in material.lower() and "hdtp" in material.lower():
            material = "SHELF, HDTP"
        elif "strap" in material.lower() and "hdtp" in material.lower():
            material = "STRAP, HDTP"
        elif "base" in material.lower():
            material = "BASE ASSEMBLY, HDTP"
        elif "column" in material.lower():
            material = "COLUMN"
        level_load  = "level load" in rest
        include_sat = "saturday" in rest or "sat" in rest
        results.append({
            "material":    material,
            "qty":         int(qty_str),
            "unit":        "TL" if "tl" in unit else "container",
            "level_load":  level_load,
            "include_sat": include_sat,
            "note":        rest.strip("() "),
        })
    return results

def _generate_level_load_schedule(
    containers: list[str],
    site_code: str,
    scac: str,
    product_type: str,
    week_start: date,
    include_sat: bool,
    priority_time: str,
) -> list[dict]:
    """
    Distribute container IDs evenly across Mon–Sat (or Mon–Fri).
    Returns list of plan entry dicts ready to insert.
    """
    n_days   = 6 if include_sat else 5
    work_days = [week_start + timedelta(days=i) for i in range(n_days)]
    slot_time = priority_time if priority_time else "8:30 AM"
    slot_num  = dict(_PLAN_SLOTS).get(slot_time, 2)

    entries = []
    for idx, cid in enumerate(containers):
        day = work_days[idx % n_days]
        entries.append({
            "week_start":   week_start,
            "appt_date":    day,
            "appt_time":    slot_time,
            "slot_num":     slot_num,
            "container_id": cid.upper().strip(),
            "scac":         scac,
            "site_code":    site_code,
            "product_type": product_type,
            "qty":          1190,
            "notes":        "",
            "status":       "SCHEDULED",
        })
    return entries


# ── Excel export helpers ──────────────────────────────────────────────────────

def _xl_setup():
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    thin     = Side(style="thin")
    border   = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_ctr   = Alignment(horizontal="center", vertical="center")
    al_mid   = Alignment(vertical="center")
    return hdr_fill, hdr_font, border, al_ctr, al_mid

def _xl_carrier_fills():
    from openpyxl.styles import PatternFill
    return {
        "ATMI": PatternFill("solid", fgColor="FFF2CC"),
        "ARVY": PatternFill("solid", fgColor="DEEAF1"),
        "HDDR": PatternFill("solid", fgColor="E2EFDA"),
        "RKNE": PatternFill("solid", fgColor="F3E5F5"),
        "TGHE": PatternFill("solid", fgColor="FCE4EC"),
    }

def _xl_header(ws, cols):
    from openpyxl.utils import get_column_letter
    hdr_fill, hdr_font, border, al_ctr, _ = _xl_setup()
    ws.append(cols)
    for cell in ws[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.border = border; cell.alignment = al_ctr
    ws.row_dimensions[1].height = 18

def _xl_col_w(ws, widths):
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def _to_bytes(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def _export_compiled_excel(plan_df: pd.DataFrame, week_days: list, week_start: date) -> bytes:
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    fills = _xl_carrier_fills()
    _, _, border, _, al_mid = _xl_setup()
    COLS = ["#","Day","Site","Appt Time","Container #","Carrier","Product Type","Qty","Notes","Status"]
    W    = [4,12,8,10,16,8,14,8,28,12]

    # summary sheet first
    ws0 = wb.create_sheet("All Sites Summary")
    _xl_header(ws0, COLS)
    active = plan_df[plan_df["status"] != "PENDING"].sort_values(["appt_date","site_code","slot_num"])
    for n, r in enumerate(active.itertuples(), 1):
        try: day_s = date.fromisoformat(str(r.appt_date)).strftime("%A")
        except: day_s = str(r.appt_date or "TBD")
        ws0.append([n, day_s, r.site_code or "", r.appt_time,
                    r.container_id, r.carrier, r.product_type or "",
                    r.qty or "", r.notes or "", r.status])
        fill = fills.get(r.carrier)
        for cell in ws0[ws0.max_row]:
            cell.border = border; cell.alignment = al_mid
            if fill: cell.fill = fill
    _xl_col_w(ws0, W)

    # day sheets
    day_abbr = ["Mon","Tue","Wed","Thu","Fri","Sat"]
    for dd, abbr in zip(week_days, day_abbr):
        lbl = f"{abbr} {dd.month}-{dd.day}"
        ws = wb.create_sheet(lbl)
        _xl_header(ws, COLS)
        day_df = plan_df[(plan_df["appt_date"]==dd.isoformat()) & (plan_df["status"]!="PENDING")].sort_values(["site_code","slot_num"])
        for n, r in enumerate(day_df.itertuples(), 1):
            ws.append([n, dd.strftime("%A"), r.site_code or "", r.appt_time,
                       r.container_id, r.carrier, r.product_type or "",
                       r.qty or "", r.notes or "", r.status])
            fill = fills.get(r.carrier)
            for cell in ws[ws.max_row]:
                cell.border = border; cell.alignment = al_mid
                if fill: cell.fill = fill
        _xl_col_w(ws, W)

    return _to_bytes(wb)

def _export_site_excel(plan_df: pd.DataFrame, site_code: str, week_start: date) -> bytes:
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    fills = _xl_carrier_fills()
    _, _, border, _, al_mid = _xl_setup()
    from openpyxl.styles import Font as XlFont
    COLS = ["#","Day","Appt Time","Container #","Carrier","Product Type","Qty","Notes","Status"]
    W    = [4,12,10,16,8,14,8,28,12]
    site_df = plan_df[plan_df["site_code"]==site_code]

    ws = wb.create_sheet(f"{site_code} — All Carriers")
    _xl_header(ws, COLS)
    for n, r in enumerate(site_df[site_df["status"]!="PENDING"].sort_values(["appt_date","slot_num"]).itertuples(), 1):
        try: day_s = date.fromisoformat(str(r.appt_date)).strftime("%A")
        except: day_s = "TBD"
        ws.append([n, day_s, r.appt_time, r.container_id, r.carrier,
                   r.product_type or "", r.qty or "", r.notes or "", r.status])
        fill = fills.get(r.carrier)
        for cell in ws[ws.max_row]:
            cell.border = border; cell.alignment = al_mid
            if fill: cell.fill = fill
    _xl_col_w(ws, W)

    for scac in sorted(site_df["carrier"].dropna().unique()):
        ws2 = wb.create_sheet(f"{site_code} — {scac}")
        _xl_header(ws2, COLS)
        cdf = site_df[site_df["carrier"]==scac]
        for n, r in enumerate(cdf[cdf["status"]!="PENDING"].sort_values(["appt_date","slot_num"]).itertuples(), 1):
            try: day_s = date.fromisoformat(str(r.appt_date)).strftime("%A")
            except: day_s = "TBD"
            ws2.append([n, day_s, r.appt_time, r.container_id, r.carrier,
                        r.product_type or "", r.qty or "", r.notes or "", r.status])
            fill = fills.get(r.carrier)
            for cell in ws2[ws2.max_row]:
                cell.border = border; cell.alignment = al_mid
                if fill: cell.fill = fill
        _xl_col_w(ws2, W)

    return _to_bytes(wb)

def _export_carrier_excel(plan_df: pd.DataFrame, scac: str, carrier_name: str, week_start: date) -> bytes:
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    fills = _xl_carrier_fills(); fill = fills.get(scac)
    _, _, border, _, al_mid = _xl_setup()
    from openpyxl.styles import Font as XlFont
    COLS = ["#","Day","Site","Appt Time","Container #","Product Type","Qty","Notes"]
    W    = [4,12,8,10,16,14,8,28]
    c_df = plan_df[plan_df["carrier"]==scac]

    ws = wb.create_sheet(f"{carrier_name} — All Sites")
    _xl_header(ws, COLS)
    for n, r in enumerate(c_df[c_df["status"]!="PENDING"].sort_values(["appt_date","site_code","slot_num"]).itertuples(), 1):
        try: day_s = date.fromisoformat(str(r.appt_date)).strftime("%A")
        except: day_s = "TBD"
        ws.append([n, day_s, r.site_code or "", r.appt_time,
                   r.container_id, r.product_type or "", r.qty or "", r.notes or ""])
        for cell in ws[ws.max_row]:
            cell.border = border; cell.alignment = al_mid
            if fill: cell.fill = fill
    pend = c_df[c_df["status"]=="PENDING"]
    if not pend.empty:
        ws.append([])
        ws.append(["PENDING — confirm availability:"] + [""]*7)
        for cell in ws[ws.max_row]:
            cell.font = XlFont(bold=True, italic=True, color="7F7F7F")
        for r in pend.itertuples():
            ws.append([None,"TBD",r.site_code or "","TBD",r.container_id,r.product_type or "","",r.notes or ""])
            for cell in ws[ws.max_row]: cell.border = border
    _xl_col_w(ws, W)

    for site in sorted(c_df["site_code"].dropna().unique()):
        ws2 = wb.create_sheet(site)
        _xl_header(ws2, COLS)
        for n, r in enumerate(c_df[(c_df["site_code"]==site)&(c_df["status"]!="PENDING")].sort_values(["appt_date","slot_num"]).itertuples(), 1):
            try: day_s = date.fromisoformat(str(r.appt_date)).strftime("%A")
            except: day_s = "TBD"
            ws2.append([n, day_s, site, r.appt_time, r.container_id,
                        r.product_type or "", r.qty or "", r.notes or ""])
            for cell in ws2[ws2.max_row]:
                cell.border = border; cell.alignment = al_mid
                if fill: cell.fill = fill
        _xl_col_w(ws2, W)

    return _to_bytes(wb)


# ── shared UI helpers ─────────────────────────────────────────────────────────

def _plan_row_table(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    rows = []
    for r in df.itertuples():
        try:
            day_s = date.fromisoformat(str(r.appt_date)).strftime("%a %m/%d")
        except Exception:
            day_s = "TBD"
        all_cols = {
            "Day":         day_s if getattr(r,"status","") != "PENDING" else "TBD",
            "Time":        r.appt_time if getattr(r,"status","") != "PENDING" else "TBD",
            "Site":        getattr(r,"site_code","") or "",
            "Container #": r.container_id,
            "Carrier":     r.carrier,
            "Product":     r.product_type or "",
            "Qty":         r.qty or "",
            "Notes":       r.notes or "",
            "Status":      _STATUS_BADGE.get(r.status, r.status),
        }
        rows.append({c: all_cols[c] for c in cols if c in all_cols})
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=cols)

def _status_editor(df: pd.DataFrame, key_prefix: str):
    if df.empty:
        return
    _sites_e   = _get_sites_df()
    _carriers_e = _get_carriers_df()
    _site_opts  = _sites_e["site_code"].tolist() if not _sites_e.empty else []
    _scac_opts  = _carriers_e["scac"].tolist() if not _carriers_e.empty else list(_SCAC_COLOR.keys())

    ids    = df["id"].tolist()
    labels = [
        f"{r.container_id}  |  {getattr(r,'site_code','') or '?'}  |  {r.carrier}  |  "
        f"{r.appt_time}  ({_STATUS_BADGE.get(r.status, r.status)})"
        for r in df.itertuples()
    ]
    with st.expander("Edit Entry", expanded=False):
        sel    = st.selectbox("Select container", labels, key=f"{key_prefix}_sel")
        row_id = ids[labels.index(sel)]
        sel_row = df[df["id"] == row_id].iloc[0]

        ef1, ef2, ef3 = st.columns(3)
        with ef1:
            st.caption("**Schedule**")
            # date
            try:
                _cur_date = date.fromisoformat(str(sel_row["appt_date"])) if pd.notna(sel_row.get("appt_date")) else date.today()
            except Exception:
                _cur_date = date.today()
            new_date = st.date_input("Delivery date", value=_cur_date, key=f"{key_prefix}_date")
            # time
            _t_opts = [t for t, _ in _PLAN_SLOTS]
            _cur_time = str(sel_row.get("appt_time", "8:30 AM") or "8:30 AM")
            _t_idx = _t_opts.index(_cur_time) if _cur_time in _t_opts else 1
            new_time = st.selectbox("Appointment time", _t_opts, index=_t_idx, key=f"{key_prefix}_time")
            new_slot = dict(_PLAN_SLOTS)[new_time]

        with ef2:
            st.caption("**Assignment**")
            _cur_site = str(sel_row.get("site_code") or "")
            _s_idx = _site_opts.index(_cur_site) if _cur_site in _site_opts else 0
            new_site = st.selectbox("Site", _site_opts or ["RIC6"], index=_s_idx, key=f"{key_prefix}_site")
            _cur_scac = str(sel_row.get("carrier") or "")
            _c_idx = _scac_opts.index(_cur_scac) if _cur_scac in _scac_opts else 0
            new_scac = st.selectbox("Carrier", _scac_opts, index=_c_idx, key=f"{key_prefix}_carrier")
            _prod_opts = _PLAN_PRODUCTS
            _cur_prod = str(sel_row.get("product_type") or _prod_opts[0])
            _p_idx = _prod_opts.index(_cur_prod) if _cur_prod in _prod_opts else 0
            new_prod = st.selectbox("Product type", _prod_opts, index=_p_idx, key=f"{key_prefix}_prod")

        with ef3:
            st.caption("**Status & Notes**")
            _cur_stat = str(sel_row.get("status") or "SCHEDULED")
            _st_idx = _PLAN_STATUSES.index(_cur_stat) if _cur_stat in _PLAN_STATUSES else 0
            new_stat  = st.selectbox("Status", _PLAN_STATUSES, index=_st_idx, key=f"{key_prefix}_stat")
            _cur_qty  = int(sel_row.get("qty") or 1190)
            new_qty   = st.number_input("Qty", value=_cur_qty, step=10, min_value=0, key=f"{key_prefix}_qty")
            new_notes = st.text_input("Notes", value=str(sel_row.get("notes") or ""), key=f"{key_prefix}_notes")
            del_flag  = st.checkbox("Delete this entry", key=f"{key_prefix}_del")

        if st.button("Save changes", key=f"{key_prefix}_apply", type="primary"):
            if del_flag:
                _delete_plan_entry(row_id)
                st.success("Entry deleted.")
            else:
                _update_plan_entry(row_id,
                    appt_date=new_date.isoformat(),
                    appt_time=new_time,
                    slot_num=new_slot,
                    site_code=new_site,
                    carrier=new_scac,
                    product_type=new_prod,
                    status=new_stat,
                    qty=int(new_qty),
                    notes=new_notes.strip(),
                    week_start=_plan_week_start(new_date).isoformat(),
                )
                st.success(f"Saved — {sel_row['container_id']} updated.")
            st.rerun()

def _week_grid(plan_df: pd.DataFrame, week_days: list):
    day_abbrs = ["Mon","Tue","Wed","Thu","Fri","Sat"]
    gcols = st.columns(6)
    for ci, (gc, gd) in enumerate(zip(gcols, week_days)):
        day_iso = gd.isoformat()
        ent = plan_df[(plan_df["appt_date"]==day_iso)&(plan_df["status"]!="PENDING")].sort_values("slot_num")
        with gc:
            dc = len(ent)
            st.markdown(f"**{day_abbrs[ci]} {gd.month}/{gd.day}**")
            cap_c = "red" if dc >= 9 else "gray"
            st.markdown(f"<small style='color:{cap_c}'>{dc}/9</small>", unsafe_allow_html=True)
            if ent.empty:
                st.markdown("<small style='color:#ccc'>—</small>", unsafe_allow_html=True)
            else:
                for _, er in ent.iterrows():
                    clr  = _SCAC_COLOR.get(er["carrier"], "#888")
                    emj  = _STATUS_BADGE.get(er["status"], "")
                    site = er.get("site_code") or ""
                    st.markdown(
                        f"""<div style="background:{clr}22;border-left:3px solid {clr};
                            padding:4px 6px;margin:3px 0;border-radius:3px;font-size:11px;line-height:1.5">
                            <b>{er["appt_time"]}</b> {emj}<br/>
                            <code style="font-size:10px">{er["container_id"]}</code><br/>
                            <span style="color:{clr};font-weight:600">{er["carrier"]}</span>
                            {"<span style=\'color:#666\'> · " + site + "</span>" if site else ""}
                        </div>""",
                        unsafe_allow_html=True
                    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 BODY
# ══════════════════════════════════════════════════════════════════════════════
with tab7:
    st.subheader("Delivery Plan Scheduler")
    st.caption("🎯 **Purpose:** Build, manage, and distribute the weekly container delivery plan across all sites and carriers.")
    st.caption("📋 **How to use:** Select the week, add containers in Plan Builder, generate Slack messages from By Site, and export carrier-specific plans from Carrier View.")

    # ── global week selector ──────────────────────────────────────────────────
    _wc1, _wc2, _wc3, _wc4 = st.columns([3,1,1,1])
    with _wc1:
        _today = date.today()
        _wk_in = st.date_input("Week (any day)", value=_plan_week_start(_today), key="plan_wk_v2")
        _ws    = _plan_week_start(_wk_in)
        _wdays = [_ws + timedelta(days=i) for i in range(6)]
        st.caption(f"**{_ws.strftime('%b %d')} – {(_ws+timedelta(days=5)).strftime('%b %d, %Y')}**")

    _pdf = _get_plan(_ws.isoformat())
    _pdf_act = _pdf[_pdf["status"] != "PENDING"]

    with _wc2: st.metric("Total loads", len(_pdf_act))
    with _wc3: st.metric("Sites", _pdf_act["site_code"].nunique() if not _pdf_act.empty else 0)
    with _wc4: st.metric("Pending", len(_pdf[_pdf["status"]=="PENDING"]))


    # ── receiving day selector (persisted per week) ───────────────────────────
    _saved_days   = _get_week_config(_ws.isoformat())
    _day_abbr_map = {"Sun": 6, "Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3, "Fri": 4, "Sat": 5}

    st.markdown("**Receiving days this week:**")
    _day_cols_row = st.columns(9)
    _new_active = []
    for _di, _dname in enumerate(_ALL_DAYS):
        with _day_cols_row[_di]:
            _checked = st.checkbox(_dname, value=(_dname in _saved_days), key=f"daytoggle_{_dname}_{_ws.isoformat()}")
            if _checked:
                _new_active.append(_dname)

    # detect changes + save
    if set(_new_active) != set(_saved_days):
        # check for containers on newly-deactivated days
        _deactivated = [d for d in _saved_days if d not in _new_active]
        _conflicts = []
        if _deactivated and not _pdf.empty:
            for _dday in _deactivated:
                _didx   = _day_abbr_map.get(_dday)
                _day_dates = [_ws + timedelta(days=(_didx - _ws.weekday()) % 7)]
                for _dd2 in _day_dates:
                    _hits = _pdf[(_pdf["appt_date"]==_dd2.isoformat())&(_pdf["status"].isin(["SCHEDULED","HOLD"]))]
                    if not _hits.empty:
                        _conflicts.append((_dday, _dd2.strftime("%b %d"), len(_hits)))
        _save_week_config(_ws.isoformat(), _new_active)
        if _conflicts:
            for _cday, _cdate, _ccnt in _conflicts:
                st.warning(
                    f"**{_ccnt} container(s) still assigned to {_cday} {_cdate}** — "
                    f"go to Plan Builder → Edit to reassign, or use **Redistribute** below."
                )
        else:
            st.rerun()

    # redistribute prompt (shown when deactivated-day containers exist)
    _active_day_dates = []
    for _adn in _new_active:
        _adidx = _day_abbr_map.get(_adn, 0)
        _addate = _ws + timedelta(days=(_adidx - _ws.weekday()) % 7)
        if _addate >= _ws and _addate <= _ws + timedelta(days=6):
            _active_day_dates.append(_addate)

    _orphan_df = pd.DataFrame()
    if not _pdf.empty and _active_day_dates:
        _active_isos = {d.isoformat() for d in _active_day_dates}
        _orphan_df = _pdf[
            (~_pdf["appt_date"].isin(_active_isos)) &
            (_pdf["appt_date"].notna()) &
            (_pdf["status"].isin(["SCHEDULED","HOLD"]))
        ]

    if not _orphan_df.empty:
        st.error(f"**{len(_orphan_df)} container(s) assigned to inactive days.** Redistribute or manually reassign.")
        with st.expander("Redistribute to active days", expanded=False):
            st.dataframe(
                _plan_row_table(_orphan_df, ["Day","Site","Container #","Carrier","Product","Status"]),
                use_container_width=True, hide_index=True
            )
            if st.button("Redistribute evenly across active days", key="redistribute_btn", type="primary"):
                if _active_day_dates:
                    sorted_active = sorted(_active_day_dates)
                    for idx, row in enumerate(_orphan_df.itertuples()):
                        new_day = sorted_active[idx % len(sorted_active)]
                        _scm3 = _get_scm_df()
                        scm_r = _scm3[(_scm3["site_code"]==getattr(row,"site_code",""))&(_scm3["scac"]==row.carrier)]
                        def_t = scm_r["priority_time"].iloc[0] if not scm_r.empty and pd.notna(scm_r["priority_time"].iloc[0]) else "8:30 AM"
                        sn    = dict(_PLAN_SLOTS).get(def_t, 2)
                        _update_plan_entry(row.id,
                            appt_date=new_day.isoformat(),
                            appt_time=def_t,
                            slot_num=sn,
                            week_start=_ws.isoformat()
                        )
                    st.success(f"Redistributed {len(_orphan_df)} container(s) across {len(sorted_active)} active day(s).")
                    st.rerun()
                else:
                    st.error("No active days selected — select at least one day first.")

    st.divider()

    _tp1, _tp2, _tp3, _tp4, _tp5, _tp6, _tp7 = st.tabs([
        "Plan Builder", "All Sites", "By Site",
        "Carrier View", "WoW / History", "Config", "Import History",
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # PLAN BUILDER
    # ══════════════════════════════════════════════════════════════════════════
    with _tp1:
        _sites_df   = _get_sites_df()
        _cdf        = _get_carriers_df()
        _scm        = _get_scm_df()
        _act_sites  = _sites_df[_sites_df["active"]==1]["site_code"].tolist() if not _sites_df.empty else []
        _act_scacs  = _cdf[_cdf["active"]==1]["scac"].tolist() if not _cdf.empty else list(_SCAC_COLOR.keys())

        # ── Step 1: Parse stakeholder request ────────────────────────────────
        with st.expander("Step 1 — Paste Stakeholder Request", expanded=True):
            st.caption("Paste the message from Miguel or site ops. The tool will extract material needs and generate a plan structure.")
            _req_text = st.text_area(
                "Stakeholder request",
                height=160, key="req_text",
                placeholder='SHELF, HDTP: 52 containers (level loaded including saturday)\nSTRAP, HDTP: 5 containers (level loaded including saturday)',
            )
            _req_site = st.selectbox("Site this request is for", _act_sites or ["RIC6"], key="req_site")

            if st.button("Parse Request", key="req_parse", type="primary") and _req_text.strip():
                parsed = _parse_site_request(_req_text)
                if parsed:
                    st.session_state["parsed_items"] = parsed
                    st.session_state["parsed_site"]  = _req_site
                    st.success(f"Parsed {len(parsed)} material line(s).")
                else:
                    st.warning("No recognizable material lines found. Expected format: `MATERIAL TYPE: N containers (instructions)`")

            # show parsed items if available
            if st.session_state.get("parsed_items"):
                items = st.session_state["parsed_items"]
                parsed_site = st.session_state.get("parsed_site", _req_site)
                st.markdown("**Parsed demand:**")
                for i, item in enumerate(items):
                    col_a, col_b, col_c, col_d = st.columns([3,1,2,2])
                    with col_a:
                        st.markdown(f"**{item['material']}**")
                    with col_b:
                        st.markdown(f"{item['qty']} {item['unit']}(s)")
                    with col_c:
                        ll = "Level load" if item["level_load"] else "Manual"
                        sat = " + Sat" if item["include_sat"] else ""
                        st.caption(f"{ll}{sat}")
                    with col_d:
                        # look up how many already in plan for this material+site
                        _in_plan = len(_pdf[
                            (_pdf["site_code"]==parsed_site) &
                            (_pdf["product_type"].str.upper()==item["material"].upper())
                        ]) if not _pdf.empty else 0
                        st.caption(f"In plan: {_in_plan} / {item['qty']}")

        # ── Step 2: Add containers ────────────────────────────────────────────
        with st.expander("Step 2 — Add Container IDs", expanded=False):
            _add_tab_s, _add_tab_b = st.tabs(["Single", "Bulk Paste"])

            with _add_tab_s:
                # pre-fill from parsed request if available
                _parsed_items = st.session_state.get("parsed_items", [])
                _parsed_site  = st.session_state.get("parsed_site", _act_sites[0] if _act_sites else "RIC6")
                _mat_opts     = _PLAN_PRODUCTS
                _default_mat  = _parsed_items[0]["material"] if _parsed_items and _parsed_items[0]["material"] in _mat_opts else _mat_opts[0]

                sa1, sa2, sa3, sa4 = st.columns(4)
                with sa1:
                    _s_site   = st.selectbox("Site", _act_sites or ["RIC6"], index=(_act_sites.index(_parsed_site) if _parsed_site in _act_sites else 0), key="s_site")
                    _mapped   = _scm[_scm["site_code"]==_s_site]["scac"].tolist() if not _scm.empty else _act_scacs
                    _s_scac   = st.selectbox("Carrier", _mapped or _act_scacs, key="s_scac")
                    _s_status = st.selectbox("Status", ["SCHEDULED","PENDING","HOLD"], key="s_status")
                with sa2:
                    _s_cid     = st.text_input("Container #", placeholder="TCNU3773041", key="s_cid")
                    _s_product = st.selectbox("Product Type", _mat_opts, index=_mat_opts.index(_default_mat) if _default_mat in _mat_opts else 0, key="s_product")
                    _s_qty     = st.number_input("Qty (units)", value=1190, step=10, min_value=0, key="s_qty")
                with sa3:
                    _s_date = st.date_input("Delivery Date", value=_today, key="s_date")
                    _scm_r  = _scm[(_scm["site_code"]==_s_site)&(_scm["scac"]==_s_scac)]
                    _def_t  = _scm_r["priority_time"].iloc[0] if not _scm_r.empty and pd.notna(_scm_r["priority_time"].iloc[0]) else "8:30 AM"
                    _t_opts = [t for t,_ in _PLAN_SLOTS]
                    _s_time = st.selectbox("Appt Time", _t_opts, index=_t_opts.index(_def_t) if _def_t in _t_opts else 1, key="s_time")
                    _s_slot = dict(_PLAN_SLOTS)[_s_time]
                with sa4:
                    _s_notes = st.text_area("Notes", height=120, key="s_notes", placeholder="Optional")

                if st.button("Add", type="primary", key="s_add"):
                    if _s_cid.strip():
                        _ad = _s_date if _s_status != "PENDING" else None
                        _at = _s_time if _s_status != "PENDING" else "TBD"
                        _an = _s_slot if _s_status != "PENDING" else 99
                        _add_plan_entry(_plan_week_start(_s_date),_ad,_at,_an,
                                        _s_cid,_s_scac,_s_site,_s_product,int(_s_qty),_s_notes,_s_status)
                        st.success(f"{_s_cid.upper().strip()} ({_s_scac} → {_s_site})")
                        st.rerun()
                    else:
                        st.warning("Container # required.")

            with _add_tab_b:
                st.caption(
                    "**Format (one container per line):** `CONTAINER_ID  CARRIER  SITE`  \n"
                    "Carrier and Site are optional if you set defaults below.  \n"
                    "Example: `TCNU3773041  HDDR  RIC6`"
                )
                _bulk_text = st.text_area("Paste containers", height=200, key="bulk_text",
                                          placeholder="TCNU3773041  HDDR  RIC6\nMRKU4103422  ATMI  ILM1")
                bb1, bb2, bb3, bb4, bb5 = st.columns(5)
                with bb1: _b_site    = st.selectbox("Default Site",    _act_sites or ["RIC6"], key="b_site")
                with bb2: _b_scac    = st.selectbox("Default Carrier", _act_scacs, key="b_scac")
                with bb3: _b_product = st.selectbox("Default Product", _mat_opts, key="b_product")
                with bb4: _b_date    = st.date_input("Delivery Date",  value=_today, key="b_date")
                with bb5: _b_status  = st.selectbox("Status",          ["SCHEDULED","PENDING"], key="b_status")

                # level-load checkbox
                _b_ll     = st.checkbox("Level load across the week", value=True, key="b_ll")
                _b_incsat = st.checkbox("Include Saturday",            value=True, key="b_incsat")

                if st.button("Parse & Add All", type="primary", key="b_add"):
                    if _bulk_text.strip():
                        raw_cids, added, skipped = [], 0, []
                        for ln in _bulk_text.strip().splitlines():
                            parts = ln.strip().split()
                            if not parts: continue
                            cid   = parts[0].upper()
                            if len(cid) < 4: skipped.append(ln); continue
                            b_sc  = parts[1].upper() if len(parts) > 1 and parts[1].upper() in _act_scacs else _b_scac
                            b_si  = parts[2].upper() if len(parts) > 2 else _b_site
                            raw_cids.append((cid, b_sc, b_si))

                        if _b_ll and raw_cids:
                            # use the configured active receiving days for this week
                            _ll_active = _get_week_config(_plan_week_start(_b_date).isoformat())
                            _ll_day_map = {"Sun":6,"Mon":0,"Tue":1,"Wed":2,"Thu":3,"Fri":4,"Sat":5}
                            _ll_ws = _plan_week_start(_b_date)
                            work_days = sorted([
                                _ll_ws + timedelta(days=(_ll_day_map[d]-_ll_ws.weekday())%7)
                                for d in _ll_active if d in _ll_day_map
                            ])
                            if not work_days:
                                work_days = [_plan_week_start(_b_date) + timedelta(days=i) for i in range(5)]
                            for idx, (cid, b_sc, b_si) in enumerate(raw_cids):
                                day = work_days[idx % n_days]
                                _scm_r2 = _scm[(_scm["site_code"]==b_si)&(_scm["scac"]==b_sc)]
                                _def_t2 = _scm_r2["priority_time"].iloc[0] if not _scm_r2.empty and pd.notna(_scm_r2["priority_time"].iloc[0]) else "8:30 AM"
                                _sn2    = dict(_PLAN_SLOTS).get(_def_t2, 2)
                                _add_plan_entry(_plan_week_start(day), day, _def_t2, _sn2,
                                                cid, b_sc, b_si, _b_product, 1190, "", _b_status)
                                added += 1
                        else:
                            for cid, b_sc, b_si in raw_cids:
                                _scm_r2 = _scm[(_scm["site_code"]==b_si)&(_scm["scac"]==b_sc)]
                                _def_t2 = _scm_r2["priority_time"].iloc[0] if not _scm_r2.empty and pd.notna(_scm_r2["priority_time"].iloc[0]) else "8:30 AM"
                                _sn2    = dict(_PLAN_SLOTS).get(_def_t2, 2)
                                _ad2    = _b_date if _b_status != "PENDING" else None
                                _at2    = _def_t2 if _b_status != "PENDING" else "TBD"
                                _add_plan_entry(_plan_week_start(_b_date), _ad2, _at2,
                                                _sn2 if _b_status != "PENDING" else 99,
                                                cid, b_sc, b_si, _b_product, 1190, "", _b_status)
                                added += 1

                        msg = f"Added {added} container(s)."
                        if skipped: msg += f" Skipped {len(skipped)} line(s)."
                        st.success(msg)
                        st.rerun()
                    else:
                        st.warning("Nothing to parse.")

        # ── Week grid ─────────────────────────────────────────────────────────
        _pdf = _get_plan(_ws.isoformat())
        if _pdf.empty:
            st.info("No entries yet. Use the steps above to build the plan.")
        else:
            st.divider()
            _week_grid(_pdf[_pdf["status"]!="PENDING"], _wdays)
            st.divider()
            _status_editor(_pdf, "pb")


    # ══════════════════════════════════════════════════════════════════════════
    # ALL SITES (COMPILED)
    # ══════════════════════════════════════════════════════════════════════════
    with _tp2:
        _pdf = _get_plan(_ws.isoformat())
        if _pdf.empty:
            st.info("No plan entries for this week.")
        else:
            _act = _pdf[_pdf["status"]!="PENDING"]
            # per-site metrics
            _sc_cnt = _act.groupby("site_code")["id"].count() if not _act.empty else pd.Series(dtype=int)
            _mc = st.columns(max(len(_sc_cnt), 1))
            for mi, (sc, cnt) in enumerate(sorted(_sc_cnt.items())):
                with _mc[mi % len(_mc)]: st.metric(sc, int(cnt))
            st.divider()
            COLS = ["Day","Time","Site","Container #","Carrier","Product","Qty","Notes","Status"]
            st.dataframe(_plan_row_table(_act.sort_values(["appt_date","site_code","slot_num"]), COLS),
                         use_container_width=True, hide_index=True)
            _pend = _pdf[_pdf["status"]=="PENDING"]
            if not _pend.empty:
                st.caption(f"**{len(_pend)} pending (unscheduled)**")
                st.dataframe(_plan_row_table(_pend, ["Site","Container #","Carrier","Product","Qty","Notes"]),
                             use_container_width=True, hide_index=True)
            xl = _export_compiled_excel(_pdf, _wdays, _ws)
            st.download_button("Export — All Sites", data=xl,
                file_name=f"Delivery Plan All Sites {_ws.strftime('%m.%d.%y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary")

    # ══════════════════════════════════════════════════════════════════════════
    # BY SITE (INTERNAL)
    # ══════════════════════════════════════════════════════════════════════════
    with _tp3:
        _pdf = _get_plan(_ws.isoformat())
        _avail_sites = sorted(_pdf["site_code"].dropna().unique().tolist()) if not _pdf.empty else []
        if not _avail_sites:
            st.info("No entries for this week.")
        else:
            _sel_site = st.selectbox("Site", _avail_sites, key="site_sel")
            _sdf  = _pdf[_pdf["site_code"]==_sel_site]
            _sact = _sdf[_sdf["status"]!="PENDING"]
            # carrier breakdown badges
            _sc_c = _sact.groupby("carrier")["id"].count() if not _sact.empty else pd.Series(dtype=int)
            _bcs  = st.columns(max(len(_sc_c), 1))
            for bi, (sc, cnt) in enumerate(sorted(_sc_c.items())):
                with _bcs[bi]:
                    clr = _SCAC_COLOR.get(sc, "#888")
                    st.markdown(
                        f"<div style='border-left:4px solid {clr};padding:6px 10px'>"
                        f"<b style='color:{clr}'>{sc}</b><br/><span style='font-size:22px'>{cnt}</span></div>",
                        unsafe_allow_html=True)
            st.divider()
            COLS = ["Day","Time","Container #","Carrier","Product","Qty","Notes","Status"]
            st.dataframe(_plan_row_table(_sact.sort_values(["appt_date","slot_num"]), COLS),
                         use_container_width=True, hide_index=True)
            _sp = _sdf[_sdf["status"]=="PENDING"]
            if not _sp.empty:
                st.caption(f"**{len(_sp)} pending**")
                st.dataframe(_plan_row_table(_sp, ["Container #","Carrier","Product","Qty","Notes"]),
                             use_container_width=True, hide_index=True)
            # ── Daily notification ────────────────────────────────────────────────
            with st.expander("Daily Notification — Copy for Slack", expanded=False):
                st.caption("Generates tomorrow's delivery list for this site, ready to paste into Slack.")
                _notif_date = st.date_input(
                    "Delivery date to notify",
                    value=_today + timedelta(days=1),
                    key=f"notif_date_{_sel_site}"
                )
                _notif_df = _pdf[
                    (_pdf["site_code"] == _sel_site) &
                    (_pdf["appt_date"] == _notif_date.isoformat()) &
                    (_pdf["status"].isin(["SCHEDULED", "HOLD"]))
                ].sort_values("slot_num")

                if _notif_df.empty:
                    st.info(f"No scheduled deliveries for {_sel_site} on {_notif_date.strftime('%b %d')}.")
                else:
                    _day_label = _notif_date.strftime("%A, %B %d")
                    lines = [f"*{_sel_site} — Scheduled Deliveries for {_day_label}*", ""]
                    for r in _notif_df.itertuples():
                        _cn = _get_carriers_df()
                        _cname_row = _cn[_cn["scac"] == r.carrier]
                        _cname_n   = _cname_row["carrier_name"].iloc[0] if not _cname_row.empty else r.carrier
                        lines.append(f"{r.appt_time}  |  {r.container_id}  |  {_cname_n}  |  {r.product_type or ''}")
                    lines.append("")
                    lines.append(f"Total: {len(_notif_df)} container(s)")
                    notif_text = "\n".join(lines)
                    st.text_area("Copy this", value=notif_text, height=200, key=f"notif_txt_{_sel_site}")

            _status_editor(_sdf, f"site_{_sel_site}")
            if not _sdf.empty:
                xl = _export_site_excel(_pdf, _sel_site, _ws)
                st.download_button(f"Export {_sel_site} to Excel", data=xl,
                    file_name=f"{_sel_site} Delivery Plan {_ws.strftime('%m.%d.%y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ══════════════════════════════════════════════════════════════════════════
    # CARRIER VIEW (EXTERNAL / SHAREABLE)
    # ══════════════════════════════════════════════════════════════════════════
    with _tp4:
        _pdf = _get_plan(_ws.isoformat())
        _cdf2 = _get_carriers_df()
        _avail_scacs = sorted(_pdf["carrier"].dropna().unique().tolist()) if not _pdf.empty else []
        if not _avail_scacs:
            st.info("No entries for this week.")
        else:
            _sel_scac = st.selectbox("Carrier", _avail_scacs, key="carr_sel")
            _cn_r = _cdf2[_cdf2["scac"]==_sel_scac]
            _cname = _cn_r["carrier_name"].iloc[0] if not _cn_r.empty else _sel_scac
            _ccontact = _cn_r["ops_contact"].iloc[0] if not _cn_r.empty else ""
            clr_cv = _SCAC_COLOR.get(_sel_scac, "#888")
            st.markdown(
                f"<div style='background:{clr_cv}15;border-left:4px solid {clr_cv};"
                f"padding:10px 14px;border-radius:4px;margin-bottom:8px'>"
                f"<b style='font-size:16px'>{_cname} ({_sel_scac})</b>"
                f"{'<br/><small>Ops: ' + _ccontact + '</small>' if _ccontact else ''}"
                f"</div>", unsafe_allow_html=True)
            _cdf_w = _pdf[_pdf["carrier"]==_sel_scac]
            _cact  = _cdf_w[_cdf_w["status"]!="PENDING"]
            # site breakdown
            _csite_c = _cact.groupby("site_code")["id"].count() if not _cact.empty else pd.Series(dtype=int)
            cmc = st.columns(max(len(_csite_c), 1))
            for cmi, (cs, cc) in enumerate(sorted(_csite_c.items())):
                with cmc[cmi]: st.metric(cs, int(cc))
            st.divider()
            st.caption("**Shareable view** — no internal status labels.")
            ext_rows = []
            for r in _cact.sort_values(["appt_date","site_code","slot_num"]).itertuples():
                try: day_s = date.fromisoformat(str(r.appt_date)).strftime("%a %m/%d")
                except: day_s = "TBD"
                ext_rows.append({"Day": day_s, "Site": r.site_code or "",
                                  "Appt Time": r.appt_time, "Container #": r.container_id,
                                  "Product Type": r.product_type or "", "Qty": r.qty or "",
                                  "Notes": r.notes or ""})
            if ext_rows:
                st.dataframe(pd.DataFrame(ext_rows), use_container_width=True, hide_index=True)
            _cpend = _cdf_w[_cdf_w["status"]=="PENDING"]
            if not _cpend.empty:
                st.caption(f"**{len(_cpend)} pending**")
                pext = [{"Site": r.site_code or "", "Container #": r.container_id,
                          "Product Type": r.product_type or "", "Notes": r.notes or ""}
                         for r in _cpend.itertuples()]
                st.dataframe(pd.DataFrame(pext), use_container_width=True, hide_index=True)
            if not _cdf_w.empty:
                xl = _export_carrier_excel(_pdf, _sel_scac, _cname, _ws)
                st.download_button(f"Export {_cname} Plan to Excel", data=xl,
                    file_name=f"{_sel_scac} Delivery Plan {_ws.strftime('%m.%d.%y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary")

    # ══════════════════════════════════════════════════════════════════════════
    # WoW / HISTORY
    # ══════════════════════════════════════════════════════════════════════════
    with _tp5:
        _wow_t, _hist_t = st.tabs(["Week-over-Week", "History Search"])

        with _wow_t:
            _pdf      = _get_plan(_ws.isoformat())
            _prior    = _get_plan((_ws - timedelta(weeks=1)).isoformat())
            cur_ids   = set(_pdf["container_id"].str.upper()) if not _pdf.empty else set()
            prior_ids = set(_prior["container_id"].str.upper()) if not _prior.empty else set()
            new_ids     = cur_ids - prior_ids
            rolled_ids  = cur_ids & prior_ids
            dropped_ids = prior_ids - cur_ids
            w1, w2, w3 = st.columns(3)
            with w1: st.metric("New this week",         len(new_ids))
            with w2: st.metric("Rolled from prior week", len(rolled_ids))
            with w3: st.metric("Dropped / delivered",    len(dropped_ids))
            st.divider()
            if new_ids:
                st.markdown("**New containers:**")
                st.dataframe(_plan_row_table(_pdf[_pdf["container_id"].str.upper().isin(new_ids)],
                    ["Day","Site","Container #","Carrier","Product","Status"]),
                    use_container_width=True, hide_index=True)
            if rolled_ids:
                st.markdown("**Carried over from prior week:**")
                roll_rows = []
                for r in _pdf[_pdf["container_id"].str.upper().isin(rolled_ids)].itertuples():
                    pr = _prior[_prior["container_id"].str.upper()==r.container_id.upper()]
                    pr_stat = pr["status"].iloc[0] if not pr.empty else "?"
                    try: day_s = date.fromisoformat(str(r.appt_date)).strftime("%a %m/%d")
                    except: day_s = "TBD"
                    roll_rows.append({
                        "Container #": r.container_id, "Site": getattr(r,"site_code","") or "",
                        "Carrier": r.carrier,
                        "This Week": f"{day_s} {_STATUS_BADGE.get(r.status,'')}",
                        "Prior Week Status": _STATUS_BADGE.get(pr_stat, pr_stat),
                    })
                st.dataframe(pd.DataFrame(roll_rows), use_container_width=True, hide_index=True)
            if dropped_ids:
                prior_lbl = (_ws - timedelta(weeks=1)).strftime("%b %d")
                st.markdown(f"**Not on this week (were on {prior_lbl}):**")
                st.dataframe(_plan_row_table(_prior[_prior["container_id"].str.upper().isin(dropped_ids)],
                    ["Day","Site","Container #","Carrier","Status"]),
                    use_container_width=True, hide_index=True)

        with _hist_t:
            st.caption("Search all stored plans.")
            hc1, hc2, hc3 = st.columns([2,1,1])
            with hc1: _hq  = st.text_input("Container # or keyword", key="hq")
            with hc2: _hsc = st.selectbox("Carrier", ["All"]+list(_SCAC_COLOR.keys()), key="hsc")
            with hc3: _hsi = st.text_input("Site", placeholder="RIC6", key="hsi")
            _hdr = st.date_input("Date range", value=[date(2026,1,1), _today], key="hdr")
            if st.button("Search", key="hsearch"):
                conn = get_db()
                q = "SELECT * FROM delivery_plan WHERE 1=1"
                p = []
                if _hq.strip():
                    q += " AND (container_id LIKE ? OR notes LIKE ?)"; p += [f"%{_hq.strip().upper()}%", f"%{_hq.strip()}%"]
                if _hsc != "All":
                    q += " AND carrier=?"; p.append(_hsc)
                if _hsi.strip():
                    q += " AND site_code LIKE ?"; p.append(f"%{_hsi.strip().upper()}%")
                if len(_hdr) == 2:
                    q += " AND (appt_date >= ? AND appt_date <= ?)"; p += [str(_hdr[0]), str(_hdr[1])]
                q += " ORDER BY appt_date DESC LIMIT 300"
                hdf = pd.read_sql_query(q, conn, params=p); conn.close()
                if hdf.empty:
                    st.info("No results.")
                else:
                    st.caption(f"**{len(hdf)} result(s)**")
                    st.dataframe(_plan_row_table(hdf, ["Day","Site","Container #","Carrier","Product","Status","Notes"]),
                                 use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # CONFIG
    # ══════════════════════════════════════════════════════════════════════════
    with _tp6:
        _cfg1, _cfg2, _cfg3 = st.tabs(["Sites", "Carriers", "Site–Carrier Map"])

        with _cfg1:
            _sdf3 = _get_sites_df()
            st.dataframe(_sdf3.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)
            with st.expander("Add / Update Site"):
                cx1, cx2, cx3 = st.columns(3)
                with cx1:
                    _csc = st.text_input("Site Code", placeholder="RIC6", key="cfg_sc")
                    _cfn = st.text_input("FC Name",   placeholder="Richmond FC", key="cfg_fn")
                    _cpo = st.text_input("Port",      placeholder="ORF", key="cfg_po")
                with cx2:
                    _crg = st.text_input("Region", placeholder="East", key="cfg_rg")
                    _ccp = st.number_input("Capacity/Day", value=9, min_value=1, key="cfg_cp")
                    _cac = st.checkbox("Active", value=True, key="cfg_ac")
                with cx3:
                    _cno = st.text_area("Notes", height=100, key="cfg_no")
                if st.button("Save Site", key="cfg_ss"):
                    if _csc.strip():
                        _config_write(
                            """INSERT INTO plan_sites (site_code,fc_name,port,region,capacity_day,active,notes)
                               VALUES (?,?,?,?,?,?,?)
                               ON CONFLICT(site_code) DO UPDATE SET fc_name=excluded.fc_name,port=excluded.port,
                               region=excluded.region,capacity_day=excluded.capacity_day,
                               active=excluded.active,notes=excluded.notes""",
                            (_csc.strip().upper(),_cfn.strip(),_cpo.strip().upper(),
                             _crg.strip(),int(_ccp),int(_cac),_cno.strip()))
                        st.success(f"Saved {_csc.upper()}."); st.rerun()
                    else: st.warning("Site code required.")

        with _cfg2:
            _cdf3 = _get_carriers_df()
            st.dataframe(_cdf3.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)
            with st.expander("Add / Update Carrier"):
                cy1, cy2 = st.columns(2)
                with cy1:
                    _cc_scac = st.text_input("SCAC", key="cfg_cscac")
                    _cc_name = st.text_input("Carrier Name", key="cfg_cname")
                with cy2:
                    _cc_ctc  = st.text_input("Ops Contact", key="cfg_cctc")
                    _cc_eml  = st.text_input("Ops Email",   key="cfg_ceml")
                _cc_act = st.checkbox("Active", value=True, key="cfg_cact")
                if st.button("Save Carrier", key="cfg_sc2"):
                    if _cc_scac.strip():
                        _config_write(
                            """INSERT INTO plan_carriers (scac,carrier_name,ops_contact,ops_email,active)
                               VALUES (?,?,?,?,?)
                               ON CONFLICT(scac) DO UPDATE SET carrier_name=excluded.carrier_name,
                               ops_contact=excluded.ops_contact,ops_email=excluded.ops_email,active=excluded.active""",
                            (_cc_scac.strip().upper(),_cc_name.strip(),_cc_ctc.strip(),_cc_eml.strip(),int(_cc_act)))
                        st.success(f"Saved {_cc_scac.upper()}."); st.rerun()
                    else: st.warning("SCAC required.")

        with _cfg3:
            _scm2 = _get_scm_df()
            if not _scm2.empty:
                dc = ["site_code","scac","carrier_name","site_contact","site_contact_email",
                      "priority_time","allocation_pct","active","notes"]
                st.dataframe(_scm2[[c for c in dc if c in _scm2.columns]], use_container_width=True, hide_index=True)
            with st.expander("Add / Update Mapping"):
                _sdf4 = _get_sites_df(); _cdf4 = _get_carriers_df()
                cm1, cm2, cm3 = st.columns(3)
                with cm1:
                    _cm_site  = st.selectbox("Site",    _sdf4["site_code"].tolist() if not _sdf4.empty else [], key="cm_s")
                    _cm_scac  = st.selectbox("Carrier", _cdf4["scac"].tolist() if not _cdf4.empty else [], key="cm_c")
                    _cm_alloc = st.number_input("Allocation %", value=0.0, step=5.0, key="cm_a")
                with cm2:
                    _cm_ctc  = st.text_input("Site Contact Name",  key="cm_ctc")
                    _cm_eml  = st.text_input("Site Contact Email", key="cm_eml")
                    _cm_pt   = st.selectbox("Priority Time Slot", [t for t,_ in _PLAN_SLOTS], key="cm_pt")
                with cm3:
                    _cm_act  = st.checkbox("Active", value=True, key="cm_act")
                    _cm_notes= st.text_area("Notes", height=80, key="cm_no")
                if st.button("Save Mapping", key="cm_save"):
                    _config_write(
                        """INSERT INTO plan_site_carrier
                           (site_code,scac,site_contact,site_contact_email,priority_time,allocation_pct,active,notes)
                           VALUES (?,?,?,?,?,?,?,?)
                           ON CONFLICT(site_code,scac) DO UPDATE SET site_contact=excluded.site_contact,
                           site_contact_email=excluded.site_contact_email,priority_time=excluded.priority_time,
                           allocation_pct=excluded.allocation_pct,active=excluded.active,notes=excluded.notes""",
                        (_cm_site,_cm_scac,_cm_ctc.strip(),_cm_eml.strip(),
                         _cm_pt,float(_cm_alloc),int(_cm_act),_cm_notes.strip()))
                    st.success(f"Saved {_cm_site} ↔ {_cm_scac}."); st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # IMPORT HISTORY
    # ══════════════════════════════════════════════════════════════════════════
    with _tp7:
        st.subheader("Import Historical Delivery Data")
        st.caption(
            "Upload the **ToteASERs Robotics DBR Tracker.xlsx** (downloaded from SharePoint) "
            "to load container history. Run weekly before planning to keep WoW comparisons current. "
            "Containers already in the database are skipped automatically."
        )

        _up = st.file_uploader(
            "ToteASERs Robotics DBR Tracker.xlsx",
            type=["xlsx"],
            key="imp_dbr_upload",
            help="Download the file from SharePoint, then upload it here.",
        )
        _imp_bytes = _up.read() if _up else None

        if _imp_bytes:
            _imp_rows, _imp_errors = _parse_dbr_tracker(_imp_bytes)

            if not _imp_rows:
                st.error("No rows parsed — confirm this is the DBR Tracker file.")
                for _e in _imp_errors[:5]:
                    st.warning(_e)
            else:
                # Deduplication check
                _con2 = get_db()
                _existing_ids = set(
                    r[0] for r in _con2.execute(
                        "SELECT DISTINCT container_id FROM delivery_plan"
                    ).fetchall()
                )
                _con2.close()
                _new_rows  = [r for r in _imp_rows if r["container_id"] not in _existing_ids]
                _dup_count = len(_imp_rows) - len(_new_rows)

                _im1, _im2, _im3, _im4 = st.columns(4)
                _im1.metric("In file",          len(_imp_rows))
                _im2.metric("New (to import)",   len(_new_rows))
                _im3.metric("Already in DB",     _dup_count)
                _im4.metric("Parse warnings",    len(_imp_errors))

                if _imp_errors:
                    with st.expander(f"Parse warnings ({len(_imp_errors)})"):
                        for _e in _imp_errors[:30]:
                            st.caption(_e)

                if _new_rows:
                    _prev_df = pd.DataFrame(_new_rows)

                    _ps1, _ps2, _ps3 = st.columns(3)
                    with _ps1:
                        st.caption("By Site")
                        st.dataframe(
                            _prev_df.groupby("site_code").size().reset_index(name="containers"),
                            hide_index=True, use_container_width=True,
                        )
                    with _ps2:
                        st.caption("By Carrier")
                        st.dataframe(
                            _prev_df.groupby("carrier").size().reset_index(name="containers"),
                            hide_index=True, use_container_width=True,
                        )
                    with _ps3:
                        st.caption("By Status")
                        st.dataframe(
                            _prev_df.groupby("status").size().reset_index(name="containers"),
                            hide_index=True, use_container_width=True,
                        )

                    _dr = sorted(_prev_df["appt_date"].dropna().tolist())
                    if _dr:
                        st.caption(f"Date range: {_dr[0]} through {_dr[-1]}")

                    with st.expander("Preview first 25 rows"):
                        st.dataframe(
                            _prev_df[[
                                "container_id","carrier","site_code","appt_date",
                                "product_type","qty","status","notes",
                            ]].head(25),
                            hide_index=True, use_container_width=True,
                        )

                    if st.button(
                        f"Import {len(_new_rows)} containers", type="primary", key="imp_go"
                    ):
                        _now = datetime.now().isoformat()
                        _icon = get_db()
                        _icon.executemany(
                            """INSERT INTO delivery_plan
                               (week_start,appt_date,appt_time,slot_num,container_id,carrier,
                                site_code,product_type,qty,notes,status,created_at,updated_at)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            [(r["week_start"],r["appt_date"],r["appt_time"],r["slot_num"],
                              r["container_id"],r["carrier"],r["site_code"],r["product_type"],
                              r["qty"],r["notes"],r["status"],_now,_now)
                             for r in _new_rows],
                        )
                        _icon.commit(); _icon.close()
                        if S3_ENABLED:
                            data_sync.push_db_to_s3(AWS_KEY, AWS_SECRET, AWS_REGION, S3_BUCKET)
                        st.success(
                            f"Imported {len(_new_rows)} containers. "
                            "WoW / History tab now has full historical data."
                        )
                        st.rerun()
                else:
                    st.info("All containers in this file are already in the database — nothing new to import.")

