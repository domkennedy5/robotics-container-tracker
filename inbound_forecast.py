"""
inbound_forecast.py
Inbound Forecast & Carrier Allocation tab for the Robotics Container Tracker.
Panels: Pipeline Forecast | Allocation Manager | Simulator
"""
from __future__ import annotations
import sqlite3
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

import streamlit as st
import pandas as pd
import plotly.express as px

_PHX = ZoneInfo("America/Phoenix")

# ─────────────────────────────────────────────────────────────────────────────
# DB INIT
# ─────────────────────────────────────────────────────────────────────────────

def init_inbound_forecast_db(conn: sqlite3.Connection) -> None:
    """Create tables and migrate plan_sites for site_type support."""
    for col_sql in [
        "ALTER TABLE plan_sites ADD COLUMN site_type TEXT DEFAULT 'Static'",
        "ALTER TABLE plan_sites ADD COLUMN site_type_effective TEXT",
        "ALTER TABLE plan_sites ADD COLUMN recv_days_per_week INTEGER DEFAULT 4",
    ]:
        try:
            conn.execute(col_sql)
        except sqlite3.OperationalError:
            pass

    conn.execute("""
        CREATE TABLE IF NOT EXISTS ar_inbound_unified (
            id                     INTEGER PRIMARY KEY AUTOINCREMENT,
            report_date            TEXT,
            container_id           TEXT,
            po_number              TEXT,
            po_line_number         INTEGER,
            po_shipment_number     INTEGER,
            part_number            TEXT,
            item_description       TEXT,
            supplier               TEXT,
            carrier_scac           TEXT,
            a_code                 TEXT,
            fc_identifier          TEXT,
            warehouse_operator     TEXT,
            warehouse_address      TEXT,
            node_type              TEXT,
            destination_ocean_port TEXT,
            origin_ocean_port      TEXT,
            discharged_port        TEXT,
            vessel_name            TEXT,
            container_size         TEXT,
            container_status       TEXT,
            market                 TEXT,
            departure_date         TEXT,
            ocean_eta              TEXT,
            actual_arrival_port    TEXT,
            off_vessel_date        TEXT,
            ready_date             TEXT,
            lfd_demurrage          TEXT,
            lfd_per_diem           TEXT,
            actual_pickup          TEXT,
            enter_facility         TEXT,
            container_empty        TEXT,
            return_to_port         TEXT,
            po_promised_date       TEXT,
            eta_final_destination  TEXT,
            ordered_units          INTEGER,
            shipped_cartons        INTEGER,
            shipped_weight_kg      REAL,
            shipped_volume_cbm     REAL,
            units_per_carton       REAL,
            shipped_units          INTEGER,
            data_source            TEXT,
            loads_refresh_date     TEXT,
            iss_refresh_date       TEXT,
            uploaded_at            TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ar_site_lookup (
            a_code                TEXT PRIMARY KEY,
            fc_identifier         TEXT,
            warehouse_operator    TEXT,
            warehouse_address     TEXT,
            node_type             TEXT,
            destination_ocean_port TEXT,
            updated_at            TEXT
        )
    """)
    conn.commit()


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN NORMALIZER
# ─────────────────────────────────────────────────────────────────────────────

_SPEC_COLS = [
    "report_date","container_id","po_number","po_line_number","po_shipment_number",
    "part_number","item_description","supplier","carrier_scac",
    "a_code","fc_identifier","warehouse_operator","warehouse_address","node_type",
    "destination_ocean_port","origin_ocean_port","discharged_port","vessel_name",
    "container_size","container_status","market",
    "departure_date","ocean_eta","actual_arrival_port","off_vessel_date",
    "ready_date","lfd_demurrage","lfd_per_diem","actual_pickup",
    "enter_facility","container_empty","return_to_port",
    "po_promised_date","eta_final_destination",
    "ordered_units","shipped_cartons","shipped_weight_kg","shipped_volume_cbm",
    "units_per_carton","shipped_units","data_source","loads_refresh_date","iss_refresh_date",
]

# ── Source-specific column alias maps ────────────────────────────────────────

# Loads: Oracle/AR Inbound Loads Report (daily, Sun-Sat)
_LOADS_ALIASES = {
    "container_id":         ["container id","container_id","container no","container_no","container"],
    "po_number":            ["po number","po_number","po #","purchase order number","po"],
    "po_line_number":       ["po line number","po_line_number","line number","line #"],
    "po_shipment_number":   ["po shipment number","po_shipment_number","shipment number","shipment #"],
    "part_number":          ["part number","part_number","part no","part#","item number"],
    "item_description":     ["po line item description","item description","description",
                             "line item description","part description"],
    "supplier":             ["supplier","vendor","vendor name"],
    "a_code":               ["destination fc","destination_fc","dest fc","ar org number","org number"],
    "origin_ocean_port":    ["international origin","origin port","origin","intl origin",
                             "port of origin","origin_port"],
    "vessel_name":          ["vessel name","vessel_name","vessel"],
    "container_size":       ["container size","container_size","size","iso size"],
    "shipped_weight_kg":    ["weight","weight (kg)","weight_kg","shipped weight",
                             "shipped weight (kg)","container weight"],
    "departure_date":       ["departure date at orign","departure date at origin",
                             "departure date","origin departure","depart date"],
    "actual_arrival_port":  ["actual arrival at discharge port","actual arrival at port",
                             "actual arrival","arrival date","arrived port"],
    "po_promised_date":     ["po promised date","promised date","need by date","need-by date"],
    "eta_final_destination":["eta to final destination","eta final destination",
                             "final eta","eta final dest","eta_final_destination"],
    "ordered_units":        ["quantity","qty","ordered units","ordered qty","order qty"],
}

# ISS: Import Shipment Status / Inbound by PO (Mon-Fri)
_ISS_ALIASES = {
    "container_id":      ["container","container_id","container id","container no","container_no"],
    "_po_iss":           ["po","po number","purchase order"],   # raw "1102935-16-1" format
    "carrier_scac":      ["full carrier","ssl","carrier","scac","dray scac",
                          "carrier_scac","carrier code"],
    "market":            ["market","region"],
    "discharged_port":   ["discharged port","discharge port","port","harbor"],
    "container_status":  ["container status","status","ctr status"],
    "_site_facility":    ["facility","fc","destination","facility code"],  # "A320-RBTCS"
    "shipped_volume_cbm":["cbm","volume cbm","volume (cbm)","shipped volume cbm","volume"],
    "shipped_cartons":   ["cartoncount","carton count","cartons","shipped cartons","total cartons"],
    "ocean_eta":         ["ocean eta","eta","vessel eta","ocean_eta"],
    "off_vessel_date":   ["off vessel","off_vessel","vessel discharge date","discharged date"],
    "ready_date":        ["ready date/time","ready date","ready_date","ready datetime"],
    "lfd_demurrage":     ["lfd demurrage","lfd_demurrage","demurrage lfd"],
    "lfd_per_diem":      ["lfd per diem","lfd_per_diem","per diem lfd","detention lfd"],
    "actual_pickup":     ["actual pickup at port","actual pickup","pickup date","picked up"],
    "enter_facility":    ["enter facility","enter_facility","facility arrival","arrived facility"],
    "container_empty":   ["container empty","container_empty","empty date","emptied"],
    "return_to_port":    ["return to port","return_to_port","returned to port","empty return"],
}

# Robotics Lanes (quarterly site reference)
_LANES_ALIASES = {
    "a_code":                  ["destination node code","a-code","a_code","node code",
                                 "dest node code","destination code"],
    "_full_address":           ["destination final node address","destination address",
                                "address","facility address","node address"],
    "node_type":               ["node type","node_type","type"],
    "destination_ocean_port":  ["destination ocean port","ocean port","dest port",
                                "destination port","dest ocean port"],
}


def _col_norm(df: pd.DataFrame, alias_map: dict) -> pd.DataFrame:
    """Case/space-insensitive column rename via alias map."""
    lower_map = {
        c.lower().strip().replace(" ", "_").replace("/", "_"): c
        for c in df.columns
    }
    rename = {}
    for target, aliases in alias_map.items():
        if target in df.columns:
            continue
        for alias in aliases:
            key = alias.lower().replace(" ", "_").replace("/", "_")
            if key in lower_map:
                rename[lower_map[key]] = target
                break
    return df.rename(columns=rename)


def _normalize_loads(df: pd.DataFrame) -> pd.DataFrame:
    """Map Inbound Loads Report columns to unified schema."""
    df = _col_norm(df, _LOADS_ALIASES)
    # String columns
    for col in ["container_id","po_number","part_number","item_description","supplier",
                "a_code","origin_ocean_port","vessel_name","container_size"]:
        df[col] = df[col].astype(str).str.strip() if col in df.columns else None
    # Clean numeric PO (remove ".0" from Excel coercion)
    df["po_number"] = df["po_number"].str.split(".").str[0]
    df["a_code"]    = df["a_code"].str.strip()
    # Numeric measures
    for col in ["po_line_number","po_shipment_number","ordered_units"]:
        df[col] = pd.to_numeric(df.get(col), errors="coerce") if col in df.columns else None
    df["shipped_weight_kg"] = pd.to_numeric(df.get("shipped_weight_kg"), errors="coerce")                                if "shipped_weight_kg" in df.columns else None
    # Date columns
    for col in ["departure_date","actual_arrival_port","po_promised_date","eta_final_destination"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d")
            df[col] = df[col].where(df[col].notna() & (df[col] != "NaT"), None)
        else:
            df[col] = None
    return df


def _normalize_iss(df: pd.DataFrame) -> pd.DataFrame:
    """Map ISS / Inbound by PO columns to unified schema."""
    if df.empty:
        return df
    df = _col_norm(df, _ISS_ALIASES)
    # Extract po_number root from raw ISS PO field ("1102935-16-1" → "1102935")
    if "_po_iss" in df.columns:
        df["po_number"] = df["_po_iss"].astype(str).str.split("-").str[0].str.strip()
    elif "po_number" not in df.columns:
        df["po_number"] = None
    # Extract a_code from facility field ("A320-RBTCS" → "A320")
    if "_site_facility" in df.columns:
        df["a_code"] = df["_site_facility"].astype(str).str.split("-").str[0].str.strip()
    elif "a_code" not in df.columns:
        df["a_code"] = None
    # Clean join keys
    df["container_id"] = df["container_id"].astype(str).str.strip()                           if "container_id" in df.columns else None
    df["po_number"]    = df["po_number"].astype(str).str.strip()
    # Numerics
    for col in ["shipped_cartons","shipped_volume_cbm"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        else:
            df[col] = None
    # Date columns
    for col in ["ocean_eta","lfd_demurrage","lfd_per_diem"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d")
            df[col] = df[col].where(df[col].notna() & (df[col] != "NaT"), None)
        else:
            df[col] = None
    # Datetime columns
    for col in ["off_vessel_date","ready_date","actual_pickup","enter_facility",
                "container_empty","return_to_port"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").astype(str)
            df[col] = df[col].where(df[col].notna() & (df[col] != "NaT"), None)
        else:
            df[col] = None
    for col in ["carrier_scac","market","discharged_port","container_status"]:
        if col not in df.columns:
            df[col] = None
    return df


def _parse_lanes(df: pd.DataFrame) -> pd.DataFrame:
    """Parse Robotics Lanes Excel into ar_site_lookup shape."""
    import re as _re
    df = _col_norm(df, _LANES_ALIASES)
    if "_full_address" in df.columns:
        addr = df["_full_address"].astype(str)
        # FC code lives in parentheses at end: "Tighe Logistics - 483 ... (ARW0)"
        df["fc_identifier"]      = addr.str.extract(r"\(([A-Z0-9]{3,6})\)\s*$")[0]
        df["warehouse_operator"] = addr.str.extract(r"^([^-]+?)(?:\s*-|$)")[0].str.strip()
        df["warehouse_address"]  = addr.str.extract(
            r"-\s*(.+?)(?:\s*\([A-Z0-9]+\))?\s*$"
        )[0].str.strip()
    else:
        df["fc_identifier"]      = None
        df["warehouse_operator"] = None
        df["warehouse_address"]  = None
    for col in ["a_code","node_type","destination_ocean_port"]:
        if col not in df.columns:
            df[col] = None
    df["updated_at"] = datetime.now(_PHX).isoformat()
    keep = ["a_code","fc_identifier","warehouse_operator","warehouse_address",
            "node_type","destination_ocean_port","updated_at"]
    return df[keep].dropna(subset=["a_code"])


def _merge_to_unified(df_loads: pd.DataFrame, df_iss: pd.DataFrame,
                      df_lanes: pd.DataFrame, report_date: str) -> pd.DataFrame:
    """JOIN Loads + ISS + Lanes → ar_inbound_unified rows."""
    today = date.today().isoformat()

    # ISS columns to bring into merge
    _iss_cols = ["container_id","po_number","carrier_scac","market","discharged_port",
                 "container_status","a_code","shipped_volume_cbm","shipped_cartons",
                 "ocean_eta","off_vessel_date","ready_date","lfd_demurrage","lfd_per_diem",
                 "actual_pickup","enter_facility","container_empty","return_to_port"]
    if not df_iss.empty:
        iss_sub = df_iss[[c for c in _iss_cols if c in df_iss.columns]]                        .drop_duplicates(subset=["container_id","po_number"])
        merged = df_loads.merge(iss_sub, on=["container_id","po_number"],
                                how="left", suffixes=("","_iss"))
        # Prefer Loads a_code; fall back to ISS a_code
        if "a_code_iss" in merged.columns:
            merged["a_code"] = merged["a_code"].where(
                merged["a_code"].notna() &
                ~merged["a_code"].isin(["None","nan",""]),
                merged["a_code_iss"]
            )
            merged = merged.drop(columns=["a_code_iss"])
        has_iss = merged["carrier_scac"].notna() | merged["container_status"].notna()
        merged["data_source"] = "LOADS_ONLY"
        merged.loc[has_iss, "data_source"] = "MATCHED"
    else:
        merged = df_loads.copy()
        merged["data_source"] = "LOADS_ONLY"
        for col in [c for c in _iss_cols if c not in ["container_id","po_number","a_code"]]:
            if col not in merged.columns:
                merged[col] = None

    # Enrich with site lookup (Lanes)
    if not df_lanes.empty and "a_code" in merged.columns and "a_code" in df_lanes.columns:
        lane_cols = [c for c in ["a_code","fc_identifier","warehouse_operator",
                                  "warehouse_address","node_type","destination_ocean_port"]
                     if c in df_lanes.columns]
        merged = merged.merge(df_lanes[lane_cols], on="a_code", how="left")

    # Derived measures
    ordered = pd.to_numeric(merged.get("ordered_units"), errors="coerce")
    cartons  = pd.to_numeric(merged.get("shipped_cartons"), errors="coerce")
    merged["units_per_carton"] = (ordered / cartons).where(cartons > 0)
    merged["shipped_units"]    = (cartons * merged["units_per_carton"]).round(0)

    # Metadata
    merged["report_date"]       = report_date
    merged["loads_refresh_date"] = today
    merged["iss_refresh_date"]   = today
    merged["uploaded_at"]        = datetime.now(_PHX).isoformat()

    # Ensure all spec columns exist
    for col in _SPEC_COLS:
        if col not in merged.columns:
            merged[col] = None

    return merged[_SPEC_COLS + ["uploaded_at"]]


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def _load_latest_unified(conn: sqlite3.Connection) -> pd.DataFrame:
    row = conn.execute("SELECT MAX(report_date) AS d FROM ar_inbound_unified").fetchone()
    latest = row["d"] if row else None
    if not latest:
        return pd.DataFrame()
    rows = conn.execute(
        "SELECT * FROM ar_inbound_unified WHERE report_date=?", (latest,)
    ).fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def _load_carrier_sla(conn: sqlite3.Connection) -> dict[str, float]:
    try:
        rows = conn.execute("""
            SELECT carrier, av_oa_sla_pct FROM wbr_carrier_results
            WHERE (carrier, generated_at) IN (
                SELECT carrier, MAX(generated_at)
                FROM wbr_carrier_results GROUP BY carrier
            )
        """).fetchall()
        return {r["carrier"]: r["av_oa_sla_pct"] for r in rows if r["av_oa_sla_pct"] is not None}
    except Exception:
        return {}


def _load_site_cfg(conn: sqlite3.Connection) -> pd.DataFrame:
    try:
        rows = conn.execute(
            "SELECT site_code, fc_name, port, capacity_day, site_type, "
            "site_type_effective, recv_days_per_week "
            "FROM plan_sites WHERE active=1 ORDER BY site_code"
        ).fetchall()
        df = pd.DataFrame([dict(r) for r in rows])
        if df.empty:
            return df
        df["site_type"] = df["site_type"].fillna("Static")
        df["recv_days_per_week"] = pd.to_numeric(
            df["recv_days_per_week"], errors="coerce"
        ).fillna(4).astype(int)
        df["weekly_capacity"] = (
            df["capacity_day"].fillna(9).astype(int) * df["recv_days_per_week"]
        )
        return df
    except Exception:
        return pd.DataFrame()


def _load_carrier_sites(conn: sqlite3.Connection) -> dict[str, list[str]]:
    try:
        rows = conn.execute(
            "SELECT scac, site_code FROM plan_site_carrier WHERE active=1"
        ).fetchall()
        result: dict[str, list[str]] = {}
        for r in rows:
            result.setdefault(r["scac"], []).append(r["site_code"])
        return result
    except Exception:
        return {}


def _load_carrier_targets(conn: sqlite3.Connection) -> dict[str, float]:
    try:
        rows = conn.execute("""
            SELECT sc.scac,
                   SUM(sc.allocation_pct
                       * COALESCE(ps.capacity_day,9)
                       * COALESCE(ps.recv_days_per_week,4))
                   / NULLIF(SUM(
                       COALESCE(ps.capacity_day,9)
                       * COALESCE(ps.recv_days_per_week,4)), 0) AS wpct
            FROM plan_site_carrier sc
            JOIN plan_sites ps ON ps.site_code = sc.site_code
            WHERE sc.active=1 AND ps.active=1
            GROUP BY sc.scac
        """).fetchall()
        return {r["scac"]: round(float(r["wpct"]),1) for r in rows if r["wpct"] is not None}
    except Exception:
        return {}


def _load_site_lookup(conn: sqlite3.Connection) -> pd.DataFrame:
    try:
        rows = conn.execute(
            "SELECT a_code,fc_identifier,warehouse_operator,warehouse_address,"
            "node_type,destination_ocean_port,updated_at FROM ar_site_lookup ORDER BY a_code"
        ).fetchall()
        return pd.DataFrame([dict(r) for r in rows])
    except Exception:
        return pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _health(sla_pct, vol_pct) -> str:
    sla_red  = sla_pct is not None and sla_pct < 85
    vol_over = vol_pct is not None and vol_pct > 100
    if sla_red or vol_over:
        return "\U0001f534"  # red circle
    sla_warn = sla_pct is not None and sla_pct < 95
    vol_warn = vol_pct is not None and vol_pct > 85
    if sla_warn or vol_warn:
        return "\U0001f7e1"  # yellow circle
    return "\U0001f7e2"       # green circle


def _site_badge(site_type) -> str:
    if not site_type:
        return ""
    if "Brownfield" in str(site_type):
        return "BF"
    if "Greenfield" in str(site_type):
        return "GF"
    return ""


def _week_start(d: date) -> date:
    return d - timedelta(days=d.weekday())


def _week_label(d: date) -> str:
    return f"W{d.isocalendar()[1]} ({_week_start(d).strftime('%m/%d')})"


def _bucket_weeks(df_ctr: pd.DataFrame, date_col: str, n_weeks: int) -> pd.DataFrame:
    today  = datetime.now(_PHX).date()
    cutoff = today + timedelta(weeks=n_weeks)
    df = df_ctr.copy()
    df["_dt"] = pd.to_datetime(df[date_col], errors="coerce").dt.date
    df = df[(df["_dt"] >= today) & (df["_dt"] <= cutoff)].copy()
    df["week_label"] = df["_dt"].apply(_week_label)
    df["_wk_start"]  = df["_dt"].apply(_week_start)
    return df


_MILESTONE_MAP = {
    "Ocean ETA":             "ocean_eta",
    "ETA Final Destination": "eta_final_destination",
    "PO Promise Date":       "po_promised_date",
    "Ready for Pickup":      "ready_date",
}

_ACTIVE_STATUSES = ["On Water", "In Yard Full", "In Transit to Destination"]

_CARRIER_COLORS = {
    "ATMI": "#4BACC6", "HDDR": "#E8A838", "ARVY": "#70AD47",
    "RKNE": "#9B59B6", "TGHE": "#E74C3C", "AOYV": "#34495E",
}


# ─────────────────────────────────────────────────────────────────────────────
# MAIN RENDER — PANEL 1: PIPELINE FORECAST
# ─────────────────────────────────────────────────────────────────────────────

def _render_pipeline(df_all, df_active_ctr, site_cfg, conn):
    st.markdown("### Container Pipeline — What's Coming & When")

    f1, f2, f3, f4 = st.columns(4)
    f_carriers = f1.multiselect(
        "Carrier", sorted(df_all["carrier_scac"].dropna().unique()), key="pf_car"
    )
    f_status = f2.multiselect(
        "Status", _ACTIVE_STATUSES, default=_ACTIVE_STATUSES, key="pf_status"
    )
    f_port = f3.multiselect(
        "Port", sorted(df_all["discharged_port"].dropna().unique()), key="pf_port"
    )
    milestone = f4.selectbox(
        "Forecast by", list(_MILESTONE_MAP.keys()), key="pf_milestone"
    )
    horizon = st.select_slider("Horizon (weeks)", [4,6,8,12], value=8, key="pf_horizon")

    dff = df_all.copy()
    if f_carriers: dff = dff[dff["carrier_scac"].isin(f_carriers)]
    if f_status:   dff = dff[dff["container_status"].isin(f_status)]
    if f_port:     dff = dff[dff["discharged_port"].isin(f_port)]
    dff_ctr = dff.drop_duplicates(subset=["container_id"])

    date_col = _MILESTONE_MAP[milestone]
    dff_wk   = _bucket_weeks(dff_ctr, date_col, horizon)

    if dff_wk.empty:
        st.info(f"No containers with {milestone} in the next {horizon} weeks.")
        return

    week_order = sorted(
        dff_wk["week_label"].unique(),
        key=lambda w: int(w.split("W")[1].split(" ")[0])
    )
    wk_grp = (
        dff_wk.groupby(["week_label","carrier_scac"])["container_id"]
        .nunique().reset_index()
        .rename(columns={"container_id":"Containers","carrier_scac":"Carrier"})
    )
    color_map = {c: _CARRIER_COLORS.get(c,"#95A5A6") for c in wk_grp["Carrier"].unique()}
    fig = px.bar(
        wk_grp, x="week_label", y="Containers", color="Carrier",
        category_orders={"week_label": week_order},
        color_discrete_map=color_map,
        title=f"Inbound Containers by Week — {milestone}",
        labels={"week_label":"Week"}, text_auto=True,
    )
    fig.update_layout(
        height=380, plot_bgcolor="white", paper_bgcolor="white",
        legend_title_text="Carrier", font=dict(size=12), title_font_size=14,
    )
    st.plotly_chart(fig, use_container_width=True)

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Containers (window)", dff_wk["container_id"].nunique())
    k2.metric("Carriers", dff_wk["carrier_scac"].nunique())
    wt  = dff_wk["shipped_weight_kg"].sum()  if "shipped_weight_kg"  in dff_wk else 0
    cbm = dff_wk["shipped_volume_cbm"].sum() if "shipped_volume_cbm" in dff_wk else 0
    k3.metric("Weight (KG)",   f"{wt:,.0f}"  if wt  > 0 else "—")
    k4.metric("Volume (CBM)",  f"{cbm:,.0f}" if cbm > 0 else "—")

    with st.expander("\U0001f4ca Site (A-Code) \u00d7 Week breakdown", expanded=False):
        if not site_cfg.empty:
            type_map = dict(zip(site_cfg["site_code"], site_cfg["site_type"]))
        else:
            type_map = {}
        site_piv = (
            dff_wk.groupby(["a_code","week_label"])["container_id"]
            .nunique().reset_index()
            .rename(columns={"container_id":"ctrs","a_code":"Site"})
            .pivot_table(index="Site", columns="week_label", values="ctrs", fill_value=0)
        )
        site_piv = site_piv.reindex(
            columns=[c for c in week_order if c in site_piv.columns]
        )
        site_piv["Total"] = site_piv.sum(axis=1)
        site_piv.index = [
            f"{s}  [{_site_badge(type_map.get(s))}]"
            if _site_badge(type_map.get(s)) else s
            for s in site_piv.index
        ]
        st.dataframe(site_piv, use_container_width=True)

    # Demurrage risk
    if "lfd_demurrage" in dff_ctr.columns:
        today  = datetime.now(_PHX).date()
        d_risk = dff_ctr[dff_ctr["container_status"]=="In Yard Full"].copy()
        d_risk["_lfd"] = pd.to_datetime(d_risk["lfd_demurrage"], errors="coerce").dt.date
        at_risk = d_risk[d_risk["_lfd"].notna() & (d_risk["_lfd"] <= today + timedelta(days=3))]
        if not at_risk.empty:
            st.error(
                f"\U0001f6a8 **{len(at_risk)} container(s) at demurrage risk** "
                f"(LFD \u2264 3 days) — Carriers: "
                f"{', '.join(at_risk['carrier_scac'].dropna().unique())}"
            )
            st.dataframe(
                at_risk[["container_id","carrier_scac","discharged_port","a_code","lfd_demurrage"]]
                .sort_values("lfd_demurrage"),
                hide_index=True, use_container_width=True,
            )


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 2: ALLOCATION MANAGER
# ─────────────────────────────────────────────────────────────────────────────

def _render_allocation_manager(df_active_ctr, total_active, carrier_sla,
                                carrier_tgt, carrier_sites, site_cfg, conn):
    st.markdown("### Carrier Allocation Manager")
    st.caption(
        "Review actual pipeline distribution against target allocations and WBR SLA performance. "
        "Site type context shown throughout. Edit config in **Planning → Config**."
    )

    # Site type config — read-only; edits live in Planning -> Config
    with st.expander("🏭 Site Configuration (read-only)", expanded=False):
        st.info("To edit site configuration, go to **Planning → Config**.")
        if site_cfg.empty:
            st.caption("No active sites configured.")
        else:
            display_cols = [c for c in ["site_code","fc_name","site_type",
                "site_type_effective","capacity_day","recv_days_per_week"]
                if c in site_cfg.columns]
            rename_map = {
                "site_code": "Site", "fc_name": "FC",
                "site_type": "Type", "site_type_effective": "Effective",
                "capacity_day": "Max/day", "recv_days_per_week": "Days/wk",
            }
            disp = site_cfg[display_cols].rename(columns=rename_map).copy()
            if "Days/wk" in disp.columns and "Max/day" in disp.columns:
                disp["Weekly Cap"] = disp["Max/day"] * disp["Days/wk"]
            st.dataframe(disp, hide_index=True, use_container_width=True)

    # Target vs. Actual table
    st.markdown("#### Current Allocation — Target vs. Actual vs. SLA")
    actual_by_scac: dict[str,int] = {}
    if total_active > 0 and not df_active_ctr.empty:
        for scac, grp in df_active_ctr.groupby("carrier_scac"):
            actual_by_scac[scac] = grp["container_id"].nunique()

    all_scacs = sorted(set(list(actual_by_scac.keys())) | set(carrier_tgt.keys()))
    alloc_rows = []
    for scac in all_scacs:
        actual_n   = actual_by_scac.get(scac, 0)
        actual_pct = round(actual_n / total_active * 100, 1) if total_active else None
        target_pct = carrier_tgt.get(scac)
        sla_pct    = carrier_sla.get(scac)
        gap        = round((actual_pct or 0) - (target_pct or 0), 1) if (actual_pct and target_pct) else None
        sites      = carrier_sites.get(scac, [])
        badges     = []
        if not site_cfg.empty:
            for s in sites:
                sr = site_cfg[site_cfg["site_code"] == s]
                b  = _site_badge(sr["site_type"].iloc[0] if not sr.empty else None)
                if b:
                    badges.append(f"{s}[{b}]")
        health = _health(sla_pct, None)
        alloc_rows.append({
            "":            health,
            "Carrier":     scac,
            "Active Ctrs": actual_n,
            "Actual %":    f"{actual_pct:.1f}%" if actual_pct is not None else "—",
            "Target %":    f"{target_pct:.0f}%" if target_pct is not None else "—",
            "Gap":         f"{gap:+.1f}%" if gap is not None else "—",
            "SLA (AV->OA)": f"{sla_pct:.0f}%" if sla_pct is not None else "no WBR data",
            "Site Flags":  ", ".join(badges) if badges else "",
        })

    if alloc_rows:
        st.dataframe(pd.DataFrame(alloc_rows), hide_index=True, use_container_width=True)
        st.caption(
            f"Total active containers: **{total_active}** · "
            "SLA from most recent WBR run · Gap = Actual \u2212 Target"
        )
    else:
        st.info("Upload ar_inbound_unified and confirm carrier config in Planning tab.")

    with st.expander("✏️ Target Allocations (read-only)", expanded=False):
        st.info("To edit target allocations, go to **Planning → Config**.")
        sc_rows = conn.execute("""
            SELECT sc.site_code, sc.scac, c.carrier_name, sc.allocation_pct
            FROM plan_site_carrier sc
            JOIN plan_carriers c ON c.scac = sc.scac
            WHERE sc.active=1 ORDER BY sc.site_code, sc.scac
        """).fetchall()
        if not sc_rows:
            st.caption("No site-carrier assignments configured.")
        else:
            rows = [{"Site": r["site_code"], "Carrier": f"{r['carrier_name']} ({r['scac']})",
                     "Target %": f"{r['allocation_pct']:.0f}%" if r["allocation_pct"] is not None else "—"}
                    for r in sc_rows]
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 3: SIMULATOR
# ─────────────────────────────────────────────────────────────────────────────

def _render_simulator(df_active_ctr, total_active, carrier_sla,
                      carrier_tgt, carrier_sites, site_cfg, conn):
    st.markdown("### Allocation Simulator")
    st.caption(
        "Adjust carrier allocation sliders to see projected container counts, "
        "capacity utilisation, and SLA-based health indicators."
    )

    if total_active == 0:
        st.info("No active container data. Upload ar_inbound_unified to enable simulation.")
        return

    sc1, sc2 = st.columns(2)
    sim_horizon   = sc1.select_slider("Projection horizon (weeks)", [2,4,6,8], value=4, key="sim_horizon")
    sim_milestone = sc2.selectbox(
        "Forecast date basis", ["Ocean ETA","ETA Final Destination"], key="sim_milestone"
    )

    fwd_col   = _MILESTONE_MAP[sim_milestone]
    dff_fwd   = _bucket_weeks(df_active_ctr, fwd_col, sim_horizon)
    fwd_total = dff_fwd["container_id"].nunique() if not dff_fwd.empty else total_active

    st.markdown(
        f"**Projected containers — next {sim_horizon} weeks: {fwd_total}**  "
        f"(basis: {sim_milestone})"
    )
    st.markdown("---")

    sim_scacs = sorted(
        set(
            (dff_fwd["carrier_scac"].dropna().unique().tolist() if not dff_fwd.empty else [])
            + list(carrier_tgt.keys())
        )
    )
    if not sim_scacs:
        st.info("No carriers in forward pipeline.")
        return

    st.markdown("**Adjust allocation % per carrier** — target: sum to 100%")
    actual_by_scac: dict[str,int] = {}
    if total_active > 0 and not df_active_ctr.empty:
        for scac, grp in df_active_ctr.groupby("carrier_scac"):
            actual_by_scac[scac] = grp["container_id"].nunique()

    sim_vals: dict[str,float] = {}
    n_cols      = min(len(sim_scacs), 3)
    slider_cols = st.columns(n_cols)
    for i, scac in enumerate(sim_scacs):
        actual_pct_now = (
            round(actual_by_scac.get(scac,0) / total_active * 100, 0)
            if total_active else 0.0
        )
        default = carrier_tgt.get(scac, actual_pct_now)
        with slider_cols[i % n_cols]:
            sim_vals[scac] = st.slider(
                scac, 0.0, 100.0, value=float(round(default,0)),
                step=5.0, key=f"sim_{scac}"
            )

    total_sim  = sum(sim_vals.values())
    sum_col, pct_col = st.columns([3,1])
    if abs(total_sim - 100.0) > 0.5:
        sum_col.warning(f"\u26a0\ufe0f Allocations sum to **{total_sim:.0f}%** — adjust to reach 100%")
    else:
        sum_col.success("\u2705 Allocations sum to 100%")
    pct_col.metric("Total", f"{total_sim:.0f}%")

    st.markdown("---")
    st.markdown("#### Projected Outcomes")

    sim_rows = []
    for scac in sim_scacs:
        sim_pct    = sim_vals[scac]
        proj_ctrs  = round(fwd_total * sim_pct / 100)
        sla_pct    = carrier_sla.get(scac)
        sites      = carrier_sites.get(scac, [])
        carrier_cap = None
        site_flag_str = ""
        if not site_cfg.empty and sites:
            sc_cfg = site_cfg[site_cfg["site_code"].isin(sites)]
            if not sc_cfg.empty:
                carrier_cap = int(sc_cfg["weekly_capacity"].sum())
                flags = [
                    f"{r['site_code']} [{_site_badge(r['site_type'])}]"
                    for _, r in sc_cfg.iterrows()
                    if _site_badge(r["site_type"])
                ]
                site_flag_str = ", ".join(flags)

        vol_pct = round(proj_ctrs / carrier_cap * 100, 1) if carrier_cap else None
        health  = _health(sla_pct, vol_pct)
        sim_rows.append({
            "":            health,
            "Carrier":     scac,
            "Sim %":       f"{sim_pct:.0f}%",
            "Proj. Ctrs":  proj_ctrs,
            "Wkly Cap":    str(carrier_cap) if carrier_cap else "—",
            "Cap Used":    f"{vol_pct:.0f}%" if vol_pct is not None else "—",
            "SLA (AV->OA)": f"{sla_pct:.0f}%" if sla_pct else "no data",
            "Site Flags":  site_flag_str,
        })

    st.dataframe(pd.DataFrame(sim_rows), hide_index=True, use_container_width=True)
    st.markdown(
        "**\U0001f7e2 Green** — SLA \u2265 95% and volume \u2264 85% of weekly site capacity  \n"
        "**\U0001f7e1 Yellow** — SLA 85\u201394% OR volume 85\u2013100% of capacity  \n"
        "**\U0001f534 Red** — SLA < 85% OR volume exceeds site capacity"
    )

    # Rebalancing suggestions
    red   = [r for r in sim_rows if r[""] == "\U0001f534"]
    green = [r for r in sim_rows if r[""] == "\U0001f7e2"]
    if red and green:
        st.markdown("---")
        st.markdown("#### \U0001f4a1 Rebalancing Suggestions")
        for rc in red:
            scac    = rc["Carrier"]
            reasons = []
            sla_v   = carrier_sla.get(scac)
            cap_s   = rc["Cap Used"]
            if sla_v and sla_v < 85:
                reasons.append(f"SLA at {sla_v:.0f}% (below 85% floor)")
            if cap_s != "—" and float(cap_s.replace("%","")) > 100:
                reasons.append(f"over site capacity ({cap_s})")
            green_names = [g["Carrier"] for g in green]
            st.warning(
                f"**{scac}** is \U0001f534 — {'; '.join(reasons) or 'see indicators above'}. "
                f"Consider shifting volume to: {', '.join(green_names)}"
            )

    # Simulated weekly chart
    if not dff_fwd.empty:
        st.markdown("---")
        st.markdown("#### Simulated Weekly Distribution")
        wk_actual = (
            dff_fwd.groupby(["week_label","carrier_scac"])["container_id"]
            .nunique().reset_index()
            .rename(columns={"container_id":"actual","carrier_scac":"Carrier"})
        )
        wk_totals = wk_actual.groupby("week_label")["actual"].sum()
        wk_actual["simulated"] = wk_actual.apply(
            lambda row: round(
                wk_totals.get(row["week_label"],0) * sim_vals.get(row["Carrier"],0) / 100
            ), axis=1
        )
        week_order_sim = sorted(
            wk_actual["week_label"].unique(),
            key=lambda w: int(w.split("W")[1].split(" ")[0])
        )
        color_map_sim = {c: _CARRIER_COLORS.get(c,"#95A5A6") for c in wk_actual["Carrier"].unique()}
        fig2 = px.bar(
            wk_actual, x="week_label", y="simulated", color="Carrier",
            category_orders={"week_label": week_order_sim},
            color_discrete_map=color_map_sim,
            title="Simulated Container Distribution by Week",
            labels={"week_label":"Week","simulated":"Projected Containers"},
            text_auto=True,
        )
        fig2.update_layout(
            height=340, plot_bgcolor="white", paper_bgcolor="white",
            font=dict(size=12), title_font_size=14,
        )
        st.plotly_chart(fig2, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def render_inbound_forecast_tab(conn: sqlite3.Connection) -> None:
    init_inbound_forecast_db(conn)

    st.markdown("## \U0001f4e6 Inbound Forecast & Carrier Allocation")
    st.caption(
        "Daily `ar_inbound_unified` snapshot powers container pipeline forecasts, "
        "carrier allocation targets, and volume simulation."
    )

    # ── Robotics Lanes (Site Reference — quarterly) ──────────────────────────
    with st.expander("\U0001f5fa\ufe0f  Robotics Lanes — Site Reference (quarterly upload)", expanded=False):
        st.markdown(
            "Enables **A-code → warehouse / FC / port enrichment** on every daily upload. "
            "Upload once, then only re-upload when Robotics adds or changes sites (~every 2\u20134 months).\n\n"
            "**File:** The Robotics-managed Excel file titled *Robotics Lanes* (or similar). "
            "Maintained by the AR team.\n\n"
            "**Required columns the app looks for:**\n"
            "- *Destination Node Code* — A-code (e.g., A320)\n"
            "- *Destination Final Node Address* — full address including FC code in parentheses "
            "(e.g., `Tighe Logistics - 483 Wildwood Ave, Woburn, MA 01801, USA (ARW0)`)\n"
            "- *Node Type* — MC, OS, or FC\n"
            "- *Destination Ocean Port* — UN/LOCODE (e.g., USORF)"
        )
        df_lanes_status = _load_site_lookup(conn)
        if not df_lanes_status.empty:
            last_upd = df_lanes_status["updated_at"].max()[:10] if "updated_at" in df_lanes_status.columns else "unknown"
            st.success(f"\u2705 {len(df_lanes_status)} A-codes loaded \u00b7 last updated {last_upd}")
            with st.expander("View current site reference", expanded=False):
                st.dataframe(df_lanes_status.drop(columns=["updated_at"], errors="ignore"),
                             hide_index=True, use_container_width=True)
        else:
            st.warning("No site reference loaded. Upload to enable warehouse/FC enrichment (optional).")
        lanes_file = st.file_uploader("Robotics Lanes XLSX", type=["xlsx","csv"],
                                       key="lanes_upload", label_visibility="collapsed")
        if lanes_file and st.button("\U0001f4be Save Site Reference", key="lanes_save"):
            try:
                raw_l = (pd.read_excel(lanes_file) if lanes_file.name.endswith(".xlsx")
                         else pd.read_csv(lanes_file))
                df_lanes_new = _parse_lanes(raw_l)
                conn.execute("DELETE FROM ar_site_lookup")
                df_lanes_new.to_sql("ar_site_lookup", conn, if_exists="append", index=False)
                conn.commit()
                st.success(f"\u2705 {len(df_lanes_new)} A-codes saved.")
                st.rerun()
            except Exception as e:
                st.error(f"Lanes parse failed: {e}")

    # ── Daily Merge: Inbound Loads + ISS → ar_inbound_unified ────────────────
    with st.expander("\U0001f4c2 Daily Upload — Merge Inbound Loads + ISS", expanded=True):
        st.caption(
            "Upload both source reports and click **Merge & Load**. "
            "The app joins them on Container ID + PO number, enriches with the site reference, "
            "and writes the result to the forecast data store."
        )
        dc1, dc2 = st.columns(2)

        dc1.markdown(
            "**\U0001f4e6 Inbound Loads Report** *(required)*\n\n"
            "**Source:** Oracle / AR system\n\n"
            "**Cadence:** Pull daily, Sun\u2013Sat\n\n"
            "**What it provides:** PO-level material detail \u2014 part numbers, quantities, "
            "item descriptions, suppliers, vessel name, container size, weight, "
            "and date milestones (departure, arrival, PO promise date)\n\n"
            "**File name:** Typically *Amazon Robotics Inbound Loads Report* or similar (.xlsx)"
        )
        loads_file = dc1.file_uploader("Inbound Loads Report (.xlsx / .csv)", type=["xlsx","csv"],
                                        key="loads_upload")

        dc2.markdown(
            "**\U0001f6a2 Import Shipment Status (ISS)** *(optional)*\n\n"
            "**Source:** OBLT / ISS system\n\n"
            "**Cadence:** Pull Mon\u2013Fri (not available weekends)\n\n"
            "**What it adds:** Carrier SCAC, discharged port, container status, CBM, "
            "carton count, and all port-to-facility milestone dates "
            "(off vessel, ready date, LFD demurrage, enter facility, etc.)\n\n"
            "**File name:** Typically *Import Shipment Status* (.xlsx). "
            "Skip on weekends \u2014 Loads-only upload still works."
        )
        iss_file = dc2.file_uploader("Import Shipment Status (.xlsx / .csv)", type=["xlsx","csv"],
                                      key="iss_upload")
        rpt_date_val = st.date_input(
            "Report date", value=datetime.now(_PHX).date(), key="merge_rpt_date",
            help="Date this snapshot represents (defaults to today)"
        )
        if st.button("\u2699\ufe0f  Merge & Load", key="merge_load",
                     use_container_width=True, disabled=(loads_file is None)):
            with st.spinner("Merging reports\u2026"):
                try:
                    df_loads_raw = (pd.read_excel(loads_file) if loads_file.name.endswith(".xlsx")
                                    else pd.read_csv(loads_file))
                    if iss_file:
                        df_iss_raw = (pd.read_excel(iss_file) if iss_file.name.endswith(".xlsx")
                                      else pd.read_csv(iss_file))
                    else:
                        df_iss_raw = pd.DataFrame()
                    df_lanes_ref = _load_site_lookup(conn)
                    df_loads_n   = _normalize_loads(df_loads_raw)
                    df_iss_n     = _normalize_iss(df_iss_raw) if not df_iss_raw.empty else pd.DataFrame()
                    df_unified   = _merge_to_unified(
                        df_loads_n, df_iss_n, df_lanes_ref, rpt_date_val.isoformat()
                    )
                    conn.execute("DELETE FROM ar_inbound_unified WHERE report_date=?",
                                 (rpt_date_val.isoformat(),))
                    df_unified.to_sql("ar_inbound_unified", conn, if_exists="append", index=False)
                    conn.commit()
                    n_ctr     = df_unified["container_id"].nunique()
                    n_matched = (df_unified["data_source"] == "MATCHED").sum()
                    src_note  = f"{n_matched:,} rows matched" if not df_iss_raw.empty else "Loads only (no ISS)"
                    st.success(
                        f"\u2705 {len(df_unified):,} rows loaded \u00b7 "
                        f"{n_ctr} containers \u00b7 {src_note} \u00b7 as-of {rpt_date_val}"
                    )
                    st.rerun()
                except Exception as e:
                    import traceback
                    st.error(f"Merge failed: {e}")
                    st.code(traceback.format_exc())

    # Load
    df_all        = _load_latest_unified(conn)
    carrier_sla   = _load_carrier_sla(conn)
    site_cfg      = _load_site_cfg(conn)
    carrier_sites = _load_carrier_sites(conn)
    carrier_tgt   = _load_carrier_targets(conn)

    if df_all.empty:
        st.info(
            "No inbound data loaded yet. "
            "Upload your **Inbound Loads Report** using the expander above to get started. "
            "Adding the ISS report enriches with carrier, port, and milestone data."
        )
        return

    # Freshness banner
    rpt_date = df_all["report_date"].max()
    try:
        days_old = (datetime.now(_PHX).date() - pd.to_datetime(rpt_date).date()).days
    except Exception:
        days_old = 0
    if days_old > 1:
        st.warning(f"\u26a0\ufe0f Data is **{days_old} day(s) old** (as-of {rpt_date}). Upload today's file.")
    else:
        n_ctr_total = df_all["container_id"].nunique()
        st.caption(
            f"Data as-of: **{rpt_date}** \u00b7 "
            f"**{n_ctr_total}** containers \u00b7 {len(df_all):,} rows"
        )

    df_active     = df_all[df_all["container_status"].isin(_ACTIVE_STATUSES)].copy()
    df_active_ctr = df_active.drop_duplicates(subset=["container_id"]).copy()
    total_active  = df_active_ctr["container_id"].nunique()

    p1, p2, p3 = st.tabs([
        "\U0001f4c5 Pipeline Forecast",
        "\U0001f3af Allocation Manager",
        "\U0001f52c Simulator",
    ])

    with p1:
        _render_pipeline(df_all, df_active_ctr, site_cfg, conn)

    with p2:
        _render_allocation_manager(
            df_active_ctr, total_active, carrier_sla,
            carrier_tgt, carrier_sites, site_cfg, conn
        )

    with p3:
        _render_simulator(
            df_active_ctr, total_active, carrier_sla,
            carrier_tgt, carrier_sites, site_cfg, conn
        )
