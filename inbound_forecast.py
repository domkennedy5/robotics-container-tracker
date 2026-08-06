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

_ALIASES: dict[str, list[str]] = {
    "container_id":          ["container_no","container no","containerid","container id","container"],
    "carrier_scac":          ["scac","dray scac","carrier_code","carrier code"],
    "fc_identifier":         ["fc","fc id","facility","fc code"],
    "ocean_eta":             ["eta","vessel eta","ocean eta"],
    "eta_final_destination": ["final eta","final dest eta","eta final destination"],
    "container_status":      ["status","container status"],
    "discharged_port":       ["port","discharge port","discharged port"],
    "shipped_weight_kg":     ["weight kg","weight (kg)","shipped weight kg","weight_kg"],
    "shipped_volume_cbm":    ["volume cbm","cbm","shipped volume cbm","volume_cbm"],
}


def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    lower_map = {c.lower().strip().replace(" ","_"): c for c in df.columns}
    rename = {}
    for spec_col in _SPEC_COLS:
        if spec_col in df.columns:
            continue
        if spec_col in lower_map:
            rename[lower_map[spec_col]] = spec_col
            continue
        for alias in _ALIASES.get(spec_col, []):
            key = alias.lower().replace(" ","_")
            if key in lower_map:
                rename[lower_map[key]] = spec_col
                break
    df = df.rename(columns=rename)
    for col in _SPEC_COLS:
        if col not in df.columns:
            df[col] = None
    return df[_SPEC_COLS]


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
        "Set carrier target allocations, review actual pipeline distribution, "
        "and compare against WBR SLA performance. Site type context visible throughout."
    )

    # Site type config
    with st.expander("\U0001f3ed Site Type Configuration", expanded=False):
        st.caption(
            "Tag each site: **Static** (steady-state) · "
            "**BF** (Brownfield/Retrofit — constrained) · "
            "**GF** (Greenfield/New Build — ramping). "
            "Effective date tracks transitions over time."
        )
        if site_cfg.empty:
            st.info("No active sites. Configure in the Planning tab.")
        else:
            hdr = st.columns([1,1,2,2,1,1])
            for lbl, col in zip(["Site","FC","Type","Effective","Max/day","Days/wk"], hdr):
                col.caption(f"**{lbl}**")

            edits: list[tuple] = []
            for _, row in site_cfg.iterrows():
                ec = st.columns([1,1,2,2,1,1])
                ec[0].markdown(f"**{row['site_code']}**")
                ec[1].caption(row.get("fc_name") or "")
                opts = ["Static","BF (Brownfield/Retrofit)","GF (Greenfield/New Build)"]
                cur  = row.get("site_type") or "Static"
                idx  = opts.index(cur) if cur in opts else 0
                new_type = ec[2].selectbox(
                    "t", opts, index=idx,
                    key=f"stype_{row['site_code']}", label_visibility="collapsed"
                )
                eff_default = pd.to_datetime(
                    row.get("site_type_effective") or date.today()
                ).date()
                new_eff = ec[3].date_input(
                    "e", value=eff_default,
                    key=f"seff_{row['site_code']}", label_visibility="collapsed"
                )
                new_cap  = ec[4].number_input(
                    "c", value=int(row.get("capacity_day") or 9),
                    min_value=1, max_value=30, step=1,
                    key=f"scap_{row['site_code']}", label_visibility="collapsed"
                )
                new_days = ec[5].number_input(
                    "d", value=int(row.get("recv_days_per_week") or 4),
                    min_value=1, max_value=6, step=1,
                    key=f"sdays_{row['site_code']}", label_visibility="collapsed"
                )
                edits.append((new_type, new_eff.isoformat(), new_cap, new_days, row["site_code"]))

            if st.button("\U0001f4be Save Site Configuration", key="save_site_cfg"):
                for (st_t, st_e, st_c, st_d, sc) in edits:
                    conn.execute(
                        "UPDATE plan_sites SET site_type=?, site_type_effective=?, "
                        "capacity_day=?, recv_days_per_week=? WHERE site_code=?",
                        (st_t, st_e, st_c, st_d, sc)
                    )
                conn.commit()
                st.success("Site configuration saved.")
                st.rerun()

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

    with st.expander("\u270f\ufe0f Edit Target Allocations (by Site-Carrier)", expanded=False):
        sc_rows = conn.execute("""
            SELECT sc.id, sc.site_code, sc.scac, c.carrier_name, sc.allocation_pct
            FROM plan_site_carrier sc
            JOIN plan_carriers c ON c.scac = sc.scac
            WHERE sc.active=1 ORDER BY sc.site_code, sc.scac
        """).fetchall()
        if not sc_rows:
            st.info("No site-carrier assignments. Configure in the Planning tab.")
        else:
            edits_alloc: dict[int,float] = {}
            cur_site = None
            for r in sc_rows:
                if r["site_code"] != cur_site:
                    if cur_site: st.markdown("---")
                    st.markdown(f"**{r['site_code']}**")
                    cur_site = r["site_code"]
                ac1, ac2 = st.columns([3,1])
                ac1.caption(f"{r['carrier_name']} ({r['scac']})")
                edits_alloc[r["id"]] = ac2.number_input(
                    "%", value=float(r["allocation_pct"] or 0),
                    min_value=0.0, max_value=100.0, step=5.0,
                    key=f"al_{r['id']}", label_visibility="collapsed"
                )
            if st.button("\U0001f4be Save Target Allocations", key="save_alloc"):
                for rid, pct in edits_alloc.items():
                    conn.execute(
                        "UPDATE plan_site_carrier SET allocation_pct=? WHERE id=?",
                        (pct, rid)
                    )
                conn.commit()
                st.success("Allocations saved.")
                st.rerun()


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

    # Upload section
    with st.expander("\U0001f4c2 Upload ar_inbound_unified (daily CSV / XLSX)", expanded=False):
        st.caption(
            "Upload `ar_inbound_unified_YYYY-MM-DD.csv` (or .xlsx). "
            "Previous data for the same report_date is replaced automatically."
        )
        uc1, uc2 = st.columns([3,1])
        up_file  = uc1.file_uploader(
            "File", type=["csv","xlsx"], key="aru_upload", label_visibility="collapsed"
        )
        do_upload = uc2.button("\u2b06\ufe0f Load", key="aru_load", use_container_width=True)

        if up_file and do_upload:
            with st.spinner("Parsing and loading\u2026"):
                try:
                    raw = (pd.read_excel(up_file) if up_file.name.endswith(".xlsx")
                           else pd.read_csv(up_file))
                    df_up = _normalize_df(raw)
                    rpt_date = (
                        str(df_up["report_date"].dropna().iloc[0])[:10]
                        if df_up["report_date"].notna().any()
                        else datetime.now(_PHX).date().isoformat()
                    )
                    df_up["report_date"] = rpt_date
                    df_up["uploaded_at"] = datetime.now(_PHX).isoformat()
                    conn.execute(
                        "DELETE FROM ar_inbound_unified WHERE report_date=?", (rpt_date,)
                    )
                    df_up.to_sql("ar_inbound_unified", conn, if_exists="append", index=False)
                    conn.commit()
                    n_ctr = df_up["container_id"].nunique()
                    st.success(
                        f"\u2705 {len(df_up):,} rows loaded \u00b7 {n_ctr} containers \u00b7 as-of {rpt_date}"
                    )
                except Exception as e:
                    st.error(f"Upload failed: {e}")

    # Load
    df_all        = _load_latest_unified(conn)
    carrier_sla   = _load_carrier_sla(conn)
    site_cfg      = _load_site_cfg(conn)
    carrier_sites = _load_carrier_sites(conn)
    carrier_tgt   = _load_carrier_targets(conn)

    if df_all.empty:
        st.info(
            "No inbound data loaded yet. Upload the `ar_inbound_unified` file above "
            "to enable forecasting and allocation analysis."
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
