"""
wbr_engine.py — WBR metric calculations for Robotics Destination Dray

Methodology (updated Aug 2026 — OBLT-primary):
- Population:  OBLT containers with EVENT_602 (OA/gate-out) date in week Sun–Sat.
               De-duplicated: each container counted once, in the week its EVENT_602 fires.
- AV->OA:      OBLT EVENT_308 (AV) → EVENT_602 (OA); measurable = has_OA OR (has_AV > 3 frac days
               before report with no OA). GVT Terminal Avail supplements missing AV records.
- OA->Del:     OBLT OA (EVENT_602) → RD (EVENT_511); BOS-only, excl A320-RBTCS; in-transit = elapsed.
- Empty->Term: OBLT EVENT_602 (OA) → EVENT_511 (RD); ALL Robotics markets; completed only
               (both OA and RD required). Do NOT use GVT Container Empty / Return to Port.
- E2E:         OBLT VD (Vessel Departure) → GVT Enter Facility; completed only.
- OTP:         GVT Current Dray SLA vs. Enter Facility (primary); Inbound Loads fallback.

Data source roles (Aug 2026):
  OBLT   — PRIMARY: population bucketing (EVENT_602), AV→OA, Empty→Term, OA→Del
  GVT    — E2E Transit (Enter Facility) + OTP (delivery date); facility/scac lookup
  IL     — OTP fallback when GVT Dray SLA is absent
  ISS    — Supplemental context only
"""
from __future__ import annotations

import io
import math
from datetime import date, datetime, timedelta
from typing import Any

import numpy as np
import openpyxl
import pandas as pd

AV_OA_SLA_DAYS  = 3
OA_DEL_SLA_DAYS = 3
EMPTY_TERM_SLA  = 3
EXCL_FACILITY   = 'A320-RBTCS'


def week_bounds(week_num: int, year: int) -> tuple:
    iso_monday = date.fromisocalendar(year, week_num, 1)
    sun = iso_monday - timedelta(days=1)   # Sunday opening the WBR week
    sat = sun + timedelta(days=6)          # Saturday closing
    return sun, sat


def guess_week(ready_dates, report_date=None) -> tuple:
    """Return (week_num, year) of the WBR week.

    GVT files span 6-12 prior weeks so a naive most-common picks the wrong
    week.  When report_date is provided we narrow to a 7-day backward window
    (the WBR Sun-Sat week sits immediately before report_date Monday).
    """
    dates = pd.to_datetime(ready_dates, errors='coerce').dropna()
    if dates.empty:
        return None, date.today().year
    if report_date is not None:
        rpt_ts = pd.Timestamp(report_date)
        window = dates[
            (dates >= rpt_ts - pd.Timedelta(days=7)) &
            (dates <= rpt_ts + pd.Timedelta(days=1))
        ]
        if not window.empty:
            dates = window
    iso_weeks = dates.apply(lambda d: (d + timedelta(days=1)).isocalendar()[:2])
    most_common = pd.Series(iso_weeks).value_counts().idxmax()
    return int(most_common[1]), int(most_common[0])


def guess_week_from_oblt(oblt: pd.DataFrame, report_date=None) -> tuple:
    """Return (week_num, year) from OBLT EVENT_602 (OA/gate-out) dates.

    Preferred over GVT-based detection when OBLT is the population source.
    Falls back to None if no OA events are present.
    """
    oa_dates = oblt[oblt['status'] == 'OA']['date'].dropna()
    if oa_dates.empty:
        return None, date.today().year
    return guess_week(oa_dates, report_date)


def _to_date(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    if isinstance(v, str):
        try: return datetime.fromisoformat(v).date()
        except: return None
    return None


def _to_ts(v):
    if v is None: return pd.NaT
    if isinstance(v, (datetime, date)): return pd.Timestamp(v)
    if isinstance(v, str):
        try: return pd.Timestamp(v)
        except: return pd.NaT
    return pd.NaT


def load_gvt(file_bytes: bytes) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    h = {v: i for i, v in enumerate(hdrs) if v}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        ctr = row[h.get('Container', -1)]
        if not ctr: continue
        rows.append({
            'container':      str(ctr).strip(),
            'ready_date':     _to_date(row[h.get('Ready Date/Time', -1)]),
            'port':           row[h.get('Discharged Port', -1)],
            'facility':       row[h.get('Facility', -1)],
            'market':         row[h.get('Market', -1)],
            'scac':           row[h.get('Dray SCAC(FL)', -1)],
            'enter_facility': _to_ts(row[h.get('Enter Facility', -1)]),
            'container_empty':_to_ts(row[h.get('Container Empty', -1)]),
            'return_to_port': _to_ts(row[h.get('Return to Port', -1)]),
            'status':         row[h.get('Container Status', -1)],
            'terminal_avail': _to_ts(row[h.get('Terminal Avail', -1)]),
            'current_dray_sla': _to_ts(row[h.get('Current Dray SLA', -1)]),
            'category':       row[h.get('Category', -1)],
        })
    wb.close()
    return pd.DataFrame(rows)


def load_oblt(file_bytes: bytes) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb['ObltData'] if 'ObltData' in wb.sheetnames else wb.active
    hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    h = {v: i for i, v in enumerate(hdrs) if v}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        ctr = row[h.get('tracking_id', -1)]
        status = row[h.get('status', -1)]
        if not ctr or status not in ('AV','OA','RD','VD'): continue
        rows.append({'container': str(ctr).strip(), 'status': status, 'date': _to_ts(row[h.get('status_date', -1)])})
    wb.close()
    df = pd.DataFrame(rows)
    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    return df


def load_inbound_loads(file_bytes: bytes) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb['Inbound Loads'] if 'Inbound Loads' in wb.sheetnames else list(wb)[0]
    rows_raw = list(ws.iter_rows(values_only=True))
    if not rows_raw: return pd.DataFrame()
    hdrs = rows_raw[0]
    h = {v: i for i, v in enumerate(hdrs) if v}
    rows = []
    for row in rows_raw[1:]:
        ctr = row[h.get('Container Id', -1)]
        if not ctr: continue
        rows.append({
            'container':      str(ctr).strip(),
            'po_promised':    _to_date(row[h.get('PO Promised Date', -1)]),
            'actual_arrival': _to_date(row[h.get('Actual Arrival At Final Destination', -1)]),
        })
    wb.close()
    return pd.DataFrame(rows)


def _latest(oblt: pd.DataFrame, status: str):
    sub = oblt[oblt['status'] == status].sort_values('date')
    return sub.groupby('container')['date'].last()


def compute_metrics(gvt, oblt, inbound, week_start, week_end, report_date) -> dict:
    # ── Population: OBLT EVENT_602 (OA) date in reporting week ──────────────────
    # Each container counted once, in the week its OA event fires.
    # De-duplicated: latest OA within the week if a container has multiple records.
    oa_wk = oblt[
        (oblt['status'] == 'OA') &
        oblt['date'].notna() &
        (oblt['date'].dt.date >= week_start) &
        (oblt['date'].dt.date <= week_end)
    ].sort_values('date')
    pop_oa = oa_wk.groupby('container').last().reset_index()   # one row per container
    pop_ctrs = set(pop_oa['container'])
    n = len(pop_ctrs)
    if n == 0:
        return _empty_metrics()

    # Week-specific OA timestamps keyed by container (used for AV→OA and Empty→Term)
    oa_ev = pop_oa.set_index('container')['date']

    # ── GVT subset: facility, scac, E2E (Enter Facility), OTP (Dray SLA) ────────
    gvt_pop = gvt[gvt['container'].isin(pop_ctrs)].copy()
    if not gvt_pop.empty:
        # De-dup GVT: keep one row per container (latest enter_facility if available)
        gvt_pop = (gvt_pop.sort_values('enter_facility', na_position='last')
                          .groupby('container').last().reset_index())
    gvt_facility = (gvt_pop.set_index('container')['facility']
                    if not gvt_pop.empty else pd.Series(dtype=str))

    # ── OBLT milestone events (full file — not week-restricted) ─────────────────
    av_ev = _latest(oblt, 'AV')
    rd_ev = _latest(oblt, 'RD')   # EVENT_511 — empty container return to terminal
    vd_ev = _latest(oblt, 'VD')   # Vessel Departure
    rpt   = pd.Timestamp(report_date)

    # Supplement OBLT AV with GVT Terminal Avail for containers missing an AV record
    if not gvt_pop.empty and 'terminal_avail' in gvt_pop.columns:
        ta = gvt_pop.set_index('container')['terminal_avail'].dropna()
        missing_av = ta[~ta.index.isin(av_ev.index)]
        if not missing_av.empty:
            av_ev = pd.concat([av_ev, missing_av])

    # ── AV→OA ──────────────────────────────────────────────────────────────────
    # Denominator: containers with OA + containers with AV expired (>3d, no OA)
    # Avg/P90: only containers with both AV+OA (computable delta)
    ao_days = []; ao_met = 0; ao_den = 0
    for c in pop_ctrs:
        hav = c in av_ev.index; hoa = c in oa_ev.index
        if hoa and hav:
            d = (oa_ev[c] - av_ev[c]).total_seconds() / 86400
            if d > 0: ao_days.append(d)  # skip bad OBLT rows where OA < AV
            ao_den += 1
            if d <= AV_OA_SLA_DAYS: ao_met += 1
        elif hoa and not hav:
            # OA present, no AV recorded — in denominator, count as SLA met
            ao_den += 1; ao_met += 1
        elif hav and not hoa:
            # Expired = AV > SLA_DAYS before report_date with no OA — per spec
            if (rpt - av_ev[c]).total_seconds() / 86400 > AV_OA_SLA_DAYS:
                ao_den += 1  # expired without OA = miss

    # ── OA→Del (BOS only, excl A320-RBTCS) ────────────────────────────────────
    pop_bos_ctrs = {c for c in pop_ctrs if gvt_facility.get(c, '') != EXCL_FACILITY}
    od_days = []; od_met = 0; od_den = 0
    for c in pop_bos_ctrs:
        if c not in oa_ev.index: continue
        oa = oa_ev[c]
        d = (rd_ev[c] - oa).total_seconds()/86400 if c in rd_ev.index else (rpt - oa).total_seconds()/86400
        if d > 0: od_days.append(d)  # skip bad rows where delivery < outgate
        od_den += 1
        if d <= OA_DEL_SLA_DAYS: od_met += 1

    # ── Empty→Term: OBLT EVENT_602 (OA) → EVENT_511 (RD), ALL markets ──────────
    # Completed only: both OA and RD must be present. Do NOT use GVT fields.
    et_met = 0; et_den = 0
    for c in pop_ctrs:
        if c not in oa_ev.index or c not in rd_ev.index: continue
        et_den += 1
        d = (rd_ev[c] - oa_ev[c]).total_seconds() / 86400
        if d <= EMPTY_TERM_SLA: et_met += 1

    # ── E2E: OBLT VD → GVT Enter Facility, completed only ──────────────────────
    gvt_ef = (gvt_pop.set_index('container')['enter_facility']
              if not gvt_pop.empty and 'enter_facility' in gvt_pop.columns
              else pd.Series(dtype='datetime64[ns]'))
    e2e = []
    for c in pop_ctrs:
        if c not in vd_ev.index: continue
        ef = gvt_ef.get(c)
        if pd.isna(ef): continue
        d = (ef - vd_ev[c]).total_seconds() / 86400
        if d > 0:
            e2e.append(d)

    # ── OTP: GVT Current Dray SLA + Enter Facility (primary); IL fallback ───────
    otp_met = 0; otp_den = 0
    if (not gvt_pop.empty
            and 'current_dray_sla' in gvt_pop.columns
            and gvt_pop['current_dray_sla'].notna().any()):
        scored = gvt_pop[
            gvt_pop['enter_facility'].notna() & gvt_pop['current_dray_sla'].notna()
        ].copy()
        if not scored.empty:
            scored['del_date'] = scored['enter_facility'].dt.normalize()
            scored['sla_date'] = scored['current_dray_sla'].dt.normalize()
            otp_met = int((scored['del_date'] <= scored['sla_date']).sum())
            otp_den = len(scored)
    if otp_den == 0 and inbound is not None and not inbound.empty:
        il = inbound[inbound['container'].isin(pop_ctrs)].dropna(
            subset=['po_promised', 'actual_arrival'])
        otp_met = int((il['actual_arrival'] <= il['po_promised']).sum())
        otp_den = len(il)

    def _pct(m, d): return round(m/d*100) if d else None
    def _avg(lst):  return round(float(np.mean(lst)), 1) if lst else None
    def _p90(lst):  return math.ceil(float(np.percentile(lst, 90))) if lst else None

    return {
        'containers':     n,
        'av_oa_avg':      _avg(ao_days),
        'av_oa_sla_pct':  _pct(ao_met, ao_den),
        'av_oa_p90':      _p90(ao_days),
        'oa_del_avg':     _avg(od_days),
        'oa_del_sla_pct': _pct(od_met, od_den),
        'oa_del_p90':     _p90(od_days),
        'empty_term_pct': _pct(et_met, et_den) if et_den else None,  # None = no CE/RTP data → shown as – on slide
        'e2e_avg':        round(float(np.mean(e2e))) if e2e else None,
        'otp_pct':        _pct(otp_met, otp_den),
        'week_start':     week_start.isoformat(),
        'week_end':       week_end.isoformat(),
    }



def compute_carrier_scorecard(gvt, oblt, week_start, week_end, report_date) -> list:
    """
    Per-carrier AV->OA and OA->Del breakdown for the Enhanced WBR carrier scorecard.
    Population: OBLT EVENT_602 (OA) date in week window.
    Carrier assignment: GVT scac lookup by container.
    Returns list of dicts sorted by volume descending.
    """
    # ── Population: OBLT OA-based (EVENT_602) ───────────────────────────────
    oa_wk = oblt[
        (oblt['status'] == 'OA') &
        oblt['date'].notna() &
        (oblt['date'].dt.date >= week_start) &
        (oblt['date'].dt.date <= week_end)
    ].sort_values('date')
    pop_oa = oa_wk.groupby('container').last().reset_index()
    pop_ctrs = set(pop_oa['container'])
    if not pop_ctrs:
        return []

    # Week-specific OA timestamps
    oa_ev = pop_oa.set_index('container')['date']

    # GVT for scac and facility lookup
    gvt_pop = gvt[gvt['container'].isin(pop_ctrs)].copy()
    if not gvt_pop.empty:
        gvt_pop = (gvt_pop.sort_values('enter_facility', na_position='last')
                          .groupby('container').last().reset_index())
    gvt_scac     = (gvt_pop.set_index('container')['scac']
                    if not gvt_pop.empty and 'scac' in gvt_pop.columns
                    else pd.Series(dtype=str))
    gvt_facility = (gvt_pop.set_index('container')['facility']
                    if not gvt_pop.empty else pd.Series(dtype=str))

    av_ev = _latest(oblt, 'AV')
    rd_ev = _latest(oblt, 'RD')
    rpt   = pd.Timestamp(report_date)

    # Build per-carrier container groups via GVT scac
    carrier_map: dict = {}
    for c in pop_ctrs:
        scac = gvt_scac.get(c)
        label = str(scac).strip() if scac and not pd.isna(scac) and str(scac).strip() else 'Unknown'
        carrier_map.setdefault(label, set()).add(c)

    scorecard = []
    for carrier_label, ctrs in carrier_map.items():
        # AV→OA
        ao_days = []; ao_met = 0; ao_den = 0
        for c in ctrs:
            hav = c in av_ev.index; hoa = c in oa_ev.index
            if hoa and hav:
                d = (oa_ev[c] - av_ev[c]).total_seconds() / 86400
                if d > 0: ao_days.append(d)
                ao_den += 1
                if d <= AV_OA_SLA_DAYS: ao_met += 1
            elif hoa and not hav:
                ao_den += 1; ao_met += 1
            elif hav and not hoa:
                if (rpt - av_ev[c]).total_seconds() / 86400 > AV_OA_SLA_DAYS:
                    ao_den += 1

        # OA→Del (BOS only, excl A320-RBTCS)
        bos_ctrs = {c for c in ctrs if gvt_facility.get(c, '') != EXCL_FACILITY}
        od_days = []; od_met = 0; od_den = 0
        for c in bos_ctrs:
            if c not in oa_ev.index: continue
            oa = oa_ev[c]
            d = (rd_ev[c] - oa).total_seconds()/86400 if c in rd_ev.index else (rpt - oa).total_seconds()/86400
            if d > 0: od_days.append(d)
            od_den += 1
            if d <= OA_DEL_SLA_DAYS: od_met += 1

        scorecard.append({
            'carrier':        carrier_label,
            'volume':         len(ctrs),
            'av_oa_sla_pct':  round(ao_met/ao_den*100) if ao_den else None,
            'av_oa_misses':   (ao_den - ao_met) if ao_den else 0,
            'av_oa_avg':      round(float(np.mean(ao_days)), 1) if ao_days else None,
            'av_oa_p90':      round(float(np.percentile(ao_days, 90))) if ao_days else None,
            'oa_del_sla_pct': round(od_met/od_den*100) if od_den else None,
            'oa_del_misses':  (od_den - od_met) if od_den else 0,
            'oa_del_avg':     round(float(np.mean(od_days)), 1) if od_days else None,
            'oa_del_p90':     round(float(np.percentile(od_days, 90))) if od_days else None,
        })

    return sorted(scorecard, key=lambda x: -x['volume'])

def _empty_metrics() -> dict:
    return {k: None for k in [
        'containers','av_oa_avg','av_oa_sla_pct','av_oa_p90',
        'oa_del_avg','oa_del_sla_pct','oa_del_p90',
        'empty_term_pct','e2e_avg','otp_pct','week_start','week_end'
    ]}


def compute_totals(weeks_data: list) -> dict:
    valid = [w for w in weeks_data if w and w.get('containers')]
    if not valid: return _empty_metrics()
    total_vol = sum(w['containers'] for w in valid)

    def wavg(k):
        vals = [(w[k], w['containers']) for w in valid if w.get(k) is not None]
        return round(sum(v*c for v,c in vals)/sum(c for _,c in vals), 1) if vals else None

    def _raw(k):
        vals = [(w[k], w['containers']) for w in valid if w.get(k) is not None]
        return sum(v*c for v,c in vals)/sum(c for _,c in vals) if vals else None

    def wavgi(k):
        v = _raw(k); return int(v + 0.5) if v is not None else None

    return {
        'containers':     total_vol,
        'av_oa_avg':      wavg('av_oa_avg'),
        'av_oa_sla_pct':  wavgi('av_oa_sla_pct'),
        'av_oa_p90':      wavgi('av_oa_p90'),
        'oa_del_avg':     wavg('oa_del_avg'),
        'oa_del_sla_pct': wavgi('oa_del_sla_pct'),
        'oa_del_p90':     wavgi('oa_del_p90'),
        'empty_term_pct': wavgi('empty_term_pct'),
        'e2e_avg':        wavgi('e2e_avg'),
        'otp_pct':        wavgi('otp_pct'),
        'week_start':     None, 'week_end': None,
    }


WBR_SCHEMA = '''
CREATE TABLE IF NOT EXISTS wbr_results (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    year           INTEGER NOT NULL,
    week_num       INTEGER NOT NULL,
    week_label     TEXT,
    week_start     TEXT,
    week_end       TEXT,
    containers     INTEGER,
    av_oa_avg      REAL,
    av_oa_sla_pct  INTEGER,
    av_oa_p90      INTEGER,
    oa_del_avg     REAL,
    oa_del_sla_pct INTEGER,
    oa_del_p90     INTEGER,
    empty_term_pct INTEGER,
    e2e_avg        INTEGER,
    otp_pct        INTEGER,
    generated_at   TEXT,
    UNIQUE(year, week_num)
);
'''


def save_week_to_db(conn, year, week_num, m, generated_at):
    conn.execute('''
        INSERT OR REPLACE INTO wbr_results
        (year,week_num,week_label,week_start,week_end,containers,
         av_oa_avg,av_oa_sla_pct,av_oa_p90,
         oa_del_avg,oa_del_sla_pct,oa_del_p90,
         empty_term_pct,e2e_avg,otp_pct,generated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ''', (year, week_num, f'W{week_num}', m.get('week_start'), m.get('week_end'), m.get('containers'),
          m.get('av_oa_avg'), m.get('av_oa_sla_pct'), m.get('av_oa_p90'),
          m.get('oa_del_avg'), m.get('oa_del_sla_pct'), m.get('oa_del_p90'),
          m.get('empty_term_pct'), m.get('e2e_avg'), m.get('otp_pct'), generated_at))
    conn.commit()


def load_weeks_from_db(conn, year, week_nums):
    placeholders = ','.join('?' * len(week_nums))
    # Exclude seed/placeholder rows — only load data from real WBR runs
    rows = conn.execute(
        f'SELECT * FROM wbr_results WHERE year=? AND week_num IN ({placeholders})'
        f' AND (generated_at IS NULL OR generated_at != ?)',
        [year] + list(week_nums) + ['seed']
    ).fetchall()
    result = {}
    for row in rows:
        d = dict(row)
        result[d['week_num']] = {
            'containers':     d['containers'],
            'av_oa_avg':      d['av_oa_avg'],
            'av_oa_sla_pct':  d['av_oa_sla_pct'],
            'av_oa_p90':      d['av_oa_p90'],
            'oa_del_avg':     d['oa_del_avg'],
            'oa_del_sla_pct': d['oa_del_sla_pct'],
            'oa_del_p90':     d['oa_del_p90'],
            'empty_term_pct': d['empty_term_pct'],
            'e2e_avg':        d['e2e_avg'],
            'otp_pct':        d['otp_pct'],
            'week_start':     d.get('week_start'),
            'week_end':       d.get('week_end'),
        }
    return result




def load_market_context(conn, days_back: int = 14) -> dict:
    """
    Load the latest port intel, freight rate, and macro signal data from DB.
    Returns a dict ready for bridge injection.
    Called at WBR generate time — data populated daily by port-intel-scraper Lambda.

    Returns:
        {
          'freight':  str  — Drewry WCI headline (latest)
          'macro':    str  — GSCPI reading (latest)
          'ports': {
              'USSAV': [headlines...],
              'USORF': [headlines...],
              'USLAX': [headlines...],
              'USBOS': [headlines...],
              'ALL':   [ocean freight headlines...],
          },
          'has_data': bool
        }
    """
    from datetime import datetime, timedelta, timezone
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")

    result = {"freight": None, "macro": None, "ports": {}, "has_data": False}

    try:
        # Freight rate — latest Drewry WCI
        row = conn.execute(
            "SELECT headline, summary FROM port_intel "
            "WHERE source='Drewry WCI' ORDER BY scraped_at DESC LIMIT 1"
        ).fetchone()
        if row:
            result["freight"] = row[0]
            result["has_data"] = True

        # Macro signal — latest GSCPI
        row = conn.execute(
            "SELECT value, period FROM macro_signals "
            "WHERE source='NY Fed' AND metric='GSCPI' ORDER BY scraped_at DESC LIMIT 1"
        ).fetchone()
        if row:
            v, period = row
            if v > 1.0:
                level = "significantly elevated"
            elif v > 0.5:
                level = "elevated"
            elif v < -1.0:
                level = "significantly below normal"
            elif v < -0.5:
                level = "below normal"
            else:
                level = "near normal"
            result["macro"] = f"GSCPI {v:+.2f}σ ({period}) — supply chain pressure {level}"
            result["has_data"] = True

        # Port headlines — top 3 per port from past days_back days
        for port in ("USSAV", "USORF", "USLAX", "USBOS", "ALL"):
            rows = conn.execute(
                "SELECT headline, source, published_date FROM port_intel "
                "WHERE port_code=? AND source != 'Drewry WCI' AND scraped_at >= ? "
                "ORDER BY scraped_at DESC LIMIT 3",
                (port, cutoff)
            ).fetchall()
            if rows:
                result["ports"][port] = [
                    {"headline": r[0], "source": r[1], "date": r[2]} for r in rows
                ]
                result["has_data"] = True

    except Exception as e:
        # Table may not exist yet (Lambda not yet deployed) — fail silently
        print(f"load_market_context: {e}")

    return result


def format_market_context_block(ctx: dict) -> str:
    """
    Format market context dict into a plain-text bridge block.
    Returns empty string if no data available.
    """
    if not ctx.get("has_data"):
        return ""

    lines = ["[Market Context]"]

    if ctx.get("freight"):
        lines.append(f"Freight Rates: {ctx['freight']}")

    if ctx.get("macro"):
        lines.append(f"Macro: {ctx['macro']}")

    port_labels = {
        "USSAV": "Savannah (USSAV)",
        "USORF": "Norfolk/Virginia (USORF)",
        "USLAX": "Los Angeles (USLAX)",
        "USBOS": "Boston (USBOS)",
        "ALL":   "Ocean / Industry",
    }

    port_section = []
    for port, label in port_labels.items():
        headlines = ctx["ports"].get(port, [])
        if headlines:
            # Use the most recent headline only for bridge brevity
            h = headlines[0]
            port_section.append(f"  {label}: {h['headline']}")

    if port_section:
        lines.append("Port Intelligence:")
        lines.extend(port_section)

    return "\n".join(lines)


def parse_wbr_pdf(pdf_bytes: bytes):
    """Parse a previously generated WBR PDF and extract the weekly table data.

    Coordinate-independent: scans all text on the page, locates week-column headers
    by pattern matching (e.g. "W25", "W 25", "2026 W 25"), then reads cell values
    by proximity. Works regardless of which app version generated the PDF.

    Returns:
        week_labels : list[str]  – e.g. ["W25","W26","W27","W28","W29"]
        week_data   : dict[int, dict]  – keyed by week_num int, values are metrics dicts
    """
    import fitz
    import re

    ROW_KEYS = [
        "containers",  "av_oa_avg",  "av_oa_sla_pct",  "av_oa_p90",
        "oa_del_avg",  "oa_del_sla_pct",  "oa_del_p90",
        "empty_term_pct", "e2e_avg", "otp_pct",
    ]
    INT_KEYS = {"containers", "av_oa_sla_pct", "av_oa_p90",
                "oa_del_sla_pct", "oa_del_p90",
                "empty_term_pct", "e2e_avg", "otp_pct"}

    doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    PAGE_H = page.rect.height

    # Extract ALL words with bounding boxes: (x0,y0,x1,y1,text,block,line,word)
    all_words = page.get_text("words")
    doc.close()

    if not all_words:
        return [], {}

    words_list = [(w[0], w[1], w[2], w[3], w[4]) for w in all_words]  # x0,y0,x1,y1,text

    # ── Step 1: Find week-column headers by scanning for "W XX" / "WXX" tokens ──
    col_headers = {}   # week_num → x_center
    col_header_ys = {}  # week_num → y0 of the "W" token — tracked in same scan as col_headers
    i = 0
    while i < len(words_list):
        x0, y0, x1, y1, txt = words_list[i]
        # Match "W25" as single token
        m = re.fullmatch(r"W(\d{1,2})", txt.strip(), re.IGNORECASE)
        if m:
            wn = int(m.group(1))
            if 1 <= wn <= 53:
                col_headers[wn] = (x0 + x1) / 2
                col_header_ys[wn] = y0
            i += 1
            continue
        # Match "W" followed by a digit token on the same y-band (e.g. "2026 W 25")
        if txt.strip().upper() == "W" and i + 1 < len(words_list):
            nx0, ny0, nx1, ny1, ntxt = words_list[i + 1]
            if abs(ny0 - y0) < 8:
                try:
                    wn = int(ntxt.strip())
                    if 1 <= wn <= 53:
                        col_headers[wn] = (x0 + nx1) / 2
                        col_header_ys[wn] = y0
                    i += 2
                    continue
                except ValueError:
                    pass
        i += 1

    if not col_headers:
        return [], {}

    # ── Step 2: Determine header row y, then derive all row y-positions ──
    # Use y0 from the SAME scan that built col_headers (last assignment = main table).
    # This avoids the fallback PAGE_H*0.57 which was 4.6pt off, causing every row
    # window to capture the row above instead of the correct one.
    hdr_y = float(np.median(list(col_header_ys.values()))) if col_header_ys else PAGE_H * 0.57

    ROW_H = 16.0  # fixed row height in both app and Quick AI slides
    # row_ys[ri] = y-center of data row ri (ri=0 → Containers, ri=9 → OTP)
    # In fitz coords (y increases downward), rows are below the header
    row_ys = [hdr_y + ROW_H * (ri + 1) for ri in range(len(ROW_KEYS))]

    # ── Step 3: Extract cell values by proximity ──────────────────────────
    # Sort col_headers by x position
    sorted_cols = sorted(col_headers.items(), key=lambda kv: kv[1])  # (wn, x_center)
    week_nums_sorted = [kv[0] for kv in sorted_cols]
    col_xs = [kv[1] for kv in sorted_cols]

    # Column width estimate
    if len(col_xs) > 1:
        col_w = min(abs(col_xs[j+1] - col_xs[j]) for j in range(len(col_xs)-1))
    else:
        col_w = 80.0

    # Row height estimate
    if len(row_ys) > 1:
        row_h = np.median([abs(row_ys[j+1] - row_ys[j]) for j in range(len(row_ys)-1)])
    else:
        row_h = ROW_H_EST

    # Skip "Total" column (typically the rightmost; x > last data col + col_w*0.6)
    # It's already excluded since we only captured "W XX" headers (not "Total")

    week_data = {}
    for ri, key in enumerate(ROW_KEYS):
        ry = row_ys[ri]
        for ci, wn in enumerate(week_nums_sorted):
            cx = col_xs[ci]
            # Clip rect centered on (cx, ry) with ±col_w/2 x, ±row_h/2 y
            half_w = col_w * 0.48
            half_h = row_h * 0.55
            cell_words = [
                (x0, y0, x1, y1, txt)
                for x0, y0, x1, y1, txt in words_list
                if abs((x0+x1)/2 - cx) < half_w and abs((y0+y1)/2 - ry) < half_h
            ]
            if not cell_words:
                continue
            # Join all text in cell, strip %, dashes
            raw = " ".join(w[4] for w in sorted(cell_words, key=lambda w: w[0]))
            raw = raw.replace("%", "").replace("\u2013", "").replace("\u2014", "").replace("–", "").replace("—", "").strip()
            if not raw:
                continue
            try:
                val = float(raw)
                if wn not in week_data:
                    week_data[wn] = {}
                week_data[wn][key] = int(round(val)) if key in INT_KEYS else round(val, 1)
            except ValueError:
                pass

    week_labels = [f"W{wn}" for wn in week_nums_sorted]
    return week_labels, week_data
