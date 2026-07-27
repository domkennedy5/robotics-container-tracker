"""
wbr_engine.py — WBR metric calculations for Robotics Destination Dray

Methodology (locked against WK29 validation):
- Population:  GVT containers with Ready Date/Time in week Sun–Sat
- AV->OA:      OBLT latest AV/OA; measurable = has_OA OR (has_AV > 3 frac days before report)
- OA->Del:     OBLT OA->RD; BOS-only (excludes A320-RBTCS ORF facility); in-transit = elapsed
- Empty->Term: GVT Container Empty -> Return to Port; completed only; SLA <= 3 days
- E2E:         OBLT VD -> GVT Enter Facility; completed only
- OTP:         Inbound Loads delivery_date <= PO Promised Date; delivered containers only
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
    pop = gvt[
        gvt['ready_date'].notna() &
        (gvt['ready_date'] >= week_start) &
        (gvt['ready_date'] <= week_end)
    ].copy()

    n = len(pop)
    if n == 0: return _empty_metrics()

    av_ev = _latest(oblt, 'AV')
    oa_ev = _latest(oblt, 'OA')
    rd_ev = _latest(oblt, 'RD')
    vd_ev = _latest(oblt, 'VD')
    rpt   = pd.Timestamp(report_date)

    # AV->OA
    # Denominator: containers with OA (regardless of AV) + containers with AV expired (>3d, no OA)
    # Avg/P90: only containers with both AV+OA (computable delta)
    # SLA met: OA-AV <= 3 days (if both present) OR has OA but no AV (treat as met, delta unknown)
    ao_days = []; ao_met = 0; ao_den = 0
    for c in pop['container']:
        hav = c in av_ev.index; hoa = c in oa_ev.index
        if hoa and hav:
            d = (oa_ev[c] - av_ev[c]).total_seconds() / 86400
            if d > 0: ao_days.append(d)  # skip bad OBLT rows where OA < AV
            ao_den += 1
            if d <= AV_OA_SLA_DAYS: ao_met += 1
        elif hoa and not hav:
            # Has OA but no AV recorded — in denominator, count as SLA met (arrived, OA happened)
            ao_den += 1; ao_met += 1
        elif hav and not hoa:
            # Use week_end (not report_date) as expiry threshold: a container
            # available before Saturday close is in-scope regardless of when we run.
            # Expired = AV > SLA_DAYS before report_date (not week_end) — per spec
            if (rpt - av_ev[c]).total_seconds() / 86400 > AV_OA_SLA_DAYS:
                ao_den += 1  # expired without OA = miss

    # OA->Del (BOS only, excl A320-RBTCS)
    pop_bos = pop[pop['facility'] != EXCL_FACILITY]
    od_days = []; od_met = 0; od_den = 0
    for c in pop_bos['container']:
        if c not in oa_ev.index: continue
        oa = oa_ev[c]
        d = (rd_ev[c] - oa).total_seconds()/86400 if c in rd_ev.index else (rpt - oa).total_seconds()/86400
        if d > 0: od_days.append(d)  # skip bad rows where delivery < outgate
        od_den += 1
        if d <= OA_DEL_SLA_DAYS: od_met += 1

    # Empty->Term (completed only)
    et_met = 0; et_den = 0
    for _, r in pop.iterrows():
        ce, rp = r['container_empty'], r['return_to_port']
        if pd.isna(ce) or pd.isna(rp): continue
        et_den += 1
        if (rp - ce).total_seconds()/86400 <= EMPTY_TERM_SLA: et_met += 1

    # E2E (VD->Enter Facility)
    # Completed containers use actual Enter Facility date.
    # In-transit containers that have voyaged 40-75 days (same voyage window
    # as completed ones) use report_date as proxy.  Containers with recent
    # VD dates (still at sea) are excluded — their voyage isn't measurable yet.
    e2e = []
    for _, r in pop.iterrows():
        c, ef = r['container'], r['enter_facility']
        if c not in vd_ev.index: continue
        vd = vd_ev[c]
        days_since_vd = (rpt - vd).total_seconds() / 86400
        if not pd.isna(ef):
            d = (ef - vd).total_seconds() / 86400
            if d > 0: e2e.append(d)
        elif 40 <= days_since_vd <= 90:
            # In-transit but voyaged long enough to be measurable
            e2e.append(days_since_vd)

    # OTP
    il = inbound[inbound['container'].isin(set(pop['container']))].dropna(subset=['po_promised','actual_arrival'])
    otp_met = (il['actual_arrival'] <= il['po_promised']).sum()
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
    Returns list of dicts sorted by volume descending.
    """
    pop = gvt[
        gvt['ready_date'].notna() &
        (gvt['ready_date'] >= week_start) &
        (gvt['ready_date'] <= week_end)
    ].copy()

    if pop.empty:
        return []

    av_ev     = _latest(oblt, 'AV')
    oa_ev     = _latest(oblt, 'OA')
    rd_ev     = _latest(oblt, 'RD')
    rpt       = pd.Timestamp(report_date)
    wk_end_ts = pd.Timestamp(week_end)

    scorecard = []
    for carrier, grp in pop.groupby('scac'):
        carrier_label = str(carrier).strip() if carrier else 'Unknown'

        # AV→OA
        ao_days = []; ao_met = 0; ao_den = 0
        for c in grp['container']:
            hav = c in av_ev.index; hoa = c in oa_ev.index
            if hoa and hav:
                d = (oa_ev[c] - av_ev[c]).total_seconds() / 86400
                if d > 0: ao_days.append(d)  # skip bad OBLT rows where OA < AV
                ao_den += 1
                if d <= AV_OA_SLA_DAYS: ao_met += 1
            elif hoa and not hav:
                ao_den += 1; ao_met += 1
            elif hav and not hoa:
                # Expired = AV > SLA_DAYS before report_date — per spec
                if (rpt - av_ev[c]).total_seconds() / 86400 > AV_OA_SLA_DAYS:
                    ao_den += 1  # expired miss

        # OA→Del (BOS only, excl A320-RBTCS)
        bos_grp = grp[grp['facility'] != EXCL_FACILITY]
        od_days = []; od_met = 0; od_den = 0
        for c in bos_grp['container']:
            if c not in oa_ev.index: continue
            oa = oa_ev[c]
            d = (rd_ev[c] - oa).total_seconds()/86400 if c in rd_ev.index else (rpt - oa).total_seconds()/86400
            if d > 0: od_days.append(d)  # skip bad rows where delivery < outgate
            od_den += 1
            if d <= OA_DEL_SLA_DAYS: od_met += 1

        scorecard.append({
            'carrier':        carrier_label,
            'volume':         len(grp),
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

    # Row label patterns — matched against full left-side line text.
    # The PDF uses labels like "AV→OA Avg (Days)", "OA→Del SLA %", etc.
    # Use simple substrings / OR patterns that are unique to each row.
    ROW_LABEL_PATTERNS = [
        r"Container",           # Containers
        r"AV.*Avg",             # AV→OA Avg (Days)
        r"AV.*SLA",             # AV→OA SLA %
        r"AV.*P90",             # AV→OA P90 (Days)
        r"OA.*Del.*Avg|Del.*Avg",  # OA→Del Avg (Days)
        r"OA.*Del.*SLA|Del.*SLA",  # OA→Del SLA %
        r"OA.*Del.*P90|Del.*P90",  # OA→Del P90 (Days)
        r"Empty",               # Empty→Term SLA %
        r"E2E|Transit",         # E2E Transit Avg (Days)
        r"On.Time|OTP",         # On-Time to Promise %
    ]

    doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    PAGE_W = page.rect.width
    PAGE_H = page.rect.height

    # Extract ALL words with bounding boxes: (x0,y0,x1,y1,text,block,line,word)
    all_words = page.get_text("words")
    doc.close()

    if not all_words:
        return [], {}

    # ── Step 1: Find week-column headers ─────────────────────────────────
    # Collect words that look like week numbers ("25","26"... or "W25","W 25", "2026")
    # Strategy: find occurrences of 2-digit numbers 1..53 that appear in the header band.
    # The header band is the row containing "W XX" patterns — find it by y-coord of
    # any word immediately preceded or followed by "W".

    # Build a word-by-word pass: look for pattern "W" + integer, or "2026" + "W" + integer
    col_headers = {}   # week_num → x_center
    i = 0
    words_list = [(w[0], w[1], w[2], w[3], w[4]) for w in all_words]  # x0,y0,x1,y1,text
    while i < len(words_list):
        x0, y0, x1, y1, txt = words_list[i]
        # Match "W25" in a single token
        m = re.fullmatch(r"W(\d{1,2})", txt.strip(), re.IGNORECASE)
        if m:
            wn = int(m.group(1))
            if 1 <= wn <= 53:
                col_headers[wn] = (x0 + x1) / 2
            i += 1
            continue
        # Match "W" followed by a digit token on the same y-band
        if txt.strip().upper() == "W" and i + 1 < len(words_list):
            nx0, ny0, nx1, ny1, ntxt = words_list[i + 1]
            if abs(ny0 - y0) < 8:  # same line
                try:
                    wn = int(ntxt.strip())
                    if 1 <= wn <= 53:
                        col_headers[wn] = ((x0 + nx1) / 2)
                    i += 2
                    continue
                except ValueError:
                    pass
        i += 1

    if not col_headers:
        return [], {}

    # Header y-band: median y of all header words
    hdr_y_vals = []
    for x0, y0, x1, y1, txt in words_list:
        m = re.fullmatch(r"W(\d{1,2})", txt.strip(), re.IGNORECASE)
        if m and 1 <= int(m.group(1)) <= 53:
            hdr_y_vals.append((y0 + y1) / 2)
    hdr_y = float(np.median(hdr_y_vals)) if hdr_y_vals else PAGE_H * 0.57

    # ── Step 2: Find row y-positions from row labels on the LEFT side ────
    # Row labels are to the left of the table (x < 200 typically)
    left_words = [(x0, y0, x1, y1, txt) for x0, y0, x1, y1, txt in words_list if x1 < 220]

    # ── Step 2: Locate row y-positions ──────────────────────────────────
    # Build full-line text by grouping left-side words on same y-band (within 4pt)
    # so multi-word labels like "AV→OA Avg (Days)" can be matched as one string.
    ROW_H_EST = 16.0
    from collections import defaultdict as _ddict
    _line_words = _ddict(list)
    for x0, y0, x1, y1, txt in left_words:
        y_key = round((y0 + y1) / 2 / 2) * 2   # bucket to nearest 2pt
        _line_words[y_key].append((x0, txt))
    full_lines = {yk: " ".join(t for _, t in sorted(items)) for yk, items in _line_words.items()}

    def _find_row_y(pattern):
        """Return y-center of the left-side label line matching the regex."""
        rx = re.compile(pattern, re.IGNORECASE)
        for y_key, line_txt in full_lines.items():
            if rx.search(line_txt):
                return float(y_key)
        return None

    # Try to locate each row label; evenly-spaced fallback from header y
    row_ys = []
    for pi, pat in enumerate(ROW_LABEL_PATTERNS):
        y = _find_row_y(pat)
        if y is None:
            y = hdr_y + ROW_H_EST * (pi + 1)
        row_ys.append(y)

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
