"""
wbr_pdf.py — WBR PDF slide generator.
Pixel-faithful match to GLS_Robotics_2026-7-20 (3).pdf (forensic analysis 2026-07-24).

Forensic spec (all coords in ReportLab pts, y=0 at bottom):
  Page:        792x612 pts (landscape letter), white background
  Header bg:   dark rect(0, 567, 792, 45) fill=(0.2,0.2,0.2)
  Title:       Helvetica-Bold 24pt, WHITE, baseline y=606, x=30
  Separator:   #e8a838 lw=3.0, y=567 full width
  SLA box:     x=25, y_bot=482, w=200, h=75, lw=0.7
  Charts R1:   bands (268-405.3, 445.3-582.7, 622.7-760.0), y 449-544
  Charts R2:   bands (53-262, 302-511, 551-760), y 306-401
  Table:       label col x=51 w=130; 6 week cols w=80; total col w=80 bg=#fff2cc
  Footer:      Helvetica 7.5pt #666666, y=13
"""
from __future__ import annotations
import io, math
from datetime import date, datetime
from typing import Optional

from reportlab.lib import colors
from reportlab.pdfgen import canvas as rl_canvas

# ── Colors (forensic) ──────────────────────────────────────────────────────
C_BLACK      = colors.black
C_GRAY       = colors.Color(0.400, 0.400, 0.400)   # #666666
C_DOT        = colors.Color(0.294, 0.675, 0.776)   # #4BACC6 teal
C_ACCENT     = colors.Color(0.910, 0.659, 0.220)   # #e8a838 orange
C_TOTAL_BG   = colors.Color(1.000, 0.950, 0.800)   # #fff2cc
C_WHITE      = colors.white
C_HEADER_BG  = colors.Color(0.200, 0.200, 0.200)   # dark header
C_GRID       = colors.Color(0.875, 0.875, 0.875)   # #dfdfdf gridlines
C_CELL_BDR   = colors.black
C_RED_SLA    = colors.Color(1.000, 0.000, 0.000)   # #ff0000
C_GREEN_SLA  = colors.Color(0.000, 0.686, 0.310)   # #00af4f

PAGE_W, PAGE_H = 792.0, 612.0

# ── Chart geometry ──────────────────────────────────────────────────────────
# Row 1: right-shifted (SLA box occupies upper-left)
R1_BANDS    = [(268.0, 405.3), (445.3, 582.7), (622.7, 760.0)]
R1_DOT_SPC  = 27.5
R1_BOT      = 449.0
R1_TOP      = 544.0
R1_TITLE_Y  = 563.0

# Row 2: full-width
R2_BANDS    = [(53.0, 262.0), (302.0, 511.0), (551.0, 760.0)]
R2_DOT_SPC  = 41.8
R2_BOT      = 306.0
R2_TOP      = 401.0
R2_TITLE_Y  = 420.0

PLOT_H = 95.0   # R*_TOP - R*_BOT

# Fixed y-axis ceilings per metric key (None = auto-scale)
CHART_YMAX = {
    "av_oa_avg":  10,
    "oa_del_avg":  8,
    "e2e_avg":    70,
    "otp_pct":   120,
}

# ── Table geometry ──────────────────────────────────────────────────────────
TBL_LABEL_X     =  51.0
TBL_LABEL_W     = 130.0
TBL_WEEK_W      =  80.0
TBL_TOTAL_W     =  80.0
TBL_ROW_H       =  16.0
TBL_COL_HDR_TOP = 260.0   # y-top of column header row

# ── Row definitions ──────────────────────────────────────────────────────────
ROW_LABELS = [
    "Containers",
    "AV\u2192OA Avg (Days)",
    "AV\u2192OA SLA %",
    "AV\u2192OA P90 (Days)",
    "OA\u2192Del Avg (Days)",
    "OA\u2192Del SLA %",
    "OA\u2192Del P90 (Days)",
    "Empty\u2192Term SLA %",
    "E2E Transit Avg (Days)",
    "On-Time to Promise %",
]
ROW_KEYS = [
    "containers", "av_oa_avg", "av_oa_sla_pct", "av_oa_p90",
    "oa_del_avg",  "oa_del_sla_pct",  "oa_del_p90",
    "empty_term_pct", "e2e_avg", "otp_pct",
]
# Row indices (0-based) that get SLA pass/fail colored inset border on current week
SLA_ROWS   = {2, 5, 7, 9}   # AV→OA SLA, OA→Del SLA, Empty→Term SLA, OTP
SLA_TARGET = 95


def _fmt(v, key: str) -> str:
    if v is None:
        return "\u2013"
    if key in ("av_oa_sla_pct", "oa_del_sla_pct", "empty_term_pct", "otp_pct"):
        return f"{int(round(float(v)))}%"
    if key in ("av_oa_avg", "oa_del_avg", "e2e_avg"):
        return f"{float(v):.1f}"
    return str(int(round(float(v))))


def _nice_max(raw_max: float, n_intervals: int = 5) -> float:
    """Round raw_max * 1.2 up to a value divisible by n_intervals."""
    target = raw_max * 1.20
    if target <= 0:
        return float(n_intervals)
    magnitude = 10 ** math.floor(math.log10(target))
    for step in [1, 2, 2.5, 5, 10]:
        candidate = math.ceil(target / (magnitude * step)) * magnitude * step
        if candidate >= target:
            return candidate
    return target


def _chart_vals(weeks_data, key):
    return [w.get(key) if w else None for w in weeks_data]


# ── SLA Goals box ────────────────────────────────────────────────────────────
def _draw_sla_goals(c: rl_canvas.Canvas):
    """Forensic: x=25, y_bot=482, w=200, h=75, lw=0.7."""
    bx, by, bw, bh = 25.0, 482.0, 200.0, 75.0

    c.setStrokeColor(C_CELL_BDR)
    c.setLineWidth(0.7)
    c.setFillColor(C_WHITE)
    c.rect(bx, by, bw, bh, fill=1, stroke=1)

    # "SLA Goals" title — Helvetica-Bold 10pt at y=553.7
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 10.0)
    c.drawString(bx + 8, 553.7, "SLA Goals")

    # Underline below title at y=541
    c.setStrokeColor(C_BLACK)
    c.setLineWidth(0.5)
    c.line(33, 541, 220, 541)

    # Bullet rows — Helvetica 9pt at y=536.7, 522.7, 508.7, 494.7
    goals = [
        "\u2022 AV to OA <= 3 Days",
        "\u2022 OA to Delivery <= 3 Days",
        "\u2022 Empty to Termination <= 3 Days",
        "\u2022 E2E Transit / On-Time to Promise >= 95%",
    ]
    c.setFont("Helvetica", 9.0)
    for g, rl_y in zip(goals, [536.7, 522.7, 508.7, 494.7]):
        c.drawString(bx + 8, rl_y, g)


# ── Single line chart ────────────────────────────────────────────────────────
def _draw_chart(
    c: rl_canvas.Canvas,
    band_idx: int,        # 0, 1, or 2
    row: int,             # 1 or 2
    title: str,
    key: str,             # metric key for CHART_YMAX lookup
    values: list,
    week_labels: list,
    year_2d: str = "26",
    is_pct: bool = False,
):
    bands    = R1_BANDS   if row == 1 else R2_BANDS
    dot_spc  = R1_DOT_SPC if row == 1 else R2_DOT_SPC
    plot_bot = R1_BOT     if row == 1 else R2_BOT
    plot_top = R1_TOP     if row == 1 else R2_TOP
    title_y  = R1_TITLE_Y if row == 1 else R2_TITLE_Y

    x_start, x_end = bands[band_idx]
    cx = (x_start + x_end) / 2
    n  = len(values)
    dot_xs = [x_start + i * dot_spc for i in range(n)]

    # ── Y scale ──────────────────────────────────────────────────────────────
    clean   = [float(v) for v in values if v is not None]
    raw_max = max(clean) if clean else 1.0
    if raw_max == 0:
        raw_max = 1.0
    fixed = CHART_YMAX.get(key)
    if fixed is not None:
        max_v = float(fixed)
        # expand only if data actually exceeds the fixed ceiling
        if clean and max(clean) > max_v:
            max_v = _nice_max(max(clean))
    else:
        max_v = _nice_max(raw_max)

    def _y(v):
        if v is None:
            return None
        return plot_bot + (float(v) / max_v) * PLOT_H

    # ── Chart title (Helvetica-Bold 9pt) ──────────────────────────────────
    c.setFont("Helvetica-Bold", 9.0)
    c.setFillColor(C_BLACK)
    c.drawCentredString(cx, title_y, title)

    # ── Gridlines (6 lines, 5 intervals, lw=0.40) ─────────────────────────
    c.setStrokeColor(C_GRID)
    c.setLineWidth(0.40)
    for i in range(6):
        gy = plot_bot + i * (PLOT_H / 5)
        c.line(x_start, gy, x_end, gy)

    # ── Y-axis value labels (Helvetica 6pt gray, right-aligned, above gridline) ──
    c.setFont("Helvetica", 6.0)
    c.setFillColor(C_GRAY)
    for i in range(6):
        gy    = plot_bot + i * (PLOT_H / 5)
        v_lbl = max_v * i / 5
        lbl   = str(int(v_lbl)) if v_lbl == int(v_lbl) else f"{v_lbl:.1f}"
        c.drawRightString(x_start - 2, gy + 4.4, lbl)

    # ── Connecting line (lw=2.2 teal) ────────────────────────────────────
    c.setStrokeColor(C_DOT)
    c.setLineWidth(2.2)
    pts  = [(dot_xs[i], _y(values[i])) for i in range(n)]
    prev = None
    for px, py in pts:
        if py is not None:
            if prev is not None:
                c.line(prev[0], prev[1], px, py)
            prev = (px, py)
        else:
            prev = None

    # ── Dots (circles radius=3.5, teal) + value labels ───────────────────
    for i in range(n):
        px = dot_xs[i]
        py = _y(values[i])
        v  = values[i]

        # X-axis label
        wk   = week_labels[i] if i < len(week_labels) else ""
        xlbl = f"{year_2d} W {wk[1:]}" if wk.startswith("W") else wk
        c.setFont("Helvetica", 6.0)
        c.setFillColor(C_GRAY)
        c.drawCentredString(px, plot_bot + 15, xlbl)

        if py is not None and v is not None:
            # Filled circle
            c.setFillColor(C_DOT)
            c.setStrokeColor(C_DOT)
            c.circle(px, py, 3.5, fill=1, stroke=0)

            # Value label above dot (Helvetica-Bold 8pt)
            c.setFillColor(C_BLACK)
            c.setFont("Helvetica-Bold", 8.0)
            lbl = f"{int(round(float(v)))}" if is_pct else str(round(float(v), 1) if float(v) != int(float(v)) else int(float(v)))
            c.drawCentredString(px, py + 15.5, lbl)


# ── Performance table ────────────────────────────────────────────────────────
def _draw_table(
    c: rl_canvas.Canvas,
    week_labels: list,
    weeks_data: list,
    totals: Optional[dict],
    year_str: str = "2026",
    sites_str: str = "USORF, USBOS, USSAV, USLAX",
):
    n_weeks     = len(week_labels)
    table_right = TBL_LABEL_X + TBL_LABEL_W + n_weeks * TBL_WEEK_W + TBL_TOTAL_W

    # ── Section header — "Weekly Performance" Helvetica-Bold 9pt at y=283 ──
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 9.0)
    tbl_center = (TBL_LABEL_X + table_right) / 2
    c.drawCentredString(tbl_center, 283, f"Weekly Performance ({sites_str})")

    # ── Column header row ─────────────────────────────────────────────────
    col_hdr_bot = TBL_COL_HDR_TOP - TBL_ROW_H   # = 244.0
    hdr_txt_y   = col_hdr_bot + TBL_ROW_H - 1.4  # text baseline

    # Label header cell
    c.setFillColor(C_WHITE)
    c.setStrokeColor(C_CELL_BDR)
    c.setLineWidth(0.5)
    c.rect(TBL_LABEL_X, col_hdr_bot, TBL_LABEL_W, TBL_ROW_H, fill=1, stroke=1)
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 9.0)
    c.drawCentredString(TBL_LABEL_X + TBL_LABEL_W / 2, hdr_txt_y, "Week")

    # Week header cells
    col_x = TBL_LABEL_X + TBL_LABEL_W
    for wk in week_labels:
        wk_display = f"{year_str} W {wk[1:]}" if wk.startswith("W") else wk
        c.setFillColor(C_WHITE)
        c.setStrokeColor(C_CELL_BDR)
        c.setLineWidth(0.5)
        c.rect(col_x, col_hdr_bot, TBL_WEEK_W, TBL_ROW_H, fill=1, stroke=1)
        c.setFillColor(C_BLACK)
        c.setFont("Helvetica-Bold", 9.0)
        c.drawCentredString(col_x + TBL_WEEK_W / 2, hdr_txt_y, wk_display)
        col_x += TBL_WEEK_W

    # Total header cell (yellow bg)
    c.setFillColor(C_TOTAL_BG)
    c.setStrokeColor(C_CELL_BDR)
    c.setLineWidth(0.5)
    c.rect(col_x, col_hdr_bot, TBL_TOTAL_W, TBL_ROW_H, fill=1, stroke=1)
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 9.0)
    c.drawCentredString(col_x + TBL_TOTAL_W / 2, hdr_txt_y, "Total")

    # ── Data rows ─────────────────────────────────────────────────────────
    last_wk_idx = n_weeks - 1
    for ri, (label, key) in enumerate(zip(ROW_LABELS, ROW_KEYS)):
        row_top = col_hdr_bot - ri * TBL_ROW_H
        row_bot = row_top - TBL_ROW_H
        txt_y   = row_bot + TBL_ROW_H - 1.4   # text baseline

        # Label cell
        c.setFillColor(C_WHITE)
        c.setStrokeColor(C_CELL_BDR)
        c.setLineWidth(0.5)
        c.rect(TBL_LABEL_X, row_bot, TBL_LABEL_W, TBL_ROW_H, fill=1, stroke=1)
        c.setFillColor(C_BLACK)
        c.setFont("Helvetica-Bold", 9.0)
        c.drawString(TBL_LABEL_X + 4, txt_y, label)

        # Week data cells
        col_x = TBL_LABEL_X + TBL_LABEL_W
        for wi, wk_data in enumerate(weeks_data):
            v   = wk_data.get(key) if wk_data else None
            txt = _fmt(v, key)
            is_last = (wi == last_wk_idx)

            # Base cell (white fill, thin black border)
            c.setFillColor(C_WHITE)
            c.setStrokeColor(C_CELL_BDR)
            c.setLineWidth(0.5)
            c.rect(col_x, row_bot, TBL_WEEK_W, TBL_ROW_H, fill=1, stroke=1)

            # Inset SLA border on current-week SLA rows
            if ri in SLA_ROWS and is_last and v is not None:
                border_c = C_GREEN_SLA if float(v) >= SLA_TARGET else C_RED_SLA
                c.setStrokeColor(border_c)
                c.setLineWidth(2.2)
                c.rect(col_x + 1, row_bot + 1, TBL_WEEK_W - 2, TBL_ROW_H - 2,
                       fill=0, stroke=1)

            c.setFillColor(C_BLACK)
            c.setFont("Helvetica-Bold" if is_last else "Helvetica", 9.0)
            c.drawCentredString(col_x + TBL_WEEK_W / 2, txt_y, txt)
            col_x += TBL_WEEK_W

        # Total cell (yellow bg)
        tv  = totals.get(key) if totals else None
        ttx = _fmt(tv, key)
        c.setFillColor(C_TOTAL_BG)
        c.setStrokeColor(C_CELL_BDR)
        c.setLineWidth(0.5)
        c.rect(col_x, row_bot, TBL_TOTAL_W, TBL_ROW_H, fill=1, stroke=1)
        c.setFillColor(C_BLACK)
        c.setFont("Helvetica-Bold", 9.0)
        c.drawCentredString(col_x + TBL_TOTAL_W / 2, txt_y, ttx)


# ── Main generator ───────────────────────────────────────────────────────────
def generate_standard_wbr(
    week_labels: list,
    weeks_data: list,
    totals: Optional[dict],
    report_date: date,
    current_week_label: str = "",
    report_time: str = "10:00am (CT) | 9:00am (MT) | 8:00am (PT)",
    sites_str: str = "USORF, USBOS, USSAV, USLAX",
) -> bytes:
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    year_str = str(report_date.year)
    year_2d  = year_str[2:]

    # ── Dark header background (y=567-612) ───────────────────────────────
    c.setFillColor(C_HEADER_BG)
    c.rect(0, 567, PAGE_W, 45, fill=1, stroke=0)

    # ── Title — Helvetica-Bold 24pt WHITE at y=606, x=30 ─────────────────
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(30, 606, "Robotics Destination Dray Metrics \u2014 NA")

    # ── Orange separator line — full width, lw=3.0 at y=567 ──────────────
    c.setStrokeColor(C_ACCENT)
    c.setLineWidth(3.0)
    c.line(0, 567, PAGE_W, 567)

    # ── SLA Goals box ────────────────────────────────────────────────────
    _draw_sla_goals(c)

    # ── Charts Row 1 (AV→OA | OA→Del | E2E Transit) ─────────────────────
    charts_r1 = [
        ("Leg: AV to OA (avg. days)",       "av_oa_avg",  False),
        ("Leg: OA to Delivery (avg. days)", "oa_del_avg", False),
        ("Leg: E2E Transit (avg. days)",    "e2e_avg",    False),
    ]
    for band_idx, (title, key, is_pct) in enumerate(charts_r1):
        _draw_chart(c, band_idx, 1, title, key,
                    _chart_vals(weeks_data, key), week_labels, year_2d, is_pct)

    # ── Charts Row 2 (Empty→Term | OTP | Volume) ─────────────────────────
    charts_r2 = [
        ("Leg: Empty to Termination (%)",  "empty_term_pct", True),
        ("Leg: On-Time to Promise %",       "otp_pct",        True),
        ("Volume (containers)",             "containers",     False),
    ]
    for band_idx, (title, key, is_pct) in enumerate(charts_r2):
        _draw_chart(c, band_idx, 2, title, key,
                    _chart_vals(weeks_data, key), week_labels, year_2d, is_pct)

    # ── Performance table ────────────────────────────────────────────────
    _draw_table(c, week_labels, weeks_data, totals, year_str, sites_str)

    # ── Footer ───────────────────────────────────────────────────────────
    try:
        date_str = report_date.strftime("%-B %-d, %Y")
    except ValueError:
        date_str = report_date.strftime("%B %d, %Y").replace(" 0", " ")

    c.setFillColor(C_GRAY)
    c.setFont("Helvetica", 7.5)
    c.drawString(30, 13, f"{date_str} {report_time}")
    c.drawCentredString(PAGE_W / 2, 13, "\u25a0 Powered by Amazon Quick")

    c.save()
    return buf.getvalue()
