"""
wbr_pdf.py — WBR PDF slide generator.
Pixel-faithful match to GLS_Robotics_2026-7-20_FINAL.pdf (forensic analysis 2026-07-24).

Forensic spec:
  Page:        792×612 pts (landscape letter), white background
  Title:       Helvetica-Bold 20pt, black, baseline RL y=576, x=30
  Accent line: #e8a838, lw=2, y=568, x=30-762
  SLA box:     x=30-230, RL y=492-556, no fill, black border lw=0.5
  Charts:      3 bands (x=238-382, 428-572, 618-762), 2 rows (RL y=435-522 / 290-377)
               6 gridlines per chart, spacing 17.4; dots 5×5 #4BACC6; line lw=1.5
  Table:       label col x=51 w=130; 6 week cols w=80; total col w=80 bg=#fff2cc
               col hdr RL y=199-215; data rows 16pt each down from 199
  Footer:      Helvetica 7.5pt #666666, RL y=13
"""
from __future__ import annotations
import io, math
from datetime import date, datetime
from typing import Optional

from reportlab.lib import colors
from reportlab.pdfgen import canvas as rl_canvas

# ── Colors (forensic) ─────────────────────────────────────────────────────────
C_BLACK      = colors.black
C_GRAY       = colors.Color(0.400, 0.400, 0.400)   # #666666
C_DOT        = colors.Color(0.294, 0.675, 0.776)   # #4BACC6
C_ACCENT     = colors.Color(0.910, 0.659, 0.220)   # #e8a838 (orange accent line)
C_TOTAL_BG   = colors.Color(1.000, 0.950, 0.800)   # #fff2cc (total col)
C_WHITE      = colors.white
C_GRIDLIGHT  = colors.Color(0.875, 0.875, 0.875)   # #dfdfdf (chart gridlines)
C_CELL_BDR   = colors.black                         # table cell borders
C_RED_SLA    = colors.Color(1.000, 0.000, 0.000)   # #ff0000 (SLA fail)
C_GREEN_SLA  = colors.Color(0.000, 0.686, 0.310)   # #00af4f (SLA pass)

PAGE_W, PAGE_H = 792.0, 612.0

# ── Chart geometry (all in ReportLab coords, y=0 at bottom) ──────────────────
# Plot x bands — row 1 is right-shifted (SLA box occupies upper-left)
# row 2 spans full page width. Derived from gold PDF forensics.
R1_BANDS = [(266, 404), (443, 581), (618, 756)]  # row 1 (shifted right)
R2_BANDS = [(51,  260), (298, 507), (547, 756)]  # row 2 (full width)
R1_DOT_SPACING = 27.5   # row 1: 6 dots spanning ~138pt
R2_DOT_SPACING = 41.8   # row 2: 6 dots spanning ~209pt
# Legacy alias so nothing else breaks
BANDS = R1_BANDS
DOT_SPACING = R1_DOT_SPACING

# Plot y ranges per row
R1_BOT, R1_TOP = 447.0, 542.0   # row 1 charts (fitz y=165→70)
R2_BOT, R2_TOP = 304.0, 399.0   # row 2 charts (fitz y=308→213)

# Chart title y (baseline)
R1_TITLE_Y  = 554.0   # row 1 (fitz y=58)
R2_TITLE_Y  = 411.0   # row 2 (fitz y=201)

# X-axis label y (inside chart near bottom)
R1_XLABEL_Y = 444.0   # fitz y=168
R2_XLABEL_Y = 301.0   # fitz y=307.3

# ── Table geometry ────────────────────────────────────────────────────────────
TBL_LABEL_X   = 51.0    # table left x (label column starts here)
TBL_LABEL_W   = 130.0   # label column width
TBL_WEEK_W    = 80.0    # per-week column width
TBL_TOTAL_W   = 80.0    # total column width
TBL_ROW_H     = 16.0    # row height (pts)
TBL_COL_HDR_TOP = 260.0 # RL y-top of column header row

# ── Row definitions ───────────────────────────────────────────────────────────
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
# Row indices (0-based) that get SLA pass/fail color border on current week
SLA_ROWS = {2, 5, 7}   # AV→OA SLA, OA→Del SLA, Empty→Term SLA
SLA_TARGET = 95


def _fmt(v, key: str) -> str:
    if v is None:
        return "\u2013"
    if key in ("av_oa_sla_pct", "oa_del_sla_pct", "empty_term_pct", "otp_pct"):
        return f"{int(round(float(v)))}%"
    if key in ("av_oa_avg", "oa_del_avg"):
        return f"{float(v):.1f}"
    return str(int(round(float(v))))


def _nice_max(raw_max: float, n_intervals: int = 5) -> float:
    """Round raw_max * 1.2 up to a nice number divisible by n_intervals."""
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


# ── SLA Goals box ─────────────────────────────────────────────────────────────
def _draw_sla_goals(c: rl_canvas.Canvas):
    """Forensic: x=30-265, RL y_bottom=477, h=82, black border lw=0.5."""
    bx, by, bw, bh = 30.0, 477.0, 235.0, 82.0

    c.setStrokeColor(C_CELL_BDR)
    c.setLineWidth(0.5)
    c.setFillColor(C_WHITE)
    c.rect(bx, by, bw, bh, fill=1, stroke=1)

    # "SLA Goals" header — Helvetica-Bold 10pt, fitz y=69 → RL baseline 543
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 10.0)
    c.drawString(bx + 3, 543, "SLA Goals")

    goals = [
        "\u2022 AV to OA <= 3 Days",
        "\u2022 OA to Delivery <= 3 Days",
        "\u2022 Empty to Termination <= 3 Days",
        "\u2022 E2E Transit / On-Time to Promise >= 95%",
    ]
    c.setFont("Helvetica", 9.0)
    # From forensic: bullet rows at fitz y=85,99,113,127 → RL 527,513,499,485
    for g, rl_y in zip(goals, [527, 513, 499, 485]):
        c.drawString(bx + 4, rl_y, g)


# ── Single line chart (dot+line style) ────────────────────────────────────────
def _draw_chart(
    c: rl_canvas.Canvas,
    band_idx: int,         # 0, 1, or 2
    row: int,              # 1 or 2
    title: str,
    values: list,
    week_labels: list,
    year_2d: str = "26",
    is_pct: bool = False,
):
    bands      = R1_BANDS if row == 1 else R2_BANDS
    dot_spc    = R1_DOT_SPACING if row == 1 else R2_DOT_SPACING
    x_start, x_end = bands[band_idx]
    plot_bot = R1_BOT if row == 1 else R2_BOT
    plot_top = R1_TOP if row == 1 else R2_TOP
    title_y  = R1_TITLE_Y if row == 1 else R2_TITLE_Y
    xlabel_y = R1_XLABEL_Y if row == 1 else R2_XLABEL_Y
    plot_h   = plot_top - plot_bot

    # Dot x positions — 6 points, row-specific spacing
    n = len(values)
    dot_xs = [x_start + i * dot_spc for i in range(n)]

    # Y scale
    clean = [float(v) for v in values if v is not None]
    raw_max = max(clean) if clean else 1.0
    if raw_max == 0:
        raw_max = 1.0
    max_v = _nice_max(raw_max)

    def _y(v):
        """Convert data value to RL y coordinate."""
        if v is None:
            return None
        return plot_bot + (float(v) / max_v) * plot_h

    # ── Title ──────────────────────────────────────────────────────────────
    cx = (x_start + x_end) / 2
    c.setFont("Helvetica-Bold", 9.0)
    c.setFillColor(C_BLACK)
    c.drawCentredString(cx, title_y, title)

    # ── Gridlines (6 horizontal, spacing 17.4 pts) ──────────────────────
    c.setStrokeColor(C_GRIDLIGHT)
    c.setLineWidth(0.3)
    for i in range(6):
        gy = plot_bot + i * (plot_h / 5)
        c.line(x_start, gy, x_end, gy)

    # ── Y-axis value labels (Helvetica 5.5pt gray, right-aligned 2pt left of x_start) ──
    c.setFont("Helvetica", 6.0)
    c.setFillColor(C_GRAY)
    for i in range(6):
        gy = plot_bot + i * (plot_h / 5)
        v_label = max_v * i / 5
        # Format: integers for most, skip decimals
        if v_label == int(v_label):
            lbl = str(int(v_label))
        else:
            lbl = f"{v_label:.1f}"
        c.drawRightString(x_start - 2, gy - 2, lbl)

    # ── Connecting line ─────────────────────────────────────────────────
    c.setStrokeColor(C_DOT)
    c.setLineWidth(1.5)
    pts = [(dot_xs[i], _y(values[i])) for i in range(n)]
    prev = None
    for px, py in pts:
        if py is not None:
            if prev is not None:
                c.line(prev[0], prev[1], px, py)
            prev = (px, py)
        else:
            prev = None

    # ── Dots (5×5 squares) + value labels ───────────────────────────────
    DOT_H = 2.5  # half of 5pt square
    for i in range(n):
        px = dot_xs[i]
        py = _y(values[i])
        v  = values[i]

        # X-axis label — Helvetica 5pt gray
        wk = week_labels[i] if i < len(week_labels) else ""
        if wk.startswith("W"):
            xlbl = f"{year_2d} W {wk[1:]}"
        else:
            xlbl = wk
        c.setFont("Helvetica", 6.0)
        c.setFillColor(C_GRAY)
        c.drawCentredString(px, xlabel_y, xlbl)

        if py is not None and v is not None:
            # 5×5 square dot
            c.setFillColor(C_DOT)
            c.setStrokeColor(C_DOT)
            c.rect(px - DOT_H, py - DOT_H, 5, 5, fill=1, stroke=0)

            # Value label above dot — Helvetica-Bold 6.5pt black
            c.setFillColor(C_BLACK)
            c.setFont("Helvetica-Bold", 8.0)
            if is_pct:
                lbl = f"{int(round(float(v)))}"
            else:
                lbl = str(round(float(v)))
            c.drawCentredString(px, py + DOT_H + 3, lbl)


# ── Performance table ─────────────────────────────────────────────────────────
def _draw_table(
    c: rl_canvas.Canvas,
    week_labels: list,
    weeks_data: list,
    totals: Optional[dict],
    year_str: str = "2026",
    sites_str: str = "USORF, USBOS, USSAV, USLAX",
):
    n_weeks = len(week_labels)
    table_right = TBL_LABEL_X + TBL_LABEL_W + n_weeks * TBL_WEEK_W + TBL_TOTAL_W

    # ── Section header ─────────────────────────────────────────────────────
    # Forensic: centered at x≈396, RL y≈225 baseline, Helvetica-Bold 9.5pt
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 11.0)
    tbl_center = (TBL_LABEL_X + table_right) / 2
    c.drawCentredString(tbl_center, 272, f"Weekly Performance ({sites_str})")

    # ── Column header row ──────────────────────────────────────────────────
    # Forensic: cells from RL y=199 to y=215, each cell drawn as white-filled rect
    col_hdr_bot = TBL_COL_HDR_TOP - TBL_ROW_H   # = 199.0

    # Label cell
    c.setFillColor(C_WHITE)
    c.setStrokeColor(C_CELL_BDR)
    c.setLineWidth(0.5)
    c.rect(TBL_LABEL_X, col_hdr_bot, TBL_LABEL_W, TBL_ROW_H, fill=1, stroke=1)
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 9.0)
    c.drawCentredString(TBL_LABEL_X + TBL_LABEL_W / 2, col_hdr_bot + 5, "Week")

    # Week cells
    col_x = TBL_LABEL_X + TBL_LABEL_W
    for wk in week_labels:
        if wk.startswith("W"):
            wk_display = f"{year_str} W {wk[1:]}"   # "2026 W 24"
        else:
            wk_display = wk
        c.setFillColor(C_WHITE)
        c.setStrokeColor(C_CELL_BDR)
        c.setLineWidth(0.5)
        c.rect(col_x, col_hdr_bot, TBL_WEEK_W, TBL_ROW_H, fill=1, stroke=1)
        c.setFillColor(C_BLACK)
        c.setFont("Helvetica-Bold", 9.0)
        c.drawCentredString(col_x + TBL_WEEK_W / 2, col_hdr_bot + 5, wk_display)
        col_x += TBL_WEEK_W

    # Total cell (yellow bg)
    c.setFillColor(C_TOTAL_BG)
    c.setStrokeColor(C_CELL_BDR)
    c.setLineWidth(0.5)
    c.rect(col_x, col_hdr_bot, TBL_TOTAL_W, TBL_ROW_H, fill=1, stroke=1)
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 9.0)
    c.drawCentredString(col_x + TBL_TOTAL_W / 2, col_hdr_bot + 5, "Total")

    # ── Data rows ──────────────────────────────────────────────────────────
    last_wk_idx = n_weeks - 1
    for ri, (label, key) in enumerate(zip(ROW_LABELS, ROW_KEYS)):
        row_top = col_hdr_bot - ri * TBL_ROW_H          # top of this row
        row_bot = row_top - TBL_ROW_H                    # bottom of this row
        txt_y   = row_bot + 5                            # text baseline (RL)

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

            # Determine border for current-week SLA cells
            if ri in SLA_ROWS and is_last and v is not None:
                border_c = C_GREEN_SLA if float(v) >= SLA_TARGET else C_RED_SLA
                border_lw = 2.0
            else:
                border_c  = C_CELL_BDR
                border_lw = 0.5

            c.setFillColor(C_WHITE)
            c.setStrokeColor(border_c)
            c.setLineWidth(border_lw)
            c.rect(col_x, row_bot, TBL_WEEK_W, TBL_ROW_H, fill=1, stroke=1)

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


# ── Main generator ────────────────────────────────────────────────────────────
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
    year_2d  = year_str[2:]   # "26"

    # ── Title ─────────────────────────────────────────────────────────────
    # Forensic: Helvetica-Bold 20pt, black, RL baseline y=576, x=30
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(30, 580, "Robotics Destination Dray Metrics \u2014 NA")

    # ── Orange accent line ────────────────────────────────────────────────
    # Forensic: RL y=568, x=30-762, stroke #e8a838, lw=2.0
    c.setStrokeColor(C_ACCENT)
    c.setLineWidth(2.0)
    c.line(30, 568, 762, 568)

    # ── SLA Goals box ─────────────────────────────────────────────────────
    _draw_sla_goals(c)

    # ── Charts — Row 1 (AV→OA | OA→Del | E2E Transit) ────────────────────
    charts_r1 = [
        ("Leg: AV to OA (avg. days)",        "av_oa_avg",       False),
        ("Leg: OA to Delivery (avg. days)",  "oa_del_avg",      False),
        ("Leg: E2E Transit (avg. days)",      "e2e_avg",         False),
    ]
    for band_idx, (title, key, is_pct) in enumerate(charts_r1):
        _draw_chart(c, band_idx, 1, title,
                    _chart_vals(weeks_data, key), week_labels, year_2d, is_pct)

    # ── Charts — Row 2 (Empty→Term | OTP | Volume) ────────────────────────
    charts_r2 = [
        ("Leg: Empty to Termination (avg. days)", "empty_term_pct", False),
        ("Leg: On-Time to Promise %",              "otp_pct",        True),
        ("Volume (containers)",                    "containers",     False),
    ]
    for band_idx, (title, key, is_pct) in enumerate(charts_r2):
        _draw_chart(c, band_idx, 2, title,
                    _chart_vals(weeks_data, key), week_labels, year_2d, is_pct)

    # ── Performance table ─────────────────────────────────────────────────
    _draw_table(c, week_labels, weeks_data, totals, year_str, sites_str)

    # ── Footer ────────────────────────────────────────────────────────────
    # Forensic: Helvetica 7.5pt #666666, RL y=13, x=30
    # Format: "July 20, 2026 10:00am (CT) / 9:00am (MT) / 8:00am (PT)"
    try:
        date_str = report_date.strftime("%-B %-d, %Y")
    except ValueError:
        # Windows doesn't support %-d; strip leading zero manually
        date_str = report_date.strftime("%B %d, %Y").replace(" 0", " ")

    footer_left  = f"{date_str} {report_time}"

    c.setFillColor(C_GRAY)
    c.setFont("Helvetica", 7.5)
    c.drawString(30, 13, footer_left)

    c.save()
    return buf.getvalue()
