"""
wbr_pdf.py — Generate the standard WBR PDF slide matching GLS_Robotics layout.

Page: 792×612 landscape letter.
Colors sourced directly from GLS_Robotics_2026-7-20.pdf via PyMuPDF inspection.
"""
from __future__ import annotations
import io
from datetime import date, datetime
from typing import Optional

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas as rl_canvas

# ── Brand colors (from PDF inspection) ────────────────────────────────────────
C_HEADER_BG  = colors.Color(0.200, 0.200, 0.200)   # dark gray header
C_BAR        = colors.Color(0.294, 0.675, 0.776)   # teal bar #4BACC6
C_BAR_CURR   = colors.Color(0.173, 0.518, 0.616)   # slightly darker for current week
C_SLA_LINE   = colors.Color(0.8, 0.2, 0.2)         # red SLA threshold line
C_GREEN      = colors.Color(0.133, 0.545, 0.133)
C_RED        = colors.Color(0.800, 0.133, 0.133)
C_TOTAL_BG   = colors.Color(1.000, 0.980, 0.804)   # light yellow #FFFACD
C_TABLE_HDR  = colors.Color(0.200, 0.200, 0.200)
C_WHITE      = colors.white
C_BLACK      = colors.black
C_LTGRAY     = colors.Color(0.85, 0.85, 0.85)
C_GRIDLINE   = colors.Color(0.60, 0.60, 0.60)

PAGE_W, PAGE_H = 792.0, 612.0

# ── Row labels and display helpers ────────────────────────────────────────────
ROW_LABELS = [
    "Containers",
    "AV\u2192OA Avg",
    "AV\u2192OA SLA%",
    "AV\u2192OA P90",
    "OA\u2192Del Avg",
    "OA\u2192Del SLA%",
    "OA\u2192Del P90",
    "Empty\u2192Term SLA%",
    "E2E Transit Avg",
    "On-Time to Promise%",
]

# SLA threshold for coloring (None = no border)
ROW_SLA = [None, None, 95, None, None, 95, None, None, None, 95]


def _fmt(v, row_idx: int) -> str:
    if v is None:
        return "–"
    if row_idx in (2, 5, 7, 9):   # SLA%
        return f"{int(v)}%"
    if row_idx in (1, 4):          # avg (decimal)
        return f"{float(v):.1f}"
    return str(int(v))


def _row_val(m: Optional[dict], row_idx: int):
    if not m:
        return None
    keys = [
        "containers", "av_oa_avg", "av_oa_sla_pct", "av_oa_p90",
        "oa_del_avg", "oa_del_sla_pct", "oa_del_p90",
        "empty_term_pct", "e2e_avg", "otp_pct",
    ]
    return m.get(keys[row_idx])


def _chart_vals(weeks_data: list[Optional[dict]], key: str) -> list:
    return [w.get(key) if w else None for w in weeks_data]


# ── Bar chart drawing ──────────────────────────────────────────────────────────
def _draw_bar_chart(
    c: rl_canvas.Canvas,
    x: float, y: float, w: float, h: float,
    title: str,
    values: list,          # list of numbers or None; last = current week
    week_labels: list[str],
    sla_val: Optional[float] = None,
    is_pct: bool = False,
    current_idx: int = -1,
):
    n = len(values)
    clean = [v for v in values if v is not None]
    max_v = max(clean) if clean else 1
    if sla_val and sla_val > max_v:
        max_v = sla_val * 1.05
    max_v = max_v * 1.15 or 1

    bar_area_x = x + 2
    bar_area_y = y + 18        # leave room for x-axis labels
    bar_area_w = w - 4
    bar_area_h = h - 32        # leave room for title

    # Title
    c.setFont("Helvetica-Bold", 7)
    c.setFillColor(C_BLACK)
    c.drawCentredString(x + w / 2, y + h - 1, title)

    # SLA reference line
    if sla_val is not None and max_v > 0:
        sla_y = bar_area_y + (sla_val / max_v) * bar_area_h
        c.setStrokeColor(C_SLA_LINE)
        c.setLineWidth(0.8)
        c.setDash(3, 2)
        c.line(bar_area_x, sla_y, bar_area_x + bar_area_w, sla_y)
        c.setDash()

    # Bars
    bar_w = bar_area_w / n * 0.6
    gap   = bar_area_w / n

    for i, v in enumerate(values):
        bx = bar_area_x + i * gap + (gap - bar_w) / 2
        if v is None:
            continue
        bh = (v / max_v) * bar_area_h if max_v else 0
        by = bar_area_y

        fill = C_BAR_CURR if i == len(values) - 1 else C_BAR
        c.setFillColor(fill)
        c.setStrokeColor(fill)
        c.rect(bx, by, bar_w, bh, fill=1, stroke=0)

        # Value label on top
        c.setFillColor(C_BLACK)
        c.setFont("Helvetica-Bold", 6)
        label = f"{round(v)}%" if is_pct else str(round(v))
        c.drawCentredString(bx + bar_w / 2, by + bh + 1, label)

        # Week label below
        c.setFont("Helvetica", 5.5)
        c.drawCentredString(bx + bar_w / 2, bar_area_y - 9, week_labels[i] if i < len(week_labels) else "")


# ── SLA goals box ──────────────────────────────────────────────────────────────
def _draw_sla_goals(c: rl_canvas.Canvas, x: float, y: float, w: float, h: float):
    c.setFillColor(C_WHITE)
    c.setStrokeColor(C_GRIDLINE)
    c.setLineWidth(0.5)
    c.rect(x, y, w, h, fill=1, stroke=1)

    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(x + 3, y + h - 9, "SLA Goals")

    goals = [
        "AV to OA \u2264 3 Days",
        "OA to Delivery \u2264 3 Days",
        "Empty to Term \u2264 3 Days",
        "E2E Transit",
        "On-Time to Promise \u2265 95%",
    ]
    c.setFont("Helvetica", 6.5)
    for i, g in enumerate(goals):
        c.drawString(x + 4, y + h - 20 - i * 11, f"\u25aa {g}")


# ── Performance table ──────────────────────────────────────────────────────────
def _draw_performance_table(
    c: rl_canvas.Canvas,
    x: float, y: float, w: float, h: float,
    week_labels: list[str],
    weeks_data: list[Optional[dict]],
    totals: Optional[dict],
    sites_str: str = "USORF, USBOS, USSAV, USLAX",
):
    n_weeks = len(week_labels)
    col_count = 1 + n_weeks + 1   # row label + weeks + total

    # Dynamic column widths: label=105, week=73, total=80
    label_w = 105
    total_w = 80
    week_w  = (w - label_w - total_w) / n_weeks

    row_count  = len(ROW_LABELS)
    header_h   = 14
    subhdr_h   = 12
    row_h      = (h - header_h - subhdr_h) / row_count

    cur_y = y + h

    # Section header
    cur_y -= header_h
    c.setFillColor(C_TABLE_HDR)
    c.rect(x, cur_y, w, header_h, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(x + 4, cur_y + 3, f"Weekly Performance ({sites_str})")

    # Column headers
    cur_y -= subhdr_h
    c.setFillColor(C_LTGRAY)
    c.rect(x, cur_y, w, subhdr_h, fill=1, stroke=0)

    # Draw col headers
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(x + 3, cur_y + 3, "Week")

    col_x = x + label_w
    for wk in week_labels:
        c.setFont("Helvetica-Bold", 7)
        cx = col_x + week_w / 2
        year_str = "2026 " if wk == week_labels[0] else ""
        c.drawCentredString(cx, cur_y + 3, year_str + wk)
        col_x += week_w

    # Total header (yellow bg)
    c.setFillColor(C_TOTAL_BG)
    c.rect(col_x, cur_y, total_w, subhdr_h, fill=1, stroke=0)
    c.setFillColor(C_BLACK)
    c.setFont("Helvetica-Bold", 7)
    c.drawCentredString(col_x + total_w / 2, cur_y + 3, "Total")

    # Data rows
    for ri, label in enumerate(ROW_LABELS):
        cur_y -= row_h
        # Alternating row shading
        bg = colors.Color(0.96, 0.96, 0.96) if ri % 2 == 1 else C_WHITE
        c.setFillColor(bg)
        c.rect(x, cur_y, w - total_w, row_h, fill=1, stroke=0)

        # Total cell background
        c.setFillColor(C_TOTAL_BG)
        c.rect(x + w - total_w, cur_y, total_w, row_h, fill=1, stroke=0)

        # Row label
        c.setFillColor(C_BLACK)
        c.setFont("Helvetica-Bold", 6.5)
        c.drawString(x + 3, cur_y + row_h / 2 - 2.5, label)

        # Week values
        col_x = x + label_w
        sla_threshold = ROW_SLA[ri]
        last_wk_idx   = len(week_labels) - 1

        for wi, wk_data in enumerate(weeks_data):
            v   = _row_val(wk_data, ri)
            txt = _fmt(v, ri)
            cx  = col_x + week_w / 2
            cy  = cur_y + row_h / 2 - 2.5

            # SLA border on current week SLA% cells
            if sla_threshold is not None and wi == last_wk_idx and v is not None:
                border_color = C_GREEN if v >= sla_threshold else C_RED
                c.setStrokeColor(border_color)
                c.setLineWidth(1.5)
                c.rect(col_x + 1, cur_y + 1, week_w - 2, row_h - 2, fill=0, stroke=1)

            c.setFillColor(C_BLACK)
            c.setFont("Helvetica-Bold" if wi == last_wk_idx else "Helvetica", 7)
            c.drawCentredString(cx, cy, txt)
            col_x += week_w

        # Total cell
        tv  = _row_val(totals, ri)
        ttx = _fmt(tv, ri)
        c.setFillColor(C_BLACK)
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(x + w - total_w + total_w / 2, cur_y + row_h / 2 - 2.5, ttx)

    # Outer border
    c.setStrokeColor(C_GRIDLINE)
    c.setLineWidth(0.5)
    c.rect(x, cur_y, w, y + h - header_h - subhdr_h - cur_y, fill=0, stroke=1)

    # Vertical grid lines
    col_x = x + label_w
    table_top = y + h - header_h - subhdr_h
    for _ in week_labels:
        c.line(col_x, cur_y, col_x, table_top)
        col_x += week_w
    c.line(col_x, cur_y, col_x, table_top)  # before total


# ── Main PDF generator ────────────────────────────────────────────────────────
def generate_standard_wbr(
    week_labels: list[str],
    weeks_data: list[Optional[dict]],
    totals: Optional[dict],
    report_date: date,
    current_week_label: str = "",
    report_time: str = "10:00am (CST) | 9:00am (MT) | 8:00am (PT)",
    sites_str: str = "USORF, USBOS, USSAV, USLAX",
) -> bytes:
    """
    Generate the standard 1-page WBR PDF slide.
    week_labels: e.g. ["W24","W25","W26","W27","W28","W29"]
    weeks_data:  list of metric dicts (same length); None = carry-forward missing
    totals:      pre-computed volume-weighted total dict
    report_date: date of WBR (Monday submission date)
    """
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    # ── Header bar ────────────────────────────────────────────────────────
    c.setFillColor(C_HEADER_BG)
    c.rect(0, PAGE_H - 45, PAGE_W, 45, fill=1, stroke=0)

    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, PAGE_H - 28, "Robotics Destination Dray Metrics \u2014 NA")

    date_str = f"{report_date.month}/{report_date.day}/{report_date.year}"
    c.setFont("Helvetica", 8)
    footer_txt = f"{date_str} {report_time}"
    c.drawRightString(PAGE_W - 10, PAGE_H - 38, footer_txt)

    # ── Layout constants ──────────────────────────────────────────────────
    # Two rows of 3 charts each
    # Row 1 (top): y range 452–562 (reportlab coords)
    # Row 2 (mid): y range 318–428
    # Table: y range 30–300
    # SLA Goals box: left of Row 1 charts

    CHART_TOP1 = PAGE_H - 45 - 120   # top of row-1 charts = ~447
    CHART_H    = 110
    CHART_TOP2 = CHART_TOP1 - 130
    TABLE_TOP  = CHART_TOP2 - 12
    TABLE_H    = TABLE_TOP - 20

    # SLA Goals box (left panel, row 1 area)
    SLA_X = 5
    SLA_W = 255
    _draw_sla_goals(c, SLA_X, CHART_TOP1 - CHART_H, SLA_W, CHART_H)

    # Row 1: AV->OA | OA->Del | E2E Transit
    chart_row1_x = SLA_X + SLA_W + 4
    chart_w1 = (PAGE_W - chart_row1_x - 6) / 3

    n_wk = len(week_labels)
    short_labels = week_labels   # ["W24","W25",...]

    _draw_bar_chart(c,
        chart_row1_x, CHART_TOP1 - CHART_H, chart_w1, CHART_H,
        "Leg: AV to OA (avg. days)",
        _chart_vals(weeks_data, "av_oa_avg"), short_labels, sla_val=3.0,
    )
    _draw_bar_chart(c,
        chart_row1_x + chart_w1 + 3, CHART_TOP1 - CHART_H, chart_w1, CHART_H,
        "Leg: OA to Delivery (avg. days)",
        _chart_vals(weeks_data, "oa_del_avg"), short_labels, sla_val=3.0,
    )
    _draw_bar_chart(c,
        chart_row1_x + 2*(chart_w1 + 3), CHART_TOP1 - CHART_H, chart_w1, CHART_H,
        "Leg: E2E Transit (avg. days)",
        _chart_vals(weeks_data, "e2e_avg"), short_labels,
    )

    # Row 2: Empty->Term | OTP | Volume
    row2_y = CHART_TOP2 - CHART_H
    chart_w2 = (PAGE_W - 14) / 3

    _draw_bar_chart(c,
        8, row2_y, chart_w2, CHART_H,
        "Leg: Empty to Termination (avg. days)",
        _chart_vals(weeks_data, "empty_term_pct"), short_labels,
        is_pct=True,
    )
    _draw_bar_chart(c,
        8 + chart_w2 + 4, row2_y, chart_w2, CHART_H,
        "Leg: On-Time to Promise %",
        _chart_vals(weeks_data, "otp_pct"), short_labels,
        sla_val=95.0, is_pct=True,
    )
    _draw_bar_chart(c,
        8 + 2*(chart_w2 + 4), row2_y, chart_w2, CHART_H,
        "Volume (containers)",
        _chart_vals(weeks_data, "containers"), short_labels,
    )

    # ── Performance table ──────────────────────────────────────────────────
    _draw_performance_table(
        c,
        x=5, y=20, w=PAGE_W - 10, h=TABLE_H,
        week_labels=week_labels,
        weeks_data=weeks_data,
        totals=totals,
        sites_str=sites_str,
    )

    # ── Footer ─────────────────────────────────────────────────────────────
    c.setFillColor(colors.Color(0.5, 0.5, 0.5))
    c.setFont("Helvetica", 6)
    c.drawString(5, 8, f"Amazon Global Logistics | AGL Robotics Dray Program | Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    c.save()
    return buf.getvalue()
