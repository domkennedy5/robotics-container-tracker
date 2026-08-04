"""
port_intel_scraper/handler.py
Lambda function — scrapes port operational intelligence and freight rate data.
Runs daily at 6 AM ET. Downloads tracker.db from S3, appends new rows, re-uploads.

Sources:
  - Google News RSS: USSAV, USORF, USLAX, USBOS port headlines + ocean freight news
  - Drewry WCI: weekly composite rate from meta description (static HTML)
  - NY Fed GSCPI: monthly supply chain pressure index (XLS download)
"""

import boto3
import io
import os
import re
import sqlite3
import ssl
import tempfile
import urllib.request
from datetime import datetime, timezone


# ── Config ────────────────────────────────────────────────────────────────────
S3_BUCKET = os.environ.get("S3_BUCKET", "robotics-container-tracker")
DB_KEY    = os.environ.get("DB_KEY", "tracker.db")

GOOGLE_NEWS_SOURCES = [
    {"port": "USSAV", "name": "Port of Savannah",      "query": "Savannah+port+operations+congestion+delay"},
    {"port": "USORF", "name": "Port of Virginia",       "query": "%22Port+of+Virginia%22+operations+cargo"},
    {"port": "USLAX", "name": "Port of Los Angeles",    "query": "%22Port+of+Los+Angeles%22+congestion+operations"},
    {"port": "USBOS", "name": "Port of Boston",         "query": "%22Conley+Terminal%22+OR+%22Boston+port%22+cargo+operations"},
    {"port": "ALL",   "name": "Ocean Freight / Rates",  "query": "transpacific+container+freight+rates+2026"},
    {"port": "ALL",   "name": "Supply Chain",           "query": "supply+chain+disruption+ports+2026"},
]

DREWRY_URL = ("https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise"
              "/world-container-index-assessed-by-drewry")
GSCPI_URL  = ("https://www.newyorkfed.org/medialibrary/research/interactives"
              "/gscpi/downloads/gscpi_data.xlsx")


# ── Helpers ───────────────────────────────────────────────────────────────────
def _ctx():
    c = ssl.create_default_context()
    c.check_hostname = False
    c.verify_mode = ssl.CERT_NONE
    return c

def _get(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout, context=_ctx()) as r:
        return r.read()

def _get_text(url, timeout=20):
    return _get(url, timeout).decode("utf-8", errors="ignore")


# ── DB setup ──────────────────────────────────────────────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS port_intel (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    port_code      TEXT,
    source         TEXT,
    headline       TEXT,
    summary        TEXT,
    url            TEXT,
    published_date TEXT,
    scraped_at     TEXT
);

CREATE TABLE IF NOT EXISTS freight_rates (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    source         TEXT,
    route          TEXT,
    rate_usd       REAL,
    wow_change_pct REAL,
    report_date    TEXT,
    scraped_at     TEXT,
    UNIQUE(source, route, report_date)
);

CREATE TABLE IF NOT EXISTS macro_signals (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    source      TEXT,
    metric      TEXT,
    value       REAL,
    period      TEXT,
    scraped_at  TEXT,
    UNIQUE(source, metric, period)
);
"""

def _ensure_schema(conn):
    for stmt in SCHEMA.strip().split(";"):
        s = stmt.strip()
        if s:
            conn.execute(s)
    conn.commit()


# ── Scrapers ──────────────────────────────────────────────────────────────────
def scrape_google_news(conn, now):
    """Pull top headlines from Google News RSS for each port + freight queries."""
    for src in GOOGLE_NEWS_SOURCES:
        try:
            url  = (f"https://news.google.com/rss/search?q={src['query']}"
                    f"&hl=en-US&gl=US&ceid=US:en")
            body = _get_text(url)
            items = re.findall(r"<item>(.*?)</item>", body, re.DOTALL)
            inserted = 0
            for item in items:
                if inserted >= 5:
                    break
                tm = re.search(r"<title>(.*?)</title>", item, re.DOTALL)
                lm = re.search(r"<link>(.*?)</link>",   item, re.DOTALL)
                dm = re.search(r"<pubDate>(.*?)</pubDate>", item, re.DOTALL)
                if not tm:
                    continue
                headline = re.sub(r"<[^>]+>", "", tm.group(1)).strip()
                link     = lm.group(1).strip() if lm else ""
                pub_date = dm.group(1).strip()[:25] if dm else now[:10]
                # Skip duplicates by headline+port
                exists = conn.execute(
                    "SELECT 1 FROM port_intel WHERE port_code=? AND headline=?",
                    (src["port"], headline)
                ).fetchone()
                if not exists:
                    conn.execute(
                        "INSERT INTO port_intel "
                        "(port_code, source, headline, url, published_date, scraped_at) "
                        "VALUES (?,?,?,?,?,?)",
                        (src["port"], src["name"], headline, link, pub_date, now)
                    )
                    inserted += 1
            print(f"  Google News [{src['port']} / {src['name']}]: {inserted} new headlines")
        except Exception as e:
            print(f"  Google News [{src['port']}] error: {e}")


def scrape_drewry(conn, now):
    """
    Drewry WCI — rate + WoW change live in <meta name="description">.
    Example: '30 Jul 2026: Drewry's WCI decreased 3% to $4,255 per 40ft container...'
    """
    try:
        body  = _get_text(DREWRY_URL)
        m_desc = re.search(r'<meta name="description" content="([^"]+)"', body)
        if not m_desc:
            print("  Drewry: meta description not found")
            return

        raw = (m_desc.group(1)
               .replace("&#x24;", "$").replace("&#x25;", "%")
               .replace("&#x3a;", ":").replace("&rsquo;", "'")
               .replace("&#x28;", "(").replace("&#x29;", ")")
               .replace("&ndash;", "–").replace("&mdash;", "—")
               .replace("&amp;", "&"))

        rate_m = re.search(r"\$([\d,]+)", raw)
        chg_m  = re.search(r"(increased|decreased)\s+(\d+(?:\.\d+)?)\s*%", raw, re.IGNORECASE)
        date_m = re.search(r"^(\d{1,2}\s+\w+\s+\d{4})", raw.strip())

        rate      = float(rate_m.group(1).replace(",", "")) if rate_m else None
        direction = chg_m.group(1).lower() if chg_m else None
        pct_val   = float(chg_m.group(2)) if chg_m else None
        wow       = (-pct_val if direction == "decreased" else pct_val) if pct_val else None
        rpt_date  = date_m.group(1) if date_m else now[:10]

        rate_str  = f"${rate_m.group(1)}" if rate_m else "–"
        arrow     = "↓" if direction == "decreased" else ("↑" if direction == "increased" else "")
        pct_str   = f"{arrow}{pct_val:.0f}%" if pct_val else ""
        headline  = f"Drewry WCI {pct_str} to {rate_str}/FEU ({rpt_date})"

        conn.execute(
            "INSERT OR IGNORE INTO freight_rates "
            "(source, route, rate_usd, wow_change_pct, report_date, scraped_at) "
            "VALUES (?,?,?,?,?,?)",
            ("Drewry WCI", "Global Composite", rate, wow, rpt_date, now)
        )
        conn.execute(
            "INSERT INTO port_intel "
            "(port_code, source, headline, summary, url, published_date, scraped_at) "
            "VALUES (?,?,?,?,?,?,?)",
            ("ALL", "Drewry WCI", headline, raw[:400], DREWRY_URL, rpt_date, now)
        )
        print(f"  Drewry WCI: {headline}")
    except Exception as e:
        print(f"  Drewry error: {e}")


def scrape_gscpi(conn, now):
    """
    NY Fed GSCPI — download Excel file, extract latest period + value.
    File uses old .xls (OLE2) format — requires xlrd.
    """
    try:
        data = _get(GSCPI_URL, timeout=25)
        try:
            import xlrd
            wb  = xlrd.open_workbook(file_contents=data)
            ws  = wb.sheet_by_index(0)
            last_period = last_value = None
            for rx in range(ws.nrows):
                row = ws.row_values(rx)
                if row[0] and row[1] is not None:
                    try:
                        v = float(row[1])
                        last_period = str(row[0])
                        last_value  = v
                    except (ValueError, TypeError):
                        pass
        except ImportError:
            # xlrd not available — try openpyxl (works if file is actually xlsx)
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(data))
            ws = wb.active
            last_period = last_value = None
            for row in ws.iter_rows(values_only=True):
                if row[0] is not None and row[1] is not None:
                    try:
                        last_period = str(row[0])
                        last_value  = float(row[1])
                    except (ValueError, TypeError):
                        pass

        if last_value is not None:
            conn.execute(
                "INSERT OR IGNORE INTO macro_signals "
                "(source, metric, value, period, scraped_at) VALUES (?,?,?,?,?)",
                ("NY Fed", "GSCPI", last_value, last_period, now)
            )
            direction = "elevated" if last_value > 0.5 else ("low" if last_value < -0.5 else "neutral")
            print(f"  GSCPI: {last_period} = {last_value:.2f} ({direction})")
        else:
            print("  GSCPI: no data rows found")
    except Exception as e:
        print(f"  GSCPI error: {e}")


# ── S3 helpers ────────────────────────────────────────────────────────────────
def _download_db(s3, bucket, key):
    obj = s3.get_object(Bucket=bucket, Key=key)
    return obj["Body"].read()

def _upload_db(s3, bucket, key, path):
    with open(path, "rb") as f:
        s3.put_object(Bucket=bucket, Key=key, Body=f.read())


# ── Lambda handler ────────────────────────────────────────────────────────────
def handler(event, context):
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    print(f"Port Intel Scraper — {now}")

    s3 = boto3.client("s3")

    # Download DB from S3 to temp file
    db_bytes = _download_db(s3, S3_BUCKET, DB_KEY)
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.write(db_bytes)
    tmp.close()
    tmp_path = tmp.name

    conn = sqlite3.connect(tmp_path)
    _ensure_schema(conn)

    print("Scraping sources...")
    scrape_google_news(conn, now)
    scrape_drewry(conn, now)
    scrape_gscpi(conn, now)

    conn.commit()
    conn.close()

    # Upload updated DB back to S3
    _upload_db(s3, S3_BUCKET, DB_KEY, tmp_path)
    os.unlink(tmp_path)

    print("Done — DB updated and uploaded.")
    return {"statusCode": 200, "body": "ok"}
