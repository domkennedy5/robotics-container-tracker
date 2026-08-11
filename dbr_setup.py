"""
dbr_setup.py — One-time setup: creates inbound_containers + carrier_contacts tables
and seeds them from Robotics DBR Tracker.xlsx.

Run once: python dbr_setup.py
"""
import sqlite3
import openpyxl
from datetime import datetime
import os
import sys

DB_PATH    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker.db")
EXCEL_PATH = r"C:\Users\kennewdo\Downloads\Robotics DBR Tracker.xlsx"

if not os.path.exists(EXCEL_PATH):
    print(f"Excel not found at {EXCEL_PATH} — update EXCEL_PATH and re-run.")
    sys.exit(1)

conn = sqlite3.connect(DB_PATH)

# ── Create inbound_containers ─────────────────────────────────────────────────
conn.execute("""
    CREATE TABLE IF NOT EXISTS inbound_containers (
        id               INTEGER PRIMARY KEY AUTOINCREMENT,
        container_id     TEXT    NOT NULL UNIQUE,
        carrier          TEXT,
        port_region      TEXT,
        yard_dest        TEXT,
        yard_sched       TEXT,
        yard_actual      TEXT,
        fc_dest          TEXT,
        fc_sched         TEXT,
        fc_actual        TEXT,
        empty_returned   TEXT,
        status           TEXT    DEFAULT 'pending',
        notes            TEXT,
        source_email_date TEXT,
        live_drop        TEXT,
        date_received    TEXT,
        last_updated     TEXT,
        source_file      TEXT
    )
""")

# ── Create carrier_contacts ───────────────────────────────────────────────────
conn.execute("""
    CREATE TABLE IF NOT EXISTS carrier_contacts (
        id               INTEGER PRIMARY KEY AUTOINCREMENT,
        scac             TEXT    NOT NULL UNIQUE,
        display_name     TEXT,
        contact_name     TEXT,
        email            TEXT,
        reminder_enabled INTEGER DEFAULT 0,
        created_at       TEXT    DEFAULT (datetime('now')),
        updated_at       TEXT    DEFAULT (datetime('now'))
    )
""")

# ── Seed carrier_contacts (fill email later via app) ─────────────────────────
_CARRIERS = [
    ("ATMI", "ATMI (Cargomatic)",  "Tyler Domingues",                None, 0),
    ("ARVY", "ARVY (Arrive)",      "Tyler Spangler",                 None, 0),
    ("HDDR", "HUDD (Maersk)",      "Sandji Ruffin / Jerry Nesbit",   None, 0),
    ("RKNE", "RKNE (RoadOne)",     "Mark Brennan",                   None, 0),
    ("TGHE", "TGHE",               None,                             None, 0),
]
for scac, dname, contact, email, enabled in _CARRIERS:
    conn.execute("""
        INSERT OR IGNORE INTO carrier_contacts
            (scac, display_name, contact_name, email, reminder_enabled)
        VALUES (?,?,?,?,?)
    """, (scac, dname, contact, email, enabled))

# ── Import from Excel Delivery Plan sheet ────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

def _scac(carrier_str):
    if not carrier_str:
        return None
    code = str(carrier_str).strip().split()[0].upper()
    return "HDDR" if code == "HUDD" else code

def _isodate(val):
    if val is None:
        return None
    if hasattr(val, "date"):
        return val.date().isoformat()
    try:
        return datetime.strptime(str(val).strip(), "%Y-%m-%d").date().isoformat()
    except Exception:
        return None

STATUS_MAP = {
    "Delivered": "delivered",
    "At Yard":   "at_yard",
    "Pending":   "pending",
}

ws = wb["Delivery Plan"]
now = datetime.now().isoformat()
inserted = updated = skipped = 0

for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i <= 1:  # row 0 = group header, row 1 = column headers
        continue
    container = row[3]
    if not container:
        continue
    container = str(container).strip()

    carrier       = _scac(row[1])
    port_region   = str(row[2]).strip() if row[2] else None
    yard_dest     = str(row[4]).strip() if row[4] else None
    yard_sched    = _isodate(row[5])
    yard_actual   = _isodate(row[6])
    fc_dest       = str(row[7]).strip() if row[7] else None
    fc_sched      = _isodate(row[8])
    fc_actual     = _isodate(row[9])
    empty_ret     = _isodate(row[10])
    status_raw    = str(row[11]).strip() if row[11] else "Pending"
    notes         = str(row[12]).strip() if row[12] else None
    src_email_dt  = str(row[13]).strip() if row[13] else None
    live_drop     = str(row[14]).strip() if row[14] else None
    date_received = _isodate(row[0])

    # Derive status from actual data (more reliable than the cell text)
    if empty_ret:
        status = "empty_returned"
    elif fc_actual:
        status = "delivered"
    elif yard_actual:
        status = "at_yard"
    else:
        status = STATUS_MAP.get(status_raw, "pending")

    cur = conn.execute(
        "SELECT id FROM inbound_containers WHERE container_id = ?", (container,)
    )
    exists = cur.fetchone()

    if exists:
        # Update only if any delivery dates are newly available
        conn.execute("""
            UPDATE inbound_containers SET
                carrier           = COALESCE(carrier, ?),
                port_region       = COALESCE(port_region, ?),
                yard_dest         = COALESCE(yard_dest, ?),
                yard_sched        = COALESCE(yard_sched, ?),
                yard_actual       = COALESCE(yard_actual, ?),
                fc_dest           = COALESCE(fc_dest, ?),
                fc_sched          = COALESCE(fc_sched, ?),
                fc_actual         = COALESCE(fc_actual, ?),
                empty_returned    = COALESCE(empty_returned, ?),
                status            = ?,
                notes             = COALESCE(notes, ?),
                source_email_date = COALESCE(source_email_date, ?),
                live_drop         = COALESCE(live_drop, ?),
                last_updated      = ?
            WHERE container_id = ?
        """, (carrier, port_region, yard_dest, yard_sched, yard_actual,
              fc_dest, fc_sched, fc_actual, empty_ret, status,
              notes, src_email_dt, live_drop, now, container))
        updated += 1
    else:
        conn.execute("""
            INSERT INTO inbound_containers
                (container_id, carrier, port_region, yard_dest, yard_sched, yard_actual,
                 fc_dest, fc_sched, fc_actual, empty_returned, status, notes,
                 source_email_date, live_drop, date_received, last_updated, source_file)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (container, carrier, port_region, yard_dest, yard_sched, yard_actual,
              fc_dest, fc_sched, fc_actual, empty_ret, status, notes,
              src_email_dt, live_drop, date_received, now, "Robotics DBR Tracker.xlsx"))
        inserted += 1

conn.commit()
conn.close()
print(f"Done.  Inserted: {inserted}  Updated: {updated}  Skipped: {skipped}")
print("Tables created: inbound_containers, carrier_contacts")
