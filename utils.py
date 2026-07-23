"""
utils.py — shared logic used by both app.py and Lambda functions.
Keeps container parsing, matching, and file parsing in one place.
"""
import re
import io
import openpyxl
import pandas as pd


# ── container ID helpers ───────────────────────────────────────────────────────

def normalize_container(cid: str) -> str:
    """Strip dashes and spaces so 'TCNU389902-4' and 'TCNU3899024' are the same."""
    return re.sub(r"[-\s]", "", str(cid).strip().upper())


def containers_match(query_norm: str, dbr_norm: str) -> bool:
    """Prefix-tolerant match — handles IDs submitted without the trailing check digit.
    'GCXU542076' matches 'GCXU5420760' because lengths differ by exactly 1.
    """
    if query_norm == dbr_norm:
        return True
    if abs(len(query_norm) - len(dbr_norm)) == 1:
        short, long_ = (
            (query_norm, dbr_norm) if len(query_norm) < len(dbr_norm)
            else (dbr_norm, query_norm)
        )
        return long_.startswith(short)
    return False


def parse_container_list(text: str) -> list:
    """Accept newline, comma, semicolon, or space-separated container IDs."""
    ids = re.split(r"[\n,;\s]+", text.strip())
    return [i.strip() for i in ids if i.strip()]


# ── carrier file parser ────────────────────────────────────────────────────────

def parse_carrier_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Parse any carrier submission Excel or CSV.

    Handles multiple formats:
    - AGL standard template (header on row 2, row 1 is a note)
    - Legacy ARVY / Carrier DBR formats (header on row 1 or 2)
    - CSV files

    Searches every sheet for ANY row containing 'Container' to find the header,
    then extracts container IDs and associated fields.
    Returns DataFrame with columns: container_id, sheet_source, terminal, status, notes.
    """
    rows_all = []

    if filename.lower().endswith(".csv"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), dtype=str)
            cont_col = next((c for c in df.columns if "container" in c.lower()), None)
            if cont_col:
                for _, row in df.iterrows():
                    if pd.notna(row[cont_col]) and str(row[cont_col]).strip():
                        rows_all.append(_extract_row_fields(row.to_dict(), "Sheet1", cont_col))
        except Exception:
            pass
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            # Find header row — scan first 8 rows for one containing 'Container'
            header_idx = next(
                (i for i, row in enumerate(rows[:8])
                 if any(c and "container" in str(c).lower() for c in row)),
                None
            )
            if header_idx is None:
                continue

            headers = [
                str(h).strip() if h and str(h).strip() else f"_col_{i}"
                for i, h in enumerate(rows[header_idx])
            ]
            cont_col = next((h for h in headers if "container" in h.lower()), None)
            if not cont_col:
                continue

            for row in rows[header_idx + 1:]:
                if not any(c for c in row if c is not None):
                    continue
                rec = dict(zip(headers, row))
                val = rec.get(cont_col)
                if not val or not str(val).strip() or str(val).strip() in ("None", "nan"):
                    continue
                rows_all.append(_extract_row_fields(rec, sname, cont_col))
        wb.close()

    if not rows_all:
        return pd.DataFrame()
    return pd.DataFrame(rows_all)


def _extract_row_fields(rec: dict, sheet: str, cont_col: str) -> dict:
    """Pull standardized fields from a raw carrier row dict."""
    def _find(keys):
        for k in rec:
            if any(kw in str(k).lower() for kw in keys):
                v = rec[k]
                if v and str(v).strip() not in ("None", "nan", ""):
                    return str(v).strip()
        return None

    return {
        "container_id":  normalize_container(str(rec[cont_col])),
        "sheet_source":  sheet,
        "terminal":      _find(["terminal", "port"]),
        "status":        _find(["status", "state"]),
        "notes":         _find(["note", "comment", "reason", "bridge"]),
    }
