"""
data_sync.py — S3 sync for DBR file and SQLite DB.
Mirrors the Signal Tower pattern: pull on startup, push after writes.
Falls back to local files gracefully if S3 is unreachable.
"""
import os
import io
import json
import logging
from datetime import datetime, timezone

import boto3
from botocore.exceptions import BotoCoreError, ClientError

logger = logging.getLogger(__name__)

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
LOCAL_DBR   = os.path.join(BASE_DIR, "dbr", "latest.xlsx")
LOCAL_DB    = os.path.join(BASE_DIR, "tracker.db")
META_KEY    = "last_updated.json"
DBR_KEY     = "dbr/latest.xlsx"
DB_KEY      = "db/tracker.db"
CACHE_TTL   = 300  # seconds — don't re-pull from S3 more than once per 5 min


def _client(aws_key: str, aws_secret: str, region: str):
    return boto3.client(
        "s3",
        aws_access_key_id=aws_key,
        aws_secret_access_key=aws_secret,
        region_name=region,
    )


def _meta(s3, bucket: str) -> dict:
    try:
        obj = s3.get_object(Bucket=bucket, Key=META_KEY)
        return json.loads(obj["Body"].read())
    except Exception:
        return {}


def _set_meta(s3, bucket: str, updates: dict):
    try:
        current = _meta(s3, bucket)
        current.update(updates)
        s3.put_object(Bucket=bucket, Key=META_KEY,
                      Body=json.dumps(current).encode(),
                      ContentType="application/json")
    except Exception as e:
        logger.warning(f"Could not update S3 metadata: {e}")


# ── DBR ───────────────────────────────────────────────────────────────────────

def pull_dbr_from_s3(aws_key: str, aws_secret: str, region: str, bucket: str) -> bytes | None:
    """Download the latest DBR from S3. Returns file bytes or None on failure."""
    try:
        s3 = _client(aws_key, aws_secret, region)
        obj = s3.get_object(Bucket=bucket, Key=DBR_KEY)
        data = obj["Body"].read()
        os.makedirs(os.path.dirname(LOCAL_DBR), exist_ok=True)
        with open(LOCAL_DBR, "wb") as f:
            f.write(data)
        logger.info("DBR pulled from S3")
        return data
    except ClientError as e:
        if e.response["Error"]["Code"] == "NoSuchKey":
            logger.info("No DBR in S3 yet")
        else:
            logger.warning(f"S3 pull failed: {e}")
    except Exception as e:
        logger.warning(f"S3 pull failed: {e}")
    # fallback to local
    if os.path.exists(LOCAL_DBR):
        with open(LOCAL_DBR, "rb") as f:
            return f.read()
    return None


def push_dbr_to_s3(file_bytes: bytes, aws_key: str, aws_secret: str,
                   region: str, bucket: str, filename: str = ""):
    """Upload a new DBR to S3 and archive the previous one."""
    try:
        s3 = _client(aws_key, aws_secret, region)
        stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        # archive current before overwriting
        try:
            s3.copy_object(
                Bucket=bucket,
                CopySource={"Bucket": bucket, "Key": DBR_KEY},
                Key=f"dbr/archive/{stamp}.xlsx",
            )
        except ClientError:
            pass  # nothing to archive yet
        s3.put_object(Bucket=bucket, Key=DBR_KEY, Body=file_bytes,
                      ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        _set_meta(s3, bucket, {"dbr": datetime.now(timezone.utc).isoformat(),
                                "dbr_filename": filename})
        # cache locally
        os.makedirs(os.path.dirname(LOCAL_DBR), exist_ok=True)
        with open(LOCAL_DBR, "wb") as f:
            f.write(file_bytes)
        logger.info("DBR pushed to S3")
        return True
    except Exception as e:
        logger.warning(f"S3 push failed: {e}")
        return False


def get_dbr_last_updated(aws_key: str, aws_secret: str,
                          region: str, bucket: str) -> str | None:
    """Return ISO timestamp string of when the DBR was last uploaded, or None."""
    try:
        s3 = _client(aws_key, aws_secret, region)
        meta = _meta(s3, bucket)
        return meta.get("dbr")
    except Exception:
        return None


# ── SQLite DB ─────────────────────────────────────────────────────────────────

def pull_db_from_s3(aws_key: str, aws_secret: str, region: str, bucket: str) -> bool:
    """Download tracker.db from S3. Returns True on success."""
    try:
        s3 = _client(aws_key, aws_secret, region)
        obj = s3.get_object(Bucket=bucket, Key=DB_KEY)
        with open(LOCAL_DB, "wb") as f:
            f.write(obj["Body"].read())
        logger.info("DB pulled from S3")
        return True
    except ClientError as e:
        if e.response["Error"]["Code"] == "NoSuchKey":
            logger.info("No DB in S3 yet — starting fresh")
        else:
            logger.warning(f"DB pull failed: {e}")
    except Exception as e:
        logger.warning(f"DB pull failed: {e}")
    return False


def push_db_to_s3(aws_key: str, aws_secret: str, region: str, bucket: str) -> bool:
    """Upload local tracker.db to S3 after writes."""
    try:
        if not os.path.exists(LOCAL_DB):
            return False
        s3 = _client(aws_key, aws_secret, region)
        with open(LOCAL_DB, "rb") as f:
            s3.put_object(Bucket=bucket, Key=DB_KEY, Body=f.read())
        _set_meta(s3, bucket, {"db": datetime.now(timezone.utc).isoformat()})
        logger.info("DB pushed to S3")
        return True
    except Exception as e:
        logger.warning(f"DB push failed: {e}")
        return False


# ── WBR PDF ───────────────────────────────────────────────────────────────────

WBR_PDF_KEY  = "wbr/latest.pdf"
WBR_META_KEY = "wbr/latest_meta.json"


def push_wbr_pdf_to_s3(pdf_bytes: bytes, week_num: int, report_date_iso: str,
                        aws_key: str, aws_secret: str, region: str, bucket: str) -> bool:
    """Upload the generated WBR PDF + metadata to S3 so dashboards persist across sessions."""
    try:
        s3 = _client(aws_key, aws_secret, region)
        s3.put_object(Bucket=bucket, Key=WBR_PDF_KEY, Body=pdf_bytes,
                      ContentType="application/pdf")
        meta = {
            "week": week_num,
            "date": report_date_iso,
            "generated_at": datetime.now(timezone.utc).isoformat(),
        }
        s3.put_object(Bucket=bucket, Key=WBR_META_KEY,
                      Body=json.dumps(meta).encode(),
                      ContentType="application/json")
        logger.info(f"WBR PDF pushed to S3 (W{week_num})")
        return True
    except Exception as e:
        logger.warning(f"WBR PDF S3 push failed: {e}")
        return False


def pull_wbr_pdf_from_s3(aws_key: str, aws_secret: str,
                          region: str, bucket: str) -> tuple:
    """Download the latest WBR PDF + metadata from S3.
    Returns (pdf_bytes, meta_dict) or (None, {}) on failure."""
    try:
        s3 = _client(aws_key, aws_secret, region)
        pdf_obj  = s3.get_object(Bucket=bucket, Key=WBR_PDF_KEY)
        pdf_bytes = pdf_obj["Body"].read()
        try:
            meta_obj = s3.get_object(Bucket=bucket, Key=WBR_META_KEY)
            meta = json.loads(meta_obj["Body"].read())
        except Exception:
            meta = {}
        logger.info("WBR PDF pulled from S3")
        return pdf_bytes, meta
    except ClientError as e:
        if e.response["Error"]["Code"] == "NoSuchKey":
            logger.info("No WBR PDF in S3 yet")
        else:
            logger.warning(f"WBR PDF S3 pull failed: {e}")
    except Exception as e:
        logger.warning(f"WBR PDF S3 pull failed: {e}")
    return None, {}
