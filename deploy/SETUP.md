# Deployment Guide — Robotics Container Tracker

## Prerequisites
- AWS account 844000647671 with console access
- SES already in production mode (confirmed)
- GitHub account to push the repo

---

## Step 1 — S3 bucket

In AWS Console → S3 → Create bucket:
- Name: `robotics-container-tracker`
- Region: us-east-1
- Block all public access: ON
- Versioning: optional but recommended

---

## Step 2 — IAM user for Streamlit Cloud

In AWS Console → IAM → Users → Create user: `robotics-tracker-reader`

Attach inline policy:
```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": ["s3:GetObject", "s3:PutObject", "s3:CopyObject", "s3:ListBucket"],
      "Resource": [
        "arn:aws:s3:::robotics-container-tracker",
        "arn:aws:s3:::robotics-container-tracker/*"
      ]
    }
  ]
}
```

Create access key → save in Streamlit Cloud secrets.

---

## Step 3 — SES verified email

In AWS Console → SES → Verified identities → Verify your sending email address.
(If using a custom domain for the inbound address, also verify the domain.)

---

## Step 4 — Email ingestion (optional but recommended)

### 4a. SES receiving rule
In SES → Email receiving → Create rule set:
- Condition: recipient = `containers@yourdomain.com`
- Action 1: S3 — save to `robotics-container-tracker/inbound/`
- Action 2: (Lambda trigger set up in 4b)

### 4b. Lambda — email_ingest
In AWS Console → Lambda → Create function:
- Name: `robotics-email-ingest`
- Runtime: Python 3.12
- Role: create new with S3 + SES permissions

Upload code:
```bash
cd lambda
zip email_ingest.zip email_ingest.py
# upload zip in Lambda console
```

Add Lambda Layer with openpyxl + pandas (or use a public layer ARN).

Environment variables:
```
S3_BUCKET        = robotics-container-tracker
DB_S3_KEY        = db/tracker.db
SES_FROM_EMAIL   = containers-noreply@yourdomain.com
```

Add S3 trigger: bucket = `robotics-container-tracker`, prefix = `inbound/`

---

## Step 5 — Lambda — empty return alerts

In Lambda → Create function:
- Name: `robotics-empty-return-alerts`
- Runtime: Python 3.12
- Same role as above + SES send permission

Upload `lambda/empty_return_alerts.py`.

Environment variables:
```
S3_BUCKET        = robotics-container-tracker
DBR_S3_KEY       = dbr/latest.xlsx
DB_S3_KEY        = db/tracker.db
SES_FROM_EMAIL   = containers-noreply@yourdomain.com
ALERT_RECIPIENTS = you@amazon.com,teammate@amazon.com
```

EventBridge trigger: `cron(0 13 * * ? *)` = 8 AM EST daily

---

## Step 6 — Streamlit Cloud deploy

1. Push this repo to GitHub (private repo)
2. Go to share.streamlit.io → New app
3. Select repo → branch `main` → main file: `app.py`
4. Advanced → Secrets → paste contents of `.streamlit/secrets.toml` with real values
5. Deploy

App will be live at `https://your-app-name.streamlit.app`
Share that URL with internal partners (they'll need the password).
Carriers get the same URL for the submission tab.

---

## Summary of AWS resources created

| Resource | Name | Purpose |
|---|---|---|
| S3 bucket | robotics-container-tracker | DBR storage, DB, inbound email, logs |
| IAM user | robotics-tracker-reader | Streamlit Cloud S3 access |
| Lambda | robotics-email-ingest | Parse carrier emails → DB |
| Lambda | robotics-empty-return-alerts | Daily overdue/due-soon alerts |
| SES rule | inbound email | Route carrier emails to S3 |
| EventBridge rule | daily 8am | Trigger alert Lambda |
