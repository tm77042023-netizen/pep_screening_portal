# Softdayta Risk

Standalone Flask application for subscriber-based PIP/PEP screening, adverse media review, relationship/RCA capture, and monitoring.

## Features

- Subscriber login and request submission.
- Single-name screening.
- Excel template download and bulk upload.
- Screening results with match scores and audit trail.
- Monitoring list for recurring/watchlist-style requests.
- Admin dashboard for full PEP database search.
- Admin CRUD for PEP/PIP records.
- Public source registry and manual update workflow.
- DailyNews PDF archive importer with temporary downloads and review-candidate creation.
- Seeded demo admin and subscriber accounts.

## Run Locally

```powershell
cd "C:\Users\HP\OneDrive\Documents\New project\pep_screening_portal"
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\python app.py
```

Then open:

```text
http://127.0.0.1:5055
```

The default SQLite database is stored under:

```text
%LOCALAPPDATA%\Softdayta Risk\pep_portal.db
```

Set `PEP_DATABASE_URI` if you want to use another database.

## Production Deployment (Linux VPS)

This repo includes example deployment templates under `deploy/`:

- `deploy/softdayta-risk.service` (systemd + gunicorn)
- `deploy/nginx_site.conf` (nginx reverse proxy)

Required environment variables:

```text
SECRET_KEY=<long-random-secret>
SOFTDAYTA_RISK_DATA_DIR=/var/lib/softdayta-risk
PEP_DATABASE_URI=sqlite:////var/lib/softdayta-risk/pep_portal.db   (or Postgres URI)
SEED_DEMO_USERS=0
```

## Demo Users

```text
Admin: admin@example.com / admin123
Subscriber: client@example.com / client123
```

Do not enable `SEED_DEMO_USERS` in production.
