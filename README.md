# Billable Hours Tracker

Flask application for weekly consultant billable-hours collection by SMS.

The app sends weekly Twilio prompts, receives consultant replies at `/sms`, stores hours in SQLite, and provides a password-protected dashboard with weekly/monthly exports.

## Production Status

Ready for a small production pilot after the Twilio production setup is complete.

Production pieces already in the app:
- Dashboard Basic Auth using `ADMIN_USERNAME` and `ADMIN_PASSWORD`
- Twilio webhook signature validation using `VALIDATE_TWILIO_WEBHOOKS=true`
- Weekly and monthly CSV exports
- Monthly Excel export with summary, detail, and chart sheets
- Month-split weekly reporting segments, for example `2026-W18A` and `2026-W18B`
- UTC scheduler for weekly prompts and monthly report generation
- Optional Sentry error reporting
- `/healthz` endpoint for host health checks
- Gunicorn `Procfile` for Railway/Render-style deployment

## Twilio Production Setup

For U.S. texting from a 10DLC Twilio number, complete A2P 10DLC registration before production use.

Recommended Twilio steps:
1. Upgrade the Twilio account and add billing.
2. Buy or keep one production SMS-capable U.S. number.
3. Register the Brand in Twilio Trust Hub.
4. Register the Campaign/use case for weekly operational consultant messages.
5. Include opt-in, opt-out, and help language in the campaign submission.
6. After approval, attach the Twilio number to the approved Messaging Service or campaign.
7. Set the inbound messaging webhook to:

```text
https://your-production-domain/sms
```

Twilio currently requires A2P registration for application-to-person SMS to U.S. recipients over 10DLC numbers. Twilio documentation says campaign review can take 10-15 days during busy periods.

## Environment Variables

Use `.env.production.example` as the deployment checklist.

Required:
- `TWILIO_ACCOUNT_SID`
- `TWILIO_AUTH_TOKEN`
- `TWILIO_MESSAGING_SERVICE_SID` for the A2P-linked Messaging Service, or `TWILIO_PHONE` for direct number sending
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `PUBLIC_BASE_URL`
- `VALIDATE_TWILIO_WEBHOOKS=true`
- `DB_PATH`
- `REPORT_OUTPUT_DIR`

Optional:
- `SENTRY_DSN`
- `SENTRY_TRACES_SAMPLE_RATE`
- `WEEKLY_PROMPT_UTC_HOUR`
- `MONTHLY_REPORT_UTC_HOUR`

Scheduler values are UTC. Keep the host cron/scheduler configuration in UTC too.

## Local Development

Install dependencies:

```bash
pip install -r requirements-dev.txt
```

Create local environment:

```bash
copy .env.example .env
```

Run locally:

```bash
python app.py
```

Local app URL:

```text
http://127.0.0.1:8000
```

For local Twilio webhook testing, expose the app with ngrok:

```bash
ngrok http 8000
```

Then set the Twilio inbound webhook to:

```text
https://your-ngrok-domain/sms
```

If webhook validation is enabled locally, set `PUBLIC_BASE_URL` to the exact ngrok domain.

## Production Deployment

Chosen low-cost deployment:
- Host: Render
- Runtime: Python
- Start command: from `Procfile`
- Database: SQLite on a persistent mounted volume for the initial 50-consultant pilot
- Monitoring: Sentry free tier
- SMS: Twilio 10DLC with approved A2P campaign

This repo includes `render.yaml` for a Render Blueprint deploy.

Render settings used by the Blueprint:
- Service type: Web Service
- Runtime: Python
- Branch: `main`
- Plan: `starter`
- Region: `ohio`
- Health check path: `/healthz`
- Persistent disk: `/data`
- SQLite database path: `/data/hours.db`
- Report output path: `/data/reports`

Production command:

```bash
gunicorn --workers 1 --threads 4 --timeout 120 --bind 0.0.0.0:$PORT app:app
```

Use one Gunicorn worker because APScheduler runs inside the Flask process. Multiple workers would start duplicate scheduled jobs.

Set persistent paths on the host:

```text
DB_PATH=/data/hours.db
REPORT_OUTPUT_DIR=/data/reports
```

Use the actual mounted volume path from your host if it differs from `/data`.

### Deploy on Render

1. Push `main` to GitHub.
2. In Render, choose **New > Blueprint**.
3. Connect the GitHub repo.
4. Select the repo root `render.yaml`.
5. Fill the `sync: false` environment variables in Render:
   - `TWILIO_ACCOUNT_SID`
   - `TWILIO_AUTH_TOKEN`
   - `TWILIO_MESSAGING_SERVICE_SID` preferred for A2P, or `TWILIO_PHONE`
   - `ADMIN_PASSWORD`
   - `PUBLIC_BASE_URL`
   - `SENTRY_DSN` if using Sentry
6. Deploy the service.
7. After Render gives the live URL, set `PUBLIC_BASE_URL` to that exact URL.
8. In Twilio, set the inbound webhook for the production number or Messaging Service to:

```text
https://your-render-service.onrender.com/sms
```

Render services have an ephemeral filesystem by default, so the persistent disk is required while using SQLite. Render's disk docs say only files written under the mount path are preserved across deploys and restarts.

## Smoke Test After Deploy

1. Open `https://your-production-domain/healthz`
2. Confirm it returns `"twilio_configured": true`
3. Open the dashboard and confirm Basic Auth is required
4. Add one internal test consultant number
5. Send one manual weekly prompt
6. Reply from the phone
7. Confirm the dashboard shows the response
8. Download weekly CSV
9. Download monthly CSV and monthly Excel
10. Check Twilio logs for delivery errors
11. Check Sentry for errors

## A2P SMS Legal URLs

Twilio A2P campaign review requires public URLs for SMS terms and SMS privacy disclosures.
After deployment, use these app-hosted pages:

```text
https://your-production-domain/sms-terms
https://your-production-domain/sms-privacy
https://your-production-domain/sms-opt-in
```

For the Render service used by this project:

```text
https://billable-hours-tracker.onrender.com/sms-terms
https://billable-hours-tracker.onrender.com/sms-privacy
https://billable-hours-tracker.onrender.com/sms-opt-in
```

## Tests

Run:

```bash
pytest -q
```

The tests cover core parsing, current week calculation, SMS reply handling, exports, split-month week reporting, and the health check.
