# Billable Hours Tracker

Sends a weekly SMS to your consulting team asking for their billable hour forecast.
Collects replies and shows them in a dashboard. Exports to CSV.

## Prototype Status

This repository is a working prototype/MVP, not a production-hardened application.

Current prototype capabilities:
- Sends weekly SMS prompts through Twilio
- Receives replies through a Twilio webhook
- Stores responses in SQLite
- Provides a local admin dashboard for reporting, reminders, and exports

Not yet production-ready:
- No authentication or role-based access control
- Uses SQLite instead of a production database
- Limited error monitoring and operational logging
- Local environment and Twilio configuration are expected

Before handing off to a production developer:
- Keep `.env`, `hours.db`, and local virtual environment files out of version control
- Review Twilio credentials, webhook configuration, and deployment approach
- Plan for hosting, monitoring, backups, and security hardening

## How It Works

1. Every Monday at 9 AM, every consultant gets a text:
   > "Hi Sarah! 👋 How many billable hours will you log this week? (Reply with just a number, e.g. 40)"
2. They reply with a number. The app parses it and stores it.
3. You check the dashboard at `http://your-server/` to see totals and per-person breakdowns.

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Set up Twilio
- Sign up at https://twilio.com (free trial available)
- Buy a phone number (~$1/month)
- Get your Account SID and Auth Token from the Twilio console

### 3. Configure environment variables
```bash
cp .env.example .env
# Edit .env with your Twilio credentials
```

### 4. Run the app
```bash
# Load env vars and start
export $(cat .env | xargs)
python app.py
```

The app runs on http://localhost:5000

### 5. Expose to the internet (so Twilio can reach your webhook)

For local development, use [ngrok](https://ngrok.com):
```bash
ngrok http 5000
# Copy the https URL, e.g. https://abc123.ngrok.io
```

In your Twilio console → Phone Numbers → your number → Messaging:
- Set "A message comes in" → Webhook → `https://abc123.ngrok.io/sms`

For production, deploy to [Railway](https://railway.app) or [Render](https://render.com) (both have free tiers).

## Dashboard Features

- **Consultants panel** — Add/remove team members by name + phone number
- **Weekly tabs** — Browse responses by week
- **Stats** — Total hours, # of responses, average per person
- **Send Now** — Trigger the text blast manually (great for testing)
- **Export CSV** — Download all data as a spreadsheet

## Database

SQLite file (`hours.db`) with three tables:
- `consultants` — roster of names + phone numbers
- `responses` — one row per person per week (week_of, hours, raw reply)
- `sent_log` — record of every outbound message

## Deploying to Railway (recommended)

1. Push this folder to a GitHub repo
2. Go to railway.app → New Project → Deploy from GitHub
3. Add environment variables in the Railway dashboard
4. Railway gives you a public URL — point your Twilio webhook there
5. Done. It'll run forever on their free tier for small teams.
