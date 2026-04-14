"""
Billable Hours Tracker
- Sends weekly SMS to consultants asking for billable hours
- Receives replies and stores them in SQLite
- Provides a dashboard to view/export data
"""

import csv
import io
import os
import re
import sqlite3
from datetime import datetime, timedelta

from apscheduler.schedulers.background import BackgroundScheduler
from flask import Flask, jsonify, render_template, request, send_file
from twilio.base.exceptions import TwilioRestException
from twilio.rest import Client
from twilio.twiml.messaging_response import MessagingResponse

app = Flask(__name__)

MESSAGE_TYPE_LABELS = {
    "weekly_prompt": "Weekly prompt",
    "reminder": "Reminder",
    "manual_follow_up": "Manual follow-up",
    "response": "Consultant reply",
}

# ─── Config (set these as environment variables) ───────────────────────────────
TWILIO_ACCOUNT_SID = os.environ.get("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN  = os.environ.get("TWILIO_AUTH_TOKEN")
TWILIO_PHONE       = os.environ.get("TWILIO_PHONE")   # e.g. "+16175551234"
DB_PATH            = os.environ.get("DB_PATH", "hours.db")
PORT               = int(os.environ.get("PORT", "8000"))

# ─── Helpers ───────────────────────────────────────────────────────────────────

def current_week_monday():
    """Return the ISO date string for the active reporting cycle Monday."""
    configured_cycle = get_reporting_cycle_date()
    if configured_cycle:
        return configured_cycle
    today = datetime.today()
    monday = today - timedelta(days=today.weekday())
    return monday.strftime("%Y-%m-%d")

def validate_reporting_cycle_date(value: str):
    """Validate a YYYY-MM-DD date and normalize it to that week's Monday."""
    if not value:
        raise ValueError("reporting_cycle_date is required")
    parsed = datetime.strptime(value, "%Y-%m-%d")
    monday = parsed - timedelta(days=parsed.weekday())
    return monday.strftime("%Y-%m-%d")

def reporting_period_parts(week_of: str):
    parsed = datetime.strptime(week_of, "%Y-%m-%d")
    iso_year, iso_week, _ = parsed.isocalendar()
    return {
        "reporting_year": iso_year,
        "reporting_week": iso_week,
        "reporting_week_label": f"{iso_year}-W{iso_week:02d}",
    }

def normalize_phone(phone: str):
    """Normalize phone values to a simple E.164-like format."""
    digits = re.sub(r"\D", "", phone or "")
    if not digits:
        return ""
    if len(digits) == 10:
        digits = f"1{digits}"
    return f"+{digits}"

def parse_hours(text: str):
    """Extract a number from a free-form reply like '32', 'about 40', '~35.5'."""
    match = re.search(r'\d+(\.\d+)?', text)
    return float(match.group()) if match else None

# ─── Database ──────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def ensure_week_metadata_columns(db, table_name: str):
    columns = {row["name"] for row in db.execute(f"PRAGMA table_info({table_name})").fetchall()}
    desired_columns = {
        "reporting_year": "INTEGER",
        "reporting_week": "INTEGER",
        "reporting_week_label": "TEXT",
    }
    for column_name, column_type in desired_columns.items():
        if column_name not in columns:
            db.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")

def backfill_week_metadata(db, table_name: str):
    rows = db.execute(
        f"""
        SELECT id, week_of
        FROM {table_name}
        WHERE reporting_year IS NULL
           OR reporting_week IS NULL
           OR reporting_week_label IS NULL
        """
    ).fetchall()
    for row in rows:
        parts = reporting_period_parts(row["week_of"])
        db.execute(
            f"""
            UPDATE {table_name}
            SET reporting_year = ?,
                reporting_week = ?,
                reporting_week_label = ?
            WHERE id = ?
            """,
            (
                parts["reporting_year"],
                parts["reporting_week"],
                parts["reporting_week_label"],
                row["id"],
            )
        )

def init_db():
    with get_db() as db:
        db.executescript("""
            CREATE TABLE IF NOT EXISTS consultants (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                name      TEXT NOT NULL,
                phone     TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS responses (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                consultant_id INTEGER NOT NULL,
                week_of       TEXT NOT NULL,   -- ISO date of Monday
                reporting_year INTEGER,
                reporting_week INTEGER,
                reporting_week_label TEXT,
                hours         REAL,
                raw_reply     TEXT,
                received_at   TEXT NOT NULL,
                FOREIGN KEY (consultant_id) REFERENCES consultants(id),
                UNIQUE(consultant_id, week_of)   -- one entry per person per week
            );

            CREATE TABLE IF NOT EXISTS sent_log (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                consultant_id INTEGER NOT NULL,
                week_of       TEXT NOT NULL,
                reporting_year INTEGER,
                reporting_week INTEGER,
                reporting_week_label TEXT,
                sent_at       TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS message_log (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                consultant_id INTEGER NOT NULL,
                week_of       TEXT NOT NULL,
                reporting_year INTEGER,
                reporting_week INTEGER,
                reporting_week_label TEXT,
                message_type  TEXT NOT NULL,
                body          TEXT NOT NULL,
                status        TEXT NOT NULL,
                error_message TEXT,
                twilio_sid    TEXT,
                sent_at       TEXT NOT NULL,
                FOREIGN KEY (consultant_id) REFERENCES consultants(id)
            );

            CREATE TABLE IF NOT EXISTS app_settings (
                key   TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
        """)
        ensure_week_metadata_columns(db, "responses")
        ensure_week_metadata_columns(db, "sent_log")
        ensure_week_metadata_columns(db, "message_log")
        consultants = db.execute("SELECT id, phone FROM consultants").fetchall()
        for consultant in consultants:
            normalized_phone = normalize_phone(consultant["phone"])
            if normalized_phone and normalized_phone != consultant["phone"]:
                db.execute(
                    "UPDATE consultants SET phone = ? WHERE id = ?",
                    (normalized_phone, consultant["id"])
                )
        backfill_week_metadata(db, "responses")
        backfill_week_metadata(db, "sent_log")
        backfill_week_metadata(db, "message_log")

init_db()

def get_reporting_cycle_date():
    with get_db() as db:
        row = db.execute(
            "SELECT value FROM app_settings WHERE key = ?",
            ("reporting_cycle_date",)
        ).fetchone()
    return row["value"] if row else None

def set_reporting_cycle_date(value: str):
    normalized_value = validate_reporting_cycle_date(value)
    with get_db() as db:
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            ("reporting_cycle_date", normalized_value)
        )
    return normalized_value

def send_sms(to: str, body: str):
    client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
    return client.messages.create(
        to=to,
        from_=TWILIO_PHONE,
        body=body
    )

def default_reminder_message(name: str, week_of: str):
    pretty_date = datetime.strptime(week_of, "%Y-%m-%d").strftime("%B %-d, %Y")
    return (
        f"Hi {name}, quick follow-up on your billable hours forecast for the week of {pretty_date}. "
        "Please reply with the number of hours you expect to log."
    )

def log_message_attempt(consultant_id: int, week_of: str, message_type: str, body: str, status: str,
                        twilio_sid: str = None, error_message: str = None, sent_at: str = None):
    sent_at = sent_at or datetime.utcnow().isoformat()
    period = reporting_period_parts(week_of)
    with get_db() as db:
        db.execute(
            """
            INSERT INTO message_log
                (
                    consultant_id, week_of, reporting_year, reporting_week, reporting_week_label,
                    message_type, body, status, error_message, twilio_sid, sent_at
                )
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                consultant_id,
                week_of,
                period["reporting_year"],
                period["reporting_week"],
                period["reporting_week_label"],
                message_type,
                body,
                status,
                error_message,
                twilio_sid,
                sent_at,
            )
        )

def send_consultant_message(consultant, body: str, week_of: str, message_type: str):
    try:
        msg = send_sms(consultant["phone"], body)
        sent_at = datetime.utcnow().isoformat()
        log_message_attempt(
            consultant["id"],
            week_of,
            message_type,
            body,
            "sent",
            twilio_sid=msg.sid,
            sent_at=sent_at
        )
        if message_type == "weekly_prompt":
            period = reporting_period_parts(week_of)
            with get_db() as db:
                db.execute(
                    """
                    INSERT OR IGNORE INTO sent_log
                        (consultant_id, week_of, reporting_year, reporting_week, reporting_week_label, sent_at)
                    VALUES (?,?,?,?,?,?)
                    """,
                    (
                        consultant["id"],
                        week_of,
                        period["reporting_year"],
                        period["reporting_week"],
                        period["reporting_week_label"],
                        sent_at,
                    )
                )
        return {"ok": True, "sid": msg.sid}
    except TwilioRestException as exc:
        error_message = exc.msg
    except Exception as exc:
        error_message = str(exc)

    log_message_attempt(
        consultant["id"],
        week_of,
        message_type,
        body,
        "failed",
        error_message=error_message
    )
    return {"ok": False, "error": error_message}

def get_consultant_by_id(cid: int):
    with get_db() as db:
        consultant = db.execute("SELECT * FROM consultants WHERE id = ?", (cid,)).fetchone()
    return consultant

def get_weekly_reporting_status(week_of: str):
    with get_db() as db:
        consultants = db.execute("SELECT * FROM consultants ORDER BY name").fetchall()
        responses = db.execute(
            """
            SELECT consultant_id, hours, raw_reply, received_at
            FROM responses
            WHERE week_of = ?
            """,
            (week_of,)
        ).fetchall()
        messages = db.execute(
            """
            SELECT ml.*
            FROM message_log ml
            JOIN (
                SELECT consultant_id, MAX(id) AS max_id
                FROM message_log
                WHERE week_of = ?
                GROUP BY consultant_id
            ) latest ON latest.max_id = ml.id
            """,
            (week_of,)
        ).fetchall()

    responses_by_consultant = {row["consultant_id"]: dict(row) for row in responses}
    messages_by_consultant = {row["consultant_id"]: dict(row) for row in messages}
    items = []
    for consultant in consultants:
        response = responses_by_consultant.get(consultant["id"])
        message = messages_by_consultant.get(consultant["id"])
        if response:
            status_key = "replied"
            status_label = "Replied"
        elif message and message["status"] == "failed":
            status_key = "failed"
            status_label = "Failed to send"
        elif message and message["message_type"] in ("reminder", "manual_follow_up"):
            status_key = "needs_follow_up"
            status_label = "Needs follow-up"
        elif message:
            status_key = "pending"
            status_label = "Pending"
        else:
            status_key = "pending"
            status_label = "Pending"

        items.append({
            "consultant_id": consultant["id"],
            "name": consultant["name"],
            "phone": consultant["phone"],
            "week_of": week_of,
            "hours": response["hours"] if response else None,
            "raw_reply": response["raw_reply"] if response else None,
            "received_at": response["received_at"] if response else None,
            "status": status_key,
            "status_label": status_label,
            "last_message_type": message["message_type"] if message else None,
            "last_message_status": message["status"] if message else None,
            "last_message_at": message["sent_at"] if message else None,
            "last_error": message["error_message"] if message else None,
        })
    return items

def get_history_for_consultant(consultant_id: int, week_of: str):
    with get_db() as db:
        message_rows = db.execute(
            """
            SELECT id, message_type, body, status, error_message, sent_at, twilio_sid
            FROM message_log
            WHERE consultant_id = ? AND week_of = ?
            ORDER BY sent_at DESC, id DESC
            """,
            (consultant_id, week_of)
        ).fetchall()
        response_rows = db.execute(
            """
            SELECT id, hours, raw_reply, received_at
            FROM responses
            WHERE consultant_id = ? AND week_of = ?
            ORDER BY received_at DESC, id DESC
            """,
            (consultant_id, week_of)
        ).fetchall()

    items = [
        {
            "kind": "message",
            "label": MESSAGE_TYPE_LABELS.get(row["message_type"], row["message_type"]),
            "message_type": row["message_type"],
            "body": row["body"],
            "status": row["status"],
            "error_message": row["error_message"],
            "timestamp": row["sent_at"],
            "twilio_sid": row["twilio_sid"],
        }
        for row in message_rows
    ]
    items.extend(
        {
            "kind": "response",
            "label": MESSAGE_TYPE_LABELS["response"],
            "hours": row["hours"],
            "body": row["raw_reply"],
            "status": "received",
            "timestamp": row["received_at"],
        }
        for row in response_rows
    )
    return sorted(items, key=lambda item: item["timestamp"], reverse=True)

# ─── Weekly job ────────────────────────────────────────────────────────────────

def send_weekly_texts():
    """Runs every Monday morning. Sends the question to all consultants."""
    week_of = current_week_monday()
    results = {"sent": 0, "failed": []}
    with get_db() as db:
        consultants = db.execute("SELECT * FROM consultants").fetchall()
    for c in consultants:
        body = (
            f"Hi {c['name']}! 👋 How many billable hours will you log this week? "
            f"(Reply with just a number, e.g. 40)"
        )
        result = send_consultant_message(c, body, week_of, "weekly_prompt")
        if result["ok"]:
            results["sent"] += 1
        else:
            message = f"{c['name']} ({c['phone']}): {result['error']}"
            print(f"Failed to text {message}")
            results["failed"].append(message)
    return results

# Schedule for every Monday at 9:00 AM
scheduler = BackgroundScheduler()
scheduler.add_job(send_weekly_texts, "cron", day_of_week="mon", hour=9, minute=0)
scheduler.start()

# ─── Twilio webhook (handles inbound replies) ──────────────────────────────────

@app.route("/sms", methods=["POST"])
def sms_reply():
    from_number = normalize_phone(request.form.get("From", "").strip())
    body        = request.form.get("Body", "").strip()
    week_of     = current_week_monday()
    hours       = parse_hours(body)

    resp = MessagingResponse()

    with get_db() as db:
        consultant = db.execute(
            "SELECT * FROM consultants WHERE phone = ?", (from_number,)
        ).fetchone()

    if not consultant:
        resp.message("Sorry, your number isn't registered. Ask your manager to add you.")
        return str(resp)

    with get_db() as db:
        period = reporting_period_parts(week_of)
        db.execute(
            """INSERT INTO responses
               (consultant_id, week_of, reporting_year, reporting_week, reporting_week_label, hours, raw_reply, received_at)
               VALUES (?,?,?,?,?,?,?,?)
               ON CONFLICT(consultant_id, week_of) DO UPDATE SET
                   reporting_year=excluded.reporting_year,
                   reporting_week=excluded.reporting_week,
                   reporting_week_label=excluded.reporting_week_label,
                   hours=excluded.hours,
                   raw_reply=excluded.raw_reply,
                   received_at=excluded.received_at""",
            (
                consultant["id"],
                week_of,
                period["reporting_year"],
                period["reporting_week"],
                period["reporting_week_label"],
                hours,
                body,
                datetime.utcnow().isoformat(),
            )
        )
    log_message_attempt(
        consultant["id"],
        week_of,
        "response",
        body,
        "received",
        sent_at=datetime.utcnow().isoformat()
    )

    if hours is not None:
        resp.message(f"Got it — {hours:.0f} hours logged for the week of {week_of}. Thanks!")
    else:
        resp.message("Hmm, I didn't catch a number. Please reply with just digits, e.g. 40.")

    return str(resp)

# ─── REST API ─────────────────────────────────────────────────────────────────

@app.route("/api/consultants", methods=["GET"])
def list_consultants():
    with get_db() as db:
        rows = db.execute("SELECT * FROM consultants ORDER BY name").fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/consultants", methods=["POST"])
def add_consultant():
    data = request.json
    name  = data.get("name", "").strip()
    phone = normalize_phone(data.get("phone", "").strip())
    if not name or not phone:
        return jsonify({"error": "name and phone required"}), 400
    try:
        with get_db() as db:
            db.execute("INSERT INTO consultants (name, phone) VALUES (?,?)", (name, phone))
    except sqlite3.IntegrityError:
        return jsonify({"error": "consultant already exists for that phone number"}), 409
    return jsonify({"ok": True}), 201

@app.route("/api/consultants/<int:cid>", methods=["DELETE"])
def remove_consultant(cid):
    with get_db() as db:
        db.execute("DELETE FROM consultants WHERE id=?", (cid,))
    return jsonify({"ok": True})

@app.route("/api/responses", methods=["GET"])
def list_responses():
    weeks = request.args.get("weeks", 8, type=int)
    cutoff = (datetime.today() - timedelta(weeks=weeks)).strftime("%Y-%m-%d")
    with get_db() as db:
        rows = db.execute("""
            SELECT r.week_of, c.name, r.hours, r.raw_reply, r.received_at
                 , r.reporting_year, r.reporting_week, r.reporting_week_label
            FROM responses r JOIN consultants c ON c.id = r.consultant_id
            WHERE r.week_of >= ?
            ORDER BY r.week_of DESC, c.name
        """, (cutoff,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/reporting-status", methods=["GET"])
def reporting_status():
    week_of = request.args.get("week_of", type=str) or current_week_monday()
    try:
        normalized_week = validate_reporting_cycle_date(week_of)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify(get_weekly_reporting_status(normalized_week))

@app.route("/api/summary", methods=["GET"])
def weekly_summary():
    """Total hours per week + per-consultant breakdown."""
    week_of = request.args.get("week_of", type=str)
    if week_of:
        try:
            normalized_week = validate_reporting_cycle_date(week_of)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        with get_db() as db:
            row = db.execute("""
                SELECT r.week_of,
                       r.reporting_year,
                       r.reporting_week,
                       r.reporting_week_label,
                       COUNT(r.id)    AS responses,
                       SUM(r.hours)   AS total_hours,
                       AVG(r.hours)   AS avg_hours
                FROM responses r
                WHERE r.week_of = ?
                GROUP BY r.week_of
            """, (normalized_week,)).fetchone()
        return jsonify([dict(row)] if row else [])

    weeks = request.args.get("weeks", 8, type=int)
    cutoff = (datetime.today() - timedelta(weeks=weeks)).strftime("%Y-%m-%d")
    with get_db() as db:
        rows = db.execute("""
            SELECT r.week_of,
                   r.reporting_year,
                   r.reporting_week,
                   r.reporting_week_label,
                   COUNT(r.id)    AS responses,
                   SUM(r.hours)   AS total_hours,
                   AVG(r.hours)   AS avg_hours
            FROM responses r
            WHERE r.week_of >= ?
            GROUP BY r.week_of
            ORDER BY r.week_of DESC
        """, (cutoff,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/reporting-cycle", methods=["GET"])
def get_reporting_cycle():
    cycle_date = current_week_monday()
    return jsonify({"reporting_cycle_date": cycle_date})

@app.route("/api/reporting-cycle", methods=["POST"])
def update_reporting_cycle():
    data = request.json or {}
    try:
        cycle_date = set_reporting_cycle_date(data.get("reporting_cycle_date", "").strip())
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({
        "ok": True,
        "reporting_cycle_date": cycle_date,
        "message": f"Reporting cycle updated to {cycle_date}."
    })

@app.route("/api/send-now", methods=["POST"])
def send_now():
    """Manually trigger the weekly text blast (for testing)."""
    results = send_weekly_texts()
    if results["failed"]:
        details = " | ".join(results["failed"][:3])
        return jsonify({
            "ok": False,
            "message": f"Sent {results['sent']} message(s). Failed: {details}"
        }), 502
    return jsonify({"ok": True, "message": f"Sent {results['sent']} text(s)."})

@app.route("/api/messages/send", methods=["POST"])
def send_follow_up():
    data = request.json or {}
    consultant_id = data.get("consultant_id")
    body = (data.get("body") or "").strip()
    message_type = (data.get("message_type") or "manual_follow_up").strip()
    week_of = data.get("week_of", "").strip() or current_week_monday()

    if not consultant_id:
        return jsonify({"error": "consultant_id is required"}), 400
    if not body:
        return jsonify({"error": "body is required"}), 400

    try:
        week_of = validate_reporting_cycle_date(week_of)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    consultant = get_consultant_by_id(int(consultant_id))
    if not consultant:
        return jsonify({"error": "consultant not found"}), 404

    result = send_consultant_message(consultant, body, week_of, message_type)
    if not result["ok"]:
        return jsonify({"ok": False, "message": result["error"]}), 502
    return jsonify({"ok": True, "message": f"Message sent to {consultant['name']}."})

@app.route("/api/messages/bulk-remind", methods=["POST"])
def bulk_remind():
    data = request.json or {}
    week_of = data.get("week_of", "").strip() or current_week_monday()
    try:
        week_of = validate_reporting_cycle_date(week_of)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    statuses = get_weekly_reporting_status(week_of)
    pending = [item for item in statuses if item["status"] in ("pending", "needs_follow_up", "failed")]
    sent = 0
    failed = []

    for item in pending:
        consultant = get_consultant_by_id(item["consultant_id"])
        body = default_reminder_message(consultant["name"], week_of)
        result = send_consultant_message(consultant, body, week_of, "reminder")
        if result["ok"]:
            sent += 1
        else:
            failed.append(f"{consultant['name']}: {result['error']}")

    if failed:
        details = " | ".join(failed[:3])
        return jsonify({
            "ok": False,
            "message": f"Sent {sent} reminder(s). Failed: {details}"
        }), 502
    return jsonify({"ok": True, "message": f"Sent {sent} reminder(s)."})

@app.route("/api/consultants/<int:cid>/history", methods=["GET"])
def consultant_history(cid):
    week_of = request.args.get("week_of", type=str) or current_week_monday()
    try:
        week_of = validate_reporting_cycle_date(week_of)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    consultant = get_consultant_by_id(cid)
    if not consultant:
        return jsonify({"error": "consultant not found"}), 404

    return jsonify({
        "consultant": {
            "id": consultant["id"],
            "name": consultant["name"],
            "phone": consultant["phone"],
        },
        "week_of": week_of,
        "history": get_history_for_consultant(cid, week_of),
    })

@app.route("/api/export", methods=["GET"])
def export_csv():
    with get_db() as db:
        rows = db.execute("""
            SELECT r.week_of, c.name, c.phone, r.hours, r.raw_reply, r.received_at
                 , r.reporting_year, r.reporting_week, r.reporting_week_label
            FROM responses r JOIN consultants c ON c.id = r.consultant_id
            ORDER BY r.week_of DESC, c.name
        """).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "week_of", "reporting_year", "reporting_week", "reporting_week_label",
        "name", "phone", "hours", "raw_reply", "received_at"
    ])
    for row in rows:
        writer.writerow(list(row))

    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode()),
        mimetype="text/csv",
        as_attachment=True,
        download_name="billable_hours.csv"
    )

# ─── Dashboard ────────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    return render_template("dashboard.html")

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=PORT)
