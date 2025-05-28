import sqlite3
import csv
import os
from datetime import datetime, timedelta

# ---- USER VARIABLE: Set this to your Calendar.sqlitedb file (can be a backup copy) ----
db_path = '/Users/phobrla/Documents/Calendar.sqlitedb'
output_csv = os.path.expanduser('~/Documents/calendar_event_extremes.csv')

def apple_core_date_to_datetime(ts):
    """Convert Apple Core Data timestamp (seconds since 2001-01-01) to datetime."""
    if ts is None:
        return ''
    try:
        return datetime(2001, 1, 1) + timedelta(seconds=float(ts))
    except Exception:
        return ''

def iso(dt):
    if not dt or isinstance(dt, str):
        return ''
    return dt.strftime('%Y-%m-%dT%H:%M:%S')

# Generate export timestamp
exported_at = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')

# Connect to the database (read-only mode)
conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
cur = conn.cursor()

# Query for all calendars (get ROWID, title, owner_identity_email)
cur.execute("SELECT ROWID, title, owner_identity_email FROM Calendar")
calendars = cur.fetchall()

rows = []

for cal_id, cal_title, cal_owner_email in calendars:
    # Query all events for this calendar
    cur.execute('''
        SELECT summary, start_date
        FROM CalendarItem
        WHERE calendar_id = ?
        AND start_date IS NOT NULL
    ''', (cal_id,))
    events = cur.fetchall()
    
    if not events:
        # No events for this calendar
        calendar_field = f"{cal_owner_email or ''} — {cal_title}"
        rows.append([calendar_field, '', '', '', '', exported_at])
        continue

    # Sort events by start date (start_date is Apple Core Data timestamp)
    events_sorted = sorted(events, key=lambda x: x[1])
    earliest = events_sorted[0]
    latest = events_sorted[-1]

    early_dt = apple_core_date_to_datetime(earliest[1])
    late_dt = apple_core_date_to_datetime(latest[1])
    early_summary = earliest[0] or ''
    late_summary = latest[0] or ''

    calendar_field = f"{cal_owner_email or ''} — {cal_title}"

    rows.append([
        calendar_field,
        iso(early_dt),
        early_summary,
        iso(late_dt),
        late_summary,
        exported_at
    ])

# Write to CSV
with open(output_csv, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['Calendar', 'Earliest Timestamp', 'Earliest Name', 'Latest Timestamp', 'Latest Name', 'Exported At'])
    writer.writerows(rows)

print(f"Exported to {output_csv}")

# Clean up
conn.close()
