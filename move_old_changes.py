#!/usr/bin/env python3
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv

load_dotenv()

SERVICE_ACCOUNT_JSON = os.getenv("SERVICE_ACCOUNT_JSON", "service-account-creds.json")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
if not SPREADSHEET_ID:
    raise RuntimeError("Missing SPREADSHEET_ID environment variable")

creds = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_JSON,
    scopes = ["https://www.googleapis.com/auth/spreadsheets"],
)
gc      = gspread.authorize(creds)
spread  = gc.open_by_key(SPREADSHEET_ID)

recent_ws     = spread.worksheet("Recent Changes")
historical_ws = spread.worksheet("Historical")

header = recent_ws.row_values(1)
try:
    ts_col = header.index("Change Timestamp") + 1
except ValueError:
    raise RuntimeError("Column header 'Change Timestamp' not found")

now_ro    = datetime.now(timezone(timedelta(hours=3)))
cutoff    = now_ro - timedelta(days=14)

all_rows = recent_ws.get_all_values()[1:]  # skip header

rows_to_move = []      # full row values
rows_to_kill = []      # 1-based row indices in Recent Changes

for idx, row in enumerate(all_rows, start=2):    # start=2 because header is row 1
    ts_str = row[ts_col - 1].strip()
    if not ts_str:
        continue                                # blank timestamp → ignore
    try:
        ts = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=timezone.utc)
    except ValueError:
        try:
            ts = datetime.strptime(ts_str, "%m/%d/%Y %H:%M:%S")
            ts = ts.replace(tzinfo=timezone.utc)
        except ValueError:
            print(f"⚠️  Couldn’t parse timestamp “{ts_str}” on row {idx}")
            continue

    if ts < cutoff:
        rows_to_move.append(row)
        rows_to_kill.append(idx)

if not rows_to_move:
    print("No rows older than 14 days; sheet already up-to-date.")
    exit(0)

# Sort the new rows by timestamp in descending order and insert at top
def extract_timestamp(row):
    try:
        ts_str = row[ts_col - 1].strip()
        if ts_str.endswith("Z"):
            ts = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        else:
            try:
                ts = datetime.fromisoformat(ts_str)
            except ValueError:
                ts = datetime.strptime(ts_str, "%m/%d/%Y %H:%M:%S")
        return ts
    except Exception:
        return datetime.min

rows_to_move.sort(key=extract_timestamp, reverse=True)
# Insert rows below the header (row 2) so they appear at the top of Historical
historical_ws.insert_rows(rows_to_move, row=2, value_input_option="USER_ENTERED")

num_cols = len(header)
clear_ranges = [
    f"A{row}:{gspread.utils.rowcol_to_a1(row, num_cols)[1:]}" for row in rows_to_kill
]

recent_ws.batch_clear([
    f"A{row}:{gspread.utils.rowcol_to_a1(row, num_cols)[1:]}"
    for row in rows_to_kill
])

all_data = recent_ws.get_all_values()[1:]
non_empty_rows = [row for row in all_data if any(cell.strip() for cell in row)]

recent_ws.batch_clear([f"A2:{gspread.utils.rowcol_to_a1(len(all_data) + 1, num_cols)[1:]}"])

if non_empty_rows:
    recent_ws.update(f"A2", non_empty_rows)

print(f"Moved {len(rows_to_move)} row(s) to 'Historical'.")