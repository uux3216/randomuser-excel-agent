import json
import os
import sys
import time
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

API_URL = "https://randomuser.me/api/1.4/"
COUNT = 398
TITLE = "RandomUser Threaded Data"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
SHARE_WITH_EMAIL = os.environ["SHARE_WITH_EMAIL"].strip()

def get_users():
    response = requests.get(
        API_URL,
        params={
            "results": COUNT,
            "seed": os.getenv("RANDOMUSER_SEED", "randomuser-threaded-demo"),
            "inc": "name,dob,location,phone,login",
            "noinfo": "1",
        },
        timeout=90,
    )
    response.raise_for_status()
    users = response.json().get("results", [])
    if len(users) != COUNT:
        raise RuntimeError(f"Expected {COUNT} profiles, got {len(users)}")
    return users

def full_name(u):
    n = u.get("name") or {}
    return " ".join(str(n.get(k, "")).strip() for k in ("title", "first", "last") if n.get(k)).strip()

def birthday(u):
    return str((u.get("dob") or {}).get("date", ""))[:10]

def address(u):
    loc = u.get("location") or {}
    street = loc.get("street") or {}
    street_line = " ".join(str(street.get(k, "")).strip() for k in ("number", "name") if street.get(k))
    parts = [street_line, loc.get("city"), loc.get("state"), loc.get("postcode"), loc.get("country")]
    return ", ".join(str(p).strip() for p in parts if p not in (None, ""))

def build_services():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return (
        build("sheets", "v4", credentials=creds, cache_discovery=False),
        build("drive", "v3", credentials=creds, cache_discovery=False),
    )

def main():
    users = get_users()
    sheets, drive = build_services()

    spreadsheet = sheets.spreadsheets().create(body={
        "properties": {"title": TITLE},
        "sheets": [{"properties": {"title": "Generated Data", "gridProperties": {"rowCount": COUNT, "columnCount": 5}}}],
    }, fields="spreadsheetId,spreadsheetUrl,sheets(properties(sheetId,title))").execute()

    spreadsheet_id = spreadsheet["spreadsheetId"]
    sheet_id = spreadsheet["sheets"][0]["properties"]["sheetId"]

    # Only A1:E398 is written/formatted/commented. No F column requests are made.
    values = [["Entry"] * 5 for _ in range(COUNT)]
    sheets.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="'Generated Data'!A1:E398",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    sheets.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [{
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": COUNT,
                "startColumnIndex": 0,
                "endColumnIndex": 5,
            },
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "italic": True, "fontSize": 11}}},
            "fields": "userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.italic,userEnteredFormat.textFormat.fontSize",
        }
    }]}).execute()

    comment_requests = []
    for row_idx, u in enumerate(users):
        login = u.get("login") or {}
        fields = [
            full_name(u),
            birthday(u),
            address(u),
            str(u.get("phone") or ""),
            str(login.get("password") or ""),
        ]
        for col_idx, content in enumerate(fields):
            comment_requests.append({
                "insertComment": {
                    "content": str(content or "(empty)"),
                    "coordinate": {
                        "sheetId": sheet_id,
                        "rowIndex": row_idx,
                        "columnIndex": col_idx,
                    },
                }
            })

    # Batch conservatively; each cell gets exactly one comment thread and no replies.
    for start in range(0, len(comment_requests), 100):
        batch = comment_requests[start:start + 100]
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": batch},
        ).execute()
        print(f"Inserted comments: {min(start + len(batch), len(comment_requests))}/{len(comment_requests)}")
        time.sleep(0.15)

    # Share the created Google Sheet with the user's email.
    drive.permissions().create(
        fileId=spreadsheet_id,
        body={"type": "user", "role": "writer", "emailAddress": SHARE_WITH_EMAIL},
        sendNotificationEmail=True,
        fields="id,emailAddress,role",
    ).execute()

    # Read back comments for verification.
    check = sheets.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        includeGridData=False,
        commentsViewMode="COMMENTS_VIEW_MODE_INCLUDED",
        fields="spreadsheetId,comments,sheets(properties(sheetId,title),commentAnchors)",
    ).execute()
    comment_count = len(check.get("comments", []))
    print(f"SPREADSHEET_URL={spreadsheet.get('spreadsheetUrl')}")
    print(f"SPREADSHEET_ID={spreadsheet_id}")
    print(f"EXPECTED_COMMENTS={COUNT * 5}")
    print(f"COMMENTS_READ_BACK={comment_count}")
    if comment_count != COUNT * 5:
        raise RuntimeError(f"Comment verification failed: expected 1990, got {comment_count}")
    print("SUCCESS: Google Sheets native comment threads verified.")

if __name__ == "__main__":
    try:
        main()
    except (HttpError, Exception) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
