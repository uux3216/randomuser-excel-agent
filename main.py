import os
import sys
import requests
from openpyxl import load_workbook
from openpyxl.comments import Comment

API_URL = "https://randomuser.me/api/1.4/"
WORKBOOK_PATH = os.getenv("WORKBOOK_PATH", "B7344.xlsx")
SHEET_NAME = "Work 1"
FIRST_ROW = 3
LAST_ROW = 400
FIRST_COL = 2  # B
LAST_COL = 6   # F — never access or modify columns A or G+
COUNT = LAST_ROW - FIRST_ROW + 1
SEED = os.getenv("RANDOMUSER_SEED", "excel-agent-demo")
COMMENT_AUTHOR = "RandomUser Agent"

def full_name(user):
    name = user.get("name") or {}
    return " ".join(
        str(part).strip()
        for part in (name.get("title"), name.get("first"), name.get("last"))
        if part not in (None, "")
    ).strip()

def full_address(user):
    loc = user.get("location") or {}
    street = loc.get("street") or {}
    street_text = " ".join(
        str(street.get(k, "")).strip()
        for k in ("number", "name")
        if street.get(k) not in (None, "")
    )
    parts = [
        street_text, loc.get("city"), loc.get("state"),
        loc.get("postcode"), loc.get("country")
    ]
    return ", ".join(str(x).strip() for x in parts if x not in (None, ""))

def fetch_users():
    params = {
        "results": COUNT,
        "seed": SEED,
        "inc": "name,dob,location,phone,login",
        "noinfo": "1",
    }
    response = requests.get(API_URL, params=params, timeout=60)
    response.raise_for_status()
    payload = response.json()
    if payload.get("error"):
        raise RuntimeError(f"RandomUser API error: {payload['error']}")
    users = payload.get("results")
    if not isinstance(users, list) or len(users) != COUNT:
        got = len(users) if isinstance(users, list) else "invalid response"
        raise RuntimeError(f"Expected {COUNT} users, received {got}.")
    return users

def birthday(user):
    dob = user.get("dob") or {}
    return str(dob.get("date", ""))[:10]

def set_comment_only(cell, value):
    # Called only for explicitly bounded cells B:F. Store data in comment, not cell text.
    cell.value = None
    cell.comment = Comment(str(value or ""), COMMENT_AUTHOR)

def main():
    if not os.path.isfile(WORKBOOK_PATH):
        raise FileNotFoundError(
            f"Workbook not found: {WORKBOOK_PATH}. Put the existing workbook "
            "in the repository root or set WORKBOOK_PATH."
        )

    print(f"Loading existing workbook: {WORKBOOK_PATH}")
    wb = load_workbook(WORKBOOK_PATH)
    if SHEET_NAME not in wb.sheetnames:
        raise KeyError(f"Sheet '{SHEET_NAME}' not found. Sheets: {wb.sheetnames}")

    ws = wb[SHEET_NAME]

    # Validate only the expected header cells, all within B:F.
    expected_headers = {
        2: "NAME",
        3: "BIRTH DAY",
        4: "ADRESS",
        5: "PHONE NUMBER",
        6: "PASSWORD",
    }
    for col, expected in expected_headers.items():
        actual = ws.cell(row=2, column=col).value
        if actual != expected:
            raise ValueError(
                f"Expected header {expected!r} in {ws.cell(row=2, column=col).coordinate}, "
                f"found {actual!r}; stopping without saving."
            )

    users = fetch_users()

    # Explicit hard bounds: rows 3..400 and columns 2..6 (B..F) only.
    for row, user in zip(range(FIRST_ROW, LAST_ROW + 1), users):
        login = user.get("login") or {}
        values = [
            full_name(user),                    # B: NAME
            birthday(user),                     # C: BIRTH DAY
            full_address(user),                 # D: ADRESS
            str(user.get("phone") or ""),       # E: PHONE NUMBER
            str(login.get("password") or ""),   # F: PASSWORD
        ]
        for col in range(FIRST_COL, LAST_COL + 1):
            value = values[col - FIRST_COL]
            cell = ws.cell(row=row, column=col)
            set_comment_only(cell, value)

    # Save back to the same workbook. No iteration over all worksheet cells,
    # no dimension-wide styling, and no access to columns A or G and beyond.
    wb.save(WORKBOOK_PATH)
    print(
        f"SUCCESS: updated only {SHEET_NAME}!B{FIRST_ROW}:F{LAST_ROW} "
        f"({COUNT} rows; {COUNT * (LAST_COL - FIRST_COL + 1)} comments)."
    )
    print("No script operation targets column A or columns G onward.")

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
